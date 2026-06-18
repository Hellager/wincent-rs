//! Helpers for cleaning Windows Recent folder shortcuts.

use crate::error::WincentError;
use crate::utils::{get_windows_recent_folder, paths_equal};
use crate::WincentResult;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::mpsc;
use std::thread;
use std::time::Duration;

const FILE_ATTRIBUTE_DIRECTORY: u32 = 0x0000_0010;
const SHELL_LINK_HEADER_SIZE: u32 = 0x4c;
const LINK_CLSID: [u8; 16] = [
    0x01, 0x14, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];
const HAS_LINK_TARGET_ID_LIST: u32 = 0x0000_0001;
const HAS_LINK_INFO: u32 = 0x0000_0002;
const FORCE_NO_LINK_INFO: u32 = 0x0000_0200;
const HAS_NAME: u32 = 0x0000_0004;
const HAS_RELATIVE_PATH: u32 = 0x0000_0008;
const HAS_WORKING_DIR: u32 = 0x0000_0010;
const HAS_ARGUMENTS: u32 = 0x0000_0020;
const HAS_ICON_LOCATION: u32 = 0x0000_0040;
const IS_UNICODE: u32 = 0x0000_0080;
const VOLUME_ID_AND_LOCAL_BASE_PATH: u32 = 0x0000_0001;
const COMMON_NETWORK_RELATIVE_LINK_AND_PATH_SUFFIX: u32 = 0x0000_0002;

/// Resolved shortcut target plus best-effort target type information.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct LnkResolution {
    pub(crate) path: String,
    pub(crate) is_dir: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ShellLinkSummary {
    target_path: Option<String>,
    file_attributes: u32,
    relative_path: Option<String>,
    target_is_network: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct LinkInfoSummary {
    size: usize,
    local_base_path: Option<String>,
    common_path_suffix: Option<String>,
    network_path: Option<String>,
    resolved_path: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct NetworkLinkSummary {
    net_name: Option<String>,
}

/// Deletes Recent-folder `.lnk` files whose resolved target equals `target`.
pub(crate) fn delete_recent_links_for_target(
    target: &str,
    timeout: Duration,
) -> WincentResult<Vec<PathBuf>> {
    let recent_folder = PathBuf::from(get_windows_recent_folder()?);
    delete_recent_links_for_target_in(&recent_folder, target, timeout, resolve_lnk_target)
}

fn delete_recent_links_for_target_in<F>(
    recent_folder: &Path,
    target: &str,
    timeout: Duration,
    mut resolver: F,
) -> WincentResult<Vec<PathBuf>>
where
    F: FnMut(&Path, Duration) -> WincentResult<Option<String>>,
{
    let mut deleted = Vec::new();

    for entry in fs::read_dir(recent_folder).map_err(WincentError::Io)? {
        let entry = entry.map_err(WincentError::Io)?;
        let path = entry.path();
        let file_type = entry.file_type().map_err(WincentError::Io)?;
        if !file_type.is_file() || !is_lnk_file(&path) {
            continue;
        }

        let Some(resolved_target) = resolver(&path, timeout)? else {
            continue;
        };

        if paths_equal(&resolved_target, target) {
            fs::remove_file(&path).map_err(WincentError::Io)?;
            deleted.push(path);
        }
    }

    Ok(deleted)
}

pub(crate) fn recent_lnk_paths(recent_folder: &Path) -> WincentResult<Vec<PathBuf>> {
    let mut paths = Vec::new();

    for entry in fs::read_dir(recent_folder).map_err(WincentError::Io)? {
        let entry = entry.map_err(WincentError::Io)?;
        let path = entry.path();
        let file_type = entry.file_type().map_err(WincentError::Io)?;
        if file_type.is_file() && is_lnk_file(&path) {
            paths.push(path);
        }
    }

    paths.sort();
    Ok(paths)
}

pub(crate) fn resolve_lnk_target(
    lnk_path: &Path,
    timeout: Duration,
) -> WincentResult<Option<String>> {
    Ok(resolve_lnk_with_type(lnk_path, timeout)?.map(|resolution| resolution.path))
}

pub(crate) fn resolve_lnk_with_type(
    lnk_path: &Path,
    timeout: Duration,
) -> WincentResult<Option<LnkResolution>> {
    let data = match fs::read(lnk_path) {
        Ok(data) => data,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(error) => return Err(WincentError::Io(error)),
    };
    let Some(summary) = parse_shell_link_summary(&data) else {
        return Ok(None);
    };
    let Some(path) = summary.target_path.or(summary.relative_path) else {
        return Ok(None);
    };
    let is_dir = lnk_target_is_dir(
        &path,
        summary.file_attributes,
        summary.target_is_network,
        timeout,
    );

    Ok(Some(LnkResolution { path, is_dir }))
}

fn lnk_target_is_dir(
    path: &str,
    attributes: u32,
    target_is_network: bool,
    timeout: Duration,
) -> Option<bool> {
    if attributes != 0 {
        return Some((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0);
    }

    // Some shortcuts do not carry useful file attributes. Metadata may fail for
    // missing, unavailable network, or access-denied targets; restore flows
    // intentionally treat that as unknown.
    if should_timeout_protect_metadata(path, target_is_network) {
        return metadata_is_dir_with_timeout(path.to_string(), timeout);
    }

    match fs::metadata(path) {
        Ok(metadata) => Some(metadata.is_dir()),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => None,
        Err(_) => None,
    }
}

fn should_timeout_protect_metadata(path: &str, target_is_network: bool) -> bool {
    target_is_network || looks_like_unc_path(path)
}

fn metadata_is_dir_with_timeout(path: String, timeout: Duration) -> Option<bool> {
    metadata_is_dir_with_timeout_using(path, timeout, |path| {
        fs::metadata(path).map(|metadata| metadata.is_dir())
    })
}

fn metadata_is_dir_with_timeout_using<F>(
    path: String,
    timeout: Duration,
    metadata_is_dir: F,
) -> Option<bool>
where
    F: FnOnce(&str) -> std::io::Result<bool> + Send + 'static,
{
    if timeout.is_zero() {
        return None;
    }

    let (tx, rx) = mpsc::channel();
    thread::spawn(move || {
        let result = metadata_is_dir(&path);
        let _ = tx.send(result);
    });

    match rx.recv_timeout(timeout) {
        Ok(Ok(is_dir)) => Some(is_dir),
        Ok(Err(_)) | Err(_) => None,
    }
}

fn parse_shell_link_summary(data: &[u8]) -> Option<ShellLinkSummary> {
    if !looks_like_lnk(data) {
        return None;
    }

    let flags = read_u32(data, 0x14).ok()?;
    let file_attributes = read_u32(data, 0x18).unwrap_or_default();
    let mut offset = SHELL_LINK_HEADER_SIZE as usize;

    if flags & HAS_LINK_TARGET_ID_LIST != 0 {
        let id_list_size = read_u16(data, offset).ok()? as usize;
        offset = offset.checked_add(2)?.checked_add(id_list_size)?;
        if offset > data.len() {
            return None;
        }
    }

    let mut link_info = None;
    if flags & HAS_LINK_INFO != 0
        && flags & FORCE_NO_LINK_INFO == 0
        && offset.checked_add(28).is_some_and(|end| end <= data.len())
    {
        link_info = parse_link_info(data, offset);
        if let Some(info) = &link_info {
            offset = offset.saturating_add(info.size).min(data.len());
        }
    }

    let mut relative_path = None;
    if flags & HAS_NAME != 0 {
        let (_, next) = read_lnk_string(data, offset, flags);
        offset = next;
    }
    if flags & HAS_RELATIVE_PATH != 0 {
        let (value, next) = read_lnk_string(data, offset, flags);
        relative_path = value;
        offset = next;
    }
    if flags & HAS_WORKING_DIR != 0 {
        let (_, next) = read_lnk_string(data, offset, flags);
        offset = next;
    }
    if flags & HAS_ARGUMENTS != 0 {
        let (_, next) = read_lnk_string(data, offset, flags);
        offset = next;
    }
    if flags & HAS_ICON_LOCATION != 0 {
        let (_, _next) = read_lnk_string(data, offset, flags);
    }

    let target_path = link_info.as_ref().and_then(|info| {
        info.resolved_path
            .clone()
            .or_else(|| info.local_base_path.clone())
            .or_else(|| info.network_path.clone())
            .or_else(|| info.common_path_suffix.clone())
    });

    let target_is_network = link_info
        .as_ref()
        .is_some_and(|info| info.network_path.is_some())
        || target_path.as_deref().is_some_and(looks_like_unc_path)
        || relative_path.as_deref().is_some_and(looks_like_unc_path);

    Some(ShellLinkSummary {
        target_path,
        file_attributes,
        relative_path,
        target_is_network,
    })
}

fn looks_like_lnk(data: &[u8]) -> bool {
    data.len() >= SHELL_LINK_HEADER_SIZE as usize
        && read_u32(data, 0).ok() == Some(SHELL_LINK_HEADER_SIZE)
        && data.get(4..20) == Some(&LINK_CLSID)
}

fn parse_link_info(data: &[u8], link_info_start: usize) -> Option<LinkInfoSummary> {
    let link_info_size = read_u32(data, link_info_start).ok()? as usize;
    let link_info_header_size = read_u32(data, link_info_start + 4).ok()? as usize;
    let link_info_end = link_info_start.checked_add(link_info_size)?;
    if link_info_size < 28
        || link_info_header_size < 28
        || link_info_header_size > link_info_size
        || link_info_end > data.len()
    {
        return None;
    }

    let link_info_flags = read_u32(data, link_info_start + 8).ok()?;
    let local_base_offset = read_u32(data, link_info_start + 16).ok()? as usize;
    let network_offset = read_u32(data, link_info_start + 20).ok()? as usize;
    let common_suffix_offset = read_u32(data, link_info_start + 24).ok()? as usize;
    let local_base_unicode_offset = if link_info_header_size >= 0x24 {
        read_u32(data, link_info_start + 28).ok()? as usize
    } else {
        0
    };
    let common_suffix_unicode_offset = if link_info_header_size >= 0x24 {
        read_u32(data, link_info_start + 32).ok()? as usize
    } else {
        0
    };

    let local_base_path = if link_info_flags & VOLUME_ID_AND_LOCAL_BASE_PATH != 0 {
        read_utf16_z_string_in_link_info(
            data,
            link_info_start,
            link_info_size,
            local_base_unicode_offset,
        )
        .or_else(|| {
            read_c_string_in_link_info(data, link_info_start, link_info_size, local_base_offset)
        })
    } else {
        None
    };
    let common_path_suffix = read_utf16_z_string_in_link_info(
        data,
        link_info_start,
        link_info_size,
        common_suffix_unicode_offset,
    )
    .or_else(|| {
        read_c_string_in_link_info(data, link_info_start, link_info_size, common_suffix_offset)
    });
    let network = if link_info_flags & COMMON_NETWORK_RELATIVE_LINK_AND_PATH_SUFFIX != 0 {
        parse_common_network_relative_link(data, link_info_start, link_info_size, network_offset)
    } else {
        None
    };
    let network_path = network.and_then(|network| network.net_name);

    let local_resolved_path = match (&local_base_path, &common_path_suffix) {
        (Some(base), Some(suffix)) if !base.is_empty() && !suffix.is_empty() => {
            Some(join_windows_path(base, suffix))
        }
        (Some(base), _) if looks_like_windows_path(base) => Some(base.clone()),
        (_, Some(suffix)) if looks_like_windows_path(suffix) => Some(suffix.clone()),
        _ => None,
    };
    let network_resolved_path = match (&network_path, &common_path_suffix) {
        (Some(base), Some(suffix)) if !base.is_empty() && !suffix.is_empty() => {
            Some(join_windows_path(base, suffix))
        }
        (Some(base), _) if looks_like_unc_path(base) => Some(base.clone()),
        _ => None,
    };
    let resolved_path = local_resolved_path.or(network_resolved_path);

    Some(LinkInfoSummary {
        size: link_info_size,
        local_base_path,
        common_path_suffix,
        network_path,
        resolved_path,
    })
}

fn parse_common_network_relative_link(
    data: &[u8],
    link_info_start: usize,
    link_info_size: usize,
    relative_offset: usize,
) -> Option<NetworkLinkSummary> {
    if relative_offset == 0 || relative_offset + 20 > link_info_size {
        return None;
    }
    let start = link_info_start.checked_add(relative_offset)?;
    let size = read_u32(data, start).ok()? as usize;
    if size < 20 || relative_offset + size > link_info_size {
        return None;
    }

    let net_name_offset = read_u32(data, start + 8).ok()? as usize;
    let net_name_unicode_offset = if net_name_offset > 0x14 {
        read_u32(data, start + 20).unwrap_or_default() as usize
    } else {
        0
    };
    let net_name = read_utf16_z_string_in_link_info(data, start, size, net_name_unicode_offset)
        .or_else(|| read_c_string_in_link_info(data, start, size, net_name_offset));

    Some(NetworkLinkSummary { net_name })
}

fn read_c_string_in_link_info(
    data: &[u8],
    link_info_start: usize,
    link_info_size: usize,
    relative_offset: usize,
) -> Option<String> {
    if relative_offset == 0 || relative_offset >= link_info_size {
        return None;
    }
    let absolute_offset = link_info_start.checked_add(relative_offset)?;
    let link_info_end = link_info_start.checked_add(link_info_size)?;
    read_c_string(data.get(..link_info_end)?, absolute_offset)
}

fn read_utf16_z_string_in_link_info(
    data: &[u8],
    link_info_start: usize,
    link_info_size: usize,
    relative_offset: usize,
) -> Option<String> {
    if relative_offset == 0 || relative_offset >= link_info_size {
        return None;
    }
    let absolute_offset = link_info_start.checked_add(relative_offset)?;
    let link_info_end = link_info_start.checked_add(link_info_size)?;
    read_utf16_z_string(data.get(..link_info_end)?, absolute_offset)
}

fn read_lnk_string(data: &[u8], offset: usize, flags: u32) -> (Option<String>, usize) {
    if offset + 2 > data.len() {
        return (None, data.len());
    }
    let chars =
        u16::from_le_bytes(data[offset..offset + 2].try_into().expect("2-byte chunk")) as usize;
    let string_start = offset + 2;
    if flags & IS_UNICODE != 0 {
        let string_end = string_start.saturating_add(chars.saturating_mul(2));
        if string_end > data.len() {
            return (None, data.len());
        }
        (
            Some(decode_utf16_lossy(&data[string_start..string_end])),
            string_end,
        )
    } else {
        let string_end = string_start.saturating_add(chars);
        if string_end > data.len() {
            return (None, data.len());
        }
        (
            Some(String::from_utf8_lossy(&data[string_start..string_end]).to_string()),
            string_end,
        )
    }
}

fn read_c_string(data: &[u8], offset: usize) -> Option<String> {
    if offset >= data.len() {
        return None;
    }
    let end = data[offset..].iter().position(|byte| *byte == 0)? + offset;
    Some(String::from_utf8_lossy(&data[offset..end]).to_string())
}

fn read_utf16_z_string(data: &[u8], offset: usize) -> Option<String> {
    if offset + 1 >= data.len() {
        return None;
    }

    let mut end = offset;
    while end + 1 < data.len() {
        if data[end] == 0 && data[end + 1] == 0 {
            break;
        }
        end += 2;
    }

    if end + 1 >= data.len() {
        return None;
    }

    Some(decode_utf16_lossy(&data[offset..end]))
}

fn decode_utf16_lossy(bytes: &[u8]) -> String {
    let words: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes(chunk.try_into().expect("2-byte chunk")))
        .collect();
    String::from_utf16_lossy(&words)
}

fn read_u16(data: &[u8], offset: usize) -> Result<u16, String> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| format!("unexpected end of data at offset {offset}"))?;
    Ok(u16::from_le_bytes(bytes.try_into().expect("2-byte chunk")))
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32, String> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| format!("unexpected end of data at offset {offset}"))?;
    Ok(u32::from_le_bytes(bytes.try_into().expect("4-byte chunk")))
}

fn join_windows_path(base: &str, suffix: &str) -> String {
    if looks_like_windows_path(suffix) || looks_like_unc_path(suffix) {
        return suffix.to_string();
    }
    if base.ends_with('\\') || base.ends_with('/') || suffix.is_empty() {
        format!("{base}{suffix}")
    } else {
        format!("{base}\\{suffix}")
    }
}

fn looks_like_windows_path(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() >= 3 && bytes[1] == b':' && (bytes[2] == b'\\' || bytes[2] == b'/')
}

fn looks_like_unc_path(value: &str) -> bool {
    value.starts_with("\\\\") || value.starts_with("//")
}

pub(crate) fn is_lnk_file(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.eq_ignore_ascii_case("lnk"))
        .unwrap_or(false)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;
    use tempfile::tempdir;

    #[test]
    fn lnk_extension_detection_is_case_insensitive() {
        assert!(is_lnk_file(Path::new("example.lnk")));
        assert!(is_lnk_file(Path::new("example.LNK")));
        assert!(!is_lnk_file(Path::new("example.txt")));
        assert!(!is_lnk_file(Path::new("example")));
    }

    #[test]
    fn deep_clean_deletes_only_matching_recent_links() -> WincentResult<()> {
        let dir = tempdir().map_err(WincentError::Io)?;
        let matching = dir.path().join("matching.lnk");
        let other = dir.path().join("other.lnk");
        let broken = dir.path().join("broken.lnk");
        let non_lnk = dir.path().join("matching.txt");
        fs::write(&matching, b"matching").map_err(WincentError::Io)?;
        fs::write(&other, b"other").map_err(WincentError::Io)?;
        fs::write(&broken, b"broken").map_err(WincentError::Io)?;
        fs::write(&non_lnk, b"not a shortcut").map_err(WincentError::Io)?;

        let mut targets = HashMap::new();
        targets.insert(matching.clone(), "C:\\Work\\Report.docx".to_string());
        targets.insert(other.clone(), "C:\\Work\\Other.docx".to_string());

        let deleted = delete_recent_links_for_target_in(
            dir.path(),
            "c:/work/report.docx",
            Duration::from_secs(1),
            |path, _| Ok(targets.get(path).cloned()),
        )?;

        assert_eq!(deleted, vec![matching.clone()]);
        assert!(!matching.exists());
        assert!(other.exists());
        assert!(broken.exists());
        assert!(non_lnk.exists());
        Ok(())
    }

    #[test]
    fn deep_clean_propagates_resolver_errors() -> WincentResult<()> {
        let dir = tempdir().map_err(WincentError::Io)?;
        let link = dir.path().join("matching.lnk");
        fs::write(&link, b"matching").map_err(WincentError::Io)?;

        let error = delete_recent_links_for_target_in(
            dir.path(),
            "c:/work/report.docx",
            Duration::from_secs(1),
            |_, _| Err(WincentError::Timeout("resolver timed out".to_string())),
        )
        .unwrap_err();

        assert!(link.exists());
        assert!(matches!(
            error,
            WincentError::Timeout(message) if message == "resolver timed out"
        ));
        Ok(())
    }

    #[test]
    fn parse_shell_link_rejects_invalid_header() {
        assert_eq!(parse_shell_link_summary(b"not a link"), None);
    }

    #[test]
    fn parse_shell_link_uses_header_file_attributes_for_target_type() -> WincentResult<()> {
        let mut lnk = minimal_lnk();
        write_u32(&mut lnk, 0x18, FILE_ATTRIBUTE_DIRECTORY);

        let dir = tempdir().map_err(WincentError::Io)?;
        let path = dir.path().join("folder.lnk");
        fs::write(&path, lnk).map_err(WincentError::Io)?;

        let resolution = resolve_lnk_with_type(&path, Duration::ZERO)?.unwrap();

        assert_eq!(resolution.path, "relative-target");
        assert_eq!(resolution.is_dir, Some(true));
        Ok(())
    }

    #[test]
    fn lnk_target_type_uses_attributes_without_metadata_probe() {
        assert_eq!(
            lnk_target_is_dir(
                "\\\\server\\share\\folder",
                FILE_ATTRIBUTE_DIRECTORY,
                true,
                Duration::ZERO,
            ),
            Some(true)
        );
        assert_eq!(
            lnk_target_is_dir(
                "\\\\server\\share\\file.txt",
                FILE_ATTRIBUTE_ARCHIVE_FOR_TEST,
                true,
                Duration::ZERO,
            ),
            Some(false)
        );
    }

    #[test]
    fn network_metadata_timeout_returns_unknown_target_type() {
        let result = metadata_is_dir_with_timeout_using(
            "\\\\server\\share\\slow".to_string(),
            Duration::from_millis(1),
            |_| {
                std::thread::sleep(Duration::from_millis(50));
                Ok(true)
            },
        );

        assert_eq!(result, None);
    }

    #[test]
    fn local_target_type_uses_synchronous_metadata() -> WincentResult<()> {
        let dir = tempdir().map_err(WincentError::Io)?;
        let folder = dir.path().join("folder");
        fs::create_dir(&folder).map_err(WincentError::Io)?;

        assert_eq!(
            lnk_target_is_dir(&folder.to_string_lossy(), 0, false, Duration::ZERO),
            Some(true)
        );
        Ok(())
    }

    #[test]
    fn parse_shell_link_reads_unicode_local_base_and_suffix() {
        let lnk = lnk_with_link_info(
            Some("C:\\Users\\Alice"),
            Some("Documents\\report.docx"),
            None,
            FILE_ATTRIBUTE_ARCHIVE_FOR_TEST,
        );

        let summary = parse_shell_link_summary(&lnk).unwrap();

        assert_eq!(
            summary.target_path.as_deref(),
            Some("C:\\Users\\Alice\\Documents\\report.docx")
        );
        assert_eq!(summary.file_attributes, FILE_ATTRIBUTE_ARCHIVE_FOR_TEST);
    }

    #[test]
    fn parse_shell_link_reads_unc_network_path() {
        let lnk = lnk_with_link_info(
            None,
            Some("Share\\report.docx"),
            Some("\\\\server\\team"),
            FILE_ATTRIBUTE_ARCHIVE_FOR_TEST,
        );

        let summary = parse_shell_link_summary(&lnk).unwrap();

        assert_eq!(
            summary.target_path.as_deref(),
            Some("\\\\server\\team\\Share\\report.docx")
        );
    }

    #[test]
    fn recent_lnk_paths_returns_only_lnk_files_sorted() -> WincentResult<()> {
        let dir = tempdir().map_err(WincentError::Io)?;
        let a = dir.path().join("a.lnk");
        let b = dir.path().join("b.LNK");
        let txt = dir.path().join("c.txt");
        fs::write(&b, b"b").map_err(WincentError::Io)?;
        fs::write(&txt, b"txt").map_err(WincentError::Io)?;
        fs::write(&a, b"a").map_err(WincentError::Io)?;

        assert_eq!(recent_lnk_paths(dir.path())?, vec![a, b]);
        Ok(())
    }

    const FILE_ATTRIBUTE_ARCHIVE_FOR_TEST: u32 = 0x0000_0020;

    fn minimal_lnk() -> Vec<u8> {
        let mut lnk = vec![0u8; SHELL_LINK_HEADER_SIZE as usize];
        write_u32(&mut lnk, 0, SHELL_LINK_HEADER_SIZE);
        lnk[4..20].copy_from_slice(&LINK_CLSID);
        write_u32(&mut lnk, 0x14, HAS_RELATIVE_PATH | IS_UNICODE);
        write_lnk_string(&mut lnk, "relative-target");
        lnk
    }

    fn lnk_with_link_info(
        local_base: Option<&str>,
        common_suffix: Option<&str>,
        network_base: Option<&str>,
        attributes: u32,
    ) -> Vec<u8> {
        let mut lnk = vec![0u8; SHELL_LINK_HEADER_SIZE as usize];
        write_u32(&mut lnk, 0, SHELL_LINK_HEADER_SIZE);
        lnk[4..20].copy_from_slice(&LINK_CLSID);
        write_u32(&mut lnk, 0x14, HAS_LINK_INFO | IS_UNICODE);
        write_u32(&mut lnk, 0x18, attributes);

        let link_info_start = lnk.len();
        lnk.resize(link_info_start + 0x24, 0);
        let mut flags = 0u32;
        let mut local_base_offset = 0u32;
        let mut network_offset = 0u32;
        let mut common_suffix_offset = 0u32;

        if let Some(value) = local_base {
            flags |= VOLUME_ID_AND_LOCAL_BASE_PATH;
            local_base_offset = (lnk.len() - link_info_start) as u32;
            write_utf16_z(&mut lnk, value);
        }
        if let Some(value) = network_base {
            flags |= COMMON_NETWORK_RELATIVE_LINK_AND_PATH_SUFFIX;
            network_offset = (lnk.len() - link_info_start) as u32;
            write_network_link(&mut lnk, value);
        }
        if let Some(value) = common_suffix {
            common_suffix_offset = (lnk.len() - link_info_start) as u32;
            write_utf16_z(&mut lnk, value);
        }

        let size = (lnk.len() - link_info_start) as u32;
        write_u32(&mut lnk, link_info_start, size);
        write_u32(&mut lnk, link_info_start + 4, 0x24);
        write_u32(&mut lnk, link_info_start + 8, flags);
        write_u32(&mut lnk, link_info_start + 16, local_base_offset);
        write_u32(&mut lnk, link_info_start + 20, network_offset);
        write_u32(&mut lnk, link_info_start + 24, common_suffix_offset);
        write_u32(&mut lnk, link_info_start + 28, local_base_offset);
        write_u32(&mut lnk, link_info_start + 32, common_suffix_offset);
        lnk
    }

    fn write_network_link(data: &mut Vec<u8>, net_name: &str) {
        let start = data.len();
        data.resize(start + 0x18, 0);
        let net_name_offset = 0x18u32;
        write_utf16_z(data, net_name);
        let size = (data.len() - start) as u32;
        write_u32(data, start, size);
        write_u32(data, start + 8, net_name_offset);
        write_u32(data, start + 20, net_name_offset);
    }

    fn write_lnk_string(data: &mut Vec<u8>, value: &str) {
        let words: Vec<u16> = value.encode_utf16().collect();
        data.extend_from_slice(&(words.len() as u16).to_le_bytes());
        for word in words {
            data.extend_from_slice(&word.to_le_bytes());
        }
    }

    fn write_utf16_z(data: &mut Vec<u8>, value: &str) {
        for word in value.encode_utf16() {
            data.extend_from_slice(&word.to_le_bytes());
        }
        data.extend_from_slice(&0u16.to_le_bytes());
    }

    fn write_u32(data: &mut [u8], offset: usize, value: u32) {
        data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    }
}
