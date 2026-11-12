//! Helpers for cleaning Windows Recent folder shortcuts.

mod shell_link;

use crate::error::WincentError;
#[cfg(test)]
use crate::path::normalize_path_lightweight;
use crate::path::paths_equal;
use crate::windows::get_windows_recent_folder;
use crate::WincentResult;
pub(crate) use shell_link::{is_lnk_file, resolve_lnk_target};
#[cfg(test)]
use shell_link::{
    parse_shell_link_summary, COMMON_NETWORK_RELATIVE_LINK_AND_PATH_SUFFIX, HAS_LINK_INFO,
    IS_UNICODE, LINK_CLSID, SHELL_LINK_HEADER_SIZE, VOLUME_ID_AND_LOCAL_BASE_PATH,
};
use std::fs;
use std::path::{Path, PathBuf};
use std::time::Duration;

/// Deletes Recent-folder `.lnk` files whose resolved target matches any target.
pub(crate) fn delete_recent_links_for_targets(
    targets: &[String],
    timeout: Duration,
) -> WincentResult<Vec<PathBuf>> {
    let recent_folder = PathBuf::from(get_windows_recent_folder()?);
    delete_recent_links_for_targets_in(&recent_folder, targets, timeout, resolve_lnk_target)
}

/// Finds shortcuts in the Windows Recent folder whose resolved target equals `target`.
///
/// The comparison uses Windows path semantics: it is case-insensitive and treats forward
/// slashes and backslashes as equivalent. Broken or unresolvable shortcuts are ignored.
/// Shortcut targets are not accessed, so offline drives and network paths cannot block lookup.
#[cfg(test)]
pub(crate) fn find_windows_recent_links_for_target(target: &str) -> WincentResult<Vec<PathBuf>> {
    let recent_folder = PathBuf::from(get_windows_recent_folder()?);
    let normalized_target = normalize_path_lightweight(target);
    let mut matches = Vec::new();

    for lnk_path in recent_lnk_paths(&recent_folder)? {
        let Some(resolved_target) = resolve_lnk_target(&lnk_path, Duration::ZERO)? else {
            continue;
        };

        if normalize_path_lightweight(&resolved_target) == normalized_target {
            matches.push(lnk_path);
        }
    }

    Ok(matches)
}

fn delete_recent_links_for_targets_in<F>(
    recent_folder: &Path,
    targets: &[String],
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

        if resolved_target_matches(&resolved_target, targets) {
            fs::remove_file(&path).map_err(WincentError::Io)?;
            deleted.push(path);
        }
    }

    Ok(deleted)
}

fn resolved_target_matches(resolved_target: &str, targets: &[String]) -> bool {
    targets
        .iter()
        .any(|target| paths_equal(resolved_target, target))
}

fn sorted_lnk_file_paths(entries: impl IntoIterator<Item = (PathBuf, bool)>) -> Vec<PathBuf> {
    let mut paths: Vec<_> = entries
        .into_iter()
        .filter_map(|(path, is_file)| (is_file && is_lnk_file(&path)).then_some(path))
        .collect();
    paths.sort();
    paths
}

pub(crate) fn recent_lnk_paths(recent_folder: &Path) -> WincentResult<Vec<PathBuf>> {
    let mut entries = Vec::new();

    for entry in fs::read_dir(recent_folder).map_err(WincentError::Io)? {
        let entry = entry.map_err(WincentError::Io)?;
        let path = entry.path();
        let file_type = entry.file_type().map_err(WincentError::Io)?;
        entries.push((path, file_type.is_file()));
    }

    Ok(sorted_lnk_file_paths(entries))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lnk_extension_detection_is_case_insensitive() {
        assert!(is_lnk_file(Path::new("example.lnk")));
        assert!(is_lnk_file(Path::new("example.LNK")));
        assert!(!is_lnk_file(Path::new("example.txt")));
        assert!(!is_lnk_file(Path::new("example")));
    }

    #[test]
    fn resolved_target_matching_uses_windows_path_semantics() {
        let targets = vec!["C:\\One".to_string(), "c:/two/".to_string()];

        assert!(resolved_target_matches("c:/one/", &targets));
        assert!(resolved_target_matches("C:\\TWO", &targets));
        assert!(!resolved_target_matches("C:\\Other", &targets));
    }

    #[test]
    fn parse_shell_link_rejects_invalid_header() {
        assert_eq!(parse_shell_link_summary(b"not a link"), None);
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
    fn recent_lnk_path_filter_is_pure_and_sorted() {
        let a = PathBuf::from("a.lnk");
        let b = PathBuf::from("b.LNK");
        let txt = PathBuf::from("c.txt");
        let directory = PathBuf::from("directory.lnk");

        assert_eq!(
            sorted_lnk_file_paths([
                (b.clone(), true),
                (txt, true),
                (directory, false),
                (a.clone(), true),
            ]),
            vec![a, b]
        );
    }

    const FILE_ATTRIBUTE_ARCHIVE_FOR_TEST: u32 = 0x0000_0020;

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
