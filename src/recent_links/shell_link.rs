use crate::error::WincentError;
use crate::WincentResult;
use std::fs;
use std::path::Path;

pub(super) const SHELL_LINK_HEADER_SIZE: u32 = 0x4c;
pub(super) const LINK_CLSID: [u8; 16] = [
    0x01, 0x14, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];
pub(super) const HAS_LINK_TARGET_ID_LIST: u32 = 0x0000_0001;
pub(super) const HAS_LINK_INFO: u32 = 0x0000_0002;
const FORCE_NO_LINK_INFO: u32 = 0x0000_0200;
pub(super) const HAS_NAME: u32 = 0x0000_0004;
pub(super) const HAS_RELATIVE_PATH: u32 = 0x0000_0008;
const HAS_WORKING_DIR: u32 = 0x0000_0010;
const HAS_ARGUMENTS: u32 = 0x0000_0020;
const HAS_ICON_LOCATION: u32 = 0x0000_0040;
pub(super) const IS_UNICODE: u32 = 0x0000_0080;
pub(super) const VOLUME_ID_AND_LOCAL_BASE_PATH: u32 = 0x0000_0001;
pub(super) const COMMON_NETWORK_RELATIVE_LINK_AND_PATH_SUFFIX: u32 = 0x0000_0002;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct ShellLinkSummary {
    pub(super) target_path: Option<String>,
    pub(super) file_attributes: u32,
    pub(super) relative_path: Option<String>,
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

pub(crate) fn resolve_lnk_target(lnk_path: &Path) -> WincentResult<Option<String>> {
    let Some(summary) = read_shell_link_summary(lnk_path)? else {
        return Ok(None);
    };

    if let Some(target_path) = summary.target_path {
        return Ok(Some(target_path));
    }

    let Some(relative_path) = summary.relative_path else {
        return Ok(None);
    };
    let Some(parent) = lnk_path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    else {
        return Ok(None);
    };

    Ok(Some(
        parent.join(relative_path).to_string_lossy().into_owned(),
    ))
}

fn read_shell_link_summary(lnk_path: &Path) -> WincentResult<Option<ShellLinkSummary>> {
    let data = match fs::read(lnk_path) {
        Ok(data) => data,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(error) => return Err(WincentError::Io(error)),
    };

    Ok(parse_shell_link_summary(&data))
}

pub(super) fn parse_shell_link_summary(data: &[u8]) -> Option<ShellLinkSummary> {
    if !looks_like_lnk(data) {
        return None;
    }

    let flags = read_u32(data, 0x14).ok()?;
    let file_attributes = read_u32(data, 0x18).unwrap_or_default();
    let mut offset = SHELL_LINK_HEADER_SIZE as usize;

    if flags & HAS_LINK_TARGET_ID_LIST != 0 {
        if let Some(next) = read_u16(data, offset).ok().and_then(|id_list_size| {
            offset
                .checked_add(2)
                .and_then(|value| value.checked_add(id_list_size as usize))
        }) {
            if next <= data.len() {
                offset = next;
            }
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
    let Some(length_end) = offset.checked_add(2) else {
        return (None, offset);
    };
    let Some(length_bytes) = data.get(offset..length_end) else {
        return (None, offset);
    };
    let chars = u16::from_le_bytes(length_bytes.try_into().expect("2-byte chunk")) as usize;
    let string_start = length_end;
    if flags & IS_UNICODE != 0 {
        let Some(string_end) = chars
            .checked_mul(2)
            .and_then(|length| string_start.checked_add(length))
        else {
            return (None, string_start);
        };
        if string_end > data.len() {
            return (None, string_start);
        }
        (
            Some(decode_utf16_lossy(&data[string_start..string_end])),
            string_end,
        )
    } else {
        let Some(string_end) = string_start.checked_add(chars) else {
            return (None, string_start);
        };
        if string_end > data.len() {
            return (None, string_start);
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
        .is_some_and(|extension| extension.eq_ignore_ascii_case("lnk"))
}
