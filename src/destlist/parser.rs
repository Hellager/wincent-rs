use std::fs;
use std::path::{Path, PathBuf};

use crate::error::WincentError;
use crate::WincentResult;

use super::cfb::{
    decode_utf16_lossy, read_i32, read_u16, read_u32, read_u64, CompoundFile,
};

pub const RECENT_FILES_APPID: &str = "5f7b5f1e01b83767.automaticDestinations-ms";
pub const FREQUENT_FOLDERS_APPID: &str = "f01b4d95cf55d32a.automaticDestinations-ms";

/// Parsed `.automaticDestinations-ms` file.
#[derive(Debug, Clone)]
pub struct AutomaticDestinations {
    pub cfb_info: CfbInfo,
    pub dest_list: DestList,
}

/// CFB container metadata.
#[derive(Debug, Clone)]
pub struct CfbInfo {
    pub sector_size: usize,
    pub mini_sector_size: usize,
    pub mini_cutoff_size: u32,
    pub directory_entries: Vec<CfbDirectoryEntry>,
}

/// A single CFB directory entry (stream or storage).
#[derive(Debug, Clone)]
pub struct CfbDirectoryEntry {
    pub name: String,
    pub object_type: u8,
    pub start_sector: u32,
    pub stream_size: u64,
}

/// Parsed DestList stream header + entries.
#[derive(Debug, Clone)]
pub struct DestList {
    pub version: u32,
    pub declared_entry_count: usize,
    pub pinned_entry_count: u32,
    pub last_entry_id: u64,
    pub entries: Vec<DestListEntry>,
}

/// A single DestList entry.
#[derive(Debug, Clone)]
pub struct DestListEntry {
    /// Byte offset of this entry within the DestList stream (used by write.rs).
    #[allow(dead_code)]
    pub(crate) entry_offset: usize,
    pub entry_id: u64,
    /// Hex string of `entry_id`; names the corresponding Shell Link stream.
    pub stream_name: String,
    /// Raw path as stored; may be `"knownfolder:{GUID}"`.
    pub raw_path: String,
    /// Resolved path (knownfolder GUIDs resolved via Shell Link stream).
    pub path: String,
    /// `None` if not pinned.
    pub pin_order: Option<i32>,
    pub recent_rank: i32,
    /// `0` means hidden in v4.
    pub access_count: u32,
    pub last_access_filetime: Option<u64>,
}

/// Returns the path to the Explorer recent-files `.automaticDestinations-ms` file.
pub fn recent_files_dest_path() -> PathBuf {
    let appdata = std::env::var_os("APPDATA")
        .unwrap_or_else(|| "C:\\Users\\Default\\AppData\\Roaming".into());
    PathBuf::from(appdata)
        .join("Microsoft")
        .join("Windows")
        .join("Recent")
        .join("AutomaticDestinations")
        .join(RECENT_FILES_APPID)
}

/// Returns the path to the Explorer frequent-folders `.automaticDestinations-ms` file.
pub fn frequent_folders_dest_path() -> PathBuf {
    let appdata = std::env::var_os("APPDATA")
        .unwrap_or_else(|| "C:\\Users\\Default\\AppData\\Roaming".into());
    PathBuf::from(appdata)
        .join("Microsoft")
        .join("Windows")
        .join("Recent")
        .join("AutomaticDestinations")
        .join(FREQUENT_FOLDERS_APPID)
}

/// Parses an `.automaticDestinations-ms` file from disk.
pub fn parse_file(path: impl AsRef<Path>) -> WincentResult<AutomaticDestinations> {
    let path = path.as_ref();
    let data = fs::read(path).map_err(WincentError::Io)?;
    parse_bytes(data)
}

/// Parses an `.automaticDestinations-ms` file from an in-memory buffer.
pub fn parse_bytes(data: Vec<u8>) -> WincentResult<AutomaticDestinations> {
    let cfb = CompoundFile::parse(data).map_err(WincentError::DestListParse)?;
    let dest_list = parse_dest_list(&cfb)?;
    let cfb_info = CfbInfo {
        sector_size: cfb.sector_size,
        mini_sector_size: cfb.mini_sector_size,
        mini_cutoff_size: cfb.mini_cutoff_size,
        directory_entries: cfb
            .directory
            .iter()
            .map(|e| CfbDirectoryEntry {
                name: e.name.clone(),
                object_type: e.object_type,
                start_sector: e.start_sector,
                stream_size: e.stream_size,
            })
            .collect(),
    };
    Ok(AutomaticDestinations { cfb_info, dest_list })
}

fn parse_dest_list(cfb: &CompoundFile) -> WincentResult<DestList> {
    let dest_list = cfb
        .stream("DestList")
        .map_err(WincentError::DestListParse)?
        .ok_or_else(|| WincentError::DestListParse("DestList stream not found".to_string()))?;

    if dest_list.len() < 32 {
        return Err(WincentError::DestListParse(
            "DestList stream is too small".to_string(),
        ));
    }

    let version = read_u32(&dest_list, 0).map_err(WincentError::DestListParse)?;
    if version != 4 && version != 6 {
        return Err(WincentError::DestListUnsupportedVersion(version));
    }

    let entry_count = read_u32(&dest_list, 4).map_err(WincentError::DestListParse)? as usize;
    let pinned_entry_count =
        read_u32(&dest_list, 8).map_err(WincentError::DestListParse)?;
    let last_entry_id =
        read_u64(&dest_list, 16).map_err(WincentError::DestListParse)?;
    let mut offset = 32usize;
    let mut entries = Vec::new();

    for _ in 0..entry_count {
        if offset + 130 > dest_list.len() {
            break;
        }

        let entry_id = read_u64(&dest_list, offset + 88).map_err(WincentError::DestListParse)?;
        let last_access_filetime = read_u64(&dest_list, offset + 100).ok();
        let pin_status =
            read_i32(&dest_list, offset + 108).map_err(WincentError::DestListParse)?;
        let recent_rank =
            read_i32(&dest_list, offset + 112).map_err(WincentError::DestListParse)?;
        let access_count =
            read_u32(&dest_list, offset + 116).map_err(WincentError::DestListParse)?;
        let path_chars =
            read_u16(&dest_list, offset + 128).map_err(WincentError::DestListParse)? as usize;
        let path_start = offset + 130;
        let path_end = path_start + path_chars.saturating_mul(2);
        if path_end > dest_list.len() {
            break;
        }

        let raw_path = decode_utf16_lossy(&dest_list[path_start..path_end]);
        let stream_name = format!("{entry_id:x}");
        let resolved_path = resolve_path(cfb, &stream_name, &raw_path);

        entries.push(DestListEntry {
            entry_offset: offset,
            entry_id,
            stream_name,
            raw_path,
            path: resolved_path,
            pin_order: (pin_status >= 0).then_some(pin_status),
            recent_rank,
            access_count,
            last_access_filetime,
        });

        // Versions 4 and 6 both store two UTF-16 NUL terminators after the variable path.
        offset = path_end + 4;
    }

    Ok(DestList {
        version,
        declared_entry_count: entry_count,
        pinned_entry_count,
        last_entry_id,
        entries,
    })
}

fn resolve_path(cfb: &CompoundFile, stream_name: &str, raw_path: &str) -> String {
    if raw_path.starts_with("knownfolder:") {
        let link_path = cfb
            .stream(stream_name)
            .ok()
            .flatten()
            .and_then(|stream| parse_lnk_local_path(&stream));
        return link_path.unwrap_or_else(|| raw_path.to_string());
    }
    raw_path.to_string()
}

fn parse_lnk_local_path(data: &[u8]) -> Option<String> {
    if data.len() < 0x4c || read_u32(data, 0).ok()? != 0x4c {
        return None;
    }

    let flags = read_u32(data, 0x14).ok()?;
    let mut offset = 0x4cusize;

    if flags & 0x1 != 0 {
        let id_list_size = read_u16(data, offset).ok()? as usize;
        offset = offset.checked_add(2 + id_list_size)?;
    }

    if flags & 0x2 == 0 || offset + 28 > data.len() {
        return None;
    }

    let link_info_start = offset;
    let link_info_size = read_u32(data, link_info_start).ok()? as usize;
    if link_info_size < 28 || link_info_start + link_info_size > data.len() {
        return None;
    }

    let local_base_offset = read_u32(data, link_info_start + 16).ok()? as usize;
    let common_suffix_offset = read_u32(data, link_info_start + 24).ok()? as usize;

    let base = read_c_string(data, link_info_start + local_base_offset);
    let suffix = read_c_string(data, link_info_start + common_suffix_offset);

    match (base, suffix) {
        (Some(base), Some(suffix)) if !base.is_empty() && !suffix.is_empty() => {
            Some(join_windows_path(&base, &suffix))
        }
        (Some(base), _) if looks_like_windows_path(&base) => Some(base),
        (_, Some(suffix)) if looks_like_windows_path(&suffix) => Some(suffix),
        _ => None,
    }
}

fn join_windows_path(base: &str, suffix: &str) -> String {
    if looks_like_windows_path(suffix) {
        return suffix.to_string();
    }
    if base.ends_with('\\') || suffix.is_empty() {
        format!("{base}{suffix}")
    } else {
        format!("{base}\\{suffix}")
    }
}

fn looks_like_windows_path(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() >= 3 && bytes[1] == b':' && (bytes[2] == b'\\' || bytes[2] == b'/')
}

fn read_c_string(data: &[u8], offset: usize) -> Option<String> {
    if offset >= data.len() {
        return None;
    }
    let end = data[offset..].iter().position(|byte| *byte == 0)? + offset;
    Some(String::from_utf8_lossy(&data[offset..end]).to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_bytes_rejects_wrong_magic() {
        let data = vec![0u8; 512];
        let result = parse_bytes(data);
        assert!(result.is_err());
        match result.unwrap_err() {
            WincentError::DestListParse(msg) => {
                assert!(msg.contains("OLE Compound File Binary"));
            }
            other => panic!("unexpected error: {other}"),
        }
    }

    #[test]
    fn parse_bytes_rejects_truncated() {
        let data = vec![0u8; 100];
        let result = parse_bytes(data);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), WincentError::DestListParse(_)));
    }
}
