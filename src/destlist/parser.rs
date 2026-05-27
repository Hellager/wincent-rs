use std::fs;
use std::path::{Path, PathBuf};

use crate::error::WincentError;
use crate::utils::get_windows_recent_folder;
use crate::WincentResult;

use super::cfb::{decode_utf16_lossy, read_i32, read_u16, read_u32, read_u64, CompoundFile};

/// Explorer Recent Files automatic destination AppID hash.
pub const RECENT_FILES_APPID: &str = "5f7b5f1e01b83767.automaticDestinations-ms";
/// Explorer Frequent Folders automatic destination AppID hash.
pub const FREQUENT_FOLDERS_APPID: &str = "f01b4d95cf55d32a.automaticDestinations-ms";

/// Parsed `.automaticDestinations-ms` file.
#[derive(Debug, Clone, PartialEq)]
pub struct AutomaticDestinations {
    /// Compound File Binary container metadata.
    pub(crate) cfb_info: CfbInfo,
    /// Parsed DestList stream.
    pub(crate) dest_list: DestList,
}

impl AutomaticDestinations {
    /// Compound File Binary container metadata.
    #[must_use]
    pub fn cfb_info(&self) -> &CfbInfo {
        &self.cfb_info
    }

    /// Parsed DestList stream.
    #[must_use]
    pub fn dest_list(&self) -> &DestList {
        &self.dest_list
    }
}

/// CFB container metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CfbInfo {
    /// Regular sector size in bytes.
    pub(crate) sector_size: usize,
    /// Mini-sector size in bytes.
    pub(crate) mini_sector_size: usize,
    /// Stream-size threshold below which CFB mini streams are used.
    pub(crate) mini_cutoff_size: u32,
    /// Directory entries found in the CFB container.
    pub(crate) directory_entries: Vec<CfbDirectoryEntry>,
}

impl CfbInfo {
    /// Regular sector size in bytes.
    #[must_use]
    pub fn sector_size(&self) -> usize {
        self.sector_size
    }

    /// Mini-sector size in bytes.
    #[must_use]
    pub fn mini_sector_size(&self) -> usize {
        self.mini_sector_size
    }

    /// Stream-size threshold below which CFB mini streams are used.
    #[must_use]
    pub fn mini_cutoff_size(&self) -> u32 {
        self.mini_cutoff_size
    }

    /// Directory entries found in the CFB container.
    #[must_use]
    pub fn directory_entries(&self) -> &[CfbDirectoryEntry] {
        &self.directory_entries
    }
}

/// A single CFB directory entry (stream or storage).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CfbDirectoryEntry {
    /// Directory entry name.
    pub(crate) name: String,
    /// Raw CFB object type.
    pub(crate) object_type: u8,
    /// First sector of the entry stream.
    pub(crate) start_sector: u32,
    /// Stream size in bytes.
    pub(crate) stream_size: u64,
}

impl CfbDirectoryEntry {
    /// Directory entry name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Raw CFB object type.
    #[must_use]
    pub fn object_type(&self) -> u8 {
        self.object_type
    }

    /// First sector of the entry stream.
    #[must_use]
    pub fn start_sector(&self) -> u32 {
        self.start_sector
    }

    /// Stream size in bytes.
    #[must_use]
    pub fn stream_size(&self) -> u64 {
        self.stream_size
    }
}

/// Parsed DestList stream header + entries.
#[derive(Debug, Clone, PartialEq)]
pub struct DestList {
    /// DestList format version.
    pub(crate) version: u32,
    /// Entry count declared by the DestList header.
    pub(crate) declared_entry_count: usize,
    /// Number of pinned entries declared by the DestList header.
    pub(crate) pinned_entry_count: u32,
    /// Compatibility alias for [`DestList::last_entry_number`].
    pub(crate) last_entry_id: u64,
    /// Last entry number assigned by Explorer.
    pub(crate) last_entry_number: u32,
    /// Unknown header field adjacent to [`DestList::last_entry_number`].
    pub(crate) last_entry_number_unknown: u32,
    /// Last revision number assigned by Explorer.
    pub(crate) last_revision_number: u32,
    /// Unknown header field adjacent to [`DestList::last_revision_number`].
    pub(crate) last_revision_number_unknown: u32,
    /// Parsed DestList entries.
    pub(crate) entries: Vec<DestListEntry>,
}

impl DestList {
    /// DestList format version.
    #[must_use]
    pub fn version(&self) -> u32 {
        self.version
    }

    /// Entry count declared by the DestList header.
    #[must_use]
    pub fn declared_entry_count(&self) -> usize {
        self.declared_entry_count
    }

    /// Number of pinned entries declared by the DestList header.
    #[must_use]
    pub fn pinned_entry_count(&self) -> u32 {
        self.pinned_entry_count
    }

    /// Compatibility alias for [`DestList::last_entry_number`].
    #[must_use]
    pub fn last_entry_id(&self) -> u64 {
        self.last_entry_id
    }

    /// Last entry number assigned by Explorer.
    #[must_use]
    pub fn last_entry_number(&self) -> u32 {
        self.last_entry_number
    }

    /// Unknown header field adjacent to [`DestList::last_entry_number`].
    #[must_use]
    pub fn last_entry_number_unknown(&self) -> u32 {
        self.last_entry_number_unknown
    }

    /// Last revision number assigned by Explorer.
    #[must_use]
    pub fn last_revision_number(&self) -> u32 {
        self.last_revision_number
    }

    /// Unknown header field adjacent to [`DestList::last_revision_number`].
    #[must_use]
    pub fn last_revision_number_unknown(&self) -> u32 {
        self.last_revision_number_unknown
    }

    /// Parsed DestList entries.
    #[must_use]
    pub fn entries(&self) -> &[DestListEntry] {
        &self.entries
    }
}

/// A single DestList entry.
#[derive(Debug, Clone, PartialEq)]
pub struct DestListEntry {
    /// Byte offset of this entry inside the DestList stream.
    pub(crate) entry_offset: usize,
    /// Parsed byte length of this entry.
    pub(crate) entry_len: usize,
    /// Compatibility alias for [`DestListEntry::entry_number`].
    pub(crate) entry_id: u64,
    /// Explorer entry number.
    pub(crate) entry_number: u32,
    /// Unknown field adjacent to [`DestListEntry::entry_number`].
    pub(crate) entry_number_unknown: u32,
    /// CFB stream name containing the Shell Link payload for this entry.
    pub(crate) stream_name: String,
    /// Raw path as stored; may be `"knownfolder:{GUID}"`.
    pub(crate) raw_path: String,
    /// Resolved path (knownfolder GUIDs resolved via Shell Link stream).
    pub(crate) path: String,
    /// `-1` if not pinned.
    pub(crate) pin_status: i32,
    /// Pin order when the entry is pinned, if known.
    pub(crate) pin_order: Option<i32>,
    /// Compatibility alias for [`DestListEntry::recent_rank`].
    pub(crate) rank: i32,
    /// Recent rank reported by the DestList entry.
    pub(crate) recent_rank: i32,
    /// `0` means hidden in v4.
    pub(crate) count: u32,
    /// Access count reported by the DestList entry.
    pub(crate) access_count: u32,
    /// Explorer score value reported by the DestList entry.
    pub(crate) score: f32,
    /// Compatibility alias for [`DestListEntry::last_interaction_filetime`].
    pub(crate) last_access_filetime: Option<u64>,
    /// Last interaction timestamp as a raw Windows FILETIME value.
    pub(crate) last_interaction_filetime: Option<u64>,
    /// Serialized property-store size when present.
    pub(crate) sps_size: Option<u32>,
}

impl DestListEntry {
    /// Byte offset of this entry inside the DestList stream.
    #[must_use]
    pub fn entry_offset(&self) -> usize {
        self.entry_offset
    }

    /// Parsed byte length of this entry.
    #[must_use]
    pub fn entry_len(&self) -> usize {
        self.entry_len
    }

    /// Compatibility alias for [`DestListEntry::entry_number`].
    #[must_use]
    pub fn entry_id(&self) -> u64 {
        self.entry_id
    }

    /// Explorer entry number.
    #[must_use]
    pub fn entry_number(&self) -> u32 {
        self.entry_number
    }

    /// Unknown field adjacent to [`DestListEntry::entry_number`].
    #[must_use]
    pub fn entry_number_unknown(&self) -> u32 {
        self.entry_number_unknown
    }

    /// CFB stream name containing the Shell Link payload for this entry.
    #[must_use]
    pub fn stream_name(&self) -> &str {
        &self.stream_name
    }

    /// Raw path as stored; may be `"knownfolder:{GUID}"`.
    #[must_use]
    pub fn raw_path(&self) -> &str {
        &self.raw_path
    }

    /// Resolved path (knownfolder GUIDs resolved via Shell Link stream).
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Pin status; `-1` means not pinned.
    #[must_use]
    pub fn pin_status(&self) -> i32 {
        self.pin_status
    }

    /// Pin order when the entry is pinned, if known.
    #[must_use]
    pub fn pin_order(&self) -> Option<i32> {
        self.pin_order
    }

    /// Whether this entry is pinned.
    #[must_use]
    pub fn is_pinned(&self) -> bool {
        self.pin_order.is_some()
    }

    /// Compatibility alias for [`DestListEntry::recent_rank`].
    #[must_use]
    pub fn rank(&self) -> i32 {
        self.rank
    }

    /// Recent rank reported by the DestList entry.
    #[must_use]
    pub fn recent_rank(&self) -> i32 {
        self.recent_rank
    }

    /// Compatibility alias for [`DestListEntry::access_count`].
    #[must_use]
    pub fn count(&self) -> u32 {
        self.count
    }

    /// Access count reported by the DestList entry.
    #[must_use]
    pub fn access_count(&self) -> u32 {
        self.access_count
    }

    /// Explorer score value reported by the DestList entry.
    #[must_use]
    pub fn score(&self) -> f32 {
        self.score
    }

    /// Compatibility alias for [`DestListEntry::last_interaction_filetime`].
    #[must_use]
    pub fn last_access_filetime(&self) -> Option<u64> {
        self.last_access_filetime
    }

    /// Last interaction timestamp as a raw Windows FILETIME value.
    #[must_use]
    pub fn last_interaction_filetime(&self) -> Option<u64> {
        self.last_interaction_filetime
    }

    /// Serialized property-store size when present.
    #[must_use]
    pub fn sps_size(&self) -> Option<u32> {
        self.sps_size
    }
}

/// Returns all entries from a parsed DestList.
pub fn entries(dest_list: &DestList) -> Vec<DestListEntry> {
    dest_list.entries.clone()
}

/// Returns the path to the Explorer recent-files `.automaticDestinations-ms` file.
pub fn recent_files_dest_path() -> WincentResult<PathBuf> {
    Ok(PathBuf::from(get_windows_recent_folder()?)
        .join("AutomaticDestinations")
        .join(RECENT_FILES_APPID))
}

/// Returns the path to the Explorer frequent-folders `.automaticDestinations-ms` file.
pub fn frequent_folders_dest_path() -> WincentResult<PathBuf> {
    Ok(PathBuf::from(get_windows_recent_folder()?)
        .join("AutomaticDestinations")
        .join(FREQUENT_FOLDERS_APPID))
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
    Ok(AutomaticDestinations {
        cfb_info,
        dest_list,
    })
}

fn parse_dest_list(cfb: &CompoundFile) -> WincentResult<DestList> {
    let dest_list = cfb
        .stream("DestList")
        .map_err(WincentError::DestListParse)?
        .ok_or_else(|| WincentError::DestListParse("DestList stream not found".to_string()))?;

    if dest_list.len() < 32 {
        return Ok(DestList {
            version: 0,
            declared_entry_count: 0,
            pinned_entry_count: 0,
            last_entry_id: 0,
            last_entry_number: 0,
            last_entry_number_unknown: 0,
            last_revision_number: 0,
            last_revision_number_unknown: 0,
            entries: Vec::new(),
        });
    }

    let version = read_u32(&dest_list, 0).map_err(WincentError::DestListParse)?;
    if !matches!(version, 1 | 3 | 4 | 6) {
        return Err(WincentError::DestListUnsupportedVersion(version));
    }

    let entry_count = read_u32(&dest_list, 4).map_err(WincentError::DestListParse)? as usize;
    let pinned_entry_count = read_u32(&dest_list, 8).map_err(WincentError::DestListParse)?;
    let last_entry_number = read_u32(&dest_list, 0x10).map_err(WincentError::DestListParse)?;
    let last_entry_number_unknown =
        read_u32(&dest_list, 0x14).map_err(WincentError::DestListParse)?;
    let last_revision_number = read_u32(&dest_list, 0x18).map_err(WincentError::DestListParse)?;
    let last_revision_number_unknown =
        read_u32(&dest_list, 0x1c).map_err(WincentError::DestListParse)?;
    let mut offset = 32usize;
    let mut entries = Vec::new();

    for i in 0..entry_count {
        let Some(entry) = parse_dest_list_entry(cfb, &dest_list, version, offset)
            .map_err(WincentError::DestListParse)?
        else {
            if i == 0 {
                return Err(WincentError::DestListParse(format!(
                    "DestList truncated before first entry (declared {entry_count})"
                )));
            }
            break;
        };

        offset = offset.checked_add(entry.entry_len).ok_or_else(|| {
            WincentError::DestListParse("DestList entry offset overflow".to_string())
        })?;
        entries.push(entry);
    }

    Ok(DestList {
        version,
        declared_entry_count: entry_count,
        pinned_entry_count,
        last_entry_id: last_entry_number as u64,
        last_entry_number,
        last_entry_number_unknown,
        last_revision_number,
        last_revision_number_unknown,
        entries,
    })
}

fn parse_dest_list_entry(
    cfb: &CompoundFile,
    dest_list: &[u8],
    version: u32,
    offset: usize,
) -> Result<Option<DestListEntry>, String> {
    match version {
        1 => parse_dest_list_entry_v1(cfb, dest_list, offset),
        3 | 4 | 6 => parse_dest_list_entry_v2_or_later(cfb, dest_list, offset),
        _ => Err(format!("unsupported DestList version {version}")),
    }
}

fn parse_dest_list_entry_v1(
    cfb: &CompoundFile,
    dest_list: &[u8],
    offset: usize,
) -> Result<Option<DestListEntry>, String> {
    if offset + 0x72 > dest_list.len() {
        return Ok(None);
    }

    let entry_number = read_u32(dest_list, offset + 0x58)?;
    let entry_number_unknown = read_u32(dest_list, offset + 0x5c)?;
    let score = f32::from_bits(read_u32(dest_list, offset + 0x60)?);
    let last_interaction_filetime = read_u64(dest_list, offset + 0x64).ok();
    let pin_status = read_i32(dest_list, offset + 0x6c)?;
    let path_chars = read_u16(dest_list, offset + 0x70)? as usize;
    let path_start = offset + 0x72;
    let path_end = path_start
        .checked_add(path_chars.saturating_mul(2))
        .ok_or_else(|| "DestList v1 path size overflow".to_string())?;
    if path_end > dest_list.len() {
        return Ok(None);
    }

    Ok(Some(build_entry(
        cfb,
        offset,
        path_end - offset,
        entry_number,
        entry_number_unknown,
        &dest_list[path_start..path_end],
        pin_status,
        -1,
        0,
        score,
        last_interaction_filetime,
        None,
    )))
}

fn parse_dest_list_entry_v2_or_later(
    cfb: &CompoundFile,
    dest_list: &[u8],
    offset: usize,
) -> Result<Option<DestListEntry>, String> {
    if offset + 0x82 > dest_list.len() {
        return Ok(None);
    }

    let entry_number = read_u32(dest_list, offset + 0x58)?;
    let entry_number_unknown = read_u32(dest_list, offset + 0x5c)?;
    let score = f32::from_bits(read_u32(dest_list, offset + 0x60)?);
    let last_interaction_filetime = read_u64(dest_list, offset + 0x64).ok();
    let pin_status = read_i32(dest_list, offset + 0x6c)?;
    let recent_rank = read_i32(dest_list, offset + 0x70)?;
    let access_count = read_u32(dest_list, offset + 0x74)?;
    let path_chars = read_u16(dest_list, offset + 0x80)? as usize;
    let path_start = offset + 0x82;
    let path_end = path_start
        .checked_add(path_chars.saturating_mul(2))
        .ok_or_else(|| "DestList path size overflow".to_string())?;
    if path_end > dest_list.len() || path_end + 4 > dest_list.len() {
        return Ok(None);
    }

    let sps_size = read_u32(dest_list, path_end)?;
    let entry_end = path_end
        .checked_add(4)
        .and_then(|value| value.checked_add(sps_size as usize))
        .ok_or_else(|| "DestList entry size overflow".to_string())?;
    if entry_end > dest_list.len() {
        return Ok(None);
    }

    Ok(Some(build_entry(
        cfb,
        offset,
        entry_end - offset,
        entry_number,
        entry_number_unknown,
        &dest_list[path_start..path_end],
        pin_status,
        recent_rank,
        access_count,
        score,
        last_interaction_filetime,
        Some(sps_size),
    )))
}

#[allow(clippy::too_many_arguments)]
fn build_entry(
    cfb: &CompoundFile,
    entry_offset: usize,
    entry_len: usize,
    entry_number: u32,
    entry_number_unknown: u32,
    raw_path_bytes: &[u8],
    pin_status: i32,
    recent_rank: i32,
    access_count: u32,
    score: f32,
    last_interaction_filetime: Option<u64>,
    sps_size: Option<u32>,
) -> DestListEntry {
    let raw_path = decode_utf16_lossy(raw_path_bytes);
    let stream_name = format!("{entry_number:x}");
    let resolved_path = resolve_path(cfb, &stream_name, &raw_path);

    DestListEntry {
        entry_offset,
        entry_len,
        entry_id: entry_number as u64,
        entry_number,
        entry_number_unknown,
        stream_name,
        raw_path,
        path: resolved_path,
        pin_status,
        pin_order: (pin_status >= 0).then_some(pin_status),
        rank: recent_rank,
        recent_rank,
        count: access_count,
        access_count,
        score,
        last_access_filetime: last_interaction_filetime,
        last_interaction_filetime,
        sps_size,
    }
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

pub(super) fn parse_lnk_local_path(data: &[u8]) -> Option<String> {
    if data.len() < 0x4c || read_u32(data, 0).ok()? != 0x4c {
        return None;
    }

    let flags = read_u32(data, 0x14).ok()?;
    let mut offset = 0x4cusize;

    if flags & 0x1 != 0 {
        let id_list_size = read_u16(data, offset).ok()? as usize;
        offset = offset.checked_add(2)?.checked_add(id_list_size)?;
    }

    if flags & 0x2 == 0 || offset.checked_add(28)? > data.len() {
        return None;
    }

    let link_info_start = offset;
    let link_info_size = read_u32(data, link_info_start).ok()? as usize;
    let link_info_header_size = read_u32(data, link_info_start + 4).ok()? as usize;
    let link_info_end = link_info_start.checked_add(link_info_size)?;
    if link_info_size < 28 || link_info_end > data.len() {
        return None;
    }

    let local_base_offset = read_u32(data, link_info_start + 16).ok()? as usize;
    let common_suffix_offset = read_u32(data, link_info_start + 24).ok()? as usize;
    let local_base_unicode_offset =
        if link_info_header_size >= 0x24 && link_info_start.checked_add(32)? <= data.len() {
            read_u32(data, link_info_start + 28).ok()? as usize
        } else {
            0
        };
    let common_suffix_unicode_offset =
        if link_info_header_size >= 0x24 && link_info_start.checked_add(36)? <= data.len() {
            read_u32(data, link_info_start + 32).ok()? as usize
        } else {
            0
        };

    let base = read_utf16_z_string_in_link_info(
        data,
        link_info_start,
        link_info_size,
        local_base_unicode_offset,
    )
    .or_else(|| {
        read_c_string_in_link_info(data, link_info_start, link_info_size, local_base_offset)
    });
    let suffix = read_utf16_z_string_in_link_info(
        data,
        link_info_start,
        link_info_size,
        common_suffix_unicode_offset,
    )
    .or_else(|| {
        read_c_string_in_link_info(data, link_info_start, link_info_size, common_suffix_offset)
    });

    match (base, suffix) {
        (Some(base), Some(suffix)) if !base.is_empty() && !suffix.is_empty() => {
            Some(join_windows_path(&base, &suffix))
        }
        (Some(base), _) if looks_like_windows_path(&base) => Some(base),
        (_, Some(suffix)) if looks_like_windows_path(&suffix) => Some(suffix),
        _ => None,
    }
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

    if end == offset || end + 1 >= data.len() {
        return None;
    }

    Some(decode_utf16_lossy(&data[offset..end]))
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
        assert!(matches!(
            result.unwrap_err(),
            WincentError::DestListParse(_)
        ));
    }

    #[test]
    #[ignore = "Integration test; reads system Jump List file — run with: cargo test --features destlist parse_recent_files_metadata -- --ignored --nocapture"]
    fn parse_recent_files_metadata() {
        let path = recent_files_dest_path().expect("failed to get recent files dest path");
        println!("Path: {}", path.display());

        let parsed = parse_file(&path).expect("failed to parse recent files dest");
        let all = entries(parsed.dest_list());

        println!(
            "version={} declared={} pinned={} last_entry_id={:#x}",
            parsed.dest_list().version(),
            parsed.dest_list().declared_entry_count(),
            parsed.dest_list().pinned_entry_count(),
            parsed.dest_list().last_entry_id(),
        );
        println!("entries ({}):", all.len());
        for e in &all {
            println!(
                "  id={:#x} pin={} rank={} count={} score={:.2} path={}",
                e.entry_id(),
                e.pin_status(),
                e.rank(),
                e.count(),
                e.score(),
                e.path()
            );
        }
    }

    #[test]
    #[ignore = "Integration test; reads system Jump List file — run with: cargo test --features destlist parse_frequent_folders_metadata -- --ignored --nocapture"]
    fn parse_frequent_folders_metadata() {
        let path = frequent_folders_dest_path().expect("failed to get frequent folders dest path");
        println!("Path: {}", path.display());

        let parsed = parse_file(&path).expect("failed to parse frequent folders dest");
        let all = entries(parsed.dest_list());

        println!(
            "version={} declared={} pinned={} last_entry_id={:#x}",
            parsed.dest_list().version(),
            parsed.dest_list().declared_entry_count(),
            parsed.dest_list().pinned_entry_count(),
            parsed.dest_list().last_entry_id(),
        );
        println!("entries ({}):", all.len());
        for e in &all {
            println!(
                "  id={:#x} pin={} rank={} count={} score={:.2} path={}",
                e.entry_id(),
                e.pin_status(),
                e.rank(),
                e.count(),
                e.score(),
                e.path()
            );
        }
    }
}
