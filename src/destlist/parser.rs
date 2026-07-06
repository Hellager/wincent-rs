use std::collections::HashSet;
use std::fs;
use std::path::{Path, PathBuf};

use crate::error::WincentError;
use crate::utils::{get_windows_recent_folder, normalize_path_lightweight, paths_equal};
use crate::WincentResult;

use super::cfb::{decode_utf16_lossy, read_i32, read_u16, read_u32, read_u64, CompoundFile};

/// Explorer Recent Files automatic destination AppID hash.
pub const RECENT_FILES_APPID: &str = "5f7b5f1e01b83767.automaticDestinations-ms";
/// Explorer Frequent Folders automatic destination AppID hash.
pub const FREQUENT_FOLDERS_APPID: &str = "f01b4d95cf55d32a.automaticDestinations-ms";

/// Pin state for a path in Explorer's Frequent Folders DestList.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FrequentFolderPinStatus {
    /// A matching Frequent Folders entry is pinned.
    Pinned,
    /// Matching Frequent Folders entries exist, but none are pinned.
    Unpinned,
    /// No matching Frequent Folders entry was found.
    NotFound,
}

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

/// Severity for non-fatal DestList parse diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DiagnosticSeverity {
    /// Informational parser note.
    Info,
    /// Non-fatal parse issue; hard failures are returned as [`WincentError`].
    Warning,
}

/// Non-fatal parse diagnostic.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Diagnostic {
    severity: DiagnosticSeverity,
    context: String,
    message: String,
}

impl Diagnostic {
    /// Creates an informational diagnostic.
    #[must_use]
    pub fn info(context: impl Into<String>, message: impl Into<String>) -> Self {
        Self {
            severity: DiagnosticSeverity::Info,
            context: context.into(),
            message: message.into(),
        }
    }

    /// Creates a warning diagnostic.
    #[must_use]
    pub fn warning(context: impl Into<String>, message: impl Into<String>) -> Self {
        Self {
            severity: DiagnosticSeverity::Warning,
            context: context.into(),
            message: message.into(),
        }
    }

    /// Diagnostic severity.
    #[must_use]
    pub fn severity(&self) -> DiagnosticSeverity {
        self.severity
    }

    /// Parser context that emitted this diagnostic.
    #[must_use]
    pub fn context(&self) -> &str {
        &self.context
    }

    /// Human-readable diagnostic message.
    #[must_use]
    pub fn message(&self) -> &str {
        &self.message
    }
}

/// Origin of a path candidate observed while parsing a DestList entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PathSource {
    source: String,
    value: String,
}

impl PathSource {
    fn new(source: impl Into<String>, value: impl Into<String>) -> Self {
        Self {
            source: source.into(),
            value: value.into(),
        }
    }

    /// Name of the source that produced this path candidate.
    #[must_use]
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Path candidate value.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
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
    /// Raw counter at header offset `0x0c`.
    pub(crate) header_counter_raw: u32,
    /// Raw counter interpreted as `f32`.
    pub(crate) header_counter_f32: f32,
    /// Last entry id assigned by Explorer.
    pub(crate) last_entry_id: u64,
    /// Low 32 bits of [`DestList::last_entry_id`].
    pub(crate) last_entry_number: u32,
    /// Add/delete action count stored at header offset `0x18`.
    pub(crate) add_delete_action_count: u64,
    /// Parsed DestList entries.
    pub(crate) entries: Vec<DestListEntry>,
    /// Non-fatal parse diagnostics.
    pub(crate) diagnostics: Vec<Diagnostic>,
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

    /// Raw counter at header offset `0x0c`.
    #[must_use]
    pub fn header_counter_raw(&self) -> u32 {
        self.header_counter_raw
    }

    /// Raw counter at header offset `0x0c` interpreted as `f32`.
    #[must_use]
    pub fn header_counter_f32(&self) -> f32 {
        self.header_counter_f32
    }

    /// Last entry id assigned by Explorer.
    #[must_use]
    pub fn last_entry_id(&self) -> u64 {
        self.last_entry_id
    }

    /// Low 32 bits of [`DestList::last_entry_id`].
    ///
    /// Explorer uses this value as the hexadecimal CFB stream name for the
    /// entry Shell Link payload.
    #[must_use]
    pub fn last_entry_number(&self) -> u32 {
        self.last_entry_number
    }

    /// Add/delete action count stored at header offset `0x18`.
    #[must_use]
    pub fn add_delete_action_count(&self) -> u64 {
        self.add_delete_action_count
    }

    /// Parsed DestList entries.
    #[must_use]
    pub fn entries(&self) -> &[DestListEntry] {
        &self.entries
    }

    /// Non-fatal parse diagnostics.
    #[must_use]
    pub fn diagnostics(&self) -> &[Diagnostic] {
        &self.diagnostics
    }
}

/// A single DestList entry.
#[derive(Debug, Clone, PartialEq)]
pub struct DestListEntry {
    /// Byte offset of this entry inside the DestList stream.
    pub(crate) entry_offset: usize,
    /// Parsed byte length of this entry.
    pub(crate) entry_len: usize,
    /// Physical entry position in the DestList stream.
    pub(crate) mru_position: usize,
    /// Checksum or unknown signed field at entry offset `0x00`.
    pub(crate) checksum: i64,
    /// Full Explorer entry id at entry offset `0x58`.
    pub(crate) entry_id: u64,
    /// Low 32 bits of [`DestListEntry::entry_id`].
    pub(crate) entry_number: u32,
    /// High 32 bits of [`DestListEntry::entry_id`].
    pub(crate) entry_number_unknown: u32,
    /// Hostname recorded in the DestList entry.
    pub(crate) hostname: String,
    /// Volume Droid GUID.
    pub(crate) volume_droid: String,
    /// File Droid GUID.
    pub(crate) file_droid: String,
    /// Volume birth Droid GUID.
    pub(crate) volume_birth_droid: String,
    /// File birth Droid GUID.
    pub(crate) file_birth_droid: String,
    /// MAC address embedded in the file Droid GUID.
    pub(crate) file_droid_mac: String,
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
    /// Reserved field at entry offset `0x78` for v3/v4/v6.
    pub(crate) reserved_78: Option<u32>,
    /// Reserved field at entry offset `0x7c` for v3/v4/v6.
    pub(crate) reserved_7c: Option<u32>,
    /// Path candidates observed while resolving this entry.
    pub(crate) path_sources: Vec<PathSource>,
    /// Non-fatal entry-specific parse diagnostics.
    pub(crate) warnings: Vec<Diagnostic>,
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

    /// Physical entry position in the DestList stream.
    #[must_use]
    pub fn mru_position(&self) -> usize {
        self.mru_position
    }

    /// Checksum or unknown signed field at entry offset `0x00`.
    #[must_use]
    pub fn checksum(&self) -> i64 {
        self.checksum
    }

    /// Full Explorer entry id at entry offset `0x58`.
    #[must_use]
    pub fn entry_id(&self) -> u64 {
        self.entry_id
    }

    /// Low 32 bits of [`DestListEntry::entry_id`].
    #[must_use]
    pub fn entry_number(&self) -> u32 {
        self.entry_number
    }

    /// High 32 bits of [`DestListEntry::entry_id`].
    #[must_use]
    pub fn entry_number_unknown(&self) -> u32 {
        self.entry_number_unknown
    }

    /// Hostname recorded in the DestList entry.
    #[must_use]
    pub fn hostname(&self) -> &str {
        &self.hostname
    }

    /// Volume Droid GUID.
    #[must_use]
    pub fn volume_droid(&self) -> &str {
        &self.volume_droid
    }

    /// File Droid GUID.
    #[must_use]
    pub fn file_droid(&self) -> &str {
        &self.file_droid
    }

    /// Volume birth Droid GUID.
    #[must_use]
    pub fn volume_birth_droid(&self) -> &str {
        &self.volume_birth_droid
    }

    /// File birth Droid GUID.
    #[must_use]
    pub fn file_birth_droid(&self) -> &str {
        &self.file_birth_droid
    }

    /// MAC address embedded in the file Droid GUID.
    #[must_use]
    pub fn file_droid_mac(&self) -> &str {
        &self.file_droid_mac
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

    /// Reserved field at entry offset `0x78` for v3/v4/v6.
    #[must_use]
    pub fn reserved_78(&self) -> Option<u32> {
        self.reserved_78
    }

    /// Reserved field at entry offset `0x7c` for v3/v4/v6.
    #[must_use]
    pub fn reserved_7c(&self) -> Option<u32> {
        self.reserved_7c
    }

    /// Path candidates observed while resolving this entry.
    #[must_use]
    pub fn path_sources(&self) -> &[PathSource] {
        &self.path_sources
    }

    /// Non-fatal entry-specific parse diagnostics.
    #[must_use]
    pub fn warnings(&self) -> &[Diagnostic] {
        &self.warnings
    }
}

/// Returns all entries from a parsed DestList.
///
/// This clones the parsed entries. Use [`DestList::entries`] when borrowing is
/// enough.
///
/// # Examples
///
/// ```rust,no_run
/// # fn main() -> wincent::WincentResult<()> {
/// use wincent::destlist::{entries, parse_file, recent_files_dest_path};
///
/// let parsed = parse_file(recent_files_dest_path()?)?;
/// for entry in entries(parsed.dest_list()) {
///     println!("{}", entry.path());
/// }
/// Ok(())
/// # }
/// ```
pub fn entries(dest_list: &DestList) -> Vec<DestListEntry> {
    dest_list.entries.clone()
}

/// Returns entries that are likely visible in Explorer Quick Access.
///
/// Uses Explorer-oriented heuristics for DestList v4 and v6. The
/// `normal_slot_count` controls how many non-pinned v6 normal entries are
/// considered; Explorer commonly uses 4.
#[must_use]
pub fn quick_access_entries(dest_list: &DestList, normal_slot_count: i32) -> Vec<DestListEntry> {
    match dest_list.version {
        4 => quick_access_entries_v4(&dest_list.entries),
        6 => quick_access_entries_v6(&dest_list.entries, normal_slot_count),
        _ => dest_list.entries.clone(),
    }
}

/// Returns entries likely visible in Explorer Quick Access using 4 normal slots.
#[must_use]
pub fn visible_entries(dest_list: &DestList) -> Vec<DestListEntry> {
    quick_access_entries(dest_list, 4)
}

pub(crate) fn frequent_folder_pin_status(path: &str) -> WincentResult<FrequentFolderPinStatus> {
    let parsed = parse_file(frequent_folders_dest_path()?)?;
    Ok(frequent_folder_pin_status_from_entries(
        path,
        parsed.dest_list().entries(),
    ))
}

pub(crate) fn frequent_folder_pin_status_from_entries(
    path: &str,
    entries: &[DestListEntry],
) -> FrequentFolderPinStatus {
    let mut matched = false;

    for entry in entries
        .iter()
        .filter(|entry| paths_equal(entry.path(), path))
    {
        matched = true;
        if entry.is_pinned() {
            return FrequentFolderPinStatus::Pinned;
        }
    }

    if matched {
        FrequentFolderPinStatus::Unpinned
    } else {
        FrequentFolderPinStatus::NotFound
    }
}

fn quick_access_entries_v6(
    entries: &[DestListEntry],
    normal_slot_count: i32,
) -> Vec<DestListEntry> {
    let mut pinned: Vec<_> = entries
        .iter()
        .filter(|entry| entry.pin_order.is_some())
        .cloned()
        .collect();
    pinned.sort_by_key(|entry| entry.pin_order.unwrap_or(i32::MAX));

    let mut used_paths: HashSet<String> = pinned
        .iter()
        .map(|entry| visible_entry_path_key(&entry.path))
        .collect();

    let mut normal: Vec<_> = entries
        .iter()
        .filter(|entry| {
            entry.pin_order.is_none()
                && entry.recent_rank >= 0
                && entry.recent_rank < normal_slot_count
        })
        .cloned()
        .collect();
    // Mirrors the observed v6 ordering in the reference parser. Validate with
    // real Explorer v6 samples before changing this direction.
    normal.sort_by_key(|entry| std::cmp::Reverse(entry.recent_rank));
    normal.retain(|entry| used_paths.insert(visible_entry_path_key(&entry.path)));

    pinned.extend(normal);
    pinned
}

fn quick_access_entries_v4(entries: &[DestListEntry]) -> Vec<DestListEntry> {
    if entries.iter().any(|entry| entry.pin_order.is_some()) {
        frequent_folder_entries_v4(entries)
    } else {
        recent_file_entries_v4(entries)
    }
}

fn frequent_folder_entries_v4(entries: &[DestListEntry]) -> Vec<DestListEntry> {
    let mut pinned: Vec<_> = entries
        .iter()
        .filter(|entry| entry.pin_order.is_some())
        .cloned()
        .collect();
    pinned.sort_by_key(|entry| entry.pin_order.unwrap_or(i32::MAX));

    let mut used_paths: HashSet<String> = pinned
        .iter()
        .map(|entry| visible_entry_path_key(&entry.path))
        .collect();

    let mut frequent_candidates: Vec<_> = entries
        .iter()
        .filter(|entry| {
            entry.pin_order.is_none() && entry.access_count > 1 && entry.recent_rank >= 0
        })
        .cloned()
        .collect();
    frequent_candidates.sort_by(|left, right| {
        left.recent_rank
            .cmp(&right.recent_rank)
            .then_with(|| {
                right
                    .last_interaction_filetime
                    .cmp(&left.last_interaction_filetime)
            })
            .then_with(|| right.entry_number.cmp(&left.entry_number))
    });
    frequent_candidates.retain(|entry| used_paths.insert(visible_entry_path_key(&entry.path)));
    frequent_candidates.truncate(4);

    pinned.extend(frequent_candidates);
    pinned
}

fn recent_file_entries_v4(entries: &[DestListEntry]) -> Vec<DestListEntry> {
    let mut visible = Vec::new();
    let mut used_paths = HashSet::new();
    let mut jump_list_backing_files = Vec::new();

    for entry in entries {
        if entry.access_count == 0 {
            continue;
        }

        if is_automatic_destinations_path(&entry.path) {
            jump_list_backing_files.push(entry.clone());
            continue;
        }

        if used_paths.insert(visible_entry_path_key(&entry.path)) {
            visible.push(entry.clone());
        }
    }

    if let Some(best_backing_file) = jump_list_backing_files.into_iter().max_by(|left, right| {
        left.access_count
            .cmp(&right.access_count)
            .then_with(|| {
                left.last_interaction_filetime
                    .cmp(&right.last_interaction_filetime)
            })
            .then_with(|| left.entry_number.cmp(&right.entry_number))
    }) {
        if used_paths.insert(visible_entry_path_key(&best_backing_file.path)) {
            visible.push(best_backing_file);
        }
    }

    visible
}

fn is_automatic_destinations_path(path: &str) -> bool {
    path.to_ascii_lowercase()
        .ends_with(".automaticdestinations-ms")
}

fn visible_entry_path_key(path: &str) -> String {
    normalize_path_lightweight(path)
}

/// Returns the path to the Explorer recent-files `.automaticDestinations-ms` file.
///
/// # Errors
///
/// Returns an error if the current user's Windows Recent folder cannot be
/// resolved.
pub fn recent_files_dest_path() -> WincentResult<PathBuf> {
    Ok(PathBuf::from(get_windows_recent_folder()?)
        .join("AutomaticDestinations")
        .join(RECENT_FILES_APPID))
}

/// Returns the path to the Explorer frequent-folders `.automaticDestinations-ms` file.
///
/// # Errors
///
/// Returns an error if the current user's Windows Recent folder cannot be
/// resolved.
pub fn frequent_folders_dest_path() -> WincentResult<PathBuf> {
    Ok(PathBuf::from(get_windows_recent_folder()?)
        .join("AutomaticDestinations")
        .join(FREQUENT_FOLDERS_APPID))
}

/// Parses an `.automaticDestinations-ms` file from disk.
///
/// # Errors
///
/// Returns [`WincentError::Io`] if the file cannot be read,
/// [`WincentError::DestListParse`] if the Compound File Binary container or
/// DestList stream is malformed, or
/// [`WincentError::DestListUnsupportedVersion`] if Explorer uses a DestList
/// format version this crate does not yet support.
///
/// # Examples
///
/// ```rust,no_run
/// # fn main() -> wincent::WincentResult<()> {
/// use wincent::destlist::{parse_file, recent_files_dest_path};
///
/// let parsed = parse_file(recent_files_dest_path()?)?;
/// println!("DestList version {}", parsed.dest_list().version());
/// Ok(())
/// # }
/// ```
pub fn parse_file(path: impl AsRef<Path>) -> WincentResult<AutomaticDestinations> {
    let path = path.as_ref();
    let data = fs::read(path).map_err(WincentError::Io)?;
    parse_bytes(data)
}

/// Parses an `.automaticDestinations-ms` file from an in-memory buffer.
///
/// # Errors
///
/// Returns [`WincentError::DestListParse`] if the bytes are not a supported CFB
/// container with a readable DestList stream, or
/// [`WincentError::DestListUnsupportedVersion`] if the DestList version is not
/// supported.
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
            header_counter_raw: 0,
            header_counter_f32: 0.0,
            last_entry_id: 0,
            last_entry_number: 0,
            add_delete_action_count: 0,
            entries: Vec::new(),
            diagnostics: vec![Diagnostic::warning(
                "destlist",
                format!("DestList stream is too small: {} bytes", dest_list.len()),
            )],
        });
    }

    let version = read_u32(&dest_list, 0).map_err(WincentError::DestListParse)?;
    if !matches!(version, 1 | 3 | 4 | 6) {
        return Err(WincentError::DestListUnsupportedVersion(version));
    }

    let entry_count = read_u32(&dest_list, 4).map_err(WincentError::DestListParse)? as usize;
    let pinned_entry_count = read_u32(&dest_list, 8).map_err(WincentError::DestListParse)?;
    let header_counter_raw = read_u32(&dest_list, 0x0c).map_err(WincentError::DestListParse)?;
    let header_counter_f32 = f32::from_bits(header_counter_raw);
    let last_entry_id = read_u64(&dest_list, 0x10).map_err(WincentError::DestListParse)?;
    let last_entry_number = last_entry_id as u32;
    let add_delete_action_count =
        read_u64(&dest_list, 0x18).map_err(WincentError::DestListParse)?;
    let mut offset = 32usize;
    let mut entries = Vec::new();
    let mut diagnostics = Vec::new();

    for i in 0..entry_count {
        let Some(entry) = parse_dest_list_entry(cfb, &dest_list, version, i, offset)
            .map_err(WincentError::DestListParse)?
        else {
            if i == 0 {
                return Err(WincentError::DestListParse(format!(
                    "DestList truncated before first entry (declared {entry_count})"
                )));
            }
            diagnostics.push(Diagnostic::warning(
                "destlist.entry",
                format!(
                    "stopped parsing at entry {i}, offset 0x{offset:x}; declared {entry_count}, parsed {}",
                    entries.len()
                ),
            ));
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
        header_counter_raw,
        header_counter_f32,
        last_entry_id,
        last_entry_number,
        add_delete_action_count,
        entries,
        diagnostics,
    })
}

fn parse_dest_list_entry(
    cfb: &CompoundFile,
    dest_list: &[u8],
    version: u32,
    mru_position: usize,
    offset: usize,
) -> Result<Option<DestListEntry>, String> {
    match version {
        1 => parse_dest_list_entry_v1(cfb, dest_list, mru_position, offset),
        3 | 4 | 6 => parse_dest_list_entry_v2_or_later(cfb, dest_list, mru_position, offset),
        _ => Err(format!("unsupported DestList version {version}")),
    }
}

fn parse_dest_list_entry_v1(
    cfb: &CompoundFile,
    dest_list: &[u8],
    mru_position: usize,
    offset: usize,
) -> Result<Option<DestListEntry>, String> {
    if offset + 0x72 > dest_list.len() {
        return Ok(None);
    }

    let entry_id = read_u64(dest_list, offset + 0x58)?;
    let entry_number = entry_id as u32;
    let entry_number_unknown = (entry_id >> 32) as u32;
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
        dest_list,
        offset,
        path_end - offset,
        mru_position,
        entry_id,
        entry_number,
        entry_number_unknown,
        &dest_list[path_start..path_end],
        pin_status,
        -1,
        0,
        score,
        last_interaction_filetime,
        None,
        None,
        None,
    )))
}

fn parse_dest_list_entry_v2_or_later(
    cfb: &CompoundFile,
    dest_list: &[u8],
    mru_position: usize,
    offset: usize,
) -> Result<Option<DestListEntry>, String> {
    if offset + 0x82 > dest_list.len() {
        return Ok(None);
    }

    let entry_id = read_u64(dest_list, offset + 0x58)?;
    let entry_number = entry_id as u32;
    let entry_number_unknown = (entry_id >> 32) as u32;
    let score = f32::from_bits(read_u32(dest_list, offset + 0x60)?);
    let last_interaction_filetime = read_u64(dest_list, offset + 0x64).ok();
    let pin_status = read_i32(dest_list, offset + 0x6c)?;
    let recent_rank = read_i32(dest_list, offset + 0x70)?;
    let access_count = read_u32(dest_list, offset + 0x74)?;
    let reserved_78 = read_u32(dest_list, offset + 0x78)?;
    let reserved_7c = read_u32(dest_list, offset + 0x7c)?;
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
        dest_list,
        offset,
        entry_end - offset,
        mru_position,
        entry_id,
        entry_number,
        entry_number_unknown,
        &dest_list[path_start..path_end],
        pin_status,
        recent_rank,
        access_count,
        score,
        last_interaction_filetime,
        Some(sps_size),
        Some(reserved_78),
        Some(reserved_7c),
    )))
}

#[allow(clippy::too_many_arguments)]
fn build_entry(
    cfb: &CompoundFile,
    dest_list: &[u8],
    entry_offset: usize,
    entry_len: usize,
    mru_position: usize,
    entry_id: u64,
    entry_number: u32,
    entry_number_unknown: u32,
    raw_path_bytes: &[u8],
    pin_status: i32,
    recent_rank: i32,
    access_count: u32,
    score: f32,
    last_interaction_filetime: Option<u64>,
    sps_size: Option<u32>,
    reserved_78: Option<u32>,
    reserved_7c: Option<u32>,
) -> DestListEntry {
    let raw_path = decode_utf16_lossy(raw_path_bytes);
    let stream_name = format!("{entry_number:x}");
    let resolved = resolve_path(cfb, &stream_name, &raw_path);

    DestListEntry {
        entry_offset,
        entry_len,
        mru_position,
        checksum: read_i64(dest_list, entry_offset).unwrap_or_default(),
        entry_id,
        entry_number,
        entry_number_unknown,
        hostname: decode_hostname(&dest_list[entry_offset + 0x48..entry_offset + 0x58]),
        volume_droid: format_guid_from_le_bytes(
            &dest_list[entry_offset + 0x08..entry_offset + 0x18],
        ),
        file_droid: format_guid_from_le_bytes(&dest_list[entry_offset + 0x18..entry_offset + 0x28]),
        volume_birth_droid: format_guid_from_le_bytes(
            &dest_list[entry_offset + 0x28..entry_offset + 0x38],
        ),
        file_birth_droid: format_guid_from_le_bytes(
            &dest_list[entry_offset + 0x38..entry_offset + 0x48],
        ),
        file_droid_mac: mac_from_droid_bytes(&dest_list[entry_offset + 0x18..entry_offset + 0x28]),
        stream_name,
        raw_path,
        path: resolved.best_path,
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
        reserved_78,
        reserved_7c,
        path_sources: resolved.path_sources,
        warnings: resolved.warnings,
    }
}

struct ResolvedPath {
    best_path: String,
    path_sources: Vec<PathSource>,
    warnings: Vec<Diagnostic>,
}

#[derive(Debug)]
struct ShellLinkSummary {
    local_path: Option<String>,
    network_path: Option<String>,
    relative_path: Option<String>,
    warnings: Vec<Diagnostic>,
}

#[derive(Debug)]
struct LinkInfoSummary {
    size: usize,
    local_path: Option<String>,
    network_path: Option<String>,
}

#[derive(Debug)]
struct NetworkLinkSummary {
    net_name: Option<String>,
}

fn resolve_path(cfb: &CompoundFile, stream_name: &str, raw_path: &str) -> ResolvedPath {
    let mut path_sources = vec![PathSource::new("destlist.raw_path", raw_path)];
    let mut warnings = Vec::new();

    let needs_link_resolution = starts_with_knownfolder(raw_path) || raw_path.starts_with("::");
    if !needs_link_resolution {
        return ResolvedPath {
            best_path: raw_path.to_string(),
            path_sources,
            warnings,
        };
    }

    let linked_lnk = match cfb.stream(stream_name) {
        Ok(Some(stream)) => parse_shell_link_summary(&stream),
        Ok(None) => {
            warnings.push(Diagnostic::warning(
                "destlist.link",
                format!("missing Shell Link stream `{stream_name}`"),
            ));
            None
        }
        Err(error) => {
            warnings.push(Diagnostic::warning(
                "destlist.link",
                format!("failed to read Shell Link stream `{stream_name}`: {error}"),
            ));
            None
        }
    };

    let mut best_path = None;
    if let Some(link) = linked_lnk {
        warnings.extend(link.warnings);
        if let Some(path) = link.local_path {
            path_sources.push(PathSource::new("lnk.linkinfo.local", path.clone()));
            best_path = Some(path);
        }
        if let Some(path) = link.network_path {
            path_sources.push(PathSource::new("lnk.linkinfo.network", path.clone()));
            if best_path.is_none() {
                best_path = Some(path);
            }
        }
        if let Some(path) = link.relative_path {
            path_sources.push(PathSource::new(
                "lnk.stringdata.relative_path",
                path.clone(),
            ));
            if best_path.is_none() {
                best_path = Some(path);
            }
        }
    }

    if best_path.is_none() {
        warnings.push(Diagnostic::warning(
            "destlist.path",
            format!("could not resolve Shell Link path for `{raw_path}`"),
        ));
    }

    ResolvedPath {
        best_path: best_path.unwrap_or_else(|| raw_path.to_string()),
        path_sources,
        warnings,
    }
}

pub(crate) fn starts_with_knownfolder(raw_path: &str) -> bool {
    raw_path
        .get(..raw_path.len().min("knownfolder:".len()))
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("knownfolder:"))
}

#[cfg(test)]
fn parse_lnk_local_path(data: &[u8]) -> Option<String> {
    let summary = parse_shell_link_summary(data)?;
    summary.local_path.or(summary.network_path)
}

fn parse_shell_link_summary(data: &[u8]) -> Option<ShellLinkSummary> {
    if !looks_like_lnk_stream(data) {
        return None;
    }

    let mut warnings = Vec::new();
    let flags = read_u32(data, 0x14).ok()?;
    let mut offset = 0x4cusize;

    if flags & 0x1 != 0 {
        match read_u16(data, offset) {
            Ok(id_list_size) => {
                let next = offset.checked_add(2)?.checked_add(id_list_size as usize)?;
                if next > data.len() {
                    warnings.push(Diagnostic::warning(
                        "lnk.idlist",
                        "IDList extends past stream end",
                    ));
                    offset = data.len();
                } else {
                    offset = next;
                }
            }
            Err(error) => {
                warnings.push(Diagnostic::warning(
                    "lnk.idlist",
                    format!("failed to read IDList size: {error}"),
                ));
                offset = data.len();
            }
        }
    }

    let mut local_path = None;
    let mut network_path = None;
    if flags & 0x2 != 0 {
        if offset.checked_add(28).is_some_and(|end| end <= data.len()) {
            if let Some(info) = parse_link_info(data, offset) {
                local_path = info.local_path;
                network_path = info.network_path;
                offset = offset.saturating_add(info.size).min(data.len());
            } else {
                warnings.push(Diagnostic::warning(
                    "lnk.linkinfo",
                    "failed to parse LinkInfo",
                ));
            }
        } else {
            warnings.push(Diagnostic::warning(
                "lnk.linkinfo",
                "LinkInfo header extends past stream end",
            ));
        }
    }

    let mut relative_path = None;
    if flags & 0x4 != 0 {
        let (_, next) = read_lnk_string(data, offset, flags);
        offset = next;
    }
    if flags & 0x8 != 0 {
        let (value, next) = read_lnk_string(data, offset, flags);
        relative_path = value;
        offset = next;
    }
    if flags & 0x10 != 0 {
        let (_, next) = read_lnk_string(data, offset, flags);
        offset = next;
    }
    if flags & 0x20 != 0 {
        let (_, next) = read_lnk_string(data, offset, flags);
        offset = next;
    }
    if flags & 0x40 != 0 {
        let (_, _next) = read_lnk_string(data, offset, flags);
    }

    Some(ShellLinkSummary {
        local_path,
        network_path,
        relative_path,
        warnings,
    })
}

fn looks_like_lnk_stream(data: &[u8]) -> bool {
    data.len() >= 0x4c
        && read_u32(data, 0).ok() == Some(0x4c)
        && data.get(4..20)
            == Some(&[
                0x01, 0x14, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x46,
            ])
}

fn parse_link_info(data: &[u8], link_info_start: usize) -> Option<LinkInfoSummary> {
    let link_info_size = read_u32(data, link_info_start).ok()? as usize;
    let link_info_header_size = read_u32(data, link_info_start + 4).ok()? as usize;
    let link_info_end = link_info_start.checked_add(link_info_size)?;
    if link_info_size < 28 || link_info_end > data.len() {
        return None;
    }

    let link_info_flags = read_u32(data, link_info_start + 8).ok()?;
    let local_base_offset = read_u32(data, link_info_start + 16).ok()? as usize;
    let network_offset = read_u32(data, link_info_start + 20).ok()? as usize;
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
    let network = if link_info_flags & 0x2 != 0 {
        parse_common_network_relative_link(data, link_info_start, link_info_size, network_offset)
    } else {
        None
    };
    let network_base = network.and_then(|network| network.net_name);

    let local_path = match (&base, &suffix) {
        (Some(base), Some(suffix)) if !base.is_empty() && !suffix.is_empty() => {
            Some(join_windows_path(base, suffix))
        }
        (Some(base), _) if looks_like_windows_path(base) => Some(base.clone()),
        (_, Some(suffix)) if looks_like_windows_path(suffix) => Some(suffix.clone()),
        _ => None,
    };
    let network_path = match (&network_base, &suffix) {
        (Some(base), Some(suffix)) if !base.is_empty() && !suffix.is_empty() => {
            Some(join_windows_path(base, suffix))
        }
        (Some(base), _) if looks_like_unc_path(base) => Some(base.clone()),
        _ => None,
    };

    Some(LinkInfoSummary {
        size: link_info_size,
        local_path,
        network_path,
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
    let net_name_unicode_offset = if size >= 24 {
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

fn looks_like_unc_path(value: &str) -> bool {
    value.starts_with("\\\\") || value.starts_with("//")
}

fn read_lnk_string(data: &[u8], offset: usize, flags: u32) -> (Option<String>, usize) {
    if offset + 2 > data.len() {
        return (None, data.len());
    }
    let chars =
        u16::from_le_bytes(data[offset..offset + 2].try_into().expect("2-byte chunk")) as usize;
    let string_start = offset + 2;
    if flags & 0x80 != 0 {
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

fn read_i64(data: &[u8], offset: usize) -> Result<i64, String> {
    let bytes = data
        .get(offset..offset + 8)
        .ok_or_else(|| format!("unexpected end of data at offset {offset}"))?;
    Ok(i64::from_le_bytes(bytes.try_into().expect("8-byte chunk")))
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

fn decode_hostname(bytes: &[u8]) -> String {
    let ascii = decode_ascii_nul_padded(bytes);
    if !ascii.is_empty() {
        ascii
    } else {
        decode_utf16_lossy(bytes).trim_end_matches('\0').to_string()
    }
}

fn decode_ascii_nul_padded(bytes: &[u8]) -> String {
    let end = bytes
        .iter()
        .position(|byte| *byte == 0)
        .unwrap_or(bytes.len());
    String::from_utf8_lossy(&bytes[..end]).trim().to_string()
}

fn format_guid_from_le_bytes(bytes: &[u8]) -> String {
    if bytes.len() != 16 {
        return String::new();
    }

    let data1 = u32::from_le_bytes(bytes[0..4].try_into().expect("4-byte chunk"));
    let data2 = u16::from_le_bytes(bytes[4..6].try_into().expect("2-byte chunk"));
    let data3 = u16::from_le_bytes(bytes[6..8].try_into().expect("2-byte chunk"));
    format!(
        "{data1:08x}-{data2:04x}-{data3:04x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}",
        bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14], bytes[15]
    )
}

fn mac_from_droid_bytes(bytes: &[u8]) -> String {
    if bytes.len() != 16 {
        return String::new();
    }

    bytes[10..16]
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<Vec<_>>()
        .join(":")
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::iter;

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
    fn parse_bytes_reads_true_header_and_entry_ids() {
        let parsed = parse_bytes(build_minimal_cfb_with_dest_list("C:\\Test\\file.txt"))
            .expect("minimal CFB should parse");
        let dest = parsed.dest_list();

        assert_eq!(dest.version(), 4);
        assert_eq!(dest.header_counter_raw(), 0x3fc0_0000);
        assert_eq!(dest.header_counter_f32(), 1.5);
        assert_eq!(dest.last_entry_id(), 0x0000_0001_0000_002a);
        assert_eq!(dest.last_entry_number(), 42);
        assert_eq!(dest.add_delete_action_count(), 0x0000_0002_0000_0064);

        let entry = &dest.entries()[0];
        assert_eq!(entry.entry_id(), 0x0000_0063_0000_002a);
        assert_eq!(entry.entry_number(), 42);
        assert_eq!(entry.entry_number_unknown(), 99);
        assert_eq!(entry.stream_name(), "2a");
        assert_eq!(entry.checksum(), -7);
        assert_eq!(entry.hostname(), "HOST");
        assert_eq!(entry.volume_droid(), "33221100-5544-7766-8899-aabbccddeeff");
        assert_eq!(entry.file_droid_mac(), "aa:bb:cc:dd:ee:01");
        assert_eq!(entry.reserved_78(), Some(0x1111_2222));
        assert_eq!(entry.reserved_7c(), Some(0x3333_4444));
    }

    #[test]
    fn visible_entries_v4_frequent_folders_orders_filters_and_dedupes() {
        let entries = vec![
            destlist_entry_for_test("C:\\Pinned2", 2, 2, 5, 3, Some(1)),
            destlist_entry_for_test("C:\\Pinned1", 1, 1, 5, 3, Some(0)),
            destlist_entry_for_test("c:/pinned1/", 7, 7, 0, 4, None),
            destlist_entry_for_test("C:\\FrequentA", 3, 3, 1, 2, None),
            destlist_entry_for_test("C:\\FrequentB", 4, 4, 0, 2, None),
            destlist_entry_for_test("C:\\FrequentA", 5, 5, 2, 4, None),
            destlist_entry_for_test("C:\\LowCount", 6, 6, 0, 1, None),
        ];
        let dest = dest_list_for_test(4, entries);

        let visible: Vec<_> = visible_entries(&dest)
            .into_iter()
            .map(|entry| entry.path().to_string())
            .collect();

        assert_eq!(
            visible,
            vec![
                "C:\\Pinned1",
                "C:\\Pinned2",
                "C:\\FrequentB",
                "C:\\FrequentA"
            ]
        );
    }

    #[test]
    fn visible_entries_v4_recent_files_skips_hidden_and_dedupes() {
        let entries = vec![
            destlist_entry_for_test("C:\\Hidden.txt", 1, 1, 0, 0, None),
            destlist_entry_for_test("C:\\Report.txt", 2, 2, 0, 1, None),
            destlist_entry_for_test("c:\\report.txt", 3, 3, 0, 5, None),
            destlist_entry_for_test("C:/Folder/İtem.txt/", 5, 5, 0, 3, None),
            destlist_entry_for_test("c:\\folder\\i\u{307}tem.txt", 6, 6, 0, 4, None),
            destlist_entry_for_test("C:\\Other.txt", 4, 4, 0, 1, None),
        ];
        let dest = dest_list_for_test(4, entries);

        let visible: Vec<_> = visible_entries(&dest)
            .into_iter()
            .map(|entry| entry.path().to_string())
            .collect();

        assert_eq!(
            visible,
            vec!["C:\\Report.txt", "C:/Folder/İtem.txt/", "C:\\Other.txt"]
        );
    }

    #[test]
    fn visible_entries_v6_uses_pinned_then_reverse_ranked_normal_slots() {
        let entries = vec![
            destlist_entry_for_test("C:\\Normal0", 1, 1, 0, 1, None),
            destlist_entry_for_test("C:\\Normal3", 2, 2, 3, 1, None),
            destlist_entry_for_test("C:\\Normal4", 3, 3, 4, 1, None),
            destlist_entry_for_test("C:\\Pinned", 4, 4, 0, 1, Some(0)),
            destlist_entry_for_test("c:/pinned/", 5, 5, 2, 1, None),
        ];
        let dest = dest_list_for_test(6, entries);

        let visible: Vec<_> = visible_entries(&dest)
            .into_iter()
            .map(|entry| entry.path().to_string())
            .collect();

        assert_eq!(visible, vec!["C:\\Pinned", "C:\\Normal3", "C:\\Normal0"]);
    }

    #[test]
    fn frequent_folder_pin_status_from_entries_detects_pinned_entry() {
        let entries = vec![destlist_entry_for_test("C:\\Folder", 1, 1, 0, 1, Some(0))];

        assert_eq!(
            frequent_folder_pin_status_from_entries("c:/folder", &entries),
            FrequentFolderPinStatus::Pinned
        );
    }

    #[test]
    fn frequent_folder_pin_status_from_entries_detects_unpinned_entry() {
        let entries = vec![destlist_entry_for_test("C:\\Folder", 1, 1, 0, 1, None)];

        assert_eq!(
            frequent_folder_pin_status_from_entries("C:\\Folder", &entries),
            FrequentFolderPinStatus::Unpinned
        );
    }

    #[test]
    fn frequent_folder_pin_status_from_entries_prefers_any_pinned_match() {
        let entries = vec![
            destlist_entry_for_test("C:\\Folder", 1, 1, 0, 1, None),
            destlist_entry_for_test("C:\\Folder", 2, 2, 0, 1, Some(0)),
        ];

        assert_eq!(
            frequent_folder_pin_status_from_entries("C:\\Folder", &entries),
            FrequentFolderPinStatus::Pinned
        );
    }

    #[test]
    fn frequent_folder_pin_status_from_entries_returns_not_found_for_unmatched_path() {
        let entries = vec![destlist_entry_for_test("C:\\Other", 1, 1, 0, 1, Some(0))];

        assert_eq!(
            frequent_folder_pin_status_from_entries("C:\\Folder", &entries),
            FrequentFolderPinStatus::NotFound
        );
    }

    #[test]
    fn frequent_folder_pin_status_from_entries_matches_windows_path_variants() {
        let entries = vec![destlist_entry_for_test("C:\\Folder\\", 1, 1, 0, 1, None)];

        assert_eq!(
            frequent_folder_pin_status_from_entries("c:/folder", &entries),
            FrequentFolderPinStatus::Unpinned
        );
    }

    #[test]
    fn parse_lnk_local_path_reads_unicode_local_base_and_suffix() {
        let lnk = build_lnk_with_link_info(Some(("C:\\Base", "Child.txt")), None, None);

        assert_eq!(
            parse_lnk_local_path(&lnk).as_deref(),
            Some("C:\\Base\\Child.txt")
        );
    }

    #[test]
    fn parse_lnk_local_path_reads_unc_network_path() {
        let lnk =
            build_lnk_with_link_info(None, Some(("\\\\server\\share", "Folder\\File.txt")), None);

        assert_eq!(
            parse_lnk_local_path(&lnk).as_deref(),
            Some("\\\\server\\share\\Folder\\File.txt")
        );
    }

    #[test]
    fn parse_shell_link_summary_reads_relative_path_as_source() {
        let lnk = build_lnk_with_link_info(None, None, Some("..\\Relative.txt"));
        let summary = parse_shell_link_summary(&lnk).expect("lnk should parse");

        assert_eq!(summary.relative_path.as_deref(), Some("..\\Relative.txt"));
    }

    #[test]
    #[ignore = "Integration test; reads system Jump List file - run with: cargo test parse_recent_files_metadata -- --ignored --nocapture"]
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
    #[ignore = "Integration test; reads system Jump List file - run with: cargo test parse_frequent_folders_metadata -- --ignored --nocapture"]
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

    fn dest_list_for_test(version: u32, entries: Vec<DestListEntry>) -> DestList {
        DestList {
            version,
            declared_entry_count: entries.len(),
            pinned_entry_count: entries.iter().filter(|entry| entry.is_pinned()).count() as u32,
            header_counter_raw: 0,
            header_counter_f32: 0.0,
            last_entry_id: entries.last().map(DestListEntry::entry_id).unwrap_or(0),
            last_entry_number: entries.last().map(DestListEntry::entry_number).unwrap_or(0),
            add_delete_action_count: 0,
            entries,
            diagnostics: Vec::new(),
        }
    }

    fn destlist_entry_for_test(
        path: &str,
        entry_number: u32,
        entry_number_unknown: u32,
        recent_rank: i32,
        access_count: u32,
        pin_order: Option<i32>,
    ) -> DestListEntry {
        let entry_id = ((entry_number_unknown as u64) << 32) | entry_number as u64;
        DestListEntry {
            entry_offset: 0,
            entry_len: 0,
            mru_position: 0,
            checksum: 0,
            entry_id,
            entry_number,
            entry_number_unknown,
            hostname: String::new(),
            volume_droid: String::new(),
            file_droid: String::new(),
            volume_birth_droid: String::new(),
            file_birth_droid: String::new(),
            file_droid_mac: String::new(),
            stream_name: format!("{entry_number:x}"),
            raw_path: path.to_string(),
            path: path.to_string(),
            pin_status: pin_order.unwrap_or(-1),
            pin_order,
            rank: recent_rank,
            recent_rank,
            count: access_count,
            access_count,
            score: 0.0,
            last_access_filetime: Some(100 + entry_number as u64),
            last_interaction_filetime: Some(100 + entry_number as u64),
            sps_size: None,
            reserved_78: None,
            reserved_7c: None,
            path_sources: vec![PathSource::new("test", path)],
            warnings: Vec::new(),
        }
    }

    fn build_minimal_cfb_with_dest_list(path: &str) -> Vec<u8> {
        let dest_list = build_dest_list(path);
        let num_mini_sectors = dest_list.len().div_ceil(64);
        let mini_stream_size = num_mini_sectors * 64;
        let mut file = vec![0u8; 512 + 512 * 4];

        file[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        write_u16(&mut file, 0x1a, 3);
        write_u16(&mut file, 0x1c, 0xFFFE);
        write_u16(&mut file, 0x1e, 9);
        write_u16(&mut file, 0x20, 6);
        write_u32(&mut file, 0x30, 0);
        write_u32(&mut file, 0x38, 0x1000);
        write_u32(&mut file, 0x3c, 2); // mini FAT at sector 2
        write_u32(&mut file, 0x40, 1);
        write_u32(&mut file, 0x44, 0xFFFF_FFFF);
        write_u32(&mut file, 0x48, 0);
        for index in 0..109 {
            write_u32(&mut file, 0x4c + index * 4, 0xFFFF_FFFF);
        }
        write_u32(&mut file, 0x4c, 3); // regular FAT at sector 3

        // Directory at sector 0: Root Entry (mini stream container at sector 1), DestList in mini sector 0
        write_directory_entry(&mut file, 512, "Root Entry", 5, 1, mini_stream_size as u64);
        write_directory_entry(
            &mut file,
            512 + 128,
            "DestList",
            2,
            0,
            dest_list.len() as u64,
        );

        // Mini stream container at sector 1
        file[512 + 512..512 + 512 + dest_list.len()].copy_from_slice(&dest_list);

        // Mini FAT at sector 2: chain num_mini_sectors mini sectors then END_OF_CHAIN
        let mini_fat_offset = 512 + 512 * 2;
        for i in 0..128usize {
            let val = if i + 1 < num_mini_sectors {
                i as u32 + 1
            } else if i + 1 == num_mini_sectors {
                0xFFFF_FFFE
            } else {
                0xFFFF_FFFF
            };
            write_u32(&mut file, mini_fat_offset + i * 4, val);
        }

        // Regular FAT at sector 3
        let fat_offset = 512 + 512 * 3;
        for index in 0..128 {
            write_u32(&mut file, fat_offset + index * 4, 0xFFFF_FFFF);
        }
        write_u32(&mut file, fat_offset, 0xFFFF_FFFE); // sector 0: directory
        write_u32(&mut file, fat_offset + 4, 0xFFFF_FFFE); // sector 1: mini stream container
        write_u32(&mut file, fat_offset + 8, 0xFFFF_FFFE); // sector 2: mini FAT
        write_u32(&mut file, fat_offset + 12, 0xFFFF_FFFD); // sector 3: FAT sector

        file
    }

    fn build_dest_list(path: &str) -> Vec<u8> {
        let path_bytes = utf16_bytes(path);
        let entry_len = 0x82 + path_bytes.len() + 4;
        let mut data = vec![0u8; 32 + entry_len];
        write_u32(&mut data, 0, 4);
        write_u32(&mut data, 4, 1);
        write_u32(&mut data, 8, 0);
        write_u32(&mut data, 0x0c, 0x3fc0_0000);
        write_u64(&mut data, 0x10, 0x0000_0001_0000_002a);
        write_u64(&mut data, 0x18, 0x0000_0002_0000_0064);

        let offset = 32;
        write_i64(&mut data, offset, -7);
        data[offset + 0x08..offset + 0x18].copy_from_slice(&[
            0x00, 0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xaa, 0xbb, 0xcc, 0xdd,
            0xee, 0xff,
        ]);
        data[offset + 0x18..offset + 0x28].copy_from_slice(&[
            0x10, 0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xaa, 0xbb, 0xcc, 0xdd,
            0xee, 0x01,
        ]);
        data[offset + 0x28..offset + 0x38].copy_from_slice(&[0x22; 16]);
        data[offset + 0x38..offset + 0x48].copy_from_slice(&[0x33; 16]);
        data[offset + 0x48..offset + 0x4c].copy_from_slice(b"HOST");
        write_u64(&mut data, offset + 0x58, 0x0000_0063_0000_002a);
        write_u32(&mut data, offset + 0x60, 1.5f32.to_bits());
        write_u64(&mut data, offset + 0x64, 132_537_600_000_000_000);
        write_i32(&mut data, offset + 0x6c, -1);
        write_i32(&mut data, offset + 0x70, 7);
        write_u32(&mut data, offset + 0x74, 3);
        write_u32(&mut data, offset + 0x78, 0x1111_2222);
        write_u32(&mut data, offset + 0x7c, 0x3333_4444);
        write_u16(&mut data, offset + 0x80, path.encode_utf16().count() as u16);
        data[offset + 0x82..offset + 0x82 + path_bytes.len()].copy_from_slice(&path_bytes);
        write_u32(&mut data, offset + 0x82 + path_bytes.len(), 0);
        data
    }

    fn build_lnk_with_link_info(
        local: Option<(&str, &str)>,
        network: Option<(&str, &str)>,
        relative_path: Option<&str>,
    ) -> Vec<u8> {
        let mut link_info = vec![0u8; 0x24];
        let mut flags = 0u32;
        let local_base_offset = 0u32;
        let mut network_offset = 0u32;
        let suffix_offset = 0u32;
        let mut local_base_unicode_offset = 0u32;
        let mut suffix_unicode_offset = 0u32;

        if let Some((base, suffix)) = local {
            flags |= 0x1;
            local_base_unicode_offset = link_info.len() as u32;
            link_info.extend(utf16_z_bytes(base));
            suffix_unicode_offset = link_info.len() as u32;
            link_info.extend(utf16_z_bytes(suffix));
        }

        if let Some((net_name, suffix)) = network {
            flags |= 0x2;
            network_offset = link_info.len() as u32;
            let mut network_block = vec![0u8; 0x18];
            write_u32(&mut network_block, 4, 0x2);
            write_u32(&mut network_block, 8, 0);
            write_u32(&mut network_block, 12, 0);
            write_u32(&mut network_block, 16, 0);
            write_u32(&mut network_block, 20, 0x18);
            network_block.extend(utf16_z_bytes(net_name));
            let network_size = network_block.len() as u32;
            write_u32(&mut network_block, 0, network_size);
            link_info.extend(network_block);

            suffix_unicode_offset = link_info.len() as u32;
            link_info.extend(utf16_z_bytes(suffix));
        }

        let link_info_size = link_info.len() as u32;
        write_u32(&mut link_info, 0, link_info_size);
        write_u32(&mut link_info, 4, 0x24);
        write_u32(&mut link_info, 8, flags);
        write_u32(&mut link_info, 16, local_base_offset);
        write_u32(&mut link_info, 20, network_offset);
        write_u32(&mut link_info, 24, suffix_offset);
        write_u32(&mut link_info, 28, local_base_unicode_offset);
        write_u32(&mut link_info, 32, suffix_unicode_offset);

        let mut lnk = vec![0u8; 0x4c];
        write_u32(&mut lnk, 0, 0x4c);
        lnk[4..20].copy_from_slice(&[
            0x01, 0x14, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x46,
        ]);
        let mut shell_flags = 0x2u32 | 0x80;
        if relative_path.is_some() {
            shell_flags |= 0x8;
        }
        write_u32(&mut lnk, 0x14, shell_flags);
        lnk.extend(link_info);
        if let Some(relative_path) = relative_path {
            append_lnk_string(&mut lnk, relative_path);
        }
        lnk
    }

    fn write_directory_entry(
        data: &mut [u8],
        offset: usize,
        name: &str,
        object_type: u8,
        start_sector: u32,
        stream_size: u64,
    ) {
        let name_bytes = utf16_bytes(&format!("{name}\0"));
        data[offset..offset + name_bytes.len()].copy_from_slice(&name_bytes);
        write_u16(data, offset + 64, name_bytes.len() as u16);
        data[offset + 66] = object_type;
        write_u32(data, offset + 116, start_sector);
        write_u64(data, offset + 120, stream_size);
    }

    fn append_lnk_string(data: &mut Vec<u8>, value: &str) {
        write_u16_to_vec(data, value.encode_utf16().count() as u16);
        data.extend(utf16_bytes(value));
    }

    fn utf16_bytes(value: &str) -> Vec<u8> {
        value.encode_utf16().flat_map(u16::to_le_bytes).collect()
    }

    fn utf16_z_bytes(value: &str) -> Vec<u8> {
        value
            .encode_utf16()
            .chain(iter::once(0))
            .flat_map(u16::to_le_bytes)
            .collect()
    }

    fn write_u16(data: &mut [u8], offset: usize, value: u16) {
        data[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
    }

    fn write_u32(data: &mut [u8], offset: usize, value: u32) {
        data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    }

    fn write_u64(data: &mut [u8], offset: usize, value: u64) {
        data[offset..offset + 8].copy_from_slice(&value.to_le_bytes());
    }

    fn write_i32(data: &mut [u8], offset: usize, value: i32) {
        data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    }

    fn write_i64(data: &mut [u8], offset: usize, value: i64) {
        data[offset..offset + 8].copy_from_slice(&value.to_le_bytes());
    }

    fn write_u16_to_vec(data: &mut Vec<u8>, value: u16) {
        data.extend(value.to_le_bytes());
    }
}
