//! Direct CFB parser for `.automaticDestinations-ms` Jump List backing files.
//!
//! This module provides rich per-entry metadata (access count, pin status, rank,
//! score, FILETIME) that is not available through the COM/Shell API used by the
//! rest of this crate.
//!
//! Offsets exposed by [`DestListEntry`](parser::DestListEntry) are logical
//! DestList stream offsets, not physical file offsets. CFB stream allocation
//! may be non-contiguous, and an embedded Shell Link payload is not guaranteed
//! to be byte-identical to the standalone `.lnk` in the Recent directory.
//! One Shell operation may update both Explorer backing files: adding a recent
//! file can also record access to its parent in Frequent Folders.
//!
//! # Quick Start
//!
//! ```rust,no_run
//! # {
//! use wincent::destlist::{parse_file, entries, recent_files_dest_path};
//!
//! let path = recent_files_dest_path().unwrap();
//! let parsed = parse_file(&path).unwrap();
//! for entry in entries(parsed.dest_list()) {
//!     println!("{} (count={})", entry.path(), entry.count());
//! }
//! # }
//! ```
//!
//! # Known Limitations
//!
//! **DestList versions 4 and 6 are supported.** Other persisted versions return
//! [`crate::error::WincentError::DestListUnsupportedVersion`].
//! Explorer can also keep a zero-length `DestList` stream before the first
//! item is recorded. The parser represents that uninitialized state as version
//! `0`; it is not a persisted DestList format version.
//!
//! A parsed entry is not necessarily visible in Explorer. Explorer also
//! applies version- and list-specific metadata filters and loads the expected
//! Shell Link stream for visible candidates. In observed Windows 10 v4 Recent
//! Files, a missing stream caused only that candidate to be skipped; Explorer
//! did not repair the stream, remove the dangling entry, or rebuild the file.
//! [`visible_entries`] returns metadata-level candidates and does not validate
//! stream availability or apply a caller-specific result limit. Parsed files
//! expose [`AutomaticDestinations::visible_entries`] and
//! [`AutomaticDestinations::shell_entries`] for stream-aware results.

pub(super) mod cfb;
/// Internal destructive tests for removing entries by rebuilding Explorer backing files.
///
/// These helpers are intentionally not part of the public API.
#[cfg(test)]
pub(crate) mod experimental_remove;
/// Parser for Explorer `.automaticDestinations-ms` Jump List files.
pub mod parser;
/// FILETIME conversion helpers for DestList timestamps.
pub mod time;

pub(crate) use parser::frequent_folder_pin_status;
pub use parser::{
    entries, frequent_folders_dest_path, parse_bytes, parse_bytes_with_kind, parse_file,
    parse_file_with_kind, quick_access_entries, quick_access_entries_for_kind,
    recent_files_dest_path, visible_entries, visible_entries_for_kind, AutomaticDestinations,
    CfbDirectoryEntry, CfbInfo, DestList, DestListEntry, DestListKind, Diagnostic,
    DiagnosticSeverity, FrequentFolderPinStatus, PathSource, DEFAULT_FREQUENT_FOLDERS_NORMAL_SLOTS,
    DEFAULT_RECENT_FILES_RESULT_LIMIT, FREQUENT_FOLDERS_MIN_ACCESS_COUNT,
};
pub use time::filetime_to_system_time;
