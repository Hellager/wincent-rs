//! Direct CFB parser for `.automaticDestinations-ms` Jump List backing files.
//!
//! This module provides rich per-entry metadata (access count, pin status, rank,
//! score, FILETIME) that is not available through the COM/Shell API used by the
//! rest of this crate.
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
//! **DestList versions 1, 3, 4 and 6 are supported.** Other versions return
//! [`crate::error::WincentError::DestListUnsupportedVersion`].

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
    entries, frequent_folders_dest_path, parse_bytes, parse_file, quick_access_entries,
    recent_files_dest_path, visible_entries, AutomaticDestinations, CfbDirectoryEntry, CfbInfo,
    DestList, DestListEntry, Diagnostic, DiagnosticSeverity, FrequentFolderPinStatus, PathSource,
};
pub use time::filetime_to_system_time;
