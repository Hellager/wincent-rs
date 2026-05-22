//! Direct CFB parser for `.automaticDestinations-ms` Jump List backing files.
//!
//! This module is gated behind the `destlist` Cargo feature and provides rich
//! per-entry metadata (access count, pin status, rank, score, FILETIME) that is
//! not available through the COM/Shell API used by the rest of this crate.
//!
//! # Quick Start
//!
//! ```rust,no_run
//! # #[cfg(feature = "destlist")]
//! # {
//! use wincent::destlist::{parse_file, entries, recent_files_dest_path};
//!
//! let path = recent_files_dest_path().unwrap();
//! let parsed = parse_file(&path).unwrap();
//! for entry in entries(&parsed.dest_list) {
//!     println!("{} (count={})", entry.path, entry.count);
//! }
//! # }
//! ```
//!
//! # Known Limitations
//!
//! **Only DestList versions 4 and 6 are supported.** Other versions return
//! [`WincentError::DestListUnsupportedVersion`].

pub(super) mod cfb;
pub mod parser;
pub mod time;

pub use parser::{
    entries, frequent_folders_dest_path, parse_bytes, parse_file, recent_files_dest_path,
    AutomaticDestinations, CfbDirectoryEntry, CfbInfo, DestList, DestListEntry,
};
pub use time::filetime_to_system_time;
