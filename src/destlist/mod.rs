//! Direct CFB parser for `.automaticDestinations-ms` Jump List backing files.
//!
//! This module is gated behind the `destlist` Cargo feature and provides rich
//! per-entry metadata (access count, pin order, recent rank, FILETIME) that is
//! not available through the COM/Shell API used by the rest of this crate.
//!
//! # Quick Start
//!
//! ```rust,no_run
//! # #[cfg(feature = "destlist")]
//! # {
//! use wincent::destlist::{parse_file, visible_entries, recent_files_dest_path};
//!
//! let path = recent_files_dest_path();
//! let parsed = parse_file(&path).unwrap();
//! let visible = visible_entries(&parsed.dest_list, 10);
//! for entry in &visible {
//!     println!("{} (access_count={})", entry.path, entry.access_count);
//! }
//! # }
//! ```
//!
//! # Known Limitations
//!
//! 1. **Only DestList versions 4 and 6 are supported.** Other versions return
//!    [`WincentError::DestListUnsupportedVersion`].

pub(super) mod cfb;
pub mod filter;
pub mod parser;
pub mod time;

pub use filter::visible_entries;
pub use parser::{
    frequent_folders_dest_path, parse_bytes, parse_file, recent_files_dest_path,
    AutomaticDestinations, CfbDirectoryEntry, CfbInfo, DestList, DestListEntry,
};
pub use time::filetime_to_system_time;
