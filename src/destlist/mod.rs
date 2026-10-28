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
//! expose [`AutomaticDestinations::rule_entries`] for the complete custom
//! CAutoList rule result driven by the built-in [`DestListRuleSet`] defaults,
//! [`AutomaticDestinations::shell_entries`] for a bounded, Shell-query-shaped
//! result, and [`AutomaticDestinations::cfb_entries`] for the raw parsed
//! DestList entries.

pub(crate) mod add;
pub(super) mod cfb;
/// Internal destructive tests for removing entries by rebuilding Explorer backing files.
///
/// These helpers are intentionally not part of the public API.
#[cfg(test)]
pub(crate) mod experimental_remove;
/// Parser for Explorer `.automaticDestinations-ms` Jump List files.
pub mod parser;
pub(crate) mod remove;
#[cfg(test)]
pub(crate) mod test_support;
/// FILETIME conversion helpers for DestList timestamps.
pub mod time;
mod transaction;
pub(crate) mod update;
mod validation;

pub(crate) use parser::frequent_folder_pin_status;
pub use parser::{
    entries, frequent_folders_dest_path, parse_bytes, parse_bytes_with_kind, parse_file,
    parse_file_with_kind, quick_access_entries, quick_access_entries_for_kind,
    recent_files_dest_path, visible_entries, visible_entries_for_kind, AutomaticDestinations,
    CfbDirectoryEntry, CfbInfo, CfbStream, DestList, DestListEntry, DestListKind, Diagnostic,
    DiagnosticSeverity, FrequentFolderPinStatus, PathSource, DEFAULT_FREQUENT_FOLDERS_NORMAL_SLOTS,
    DEFAULT_RECENT_FILES_RESULT_LIMIT, FREQUENT_FOLDERS_MIN_ACCESS_COUNT,
};
pub use time::filetime_to_system_time;

/// Zero-based destination for a direct DestList item move.
///
/// Recent Files positions address the raw DestList entry sequence. Frequent
/// Folders positions address the item's current pin group: pinned entries use
/// pin order, while unpinned entries use rank order.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum DestListPosition {
    /// Move the item to the first position in its list or pin group.
    First,
    /// Move the item to the last position in its list or pin group.
    Last,
    /// Move the item to a zero-based position in its list or pin group.
    Index(usize),
}

/// Requested changes for one DestList v4 entry.
///
/// Score and last-interaction time are intentionally not writable through this
/// API. Pin changes are accepted only for Frequent Folders.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub struct DestListItemUpdate {
    visible: Option<bool>,
    pinned: Option<bool>,
    position: Option<DestListPosition>,
}

impl DestListItemUpdate {
    /// Creates an empty update.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            visible: None,
            pinned: None,
            position: None,
        }
    }

    /// Returns the requested visibility change, if any.
    #[must_use]
    pub const fn visible_value(&self) -> Option<bool> {
        self.visible
    }

    /// Sets whether the entry is eligible for display.
    #[must_use]
    pub const fn with_visible(mut self, visible: bool) -> Self {
        self.visible = Some(visible);
        self
    }

    /// Marks the entry for display.
    #[must_use]
    pub const fn show(self) -> Self {
        self.with_visible(true)
    }

    /// Hides the entry without deleting its DestList entry or numeric stream.
    ///
    /// A pinned Frequent Folder cannot be hidden. Combine this with
    /// [`Self::unpin`] to hide it while retaining it as an unpinned entry.
    #[must_use]
    pub const fn hide(self) -> Self {
        self.with_visible(false)
    }

    /// Returns the requested Frequent Folders pin change, if any.
    #[must_use]
    pub const fn pinned_value(&self) -> Option<bool> {
        self.pinned
    }

    /// Sets the Frequent Folders pinned state.
    #[must_use]
    pub const fn with_pinned(mut self, pinned: bool) -> Self {
        self.pinned = Some(pinned);
        self
    }

    /// Pins a Frequent Folders entry.
    #[must_use]
    pub const fn pin(self) -> Self {
        self.with_pinned(true)
    }

    /// Converts a pinned Frequent Folders entry to an unpinned entry.
    #[must_use]
    pub const fn unpin(self) -> Self {
        self.with_pinned(false)
    }

    /// Returns the requested position change, if any.
    #[must_use]
    pub const fn position_value(&self) -> Option<DestListPosition> {
        self.position
    }

    /// Sets the destination position.
    #[must_use]
    pub const fn with_position(mut self, position: DestListPosition) -> Self {
        self.position = Some(position);
        self
    }

    /// Returns whether no changes were requested.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.visible.is_none() && self.pinned.is_none() && self.position.is_none()
    }
}
