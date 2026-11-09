//! # wincent
//!
//! `wincent` is a Rust library for managing Windows Quick Access items, providing a safe and
//! efficient interface to interact with Windows Quick Access functionality.
//!
//! ## Key Features
//!
//! ### Core Operations
//! - Query Quick Access items (recent files and frequent folders)
//! - Add/Remove items to/from Quick Access
//! - Clear Quick Access categories
//! - Check item existence
//!
//! ### Advanced Management
//! - On-disk DestList queries using rules recovered from Shell behavior
//! - Timeout protection for shell operations
//! - Force refresh support
//!
//! ### System Integration
//! - Windows API integration for reliable operations
//! - Native Shell and COM mutation paths
//! - Cross-version Windows support
//!
//! ## Quick Start
//!
//! ```rust,no_run
//! use wincent::prelude::*;
//!
//! fn main() -> WincentResult<()> {
//!     // Create manager instance
//!     let manager = QuickAccessManager::new();
//!     
//!     // Add a file to Recent Files
//!     manager.add_item("C:\\path\\to\\file.txt")?;
//!     
//!     // Query all Quick Access items as PathBuf values
//!     let items = manager.get_item_paths(QuickAccess::All)?;
//!     println!("Quick Access items: {:?}", items);
//!     
//!     Ok(())
//! }
//! ```
//!
//! ## Implementation Details
//!
//! - Implements timeout mechanism to prevent deadlocks
//! - Reads Quick Access directly from DestList data without live Shell queries
//! - Applies Quick Access UI filtering recovered from observed Explorer behavior
//! - Handles system-specific edge cases
//!
//! ## Safety and Reliability
//!
//! - Validates all paths before operations
//! - Provides comprehensive error handling
//! - Separates Recent Files backing-data rebuilds from Explorer window refreshes
//! - Manages system resources properly
//!
//! ## Best Practices
//!
//! - Treat successful add/remove calls as successful Shell mutations, not as a
//!   guarantee that persisted queries and open Explorer views are synchronized
//! - Treat mutation timeouts as indeterminate outcomes; inspect state before
//!   issuing another mutation
//! - Call [`manager::QuickAccessManager::refresh`] explicitly when an open
//!   Explorer Quick Access or Home view should be notified and refreshed
//!

mod backend;
mod com;
mod com_thread;
mod empty;
pub mod error;
mod explorer_window;
mod handle;
mod lock;
pub mod manager;
mod query;
mod recent_links;
pub mod standalone;
#[cfg(test)]
mod test_utils;
pub(crate) mod utils;

pub mod visible;

pub mod destlist;

pub mod status;

#[allow(unused)]
/// Convenient re-exports for common Quick Access operations.
pub mod prelude {
    pub use crate::error::WincentError;
    pub use crate::lock::QuickAccessLockTarget;
    pub use crate::manager::{QuickAccessManager, QuickAccessManagerBuilder};
    pub use crate::standalone::{
        add_to_recent_files, empty_frequent_folders, empty_recent_files, get_frequent_folders,
        get_recent_files, is_folder_pinned, lock_frequent_folders, lock_recent_files,
        pin_to_frequent_folders, refresh_quick_access, remove_from_frequent_folders,
        remove_from_recent_files, unlock_frequent_folders, unlock_recent_files,
        FrequentFoldersLock, RecentFilesLock,
    };
    pub use crate::status::{FrequentFolderSlot, QuickAccessStatus};

    pub use crate::visible::{is_visible, set_visible, VisibilityTarget};

    pub use crate::utils::{get_windows_version, WindowsVersion};

    pub use crate::{QuickAccess, WincentResult};

    pub use crate::destlist::{
        entries, filetime_to_system_time, frequent_folders_dest_path,
        parse_bytes as parse_dest_bytes, parse_bytes_with_kind as parse_dest_bytes_with_kind,
        parse_file as parse_dest_file, parse_file_with_kind as parse_dest_file_with_kind,
        quick_access_entries, quick_access_entries_for_kind, recent_files_dest_path,
        visible_entries, visible_entries_for_kind, AutomaticDestinations, CfbInfo, CfbStream,
        DestList, DestListEntry, DestListKind, Diagnostic, DiagnosticSeverity, PathSource,
        DEFAULT_FREQUENT_FOLDERS_NORMAL_SLOTS, DEFAULT_RECENT_FILES_RESULT_LIMIT,
        DEST_LIST_PROPERTY_STORE_STREAM, FREQUENT_FOLDERS_MIN_ACCESS_COUNT,
    };
}

use crate::error::WincentError;

pub use crate::lock::QuickAccessLockTarget;
pub use crate::standalone::{
    add_to_recent_files, empty_frequent_folders, empty_recent_files, get_frequent_folders,
    get_recent_files, is_folder_pinned, lock_frequent_folders, lock_recent_files,
    pin_to_frequent_folders, refresh_quick_access, remove_from_frequent_folders,
    remove_from_recent_files, unlock_frequent_folders, unlock_recent_files, FrequentFoldersLock,
    RecentFilesLock,
};
pub use crate::visible::VisibilityTarget;

/// Quick Access categories supported by this crate.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum QuickAccess {
    /// Frequently used or pinned folders shown by Quick Access.
    FrequentFolders,
    /// Recently used files shown by Quick Access.
    RecentFiles,
    /// Both Recent Files and Frequent Folders.
    All,
}

/// Result type used by wincent APIs.
pub type WincentResult<T> = Result<T, WincentError>;
