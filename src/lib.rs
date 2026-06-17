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
//! - Native Windows API fast paths with PowerShell fallbacks
//! - Timeout protection for shell operations
//! - Force refresh support
//!
//! ### System Integration
//! - Windows API integration for reliable operations
//! - PowerShell script execution for complex tasks
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
//!     manager.add_item(
//!         "C:\\path\\to\\file.txt",
//!         QuickAccess::RecentFiles,
//!         AddOptions::new()
//!             .force_recent_files_rebuild()
//!             .refresh_explorer(),
//!     )?;
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
//! - Supports both Windows API and PowerShell operations
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
//! - Use [`AddOptions::force_recent_files_rebuild`] when adding recent files for immediate visibility
//! - Use [`EmptyOptions::remove_pinned_folders`] only when you intend to remove user-pinned folders
//!

mod backend;
mod batch;
mod com;
mod com_thread;
mod empty;
pub mod error;
mod explorer_window;
mod handle;
pub mod manager;
mod query;
mod quick_access_lock;
mod recent_links;
mod restore;
mod retry;
mod script_executor;
mod script_storage;
mod script_strategy;
mod test_utils;
mod utils;

pub mod visible;

pub mod destlist;

#[allow(unused)]
/// Convenient re-exports for common Quick Access operations.
pub mod prelude {
    pub use crate::error::{
        PowerShellError, PowerShellErrorKind, PowerShellOperation, QuickAccessPostMutationStep,
        WincentError,
    };
    pub use crate::manager::{
        AddOptions, QuickAccessItem, QuickAccessManager, QuickAccessManagerBuilder, RemoveOptions,
    };
    pub use crate::quick_access_lock::{
        QuickAccessLock, QuickAccessLockTarget, QuickAccessUnlockFailure, QuickAccessUnlockOptions,
        QuickAccessUnlockReport,
    };
    pub use crate::restore::{
        FrequentRawPathRemoveReport, FrequentRestoreReport, RecentRestoreReport,
        RestoreDefaultsOptions, RestoreDefaultsReport,
    };

    pub use crate::visible::{
        is_frequent_folders_visible, is_recent_files_visible, is_visible,
        set_frequent_folders_visible, set_frequent_folders_visible_with_options,
        set_recent_files_visible, set_recent_files_visible_with_options, set_visible,
        set_visible_with_options, VisibilityOptions,
    };

    pub use crate::{
        BatchOptions, BatchResult, EmptyOptions, QuickAccess, RetryPolicy, WincentResult,
    };

    pub use crate::destlist::{
        entries, filetime_to_system_time, frequent_folders_dest_path,
        parse_bytes as parse_dest_bytes, parse_file as parse_dest_file, quick_access_entries,
        recent_files_dest_path, visible_entries, AutomaticDestinations, CfbInfo, DestList,
        DestListEntry, Diagnostic, DiagnosticSeverity, PathSource,
    };
}

use crate::error::WincentError;

pub use crate::batch::{BatchOptions, BatchResult};
pub use crate::empty::EmptyOptions;
pub use crate::error::QuickAccessPostMutationStep;
pub use crate::manager::{AddOptions, QuickAccessItem, RemoveOptions};
pub use crate::quick_access_lock::{
    QuickAccessLock, QuickAccessLockTarget, QuickAccessUnlockFailure, QuickAccessUnlockOptions,
    QuickAccessUnlockReport,
};
pub use crate::restore::{
    FrequentRawPathRemoveReport, FrequentRestoreReport, RecentRestoreReport,
    RestoreDefaultsOptions, RestoreDefaultsReport,
};
pub use crate::retry::RetryPolicy;

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
