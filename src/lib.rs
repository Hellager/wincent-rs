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
//!         true // force update Explorer
//!     )?;
//!     
//!     // Query all Quick Access items
//!     let items = manager.get_items(QuickAccess::All)?;
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
//! - Supports force refresh for consistency
//! - Manages system resources properly
//!
//! ## Best Practices
//!
//! - Use `force_update` when adding recent files for immediate visibility
//! - Consider `also_pinned_folders` carefully when clearing frequent folders
//!

pub mod batch;
mod com;
mod com_thread;
pub mod empty;
pub mod error;
pub mod handle;
pub mod manager;
pub mod query;
pub mod retry;
pub mod script_executor;
mod script_storage;
pub mod script_strategy;
mod test_utils;
mod utils;

#[cfg(feature = "destlist")]
pub mod destlist;

#[allow(unused)]
pub mod prelude {
    pub use crate::batch::{BatchOptions, BatchResult};
    pub use crate::empty::EmptyOptions;
    pub use crate::error::{PowerShellError, PowerShellErrorKind, WincentError};
    pub use crate::handle::AddRecentFileOptions;
    pub use crate::manager::{QuickAccessManager, QuickAccessManagerBuilder};
    pub use crate::retry::RetryPolicy;
    pub use crate::script_strategy::PSScript;
    pub use crate::{QuickAccess, WincentResult};

    #[cfg(feature = "destlist")]
    pub use crate::destlist::{
        filetime_to_system_time, frequent_folders_dest_path,
        parse_bytes as parse_dest_bytes, parse_file as parse_dest_file, recent_files_dest_path,
        entries, AutomaticDestinations, CfbInfo, DestList, DestListEntry,
    };

    // Commonly used query functions
    pub use crate::query::{
        get_frequent_folders, get_quick_access_items, get_recent_files, is_frequent_folder_exact,
        is_in_frequent_folders, is_in_quick_access, is_in_quick_access_exact, is_in_recent_files,
        is_recent_file_exact,
    };
}

use crate::error::WincentError;

#[derive(Debug, PartialEq, Clone)]
pub enum QuickAccess {
    FrequentFolders,
    RecentFiles,
    All,
}

pub type WincentResult<T> = Result<T, WincentError>;
