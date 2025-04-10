//! # wincent
//!
//! `wincent` is a Rust library for managing Windows Quick Access items, including recent files
//! and frequent folders. It provides a comprehensive API to interact with Windows Quick Access functionality.
//!
//! ## Main Features
//!
//! - Feasibility Management
//!   - Check and fix PowerShell script execution
//!   - Verify Quick Access operations support
//!
//! - Quick Access Operations
//!   - Query recent files and frequent folders
//!   - Add/Remove items from Quick Access
//!   - Check item existence
//!
//! - Visibility Control
//!   - Show/Hide recent files
//!   - Show/Hide frequent folders
//!
//! ## Basic Example
//!
//! ```no_run
//! use wincent::{
//!     feasible::{check_feasible, fix_feasible},
//!     query::get_quick_access_items,
//!     error::WincentError
//! };
//!
//! fn main() -> Result<(), WincentError> {
//!     // Ensure operations are feasible
//!     if !check_feasible()? {
//!         fix_feasible()?;
//!     }
//!
//!     // Get all Quick Access items
//!     let items = get_quick_access_items()?;
//!     println!("Found {} Quick Access items", items.len());
//!
//!     Ok(())
//! }
//! ```
//!
//! ## Advanced Example
//!
//! ```no_run
//! use wincent::{
//!     handle::{add_to_frequent_folders, remove_from_recent_files},
//!     query::{get_recent_files, is_in_quick_access},
//!     error::WincentError
//! };
//!
//! fn main() -> Result<(), WincentError> {
//!     // Pin an important project folder
//!     let project_path = "C:\\Projects\\important-project";
//!     if !is_in_quick_access(project_path)? {
//!         add_to_frequent_folders(project_path)?;
//!     }
//!
//!     // Remove sensitive files from recent items
//!     let sensitive_files = get_recent_files()?
//!         .into_iter()
//!         .filter(|path| path.contains("password") || path.contains("secret"));
//!
//!     for file in sensitive_files {
//!         remove_from_recent_files(&file)?;
//!     }
//!
//!     Ok(())
//! }
//! ```
//!
//! ## Features
//!
//! - Comprehensive error handling
//! - PowerShell script integration
//! - Registry management
//! - Windows API integration
//! - Cross-version Windows support
//!

pub mod empty;
pub mod error;
pub mod feasible;
pub mod handle;
pub mod query;
pub mod manager;
mod script_executor;
mod script_storage;
mod script_strategy;
mod test_utils;
mod utils;

#[allow(unused)]
pub mod predule {
    pub use crate::error::WincentError;
    pub use crate::manager::QuickAccessManager;
    pub use crate::{QuickAccess, WincentResult};
}

use crate::error::WincentError;

#[derive(Debug, PartialEq, Clone)]
pub enum QuickAccess {
    FrequentFolders,
    RecentFiles,
    All,
}

pub type WincentResult<T> = Result<T, WincentError>;
