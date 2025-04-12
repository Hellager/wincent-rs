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
//! - Cached script execution
//! - Timeout protection
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
//! use wincent::predule::*;
//!
//! #[tokio::main]
//! async fn main() -> WincentResult<()> {
//!     // Create manager instance
//!     let manager = QuickAccessManager::new().await?;
//!     
//!     // Add a file to Recent Files
//!     manager.add_item(
//!         "C:\\path\\to\\file.txt",
//!         QuickAccess::RecentFiles,
//!         true // force update Explorer
//!     ).await?;
//!     
//!     // Query all Quick Access items
//!     let items = manager.get_items(QuickAccess::All).await?;
//!     println!("Quick Access items: {:?}", items);
//!     
//!     Ok(())
//! }
//! ```
//!
//! ## Implementation Details
//!
//! - Uses `tokio::sync::OnceCell` for lazy initialization
//! - Implements timeout mechanism to prevent deadlocks
//! - Provides caching for performance optimization
//! - Supports both Windows API and PowerShell operations
//! - Handles system-specific edge cases
//!
//! ## Safety and Reliability
//!
//! - Validates all paths before operations
//! - Checks operation feasibility automatically
//! - Provides comprehensive error handling
//! - Supports force refresh for consistency
//! - Manages system resources properly
//!
//! ## Best Practices
//!
//! - Use `force_update` when adding recent files for immediate visibility
//! - Check operation feasibility only when necessary
//! - Clear cache after bulk operations
//! - Consider using `also_system_default` carefully when clearing items
//!

pub mod empty;
pub mod error;
pub mod feasible;
pub mod handle;
pub mod manager;
pub mod query;
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
