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
