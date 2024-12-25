//! # wincent
//! 
//! `wincent` is a Rust library for managing Windows Quick Access items, including recent files 
//! and frequent folders. It provides a simple API to interact with Windows Quick Access functionality.
//!
//! ## Main Features
//!
//! - Query recent files and frequent folders
//! - Add/Remove items from Quick Access
//! - Check item existence in Quick Access
//! - Control Quick Access visibility
//! - Registry-based feasibility checks
//!
//! ## Examples
//!
//! ```rust
//! use wincent::{
//!     check_feasible, fix_feasible, get_frequent_folders, get_quick_access_items, get_recent_files, WincentError
//! };
//! use std::io::{Error, ErrorKind};
//! 
//! #[tokio::main]
//! async fn main() -> Result<(), WincentError> {
//!     if !check_feasible()?{
//!         fix_feasible()?;
//!
//!         if !check_feasible()? {
//!             return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
//!         }
//!     }
//!
//!     let recent_files: Vec<String> = get_recent_files().await?;
//!     let danger_content = "password";
//! 
//!     for item in recent_files {
//!         if item.contains(danger_content) {
//!             remove_from_recent_files(item).await?;
//!         }
//!     }
//!
//!     Ok(())
//! }
//! ```
//!
//! ## Features
//!
//! - Async support
//! - Error handling
//! - PowerShell script integration
//! - Registry management
//! - Cross-version Windows support
//!

use visible::{is_visialbe_with_registry, set_visiable_with_registry};
use std::io::{Error, ErrorKind};

mod utils;
mod feasible;
mod query;
mod visible;
mod handle;
mod empty;

pub(crate) const SCRIPT_TIMEOUT: u64 = 5;

#[derive(Debug)]
pub enum WincentError {
    ScriptError(powershell_script::PsError),
    IoError(std::io::Error),
    ConvertError(std::array::TryFromSliceError),
    ExecuteError(tokio::task::JoinError),
    TimeoutError(tokio::time::error::Elapsed)
}

pub(crate) enum QuickAccess {
    FrequentFolders,
    RecentFiles,
    All
}

/****************************************************** Feature Feasible ******************************************************/

/// Checks whether the current script is feasible based on the registry.
///
/// # Parameters
/// None
///
/// # Example
/// ```
/// match check_feasible() {
///     Ok(true) => println!("Script is feasible"),
///     Ok(false) => println!("Script is not feasible"),
///     Err(e) => println!("Error checking feasibility: {}", e),
/// }
/// ```
pub fn check_feasible() -> Result<bool, WincentError> {
    feasible::check_script_feasible_with_registry()
}

/// Fixes the feasibility of the current script based on the registry.
///
/// # Parameters
/// None
///
/// # Example
/// ```
/// match fix_feasible() {
///     Ok(()) => println!("Feasibility fixed successfully"),
///     Err(e) => println!("Error fixing feasibility: {}", e),
/// }
/// ```
pub fn fix_feasible() -> Result<(), WincentError> {
    feasible::fix_script_feasible_with_registry()
}

/****************************************************** Query Quick Access ******************************************************/

/// Retrieves a list of recently accessed files.
///
/// # Parameters
/// None
///
/// # Example
/// ```
/// match get_recent_files().await {
///     Ok(files) => {
///         for file in files {
///             println!("Recent file: {}", file);
///         }
///     },
///     Err(e) => {
///         println!("Error getting recent files: {}", e);
///     }
/// }
/// ```
pub async fn get_recent_files() -> Result<Vec<String>, WincentError> {
    match check_feasible() {
        Ok(is_feasible) => {
            if is_feasible {
                let res = query::query_recent_with_ps_script(QuickAccess::RecentFiles).await?;
                return Ok(res);
            } else {
                return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
            }
        },
        Err(e) => return Err(e),
    }
}

/// Retrieves a list of frequently accessed folders.
///
/// # Parameters
/// None
///
/// # Example
/// ```
/// match get_frequent_folders().await {
///     Ok(folders) => {
///         for folder in folders {
///             println!("Frequent folder: {}", folder);
///         }
///     },
///     Err(e) => {
///         println!("Error getting frequent folders: {}", e);
///     }
/// }
/// ```
pub async fn get_frequent_folders() -> Result<Vec<String>, WincentError> {
    match check_feasible() {
        Ok(is_feasible) => {
            if is_feasible {
                let res = query::query_recent_with_ps_script(QuickAccess::FrequentFolders).await?;
                return Ok(res);
            } else {
                return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
            }
        },
        Err(e) => return Err(e),
    }
}

/// Retrieves a list of all quick access items, including recent files and frequent folders.
///
/// # Parameters
/// None
///
/// # Example
/// ```
/// match get_quick_access_items().await {
///     Ok(items) => {
///         for item in items {
///             println!("Quick access item: {}", item);
///         }
///     },
///     Err(e) => {
///         println!("Error getting quick access items: {}", e);
///     }
/// }
/// ```
pub async fn get_quick_access_items() -> Result<Vec<String>, WincentError> {
    match check_feasible() {
        Ok(is_feasible) => {
            if is_feasible {
                let res = query::query_recent_with_ps_script(QuickAccess::All).await?;
                return Ok(res);
            } else {
                return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
            }
        },
        Err(e) => return Err(e),
    }
}

/****************************************************** Check Quick Access ******************************************************/

/// Checks if a given keyword is present in the list of recently accessed files.
///
/// # Parameters
/// - `keyword: &str`: The keyword to search for in the list of recent files.
///
/// # Example
/// ```
/// match is_in_recent_files("important").await {
///     Ok(true) => println!("The keyword 'important' was found in the recent files."),
///     Ok(false) => println!("The keyword 'important' was not found in the recent files."),
///     Err(e) => println!("Error checking recent files: {}", e),
/// }
/// ```
pub async fn is_in_recent_files(keyword: &str) -> Result<bool, WincentError> {
    let items = get_recent_files().await?;

    Ok(items.iter().any(|item| {item.contains(keyword) }))
}

/// Checks if a given keyword is present in the list of frequently accessed folders.
///
/// # Parameters
/// - `keyword: &str`: The keyword to search for in the list of frequent folders.
///
/// # Example
/// ```
/// match is_in_frequent_folders("documents").await {
///     Ok(true) => println!("The keyword 'documents' was found in the frequent folders."),
///     Ok(false) => println!("The keyword 'documents' was not found in the frequent folders."),
///     Err(e) => println!("Error checking frequent folders: {}", e),
/// }
/// ```
pub async fn is_in_frequent_folders(keyword: &str) -> Result<bool, WincentError> {
    let items = get_frequent_folders().await?;

    Ok(items.iter().any(|item| {item.contains(keyword) }))
}

/// Checks if a given keyword is present in the list of all quick access items.
///
/// # Parameters
/// - `keyword: &str`: The keyword to search for in the list of quick access items.
///
/// # Example
/// ```
/// match is_in_quick_access("project").await {
///     Ok(true) => println!("The keyword 'project' was found in the quick access items."),
///     Ok(false) => println!("The keyword 'project' was not found in the quick access items."),
///     Err(e) => println!("Error checking quick access items: {}", e),
/// }
/// ```
pub async fn is_in_quick_access(keyword: &str) -> Result<bool, WincentError> {
    let items = get_quick_access_items().await?;

    Ok(items.iter().any(|item| {item.contains(keyword) }))
}

/****************************************************** Handle Quick Access ******************************************************/

/// Removes a file from the list of recently accessed files.
///
/// # Parameters
/// - `path: &str`: The file path to remove from the list of recently accessed files.
/// 
/// # Example
/// ```
/// match remove_from_recent_files("C:\\Users\\user\\Documents\\important.txt").await {
///     Ok(()) => println!("File removed from recent files successfully."),
///     Err(e) => println!("Error removing file from recent files: {}", e),
/// }
/// ```
pub async fn remove_from_recent_files(path: &str) -> Result<(), WincentError> {
    use std::fs;
    use std::path::Path;

    if fs::metadata(path).is_err() {
        return Err(WincentError::IoError(std::io::ErrorKind::NotFound.into()));
    }

    if !Path::new(path).is_file() {
        return Err(WincentError::IoError(std::io::ErrorKind::InvalidData.into()));
    }

    match check_feasible() {
        Ok(is_feasible) => {
            if is_feasible {
                let in_quick_access = match crate::is_in_quick_access(path).await {
                    Ok(result) => result,
                    Err(e) => return Err(e),
                };
            
                if !in_quick_access {
                    return Ok(());
                }
            
                crate::handle::handle_recent_files_with_ps_script(path, true).await?;
                return Ok(());
            } else {
                return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
            }
        },
        Err(e) => return Err(e),
    }
}

/// Adds a folder to the list of frequently accessed folders.
///
/// # Parameters
/// - `path: &str`: The folder path to add to the list of frequently accessed folders.
///
/// # Example
/// ```
/// match add_to_frequent_folders("C:\\Users\\user\\Documents").await {
///     Ok(()) => println!("Folder added to frequent folders successfully."),
///     Err(e) => println!("Error adding folder to frequent folders: {}", e),
/// }
/// ```
pub async fn add_to_frequent_folders(path: &str) -> Result<(), WincentError> {
    if let Err(e) = std::fs::metadata(path) {
        return Err(WincentError::IoError(e));
    }

    if !std::path::Path::new(path).is_dir() {
        return Err(WincentError::IoError(std::io::ErrorKind::InvalidData.into()));
    }

    match check_feasible() {
        Ok(is_feasible) => {
            if is_feasible {
                crate::handle::handle_frequent_folders_with_ps_script(path).await?;
                return Ok(());
            } else {
                return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
            }
        },
        Err(e) => return Err(e),
    }
}

/// Removes a folder from the list of frequently accessed folders.
///
/// # Parameters
/// - `path: &str`: The folder path to remove from the list of frequently accessed folders.
///
/// # Example
/// ```
/// match remove_from_frequent_folders("C:\\Users\\user\\Documents").await {
///     Ok(()) => println!("Folder removed from frequent folders successfully."),
///     Err(e) => println!("Error removing folder from frequent folders: {}", e),
/// }
/// ```
pub async fn remove_from_frequent_folders(path: &str) -> Result<(), WincentError> {
    if let Err(e) = std::fs::metadata(path) {
        return Err(WincentError::IoError(e));
    }

    if !std::path::Path::new(path).is_dir() {
        return Err(WincentError::IoError(std::io::ErrorKind::InvalidData.into()));
    }

    match check_feasible() {
        Ok(is_feasible) => {
            if is_feasible {
                let is_in_quick_access = match crate::is_in_quick_access(path).await {
                    Ok(result) => result,
                    Err(e) => return Err(e),
                };

                let is_exist = crate::is_in_frequent_folders(path).await?;
                if is_exist {
                    // if target folder already exist in Frequent Folders, there will be two conditions
                    // 1. the target folder is a pinned folder, then we just need to do it once
                    // 2. the target folder is not a pinned folder, but a frequent one, then we have to do it twice, sometimes, `removefromhome` not works
                    crate::handle::handle_frequent_folders_with_ps_script(path).await?;

                    let is_frequent = crate::is_in_frequent_folders(path).await?;
                    if is_frequent {
                        crate::handle::handle_frequent_folders_with_ps_script(path).await?;
                    }
                }
            
                if is_in_quick_access {
                    crate::handle::handle_frequent_folders_with_ps_script(path).await?;
                }
                return Ok(());
            } else {
                return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
            }
        },
        Err(e) => return Err(e),
    }
}

/****************************************************** Quick Access Visiablity ******************************************************/

/// Checks if the list of recently accessed files is visible.
///
/// # Parameters
/// None
///
/// # Example
/// ```
/// match is_recent_files_visiable() {
///     Ok(true) => println!("The list of recent files is visible."),
///     Ok(false) => println!("The list of recent files is not visible."),
///     Err(e) => println!("Error checking recent files visibility: {}", e),
/// }
/// ```
pub fn is_recent_files_visiable() -> Result<bool, WincentError> {
    is_visialbe_with_registry(QuickAccess::RecentFiles)
}

/// Checks if the list of frequently accessed folders is visible.
///
/// # Parameters
/// None
///
/// # Example
/// ```
/// match is_frequent_folders_visible() {
///     Ok(true) => println!("The list of frequent folders is visible."),
///     Ok(false) => println!("The list of frequent folders is not visible."),
///     Err(e) => println!("Error checking frequent folders visibility: {}", e),
/// }
/// ```
pub fn is_frequent_folders_visible() -> Result<bool, WincentError> {
    is_visialbe_with_registry(QuickAccess::FrequentFolders)
}

/// Sets the visibility of the list of recently accessed files.
///
/// # Parameters
/// - `is_visiable: bool`: The desired visibility state of the list of recently accessed files. `true` to make the list visible, `false` to make it not visible.
///
/// # Example
/// ```
/// match set_recent_files_visiable(true) {
///     Ok(()) => println!("The list of recent files is now visible."),
///     Err(e) => println!("Error setting recent files visibility: {}", e),
/// }
/// ```
pub fn set_recent_files_visiable(is_visiable: bool) -> Result<(), WincentError> {
    set_visiable_with_registry(QuickAccess::RecentFiles, is_visiable)
}

/// Sets the visibility of the list of frequently accessed folders.
///
/// # Parameters
/// - `is_visiable: bool`: The desired visibility state of the list of frequently accessed folders. `true` to make the list visible, `false` to make it not visible.
///
/// # Example
/// ```
/// match set_frequent_folders_visiable(true) {
///     Ok(()) => println!("The list of frequent folders is now visible."),
///     Err(e) => println!("Error setting frequent folders visibility: {}", e),
/// }
/// ```
pub fn set_frequent_folders_visiable(is_visiable: bool) -> Result<(), WincentError> {
    set_visiable_with_registry(QuickAccess::FrequentFolders, is_visiable)
}

#[cfg(test)]
mod tests {
    use utils::init_test_logger;
    use log::{debug, error};
    use super::*;

    #[test]
    fn test_feasible() -> Result<(), WincentError> {
        init_test_logger();

        let is_feasible = check_feasible()?;
        if is_feasible {
            debug!("functions feasible to run");
        } else {
            debug!("try fix feasible");
            let _ = fix_feasible()?;
            let fix_res = check_feasible()?;
            if fix_res {
                debug!("fix feasible success!");
            } else {
                error!("failed to fix feasible!!!");
            }
        }

        Ok(())
    }

    #[tokio::test]
    async fn test_query_quick_access() -> Result<(), WincentError> {
        let recent_files: Vec<String> = get_recent_files().await?;
        let frequent_folders: Vec<String> = get_frequent_folders().await?;
        let quick_access: Vec<String> = get_quick_access_items().await?;
    
        debug!("recent files");
        for (idx, item) in recent_files.iter().enumerate() {
            debug!("{}. {}", idx, item);
        }
        debug!("\n\n");
    
        debug!("frequent folders");
        for (idx, item) in frequent_folders.iter().enumerate() {
            debug!("{}. {}", idx, item);
        }
        debug!("\n\n");
    
        debug!("quick access items");
        for (idx, item) in quick_access.iter().enumerate() {
            debug!("{}. {}", idx, item);
        }

        Ok(())
    }

    #[ignore]
    #[tokio::test]
    async fn test_check_handle_quick_access() {
        // see detail in `handle` module
        assert!(true);
    }

    #[ignore]
    #[test]
    fn test_quick_access_visible() {
        // see detail in `visible` module
        assert!(true);
    }
}
