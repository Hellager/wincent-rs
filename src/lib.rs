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

mod utils;
mod feasible;
mod query;
mod visible;
mod handle;
mod scripts;
pub mod error;

use crate::{
    error::WincentError,
    visible::{is_visialbe_with_registry, set_visiable_with_registry},
};

pub(crate) enum QuickAccess {
    FrequentFolders,
    RecentFiles,
    All
}

pub type WincentResult<T> = Result<T, WincentError>; 

/****************************************************** Feature Feasible ******************************************************/

/// Checks if the PowerShell execution policy is feasible based on the registry settings.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let is_feasible = check_feasible()?;
///     if is_feasible {
///         println!("The execution policy is feasible.");
///     } else {
///         println!("The execution policy is not feasible.");
///     }
///     Ok(())
/// }
/// ```
pub fn check_feasible() -> WincentResult<bool> {
    feasible::check_script_feasible_with_registry()
}

/// Fixes the PowerShell execution policy to ensure it is feasible based on the registry settings.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let is_fixed = fix_feasible()?;
///     if is_fixed {
///         println!("The execution policy has been fixed and is now feasible.");
///     } else {
///         println!("The execution policy is still not feasible.");
///     }
///     Ok(())
/// }
/// ```
pub fn fix_feasible() -> WincentResult<bool> {
    let _ = feasible::fix_script_feasible_with_registry()?;
    Ok(check_feasible()?)
}

/****************************************************** Query Quick Access ******************************************************/

/// Retrieves a list of recent files from Quick Access.
///
/// # Returns
/// 
/// Returns the full paths in Vec<String> if success.
/// 
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let recent_files = get_recent_files()?;
///     for file in recent_files {
///         println!("{}", file);
///     }
///     Ok(())
/// }
/// ```
pub fn get_recent_files() -> WincentResult<Vec<String>> {
    match check_feasible() {
        Ok(is_feasible) => {
            if is_feasible {
                let res = query::query_recent_with_ps_script(QuickAccess::RecentFiles)?;
                return Ok(res);
            } else {
                let error = std::io::ErrorKind::PermissionDenied;
                return Err(WincentError::Io(error.into()));
            }
        },
        Err(e) => return Err(e),
    }
}

/// Retrieves a list of frequently accessed folders from Quick Access.
///
/// # Returns
/// 
/// Returns the full paths in Vec<String> if success.
/// 
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let frequent_folders = get_frequent_folders()?;
///     for folder in frequent_folders {
///         println!("{}", folder);
///     }
///     Ok(())
/// }
/// ```
pub fn get_frequent_folders() -> WincentResult<Vec<String>> {
    match check_feasible() {
        Ok(is_feasible) => {
            if is_feasible {
                let res = query::query_recent_with_ps_script(QuickAccess::FrequentFolders)?;
                return Ok(res);
            } else {
                let error = std::io::ErrorKind::PermissionDenied;
                return Err(WincentError::Io(error.into()));
            }
        },
        Err(e) => return Err(e),
    }
}

/// Retrieves a list of items from Quick Access.
///
/// # Returns
/// 
/// Returns the full paths in Vec<String> if success.
/// 
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let quick_access_items = get_quick_access_items()?;
///     for item in quick_access_items {
///         println!("{}", item);
///     }
///     Ok(())
/// }
/// ```
pub fn get_quick_access_items() -> WincentResult<Vec<String>> {
    match check_feasible() {
        Ok(is_feasible) => {
            if is_feasible {
                let res = query::query_recent_with_ps_script(QuickAccess::All)?;
                return Ok(res);
            } else {
                let error = std::io::ErrorKind::PermissionDenied;
                return Err(WincentError::Io(error.into()));
            }
        },
        Err(e) => return Err(e),
    }
}

/****************************************************** Check Quick Access ******************************************************/

/// Checks if a given keyword is present in the recent files.
///
/// # Parameters
///
/// - `keyword`: A string slice that represents the keyword to search for in the recent files.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let keyword = "example.txt";
///     let found = is_in_recent_files(keyword)?;
///     if found {
///         println!("The keyword '{}' is in the recent files.", keyword);
///     } else {
///         println!("The keyword '{}' is not found in the recent files.", keyword);
///     }
///     Ok(())
/// }
/// ```
pub fn is_in_recent_files(keyword: &str) -> WincentResult<bool> {
    let items = get_recent_files()?;

    Ok(items.iter().any(|item| {item.contains(keyword) }))
}

/// Checks if a given keyword is present in the frequently accessed folders.
///
/// # Parameters
///
/// - `keyword`: A string slice that represents the keyword to search for in the frequently accessed folders.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let keyword = "Documents";
///     let found = is_in_frequent_folders(keyword)?;
///     if found {
///         println!("The keyword '{}' is in the frequent folders.", keyword);
///     } else {
///         println!("The keyword '{}' is not found in the frequent folders.", keyword);
///     }
///     Ok(())
/// }
/// ```
pub fn is_in_frequent_folders(keyword: &str) -> WincentResult<bool> {
    let items = get_frequent_folders()?;

    Ok(items.iter().any(|item| {item.contains(keyword) }))
}

/// Checks if a given keyword is present in the Quick Access items.
///
/// # Parameters
///
/// - `keyword`: A string slice that represents the keyword to search for in the Quick Access items.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let keyword = "Project";
///     let found = is_in_quick_access(keyword)?;
///     if found {
///         println!("The keyword '{}' is in Quick Access.", keyword);
///     } else {
///         println!("The keyword '{}' is not found in Quick Access.", keyword);
///     }
///     Ok(())
/// }
/// ```
pub fn is_in_quick_access(keyword: &str) -> WincentResult<bool> {
    let items = get_quick_access_items()?;

    Ok(items.iter().any(|item| {item.contains(keyword) }))
}

/****************************************************** Handle Quick Access ******************************************************/

/// Removes a specified file from the recent files list.
///
/// # Parameters
///
/// - `path`: A string slice that represents the path of the file to be removed from the recent files.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let path = "example.txt";
///     remove_from_recent_files(path)?;
///     println!("Removed '{}' from recent files.", path);
///     Ok(())
/// }
/// ```
pub fn remove_from_recent_files(path: &str) -> WincentResult<()> {
    use std::path::Path;

    if let Err(e) = std::fs::metadata(path) {
        return Err(WincentError::Io(e));
    }

    if !Path::new(path).is_file() {
        let error = std::io::ErrorKind::InvalidData;
        return Err(WincentError::Io(error.into()));
    }

    match check_feasible() {
        Ok(is_feasible) => {
            if is_feasible {
                let in_quick_access = match crate::is_in_quick_access(path) {
                    Ok(result) => result,
                    Err(e) => return Err(e),
                };
            
                if !in_quick_access {
                    return Ok(());
                }
            
                crate::handle::remove_recent_files_with_ps_script(path)?;
                return Ok(());
            } else {
                let error = std::io::ErrorKind::PermissionDenied;
                return Err(WincentError::Io(error.into()));
            }
        },
        Err(e) => return Err(e),
    }
}

/// Adds a specified folder to the list of frequently accessed folders.
///
/// # Parameters
///
/// - `path`: A string slice that represents the path of the folder to be added to the frequent folders.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let path = "C:/Users/Example/Documents";
///     add_to_frequent_folders(path)?;
///     println!("Added '{}' to frequent folders.", path);
///     Ok(())
/// }
/// ```
pub fn add_to_frequent_folders(path: &str) -> WincentResult<()> {
    if let Err(e) = std::fs::metadata(path) {
        return Err(WincentError::Io(e));
    }

    if !std::path::Path::new(path).is_dir() {
        let error = std::io::ErrorKind::InvalidData;
        return Err(WincentError::Io(error.into()));
    }

    match check_feasible() {
        Ok(is_feasible) => {
            if is_feasible {
                crate::handle::pin_frequent_folder_with_ps_script(path)?;
                return Ok(());
            } else {
                let error = std::io::ErrorKind::PermissionDenied;
                return Err(WincentError::Io(error.into()));
            }
        },
        Err(e) => return Err(e),
    }
}

/// Removes a specified folder from the list of frequently accessed folders.
///
/// # Parameters
///
/// - `path`: A string slice that represents the path of the folder to be removed from the frequent folders.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let path = "C:/Users/Example/Documents";
///     remove_from_frequent_folders(path)?;
///     println!("Removed '{}' from frequent folders.", path);
///     Ok(())
/// }
/// ```
pub fn remove_from_frequent_folders(path: &str) -> WincentResult<()> {
    if let Err(e) = std::fs::metadata(path) {
        return Err(WincentError::Io(e));
    }

    if !std::path::Path::new(path).is_dir() {
        let error = std::io::ErrorKind::InvalidData;
        return Err(WincentError::Io(error.into()));
    }

    match check_feasible() {
        Ok(is_feasible) => {
            if is_feasible {
                let is_in_quick_access = match crate::is_in_quick_access(path) {
                    Ok(result) => result,
                    Err(e) => return Err(e),
                };

                if is_in_quick_access {
                    crate::handle::unpin_frequent_folder_with_ps_script(path)?;
                }
                return Ok(());
            } else {
                let error = std::io::ErrorKind::PermissionDenied;
                return Err(WincentError::Io(error.into()));
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
pub fn is_recent_files_visiable() -> WincentResult<bool> {
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
pub fn is_frequent_folders_visible() -> WincentResult<bool> {
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
pub fn set_recent_files_visiable(is_visiable: bool) -> WincentResult<()> {
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
pub fn set_frequent_folders_visiable(is_visiable: bool) -> WincentResult<()> {
    set_visiable_with_registry(QuickAccess::FrequentFolders, is_visiable)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test_log::test]
    fn test_feasible() -> Result<(), WincentError> {
        let is_feasible = check_feasible()?;
        if is_feasible {
            println!("functions feasible to run");
        } else {
            println!("try fix feasible");
            let _ = fix_feasible()?;
            let fix_res = check_feasible()?;
            if fix_res {
                println!("fix feasible success!");
            } else {
                println!("failed to fix feasible!!!");
            }
        }

        Ok(())
    }

    #[test_log::test]
    fn test_query_quick_access() -> Result<(), WincentError> {
        let recent_files: Vec<String> = get_recent_files()?;
        let frequent_folders: Vec<String> = get_frequent_folders()?;
        let quick_access: Vec<String> = get_quick_access_items()?;
    
        println!("recent files");
        for (idx, item) in recent_files.iter().enumerate() {
            println!("{}. {}", idx, item);
        }
        println!("\n\n");
    
        println!("frequent folders");
        for (idx, item) in frequent_folders.iter().enumerate() {
            println!("{}. {}", idx, item);
        }
        println!("\n\n");
    
        println!("quick access items");
        for (idx, item) in quick_access.iter().enumerate() {
            println!("{}. {}", idx, item);
        }

        Ok(())
    }

    #[ignore]
    #[test]
    fn test_check_handle_quick_access() {
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
