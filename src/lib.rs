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
//! ```rust
//! use wincent::{check_feasible, fix_feasible, get_quick_access_items};
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
//! ```rust
//! use wincent::{
//!     add_to_frequent_folders,
//!     is_in_quick_access,
//!     remove_from_recent_files,
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
//! - Async support
//! - Comprehensive error handling
//! - PowerShell script integration
//! - Registry management
//! - Windows API integration
//! - Cross-version Windows support
//!

mod utils;
mod feasible;
mod query;
mod visible;
mod handle;
mod scripts;
mod test_utils;
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

/// Checks if PowerShell script execution is feasible on the current system.
///
/// # Returns
///
/// Returns `true` if script execution is allowed, `false` otherwise.
///
/// # Example
///
/// ```rust
/// use wincent::check_script_feasible;
///
/// fn main() -> Result<(), WincentError> {
///     if check_script_feasible()? {
///         println!("PowerShell scripts can be executed");
///     } else {
///         println!("PowerShell script execution is restricted");
///     }
///     Ok(())
/// }
/// ```
pub fn check_script_feasible() -> WincentResult<bool> {
    feasible::check_script_feasible_with_registry()
}

/// Fixes PowerShell script execution policy to allow script execution.
///
/// # Example
///
/// ```rust
/// use wincent::{check_script_feasible, fix_script_feasible};
///
/// fn main() -> Result<(), WincentError> {
///     if !check_script_feasible()? {
///         fix_script_feasible()?;
///         assert!(check_script_feasible()?);
///     }
///     Ok(())
/// }
/// ```
pub fn fix_script_feasible() -> WincentResult<()> {
    feasible::fix_script_feasible_with_registry()
}

/// Checks if Quick Access query operations are feasible on the current system.
///
/// # Returns
///
/// Returns `true` if Quick Access query operations are supported, `false` otherwise.
///
/// # Example
///
/// ```rust
/// use wincent::check_query_feasible;
///
/// fn main() -> Result<(), WincentError> {
///     if check_query_feasible()? {
///         println!("Quick Access query operations are supported");
///     } else {
///         println!("Quick Access query operations are not supported");
///     }
///     Ok(())
/// }
/// ```
pub fn check_query_feasible() -> WincentResult<bool> {
    feasible::check_query_feasible_with_script()
}

/// Checks if pin/unpin operations are feasible on the current system.
///
/// # Returns
///
/// Returns `true` if pin/unpin operations are supported, `false` otherwise.
///
/// # Example
///
/// ```rust
/// use wincent::check_pinunpin_feasible;
///
/// fn main() -> Result<(), WincentError> {
///     if check_pinunpin_feasible()? {
///         println!("Pin/unpin operations are supported");
///     } else {
///         println!("Pin/unpin operations are not supported");
///     }
///     Ok(())
/// }
/// ```
pub fn check_pinunpin_feasible() -> WincentResult<bool> {
    feasible::check_pinunpin_feasible_with_script()
}

/// Checks if all Quick Access operations are feasible on the current system.
///
/// # Returns
///
/// Returns `true` only if all operations are supported, `false` otherwise.
///
/// # Example
///
/// ```rust
/// use wincent::{check_feasible, fix_feasible};
///
/// fn main() -> Result<(), WincentError> {
///     if !check_feasible()? {
///         println!("Some Quick Access operations are not supported");
///         // Try to fix the issues
///         if fix_feasible()? {
///             println!("Successfully enabled Quick Access operations");
///         }
///     }
///     Ok(())
/// }
/// ```
pub fn check_feasible() -> WincentResult<bool> {
    // First check script execution policy
    if !check_script_feasible()? {
        return Ok(false);
    }

    // Then check both operations
    let query_ok = check_query_feasible()?;
    let pinunpin_ok = check_pinunpin_feasible()?;

    Ok(query_ok && pinunpin_ok)
}

/// Attempts to fix Quick Access operation feasibility issues.
///
/// # Returns
///
/// Returns `true` if all operations are successfully enabled, `false` otherwise.
///
/// # Example
///
/// ```rust
/// use wincent::fix_feasible;
///
/// fn main() -> Result<(), WincentError> {
///     match fix_feasible()? {
///         true => println!("Successfully enabled Quick Access operations"),
///         false => println!("Failed to enable some Quick Access operations")
///     }
///     Ok(())
/// }
/// ```
pub fn fix_feasible() -> WincentResult<bool> {
    fix_script_feasible()?;
    check_feasible()
}

/****************************************************** Query Quick Access ******************************************************/

/// Gets a list of recent files from Windows Quick Access.
/// 
/// # Returns
///
/// Returns a vector of file paths as strings.
/// 
/// # Example
///
/// ```rust
/// use wincent::get_recent_files;
///
/// fn main() -> Result<(), WincentError> {
///     let recent_files = get_recent_files()?;
///     for file in recent_files {
///         println!("Recent file: {}", file);
///     }
///     Ok(())
/// }
/// ```
pub fn get_recent_files() -> WincentResult<Vec<String>> {
    if !check_script_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "PowerShell script execution is not feasible".to_string()
        ));
    }

    if !check_query_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "Quick Access query operation is not feasible".to_string()
        ));
    }

    query::query_recent_with_ps_script(QuickAccess::RecentFiles)
}

/// Gets a list of frequent folders from Windows Quick Access.
/// 
/// # Returns
///
/// Returns a vector of folder paths as strings.
///
/// # Example
///
/// ```rust
/// use wincent::get_frequent_folders;
///
/// fn main() -> Result<(), WincentError> {
///     let folders = get_frequent_folders()?;
///     for folder in folders {
///         println!("Frequent folder: {}", folder);
///     }
///     Ok(())
/// }
/// ```
pub fn get_frequent_folders() -> WincentResult<Vec<String>> {
    if !check_script_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "PowerShell script execution is not feasible".to_string()
        ));
    }

    if !check_query_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "Quick Access query operation is not feasible".to_string()
        ));
    }

    query::query_recent_with_ps_script(QuickAccess::FrequentFolders)
}

/// Gets a list of all items from Windows Quick Access, including both recent files and frequent folders.
///
/// # Returns
///
/// Returns a vector of strings containing the paths of all Quick Access items.
///
/// # Example
///
/// ```rust
/// use wincent::get_quick_access_items;
///
/// fn main() -> Result<(), WincentError> {
///     match get_quick_access_items() {
///         Ok(items) => {
///             println!("Found {} Quick Access items:", items.len());
///             for item in items {
///                 println!("  - {}", item);
///             }
///         },
///         Err(e) => println!("Failed to get Quick Access items: {}", e)
///     }
///     Ok(())
/// }
/// ```
pub fn get_quick_access_items() -> WincentResult<Vec<String>> {
    if !check_script_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "PowerShell script execution is not feasible".to_string()
        ));
    }

    if !check_query_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "Quick Access query operation is not feasible".to_string()
        ));
    }

    query::query_recent_with_ps_script(QuickAccess::All)
}


/****************************************************** Check Quick Access ******************************************************/

/// Checks if a file path exists in the Windows Recent Files list.
///
/// # Arguments
///
/// * `keyword` - The file path or partial path to search for
///
/// # Returns
///
/// Returns `true` if the file is found in the recent files list.
///
/// # Example
///
/// ```rust
/// use wincent::is_in_recent_files;
///
/// fn main() -> Result<(), WincentError> {
///     let file_exists = is_in_recent_files("Documents\\report.docx")?;
///     if file_exists {
///         println!("File found in recent files");
///     }
///     Ok(())
/// }
/// ```
pub fn is_in_recent_files(keyword: &str) -> WincentResult<bool> {
    let items = get_recent_files()?;

    Ok(items.iter().any(|item| {item.contains(keyword) }))
}

/// Checks if a folder path exists in the Windows Frequent Folders list.
///
/// # Arguments
///
/// * `keyword` - The folder path or partial path to search for
///
/// # Returns
///
/// Returns `true` if the folder is found in the frequent folders list.
///
/// # Example
///
/// ```rust
/// use wincent::is_in_frequent_folders;
///
/// fn main() -> Result<(), WincentError> {
///     let folder_exists = is_in_frequent_folders("Projects")?;
///     if folder_exists {
///         println!("Found folder in frequent folders list");
///     } else {
///         println!("Folder not found in frequent folders list");
///     }
///     Ok(())
/// }
/// ```
pub fn is_in_frequent_folders(keyword: &str) -> WincentResult<bool> {
    let items = get_frequent_folders()?;

    Ok(items.iter().any(|item| {item.contains(keyword) }))
}

/// Checks if a path exists in the Windows Quick Access list.
///
/// # Arguments
///
/// * `keyword` - The path or partial path to search for
///
/// # Returns
///
/// Returns `true` if the path is found in either recent files or frequent folders.
///
/// # Example
///
/// ```rust
/// use wincent::is_in_quick_access;
///
/// fn main() -> Result<(), WincentError> {
///     // Check for a specific file or folder
///     if is_in_quick_access("Documents")? {
///         println!("Found item in Quick Access");
///     }
///
///     // Check for items in a specific location
///     if is_in_quick_access("C:\\Projects\\")? {
///         println!("Found items from Projects folder");
///     }
///     Ok(())
/// }
/// ```
pub fn is_in_quick_access(keyword: &str) -> WincentResult<bool> {
    let items = get_quick_access_items()?;

    Ok(items.iter().any(|item| {item.contains(keyword) }))
}

/****************************************************** Handle Quick Access ******************************************************/

/// Adds a file to Windows Recent Files.
///
/// # Arguments
///
/// * `path` - The full path to the file to be added
///
/// # Example
///
/// ```rust
/// use wincent::add_to_recent_files;
///
/// fn main() -> Result<(), WincentError> {
///     add_to_recent_files("C:\\Documents\\report.docx")?;
///     Ok(())
/// }
/// ```
pub fn add_to_recent_files(path: &str) -> WincentResult<()> {
    if !std::path::Path::new(path).is_file() {
        return Err(WincentError::InvalidPath(format!("Not a valid file: {}", path)));
    }

    handle::add_file_to_recent_with_api(path)
}

/// Removes a file from Windows Recent Files.
///
/// # Arguments
///
/// * `path` - The full path to the file to be removed
///
/// # Example
///
/// ```rust
/// use wincent::remove_from_recent_files;
///
/// fn main() -> Result<(), WincentError> {
///     remove_from_recent_files("C:\\Documents\\report.docx")?;
///     Ok(())
/// }
/// ```
pub fn remove_from_recent_files(path: &str) -> WincentResult<()> {
    if !std::path::Path::new(path).is_file() {
        return Err(WincentError::InvalidPath(format!("Not a valid file: {}", path)));
    }

    if !check_script_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "PowerShell script execution is not feasible".to_string()
        ));
    }

    handle::remove_recent_files_with_ps_script(path)
}

/// Pins a folder to Windows Quick Access.
///
/// # Arguments
///
/// * `path` - The full path to the folder to be pinned
///
/// # Returns
///
/// Returns `Ok(())` if the folder was successfully pinned.
///
/// # Example
///
/// ```rust
/// use wincent::add_to_frequent_folders;
///
/// fn main() -> Result<(), WincentError> {
///     // Pin a project folder
///     add_to_frequent_folders("C:\\Projects\\my-project")?;
///     Ok(())
/// }
/// ```
pub fn add_to_frequent_folders(path: &str) -> WincentResult<()> {
    if !std::path::Path::new(path).is_dir() {
        return Err(WincentError::InvalidPath(format!("Not a valid directory: {}", path)));
    }

    if !check_script_feasible()? || !check_pinunpin_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "Pin operation is not feasible".to_string()
        ));
    }

    handle::pin_frequent_folder_with_ps_script(path)
}

/// Unpins a folder from Windows Quick Access.
///
/// # Arguments
///
/// * `path` - The full path to the folder to be unpinned
///
/// # Returns
///
/// Returns `Ok(())` if the folder was successfully unpinned.
///
/// # Example
///
/// ```rust
/// use wincent::remove_from_frequent_folders;
///
/// fn main() -> Result<(), WincentError> {
///     // Unpin a project folder
///     remove_from_frequent_folders("C:\\Projects\\old-project")?;
///     Ok(())
/// }
/// ```
pub fn remove_from_frequent_folders(path: &str) -> WincentResult<()> {
    if !std::path::Path::new(path).is_dir() {
        return Err(WincentError::InvalidPath(format!("Not a valid directory: {}", path)));
    }

    if !check_script_feasible()? || !check_pinunpin_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "Unpin operation is not feasible".to_string()
        ));
    }

    handle::unpin_frequent_folder_with_ps_script(path)
}

/****************************************************** Quick Access Visiablity ******************************************************/

/// Checks if Quick Access visibility settings can be modified.
///
/// # Returns
///
/// Returns `true` if Quick Access visibility can be controlled.
///
/// # Example
///
/// ```rust
/// use wincent::{is_recent_files_visiable, set_recent_files_visiable};
///
/// fn main() -> Result<(), WincentError> {
///     let is_visible = is_recent_files_visiable()?;
///     if !is_visible {
///         set_recent_files_visiable(true)?;
///     }
///     Ok(())
/// }
/// ```
pub fn is_recent_files_visiable() -> WincentResult<bool> {
    is_visialbe_with_registry(QuickAccess::RecentFiles)
}

/// Checks if frequent folders are visible in Windows Quick Access.
///
/// # Returns
///
/// Returns `true` if frequent folders are visible, `false` if they are hidden.
///
/// # Example
///
/// ```rust
/// use wincent::{is_frequent_folders_visible, set_frequent_folders_visiable};
///
/// fn main() -> Result<(), WincentError> {
///     let is_visible = is_frequent_folders_visible()?;
///     println!("Frequent folders are {}", if is_visible { "visible" } else { "hidden" });
///     
///     // Ensure frequent folders are visible
///     if !is_visible {
///         set_frequent_folders_visiable(true)?;
///     }
///     Ok(())
/// }
/// ```
pub fn is_frequent_folders_visible() -> WincentResult<bool> {
    is_visialbe_with_registry(QuickAccess::FrequentFolders)
}

/// Sets the visibility of Quick Access recent files.
///
/// # Arguments
///
/// * `is_visiable` - Whether recent files should be visible
///
/// # Example
///
/// ```rust
/// use wincent::set_recent_files_visiable;
///
/// fn main() -> Result<(), WincentError> {
///     // Hide recent files in Quick Access
///     set_recent_files_visiable(false)?;
///     Ok(())
/// }
/// ```
pub fn set_recent_files_visiable(is_visiable: bool) -> WincentResult<()> {
    set_visiable_with_registry(QuickAccess::RecentFiles, is_visiable)
}

/// Sets the visibility of frequent folders in Windows Quick Access.
///
/// # Arguments
///
/// * `is_visiable` - `true` to show frequent folders, `false` to hide them
///
/// # Returns
///
/// Returns `Ok(())` if the visibility was successfully changed.
///
/// # Example
///
/// ```rust
/// use wincent::set_frequent_folders_visiable;
///
/// fn main() -> Result<(), WincentError> {
///     // Hide frequent folders in Quick Access
///     set_frequent_folders_visiable(false)?;
///     println!("Frequent folders are now hidden");
///     Ok(())
/// }
/// ```
pub fn set_frequent_folders_visiable(is_visiable: bool) -> WincentResult<()> {
    set_visiable_with_registry(QuickAccess::FrequentFolders, is_visiable)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_utils::{setup_test_env, create_test_file, cleanup_test_env};
    use std::{thread, time::Duration};

    #[test_log::test]
    fn test_feasibility_checks() -> WincentResult<()> {
        // Test script execution feasibility
        let script_feasible = check_script_feasible()?;
        println!("Script execution feasible: {}", script_feasible);

        // Test query feasibility
        let query_feasible = check_query_feasible()?;
        println!("Query operation feasible: {}", query_feasible);

        // Test pin/unpin feasibility
        let pinunpin_feasible = check_pinunpin_feasible()?;
        println!("Pin/Unpin operation feasible: {}", pinunpin_feasible);

        // If any check fails, try fixing
        if !script_feasible || !query_feasible || !pinunpin_feasible {
            println!("Attempting to fix feasibility...");
            fix_script_feasible()?;
            
            let fixed = check_feasible()?;
            if fixed {
                println!("Successfully fixed feasibility");
            } else {
                println!("Failed to fix feasibility");
            }
        }

        Ok(())
    }

    #[test]
    fn test_quick_access_operations() -> WincentResult<()> {
        let test_dir = setup_test_env()?;
        
        // Create test files
        let test_file = create_test_file(&test_dir, "test.txt", "test content")?;
        let test_path = test_file.to_str().unwrap();
        
        // Test adding to recent files
        add_to_recent_files(test_path)?;
        thread::sleep(Duration::from_millis(500));
        
        // Verify file was added
        assert!(is_in_recent_files(test_path)?, "File should be in recent files");
        
        // Test adding folder to frequent folders
        let dir_path = test_dir.to_str().unwrap();
        add_to_frequent_folders(dir_path)?;
        thread::sleep(Duration::from_millis(500));
        
        // Verify folder was added
        assert!(is_in_frequent_folders(dir_path)?, "Folder should be in frequent folders");
        
        // Test removal operations
        remove_from_recent_files(test_path)?;
        remove_from_frequent_folders(dir_path)?;
        thread::sleep(Duration::from_millis(500));
        
        // Verify removals
        assert!(!is_in_recent_files(test_path)?, "File should not be in recent files");
        assert!(!is_in_frequent_folders(dir_path)?, "Folder should not be in frequent folders");
        
        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    fn test_visibility_operations() -> WincentResult<()> {
        // Save initial states
        let initial_recent = is_recent_files_visiable()?;
        let initial_frequent = is_frequent_folders_visible()?;
        
        // Test visibility toggling
        set_recent_files_visiable(!initial_recent)?;
        set_frequent_folders_visiable(!initial_frequent)?;
        
        // Verify changes
        assert_eq!(!initial_recent, is_recent_files_visiable()?, "Recent files visibility should be toggled");
        assert_eq!(!initial_frequent, is_frequent_folders_visible()?, "Frequent folders visibility should be toggled");
        
        // Restore initial states
        set_recent_files_visiable(initial_recent)?;
        set_frequent_folders_visiable(initial_frequent)?;
        
        Ok(())
    }
}
