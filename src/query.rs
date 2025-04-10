//! Query Windows Quick Access items, including recent files
//! and frequent folders.
//!
//! ## Example
//!
//! ```no_run
//! use wincent::{
//!     feasible::{check_script_feasible, fix_script_feasible},
//!     query::{
//!         get_frequent_folders, get_quick_access_items, get_recent_files, is_in_frequent_folders,
//!         is_in_quick_access, is_in_recent_files,
//!     },
//!     WincentResult,
//! };
//!
//! fn print_items(title: &str, items: &[String]) {
//!     println!("\n=== {} ===", title);
//!     if items.is_empty() {
//!         println!("No items found");
//!     } else {
//!         for (idx, item) in items.iter().enumerate() {
//!             println!("{}. {}", idx + 1, item);
//!         }
//!     }
//!     println!("=== End of {} ===\n", title);
//! }
//!
//! fn main() -> WincentResult<()> {
//!     // Check and ensure script execution feasibility
//!     if !check_script_feasible()? {
//!         println!("Fixing script execution policy...");
//!         fix_script_feasible()?;
//!     }
//!
//!     // Get all Quick Access items
//!     println!("Querying Quick Access items...");
//!     let all_items = get_quick_access_items()?;
//!     print_items("All Quick Access Items", &all_items);
//!
//!     // Get recently used files
//!     let recent_files = get_recent_files()?;
//!     print_items("Recent Files", &recent_files);
//!
//!     // Get frequent folders
//!     let frequent_folders = get_frequent_folders()?;
//!     print_items("Frequent Folders", &frequent_folders);
//!
//!     // Search for specific keywords
//!     let keywords = ["Documents", "Downloads", "Desktop"];
//!     println!("\nSearching for specific keywords...");
//!
//!     for keyword in keywords {
//!         println!("\nChecking for keyword: '{}'", keyword);
//!
//!         if is_in_quick_access(keyword)? {
//!             println!("'{}' found in Quick Access", keyword);
//!
//!             if is_in_recent_files(keyword)? {
//!                 println!("'{}' found in Recent Files", keyword);
//!             }
//!
//!             if is_in_frequent_folders(keyword)? {
//!                 println!("'{}' found in Frequent Folders", keyword);
//!             }
//!         } else {
//!             println!("'{}' not found in Quick Access", keyword);
//!         }
//!     }
//!
//!     Ok(())
//! }
//! ```

use crate::{
    script_executor::ScriptExecutor,
    script_strategy::PSScript,
    QuickAccess, WincentResult,
};

/// Queries recent items from Quick Access using a PowerShell script.
pub(crate) fn query_recent_with_ps_script(qa_type: QuickAccess) -> WincentResult<Vec<String>> {
    let output = match qa_type {
        QuickAccess::All => ScriptExecutor::execute_ps_script(PSScript::QueryQuickAccess, None)?,
        QuickAccess::RecentFiles => ScriptExecutor::execute_ps_script(PSScript::QueryRecentFile, None)?,
        QuickAccess::FrequentFolders => ScriptExecutor::execute_ps_script(PSScript::QueryFrequentFolder, None)?,
    };

    let data = ScriptExecutor::parse_output_to_strings(output)?;

    Ok(data)
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
/// use wincent::{query::get_recent_files, error::WincentError};
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
    query_recent_with_ps_script(QuickAccess::RecentFiles)
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
/// use wincent::{query::get_frequent_folders, error::WincentError};
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
    query_recent_with_ps_script(QuickAccess::FrequentFolders)
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
/// use wincent::{query::get_quick_access_items, error::WincentError};
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
    query_recent_with_ps_script(QuickAccess::All)
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
/// use wincent::{query::is_in_recent_files, error::WincentError};
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

    Ok(items.iter().any(|item| item.contains(keyword)))
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
/// use wincent::{query::is_in_frequent_folders, error::WincentError};
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

    Ok(items.iter().any(|item| item.contains(keyword)))
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
/// use wincent::{query::is_in_quick_access, error::WincentError};
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

    Ok(items.iter().any(|item| item.contains(keyword)))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_query_recent_files() -> WincentResult<()> {
        let files = query_recent_with_ps_script(QuickAccess::RecentFiles)?;

        if !files.is_empty() {
            assert!(
                files.iter().all(|path| !path.is_empty()),
                "Paths should not be empty"
            );

            for path in &files {
                assert!(
                    path.contains(":\\"),
                    "Path should be a valid Windows path format: {}",
                    path
                );
            }
        }

        Ok(())
    }

    #[test]
    fn test_query_frequent_folders() -> WincentResult<()> {
        let folders = query_recent_with_ps_script(QuickAccess::FrequentFolders)?;

        if !folders.is_empty() {
            assert!(
                folders.iter().all(|path| !path.is_empty()),
                "Paths should not be empty"
            );

            for path in &folders {
                assert!(
                    path.contains(":\\"),
                    "Path should be a valid Windows path format: {}",
                    path
                );
            }
        }

        Ok(())
    }

    #[test_log::test]
    fn test_query_quick_access() -> WincentResult<()> {
        let items = query_recent_with_ps_script(QuickAccess::All)?;

        if !items.is_empty() {
            assert!(
                items.iter().all(|path| !path.is_empty()),
                "Paths should not be empty"
            );

            for path in &items {
                assert!(
                    path.contains(":\\"),
                    "Path should be a valid Windows path format: {}",
                    path
                );
            }
        }

        Ok(())
    }
}
