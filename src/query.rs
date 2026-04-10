//! Windows Quick Access item retrieval and inspection
//!
//! Provides read-only access to system Quick Access metadata including:
//! - Recent file tracking
//! - Frequent folder usage
//! - Combined access patterns
//!
//! # Key Functionality
//! - Full Quick Access inventory retrieval
//! - Category-specific queries
//! - Path existence verification
//! - PowerShell-based data collection
//!
//! # Data Characteristics
//! - Updated automatically by Windows Explorer
//! - Contains user-specific activity data
//! - Maximum 20 items per category (Windows default)

use crate::{
    script_executor::ScriptExecutor, script_strategy::PSScript, QuickAccess, WincentResult,
};

/// Queries recent items from Quick Access using a PowerShell script.
pub(crate) fn query_recent_with_ps_script(qa_type: QuickAccess) -> WincentResult<Vec<String>> {
    let output = match qa_type {
        QuickAccess::All => ScriptExecutor::execute_ps_script(PSScript::QueryQuickAccess, None)?,
        QuickAccess::RecentFiles => {
            ScriptExecutor::execute_ps_script(PSScript::QueryRecentFile, None)?
        }
        QuickAccess::FrequentFolders => {
            ScriptExecutor::execute_ps_script(PSScript::QueryFrequentFolder, None)?
        }
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

/// Checks if an exact file path exists in the Windows Recent Files list.
///
/// This function performs exact path comparison. For partial/fuzzy matching,
/// use `is_in_recent_files()` instead.
///
/// # Arguments
///
/// * `path` - The exact file path to search for
///
/// # Returns
///
/// Returns `true` if the exact path is found in the recent files list.
///
/// # Example
///
/// ```rust
/// use wincent::{query::is_recent_file_exact, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Exact match only
///     let exists = is_recent_file_exact("C:\\Users\\Documents\\file.txt")?;
///     if exists {
///         println!("Exact file path found in recent files");
///     }
///     Ok(())
/// }
/// ```
pub fn is_recent_file_exact(path: &str) -> WincentResult<bool> {
    let items = get_recent_files()?;
    Ok(items.iter().any(|item| item == path))
}

/// Checks if a file path or keyword exists in the Windows Recent Files list.
///
/// **Note**: This function performs substring matching (fuzzy match). If you need
/// exact path matching, use `is_recent_file_exact()` instead.
///
/// # Arguments
///
/// * `keyword` - The file path or partial path to search for (substring match)
///
/// # Returns
///
/// Returns `true` if any recent file path contains the keyword.
///
/// # Example
///
/// ```rust
/// use wincent::{query::is_in_recent_files, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Fuzzy match - matches any path containing "Documents"
///     let file_exists = is_in_recent_files("Documents")?;
///
///     // This will match paths like:
///     // - "C:\\Users\\Documents\\report.docx"
///     // - "D:\\My Documents\\file.txt"
///
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

/// Checks if an exact folder path exists in the Windows Frequent Folders list.
///
/// This function performs exact path comparison. For partial/fuzzy matching,
/// use `is_in_frequent_folders()` instead.
///
/// # Arguments
///
/// * `path` - The exact folder path to search for
///
/// # Returns
///
/// Returns `true` if the exact path is found in the frequent folders list.
///
/// # Example
///
/// ```rust
/// use wincent::{query::is_frequent_folder_exact, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     let folder_exists = is_frequent_folder_exact("C:\\Users\\Documents")?;
///     if folder_exists {
///         println!("Exact folder path found in frequent folders list");
///     }
///     Ok(())
/// }
/// ```
pub fn is_frequent_folder_exact(path: &str) -> WincentResult<bool> {
    let items = get_frequent_folders()?;
    Ok(items.iter().any(|item| item == path))
}

/// Checks if a folder path or keyword exists in the Windows Frequent Folders list.
///
/// **Note**: This function performs substring matching (fuzzy match). If you need
/// exact path matching, use `is_frequent_folder_exact()` instead.
///
/// # Arguments
///
/// * `keyword` - The folder path or partial path to search for (substring match)
///
/// # Returns
///
/// Returns `true` if any frequent folder path contains the keyword.
///
/// # Example
///
/// ```rust
/// use wincent::{query::is_in_frequent_folders, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Fuzzy match - matches any path containing "Projects"
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

/// Checks if an exact path exists in the Windows Quick Access list.
///
/// This function performs exact path comparison. For partial/fuzzy matching,
/// use `is_in_quick_access()` instead.
///
/// # Arguments
///
/// * `path` - The exact path to search for
///
/// # Returns
///
/// Returns `true` if the exact path is found in either recent files or frequent folders.
///
/// # Example
///
/// ```rust
/// use wincent::{query::is_in_quick_access_exact, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     let exists = is_in_quick_access_exact("C:\\Users\\Documents\\file.txt")?;
///     if exists {
///         println!("Exact path found in Quick Access");
///     }
///     Ok(())
/// }
/// ```
pub fn is_in_quick_access_exact(path: &str) -> WincentResult<bool> {
    let items = get_quick_access_items()?;
    Ok(items.iter().any(|item| item == path))
}

/// Checks if a path or keyword exists in the Windows Quick Access list.
///
/// **Note**: This function performs substring matching (fuzzy match). If you need
/// exact path matching, use `is_in_quick_access_exact()` instead.
///
/// # Arguments
///
/// * `keyword` - The path or partial path to search for (substring match)
///
/// # Returns
///
/// Returns `true` if any path in recent files or frequent folders contains the keyword.
///
/// # Example
///
/// ```rust
/// use wincent::{query::is_in_quick_access, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Fuzzy match - check for items containing "Documents"
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
                    path.contains(":\\") || path.starts_with("\\\\"),
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
                    path.contains(":\\") || path.starts_with("\\\\"),
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
                    path.contains(":\\") || path.starts_with("\\\\"),
                    "Path should be a valid Windows path format: {}",
                    path
                );
            }
        }

        Ok(())
    }

    #[test]
    fn test_exact_vs_fuzzy_matching() -> WincentResult<()> {
        let items = query_recent_with_ps_script(QuickAccess::All)?;

        if let Some(full_path) = items.first() {
            // exact match with full path should succeed
            assert!(
                is_in_quick_access_exact(full_path)?,
                "exact match should find full path"
            );
            assert!(
                is_in_quick_access(full_path)?,
                "fuzzy match should also find full path"
            );

            // exact match with partial path should fail
            if full_path.len() > 3 {
                let partial = &full_path[..full_path.len() - 1];
                assert!(
                    !is_in_quick_access_exact(partial)?,
                    "exact match should not find partial path"
                );
                // fuzzy match with partial path should succeed
                assert!(
                    is_in_quick_access(partial)?,
                    "fuzzy match should find partial path"
                );
            }
        }

        // non-existent path should return false for both
        let non_existent = "Z:\\Invalid\\Path\\Test.txt";
        assert!(!is_in_quick_access_exact(non_existent)?);
        assert!(!is_in_quick_access(non_existent)?);

        Ok(())
    }
}
