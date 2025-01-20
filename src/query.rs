use crate::{
    error::WincentError,
    feasible::{check_query_feasible, check_script_feasible},
    scripts::{execute_ps_script, Script},
    QuickAccess, WincentResult,
};

/// Queries recent items from Quick Access using a PowerShell script.
pub(crate) fn query_recent_with_ps_script(qa_type: QuickAccess) -> WincentResult<Vec<String>> {
    let output = match qa_type {
        QuickAccess::All => execute_ps_script(Script::QueryQuickAccess, None)?,
        QuickAccess::RecentFiles => execute_ps_script(Script::QuertRecentFile, None)?,
        QuickAccess::FrequentFolders => execute_ps_script(Script::QueryFrequentFolder, None)?,
    };

    if output.status.success() {
        let stdout_str = String::from_utf8(output.stdout).map_err(WincentError::Utf8)?;

        let data: Vec<String> = stdout_str
            .lines()
            .map(str::trim)
            .filter(|line| !line.is_empty())
            .map(String::from)
            .collect();

        Ok(data)
    } else {
        let error = String::from_utf8(output.stderr)?;
        Err(WincentError::ScriptFailed(error))
    }
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
    if !check_script_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "PowerShell script execution is not feasible".to_string(),
        ));
    }

    if !check_query_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "Quick Access query operation is not feasible".to_string(),
        ));
    }

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
    if !check_script_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "PowerShell script execution is not feasible".to_string(),
        ));
    }

    if !check_query_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "Quick Access query operation is not feasible".to_string(),
        ));
    }

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
    if !check_script_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "PowerShell script execution is not feasible".to_string(),
        ));
    }

    if !check_query_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "Quick Access query operation is not feasible".to_string(),
        ));
    }

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
