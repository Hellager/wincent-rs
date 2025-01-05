use crate::{
    scripts::{Script, execute_ps_script},
    error::{WincentError, WincentResult}
};

/// Removes recent files using a PowerShell script.
///
/// This function executes a PowerShell script to remove recent files from the specified path.
/// It checks the output status to determine if the operation was successful.
///
/// # Parameters
///
/// - `path`: A string slice that represents the path from which recent files should be removed.
///
/// # Returns
///
/// Returns a `WincentResult<()>`. If the operation is successful, it returns `Ok(())`.
/// If the script fails, it returns `WincentError::ScriptFailed` with the error message.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let path = "C:\\Users\\User\\Recent";
///     remove_recent_files_with_ps_script(path)?;
///     println!("Recent files removed successfully.");
///     Ok(())
/// }
/// ```
pub(crate) fn remove_recent_files_with_ps_script(path: &str) -> WincentResult<()> {
    let output = execute_ps_script(Script::RemoveRecentFile, Some(path))?;

    if output.status.success() {
        Ok(())
    } else {
        let error = String::from_utf8(output.stderr)?;
        Err(WincentError::ScriptFailed(error))
    }
}

/// Pins a folder to the Frequent Folders list using a PowerShell script.
///
/// This function executes a PowerShell script to pin the specified folder to the Frequent Folders list.
/// It checks the output status to determine if the operation was successful.
///
/// # Parameters
///
/// - `path`: A string slice that represents the path of the folder to be pinned.
///
/// # Returns
///
/// Returns a `WincentResult<()>`. If the operation is successful, it returns `Ok(())`.
/// If the script fails, it returns `WincentError::ScriptFailed` with the error message.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let path = "C:\\Users\\User\\Documents";
///     pin_frequent_folder_with_ps_script(path)?;
///     println!("Folder pinned to Frequent Folders successfully.");
///     Ok(())
/// }
/// ```
pub(crate) fn pin_frequent_folder_with_ps_script(path: &str) -> WincentResult<()> {
    let output = execute_ps_script(Script::PinToFrequentFolder, Some(path))?;

    if output.status.success() {
        Ok(())
    } else {
        let error = String::from_utf8(output.stderr)?;
        Err(WincentError::ScriptFailed(error))
    }
}

/// Unpins a folder from the Frequent Folders list using a PowerShell script.
///
/// This function executes a PowerShell script to unpin the specified folder from the Frequent Folders list.
/// It checks the output status to determine if the operation was successful.
///
/// # Parameters
///
/// - `path`: A string slice that represents the path of the folder to be unpinned.
///
/// # Returns
///
/// Returns a `WincentResult<()>`. If the operation is successful, it returns `Ok(())`.
/// If the script fails, it returns `WincentError::ScriptFailed` with the error message.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let path = "C:\\Users\\User\\Documents";
///     unpin_frequent_folder_with_ps_script(path)?;
///     println!("Folder unpinned from Frequent Folders successfully.");
///     Ok(())
/// }
/// ```
pub(crate) fn unpin_frequent_folder_with_ps_script(path: &str) -> WincentResult<()> {
    let output = execute_ps_script(Script::UnpinFromFrequentFolder, Some(path))?;

    if output.status.success() {
        Ok(())
    } else {
        let error = String::from_utf8(output.stderr)?;
        Err(WincentError::ScriptFailed(error))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::error::WincentResult;
    use std::path::PathBuf;

    fn get_test_path() -> PathBuf {
        let mut path = std::env::current_dir().unwrap();
        path.push("tests");
        path.push("test_folder");
        path
    }

    #[test]
    fn test_pin_frequent_folder() -> WincentResult<()> {
        let test_path = get_test_path();
        std::fs::create_dir_all(&test_path)?;

        pin_frequent_folder_with_ps_script(test_path.to_str().unwrap())?;
        
        std::fs::remove_dir_all(&test_path)?;
        Ok(())
    }

    #[test]
    fn test_unpin_frequent_folder() -> WincentResult<()> {
        let test_path = get_test_path();
        std::fs::create_dir_all(&test_path)?;

        pin_frequent_folder_with_ps_script(test_path.to_str().unwrap())?;
        
        unpin_frequent_folder_with_ps_script(test_path.to_str().unwrap())?;
        
        std::fs::remove_dir_all(&test_path)?;
        Ok(())
    }

    #[ignore]
    #[test]
    fn test_remove_recent_files() -> WincentResult<()> {
        let test_path = get_test_path();
        std::fs::create_dir_all(&test_path)?;
        
        let test_file = test_path.join("test.txt");
        std::fs::write(&test_file, "test content")?;

        remove_recent_files_with_ps_script(test_file.to_str().unwrap())?;
        
        std::fs::remove_dir_all(&test_path)?;
        Ok(())
    }
}
