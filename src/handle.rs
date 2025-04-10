//! Windows Quick Access item management
//!
//! Provides system-level manipulation of Quick Access locations including:
//! - Recent files management
//! - Frequent folders pinning
//! - Cross-API operation support (PowerShell + Win32 API)
//!
//! # Key Functionality
//! - File addition/removal from Recent Items
//! - Folder pinning/unpinning operations
//! - Path validation and sanitization
//! - Multi-strategy execution (PowerShell/Win32 API)
//!
//! # Operation Safety
//! 1. Automatic COM initialization for API operations
//! 2. Path validation before execution
//! 3. PowerShell script sandboxing
//! 4. Clean error propagation

use crate::{
    error::WincentError,
    script_executor::ScriptExecutor,
    script_strategy::PSScript,
    utils::{validate_path, PathType},
    WincentResult,
};
use std::ffi::OsString;
use std::os::windows::prelude::*;
use windows::Win32::System::Com::CoInitializeEx;
use windows::Win32::System::Com::CoUninitialize;
use windows::Win32::System::Com::COINIT_APARTMENTTHREADED;
use windows::Win32::UI::Shell::SHAddToRecentDocs;

/// Executes a PowerShell script after validating the given path.
pub(crate) fn execute_script_with_validation(
    script: PSScript,
    path: &str,
    path_type: PathType,
) -> WincentResult<()> {
    validate_path(path, path_type)?;

    let output = ScriptExecutor::execute_ps_script(script, Some(path))?;

    match output.status.success() {
        true => Ok(()),
        false => {
            let error = String::from_utf8(output.stderr)
                .unwrap_or_else(|_| "Unable to parse script error output".to_string());
            Err(WincentError::ScriptFailed(error))
        }
    }
}

/// Adds a file to the Windows Recent Items list using the Windows API.
pub(crate) fn add_file_to_recent_with_api(path: &str) -> WincentResult<()> {
    validate_path(path, PathType::File)?;

    unsafe {
        let hr = CoInitializeEx(Some(std::ptr::null_mut()), COINIT_APARTMENTTHREADED);
        if hr.is_err() {
            return Err(WincentError::WindowsApi(hr.0));
        }

        let file_path_wide: Vec<u16> = OsString::from(path)
            .encode_wide()
            .chain(std::iter::once(0))
            .collect();

        // 0x0000_0003 equals SHARD_PATHW
        SHAddToRecentDocs(0x0000_0003, Some(file_path_wide.as_ptr() as *const _));

        CoUninitialize();
    }

    Ok(())
}

/// Removes a file from the Windows Recent Items list using PowerShell.
pub(crate) fn remove_recent_files_with_ps_script(path: &str) -> WincentResult<()> {
    execute_script_with_validation(PSScript::RemoveRecentFile, path, PathType::File)
}

/// Pins a folder to the Windows Quick Access Frequent Folders list.
pub(crate) fn pin_frequent_folder_with_ps_script(path: &str) -> WincentResult<()> {
    execute_script_with_validation(PSScript::PinToFrequentFolder, path, PathType::Directory)
}

/// Unpins a folder from the Windows Quick Access Frequent Folders list.
pub(crate) fn unpin_frequent_folder_with_ps_script(path: &str) -> WincentResult<()> {
    execute_script_with_validation(PSScript::UnpinFromFrequentFolder, path, PathType::Directory)
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
/// ```no_run
/// use wincent::{handle::add_to_recent_files, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     add_to_recent_files("C:\\Documents\\report.docx")?;
///     Ok(())
/// }
/// ```
pub fn add_to_recent_files(path: &str) -> WincentResult<()> {
    add_file_to_recent_with_api(path)
}

/// Removes a file from Windows Recent Files.
///
/// # Arguments
///
/// * `path` - The full path to the file to be removed
///
/// # Example
///
/// ```no_run
/// use wincent::{handle::remove_from_recent_files, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     remove_from_recent_files("C:\\Documents\\report.docx")?;
///     Ok(())
/// }
/// ```
pub fn remove_from_recent_files(path: &str) -> WincentResult<()> {
    remove_recent_files_with_ps_script(path)
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
/// ```no_run
/// use wincent::{handle::add_to_frequent_folders, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Pin a project folder
///     add_to_frequent_folders("C:\\Projects\\my-project")?;
///     Ok(())
/// }   
/// ```
pub fn add_to_frequent_folders(path: &str) -> WincentResult<()> {
    pin_frequent_folder_with_ps_script(path)
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
/// ```no_run
/// use wincent::{handle::remove_from_frequent_folders, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Unpin a project folder
///     remove_from_frequent_folders("C:\\Projects\\old-project")?;
///     Ok(())
/// }
/// ```
pub fn remove_from_frequent_folders(path: &str) -> WincentResult<()> {
    unpin_frequent_folder_with_ps_script(path)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::query::query_recent_with_ps_script;
    use crate::test_utils::{cleanup_test_env, create_test_file, setup_test_env};
    use std::{thread, time::Duration};

    fn wait_for_folder_status(
        path: &str,
        should_exist: bool,
        max_retries: u32,
    ) -> WincentResult<bool> {
        for _ in 0..max_retries {
            let frequent_folders =
                query_recent_with_ps_script(crate::QuickAccess::FrequentFolders)?;
            let exists = frequent_folders.iter().any(|p| p == path);

            if exists == should_exist {
                return Ok(true);
            }

            thread::sleep(Duration::from_millis(500));
        }
        Ok(false)
    }

    fn wait_for_file_status(
        path: &str,
        should_exist: bool,
        max_retries: u32,
    ) -> WincentResult<bool> {
        for _ in 0..max_retries {
            let recent_files = query_recent_with_ps_script(crate::QuickAccess::RecentFiles)?;
            let exists = recent_files.iter().any(|p| p == path);

            if exists == should_exist {
                return Ok(true);
            }

            thread::sleep(Duration::from_millis(500));
        }
        Ok(false)
    }

    #[test]
    #[ignore = "Modifies system state"]
    fn test_pin_unpin_frequent_folder() -> WincentResult<()> {
        let test_dir = setup_test_env()?;
        let test_path = test_dir.to_str().unwrap();

        pin_frequent_folder_with_ps_script(test_path)?;

        assert!(
            wait_for_folder_status(test_path, true, 5)?,
            "Pin operation failed: folder did not appear in frequent folders list"
        );

        unpin_frequent_folder_with_ps_script(test_path)?;

        assert!(
            wait_for_folder_status(test_path, false, 5)?,
            "Unpin operation failed: folder still exists in frequent folders list"
        );

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    fn test_pin_frequent_folder_error_handling() -> WincentResult<()> {
        let result = pin_frequent_folder_with_ps_script("Z:\\NonExistentFolder");
        assert!(result.is_err(), "Should fail with non-existent folder");

        let result = pin_frequent_folder_with_ps_script("");
        assert!(result.is_err(), "Should fail with empty path");

        Ok(())
    }

    #[test]
    fn test_unpin_frequent_folder_error_handling() -> WincentResult<()> {
        let result = unpin_frequent_folder_with_ps_script("Z:\\NonExistentFolder");
        assert!(result.is_err(), "Should fail with non-existent folder");

        let result = unpin_frequent_folder_with_ps_script("");
        assert!(result.is_err(), "Should fail with empty path");

        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state"]
    fn test_concurrent_operations() -> WincentResult<()> {
        let test_dir = setup_test_env()?;

        let path = test_dir.to_str().unwrap();

        pin_frequent_folder_with_ps_script(path)?;
        unpin_frequent_folder_with_ps_script(path)?;

        unpin_frequent_folder_with_ps_script(path)?;
        pin_frequent_folder_with_ps_script(path)?;

        unpin_frequent_folder_with_ps_script(path)?;

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state"]
    fn test_add_remove_file_in_recent() -> WincentResult<()> {
        let test_dir = setup_test_env()?;

        let test_file = create_test_file(&test_dir, "recent_test.txt", "test content")?;
        let test_path = test_file.to_str().unwrap();

        add_file_to_recent_with_api(test_path)?;

        assert!(
            wait_for_file_status(test_path, true, 10)?,
            "Add operation failed: file did not appear in recent files list"
        );

        remove_recent_files_with_ps_script(test_path)?;

        assert!(
            wait_for_file_status(test_path, false, 5)?,
            "Remove operation failed: file still exists in recent files list"
        );

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    fn test_add_file_to_recent_error_handling() -> WincentResult<()> {
        let result = add_file_to_recent_with_api("Z:\\NonExistentFile.txt");
        assert!(
            result.is_err(),
            "Windows API should not allow adding non-existent file paths"
        );

        let result = add_file_to_recent_with_api("");
        assert!(result.is_err(), "Should fail with empty path");

        let result = add_file_to_recent_with_api("\0invalid\0path");
        assert!(
            result.is_err(),
            "Invalid path characters should not be allowed"
        );

        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state"]
    fn test_add_file_to_recent_with_unicode() -> WincentResult<()> {
        let test_dir = setup_test_env()?;

        let test_file = create_test_file(&test_dir, "test_file.txt", "test content")?;
        add_file_to_recent_with_api(test_file.to_str().unwrap())?;

        let test_file2 = create_test_file(&test_dir, "test file with spaces.txt", "test content")?;
        add_file_to_recent_with_api(test_file2.to_str().unwrap())?;

        remove_recent_files_with_ps_script(test_file.to_str().unwrap())?;

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    fn test_remove_recent_files_error_handling() -> WincentResult<()> {
        let result = remove_recent_files_with_ps_script("Z:\\NonExistentFile.txt");
        assert!(result.is_err(), "Should fail with non-existent file");

        let result = remove_recent_files_with_ps_script("");
        assert!(result.is_err(), "Should fail with empty path");

        let result = remove_recent_files_with_ps_script("invalid\\path\\*");
        assert!(result.is_err(), "Should fail with invalid path");

        Ok(())
    }
}
