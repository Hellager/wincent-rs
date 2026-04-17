//! Windows Quick Access cleanup operations
//!
//! Provides unified interface for clearing Windows Quick Access items including:
//! - Recent files
//! - Frequent folders (both pinned and normal)
//! - Complete Quick Access history
//!
//! Implements multiple cleanup strategies with fallback mechanisms
//!
//! # Key Functionality
//! - Clear recent files using Windows Shell API
//! - Remove frequent folders through file system operations
//! - Clear pinned folders via PowerShell scripts
//! - Full Quick Access reset capabilities
//! - Atomic operations with proper cleanup sequencing

use crate::{error::WincentError, utils::get_windows_recent_folder, WincentResult};
use windows::Win32::System::Com::CoInitializeEx;
use windows::Win32::System::Com::CoUninitialize;
use windows::Win32::System::Com::COINIT_APARTMENTTHREADED;
use windows::Win32::UI::Shell::SHAddToRecentDocs;

/// Clears the Windows Recent Files list using the Windows Shell API.
pub(crate) fn empty_recent_files_with_api() -> WincentResult<()> {
    unsafe {
        let hr = CoInitializeEx(Some(std::ptr::null_mut()), COINIT_APARTMENTTHREADED);
        if hr.is_err() {
            return Err(WincentError::WindowsApi(hr.0));
        }

        // 0x0000_0003 equals SHARD_PATHW
        SHAddToRecentDocs(0x0000_0003, None);

        CoUninitialize();
    }

    Ok(())
}

/// Clears user folders from Quick Access by removing the Windows jump list file.
pub(crate) fn empty_user_folders_with_jumplist_file() -> WincentResult<()> {
    let recent_folder = get_windows_recent_folder()?;

    let jumplist_file = std::path::Path::new(&recent_folder)
        .join("AutomaticDestinations")
        .join("f01b4d95cf55d32a.automaticDestinations-ms");

    if jumplist_file.exists() {
        std::fs::remove_file(&jumplist_file).map_err(WincentError::Io)?;
    }

    Ok(())
}

/// Clear system default folders from Quick Access using PowerShell commands.
fn empty_system_default_folders_powershell() -> WincentResult<()> {
    use crate::script_executor::ScriptExecutor;
    use crate::script_strategy::PSScript;

    let start = std::time::Instant::now();
    let script_path = crate::script_storage::ScriptStorage::get_script_path(PSScript::EmptyPinnedFolders)?;
    let output = ScriptExecutor::execute_ps_script(PSScript::EmptyPinnedFolders, None)?;
    let duration = start.elapsed();
    let _ = ScriptExecutor::parse_output_to_strings(
        output,
        PSScript::EmptyPinnedFolders,
        script_path,
        None,
        duration,
    )?;

    Ok(())
}

/// Clear system default folders from Quick Access.
///
/// This function attempts to clear using native COM API first, falling back to PowerShell if COM fails.
/// The native approach is significantly faster (~50-100ms vs ~500-1000ms).
pub(crate) fn empty_system_default_folders() -> WincentResult<()> {
    // Try native COM first (fast path)
    crate::handle::empty_pinned_folders_native().or_else(|_| {
        // Fallback to PowerShell if COM fails (compatibility)
        empty_system_default_folders_powershell()
    })
}

/// Clears all items from the Windows Recent Files list.
///
/// # Returns
///
/// Returns `Ok(())` if all recent files were successfully cleared.
///
/// # Example
///
/// ```no_run
/// use wincent::{empty::empty_recent_files, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Clear all recent files
///     empty_recent_files()?;
///     println!("Recent files list has been cleared");
///     Ok(())
/// }
/// ```
pub fn empty_recent_files() -> WincentResult<()> {
    empty_recent_files_with_api()
}

/// Clears all items from the Windows Frequent Folders list, including both pinned and normal folders.
///
/// # Returns
///
/// Returns `Ok(())` if all frequent folders were successfully cleared.
///
/// # Example
///
/// ```no_run
/// use wincent::{empty::empty_frequent_folders, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Clear all frequent folders
///     empty_frequent_folders(false)?;
///     println!("Frequent folders list has been cleared");
///     Ok(())
/// }
/// ```
pub fn empty_frequent_folders(also_system_default: bool) -> WincentResult<()> {
    empty_user_folders_with_jumplist_file()?;
    if also_system_default {
        empty_system_default_folders()?;
    }
    Ok(())
}

/// Clears all items from Windows Quick Access, including both recent files and frequent folders.
///
/// # Returns
///
/// Returns `Ok(())` if all Quick Access items were successfully cleared.
///
/// # Example
///
/// ```no_run
/// use wincent::{empty::empty_quick_access, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Clear all Quick Access items
///     empty_quick_access(false)?;
///     println!("Quick Access has been completely cleared");
///     Ok(())
/// }
/// ```
pub fn empty_quick_access(also_system_default: bool) -> WincentResult<()> {
    empty_recent_files()?;
    empty_frequent_folders(also_system_default)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::handle::{add_file_to_recent_with_api, pin_frequent_folder};
    use crate::script_executor::ScriptExecutor;
    use crate::script_strategy::PSScript;
    use crate::test_utils::{cleanup_test_env, create_test_file, setup_test_env};
    use std::thread;
    use std::time::Duration;

    fn wait_for_files_empty(max_retries: u32) -> WincentResult<bool> {
        for _ in 0..max_retries {
            let start = std::time::Instant::now();
            let script_path = crate::script_storage::ScriptStorage::get_script_path(PSScript::QueryRecentFile)?;
            let output = ScriptExecutor::execute_ps_script(PSScript::QueryRecentFile, None)?;
            let duration = start.elapsed();
            let recent_files = ScriptExecutor::parse_output_to_strings(
                output,
                PSScript::QueryRecentFile,
                script_path,
                None,
                duration,
            )?;
            if recent_files.is_empty() {
                return Ok(true);
            }
            thread::sleep(Duration::from_millis(500));
        }
        Ok(false)
    }

    fn wait_for_folders_empty(max_retries: u32) -> WincentResult<bool> {
        for _ in 0..max_retries {
            let start = std::time::Instant::now();
            let script_path = crate::script_storage::ScriptStorage::get_script_path(PSScript::QueryFrequentFolder)?;
            let output = ScriptExecutor::execute_ps_script(PSScript::QueryFrequentFolder, None)?;
            let duration = start.elapsed();
            let folders = ScriptExecutor::parse_output_to_strings(
                output,
                PSScript::QueryFrequentFolder,
                script_path,
                None,
                duration,
            )?;
            if folders.is_empty() {
                return Ok(true);
            }
            thread::sleep(Duration::from_millis(500));
        }
        Ok(false)
    }

    #[test]
    #[ignore = "Modifies system state"]
    fn test_empty_recent_files() -> WincentResult<()> {
        let test_dir = setup_test_env()?;

        let test_file = create_test_file(&test_dir, "test.txt", "content")?;
        add_file_to_recent_with_api(test_file.to_str().unwrap())?;
        thread::sleep(Duration::from_secs(1));

        let start = std::time::Instant::now();
        let script_path = crate::script_storage::ScriptStorage::get_script_path(PSScript::QueryRecentFile)?;
        let output = ScriptExecutor::execute_ps_script(PSScript::QueryRecentFile, None)?;
        let duration = start.elapsed();
        let recent_files = ScriptExecutor::parse_output_to_strings(
            output,
            PSScript::QueryRecentFile,
            script_path,
            None,
            duration,
        )?;
        assert!(
            !recent_files.is_empty(),
            "File should have been added to recent list"
        );

        empty_recent_files_with_api()?;
        assert!(
            wait_for_files_empty(5)?,
            "Recent files list should be empty"
        );

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state"]
    fn test_empty_user_folders() -> WincentResult<()> {
        empty_user_folders_with_jumplist_file()?;
        thread::sleep(Duration::from_secs(1));

        let start = std::time::Instant::now();
        let script_path = crate::script_storage::ScriptStorage::get_script_path(PSScript::QueryFrequentFolder)?;
        let output = ScriptExecutor::execute_ps_script(PSScript::QueryFrequentFolder, None)?;
        let duration = start.elapsed();
        let folders = ScriptExecutor::parse_output_to_strings(
            output,
            PSScript::QueryFrequentFolder,
            script_path,
            None,
            duration,
        )?;
        assert!(
            folders.is_empty(),
            "No recent files should exist after jump list cleanup"
        );

        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state"]
    fn test_empty_pinned_folders() -> WincentResult<()> {
        let test_dir = setup_test_env()?;

        pin_frequent_folder(test_dir.to_str().unwrap())?;
        thread::sleep(Duration::from_secs(1));

        let start = std::time::Instant::now();
        let script_path = crate::script_storage::ScriptStorage::get_script_path(PSScript::QueryFrequentFolder)?;
        let output = ScriptExecutor::execute_ps_script(PSScript::QueryFrequentFolder, None)?;
        let duration = start.elapsed();
        let folders = ScriptExecutor::parse_output_to_strings(
            output,
            PSScript::QueryFrequentFolder,
            script_path,
            None,
            duration,
        )?;
        assert!(!folders.is_empty(), "Should have pinned folders");
        let test_path = test_dir.to_str().unwrap();
        assert!(
            folders.iter().any(|p| p == test_path),
            "Test folder should be in the list"
        );

        empty_system_default_folders()?;
        thread::sleep(Duration::from_secs(1));

        // Verify our test folder is gone (not necessarily all folders)
        let start = std::time::Instant::now();
        let script_path = crate::script_storage::ScriptStorage::get_script_path(PSScript::QueryFrequentFolder)?;
        let output = ScriptExecutor::execute_ps_script(PSScript::QueryFrequentFolder, None)?;
        let duration = start.elapsed();
        let folders_after = ScriptExecutor::parse_output_to_strings(
            output,
            PSScript::QueryFrequentFolder,
            script_path,
            None,
            duration,
        )?;
        assert!(
            !folders_after.iter().any(|p| p == test_path),
            "Test folder should have been removed"
        );

        cleanup_test_env(&test_dir)?;
        Ok(())
    }
}
