use crate::{
    error::WincentError, feasible::check_script_feasible,
    handle::unpin_frequent_folder_with_ps_script, query::query_recent_with_ps_script, QuickAccess,
    WincentResult,
};
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use windows::Win32::Foundation::HANDLE;
use windows::Win32::System::Com::CoInitializeEx;
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::System::Com::CoUninitialize;
use windows::Win32::System::Com::COINIT_APARTMENTTHREADED;
use windows::Win32::UI::Shell::SHAddToRecentDocs;
use windows::Win32::UI::Shell::{FOLDERID_Recent, SHGetKnownFolderPath, KNOWN_FOLDER_FLAG};

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

/// Clears normal folders from Quick Access by removing the Windows jump list file.
pub(crate) fn empty_normal_folders_with_jumplist_file() -> WincentResult<()> {
    let result = unsafe {
        SHGetKnownFolderPath(
            &FOLDERID_Recent,
            KNOWN_FOLDER_FLAG(0x00),
            HANDLE(std::ptr::null_mut()),
        )
    }?;

    let recent_folder = unsafe {
        let wide_str = OsString::from_wide(result.as_wide());
        CoTaskMemFree(Some(result.as_ptr() as _));
        wide_str
            .into_string()
            .map_err(|_| WincentError::SystemError("Invalid UTF-16".to_string()))?
    };

    let jumplist_file = std::path::Path::new(&recent_folder)
        .join("AutomaticDestinations")
        .join("f01b4d95cf55d32a.automaticDestinations-ms");

    if jumplist_file.exists() {
        std::fs::remove_file(&jumplist_file).map_err(WincentError::Io)?;
    }

    Ok(())
}

/// Removes all pinned folders from Quick Access using PowerShell commands.
pub(crate) fn empty_pinned_folders_with_script() -> WincentResult<()> {
    let folders = query_recent_with_ps_script(QuickAccess::FrequentFolders)?;

    for folder in folders {
        unpin_frequent_folder_with_ps_script(&folder)?;
    }

    Ok(())
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
    if !check_script_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "PowerShell script execution is not feasible".to_string(),
        ));
    }

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
///     empty_frequent_folders()?;
///     println!("Frequent folders list has been cleared");
///     Ok(())
/// }
/// ```
pub fn empty_frequent_folders() -> WincentResult<()> {
    if !check_script_feasible()? {
        return Err(WincentError::UnsupportedOperation(
            "PowerShell script execution is not feasible".to_string(),
        ));
    }

    empty_normal_folders_with_jumplist_file()?;
    empty_pinned_folders_with_script()?;
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
///     empty_quick_access()?;
///     println!("Quick Access has been completely cleared");
///     Ok(())
/// }
/// ```
pub fn empty_quick_access() -> WincentResult<()> {
    empty_recent_files()?;
    empty_frequent_folders()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::handle::{add_file_to_recent_with_api, pin_frequent_folder_with_ps_script};
    use crate::test_utils::{cleanup_test_env, create_test_file, setup_test_env};
    use std::thread;
    use std::time::Duration;

    fn wait_for_files_empty(max_retries: u32) -> WincentResult<bool> {
        for _ in 0..max_retries {
            let recent_files = query_recent_with_ps_script(QuickAccess::RecentFiles)?;
            if recent_files.is_empty() {
                return Ok(true);
            }
            thread::sleep(Duration::from_millis(500));
        }
        Ok(false)
    }

    fn wait_for_folders_empty(max_retries: u32) -> WincentResult<bool> {
        for _ in 0..max_retries {
            let folders = query_recent_with_ps_script(QuickAccess::FrequentFolders)?;
            if folders.is_empty() {
                return Ok(true);
            }
            thread::sleep(Duration::from_millis(500));
        }
        Ok(false)
    }

    #[test]
    #[ignore]
    fn test_empty_recent_files() -> WincentResult<()> {
        let test_dir = setup_test_env()?;

        let test_file = create_test_file(&test_dir, "test.txt", "content")?;
        add_file_to_recent_with_api(test_file.to_str().unwrap())?;
        thread::sleep(Duration::from_secs(1));

        let recent_files = query_recent_with_ps_script(QuickAccess::RecentFiles)?;
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
    #[ignore]
    fn test_empty_normal_folders() -> WincentResult<()> {
        empty_normal_folders_with_jumplist_file()?;
        thread::sleep(Duration::from_secs(1));

        let recent_files = query_recent_with_ps_script(QuickAccess::RecentFiles)?;
        assert!(
            recent_files.is_empty(),
            "No recent files should exist after jump list cleanup"
        );

        Ok(())
    }

    #[test]
    #[ignore]
    fn test_empty_pinned_folders() -> WincentResult<()> {
        let test_dir = setup_test_env()?;

        pin_frequent_folder_with_ps_script(test_dir.to_str().unwrap())?;
        thread::sleep(Duration::from_secs(1));

        let folders = query_recent_with_ps_script(QuickAccess::FrequentFolders)?;
        assert!(!folders.is_empty(), "Should have pinned folders");

        empty_pinned_folders_with_script()?;
        assert!(
            wait_for_folders_empty(5)?,
            "Pinned folders list should be empty"
        );

        cleanup_test_env(&test_dir)?;
        Ok(())
    }
}
