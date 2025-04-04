#![allow(dead_code)]

use crate::{
    error::WincentError,
    scripts::{execute_ps_script, Script},
    WincentResult,
};
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use windows::Win32::Foundation::{BOOL, HANDLE};
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::UI::Shell::IsUserAnAdmin;
use windows::Win32::UI::Shell::{FOLDERID_Recent, SHGetKnownFolderPath, KNOWN_FOLDER_FLAG};

/// Checks if the current user has administrative privileges.
pub(crate) fn is_admin() -> bool {
    unsafe { IsUserAnAdmin() == BOOL(1) }
}

/// Refreshes the Windows Explorer window using a PowerShell script.
pub(crate) fn refresh_explorer_window() -> WincentResult<()> {
    let output = execute_ps_script(Script::RefreshExplorer, None)?;

    if output.status.success() {
        Ok(())
    } else {
        let error = String::from_utf8(output.stderr)?;
        Err(WincentError::ScriptFailed(error))
    }
}

/// Get Windows Recent Folder path
pub(crate) fn get_windows_recent_folder() -> WincentResult<String> {
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

    Ok(recent_folder)
}

#[cfg(test)]
mod utils_test {
    use super::*;

    #[test]
    fn test_check_admin() {
        let is_admin = is_admin();
        assert!(is_admin || !is_admin, "Should return a boolean value");
    }

    #[test]
    fn test_refresh_explorer() -> WincentResult<()> {
        refresh_explorer_window()
    }

    #[test]
    fn test_get_windows_recent_folder() -> WincentResult<()> {
        let recent_folder = get_windows_recent_folder()?;
        assert!(
            !recent_folder.is_empty(),
            "Recent folder path should not be empty"
        );
        assert!(
            std::path::Path::new(&recent_folder).exists(),
            "Recent folder should exist"
        );
        Ok(())
    }
}
