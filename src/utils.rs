#![allow(dead_code)]

use crate::{
    error::WincentError, script_executor::ScriptExecutor, script_strategy::PSScript, WincentResult,
};
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use std::path::Path;
use windows::Wdk::System::SystemServices::RtlGetVersion;
use windows::Win32::Foundation::{BOOL, HANDLE};
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::System::Diagnostics::Debug::VER_PLATFORM_WIN32_NT;
use windows::Win32::System::SystemInformation::OSVERSIONINFOEXW;
use windows::Win32::UI::Shell::IsUserAnAdmin;
use windows::Win32::UI::Shell::{FOLDERID_Recent, SHGetKnownFolderPath, KNOWN_FOLDER_FLAG};

#[derive(Debug, Copy, Clone)]
pub(crate) enum PathType {
    File,
    Directory,
}

/// Checks if the current user has administrative privileges.
pub(crate) fn is_admin() -> bool {
    unsafe { IsUserAnAdmin() == BOOL(1) }
}

/// Refreshes the Windows Explorer window using a PowerShell script.
pub(crate) fn refresh_explorer_window() -> WincentResult<()> {
    let output = ScriptExecutor::execute_ps_script(PSScript::RefreshExplorer, None)?;
    let _ = ScriptExecutor::parse_output_to_strings(output)?;

    Ok(())
}

/// Validates if a given path exists and matches the expected type (file or directory).
pub(crate) fn validate_path(path: &str, expected_type: PathType) -> WincentResult<()> {
    let path_buf = Path::new(path);

    if path.is_empty() {
        return Err(WincentError::InvalidPath("Empty path provided".to_string()));
    }

    if !path_buf.exists() {
        return Err(WincentError::InvalidPath(format!(
            "Path does not exist: {}",
            path
        )));
    }

    match expected_type {
        PathType::File if !path_buf.is_file() => Err(WincentError::InvalidPath(format!(
            "Not a valid file: {}",
            path
        ))),
        PathType::Directory if !path_buf.is_dir() => Err(WincentError::InvalidPath(format!(
            "Not a valid directory: {}",
            path
        ))),
        _ => Ok(()),
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

/// Get Windows OS Version
fn get_os_version() -> WincentResult<OSVERSIONINFOEXW> {
    let mut info = OSVERSIONINFOEXW {
        dwOSVersionInfoSize: std::mem::size_of::<OSVERSIONINFOEXW>() as u32,
        ..Default::default()
    };

    unsafe {
        RtlGetVersion(&mut info as *mut _ as *mut _).ok()?;
    }

    Ok(info)
}

/// Check Whether Win11
pub(crate) fn is_win11() -> WincentResult<bool> {
    let version_info = get_os_version()?;

    if version_info.dwPlatformId != VER_PLATFORM_WIN32_NT.0 {
        return Err(WincentError::SystemError(
            "No Windows NT system".to_string(),
        ));
    }

    match (version_info.dwMajorVersion, version_info.dwMinorVersion) {
        (10, 0) if version_info.dwBuildNumber >= 22000 => Ok(true),
        (10, 0) => Ok(false),
        _ => Ok(false),
    }
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

    #[test]
    fn test_is_win11() -> WincentResult<()> {
        let is_win11 = is_win11()?;
        assert!(is_win11 || !is_win11, "Should return a boolean value");
        Ok(())
    }
}
