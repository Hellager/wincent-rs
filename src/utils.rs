use crate::{
    error::WincentError, script_executor::ScriptExecutor, script_strategy::PSScript, WincentResult,
};
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use std::path::Path;
use windows::Wdk::System::SystemServices::RtlGetVersion;
use windows::Win32::Foundation::{BOOL, HANDLE};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED,
    CoTaskMemFree,
};
use windows::Win32::System::Diagnostics::Debug::VER_PLATFORM_WIN32_NT;
use windows::Win32::System::SystemInformation::OSVERSIONINFOEXW;
use windows::Win32::UI::Shell::IsUserAnAdmin;
use windows::Win32::UI::Shell::{
    FOLDERID_Recent, SHGetKnownFolderPath, KNOWN_FOLDER_FLAG, IShellWindows, IWebBrowser2, ShellWindows,
};
use windows::core::Interface;

#[derive(Debug, Copy, Clone)]
pub(crate) enum PathType {
    File,
    Directory,
}

/// Checks if the current user has administrative privileges.
#[allow(dead_code)]
pub(crate) fn is_admin() -> bool {
    unsafe { IsUserAnAdmin() == BOOL(1) }
}

/// Refreshes the Windows Explorer windows.
///
/// This function attempts to refresh all open Explorer windows using native COM API first,
/// falling back to PowerShell if COM fails. The native approach is significantly faster
/// (~1-10ms vs ~100-500ms).
pub(crate) fn refresh_explorer_window() -> WincentResult<()> {
    // Try native COM first (fast path)
    refresh_explorer_native().or_else(|_| {
        // Fallback to PowerShell if COM fails (compatibility)
        refresh_explorer_powershell()
    })
}

/// Refreshes Explorer windows using native COM API (fast path).
///
/// Directly calls Shell.Application COM interface to refresh all Explorer windows.
/// This is much faster than spawning a PowerShell process.
///
/// Refreshes all Explorer windows including Quick Access, This PC, and file system views.
fn refresh_explorer_native() -> WincentResult<()> {
    unsafe {
        // Initialize COM (handle already-initialized case)
        let com_initialized = match CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok() {
            Ok(_) => true,
            Err(e) => {
                // RPC_E_CHANGED_MODE (0x80010106) means COM is already initialized in this thread
                // with a different apartment model - we can still proceed
                const RPC_E_CHANGED_MODE: i32 = -2147417850_i32; // 0x80010106
                if e.code().0 == RPC_E_CHANGED_MODE {
                    false // Don't uninitialize later
                } else {
                    return Err(WincentError::SystemError(format!("Failed to initialize COM: {}", e)));
                }
            }
        };

        let result = (|| -> WincentResult<()> {
            // Create Shell.Application object
            let shell_windows: IShellWindows = CoCreateInstance(&ShellWindows, None, CLSCTX_LOCAL_SERVER)
                .map_err(|e| WincentError::SystemError(format!("Failed to create ShellWindows: {}", e)))?;

            // Get window count
            let count = shell_windows.Count()
                .map_err(|e| WincentError::SystemError(format!("Failed to get window count: {}", e)))?;

            let mut refreshed_count = 0;
            let mut total_explorer_windows = 0;

            // Refresh each Explorer window
            for i in 0..count {
                if let Ok(dispatch) = shell_windows.Item(&i.into()) {
                    // Try to get IWebBrowser2 interface
                    if let Ok(web_browser) = dispatch.cast::<IWebBrowser2>() {
                        // Check if this is an Explorer window
                        if let Ok(location) = web_browser.LocationURL() {
                            let url = location.to_string();

                            // Refresh all Explorer windows:
                            // - file:/// for file system views
                            // - shell::: for Quick Access, This PC, etc.
                            // Skip only browser windows (http://, https://)
                            if url.starts_with("file:///") || url.starts_with("shell:::") {
                                total_explorer_windows += 1;
                                if web_browser.Refresh().is_ok() {
                                    refreshed_count += 1;
                                }
                            }
                        }
                    }
                }
            }

            // If we found Explorer windows but couldn't refresh any, return error
            // to trigger PowerShell fallback
            if total_explorer_windows > 0 && refreshed_count == 0 {
                return Err(WincentError::SystemError(
                    format!("Found {} Explorer windows but failed to refresh any", total_explorer_windows)
                ));
            }

            Ok(())
        })();

        // Only uninitialize COM if we initialized it
        if com_initialized {
            CoUninitialize();
        }

        result
    }
}

/// Refreshes Explorer windows using PowerShell script (fallback).
///
/// This is the original implementation, kept as a fallback for compatibility.
fn refresh_explorer_powershell() -> WincentResult<()> {
    let start = std::time::Instant::now();
    let script_path = crate::script_storage::ScriptStorage::get_script_path(PSScript::RefreshExplorer)?;
    let output = ScriptExecutor::execute_ps_script(PSScript::RefreshExplorer, None)?;
    let duration = start.elapsed();
    let _ = ScriptExecutor::parse_output_to_strings(
        output,
        PSScript::RefreshExplorer,
        script_path,
        None,
        duration,
    )?;

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
#[allow(dead_code)]
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
#[allow(dead_code)]
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
        // At least verify the function doesn't panic
        println!("Running as admin: {}", is_admin);
    }

    #[test]
    #[ignore = "Requires desktop session and may need elevated privileges"]
    fn test_refresh_explorer() -> WincentResult<()> {
        refresh_explorer_window()
    }

    #[test]
    #[ignore = "Performance benchmark - requires desktop session"]
    fn test_refresh_explorer_performance() -> WincentResult<()> {
        use std::time::Instant;

        println!("\n=== Refresh Explorer Performance Comparison ===\n");

        // Test native COM version
        let start = Instant::now();
        refresh_explorer_native()?;
        let native_duration = start.elapsed();
        println!("Native COM version: {:?}", native_duration);

        // Test PowerShell version
        let start = Instant::now();
        refresh_explorer_powershell()?;
        let powershell_duration = start.elapsed();
        println!("PowerShell version: {:?}", powershell_duration);

        // Calculate speedup
        let speedup = powershell_duration.as_secs_f64() / native_duration.as_secs_f64();
        println!("\nSpeedup: {:.2}x faster", speedup);
        println!("Time saved: {:?}", powershell_duration - native_duration);

        Ok(())
    }

    #[test]
    #[ignore = "Requires desktop session and may need elevated privileges"]
    fn test_refresh_explorer_native_only() -> WincentResult<()> {
        // Test that native COM version works independently
        refresh_explorer_native()
    }

    #[test]
    #[ignore = "Requires desktop session and may need elevated privileges"]
    fn test_refresh_explorer_powershell_only() -> WincentResult<()> {
        // Test that PowerShell fallback still works
        refresh_explorer_powershell()
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
        println!("Running on Windows 11: {}", is_win11);
        Ok(())
    }

    #[test]
    fn test_validate_path_file() -> WincentResult<()> {
        // Create temporary file
        let temp_file = tempfile::NamedTempFile::new()
            .map_err(|e| WincentError::SystemError(e.to_string()))?;
        let path = temp_file.path().to_str().unwrap();

        // Should succeed as file
        validate_path(path, PathType::File)?;

        // Should fail as directory
        assert!(
            validate_path(path, PathType::Directory).is_err(),
            "File should not validate as directory"
        );

        Ok(())
    }

    #[test]
    fn test_validate_path_directory() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir()
            .map_err(|e| WincentError::SystemError(e.to_string()))?;
        let path = temp_dir.path().to_str().unwrap();

        // Should succeed as directory
        validate_path(path, PathType::Directory)?;

        // Should fail as file
        assert!(
            validate_path(path, PathType::File).is_err(),
            "Directory should not validate as file"
        );

        Ok(())
    }

    #[test]
    fn test_validate_path_empty() {
        let result = validate_path("", PathType::File);
        assert!(result.is_err(), "Empty path should fail validation");

        if let Err(WincentError::InvalidPath(msg)) = result {
            assert!(msg.contains("Empty"), "Error should mention empty path");
        } else {
            panic!("Expected InvalidPath error");
        }
    }

    #[test]
    fn test_validate_path_nonexistent() {
        let result = validate_path("Z:\\NonExistent\\Path.txt", PathType::File);
        assert!(result.is_err(), "Non-existent path should fail validation");

        if let Err(WincentError::InvalidPath(msg)) = result {
            assert!(msg.contains("does not exist"), "Error should mention path doesn't exist");
        } else {
            panic!("Expected InvalidPath error");
        }
    }

    #[test]
    fn test_validate_path_type_mismatch() -> WincentResult<()> {
        // Use Windows Recent folder (a directory that definitely exists)
        let recent_folder = get_windows_recent_folder()?;

        // Should succeed as directory
        validate_path(&recent_folder, PathType::Directory)?;

        // Should fail as file
        let result = validate_path(&recent_folder, PathType::File);
        assert!(result.is_err(), "Directory should not validate as file");

        if let Err(WincentError::InvalidPath(msg)) = result {
            assert!(msg.contains("Not a valid file"), "Error should mention type mismatch");
        }

        Ok(())
    }
}

