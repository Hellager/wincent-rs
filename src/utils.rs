use crate::{
    error::WincentError, script_executor::ScriptExecutor, script_strategy::PSScript, WincentResult,
};
use std::ffi::{OsStr, OsString};
use std::os::windows::ffi::OsStrExt;
use std::os::windows::ffi::OsStringExt;
use std::path::Path;
use std::time::Duration;

/// Lightweight path normalization without I/O operations.
///
/// Converts to lowercase, replaces forward slashes with backslashes, and strips
/// a trailing backslash (except for root paths like `"C:\"`).
pub(crate) fn normalize_path_lightweight(path: &str) -> String {
    let mut result = path.to_lowercase().replace('/', "\\");

    // Remove trailing backslash unless it's a root path (e.g., "C:\")
    if result.len() > 3 && result.ends_with('\\') {
        result.pop();
    }

    result
}

/// Normalizes a Windows path for comparison, using canonicalization when available.
///
/// Falls back to [`normalize_path_lightweight`] if `Path::canonicalize` fails
/// (e.g. path doesn't exist on disk).
pub(crate) fn normalize_path_for_comparison(path: &str) -> String {
    let path_obj = Path::new(path);

    let normalized = if let Ok(canonical_path) = path_obj.canonicalize() {
        canonical_path.to_string_lossy().to_string()
    } else {
        path.to_string()
    };

    normalize_path_lightweight(&normalized)
}

/// Checks whether two Windows paths refer to the same location.
///
/// Uses a two-stage strategy:
/// 1. Fast lightweight string comparison (no I/O).
/// 2. Canonicalization-based comparison (resolves symlinks / relative paths).
pub(crate) fn paths_equal(path1: &str, path2: &str) -> bool {
    let light1 = normalize_path_lightweight(path1);
    let light2 = normalize_path_lightweight(path2);

    if light1 == light2 {
        return true;
    }

    normalize_path_for_comparison(path1) == normalize_path_for_comparison(path2)
}

pub(crate) fn os_str_to_wide_null(value: &OsStr) -> Vec<u16> {
    value.encode_wide().chain(std::iter::once(0)).collect()
}
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::UI::Shell::{FOLDERID_Recent, SHGetKnownFolderPath, KNOWN_FOLDER_FLAG};

#[derive(Debug, Copy, Clone)]
pub(crate) enum PathType {
    File,
    Directory,
}

/// Refreshes the Windows Explorer windows with a caller-supplied timeout for the PowerShell fallback.
///
/// Attempts native COM first (~1-10ms). If that fails, falls back to PowerShell
/// with `timeout` capping how long the caller waits.
pub(crate) fn refresh_explorer_window_with_timeout(timeout: Duration) -> WincentResult<()> {
    refresh_explorer_window_with_timeout_using(
        timeout,
        refresh_explorer_native_with_timeout,
        refresh_explorer_powershell_with_timeout,
    )
}

fn refresh_explorer_window_with_timeout_using<N, P>(
    timeout: Duration,
    native: N,
    powershell: P,
) -> WincentResult<()>
where
    N: FnOnce(Duration) -> WincentResult<()>,
    P: FnOnce(Duration) -> WincentResult<()>,
{
    native(timeout).or_else(|_| powershell(timeout))
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
    crate::explorer_window::refresh_explorer_shell_views()
}

fn refresh_explorer_native_with_timeout(timeout: Duration) -> WincentResult<()> {
    crate::explorer_window::refresh_explorer_shell_views_with_timeout(timeout)
}

/// Refreshes Explorer windows using PowerShell script (fallback).
///
/// This is the original implementation, kept as a fallback for compatibility.
fn refresh_explorer_powershell() -> WincentResult<()> {
    let start = std::time::Instant::now();
    let script_path =
        crate::script_storage::ScriptStorage::get_script_path(PSScript::RefreshExplorer)?;
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

/// Refreshes Explorer windows using PowerShell with a caller-supplied timeout.
fn refresh_explorer_powershell_with_timeout(timeout: Duration) -> WincentResult<()> {
    use std::time::Instant;
    let start = Instant::now();
    let script_path =
        crate::script_storage::ScriptStorage::get_script_path(PSScript::RefreshExplorer)?;
    let output =
        ScriptExecutor::execute_ps_script_with_timeout(PSScript::RefreshExplorer, None, timeout)?;
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
        return Err(WincentError::invalid_path_reason("Empty path provided"));
    }

    if !path_buf.exists() {
        return Err(WincentError::invalid_path(path, "Path does not exist"));
    }

    match expected_type {
        PathType::File if !path_buf.is_file() => {
            Err(WincentError::invalid_path(path, "Not a valid file"))
        }
        PathType::Directory if !path_buf.is_dir() => {
            Err(WincentError::invalid_path(path, "Not a valid directory"))
        }
        _ => Ok(()),
    }
}

/// Get Windows Recent Folder path
pub(crate) fn get_windows_recent_folder() -> WincentResult<String> {
    // SAFETY: SHGetKnownFolderPath writes a heap-allocated wide string into `result`.
    // FOLDERID_Recent is a well-known constant, the flag is 0, and no token handle is needed.
    let result = unsafe { SHGetKnownFolderPath(&FOLDERID_Recent, KNOWN_FOLDER_FLAG(0x00), None) }?;

    // SAFETY: `result` is a valid PWSTR allocated by the shell (above). We
    // copy its contents into an OsString before freeing the allocation with
    // CoTaskMemFree, which is the documented cleanup method for SHGetKnownFolderPath output.
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
    fn refresh_explorer_window_with_timeout_passes_timeout_to_native() -> WincentResult<()> {
        let timeout = Duration::from_millis(4321);
        let observed = std::cell::Cell::new(None);

        refresh_explorer_window_with_timeout_using(
            timeout,
            |actual| {
                observed.set(Some(actual));
                Ok(())
            },
            |_| panic!("PowerShell fallback should not run when native succeeds"),
        )?;

        assert_eq!(observed.get(), Some(timeout));
        Ok(())
    }

    #[test]
    fn refresh_explorer_window_with_timeout_passes_timeout_to_fallback() -> WincentResult<()> {
        let timeout = Duration::from_millis(4321);
        let native_observed = std::cell::Cell::new(None);
        let fallback_observed = std::cell::Cell::new(None);

        refresh_explorer_window_with_timeout_using(
            timeout,
            |actual| {
                native_observed.set(Some(actual));
                Err(WincentError::SystemError("native failed".to_string()))
            },
            |actual| {
                fallback_observed.set(Some(actual));
                Ok(())
            },
        )?;

        assert_eq!(native_observed.get(), Some(timeout));
        assert_eq!(fallback_observed.get(), Some(timeout));
        Ok(())
    }

    #[test]
    #[ignore = "Requires desktop session and may need elevated privileges — run with: cargo test test_refresh_explorer -- --ignored --nocapture"]
    fn test_refresh_explorer() -> WincentResult<()> {
        refresh_explorer_window()
    }

    #[test]
    #[ignore = "Performance benchmark; requires desktop session — run with: cargo test test_refresh_explorer_performance -- --ignored --nocapture"]
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
    #[ignore = "Requires desktop session and may need elevated privileges — run with: cargo test test_refresh_explorer_native_only -- --ignored --nocapture"]
    fn test_refresh_explorer_native_only() -> WincentResult<()> {
        // Test that native COM version works independently
        refresh_explorer_native()
    }

    #[test]
    #[ignore = "Requires desktop session and may need elevated privileges — run with: cargo test test_refresh_explorer_powershell_only -- --ignored --nocapture"]
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
    fn test_nonexistent_path_normalization_is_consistent() {
        let path = "Z:/NonExistent/Path/";

        assert_eq!(
            normalize_path_for_comparison(path),
            normalize_path_lightweight(path),
            "nonexistent paths should use the same lightweight normalization fallback"
        );
    }

    #[test]
    fn test_paths_equal_nonexistent_paths_use_lightweight_fallback() {
        assert!(
            paths_equal("Z:/NonExistent/Path/", "z:\\nonexistent\\path"),
            "nonexistent paths should still compare case-insensitively with slash normalization"
        );
    }

    #[test]
    fn test_validate_path_file() -> WincentResult<()> {
        // Create temporary file
        let temp_file =
            tempfile::NamedTempFile::new().map_err(|e| WincentError::SystemError(e.to_string()))?;
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
        let temp_dir = tempfile::tempdir().map_err(|e| WincentError::SystemError(e.to_string()))?;
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

        if let Err(WincentError::InvalidPath(error)) = result {
            assert!(
                error.reason_text().contains("Empty"),
                "Error should mention empty path"
            );
        } else {
            panic!("Expected InvalidPath error");
        }
    }

    #[test]
    fn test_validate_path_nonexistent() {
        let result = validate_path("Z:\\NonExistent\\Path.txt", PathType::File);
        assert!(result.is_err(), "Non-existent path should fail validation");

        if let Err(WincentError::InvalidPath(error)) = result {
            assert!(
                error.reason_text().contains("does not exist"),
                "Error should mention path doesn't exist"
            );
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

        if let Err(WincentError::InvalidPath(error)) = result {
            assert!(
                error.reason_text().contains("Not a valid file"),
                "Error should mention type mismatch"
            );
        }

        Ok(())
    }
}
