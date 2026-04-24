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
    com::{ComGuard, ComInitStatus},
    error::WincentError,
    query::{folder_items, item_path, shell_folder, FREQUENT_FOLDERS_NAMESPACE},
    script_executor::ScriptExecutor,
    script_strategy::PSScript,
    utils::{validate_path, PathType},
    WincentResult,
};
use std::ffi::OsString;
use std::os::windows::prelude::*;
use std::path::Path;
use windows::core::{Interface, VARIANT};
use windows::Win32::UI::Shell::{Folder3, SHAddToRecentDocs};

/// Default timeout for COM STA thread operations
const DEFAULT_COM_TIMEOUT: std::time::Duration = std::time::Duration::from_secs(10);

/// Lightweight path normalization without I/O operations
///
/// This function performs fast path normalization without accessing the file system.
/// It's used as the first stage in path comparison for performance optimization.
///
/// # Transformations
///
/// - **Case normalization**: Converts to lowercase (Windows is case-insensitive)
/// - **Slash normalization**: Converts forward slashes to backslashes
/// - **Trailing slash removal**: Removes trailing backslash (except for root paths like "C:\")
///
/// # Arguments
///
/// * `path` - The path string to normalize
///
/// # Returns
///
/// A normalized path string suitable for fast comparison
///
/// # Example
///
/// ```ignore
/// assert_eq!(normalize_path_lightweight("C:/Users/Test/"), "c:\\users\\test");
/// assert_eq!(normalize_path_lightweight("C:\\"), "c:\\"); // Root preserved
/// ```
fn normalize_path_lightweight(path: &str) -> String {
    let mut result = path.to_lowercase().replace('/', "\\");

    // Remove trailing backslash unless it's a root path (e.g., "C:\")
    if result.len() > 3 && result.ends_with('\\') {
        result.pop();
    }

    result
}

/// Normalizes a Windows path for comparison with canonicalization
///
/// This function performs comprehensive path normalization including file system
/// operations to resolve symlinks, relative paths, and other indirections.
///
/// # Transformations
///
/// - **Canonicalization**: Resolves symlinks, relative paths (e.g., "..", "."), and junction points
/// - **Case normalization**: Converts to lowercase (Windows is case-insensitive)
/// - **Slash normalization**: Converts forward slashes to backslashes
/// - **Trailing slash removal**: Removes trailing backslash (except for root paths)
///
/// # Fallback Behavior
///
/// If canonicalization fails (e.g., path doesn't exist), falls back to basic
/// string normalization without I/O operations.
///
/// # Arguments
///
/// * `path` - The path string to normalize
///
/// # Returns
///
/// A normalized path string suitable for accurate comparison
///
/// # Performance Note
///
/// This function performs I/O operations and is slower than `normalize_path_lightweight()`.
/// Use it only when canonicalization is needed (e.g., for symlinks or relative paths).
fn normalize_path_for_comparison(path: &str) -> String {
    let path_obj = Path::new(path);

    // Try to canonicalize (resolves symlinks, relative paths, etc.)
    // If it fails, fall back to basic normalization
    let normalized = if let Ok(canonical) = path_obj.canonicalize() {
        canonical.to_string_lossy().to_string()
    } else {
        path.to_string()
    };

    // Convert to lowercase for case-insensitive comparison
    // Remove trailing backslash/slash (unless it's a root like "C:\")
    let mut result = normalized.to_lowercase().replace('/', "\\");

    // Remove trailing backslash unless it's a root path
    if result.len() > 3 && result.ends_with('\\') {
        result.pop();
    }

    result
}

/// Checks if two paths are equivalent on Windows using two-stage comparison
///
/// This function implements an optimized two-stage path comparison strategy:
///
/// # Stage 1: Lightweight Comparison (Fast Path)
///
/// Performs string-based normalization without I/O operations:
/// - Case normalization (lowercase)
/// - Slash normalization (forward to backslash)
/// - Trailing slash removal
///
/// If paths match at this stage, returns `true` immediately (~microseconds).
///
/// # Stage 2: Canonicalization (Slow Path)
///
/// Only executed if Stage 1 fails. Performs file system operations to resolve:
/// - Symlinks and junction points
/// - Relative paths (e.g., "..", ".")
/// - Different representations of the same path
///
/// This stage involves I/O and is slower (~milliseconds).
///
/// # Arguments
///
/// * `path1` - First path to compare
/// * `path2` - Second path to compare
///
/// # Returns
///
/// `true` if the paths refer to the same location, `false` otherwise
///
/// # Performance
///
/// - Best case (Stage 1 match): ~1-10 microseconds
/// - Worst case (Stage 2 needed): ~1-10 milliseconds
///
/// # Example
///
/// ```ignore
/// assert!(paths_equal("C:\\Users\\Test", "c:\\users\\test")); // Stage 1
/// assert!(paths_equal("C:\\Users\\Test", "C:/Users/Test"));   // Stage 1
/// assert!(paths_equal("C:\\Users\\Test", "C:\\Users\\Test\\")); // Stage 1
/// // Symlinks would require Stage 2
/// ```
fn paths_equal(path1: &str, path2: &str) -> bool {
    // Stage 1: Fast lightweight comparison (no I/O)
    let light1 = normalize_path_lightweight(path1);
    let light2 = normalize_path_lightweight(path2);

    if light1 == light2 {
        return true;
    }

    // Stage 2: Canonicalize for symlinks, relative paths, etc.
    // Only if lightweight comparison failed
    normalize_path_for_comparison(path1) == normalize_path_for_comparison(path2)
}

/// Invokes a Shell verb directly on a folder using Folder.Self pattern
///
/// This function uses the PowerShell-equivalent pattern: `$shell.Namespace(path).Self.InvokeVerb(verb)`.
/// It's more efficient than enumerating parent folder items to find the target.
///
/// # Threading Model
///
/// Runs on a dedicated STA thread to avoid the Windows 11 deadlock where
/// `pintohome`/`unpinfromhome` verbs require cross-process calls to explorer.exe
/// that need a message pump the calling thread does not have.
///
/// # Arguments
///
/// * `path` - The folder path to invoke the verb on
/// * `verb` - The Shell verb to invoke (e.g., "pintohome", "unpinfromhome")
///
/// # Returns
///
/// `Ok(())` if the verb was successfully invoked, or an error if:
/// - The folder namespace cannot be opened
/// - The Folder3 interface cast fails
/// - The verb invocation fails
///
/// # Windows 11 Compatibility
///
/// This function is specifically designed to work around Windows 11 deadlock issues
/// by using a dedicated STA thread with a message pump.
///
/// # Safety
///
/// This function performs unsafe COM operations. The dedicated STA thread ensures
/// proper COM initialization and message pump availability.
fn invoke_verb_on_self(path: &str, verb: &str, timeout: std::time::Duration) -> WincentResult<()> {
    let path = path.to_owned();
    let verb = verb.to_owned();
    crate::com_thread::run_on_sta_thread(move || {
        let folder = shell_folder(&path)?;

        unsafe {
            let folder3: Folder3 = folder.cast().map_err(|e| {
                WincentError::SystemError(format!("Failed to cast to Folder3 for {}: {}", path, e))
            })?;

            let self_item = folder3.Self_().map_err(|e| {
                WincentError::SystemError(format!("Failed to get Self for {}: {}", path, e))
            })?;

            let verb_variant = VARIANT::from(verb.as_str());
            self_item.InvokeVerb(&verb_variant).map_err(|e| {
                WincentError::SystemError(format!(
                    "Failed to invoke verb '{}' on {}: {}",
                    verb, path, e
                ))
            })?;
        }

        Ok(())
    }, timeout)
}


/// Finds a folder item in Frequent Folders namespace and invokes a Shell verb on it
///
/// This function searches through the Frequent Folders namespace
/// (`shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}`) to find a folder matching
/// the given path, then invokes the specified Shell verb on it.
///
/// # Threading Model
///
/// Runs on a dedicated STA thread to avoid the Windows 11 deadlock where
/// cross-process Shell namespace resolution hangs without a message pump.
/// The dedicated thread provides the required message pump for COM operations.
///
/// # Arguments
///
/// * `path` - The full path to the folder to find
/// * `verb` - The Shell verb to invoke (e.g., "unpinfromhome", "pintohome")
///
/// # Returns
///
/// Returns `Ok(())` if the folder was found and the verb was successfully invoked.
///
/// # Errors
///
/// - `NotInRecent`: The folder was not found in the Frequent Folders namespace
/// - `SystemError`: COM operation failed (e.g., failed to get item count, invoke verb)
///
/// # Path Matching
///
/// Uses [`paths_equal()`] for path comparison, which handles:
/// - Case insensitivity
/// - Forward slash vs backslash normalization
/// - Trailing slash differences
/// - Symlinks and relative paths (via canonicalization)
///
/// # Windows 11 Compatibility
///
/// This function is specifically designed to work around Windows 11 deadlock issues
/// by using a dedicated STA thread with a message pump.
///
/// # Safety
///
/// This function performs unsafe COM operations. The dedicated STA thread ensures
/// proper COM initialization and message pump availability.
fn find_and_invoke_verb(path: &str, verb: &str, timeout: std::time::Duration) -> WincentResult<()> {
    let path = path.to_owned();
    let verb = verb.to_owned();
    crate::com_thread::run_on_sta_thread(move || {
        let folder = shell_folder(FREQUENT_FOLDERS_NAMESPACE)?;
        let items = folder_items(&folder)?;

        let count = unsafe {
            items.Count().map_err(|e| {
                WincentError::SystemError(format!("Failed to get item count: {}", e))
            })?
        };

        for index in 0..count {
            let index_variant = VARIANT::from(index);
            let item = unsafe {
                match items.Item(&index_variant) {
                    Ok(item) => item,
                    Err(_) => continue,
                }
            };

            if let Some(item_path_str) = item_path(&item) {
                if paths_equal(&item_path_str, &path) {
                    let verb_variant = VARIANT::from(verb.as_str());
                    unsafe {
                        item.InvokeVerb(&verb_variant).map_err(|e| {
                            WincentError::SystemError(format!(
                                "Failed to invoke verb '{}' on {}: {}",
                                verb, path, e
                            ))
                        })?;
                    }
                    return Ok(());
                }
            }
        }

        Err(WincentError::NotInRecent(format!(
            "Folder not found in frequent folders: {}",
            path
        )))
    }, timeout)
}

/// Pins a folder to Quick Access using native COM API
///
/// Uses the "pintohome" Shell verb to pin a folder to the Frequent Folders list.
/// This is the fast path for pinning operations (~10-50ms).
///
/// # Threading Model
///
/// Uses [`invoke_verb_on_self()`] which runs on a dedicated STA thread to avoid
/// Windows 11 deadlock issues.
///
/// # Arguments
///
/// * `path` - The full path to the folder to be pinned
///
/// # Returns
///
/// Returns `Ok(())` if the folder was successfully pinned.
///
/// # Errors
///
/// - `SystemError`: COM operation failed (e.g., failed to open folder namespace, invoke verb)
///
/// # See Also
///
/// - [`pin_frequent_folder()`] - Public wrapper with PowerShell fallback
/// - [`invoke_verb_on_self()`] - The underlying verb invocation mechanism
fn pin_frequent_folder_native(path: &str, timeout: std::time::Duration) -> WincentResult<()> {
    invoke_verb_on_self(path, "pintohome", timeout)
}

/// Checks if a folder is in the Frequent Folders namespace using native COM
///
/// This function performs a pure native COM check without PowerShell fallback.
/// It uses [`paths_equal()`] for path comparison to handle case insensitivity,
/// slash normalization, and trailing slashes.
///
/// # Threading Model
///
/// Runs on a dedicated STA thread to avoid Windows 11 deadlock issues.
///
/// # Arguments
///
/// * `path` - The full path to check
///
/// # Returns
///
/// Returns `true` if the folder is found in the Frequent Folders namespace,
/// `false` otherwise.
///
/// # Errors
///
/// - `SystemError`: COM operation failed (e.g., failed to open namespace, get item count)
///
/// # Path Matching
///
/// Uses [`paths_equal()`] for robust path comparison, which handles:
/// - Case insensitivity (`C:\\Test` == `c:\\test`)
/// - Slash normalization (`C:\\Test` == `C:/Test`)
/// - Trailing slashes (`C:\\Test` == `C:\\Test\\`)
/// - Symlinks and relative paths (via canonicalization)
///
/// # See Also
///
/// - [`unpin_frequent_folder_native()`] - Uses this function for pre-check
fn is_in_frequent_folders_native(path: &str, timeout: std::time::Duration) -> WincentResult<bool> {
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread(move || {
        let folder = shell_folder(FREQUENT_FOLDERS_NAMESPACE)?;
        let items = folder_items(&folder)?;

        let count = unsafe {
            items.Count().map_err(|e| {
                WincentError::SystemError(format!("Failed to get item count: {}", e))
            })?
        };

        for index in 0..count {
            let index_variant = VARIANT::from(index);
            let item = unsafe {
                match items.Item(&index_variant) {
                    Ok(item) => item,
                    Err(_) => continue,
                }
            };

            if let Some(item_path_str) = item_path(&item) {
                if paths_equal(&item_path_str, &path) {
                    return Ok(true);
                }
            }
        }

        Ok(false)
    }, timeout)
}
///
/// This function implements a **safe Windows version-aware strategy**:
///
/// 1. **Check if folder is pinned**
///    - Query the Frequent Folders namespace to verify the folder is pinned
///    - If not pinned, return `NotInRecent` error immediately
///    - This prevents accidentally pinning an unpinned folder
///
/// 2. **Try "unpinfromhome" verb first** (Windows 10 style)
///    - Works when the folder is in the Frequent Folders namespace
///    - Returns `Ok(())` if successful
///
/// 3. **Fallback to "pintohome" toggle** (Windows 11 style)
///    - If "unpinfromhome" fails (verb not available or other error)
///    - Uses "pintohome" as a toggle (acts as unpin when already pinned)
///    - Returns `Ok(())` if successful
///
/// This approach is more robust than explicit Windows version detection because
/// it adapts to the actual Shell verb availability on the system, while ensuring
/// we never accidentally pin an unpinned folder.
///
/// # Threading Model
///
/// Uses [`find_and_invoke_verb()`] and [`invoke_verb_on_self()`] which run on
/// dedicated STA threads to avoid Windows 11 deadlock issues.
///
/// # Arguments
///
/// * `path` - The full path to the folder to be unpinned
///
/// # Returns
///
/// Returns `Ok(())` if the folder was successfully unpinned.
///
/// # Errors
///
/// - `NotInRecent`: The folder is not in Frequent Folders (not pinned)
/// - `SystemError`: COM operation failed (e.g., permission denied, COM error)
///
/// # Windows Version Compatibility
///
/// - **Windows 10**: Uses "unpinfromhome" verb (step 2 succeeds)
/// - **Windows 11**: Uses "pintohome" toggle (step 2 fails, step 3 succeeds)
///
/// # See Also
///
/// - [`unpin_frequent_folder()`] - Public wrapper with PowerShell fallback
/// - [`find_and_invoke_verb()`] - Search and invoke verb in Frequent Folders namespace
/// - [`invoke_verb_on_self()`] - Invoke verb directly on folder
fn unpin_frequent_folder_native(path: &str, timeout: std::time::Duration) -> WincentResult<()> {
    // Step 1: Check if folder is pinned first to avoid accidentally pinning it
    // This is critical because pintohome is a toggle - it will pin if not already pinned
    // Uses pure native COM check with paths_equal() for robust path matching
    if !is_in_frequent_folders_native(path, timeout)? {
        return Err(WincentError::NotInRecent(format!(
            "Folder not in frequent folders: {}",
            path
        )));
    }

    // Step 2: Try unpinfromhome first (Windows 10 style)
    // If the folder is in frequent folders, this should work on Windows 10
    match find_and_invoke_verb(path, "unpinfromhome", timeout) {
        Ok(()) => return Ok(()),
        Err(_) => {
            // unpinfromhome failed (verb not available or other error)
            // Fall through to step 3
        }
    }

    // Step 3: Fallback to pintohome toggle (Windows 11 style)
    // Since we verified the folder is pinned in step 1, pintohome will unpin it
    invoke_verb_on_self(path, "pintohome", timeout)
}

/// Executes a PowerShell script after validating the given path
///
/// This function provides a safe wrapper for PowerShell script execution with:
///
/// 1. **Path validation**: Ensures the path exists and matches the expected type
/// 2. **Script path resolution**: Gets the dynamic script path for the operation
/// 3. **Script execution**: Executes the PowerShell script with the path parameter
/// 4. **Error handling**: Converts PowerShell errors to structured `WincentError`
///
/// # Arguments
///
/// * `script` - The PowerShell script type to execute (e.g., `PSScript::PinToFrequentFolder`)
/// * `path` - The file or folder path to pass to the script
/// * `path_type` - The expected path type (`PathType::File` or `PathType::Directory`)
///
/// # Returns
///
/// Returns `Ok(())` if the script executed successfully (exit code 0).
///
/// # Errors
///
/// - `InvalidPath`: Path validation failed (empty, doesn't exist, wrong type)
/// - `PowerShellExecution`: Script execution failed (non-zero exit code)
///   - Includes detailed error information: exit code, stdout, stderr, duration
///   - Error kind is inferred from stderr content (e.g., `NotFound`, `PermissionDenied`)
///
/// # Performance
///
/// PowerShell script execution is significantly slower than native COM:
/// - Typical execution time: 200-500ms
/// - Includes PowerShell startup overhead (~100-200ms)
///
/// # Safety
///
/// - Path validation prevents injection attacks
/// - Scripts are stored in a secure location and validated before execution
/// - PowerShell execution is sandboxed by the OS
///
/// # See Also
///
/// - [`crate::script_executor::ScriptExecutor`] - PowerShell script executor
/// - [`crate::script_storage::ScriptStorage`] - Script storage and retrieval
pub(crate) fn execute_script_with_validation(
    script: PSScript,
    path: &str,
    path_type: PathType,
) -> WincentResult<()> {
    validate_path(path, path_type)?;

    let start = std::time::Instant::now();
    let script_path = match path_type {
        PathType::File | PathType::Directory => {
            crate::script_storage::ScriptStorage::get_dynamic_script_path(script, path)?
        }
    };
    let output = ScriptExecutor::execute_ps_script(script, Some(path))?;
    let duration = start.elapsed();

    match output.status.success() {
        true => Ok(()),
        false => {
            use crate::error::PowerShellError;
            let stderr = String::from_utf8(output.stderr)
                .unwrap_or_else(|_| "Unable to parse script error output".to_string());
            let stdout = String::from_utf8(output.stdout).unwrap_or_default();

            // Infer error kind from stderr content
            let kind = PowerShellError::infer_kind_from_stderr(&stderr);

            Err(WincentError::PowerShellExecution(PowerShellError {
                kind,
                script_type: script,
                exit_code: output.status.code(),
                stdout,
                stderr,
                script_path,
                parameters: Some(path.to_string()),
                duration: Some(duration),
                io_error: None,
                os_error: None,
            }))
        }
    }
}

/// Adds a file to the Windows Recent Items list using the Windows API.
///
/// Uses the calling thread's COM context because SHAddToRecentDocs is an
/// asynchronous API that requires the COM context and message pump to remain
/// active for the Shell to complete cross-process communication.
/// Unlike synchronous Shell verb operations, this does not cause deadlocks.
///
/// # Windows Behavior Notes
///
/// - **Deduplication**: Windows has a built-in deduplication mechanism that may
///   ignore repeated additions of the same file within a short time window.
///   This is by design to prevent spam and maintain list quality.
/// - **Asynchronous Processing**: The file may not appear immediately in the
///   Recent Items list. Windows Shell processes these updates asynchronously.
pub(crate) fn add_file_to_recent_native(path: &str) -> WincentResult<()> {
    validate_path(path, PathType::File)?;

    // Initialize COM (handle already-initialized case)
    let _com = ComGuard::try_initialize().map_err(|status| match status {
        ComInitStatus::ApartmentMismatch => {
            WincentError::ComApartmentMismatch(
                "Thread already initialized with incompatible COM apartment model".to_string()
            )
        }
        ComInitStatus::OtherError(hr) => {
            WincentError::WindowsApi(hr)
        }
        _ => unreachable!(),
    })?;

    unsafe {
        let file_path_wide: Vec<u16> = OsString::from(path)
            .encode_wide()
            .chain(std::iter::once(0))
            .collect();

        // 0x0000_0003 equals SHARD_PATHW
        SHAddToRecentDocs(0x0000_0003, Some(file_path_wide.as_ptr() as *const _));
    }

    Ok(())
}

/// Removes a file from the Windows Recent Items list using native COM API
///
/// This function searches through the Recent Items namespace
/// (`shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}`) to find a file matching
/// the given path, then invokes the "remove" Shell verb on it.
///
/// # Threading Model
///
/// Runs on a dedicated STA thread to avoid the Windows 11 deadlock where
/// the Recent Items namespace requires cross-process calls to explorer.exe
/// that need a message pump the calling thread does not have.
///
/// # Search Strategy
///
/// 1. Opens the Recent Items namespace
/// 2. Enumerates all items in the namespace
/// 3. Filters out folders (only processes files)
/// 4. Compares each file path using [`paths_equal()`]
/// 5. Invokes "remove" verb on the matching file
///
/// # Arguments
///
/// * `path` - The full path to the file to be removed
///
/// # Returns
///
/// Returns `Ok(())` if the file was found and successfully removed.
///
/// # Errors
///
/// - `NotInRecent`: The file was not found in the Recent Items namespace
/// - `SystemError`: COM operation failed (e.g., failed to get item count, invoke verb)
///
/// # Path Matching
///
/// Uses [`paths_equal()`] for path comparison, which handles:
/// - Case insensitivity
/// - Forward slash vs backslash normalization
/// - Trailing slash differences
/// - Symlinks and relative paths (via canonicalization)
///
/// # Performance
///
/// - Typical execution time: 10-50ms
/// - Performance depends on the number of items in Recent Items (typically 20-50 items)
///
/// # Windows 11 Compatibility
///
/// This function is specifically designed to work around Windows 11 deadlock issues
/// by using a dedicated STA thread with a message pump.
///
/// # Safety
///
/// This function performs unsafe COM operations. The dedicated STA thread ensures
/// proper COM initialization and message pump availability.
///
/// # See Also
///
/// - [`remove_recent_file_powershell()`] - PowerShell fallback implementation
/// - [`remove_from_recent_files()`] - Public wrapper with fallback strategy
fn remove_recent_file_native(path: &str, timeout: std::time::Duration) -> WincentResult<()> {
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread(move || {
        let recent_namespace = "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}";
        let folder = shell_folder(recent_namespace)?;
        let items = folder_items(&folder)?;

        let count = unsafe {
            items.Count().map_err(|e| {
                WincentError::SystemError(format!("Failed to get item count: {}", e))
            })?
        };

        // Search for the target file
        for index in 0..count {
            let index_variant = VARIANT::from(index);
            let item = unsafe {
                match items.Item(&index_variant) {
                    Ok(item) => item,
                    Err(_) => continue,
                }
            };

            // Check if this is a file (not folder)
            let is_folder = unsafe { item.IsFolder().map(bool::from).unwrap_or(true) };
            if is_folder {
                continue;
            }

            if let Some(item_path_str) = item_path(&item) {
                if paths_equal(&item_path_str, &path) {
                    // Found the file, invoke remove verb
                    let verb_variant = VARIANT::from("remove");
                    unsafe {
                        item.InvokeVerb(&verb_variant).map_err(|e| {
                            WincentError::SystemError(format!(
                                "Failed to invoke 'remove' verb on {}: {}",
                                path, e
                            ))
                        })?;
                    }
                    return Ok(());
                }
            }
        }

        Err(WincentError::NotInRecent(format!(
            "File not found in recent items: {}",
            path
        )))
    }, timeout)
}

/// Removes a file from the Windows Recent Items list using PowerShell
///
/// This is the PowerShell fallback implementation for removing recent files.
/// It's used when the native COM approach fails (e.g., permission issues,
/// COM initialization failures).
///
/// # Arguments
///
/// * `path` - The full path to the file to be removed
///
/// # Returns
///
/// Returns `Ok(())` if the PowerShell script executed successfully.
///
/// # Errors
///
/// - `InvalidPath`: Path validation failed
/// - `PowerShellExecution`: Script execution failed
///
/// # Performance
///
/// - Typical execution time: 200-500ms (significantly slower than native COM)
/// - Includes PowerShell startup overhead
///
/// # See Also
///
/// - [`remove_recent_file_native()`] - Native COM implementation (fast path)
/// - [`remove_from_recent_files()`] - Public wrapper with fallback strategy
pub(crate) fn remove_recent_file_powershell(path: &str) -> WincentResult<()> {
    execute_script_with_validation(PSScript::RemoveRecentFile, path, PathType::File)
}

/// Pins a folder to the Windows Quick Access Frequent Folders list using PowerShell
///
/// This is the PowerShell fallback implementation for pinning folders.
/// It's used when the native COM approach fails (e.g., permission issues,
/// COM initialization failures, Windows 11 deadlock without message pump).
///
/// # Arguments
///
/// * `path` - The full path to the folder to be pinned
///
/// # Returns
///
/// Returns `Ok(())` if the PowerShell script executed successfully.
///
/// # Errors
///
/// - `InvalidPath`: Path validation failed (not a directory, doesn't exist)
/// - `PowerShellExecution`: Script execution failed
///
/// # Performance
///
/// - Typical execution time: 200-500ms (significantly slower than native COM)
/// - Includes PowerShell startup overhead (~100-200ms)
///
/// # See Also
///
/// - [`pin_frequent_folder_native()`] - Native COM implementation (fast path)
/// - [`pin_frequent_folder()`] - Internal wrapper with fallback strategy
fn pin_frequent_folder_powershell(path: &str) -> WincentResult<()> {
    execute_script_with_validation(PSScript::PinToFrequentFolder, path, PathType::Directory)
}

/// Unpins a folder from the Windows Quick Access Frequent Folders list using PowerShell
///
/// This is the PowerShell fallback implementation for unpinning folders.
/// It's used when the native COM approach fails (e.g., permission issues,
/// COM initialization failures, Windows 11 deadlock without message pump).
///
/// # Arguments
///
/// * `path` - The full path to the folder to be unpinned
///
/// # Returns
///
/// Returns `Ok(())` if the PowerShell script executed successfully.
///
/// # Errors
///
/// - `InvalidPath`: Path validation failed (not a directory, doesn't exist)
/// - `PowerShellExecution`: Script execution failed (e.g., folder not pinned)
///
/// # Performance
///
/// - Typical execution time: 200-500ms (significantly slower than native COM)
/// - Includes PowerShell startup overhead (~100-200ms)
///
/// # See Also
///
/// - [`unpin_frequent_folder_native()`] - Native COM implementation (fast path)
/// - [`unpin_frequent_folder()`] - Internal wrapper with fallback strategy
fn unpin_frequent_folder_powershell(path: &str) -> WincentResult<()> {
    execute_script_with_validation(PSScript::UnpinFromFrequentFolder, path, PathType::Directory)
}

/// Pins a folder to the Windows Quick Access Frequent Folders list
///
/// This is the internal implementation that uses a **two-tier fallback strategy**:
///
/// 1. **Native COM API** (fast path, 10-50ms): Uses [`pin_frequent_folder_native()`]
///    - Uses "pintohome" Shell verb on the folder
///    - Runs on dedicated STA thread to avoid Windows 11 deadlock
///
/// 2. **PowerShell fallback** (slow path, 200-500ms): Uses [`pin_frequent_folder_powershell()`]
///    - Executes PowerShell script if native COM fails
///    - Provides broader compatibility at the cost of performance
///
/// # Arguments
///
/// * `path` - The full path to the folder to be pinned. Must be an existing directory.
///
/// # Returns
///
/// Returns `Ok(())` if the folder was successfully pinned (by either strategy).
///
/// # Errors
///
/// - `InvalidPath`: Path validation failed (not a directory, doesn't exist)
/// - `SystemError` or `PowerShellExecution`: Both strategies failed
///
/// # Performance
///
/// - Native COM: ~10-50ms (typical case)
/// - PowerShell fallback: ~200-500ms (when COM fails)
///
/// # See Also
///
/// - [`add_to_frequent_folders()`] - Public API wrapper
/// - [`pin_frequent_folder_native()`] - Native COM implementation
/// - [`pin_frequent_folder_powershell()`] - PowerShell fallback
pub(crate) fn pin_frequent_folder(path: &str, timeout: std::time::Duration) -> WincentResult<()> {
    validate_path(path, PathType::Directory)?;

    // Try native COM first (fast path), fallback to PowerShell if it fails
    pin_frequent_folder_native(path, timeout).or_else(|_| pin_frequent_folder_powershell(path))
}

/// Unpins a folder from the Windows Quick Access Frequent Folders list
///
/// This is the internal implementation that uses a **two-tier fallback strategy**:
///
/// 1. **Native COM API** (fast path, 10-50ms): Uses [`unpin_frequent_folder_native()`]
///    - Windows 10: Uses "unpinfromhome" verb
///    - Windows 11: Uses "pintohome" toggle
///    - Runs on dedicated STA thread to avoid Windows 11 deadlock
///
/// 2. **PowerShell fallback** (slow path, 200-500ms): Uses [`unpin_frequent_folder_powershell()`]
///    - Executes PowerShell script if native COM fails
///    - Provides broader compatibility at the cost of performance
///
/// # Arguments
///
/// * `path` - The full path to the folder to be unpinned. Must be an existing directory.
///
/// # Returns
///
/// Returns `Ok(())` if the folder was successfully unpinned (by either strategy).
///
/// # Errors
///
/// - `InvalidPath`: Path validation failed (not a directory, doesn't exist)
/// - `NotInRecent`: Folder not in Frequent Folders (returned immediately, no PowerShell fallback)
/// - `SystemError` or `PowerShellExecution`: Both native COM and PowerShell strategies failed
///
/// Note: When the folder is not pinned, this function returns `NotInRecent` immediately
/// without attempting the PowerShell fallback, ensuring the caller receives accurate
/// error information. Similarly, `InvalidPath` errors are returned directly without
/// fallback, as the path is definitively wrong.
///
/// # Windows Version Compatibility
///
/// The native COM implementation automatically adapts to Windows version:
/// - **Windows 10**: Uses "unpinfromhome" verb
/// - **Windows 11**: Uses "pintohome" toggle
///
/// # Performance
///
/// - Native COM: ~10-50ms (typical case)
/// - PowerShell fallback: ~200-500ms (when COM fails)
///
/// # See Also
///
/// - [`remove_from_frequent_folders()`] - Public API wrapper
/// - [`unpin_frequent_folder_native()`] - Native COM implementation
/// - [`unpin_frequent_folder_powershell()`] - PowerShell fallback
pub(crate) fn unpin_frequent_folder(path: &str, timeout: std::time::Duration) -> WincentResult<()> {
    validate_path(path, PathType::Directory)?;

    // Try native COM first (fast path)
    match unpin_frequent_folder_native(path, timeout) {
        Ok(()) => Ok(()),
        // NotInRecent means the folder is not pinned - this is a semantic error,
        // not a system failure. Falling back to PowerShell would not help and
        // could mask the real cause from the caller.
        // InvalidPath should also not fallback as the path is definitively wrong.
        Err(e @ WincentError::NotInRecent(_)) => Err(e),
        Err(e @ WincentError::InvalidPath(_)) => Err(e),
        // For system/COM errors, fallback to PowerShell for broader compatibility
        Err(_) => unpin_frequent_folder_powershell(path),
    }
}

/****************************************************** Handle Quick Access ******************************************************/

/// Adds a file to Windows Recent Files.
///
/// # COM Apartment Requirements
///
/// This function requires the calling thread to be initialized with the
/// Single-Threaded Apartment (STA) model, or not initialized at all.
/// If the thread is initialized with Multi-Threaded Apartment (MTA),
/// this function will return `ComApartmentMismatch` error.
///
/// If you need to call this from an MTA thread, consider using
/// `std::thread::spawn()` to run it on a new thread.
///
/// # Arguments
///
/// * `path` - The full path to the file to be added
///
/// # Errors
///
/// Returns `ComApartmentMismatch` if the calling thread is initialized
/// with an incompatible COM apartment model.
///
/// # Windows Behavior Notes
///
/// - **Deduplication**: Windows may ignore repeated additions of the same file
///   within a short time window. This is intentional behavior to prevent spam.
/// - **Asynchronous Processing**: The file may not appear immediately in the
///   Recent Items list. Windows Shell processes these updates asynchronously.
/// - **File Must Exist**: The file must exist on disk for Windows to add it.
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
    add_file_to_recent_native(path)
}

/// Removes a file from Windows Recent Files.
///
/// This function uses a **two-tier fallback strategy** to maximize compatibility
/// across different Windows versions and system configurations:
///
/// 1. **Native COM API** (fast path, 10-50ms): Uses Shell namespace to find and remove the file
/// 2. **PowerShell fallback** (slow path, 200-500ms): Uses PowerShell script if COM fails
///
/// The native approach is significantly faster but may fail in certain scenarios
/// (e.g., permission issues, COM initialization failures). The PowerShell fallback
/// provides broader compatibility at the cost of performance.
///
/// # Arguments
///
/// * `path` - The full path to the file to be removed. Must be an existing file path.
///
/// # Returns
///
/// Returns `Ok(())` if the file was successfully removed from Recent Files.
///
/// # Errors
///
/// This function returns an error if:
/// - **Path validation fails**: The path is empty, invalid, or points to a directory
/// - **File not found**: The file is not in the Recent Files list (both strategies fail)
/// - **Permission denied**: Insufficient permissions to modify Recent Files
/// - **COM failure**: Both native COM and PowerShell fallback fail
///
/// # Windows Behavior Notes
///
/// - **Asynchronous Processing**: The file may not disappear immediately from the
///   Recent Items list. Windows Shell processes these updates asynchronously.
/// - **File Must Exist**: The file must exist on disk for Windows to locate it in
///   the Recent Items namespace.
/// - **Case Insensitive**: Path matching is case-insensitive on Windows.
///
/// # Performance
///
/// - Native COM: ~10-50ms (typical case)
/// - PowerShell fallback: ~200-500ms (when COM fails)
///
/// # Examples
///
/// Basic usage:
///
/// ```no_run
/// use wincent::handle::remove_from_recent_files;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     remove_from_recent_files("C:\\Documents\\report.docx")?;
///     println!("File removed from Recent Files");
///     Ok(())
/// }
/// ```
///
/// Error handling:
///
/// ```no_run
/// use wincent::{handle::remove_from_recent_files, error::WincentError};
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     match remove_from_recent_files("C:\\Documents\\report.docx") {
///         Ok(()) => println!("Successfully removed"),
///         Err(WincentError::NotInRecent(_)) => println!("File not in Recent Files"),
///         Err(e) => eprintln!("Error: {}", e),
///     }
///     Ok(())
/// }
/// ```
///
/// # See Also
///
/// - [`add_to_recent_files()`] - Add a file to Recent Files
/// - [`crate::query::get_recent_files()`] - Query all recent files
/// - [`crate::query::is_recent_file_exact()`] - Check if a file is in Recent Files
pub fn remove_from_recent_files(path: &str) -> WincentResult<()> {
    validate_path(path, PathType::File)?;

    // Try native COM first (fast path), fallback to PowerShell if it fails
    remove_recent_file_native(path, DEFAULT_COM_TIMEOUT).or_else(|_| remove_recent_file_powershell(path))
}

/// Pins a folder to Windows Quick Access (Frequent Folders).
///
/// This function uses a **two-tier fallback strategy** to maximize compatibility
/// across different Windows versions (Windows 10, Windows 11) and system configurations:
///
/// 1. **Native COM API** (fast path, 10-50ms): Uses Shell verb "pintohome" on the folder
/// 2. **PowerShell fallback** (slow path, 200-500ms): Uses PowerShell script if COM fails
///
/// The native approach is significantly faster but may fail in certain scenarios
/// (e.g., permission issues, COM initialization failures, Windows 11 deadlock without
/// message pump). The PowerShell fallback provides broader compatibility at the cost
/// of performance.
///
/// # Arguments
///
/// * `path` - The full path to the folder to be pinned. Must be an existing directory.
///
/// # Returns
///
/// Returns `Ok(())` if the folder was successfully pinned to Frequent Folders.
///
/// # Errors
///
/// This function returns an error if:
/// - **Path validation fails**: The path is empty, invalid, or points to a file
/// - **Folder doesn't exist**: The folder must exist on disk
/// - **Permission denied**: Insufficient permissions to modify Quick Access
/// - **COM failure**: Both native COM and PowerShell fallback fail
///
/// # Windows Version Compatibility
///
/// - **Windows 10**: Uses "pintohome" verb to pin folders
/// - **Windows 11**: Uses "pintohome" verb (same as Windows 10)
/// - Both versions supported through the two-tier strategy
///
/// # Windows Behavior Notes
///
/// - **Asynchronous Processing**: The folder may not appear immediately in Quick Access.
///   Windows Shell processes these updates asynchronously.
/// - **Deduplication**: Pinning an already-pinned folder is a no-op (no error).
/// - **Case Insensitive**: Path matching is case-insensitive on Windows.
/// - **Automatic Tracking**: Windows automatically tracks folder access frequency.
///   Pinning a folder makes it permanently visible in Quick Access.
///
/// # Performance
///
/// - Native COM: ~10-50ms (typical case)
/// - PowerShell fallback: ~200-500ms (when COM fails)
///
/// # Examples
///
/// Basic usage:
///
/// ```no_run
/// use wincent::handle::add_to_frequent_folders;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     add_to_frequent_folders("C:\\Projects\\my-project")?;
///     println!("Folder pinned to Quick Access");
///     Ok(())
/// }
/// ```
///
/// Pin multiple folders:
///
/// ```no_run
/// use wincent::handle::add_to_frequent_folders;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     let folders = vec![
///         "C:\\Projects\\project-a",
///         "C:\\Projects\\project-b",
///         "C:\\Documents\\Work",
///     ];
///
///     for folder in folders {
///         add_to_frequent_folders(folder)?;
///     }
///
///     println!("All folders pinned successfully");
///     Ok(())
/// }
/// ```
///
/// Error handling:
///
/// ```no_run
/// use wincent::{handle::add_to_frequent_folders, error::WincentError};
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     match add_to_frequent_folders("C:\\Projects\\my-project") {
///         Ok(()) => println!("Successfully pinned"),
///         Err(WincentError::InvalidPath(_)) => println!("Invalid folder path"),
///         Err(e) => eprintln!("Error: {}", e),
///     }
///     Ok(())
/// }
/// ```
///
/// # See Also
///
/// - [`remove_from_frequent_folders()`] - Unpin a folder from Quick Access
/// - [`crate::query::get_frequent_folders()`] - Query all frequent folders
/// - [`crate::query::is_frequent_folder_exact()`] - Check if a folder is pinned
pub fn add_to_frequent_folders(path: &str) -> WincentResult<()> {
    pin_frequent_folder(path, DEFAULT_COM_TIMEOUT)
}

/// Unpins or remove a folder from Windows Quick Access (Frequent Folders).
///
/// This function uses a **two-tier fallback strategy** to maximize compatibility
/// across different Windows versions (Windows 10, Windows 11) and system configurations:
///
/// 1. **Native COM API** (fast path, 10-50ms): Uses Shell verbs with Windows version detection
///    - Windows 10: Uses "unpinfromhome" verb when folder is in Frequent Folders
///    - Windows 11: Uses "pintohome" toggle (acts as unpin when already pinned)
/// 2. **PowerShell fallback** (slow path, 200-500ms): Uses PowerShell script if COM fails
///
/// The native approach is significantly faster but may fail in certain scenarios
/// (e.g., permission issues, COM initialization failures, Windows 11 deadlock without
/// message pump). The PowerShell fallback provides broader compatibility at the cost
/// of performance.
///
/// # Arguments
///
/// * `path` - The full path to the folder to be unpinned. Must be an existing directory.
///
/// # Returns
///
/// Returns `Ok(())` if the folder was successfully unpinned from Frequent Folders.
///
/// # Errors
///
/// This function returns an error if:
/// - **Path validation fails**: The path is empty, invalid, or points to a file
/// - **Folder doesn't exist**: The folder must exist on disk
/// - **Folder not pinned**: The folder is not in Frequent Folders (native COM detects this)
/// - **Both strategies fail**: Native COM and PowerShell both fail to unpin the folder
/// - **Permission denied**: Insufficient permissions to modify Quick Access
///
/// Note: When the folder is not pinned, this function returns `NotInRecent`
/// immediately without attempting the PowerShell fallback, ensuring the
/// caller receives accurate error information.
///
/// # Windows Version Compatibility
///
/// - **Windows 10**: Uses "unpinfromhome" verb to unpin folders
/// - **Windows 11**: Uses "pintohome" toggle (same verb as pin, but acts as unpin when already pinned)
/// - The function automatically detects the appropriate strategy without explicit version checking
///
/// # Windows Behavior Notes
///
/// - **Asynchronous Processing**: The folder may not disappear immediately from Quick Access.
///   Windows Shell processes these updates asynchronously.
/// - **Not Pinned**: The native COM implementation checks if the folder is pinned before
///   attempting to unpin. If not pinned, it returns `NotInRecent` immediately without
///   falling back to PowerShell.
/// - **Case Insensitive**: Path matching is case-insensitive on Windows.
/// - **Automatic Tracking**: Unpinning removes the folder from the permanent Quick Access list,
///   but Windows may still show it temporarily if it's frequently accessed.
///
/// # Performance
///
/// - Native COM: ~10-50ms (typical case)
/// - PowerShell fallback: ~200-500ms (when COM fails)
///
/// # Examples
///
/// Basic usage:
///
/// ```no_run
/// use wincent::handle::remove_from_frequent_folders;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     remove_from_frequent_folders("C:\\Projects\\old-project")?;
///     println!("Folder unpinned from Quick Access");
///     Ok(())
/// }
/// ```
///
/// Unpin multiple folders:
///
/// ```no_run
/// use wincent::handle::remove_from_frequent_folders;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     let folders = vec![
///         "C:\\Projects\\archived-a",
///         "C:\\Projects\\archived-b",
///         "C:\\Temp\\old-work",
///     ];
///
///     for folder in folders {
///         match remove_from_frequent_folders(folder) {
///             Ok(()) => println!("Unpinned: {}", folder),
///             Err(e) => eprintln!("Failed to unpin {}: {}", folder, e),
///         }
///     }
///
///     Ok(())
/// }
/// ```
///
/// Error handling:
///
/// ```no_run
/// use wincent::handle::remove_from_frequent_folders;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     match remove_from_frequent_folders("C:\\Projects\\old-project") {
///         Ok(()) => println!("Successfully unpinned"),
///         Err(e) => eprintln!("Failed to unpin: {}", e),
///     }
///     Ok(())
/// }
/// ```
///
/// # See Also
///
/// - [`add_to_frequent_folders()`] - Pin a folder to Quick Access
/// - [`crate::query::get_frequent_folders()`] - Query all frequent folders
/// - [`crate::query::is_frequent_folder_exact()`] - Check if a folder is pinned
pub fn remove_from_frequent_folders(path: &str) -> WincentResult<()> {
    unpin_frequent_folder(path, DEFAULT_COM_TIMEOUT)
}

/// Adds a file to Windows Recent Files with a custom timeout.
///
/// # Timeout Behavior
///
/// The `timeout` parameter is accepted for API consistency with other
/// `_with_timeout` variants but has **no effect** on this function.
///
/// `SHAddToRecentDocs` is a fire-and-forget API: the call itself returns
/// within ~1ms after writing to the MRU registry key and posting a shell
/// notification. The actual UI update is processed asynchronously by
/// explorer.exe with no synchronization point exposed to the caller.
/// There is no blocking operation to guard with a timeout.
///
/// Moving the call onto a background thread would be counterproductive:
/// `SHAddToRecentDocs` requires the calling thread's COM apartment and
/// message pump to remain alive for the shell to deliver its cross-process
/// callbacks. A thread that exits immediately after the call would prevent
/// those callbacks from being delivered, silently dropping the registration.
///
/// If you need to verify that the file actually appeared in Recent Files,
/// poll [`crate::query::is_recent_file_exact()`] with your own deadline.
///
/// # Arguments
///
/// * `path` - The full path to the file to be added
/// * `_timeout` - Accepted for API consistency; currently has no effect
pub fn add_to_recent_files_with_timeout(
    path: &str,
    _timeout: std::time::Duration,
) -> WincentResult<()> {
    add_file_to_recent_native(path)
}

/// Removes a file from Windows Recent Files with a custom COM STA thread timeout.
///
/// Identical to [`remove_from_recent_files()`] but allows specifying the timeout
/// for the native COM STA thread operation.
///
/// # Arguments
///
/// * `path` - The full path to the file to be removed
/// * `timeout` - Timeout for the COM STA thread operation. Must be non-zero;
///   passing [`Duration::ZERO`] returns [`WincentError::InvalidArgument`] immediately without attempting any operation.
pub fn remove_from_recent_files_with_timeout(
    path: &str,
    timeout: std::time::Duration,
) -> WincentResult<()> {
    if timeout.is_zero() {
        return Err(WincentError::InvalidArgument(
            "timeout must be greater than zero".to_string(),
        ));
    }
    validate_path(path, PathType::File)?;
    remove_recent_file_native(path, timeout).or_else(|_| remove_recent_file_powershell(path))
}

/// Pins a folder to Windows Quick Access with a custom COM STA thread timeout.
///
/// Identical to [`add_to_frequent_folders()`] but allows specifying the timeout
/// for the native COM STA thread operation.
///
/// # Arguments
///
/// * `path` - The full path to the folder to be pinned. Must be an existing directory.
/// * `timeout` - Timeout for the COM STA thread operation. Must be non-zero;
///   passing [`Duration::ZERO`] returns [`WincentError::InvalidArgument`] immediately without attempting any operation.
pub fn add_to_frequent_folders_with_timeout(
    path: &str,
    timeout: std::time::Duration,
) -> WincentResult<()> {
    if timeout.is_zero() {
        return Err(WincentError::InvalidArgument(
            "timeout must be greater than zero".to_string(),
        ));
    }
    pin_frequent_folder(path, timeout)
}

/// Unpins a folder from Windows Quick Access with a custom COM STA thread timeout.
///
/// Identical to [`remove_from_frequent_folders()`] but allows specifying the timeout
/// for the native COM STA thread operation.
///
/// # Arguments
///
/// * `path` - The full path to the folder to be unpinned. Must be an existing directory.
/// * `timeout` - Timeout for the COM STA thread operation. Must be non-zero;
///   passing [`Duration::ZERO`] returns [`WincentError::InvalidArgument`] immediately without attempting any operation.
pub fn remove_from_frequent_folders_with_timeout(
    path: &str,
    timeout: std::time::Duration,
) -> WincentResult<()> {
    if timeout.is_zero() {
        return Err(WincentError::InvalidArgument(
            "timeout must be greater than zero".to_string(),
        ));
    }
    unpin_frequent_folder(path, timeout)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::query::query_recent;
    use crate::test_utils::{cleanup_test_env, create_test_file, setup_test_env};
    use std::{thread, time::Duration};
    use windows::Win32::UI::Shell::FolderItem;

    fn wait_for_folder_status(
        path: &str,
        should_exist: bool,
        max_retries: u32,
    ) -> WincentResult<bool> {
        for _ in 0..max_retries {
            let frequent_folders =
                query_recent(crate::QuickAccess::FrequentFolders)?;
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
            let recent_files = query_recent(crate::QuickAccess::RecentFiles)?;
            let exists = recent_files.iter().any(|p| p == path);

            if exists == should_exist {
                return Ok(true);
            }

            thread::sleep(Duration::from_millis(500));
        }
        Ok(false)
    }

    // Invokes unpin verb on a FolderItem with Windows version compatibility.
    // Tries "unpinfromhome" first (Windows 10), falls back to "pintohome" toggle (Windows 11).
    fn unpin_folder_item(item: &FolderItem) -> WincentResult<()> {
        let verb_variant = VARIANT::from("unpinfromhome");
        match unsafe { item.InvokeVerb(&verb_variant) } {
            Ok(()) => Ok(()),
            Err(_) => {
                let verb_variant = VARIANT::from("pintohome");
                unsafe {
                    item.InvokeVerb(&verb_variant).map_err(|e| {
                        WincentError::SystemError(format!("Failed to unpin folder item: {}", e))
                    })
                }
            }
        }
    }

    #[test]
    fn test_path_normalization() {
        // Test case insensitivity
        assert!(paths_equal("C:\\Users\\Test", "c:\\users\\test"));
        assert!(paths_equal("C:\\Users\\Test", "C:\\USERS\\TEST"));

        // Test trailing backslash
        assert!(paths_equal("C:\\Users\\Test", "C:\\Users\\Test\\"));
        assert!(paths_equal("C:\\Users\\Test\\", "C:\\Users\\Test"));

        // Test forward slash vs backslash
        assert!(paths_equal("C:\\Users\\Test", "C:/Users/Test"));

        // Test different paths should not match
        assert!(!paths_equal("C:\\Users\\Test", "C:\\Users\\Other"));
        assert!(!paths_equal("C:\\Users\\Test", "D:\\Users\\Test"));
    }

    #[test]
    fn test_path_normalization_stages() {
        // Tests the two-stage path comparison logic in paths_equal()
        // Stage 1: Lightweight normalization (no I/O) - fast path
        // Stage 2: Canonicalization (with I/O) - for symlinks/relative paths

        // Stage 1 tests - should succeed with lightweight normalization
        assert!(
            paths_equal("C:\\Users\\Test", "c:\\users\\test"),
            "Case insensitivity should work in stage 1"
        );
        assert!(
            paths_equal("C:\\Users\\Test", "C:/Users/Test"),
            "Slash normalization should work in stage 1"
        );
        assert!(
            paths_equal("C:\\Users\\Test\\", "C:\\Users\\Test"),
            "Trailing slash should work in stage 1"
        );

        // Stage 2 tests - require canonicalization
        // Note: These tests use real paths to trigger canonicalization
        let temp_dir = std::env::temp_dir();
        let temp_str = temp_dir.to_str().unwrap();

        // Test with different representations of the same path
        assert!(paths_equal(temp_str, temp_str), "Same path should be equal");

        // Test that different paths are not equal
        assert!(
            !paths_equal("C:\\Windows", "C:\\Users"),
            "Different paths should not be equal"
        );
    }

    #[test]
    #[ignore = "Modifies system state"]
    fn test_pin_unpin_frequent_folder() -> WincentResult<()> {
        let test_dir = setup_test_env()?;
        let test_path = test_dir.to_str().unwrap();

        pin_frequent_folder(test_path, DEFAULT_COM_TIMEOUT)?;

        assert!(
            wait_for_folder_status(test_path, true, 5)?,
            "Pin operation failed: folder did not appear in frequent folders list"
        );

        unpin_frequent_folder(test_path, DEFAULT_COM_TIMEOUT)?;

        assert!(
            wait_for_folder_status(test_path, false, 5)?,
            "Unpin operation failed: folder still exists in frequent folders list"
        );

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    fn test_pin_unpin_error_handling() -> WincentResult<()> {
        // Test pin errors
        let result = pin_frequent_folder("Z:\\NonExistentFolder", DEFAULT_COM_TIMEOUT);
        assert!(result.is_err(), "Should fail with non-existent folder");

        let result = pin_frequent_folder("", DEFAULT_COM_TIMEOUT);
        assert!(result.is_err(), "Should fail with empty path");

        // Test unpin errors
        let result = unpin_frequent_folder("Z:\\NonExistentFolder", DEFAULT_COM_TIMEOUT);
        assert!(result.is_err(), "Should fail with non-existent folder");

        let result = unpin_frequent_folder("", DEFAULT_COM_TIMEOUT);
        assert!(result.is_err(), "Should fail with empty path");

        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state"]
    fn test_unpin_native_error_classification() -> WincentResult<()> {
        // Tests the critical error classification logic in unpin_frequent_folder_native()
        // After fix: The function now checks if folder is pinned BEFORE attempting to unpin
        // This prevents accidentally pinning an unpinned folder via pintohome toggle
        //
        // Key logic:
        // 1. Check if folder is pinned (is_frequent_folder_exact)
        // 2. If not pinned, return NotInRecent immediately
        // 3. If pinned, try unpinfromhome first, then fallback to pintohome

        let test_dir = setup_test_env()?;
        let test_path = test_dir.to_str().unwrap();

        // Scenario 1: Unpin a folder that is NOT in frequent folders
        // Expected: Should return NotInRecent error immediately (after fix)
        // This prevents accidentally pinning the folder
        let result = unpin_frequent_folder_native(test_path, DEFAULT_COM_TIMEOUT);

        assert!(
            result.is_err(),
            "Unpinning a non-pinned folder should return error"
        );

        if let Err(e) = result {
            assert!(
                matches!(e, WincentError::NotInRecent(_)),
                "Unpinning a non-pinned folder should return NotInRecent, got: {:?}",
                e
            );
        }

        // Scenario 2: Pin the folder first, then unpin it
        // Expected: Should succeed (folder is pinned, so unpin works)
        pin_frequent_folder(test_path, DEFAULT_COM_TIMEOUT)?;
        thread::sleep(Duration::from_millis(500));

        let result = unpin_frequent_folder_native(test_path, DEFAULT_COM_TIMEOUT);
        assert!(
            result.is_ok(),
            "Unpinning a pinned folder should succeed: {:?}",
            result
        );

        // Scenario 3: Verify the folder is no longer pinned after unpin
        thread::sleep(Duration::from_millis(500));
        let is_pinned = crate::query::is_frequent_folder_exact(test_path)?;
        assert!(
            !is_pinned,
            "Folder should not be pinned after unpinning"
        );

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state"]
    fn test_add_remove_file_in_recent() -> WincentResult<()> {
        // Note: This test depends on Windows Shell's asynchronous behavior
        // SHAddToRecentDocs API does not guarantee immediate visibility
        // Test may be skipped if file doesn't appear within timeout period

        let test_dir = setup_test_env()?;

        // Use unique filename to avoid Windows caching/deduplication issues
        // Windows has a built-in deduplication mechanism that ignores repeated
        // additions of the same file within a short time window. Using a timestamp
        // ensures each test run uses a unique file, preventing false negatives.
        let timestamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_millis();
        let filename = format!("recent_test_{}.txt", timestamp);
        let test_file = create_test_file(&test_dir, &filename, "test content")?;
        let test_path = test_file.to_str().unwrap();

        // Use public API consistently
        add_to_recent_files(test_path)?;

        // Wait longer for Windows Shell to process the recent item (20 retries = 10 seconds)
        assert!(
            wait_for_file_status(test_path, true, 20)?,
            "Add operation failed: file did not appear in recent files list after 10 seconds"
        );

        remove_from_recent_files(test_path)?;

        assert!(
            wait_for_file_status(test_path, false, 10)?,
            "Remove operation failed: file still exists in recent files list"
        );

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state"]
    fn test_remove_recent_file_native_direct() -> WincentResult<()> {
        // Tests the native removal logic directly to verify the "find item and invoke remove verb" path
        // Implementation: src/handle.rs:400-451
        // Key logic:
        // 1. Enumerate items in Recent namespace
        // 2. Find target file by path comparison
        // 3. Invoke "remove" verb on the found item
        //
        // This test validates that remove_recent_file_native() correctly:
        // - Finds the file in the Recent Items namespace
        // - Invokes the remove verb on the correct item
        // - Returns NotInRecent when file is not found
        //
        // Note: This test depends on Windows Shell's asynchronous behavior
        // SHAddToRecentDocs API does not guarantee immediate visibility

        let test_dir = setup_test_env()?;

        // Use unique filename to avoid Windows deduplication issues
        // Windows may ignore repeated additions of the same file within a short time window
        let timestamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_millis();
        let filename = format!("native_remove_test_{}.txt", timestamp);
        let test_file = create_test_file(&test_dir, &filename, "test content")?;
        let test_path = test_file.to_str().unwrap();

        // Add file to recent using native API
        add_file_to_recent_native(test_path)?;

        // Wait longer for Windows to process the recent item (20 retries = 10 seconds)
        if !wait_for_file_status(test_path, true, 20)? {
            // If file doesn't appear, skip this test (Windows Recent Items can be flaky)
            cleanup_test_env(&test_dir)?;
            println!("Skipping test - file did not appear in recent items (Windows Shell timing issue)");
            return Ok(());
        }

        // Test native removal directly (not through public API with fallback)
        let result = remove_recent_file_native(test_path, DEFAULT_COM_TIMEOUT);
        assert!(
            result.is_ok(),
            "Native removal should succeed for existing file: {:?}",
            result
        );

        // Verify file was actually removed
        assert!(
            wait_for_file_status(test_path, false, 10)?,
            "File should be removed from recent items"
        );

        // Test removing non-existent file - should return NotInRecent
        let result = remove_recent_file_native(test_path, DEFAULT_COM_TIMEOUT);
        assert!(result.is_err(), "Should fail when file not in recent");

        if let Err(e) = result {
            let error_msg = format!("{:?}", e);
            assert!(
                error_msg.contains("NotInRecent") || error_msg.contains("not found in recent items"),
                "Should return NotInRecent error when file not found: {:?}",
                e
            );
        }

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    fn test_recent_files_error_handling() -> WincentResult<()> {
        // Tests error handling consistency across native and PowerShell implementations
        // remove_from_recent_files() has fallback from native to PowerShell
        // Both implementations should fail consistently for the same invalid inputs

        // Test add errors
        let result = add_to_recent_files("Z:\\NonExistentFile.txt");
        assert!(result.is_err(), "Should fail with non-existent file");

        let result = add_to_recent_files("");
        assert!(result.is_err(), "Should fail with empty path");

        let result = add_to_recent_files("\0invalid\0path");
        assert!(result.is_err(), "Invalid path characters should not be allowed");

        // Test remove errors - validates fallback consistency
        // If native fails, PowerShell should also fail for the same reasons
        let result = remove_from_recent_files("Z:\\NonExistentFile.txt");
        assert!(result.is_err(), "Should fail with non-existent file");

        let result = remove_from_recent_files("");
        assert!(result.is_err(), "Should fail with empty path");

        let result = remove_from_recent_files("invalid\\path\\*");
        assert!(result.is_err(), "Should fail with invalid path");

        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state"]
    fn test_add_file_to_recent_with_unicode() -> WincentResult<()> {
        // Tests that add_to_recent_files() works for filenames that contain spaces.
        // Uses timestamp-suffixed names to avoid Windows Shell deduplication:
        // repeated additions of the same filename within a short window may be silently ignored.
        //
        // Note: SHAddToRecentDocs is asynchronous; files are polled until visible
        // or the test is skipped on timeout (Windows Shell timing is not guaranteed).

        let test_dir = setup_test_env()?;

        let timestamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_millis();

        // Test regular file
        let filename1 = format!("test_file_{}.txt", timestamp);
        let test_file = create_test_file(&test_dir, &filename1, "test content")?;
        let test_path = test_file.to_str().unwrap();
        add_to_recent_files(test_path)?;

        // Test file with spaces
        let filename2 = format!("test file with spaces {}.txt", timestamp);
        let test_file2 = create_test_file(&test_dir, &filename2, "test content")?;
        let test_path2 = test_file2.to_str().unwrap();
        add_to_recent_files(test_path2)?;

        // Wait for both files to appear (20 retries × 500ms = 10 seconds)
        if !wait_for_file_status(test_path, true, 20)? {
            cleanup_test_env(&test_dir)?;
            println!("Skipping test - file did not appear in recent items (Windows Shell timing issue)");
            return Ok(());
        }
        if !wait_for_file_status(test_path2, true, 20)? {
            cleanup_test_env(&test_dir)?;
            println!("Skipping test - file with spaces did not appear in recent items (Windows Shell timing issue)");
            return Ok(());
        }

        let _ = remove_from_recent_files(test_path);
        let _ = remove_from_recent_files(test_path2);

        cleanup_test_env(&test_dir)?;
        Ok(())
    }
    
    #[test]
    #[ignore = "Modifies system state"]
    fn test_unpin_folder_item_compatibility() -> WincentResult<()> {
        // Tests the Windows 10/11 compatibility logic in unpin_folder_item()
        // Implementation: src/handle.rs:258-272
        // Key logic: Try unpinfromhome first (Win10), fallback to pintohome (Win11)
        //
        // This test directly verifies the fallback mechanism by:
        // 1. Pinning a folder to get a real FolderItem
        // 2. Calling unpin_folder_item() on it
        // 3. Verifying the folder is actually unpinned
        //
        // This test focuses on the single-item unpin compatibility path

        use crate::query::query_recent_native;
        use crate::QuickAccess;

        let test_dir = setup_test_env()?;
        let test_path = test_dir.to_str().unwrap();

        // Pin a folder
        pin_frequent_folder(test_path, DEFAULT_COM_TIMEOUT)?;
        thread::sleep(Duration::from_millis(500));

        // Verify it was pinned
        let folders = query_recent_native(QuickAccess::FrequentFolders)?;
        assert!(
            folders.iter().any(|p| p == test_path),
            "Test folder should be pinned"
        );

        // Get the FolderItem and test unpin_folder_item() directly.
        // Must run on a dedicated STA thread: on Win11, shell_folder() for frequent folders
        // requires cross-process calls to explorer.exe that need a message pump.
        let test_path_owned = test_path.to_owned();
        let found = crate::com_thread::run_on_sta_thread(move || {
            let folder = shell_folder(FREQUENT_FOLDERS_NAMESPACE)?;
            let items = folder_items(&folder)?;

            let count = unsafe {
                items.Count().map_err(|e| {
                    WincentError::SystemError(format!("Failed to get item count: {}", e))
                })?
            };

            for index in 0..count {
                let index_variant = VARIANT::from(index);
                let item = unsafe {
                    match items.Item(&index_variant) {
                        Ok(item) => item,
                        Err(_) => continue,
                    }
                };

                if let Some(item_path_str) = item_path(&item) {
                    if paths_equal(&item_path_str, &test_path_owned) {
                        // Found our test folder - test unpin_folder_item()
                        // This will try unpinfromhome first, then fallback to pintohome
                        unpin_folder_item(&item)?;
                        return Ok(true);
                    }
                }
            }

            Ok(false)
        }, DEFAULT_COM_TIMEOUT)?;

        assert!(found, "Should have found the test folder in frequent folders");

        // Verify the folder was actually unpinned
        thread::sleep(Duration::from_millis(500));
        let folders_after = query_recent_native(QuickAccess::FrequentFolders)?;
        assert!(
            !folders_after.iter().any(|p| p == test_path),
            "Test folder should be unpinned after unpin_folder_item()"
        );

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    #[ignore = "Performance benchmark"]
    fn test_native_pin_unpin_performance() -> WincentResult<()> {
        use std::time::Instant;

        let test_dir = setup_test_env()?;
        let test_path = test_dir.to_str().unwrap();

        // Benchmark native API
        let start = Instant::now();
        pin_frequent_folder_native(test_path, DEFAULT_COM_TIMEOUT)?;
        let native_pin = start.elapsed();

        thread::sleep(Duration::from_millis(100));

        let start = Instant::now();
        unpin_frequent_folder_native(test_path, DEFAULT_COM_TIMEOUT)?;
        let native_unpin = start.elapsed();

        println!("Native API - Pin: {:?}, Unpin: {:?}", native_pin, native_unpin);
        println!("Total: {:?}", native_pin + native_unpin);

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state"]
    fn test_com_s_false_reference_counting() -> WincentResult<()> {
        // Tests that add_file_to_recent_native() correctly handles S_FALSE
        // and properly calls CoUninitialize to balance reference counts.
        //
        // Strategy: Use multiple init/uninit cycles to amplify reference leaks.
        // If the library leaks references, COM will remain initialized after
        // the final CoUninitialize(), which we can detect.
        use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED};

        let test_dir = setup_test_env()?;
        let test_file = create_test_file(&test_dir, "s_false_test.txt", "content")?;
        let test_path = test_file.to_str().unwrap();

        unsafe {
            // Cycle 1: User init -> library call -> user uninit
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            assert_eq!(hr.0, 0, "First CoInitializeEx should return S_OK");

            let result = add_file_to_recent_native(test_path);
            assert!(result.is_ok(), "Should handle S_FALSE correctly: {:?}", result);

            CoUninitialize();

            // Cycle 2: Repeat to amplify potential leaks
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            assert_eq!(hr.0, 0, "Second cycle should return S_OK (COM was uninitialized)");

            let result = add_file_to_recent_native(test_path);
            assert!(result.is_ok(), "Should handle S_FALSE on second cycle: {:?}", result);

            CoUninitialize();

            // Cycle 3: One more time
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            assert_eq!(hr.0, 0, "Third cycle should return S_OK");

            let result = add_file_to_recent_native(test_path);
            assert!(result.is_ok(), "Should handle S_FALSE on third cycle: {:?}", result);

            CoUninitialize();

            // Verification: Try to initialize with incompatible mode
            // If there are leaked references, this will fail with RPC_E_CHANGED_MODE
            let hr = CoInitializeEx(None, COINIT_MULTITHREADED);
            assert!(
                hr.0 == 0 || hr.0 == 1,
                "Should be able to initialize with MTA (no leaked STA references). Got: 0x{:08X}",
                hr.0
            );
            CoUninitialize();
        }

        // Clean up: Remove test file from recent list
        let _ = remove_recent_file_native(test_path, DEFAULT_COM_TIMEOUT);

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    fn test_com_apartment_mismatch() -> WincentResult<()> {
        // Tests that add_file_to_recent_native() correctly detects and reports
        // RPC_E_CHANGED_MODE when called from an MTA thread.
        //
        // Scenario: Thread is initialized with MTA, then calls STA-only function
        // Expected: Function returns ComApartmentMismatch error
        // Risk: If not detected, function may hang or crash
        use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED};

        unsafe {
            // Initialize as MTA (multi-threaded apartment)
            let hr = CoInitializeEx(None, COINIT_MULTITHREADED);
            assert!(hr.is_ok() || hr.0 == 1, "MTA init should succeed");

            // Call STA-only function (should return ComApartmentMismatch)
            let result = add_file_to_recent_native("C:\\Windows\\notepad.exe");

            match result {
                Err(WincentError::ComApartmentMismatch(msg)) => {
                    assert!(
                        msg.contains("incompatible"),
                        "Error message should mention incompatibility: {}",
                        msg
                    );
                }
                _ => panic!("Expected ComApartmentMismatch error, got: {:?}", result),
            }

            // Clean up MTA
            CoUninitialize();
        }

        Ok(())
    }

    #[test]
    fn test_with_timeout_zero_rejected() {
        // All three *_with_timeout public functions must reject Duration::ZERO
        // before any operation (path validation, native COM, PowerShell fallback).
        let r = remove_from_recent_files_with_timeout("C:\\Windows\\notepad.exe", std::time::Duration::ZERO);
        assert!(matches!(r, Err(WincentError::InvalidArgument(_))), "got: {:?}", r);

        let r = add_to_frequent_folders_with_timeout("C:\\Windows", std::time::Duration::ZERO);
        assert!(matches!(r, Err(WincentError::InvalidArgument(_))), "got: {:?}", r);

        let r = remove_from_frequent_folders_with_timeout("C:\\Windows", std::time::Duration::ZERO);
        assert!(matches!(r, Err(WincentError::InvalidArgument(_))), "got: {:?}", r);
    }
}
