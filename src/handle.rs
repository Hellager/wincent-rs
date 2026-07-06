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
//! 1. Dedicated STA threading for Shell COM operations
//! 2. Path validation before execution
//! 3. PowerShell script sandboxing
//! 4. Clean error propagation

use crate::{
    error::WincentError,
    query::{folder_items, item_path, shell_folder, FREQUENT_FOLDERS_NAMESPACE},
    script_executor::ScriptExecutor,
    script_strategy::PSScript,
    utils::{paths_equal, validate_path, PathType},
    QuickAccess, WincentResult,
};
use std::ffi::OsString;
use std::os::windows::prelude::*;
use std::time::{Duration, Instant};
use windows::core::Interface;
use windows::Win32::System::SystemInformation::OSVERSIONINFOW;
use windows::Win32::System::Variant::VARIANT;
use windows::Win32::UI::Shell::{Folder3, SHAddToRecentDocs};

/// Default timeout for COM STA thread operations
const DEFAULT_COM_TIMEOUT: std::time::Duration = std::time::Duration::from_secs(10);
const FREQUENT_FOLDER_VERIFICATION_TIMEOUT: Duration = Duration::from_secs(1);
const FREQUENT_FOLDER_VERIFICATION_POLL_INTERVAL: Duration = Duration::from_millis(100);
/// SHARD_PATHW registers a wide-string path in the recent documents list.
const SHARD_PATHW: u32 = 0x0000_0003;

#[link(name = "ntdll")]
extern "system" {
    fn RtlGetVersion(lpversioninformation: *mut OSVERSIONINFOW) -> i32;
}

fn current_windows_build_number() -> WincentResult<u32> {
    let mut version_info = OSVERSIONINFOW {
        dwOSVersionInfoSize: std::mem::size_of::<OSVERSIONINFOW>() as u32,
        ..Default::default()
    };

    // SAFETY: version_info is initialized with the correct size in `dwOSVersionInfoSize`.
    // RtlGetVersion fills the struct in-place. This is the recommended API for reliable
    // OS version detection that bypasses compatibility shims.
    let status = unsafe { RtlGetVersion(&mut version_info) };
    if status != 0 {
        return Err(WincentError::SystemError(format!(
            "Failed to get Windows version: NTSTATUS 0x{status:08X}"
        )));
    }

    Ok(version_info.dwBuildNumber)
}

fn is_windows_11_or_later() -> WincentResult<bool> {
    Ok(current_windows_build_number()? >= 22_000)
}

fn invoke_verb_on_self_current_sta(path: &str, verb: &str) -> WincentResult<()> {
    let folder = shell_folder(path)?;

    // SAFETY: This runs on a COM-initialized STA thread. `folder` is a live
    // Shell Folder interface for `path`, and the verb VARIANT lives until
    // InvokeVerb returns.
    unsafe {
        let folder3: Folder3 = folder.cast().map_err(|e| {
            WincentError::SystemError(format!("Failed to cast to Folder3 for {}: {}", path, e))
        })?;

        let self_item = folder3.Self_().map_err(|e| {
            WincentError::SystemError(format!("Failed to get Self for {}: {}", path, e))
        })?;

        let verb_variant = VARIANT::from(verb);
        self_item.InvokeVerb(&verb_variant).map_err(|e| {
            WincentError::SystemError(format!(
                "Failed to invoke verb '{}' on {}: {}",
                verb, path, e
            ))
        })?;
    }

    Ok(())
}

fn contains_frequent_folder_current_sta(path: &str) -> WincentResult<bool> {
    let folder = shell_folder(FREQUENT_FOLDERS_NAMESPACE)?;
    let items = folder_items(&folder)?;

    // SAFETY: `items` is a live FolderItems collection on the current STA thread.
    let count = unsafe {
        items
            .Count()
            .map_err(|e| WincentError::SystemError(format!("Failed to get item count: {}", e)))?
    };

    for index in 0..count {
        let index_variant = VARIANT::from(index);
        // SAFETY: The collection is live for the loop and the index VARIANT
        // lives until Item returns. Per-item failures are skipped.
        let item = unsafe {
            match items.Item(&index_variant) {
                Ok(item) => item,
                Err(_) => continue,
            }
        };

        if let Some(item_path_str) = item_path(&item) {
            if paths_equal(&item_path_str, path) {
                return Ok(true);
            }
        }
    }

    Ok(false)
}

fn try_invoke_verb_on_frequent_folder_current_sta(path: &str, verb: &str) -> WincentResult<bool> {
    let folder = shell_folder(FREQUENT_FOLDERS_NAMESPACE)?;
    let items = folder_items(&folder)?;

    // SAFETY: `items` is a live FolderItems collection on the current STA thread.
    let count = unsafe {
        items
            .Count()
            .map_err(|e| WincentError::SystemError(format!("Failed to get item count: {}", e)))?
    };

    for index in 0..count {
        let index_variant = VARIANT::from(index);
        // SAFETY: The collection is live for the loop and the index VARIANT
        // lives until Item returns. Per-item failures are skipped.
        let item = unsafe {
            match items.Item(&index_variant) {
                Ok(item) => item,
                Err(_) => continue,
            }
        };

        if let Some(item_path_str) = item_path(&item) {
            if paths_equal(&item_path_str, path) {
                let verb_variant = VARIANT::from(verb);
                // SAFETY: `item` is a live FolderItem and `verb_variant`
                // remains alive for the synchronous InvokeVerb call.
                unsafe {
                    item.InvokeVerb(&verb_variant).map_err(|e| {
                        WincentError::SystemError(format!(
                            "Failed to invoke verb '{}' on {}: {}",
                            verb, path, e
                        ))
                    })?;
                }
                return Ok(true);
            }
        }
    }

    Ok(false)
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
fn run_pin_frequent_folder_checked<C, I>(
    path: &str,
    mut contains: C,
    mut invoke_on_self: I,
) -> WincentResult<()>
where
    C: FnMut(&str) -> WincentResult<bool>,
    I: FnMut(&str, &str) -> WincentResult<()>,
{
    // On Windows 11 "pintohome" acts as a toggle. Check in the same STA flow
    // as the verb invocation so an already-pinned folder is never toggled off.
    if contains(path)? {
        return Err(WincentError::already_exists(
            path,
            QuickAccess::FrequentFolders,
        ));
    }

    invoke_on_self(path, "pintohome")
}

/// Pins a folder to Quick Access using native COM API
///
/// Uses the "pintohome" Shell verb to pin a folder to the Frequent Folders list.
/// This is the fast path for pinning operations (~10-50ms).
///
/// # Threading Model
///
/// Runs the duplicate check and `pintohome` invocation inside a single
/// dedicated STA thread to avoid Windows 11 deadlocks and reduce the race
/// window between check and invoke.
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
/// - `AlreadyExists`: The folder is already pinned
/// - `SystemError`: COM operation failed (e.g., failed to open folder namespace, invoke verb)
///
/// # See Also
///
/// - [`pin_frequent_folder()`] - Higher-level internal wrapper with PowerShell fallback
/// - [`invoke_verb_on_self_current_sta()`] - The underlying verb invocation mechanism
fn pin_frequent_folder_native(path: &str, timeout: std::time::Duration) -> WincentResult<()> {
    // Check and invoke in one STA worker to narrow the TOCTOU window.
    // On Windows 11 "pintohome" acts as a toggle. The check prevents toggling an
    // already-pinned folder off.
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread(
        move || {
            run_pin_frequent_folder_checked(
                &path,
                contains_frequent_folder_current_sta,
                invoke_verb_on_self_current_sta,
            )
        },
        timeout,
    )
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
#[allow(dead_code)]
fn is_in_frequent_folders_native(path: &str, timeout: std::time::Duration) -> WincentResult<bool> {
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread(
        move || contains_frequent_folder_current_sta(&path),
        timeout,
    )
}

fn wait_for_frequent_folder_presence_with<C, S>(
    path: &str,
    expected: bool,
    verification_timeout: Duration,
    poll_interval: Duration,
    contains: &mut C,
    sleep: &mut S,
) -> WincentResult<bool>
where
    C: FnMut(&str) -> WincentResult<bool>,
    S: FnMut(Duration),
{
    let start = Instant::now();

    while start.elapsed() < verification_timeout {
        if contains(path)? == expected {
            return Ok(true);
        }

        let remaining = verification_timeout.saturating_sub(start.elapsed());
        if remaining.is_zero() {
            break;
        }

        sleep(std::cmp::min(poll_interval, remaining));
    }

    Ok(contains(path)? == expected)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PinnedStatus {
    Pinned,
    Unpinned,
    Unknown,
}

fn frequent_folder_pinned_status(path: &str) -> PinnedStatus {
    frequent_folder_pinned_status_impl(path)
}

fn frequent_folder_pinned_status_impl(path: &str) -> PinnedStatus {
    match frequent_folder_pinned_status_from_destlist(path) {
        Ok(status) => status,
        Err(_) => PinnedStatus::Unknown,
    }
}

fn frequent_folder_pinned_status_from_destlist(path: &str) -> WincentResult<PinnedStatus> {
    let parsed = crate::destlist::parse_file(crate::destlist::frequent_folders_dest_path()?)?;
    Ok(frequent_folder_pinned_status_from_entries(
        path,
        parsed.dest_list().entries(),
    ))
}

fn frequent_folder_pinned_status_from_entries(
    path: &str,
    entries: &[crate::destlist::DestListEntry],
) -> PinnedStatus {
    let mut matched = false;

    for entry in entries
        .iter()
        .filter(|entry| paths_equal(entry.path(), path))
    {
        matched = true;
        if entry.is_pinned() {
            return PinnedStatus::Pinned;
        }
    }

    if matched {
        PinnedStatus::Unpinned
    } else {
        PinnedStatus::Unknown
    }
}

#[allow(clippy::too_many_arguments)]
fn run_unpin_frequent_folder_state_machine<C, N, S, W, Sleep>(
    path: &str,
    pinned_status: PinnedStatus,
    verification_timeout: Duration,
    poll_interval: Duration,
    mut contains: C,
    mut invoke_on_namespace: N,
    mut invoke_on_self: S,
    mut is_win11_or_later: W,
    mut sleep: Sleep,
) -> WincentResult<()>
where
    C: FnMut(&str) -> WincentResult<bool>,
    N: FnMut(&str, &str) -> WincentResult<bool>,
    S: FnMut(&str, &str) -> WincentResult<()>,
    W: FnMut() -> WincentResult<bool>,
    Sleep: FnMut(Duration),
{
    if !contains(path)? {
        return Err(WincentError::not_in_quick_access(
            path,
            QuickAccess::FrequentFolders,
        ));
    }

    if matches!(pinned_status, PinnedStatus::Unpinned) {
        invoke_on_self(path, "pintohome")?;
    }

    let _ = invoke_on_namespace(path, "unpinfromhome");
    if wait_for_frequent_folder_presence_with(
        path,
        false,
        verification_timeout,
        poll_interval,
        &mut contains,
        &mut sleep,
    )? {
        return Ok(());
    }

    invoke_on_self(path, "pintohome")?;
    if wait_for_frequent_folder_presence_with(
        path,
        false,
        verification_timeout,
        poll_interval,
        &mut contains,
        &mut sleep,
    )? {
        return Ok(());
    }

    if is_win11_or_later()? {
        invoke_on_self(path, "pintohome")?;
    } else {
        let _ = invoke_on_namespace(path, "unpinfromhome");
    }

    if wait_for_frequent_folder_presence_with(
        path,
        false,
        verification_timeout,
        poll_interval,
        &mut contains,
        &mut sleep,
    )? {
        return Ok(());
    }

    Err(WincentError::SystemError(format!(
        "Failed to remove frequent folder: {}",
        path
    )))
}

///
/// This function implements a **safe Windows version-aware strategy**:
///
/// 1. **Check if folder is present**
///    - Query the Frequent Folders namespace to verify the folder is visible
///    - If absent, return `NotInQuickAccess` error immediately
///    - If present but unpinned, pin it first so the normal unpin verbs can remove it
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
/// unpinned frequent entries are only pinned intentionally as the first half of
/// a pin-then-unpin removal.
///
/// # Threading Model
///
/// Runs the full check, mutation, verification, and fallback sequence inside a
/// single dedicated STA thread to avoid repeated COM thread dispatches.
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
/// - `NotInQuickAccess`: The folder is not in Frequent Folders
/// - `SystemError`: COM operation failed (e.g., permission denied, COM error)
///
/// # Windows Version Compatibility
///
/// - **Windows 10**: Uses "unpinfromhome" verb (step 2 succeeds)
/// - **Windows 11**: Uses "pintohome" toggle (step 2 fails, step 3 succeeds)
///
/// # See Also
///
/// - [`unpin_frequent_folder()`] - Higher-level internal wrapper with PowerShell fallback
fn unpin_frequent_folder_native(path: &str, timeout: std::time::Duration) -> WincentResult<()> {
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread(
        move || {
            run_unpin_frequent_folder_state_machine(
                &path,
                frequent_folder_pinned_status(&path),
                FREQUENT_FOLDER_VERIFICATION_TIMEOUT,
                FREQUENT_FOLDER_VERIFICATION_POLL_INTERVAL,
                contains_frequent_folder_current_sta,
                try_invoke_verb_on_frequent_folder_current_sta,
                invoke_verb_on_self_current_sta,
                is_windows_11_or_later,
                std::thread::sleep,
            )
        },
        timeout,
    )
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
/// - Paths are validated before script generation
/// - Path parameters are embedded only after PowerShell single-quote escaping
/// - Scripts are executed with `Command::args`, avoiding shell command-line interpolation
/// - Generated scripts are stored under wincent's temp-script directory and refreshed when stale
///
/// # See Also
///
/// - [`crate::script_executor::ScriptExecutor`] - PowerShell script executor
/// - [`crate::script_storage::ScriptStorage`] - Script storage and retrieval
#[allow(dead_code)]
pub(crate) fn execute_script_with_validation(
    script: PSScript,
    path: &str,
    path_type: PathType,
) -> WincentResult<()> {
    execute_script_with_validation_and_timeout(
        script,
        path,
        path_type,
        ScriptExecutor::execute_ps_script,
    )
}

fn execute_script_with_validation_and_timeout<F>(
    script: PSScript,
    path: &str,
    path_type: PathType,
    mut execute: F,
) -> WincentResult<()>
where
    F: FnMut(PSScript, Option<&str>) -> WincentResult<std::process::Output>,
{
    validate_path(path, path_type)?;

    let start = std::time::Instant::now();
    let script_path = match path_type {
        PathType::File | PathType::Directory => {
            crate::script_storage::ScriptStorage::get_dynamic_script_path(script, path)?
        }
    };
    let output = execute(script, Some(path))?;
    let duration = start.elapsed();

    parse_script_execution_result(output, script, script_path, path, duration)
}

fn parse_script_execution_result(
    output: std::process::Output,
    script: PSScript,
    script_path: std::path::PathBuf,
    path: &str,
    duration: Duration,
) -> WincentResult<()> {
    if output.status.success() {
        return Ok(());
    }

    use crate::error::PowerShellError;
    let stderr = String::from_utf8(output.stderr)
        .unwrap_or_else(|_| "Unable to parse script error output".to_string());
    let stdout = String::from_utf8(output.stdout).unwrap_or_default();

    // Infer error kind from stderr content
    let kind = PowerShellError::infer_kind_from_stderr(&stderr);

    Err(WincentError::PowerShellExecution(Box::new(
        PowerShellError::builder(script.operation())
            .kind(kind)
            .exit_code(output.status.code())
            .stdout(stdout)
            .stderr(stderr)
            .script_path(script_path)
            .parameters(path)
            .duration(duration)
            .build(),
    )))
}

/// Inner: runs on the current STA thread. Only call from within `run_on_sta_thread`.
fn add_file_to_recent_native_current_sta(path: &str) -> WincentResult<()> {
    // SAFETY: `path` was validated before reaching here; the wide string is
    // null-terminated and its lifetime extends past the SHAddToRecentDocs call.
    // SHARD_PATHW instructs the Shell to accept a wide-string path pointer.
    unsafe {
        let file_path_wide: Vec<u16> = OsString::from(path)
            .encode_wide()
            .chain(std::iter::once(0))
            .collect();
        SHAddToRecentDocs(SHARD_PATHW, Some(file_path_wide.as_ptr() as *const _));
    }
    Ok(())
}

/// Adds a file to the Windows Recent Items list via a dedicated STA thread.
///
/// `timeout` limits how long the caller waits; the underlying Shell operation
/// may still complete after the timeout elapses.
///
/// # Windows Behavior Notes
///
/// - **Deduplication**: Windows may ignore repeated additions of the same file
///   within a short time window.
/// - **Asynchronous Processing**: The file may not appear immediately in the
///   Recent Items list. Windows Shell processes these updates asynchronously.
pub(crate) fn add_file_to_recent_native(path: &str, timeout: Duration) -> WincentResult<()> {
    validate_path(path, PathType::File)?;
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread(
        move || add_file_to_recent_native_current_sta(&path),
        timeout,
    )
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
/// - `NotInQuickAccess`: The file was not found in the Recent Items namespace
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
/// - [`crate::manager::QuickAccessManager::remove_item`] - Public removal API
fn remove_recent_file_native(path: &str, timeout: std::time::Duration) -> WincentResult<()> {
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread(
        move || {
            let recent_namespace = "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}";
            let folder = shell_folder(recent_namespace)?;
            let items = folder_items(&folder)?;

            // SAFETY: `items` is a live FolderItems collection on the STA worker.
            let count = unsafe {
                items.Count().map_err(|e| {
                    WincentError::SystemError(format!("Failed to get item count: {}", e))
                })?
            };

            // Search for the target file
            for index in 0..count {
                let index_variant = VARIANT::from(index);
                // SAFETY: `items` is live for this loop and the VARIANT index
                // lives until Item returns. Per-item failures are skipped.
                let item = unsafe {
                    match items.Item(&index_variant) {
                        Ok(item) => item,
                        Err(_) => continue,
                    }
                };

                // Check if this is a file (not folder)
                // SAFETY: `item` is a live FolderItem returned by the Shell collection.
                let is_folder = unsafe { item.IsFolder().map(bool::from).unwrap_or(true) };
                if is_folder {
                    continue;
                }

                if let Some(item_path_str) = item_path(&item) {
                    if paths_equal(&item_path_str, &path) {
                        // Found the file, invoke remove verb
                        let verb_variant = VARIANT::from("remove");
                        // SAFETY: `item` is a live FolderItem and the verb
                        // VARIANT remains alive until InvokeVerb returns.
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

            Err(WincentError::not_in_quick_access(
                path,
                QuickAccess::RecentFiles,
            ))
        },
        timeout,
    )
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
/// - [`crate::manager::QuickAccessManager::remove_item`] - Public removal API
#[allow(dead_code)]
pub(crate) fn remove_recent_file_powershell(path: &str) -> WincentResult<()> {
    remove_recent_file_powershell_with_timeout(path, DEFAULT_COM_TIMEOUT)
}

fn remove_recent_file_powershell_with_timeout(path: &str, timeout: Duration) -> WincentResult<()> {
    match execute_remove_recent_file_script_with_timeout(
        path,
        timeout,
        |script, parameter, timeout| {
            ScriptExecutor::execute_ps_script_with_timeout(script, parameter, timeout)
        },
    ) {
        Err(error) => Err(map_remove_recent_file_powershell_error(path, error)),
        ok => ok,
    }
}

fn execute_remove_recent_file_script_with_timeout<F>(
    path: &str,
    timeout: Duration,
    mut execute: F,
) -> WincentResult<()>
where
    F: FnMut(PSScript, Option<&str>, Duration) -> WincentResult<std::process::Output>,
{
    execute_script_with_validation_and_timeout(
        PSScript::RemoveRecentFile,
        path,
        PathType::File,
        |script, parameter| execute(script, parameter, timeout),
    )
}

fn map_remove_recent_file_powershell_error(path: &str, error: WincentError) -> WincentError {
    const NOT_IN_QUICK_ACCESS_SENTINEL: &str = "WINCENT_NOT_IN_QUICK_ACCESS";

    match error {
        WincentError::PowerShellExecution(ref powershell_error)
            if powershell_error
                .raw_stdout()
                .contains(NOT_IN_QUICK_ACCESS_SENTINEL) =>
        {
            WincentError::not_in_quick_access(path, QuickAccess::RecentFiles)
        }
        other => other,
    }
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
#[allow(dead_code)]
fn pin_frequent_folder_powershell(path: &str) -> WincentResult<()> {
    pin_frequent_folder_powershell_with_timeout(path, DEFAULT_COM_TIMEOUT)
}

fn pin_frequent_folder_powershell_with_timeout(path: &str, timeout: Duration) -> WincentResult<()> {
    match execute_pin_frequent_folder_script_with_timeout(
        path,
        timeout,
        |script, parameter, timeout| {
            ScriptExecutor::execute_ps_script_with_timeout(script, parameter, timeout)
        },
    ) {
        Ok(()) => Ok(()),
        Err(error) => Err(map_pin_frequent_folder_powershell_error(path, error)),
    }
}

fn execute_pin_frequent_folder_script_with_timeout<F>(
    path: &str,
    timeout: Duration,
    mut execute: F,
) -> WincentResult<()>
where
    F: FnMut(PSScript, Option<&str>, Duration) -> WincentResult<std::process::Output>,
{
    execute_script_with_validation_and_timeout(
        PSScript::PinToFrequentFolder,
        path,
        PathType::Directory,
        |script, parameter| execute(script, parameter, timeout),
    )
}

fn pin_frequent_folder_with_fallback<N, P>(mut native: N, mut powershell: P) -> WincentResult<()>
where
    N: FnMut() -> WincentResult<()>,
    P: FnMut() -> WincentResult<()>,
{
    match native() {
        Ok(()) => Ok(()),
        Err(e @ WincentError::AlreadyExists { .. }) => Err(e),
        // A timed-out STA worker keeps running and may still complete the
        // `pintohome` call later. On Windows 11 that verb can behave like a
        // toggle, so running the fallback could undo a successful late pin.
        Err(e @ WincentError::Timeout(_)) => Err(e),
        Err(_) => powershell(),
    }
}

fn map_pin_frequent_folder_powershell_error(path: &str, error: WincentError) -> WincentError {
    const ALREADY_EXISTS_SENTINEL: &str = "WINCENT_ALREADY_EXISTS";

    match error {
        WincentError::PowerShellExecution(ref powershell_error)
            if powershell_error
                .raw_stdout()
                .contains(ALREADY_EXISTS_SENTINEL) =>
        {
            WincentError::already_exists(path, QuickAccess::FrequentFolders)
        }
        other => other,
    }
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
/// - `PowerShellExecution`: Script execution failed (e.g., folder could not be removed)
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
fn unpin_frequent_folder_powershell(path: &str, timeout: std::time::Duration) -> WincentResult<()> {
    match execute_unpin_frequent_folder_script_with_timeout(
        path,
        timeout,
        |script, parameter, timeout| {
            ScriptExecutor::execute_ps_script_with_timeout(script, parameter, timeout)
        },
    ) {
        Err(error) => Err(map_unpin_frequent_folder_powershell_error(path, error)),
        ok => ok,
    }
}

fn execute_unpin_frequent_folder_script_with_timeout<F>(
    path: &str,
    timeout: Duration,
    mut execute: F,
) -> WincentResult<()>
where
    F: FnMut(PSScript, Option<&str>, Duration) -> WincentResult<std::process::Output>,
{
    execute_script_with_validation_and_timeout(
        PSScript::UnpinFromFrequentFolder,
        path,
        PathType::Directory,
        |script, parameter| execute(script, parameter, timeout),
    )
}

fn map_unpin_frequent_folder_powershell_error(path: &str, error: WincentError) -> WincentError {
    const NOT_IN_QUICK_ACCESS_SENTINEL: &str = "WINCENT_NOT_IN_QUICK_ACCESS";

    match error {
        WincentError::PowerShellExecution(ref powershell_error)
            if powershell_error
                .raw_stdout()
                .contains(NOT_IN_QUICK_ACCESS_SENTINEL) =>
        {
            WincentError::not_in_quick_access(path, QuickAccess::FrequentFolders)
        }
        other => other,
    }
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
/// - [`crate::manager::QuickAccessManager::add_item`] - Public mutation API
/// - [`pin_frequent_folder_native()`] - Native COM implementation
/// - [`pin_frequent_folder_powershell()`] - PowerShell fallback
pub(crate) fn pin_frequent_folder(path: &str, timeout: std::time::Duration) -> WincentResult<()> {
    validate_path(path, PathType::Directory)?;

    // Try native COM first (fast path), fallback to PowerShell only for
    // system/COM failures. AlreadyExists is an authoritative duplicate result
    // from the native check+pin path and must not be retried via PowerShell.
    pin_frequent_folder_with_fallback(
        || pin_frequent_folder_native(path, timeout),
        || pin_frequent_folder_powershell_with_timeout(path, timeout),
    )
}

/// Removes a folder from the Windows Quick Access Frequent Folders list
///
/// This is the internal implementation that uses a **two-tier fallback strategy**:
///
/// 1. **Native COM API** (fast path, 10-50ms): Uses [`unpin_frequent_folder_native()`]
///    - Windows 10: Uses "unpinfromhome" verb
///    - Windows 11: Uses "pintohome" toggle
///    - Unpinned frequent entries are first pinned, then unpinned
///    - Runs on dedicated STA thread to avoid Windows 11 deadlock
///
/// 2. **PowerShell fallback** (slow path, 200-500ms): Uses [`unpin_frequent_folder_powershell()`]
///    - Executes PowerShell script if native COM fails
///    - Provides broader compatibility at the cost of performance
///
/// # Arguments
///
/// * `path` - The full path to the folder to be removed. Must be an existing directory.
///
/// # Returns
///
/// Returns `Ok(())` if the folder was successfully removed (by either strategy).
///
/// # Errors
///
/// - `InvalidPath`: Path validation failed (not a directory, doesn't exist)
/// - `NotInQuickAccess`: Folder not in Frequent Folders (returned immediately, no PowerShell fallback)
/// - `SystemError` or `PowerShellExecution`: Both native COM and PowerShell strategies failed
///
/// Note: When removing an unpinned frequent entry, the native path first pins it
/// and then unpins it. If the second step fails, the original error is returned
/// and the folder may remain pinned. `InvalidPath` errors are returned directly
/// without fallback, as the path is definitively wrong.
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
/// - [`crate::manager::QuickAccessManager::remove_item`] - Public mutation API
/// - [`unpin_frequent_folder_native()`] - Native COM implementation
/// - [`unpin_frequent_folder_powershell()`] - PowerShell fallback
pub(crate) fn unpin_frequent_folder(path: &str, timeout: std::time::Duration) -> WincentResult<()> {
    validate_path(path, PathType::Directory)?;

    // Try native COM first (fast path)
    match unpin_frequent_folder_native(path, timeout) {
        Ok(()) => Ok(()),
        // NotInQuickAccess means the folder is absent - this is a semantic error,
        // not a system failure. Falling back to PowerShell would not help and
        // could mask the real cause from the caller.
        // InvalidPath should also not fallback as the path is definitively wrong.
        Err(e @ WincentError::NotInQuickAccess { .. }) => Err(e),
        Err(e @ WincentError::InvalidPath(_)) => Err(e),
        // For system/COM errors, fallback to PowerShell for broader compatibility
        Err(_) => unpin_frequent_folder_powershell(path, timeout),
    }
}

/****************************************************** Handle Quick Access ******************************************************/

/// Test-only wrapper for removing a file from Windows Recent Files.
///
/// Production callers should use [`crate::manager::QuickAccessManager::remove_item`].
#[cfg(test)]
pub(crate) fn remove_from_recent_files(path: &str) -> WincentResult<()> {
    validate_path(path, PathType::File)?;

    // Try native COM first (fast path), fallback to PowerShell if it fails
    match remove_recent_file_native(path, DEFAULT_COM_TIMEOUT) {
        Ok(()) => Ok(()),
        Err(e @ WincentError::NotInQuickAccess { .. }) => Err(e),
        Err(e @ WincentError::InvalidPath(_)) => Err(e),
        Err(_) => remove_recent_file_powershell(path),
    }
}

/// Removes a file from Windows Recent Files with a custom COM STA thread timeout.
///
/// Internal timeout-aware variant used by [`crate::manager::QuickAccessManager::remove_item`].
/// It tries native COM first and falls back to PowerShell only for system-level
/// native failures.
///
/// # Arguments
///
/// * `path` - The full path to the file to be removed
/// * `timeout` - Timeout for the COM STA thread operation. Must be non-zero;
///   passing [`std::time::Duration::ZERO`] returns
///   [`WincentError::InvalidArgument`] immediately without attempting any operation.
pub(crate) fn remove_from_recent_files_with_timeout(
    path: &str,
    timeout: std::time::Duration,
) -> WincentResult<()> {
    if timeout.is_zero() {
        return Err(WincentError::InvalidArgument(
            "timeout must be greater than zero".to_string(),
        ));
    }
    validate_path(path, PathType::File)?;
    match remove_recent_file_native(path, timeout) {
        Ok(()) => Ok(()),
        Err(e @ WincentError::NotInQuickAccess { .. }) => Err(e),
        Err(e @ WincentError::InvalidPath(_)) => Err(e),
        Err(_) => remove_recent_file_powershell_with_timeout(path, timeout),
    }
}

/// Pins a folder to Windows Quick Access with a custom COM STA thread timeout.
///
/// Internal timeout-aware variant used by [`crate::manager::QuickAccessManager::add_item`].
/// It uses native COM with a PowerShell fallback and rejects zero timeouts before
/// attempting any mutation.
///
/// # Arguments
///
/// * `path` - The full path to the folder to be pinned. Must be an existing directory.
/// * `timeout` - Timeout for the COM STA thread operation. Must be non-zero;
///   passing [`std::time::Duration::ZERO`] returns
///   [`WincentError::InvalidArgument`] immediately without attempting any operation.
pub(crate) fn add_to_frequent_folders_with_timeout(
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
/// Internal timeout-aware variant used by [`crate::manager::QuickAccessManager::remove_item`].
/// It uses native COM with a PowerShell fallback and rejects zero timeouts before
/// attempting any mutation.
///
/// # Arguments
///
/// * `path` - The full path to the folder to be unpinned. Must be an existing directory.
/// * `timeout` - Timeout for the COM STA thread operation. Must be non-zero;
///   passing [`std::time::Duration::ZERO`] returns
///   [`WincentError::InvalidArgument`] immediately without attempting any operation.
pub(crate) fn remove_from_frequent_folders_with_timeout(
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

    fn path_to_str(path: &std::path::Path) -> WincentResult<&str> {
        path.to_str()
            .ok_or_else(|| WincentError::invalid_path(path, "Invalid path encoding"))
    }

    fn wait_for_folder_status(
        path: &str,
        should_exist: bool,
        max_retries: u32,
    ) -> WincentResult<bool> {
        for _ in 0..max_retries {
            let frequent_folders = query_recent(crate::QuickAccess::FrequentFolders)?;
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
        // SAFETY: Test helper receives a live FolderItem from a COM STA worker;
        // the verb VARIANT remains alive for the InvokeVerb call.
        match unsafe { item.InvokeVerb(&verb_variant) } {
            Ok(()) => Ok(()),
            Err(_) => {
                let verb_variant = VARIANT::from("pintohome");
                // SAFETY: Same live FolderItem as above, with the Windows 11
                // fallback verb VARIANT alive until InvokeVerb returns.
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
    fn test_path_normalization_stages() -> WincentResult<()> {
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
        let temp_str = path_to_str(&temp_dir)?;

        // Test with different representations of the same path
        assert!(paths_equal(temp_str, temp_str), "Same path should be equal");

        // Test that different paths are not equal
        assert!(
            !paths_equal("C:\\Windows", "C:\\Users"),
            "Different paths should not be equal"
        );

        Ok(())
    }

    #[test]
    fn test_unpin_state_machine_rejects_initial_absence() {
        let actions = std::rc::Rc::new(std::cell::RefCell::new(Vec::<String>::new()));
        let action_log = std::rc::Rc::clone(&actions);

        let result = run_unpin_frequent_folder_state_machine(
            "C:\\Folder",
            PinnedStatus::Unknown,
            Duration::from_nanos(1),
            Duration::from_nanos(1),
            |_| Ok(false),
            move |_, verb| {
                action_log.borrow_mut().push(format!("namespace:{verb}"));
                Ok(true)
            },
            {
                let actions = std::rc::Rc::clone(&actions);
                move |_, verb| {
                    actions.borrow_mut().push(format!("self:{verb}"));
                    Ok(())
                }
            },
            || Ok(true),
            |_| {},
        );

        assert!(
            matches!(result, Err(WincentError::NotInQuickAccess { .. })),
            "initial absence should be NotInQuickAccess, got: {:?}",
            result
        );
        assert!(
            actions.borrow().is_empty(),
            "no verbs should be invoked when the folder is absent"
        );
    }

    #[test]
    fn test_unpin_state_machine_removes_unpinned_present_by_pin_then_unpin() -> WincentResult<()> {
        let present = std::rc::Rc::new(std::cell::RefCell::new(true));
        let actions = std::rc::Rc::new(std::cell::RefCell::new(Vec::<String>::new()));
        let self_calls = std::rc::Rc::new(std::cell::RefCell::new(0usize));

        run_unpin_frequent_folder_state_machine(
            "C:\\Folder",
            PinnedStatus::Unpinned,
            Duration::from_nanos(1),
            Duration::from_nanos(1),
            {
                let present = std::rc::Rc::clone(&present);
                move |_| Ok(*present.borrow())
            },
            {
                let actions = std::rc::Rc::clone(&actions);
                move |_, verb| {
                    actions.borrow_mut().push(format!("namespace:{verb}"));
                    Ok(true)
                }
            },
            {
                let actions = std::rc::Rc::clone(&actions);
                let present = std::rc::Rc::clone(&present);
                let self_calls = std::rc::Rc::clone(&self_calls);
                move |_, verb| {
                    actions.borrow_mut().push(format!("self:{verb}"));
                    let mut calls = self_calls.borrow_mut();
                    *calls += 1;
                    if *calls == 2 {
                        *present.borrow_mut() = false;
                    }
                    Ok(())
                }
            },
            || Ok(true),
            |_| {},
        )?;

        assert_eq!(
            actions.borrow().as_slice(),
            [
                "self:pintohome",
                "namespace:unpinfromhome",
                "self:pintohome"
            ]
        );
        Ok(())
    }

    #[test]
    fn test_unpin_state_machine_returns_unpinned_pin_error() {
        let actions = std::rc::Rc::new(std::cell::RefCell::new(Vec::<String>::new()));

        let result = run_unpin_frequent_folder_state_machine(
            "C:\\Folder",
            PinnedStatus::Unpinned,
            Duration::from_nanos(1),
            Duration::from_nanos(1),
            |_| Ok(true),
            {
                let actions = std::rc::Rc::clone(&actions);
                move |_, verb| {
                    actions.borrow_mut().push(format!("namespace:{verb}"));
                    Ok(true)
                }
            },
            {
                let actions = std::rc::Rc::clone(&actions);
                move |_, verb| {
                    actions.borrow_mut().push(format!("self:{verb}"));
                    Err(WincentError::SystemError("pin failed".to_string()))
                }
            },
            || Ok(true),
            |_| {},
        );

        assert!(
            matches!(result, Err(WincentError::SystemError(ref message)) if message == "pin failed"),
            "pin failure should be returned directly, got: {:?}",
            result
        );
        assert_eq!(actions.borrow().as_slice(), ["self:pintohome"]);
    }

    #[test]
    fn test_unpin_state_machine_returns_unpin_error_after_unpinned_pin() {
        let actions = std::rc::Rc::new(std::cell::RefCell::new(Vec::<String>::new()));
        let self_calls = std::rc::Rc::new(std::cell::RefCell::new(0usize));

        let result = run_unpin_frequent_folder_state_machine(
            "C:\\Folder",
            PinnedStatus::Unpinned,
            Duration::from_nanos(1),
            Duration::from_nanos(1),
            |_| Ok(true),
            {
                let actions = std::rc::Rc::clone(&actions);
                move |_, verb| {
                    actions.borrow_mut().push(format!("namespace:{verb}"));
                    Ok(true)
                }
            },
            {
                let actions = std::rc::Rc::clone(&actions);
                let self_calls = std::rc::Rc::clone(&self_calls);
                move |_, verb| {
                    actions.borrow_mut().push(format!("self:{verb}"));
                    let mut calls = self_calls.borrow_mut();
                    *calls += 1;
                    if *calls == 1 {
                        Ok(())
                    } else {
                        Err(WincentError::SystemError("unpin failed".to_string()))
                    }
                }
            },
            || Ok(true),
            |_| {},
        );

        assert!(
            matches!(result, Err(WincentError::SystemError(ref message)) if message == "unpin failed"),
            "unpin failure should be returned directly, got: {:?}",
            result
        );
        assert_eq!(
            actions.borrow().as_slice(),
            [
                "self:pintohome",
                "namespace:unpinfromhome",
                "self:pintohome"
            ]
        );
    }

    #[test]
    fn test_pin_checked_rejects_existing_without_invoking() {
        let invoked = std::rc::Rc::new(std::cell::RefCell::new(false));
        let invoked_for_closure = std::rc::Rc::clone(&invoked);

        let result = run_pin_frequent_folder_checked(
            "C:\\Folder",
            |_| Ok(true),
            move |_, _| {
                *invoked_for_closure.borrow_mut() = true;
                Ok(())
            },
        );

        assert!(
            matches!(result, Err(WincentError::AlreadyExists { .. })),
            "existing folder should be AlreadyExists, got: {:?}",
            result
        );
        assert!(!*invoked.borrow(), "pintohome must not be invoked");
    }

    #[test]
    fn test_pin_checked_invokes_pintohome_when_absent() -> WincentResult<()> {
        let actions = std::rc::Rc::new(std::cell::RefCell::new(Vec::<String>::new()));
        let actions_for_closure = std::rc::Rc::clone(&actions);

        run_pin_frequent_folder_checked(
            "C:\\Folder",
            |_| Ok(false),
            move |path, verb| {
                actions_for_closure
                    .borrow_mut()
                    .push(format!("{path}:{verb}"));
                Ok(())
            },
        )?;

        assert_eq!(actions.borrow().as_slice(), ["C:\\Folder:pintohome"]);
        Ok(())
    }

    #[test]
    fn test_pin_checked_propagates_contains_error_without_invoking() {
        let invoked = std::rc::Rc::new(std::cell::RefCell::new(false));
        let invoked_for_closure = std::rc::Rc::clone(&invoked);

        let result = run_pin_frequent_folder_checked(
            "C:\\Folder",
            |_| Err(WincentError::SystemError("contains failed".to_string())),
            move |_, _| {
                *invoked_for_closure.borrow_mut() = true;
                Ok(())
            },
        );

        assert!(
            matches!(result, Err(WincentError::SystemError(ref message)) if message == "contains failed"),
            "contains error should propagate, got: {:?}",
            result
        );
        assert!(!*invoked.borrow(), "pintohome must not be invoked");
    }

    #[test]
    fn test_pin_fallback_preserves_already_exists_without_powershell() {
        let fallback_called = std::rc::Rc::new(std::cell::RefCell::new(false));
        let fallback_called_for_closure = std::rc::Rc::clone(&fallback_called);

        let result = pin_frequent_folder_with_fallback(
            || {
                Err(WincentError::already_exists(
                    "C:\\Folder",
                    QuickAccess::FrequentFolders,
                ))
            },
            move || {
                *fallback_called_for_closure.borrow_mut() = true;
                Ok(())
            },
        );

        assert!(matches!(result, Err(WincentError::AlreadyExists { .. })));
        assert!(
            !*fallback_called.borrow(),
            "PowerShell fallback must not run for AlreadyExists"
        );
    }

    #[test]
    fn test_pin_fallback_runs_powershell_for_system_error() -> WincentResult<()> {
        let fallback_called = std::rc::Rc::new(std::cell::RefCell::new(false));
        let fallback_called_for_closure = std::rc::Rc::clone(&fallback_called);

        pin_frequent_folder_with_fallback(
            || Err(WincentError::SystemError("native failed".to_string())),
            move || {
                *fallback_called_for_closure.borrow_mut() = true;
                Ok(())
            },
        )?;

        assert!(*fallback_called.borrow());
        Ok(())
    }

    #[test]
    fn test_pin_fallback_preserves_timeout_without_powershell() {
        let fallback_called = std::rc::Rc::new(std::cell::RefCell::new(false));
        let fallback_called_for_closure = std::rc::Rc::clone(&fallback_called);

        let result = pin_frequent_folder_with_fallback(
            || Err(WincentError::Timeout("native timed out".to_string())),
            move || {
                *fallback_called_for_closure.borrow_mut() = true;
                Ok(())
            },
        );

        assert!(matches!(
            result,
            Err(WincentError::Timeout(ref message)) if message == "native timed out"
        ));
        assert!(
            !*fallback_called.borrow(),
            "PowerShell fallback must not run after native timeout"
        );
    }

    #[test]
    fn test_unpin_state_machine_does_not_trust_unpinfromhome_success() -> WincentResult<()> {
        let present = std::rc::Rc::new(std::cell::RefCell::new(true));
        let actions = std::rc::Rc::new(std::cell::RefCell::new(Vec::<String>::new()));

        run_unpin_frequent_folder_state_machine(
            "C:\\Folder",
            PinnedStatus::Unknown,
            Duration::from_nanos(1),
            Duration::from_nanos(1),
            {
                let present = std::rc::Rc::clone(&present);
                move |_| Ok(*present.borrow())
            },
            {
                let actions = std::rc::Rc::clone(&actions);
                move |_, verb| {
                    actions.borrow_mut().push(format!("namespace:{verb}"));
                    Ok(true)
                }
            },
            {
                let present = std::rc::Rc::clone(&present);
                let actions = std::rc::Rc::clone(&actions);
                move |_, verb| {
                    actions.borrow_mut().push(format!("self:{verb}"));
                    *present.borrow_mut() = false;
                    Ok(())
                }
            },
            || Ok(true),
            |_| {},
        )?;

        assert_eq!(
            actions.borrow().as_slice(),
            ["namespace:unpinfromhome", "self:pintohome"]
        );
        Ok(())
    }

    #[test]
    fn test_unpin_state_machine_pinned_uses_existing_sequence() -> WincentResult<()> {
        let present = std::rc::Rc::new(std::cell::RefCell::new(true));
        let actions = std::rc::Rc::new(std::cell::RefCell::new(Vec::<String>::new()));

        run_unpin_frequent_folder_state_machine(
            "C:\\Folder",
            PinnedStatus::Pinned,
            Duration::from_nanos(1),
            Duration::from_nanos(1),
            {
                let present = std::rc::Rc::clone(&present);
                move |_| Ok(*present.borrow())
            },
            {
                let actions = std::rc::Rc::clone(&actions);
                move |_, verb| {
                    actions.borrow_mut().push(format!("namespace:{verb}"));
                    Ok(true)
                }
            },
            {
                let present = std::rc::Rc::clone(&present);
                let actions = std::rc::Rc::clone(&actions);
                move |_, verb| {
                    actions.borrow_mut().push(format!("self:{verb}"));
                    *present.borrow_mut() = false;
                    Ok(())
                }
            },
            || Ok(true),
            |_| {},
        )?;

        assert_eq!(
            actions.borrow().as_slice(),
            ["namespace:unpinfromhome", "self:pintohome"]
        );
        Ok(())
    }

    #[test]
    fn test_unpin_state_machine_uses_win11_final_pintohome() -> WincentResult<()> {
        let present = std::rc::Rc::new(std::cell::RefCell::new(true));
        let actions = std::rc::Rc::new(std::cell::RefCell::new(Vec::<String>::new()));
        let self_calls = std::rc::Rc::new(std::cell::RefCell::new(0usize));

        run_unpin_frequent_folder_state_machine(
            "C:\\Folder",
            PinnedStatus::Unknown,
            Duration::from_nanos(1),
            Duration::from_nanos(1),
            {
                let present = std::rc::Rc::clone(&present);
                move |_| Ok(*present.borrow())
            },
            {
                let actions = std::rc::Rc::clone(&actions);
                move |_, verb| {
                    actions.borrow_mut().push(format!("namespace:{verb}"));
                    Ok(true)
                }
            },
            {
                let present = std::rc::Rc::clone(&present);
                let actions = std::rc::Rc::clone(&actions);
                let self_calls = std::rc::Rc::clone(&self_calls);
                move |_, verb| {
                    actions.borrow_mut().push(format!("self:{verb}"));
                    let mut calls = self_calls.borrow_mut();
                    *calls += 1;
                    if *calls == 2 {
                        *present.borrow_mut() = false;
                    }
                    Ok(())
                }
            },
            || Ok(true),
            |_| {},
        )?;

        assert_eq!(
            actions.borrow().as_slice(),
            [
                "namespace:unpinfromhome",
                "self:pintohome",
                "self:pintohome"
            ]
        );
        Ok(())
    }

    #[test]
    fn test_unpin_state_machine_uses_win10_final_unpinfromhome() -> WincentResult<()> {
        let present = std::rc::Rc::new(std::cell::RefCell::new(true));
        let actions = std::rc::Rc::new(std::cell::RefCell::new(Vec::<String>::new()));
        let namespace_calls = std::rc::Rc::new(std::cell::RefCell::new(0usize));

        run_unpin_frequent_folder_state_machine(
            "C:\\Folder",
            PinnedStatus::Unknown,
            Duration::from_nanos(1),
            Duration::from_nanos(1),
            {
                let present = std::rc::Rc::clone(&present);
                move |_| Ok(*present.borrow())
            },
            {
                let present = std::rc::Rc::clone(&present);
                let actions = std::rc::Rc::clone(&actions);
                let namespace_calls = std::rc::Rc::clone(&namespace_calls);
                move |_, verb| {
                    actions.borrow_mut().push(format!("namespace:{verb}"));
                    let mut calls = namespace_calls.borrow_mut();
                    *calls += 1;
                    if *calls == 2 {
                        *present.borrow_mut() = false;
                    }
                    Ok(true)
                }
            },
            {
                let actions = std::rc::Rc::clone(&actions);
                move |_, verb| {
                    actions.borrow_mut().push(format!("self:{verb}"));
                    Ok(())
                }
            },
            || Ok(false),
            |_| {},
        )?;

        assert_eq!(
            actions.borrow().as_slice(),
            [
                "namespace:unpinfromhome",
                "self:pintohome",
                "namespace:unpinfromhome"
            ]
        );
        Ok(())
    }

    #[test]
    fn test_unpin_state_machine_errors_when_folder_remains_present() {
        let result = run_unpin_frequent_folder_state_machine(
            "C:\\Folder",
            PinnedStatus::Unknown,
            Duration::from_nanos(1),
            Duration::from_nanos(1),
            |_| Ok(true),
            |_, _| Ok(true),
            |_, _| Ok(()),
            || Ok(true),
            |_| {},
        );

        assert!(
            matches!(result, Err(WincentError::SystemError(ref message)) if message.contains("Failed to remove frequent folder")),
            "persistent presence should be SystemError, got: {:?}",
            result
        );
    }

    fn destlist_entry_for_test(path: &str, pin_status: i32) -> crate::destlist::DestListEntry {
        crate::destlist::DestListEntry {
            entry_offset: 0,
            entry_len: 0,
            mru_position: 0,
            checksum: 0,
            entry_id: 1,
            entry_number: 1,
            entry_number_unknown: 0,
            hostname: String::new(),
            volume_droid: String::new(),
            file_droid: String::new(),
            volume_birth_droid: String::new(),
            file_birth_droid: String::new(),
            file_droid_mac: String::new(),
            stream_name: "1".to_string(),
            raw_path: path.to_string(),
            path: path.to_string(),
            pin_status,
            pin_order: (pin_status >= 0).then_some(pin_status),
            rank: 0,
            recent_rank: 0,
            count: 1,
            access_count: 1,
            score: 0.0,
            last_access_filetime: None,
            last_interaction_filetime: None,
            sps_size: None,
            reserved_78: None,
            reserved_7c: None,
            path_sources: Vec::new(),
            warnings: Vec::new(),
        }
    }

    #[test]
    fn frequent_folder_pinned_status_from_entries_detects_pinned_entry() {
        let entries = vec![destlist_entry_for_test("C:\\Folder", 0)];

        assert_eq!(
            frequent_folder_pinned_status_from_entries("c:/folder", &entries),
            PinnedStatus::Pinned
        );
    }

    #[test]
    fn frequent_folder_pinned_status_from_entries_detects_unpinned_entry() {
        let entries = vec![destlist_entry_for_test("C:\\Folder", -1)];

        assert_eq!(
            frequent_folder_pinned_status_from_entries("C:\\Folder", &entries),
            PinnedStatus::Unpinned
        );
    }

    #[test]
    fn frequent_folder_pinned_status_from_entries_prefers_any_pinned_match() {
        let entries = vec![
            destlist_entry_for_test("C:\\Folder", -1),
            destlist_entry_for_test("C:\\Folder", 0),
        ];

        assert_eq!(
            frequent_folder_pinned_status_from_entries("C:\\Folder", &entries),
            PinnedStatus::Pinned
        );
    }

    #[test]
    fn frequent_folder_pinned_status_from_entries_returns_unknown_for_unmatched_path() {
        let entries = vec![destlist_entry_for_test("C:\\Other", 0)];

        assert_eq!(
            frequent_folder_pinned_status_from_entries("C:\\Folder", &entries),
            PinnedStatus::Unknown
        );
    }

    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_pin_unpin_frequent_folder -- --ignored --nocapture"]
    fn test_pin_unpin_frequent_folder() -> WincentResult<()> {
        let test_dir = setup_test_env()?;
        let test_path = path_to_str(&test_dir)?;

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
    #[ignore = "Modifies system state - run with: cargo test test_unpin_native_error_classification -- --ignored --nocapture"]
    fn test_unpin_native_error_classification() -> WincentResult<()> {
        // Tests the critical error classification logic in unpin_frequent_folder_native()
        // The function checks namespace presence before any Shell verb and handles
        // unpinned entries with an explicit pin-then-unpin flow.
        //
        // Key logic:
        // 1. Check if folder is present in Frequent Folders
        // 2. If absent, return NotInQuickAccess immediately
        // 3. If present, remove via unpinfromhome/pintohome Shell verbs

        let test_dir = setup_test_env()?;
        let test_path = path_to_str(&test_dir)?;

        // Scenario 1: Unpin a folder that is NOT in frequent folders
        // Expected: Should return NotInQuickAccess error immediately
        let result = unpin_frequent_folder_native(test_path, DEFAULT_COM_TIMEOUT);

        assert!(
            result.is_err(),
            "Unpinning a non-pinned folder should return error"
        );

        if let Err(e) = result {
            assert!(
                matches!(e, WincentError::NotInQuickAccess { .. }),
                "Unpinning a non-pinned folder should return NotInQuickAccess, got: {:?}",
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
        assert!(!is_pinned, "Folder should not be pinned after unpinning");

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_add_remove_file_in_recent -- --ignored --nocapture"]
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
        let test_path = path_to_str(&test_file)?;

        // Use public API consistently
        add_file_to_recent_native(test_path, DEFAULT_COM_TIMEOUT)?;

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
    #[ignore = "Modifies system state - run with: cargo test test_remove_recent_file_native_direct -- --ignored --nocapture"]
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
        // - Returns NotInQuickAccess when file is not found
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
        let test_path = path_to_str(&test_file)?;

        // Add file to recent using native API
        add_file_to_recent_native(test_path, DEFAULT_COM_TIMEOUT)?;

        // Wait longer for Windows to process the recent item (20 retries = 10 seconds)
        if !wait_for_file_status(test_path, true, 20)? {
            // If file doesn't appear, skip this test (Windows Recent Items can be flaky)
            cleanup_test_env(&test_dir)?;
            println!(
                "Skipping test - file did not appear in recent items (Windows Shell timing issue)"
            );
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

        // Test removing non-existent file - should return NotInQuickAccess
        let result = remove_recent_file_native(test_path, DEFAULT_COM_TIMEOUT);
        assert!(result.is_err(), "Should fail when file not in recent");

        if let Err(e) = result {
            let error_msg = format!("{:?}", e);
            assert!(
                error_msg.contains("NotInQuickAccess")
                    || error_msg.contains("not found in recent items"),
                "Should return NotInQuickAccess error when file not found: {:?}",
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
        let result = add_file_to_recent_native("Z:\\NonExistentFile.txt", DEFAULT_COM_TIMEOUT);
        assert!(result.is_err(), "Should fail with non-existent file");

        let result = add_file_to_recent_native("", DEFAULT_COM_TIMEOUT);
        assert!(result.is_err(), "Should fail with empty path");

        let result = add_file_to_recent_native("\0invalid\0path", DEFAULT_COM_TIMEOUT);
        assert!(
            result.is_err(),
            "Invalid path characters should not be allowed"
        );

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
    fn remove_recent_file_powershell_sentinel_maps_to_not_in_quick_access() {
        let error = WincentError::PowerShellExecution(Box::new(
            crate::error::PowerShellError::builder(
                crate::error::PowerShellOperation::RemoveRecentFile,
            )
            .stdout("WINCENT_NOT_IN_QUICK_ACCESS")
            .stderr("")
            .parameters("C:\\report.docx")
            .build(),
        ));

        let mapped = map_remove_recent_file_powershell_error("C:\\report.docx", error);

        assert!(
            matches!(
                mapped,
                WincentError::NotInQuickAccess {
                    ref path,
                    qa_type: QuickAccess::RecentFiles,
                } if path == "C:\\report.docx"
            ),
            "sentinel should map to NotInQuickAccess, got: {:?}",
            mapped
        );
    }

    #[test]
    fn remove_recent_file_powershell_keeps_non_sentinel_error() {
        let error = WincentError::PowerShellExecution(Box::new(
            crate::error::PowerShellError::builder(
                crate::error::PowerShellOperation::RemoveRecentFile,
            )
            .stdout("ordinary output")
            .stderr("ordinary failure")
            .parameters("C:\\report.docx")
            .build(),
        ));

        let mapped = map_remove_recent_file_powershell_error("C:\\report.docx", error);

        assert!(
            matches!(mapped, WincentError::PowerShellExecution(_)),
            "ordinary PowerShell failures should remain PowerShellExecution, got: {:?}",
            mapped
        );
    }

    #[test]
    fn remove_recent_file_script_uses_caller_timeout() -> WincentResult<()> {
        use std::os::windows::process::ExitStatusExt;
        use std::process::Output;

        let temp_file =
            tempfile::NamedTempFile::new().map_err(|e| WincentError::SystemError(e.to_string()))?;
        let file = temp_file.path().to_string_lossy().into_owned();
        let expected_file = file.clone();
        let expected_timeout = Duration::from_millis(1234);
        let observed_timeout = std::rc::Rc::new(std::cell::RefCell::new(None));
        let observed_timeout_for_closure = std::rc::Rc::clone(&observed_timeout);

        execute_remove_recent_file_script_with_timeout(
            &file,
            expected_timeout,
            move |script, parameter, timeout| {
                assert_eq!(script, PSScript::RemoveRecentFile);
                assert_eq!(parameter, Some(expected_file.as_str()));
                *observed_timeout_for_closure.borrow_mut() = Some(timeout);
                Ok(Output {
                    status: std::process::ExitStatus::from_raw(0),
                    stdout: Vec::new(),
                    stderr: Vec::new(),
                })
            },
        )?;

        assert_eq!(*observed_timeout.borrow(), Some(expected_timeout));
        Ok(())
    }

    #[test]
    fn pin_frequent_folder_script_uses_caller_timeout() -> WincentResult<()> {
        use std::os::windows::process::ExitStatusExt;
        use std::process::Output;

        let temp_dir = tempfile::tempdir().map_err(|e| WincentError::SystemError(e.to_string()))?;
        let folder = temp_dir.path().join("timeout-pin");
        std::fs::create_dir(&folder).map_err(WincentError::Io)?;
        let folder = folder.to_string_lossy().into_owned();
        let expected_folder = folder.clone();
        let expected_timeout = Duration::from_millis(1234);
        let observed_timeout = std::rc::Rc::new(std::cell::RefCell::new(None));
        let observed_timeout_for_closure = std::rc::Rc::clone(&observed_timeout);

        execute_pin_frequent_folder_script_with_timeout(
            &folder,
            expected_timeout,
            move |script, parameter, timeout| {
                assert_eq!(script, PSScript::PinToFrequentFolder);
                assert_eq!(parameter, Some(expected_folder.as_str()));
                *observed_timeout_for_closure.borrow_mut() = Some(timeout);
                Ok(Output {
                    status: std::process::ExitStatus::from_raw(0),
                    stdout: Vec::new(),
                    stderr: Vec::new(),
                })
            },
        )?;

        assert_eq!(*observed_timeout.borrow(), Some(expected_timeout));
        Ok(())
    }

    #[test]
    fn pin_frequent_folder_powershell_sentinel_maps_to_already_exists() {
        let error = WincentError::PowerShellExecution(Box::new(
            crate::error::PowerShellError::builder(
                crate::error::PowerShellOperation::PinFrequentFolder,
            )
            .stdout("WINCENT_ALREADY_EXISTS")
            .stderr("")
            .parameters("C:\\Folder")
            .build(),
        ));

        let mapped = map_pin_frequent_folder_powershell_error("C:\\Folder", error);

        assert!(
            matches!(
                mapped,
                WincentError::AlreadyExists {
                    ref path,
                    qa_type: QuickAccess::FrequentFolders,
                } if path == "C:\\Folder"
            ),
            "sentinel should map to AlreadyExists, got: {:?}",
            mapped
        );
    }

    #[test]
    fn pin_frequent_folder_powershell_keeps_non_sentinel_error() {
        let error = WincentError::PowerShellExecution(Box::new(
            crate::error::PowerShellError::builder(
                crate::error::PowerShellOperation::PinFrequentFolder,
            )
            .stdout("ordinary output")
            .stderr("ordinary failure")
            .parameters("C:\\Folder")
            .build(),
        ));

        let mapped = map_pin_frequent_folder_powershell_error("C:\\Folder", error);

        assert!(
            matches!(mapped, WincentError::PowerShellExecution(_)),
            "ordinary PowerShell failures should remain PowerShellExecution, got: {:?}",
            mapped
        );
    }

    #[test]
    fn unpin_frequent_folder_powershell_sentinel_maps_to_not_in_quick_access() {
        let error = WincentError::PowerShellExecution(Box::new(
            crate::error::PowerShellError::builder(
                crate::error::PowerShellOperation::UnpinFrequentFolder,
            )
            .stdout("WINCENT_NOT_IN_QUICK_ACCESS")
            .stderr("")
            .parameters("C:\\Folder")
            .build(),
        ));

        let mapped = map_unpin_frequent_folder_powershell_error("C:\\Folder", error);

        assert!(
            matches!(
                mapped,
                WincentError::NotInQuickAccess {
                    ref path,
                    qa_type: QuickAccess::FrequentFolders,
                } if path == "C:\\Folder"
            ),
            "sentinel should map to NotInQuickAccess, got: {:?}",
            mapped
        );
    }

    #[test]
    fn unpin_frequent_folder_powershell_keeps_non_sentinel_error() {
        let error = WincentError::PowerShellExecution(Box::new(
            crate::error::PowerShellError::builder(
                crate::error::PowerShellOperation::UnpinFrequentFolder,
            )
            .stdout("ordinary output")
            .stderr("ordinary failure")
            .parameters("C:\\Folder")
            .build(),
        ));

        let mapped = map_unpin_frequent_folder_powershell_error("C:\\Folder", error);

        assert!(
            matches!(mapped, WincentError::PowerShellExecution(_)),
            "ordinary PowerShell failures should remain PowerShellExecution, got: {:?}",
            mapped
        );
    }

    #[test]
    fn unpin_frequent_folder_script_uses_caller_timeout() -> WincentResult<()> {
        use std::os::windows::process::ExitStatusExt;
        use std::process::Output;

        let temp_dir = tempfile::tempdir().map_err(|e| WincentError::SystemError(e.to_string()))?;
        let folder = temp_dir.path().join("timeout-unpin");
        std::fs::create_dir(&folder).map_err(WincentError::Io)?;
        let folder = folder.to_string_lossy().into_owned();
        let expected_folder = folder.clone();
        let expected_timeout = Duration::from_millis(1234);
        let observed_timeout = std::rc::Rc::new(std::cell::RefCell::new(None));
        let observed_timeout_for_closure = std::rc::Rc::clone(&observed_timeout);

        execute_unpin_frequent_folder_script_with_timeout(
            &folder,
            expected_timeout,
            move |script, parameter, timeout| {
                assert_eq!(script, PSScript::UnpinFromFrequentFolder);
                assert_eq!(parameter, Some(expected_folder.as_str()));
                *observed_timeout_for_closure.borrow_mut() = Some(timeout);
                Ok(Output {
                    status: std::process::ExitStatus::from_raw(0),
                    stdout: Vec::new(),
                    stderr: Vec::new(),
                })
            },
        )?;

        assert_eq!(*observed_timeout.borrow(), Some(expected_timeout));
        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_remove_from_recent_files_preserves_not_in_recent -- --ignored --nocapture"]
    fn test_remove_from_recent_files_preserves_not_in_recent() -> WincentResult<()> {
        let test_dir = setup_test_env()?;
        let test_file =
            create_test_file(&test_dir, "not_in_recent_wrapper_test.txt", "test content")?;
        let test_path = path_to_str(&test_file)?;

        let _ = remove_recent_file_native(test_path, DEFAULT_COM_TIMEOUT);
        let result = remove_from_recent_files(test_path);

        cleanup_test_env(&test_dir)?;

        assert!(
            matches!(result, Err(WincentError::NotInQuickAccess { .. })),
            "public wrapper should preserve NotInQuickAccess instead of falling back to PowerShell: {:?}",
            result
        );

        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_add_file_to_recent_with_spaces -- --ignored --nocapture"]
    fn test_add_file_to_recent_with_spaces() -> WincentResult<()> {
        // Tests that add_file_to_recent_native() works for filenames that contain spaces.
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
        let test_path = path_to_str(&test_file)?;
        add_file_to_recent_native(test_path, DEFAULT_COM_TIMEOUT)?;

        // Test file with spaces
        let filename2 = format!("test file with spaces {}.txt", timestamp);
        let test_file2 = create_test_file(&test_dir, &filename2, "test content")?;
        let test_path2 = path_to_str(&test_file2)?;
        add_file_to_recent_native(test_path2, DEFAULT_COM_TIMEOUT)?;

        // Wait for both files to appear (20 retries 閼?500ms = 10 seconds)
        if !wait_for_file_status(test_path, true, 20)? {
            cleanup_test_env(&test_dir)?;
            println!(
                "Skipping test - file did not appear in recent items (Windows Shell timing issue)"
            );
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
    #[ignore = "Modifies system state - run with: cargo test test_unpin_folder_item_compatibility -- --ignored --nocapture"]
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
        let test_path = path_to_str(&test_dir)?;

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
        let found = crate::com_thread::run_on_sta_thread(
            move || {
                let folder = shell_folder(FREQUENT_FOLDERS_NAMESPACE)?;
                let items = folder_items(&folder)?;

                // SAFETY: `items` is a live FolderItems collection on the STA worker.
                let count = unsafe {
                    items.Count().map_err(|e| {
                        WincentError::SystemError(format!("Failed to get item count: {}", e))
                    })?
                };

                for index in 0..count {
                    let index_variant = VARIANT::from(index);
                    // SAFETY: `items` is live for this loop and the VARIANT
                    // index lives until Item returns. Per-item failures are skipped.
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
            },
            DEFAULT_COM_TIMEOUT,
        )?;

        assert!(
            found,
            "Should have found the test folder in frequent folders"
        );

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
    #[ignore = "Performance benchmark - run with: cargo test test_native_pin_unpin_performance -- --ignored --nocapture"]
    fn test_native_pin_unpin_performance() -> WincentResult<()> {
        use std::time::Instant;

        let test_dir = setup_test_env()?;
        let test_path = path_to_str(&test_dir)?;

        // Benchmark native API
        let start = Instant::now();
        pin_frequent_folder_native(test_path, DEFAULT_COM_TIMEOUT)?;
        let native_pin = start.elapsed();

        thread::sleep(Duration::from_millis(100));

        let start = Instant::now();
        unpin_frequent_folder_native(test_path, DEFAULT_COM_TIMEOUT)?;
        let native_unpin = start.elapsed();

        println!(
            "Native API - Pin: {:?}, Unpin: {:?}",
            native_pin, native_unpin
        );
        println!("Total: {:?}", native_pin + native_unpin);

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_com_s_false_reference_counting -- --ignored --nocapture"]
    fn test_com_s_false_reference_counting() -> WincentResult<()> {
        // add_file_to_recent_native now runs on a dedicated STA thread via
        // run_on_sta_thread, so COM reference counting on the calling thread
        // is not affected. This test verifies the calling thread's apartment
        // model is unaffected after several calls.
        use windows::Win32::System::Com::{
            CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED,
        };

        let test_dir = setup_test_env()?;
        let test_file = create_test_file(&test_dir, "s_false_test.txt", "content")?;
        let test_path = path_to_str(&test_file)?;

        // SAFETY: The test pairs each successful CoInitializeEx with
        // CoUninitialize on the same thread.
        unsafe {
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            assert_eq!(hr.0, 0, "First CoInitializeEx should return S_OK");

            for _ in 0..3 {
                let result = add_file_to_recent_native(test_path, DEFAULT_COM_TIMEOUT);
                assert!(result.is_ok(), "Should succeed: {:?}", result);
            }

            CoUninitialize();

            // The calling thread has fully uninitialised; MTA should succeed.
            let hr = CoInitializeEx(None, COINIT_MULTITHREADED);
            assert!(
                hr.0 == 0 || hr.0 == 1,
                "Should be able to initialize with MTA (no leaked STA references). Got: 0x{:08X}",
                hr.0
            );
            CoUninitialize();
        }

        let _ = remove_recent_file_native(test_path, DEFAULT_COM_TIMEOUT);
        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_com_apartment_mismatch -- --ignored --nocapture"]
    fn test_com_apartment_mismatch() -> WincentResult<()> {
        // add_file_to_recent_native now routes through run_on_sta_thread which
        // spawns a fresh STA thread regardless of the caller's apartment model.
        // Calling from an MTA thread must NOT return ComApartmentMismatch.
        use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED};

        // SAFETY: The MTA initialization in this test is balanced by
        // CoUninitialize before returning.
        unsafe {
            let hr = CoInitializeEx(None, COINIT_MULTITHREADED);
            assert!(hr.is_ok() || hr.0 == 1, "MTA init should succeed");

            let result = add_file_to_recent_native("C:\\Windows\\notepad.exe", DEFAULT_COM_TIMEOUT);
            assert!(
                !matches!(result, Err(WincentError::ComApartmentMismatch(_))),
                "MTA caller should no longer get ComApartmentMismatch, got: {:?}",
                result
            );

            CoUninitialize();
        }

        Ok(())
    }

    #[test]
    fn test_with_timeout_zero_rejected() {
        // All three *_with_timeout public functions must reject Duration::ZERO
        // before any operation (path validation, native COM, PowerShell fallback).
        let r = remove_from_recent_files_with_timeout(
            "C:\\Windows\\notepad.exe",
            std::time::Duration::ZERO,
        );
        assert!(
            matches!(r, Err(WincentError::InvalidArgument(_))),
            "got: {:?}",
            r
        );

        let r = add_to_frequent_folders_with_timeout("C:\\Windows", std::time::Duration::ZERO);
        assert!(
            matches!(r, Err(WincentError::InvalidArgument(_))),
            "got: {:?}",
            r
        );

        let r = remove_from_frequent_folders_with_timeout("C:\\Windows", std::time::Duration::ZERO);
        assert!(
            matches!(r, Err(WincentError::InvalidArgument(_))),
            "got: {:?}",
            r
        );

        let r = add_file_to_recent_native("C:\\Windows\\notepad.exe", std::time::Duration::ZERO);
        assert!(
            matches!(r, Err(WincentError::InvalidArgument(_))),
            "got: {:?}",
            r
        );
    }
}
