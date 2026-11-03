//! Windows Quick Access item management
//!
//! Provides system-level manipulation of Quick Access locations including:
//! - Recent files management
//! - Frequent folders pinning
//! - Native add and removal operations
//!
//! # Key Functionality
//! - File addition/removal from Recent Items
//! - Folder pinning/unpinning operations
//! - Path validation and sanitization
//! - Native SHAddToRecentDocs and Shell COM additions
//!
//! # Operation Safety
//! 1. Dedicated STA threading for Shell COM operations
//! 2. Path validation before execution
//! 3. Clean error propagation

use crate::{
    destlist::FrequentFolderPinStatus,
    error::WincentError,
    query::{folder_items, item_path, shell_folder, FREQUENT_FOLDERS_NAMESPACE},
    quick_access_lock::MutationPermit,
    utils::{paths_equal, validate_path, PathType},
    QuickAccess, WincentResult,
};
use std::ffi::OsString;
use std::os::windows::prelude::*;
use std::time::{Duration, Instant};
use windows::core::Interface;
use windows::Win32::System::Variant::VARIANT;
use windows::Win32::UI::Shell::{Folder3, SHAddToRecentDocs};

const FREQUENT_FOLDER_VERIFICATION_TIMEOUT: Duration = Duration::from_secs(1);
const FREQUENT_FOLDER_VERIFICATION_POLL_INTERVAL: Duration = Duration::from_millis(100);
const REMOVE_FROM_HOME_VERB: &str = "removefromhome";
const UNPIN_FROM_HOME_VERB: &str = "unpinfromhome";
/// SHARD_PATHW registers a wide-string path in the recent documents list.
const SHARD_PATHW: u32 = 0x0000_0003;

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

fn remove_frequent_folder_with_verb_current_sta(path: &str, verb: &str) -> WincentResult<()> {
    if !contains_frequent_folder_current_sta(path)? {
        return Err(WincentError::not_in_quick_access(
            path,
            QuickAccess::FrequentFolders,
        ));
    }

    if !try_invoke_verb_on_frequent_folder_current_sta(path, verb)? {
        return Err(WincentError::SystemError(format!(
            "Frequent Folders item disappeared before {verb}: {path}"
        )));
    }

    let mut contains = contains_frequent_folder_current_sta;
    let mut sleep = std::thread::sleep;
    if wait_for_frequent_folder_removed(
        path,
        FREQUENT_FOLDER_VERIFICATION_TIMEOUT,
        FREQUENT_FOLDER_VERIFICATION_POLL_INTERVAL,
        &mut contains,
        &mut sleep,
    )? {
        return Ok(());
    }

    Err(WincentError::SystemError(format!(
        "{verb} completed but Frequent Folders still contains {path}"
    )))
}

fn remove_unpinned_frequent_folder_current_sta(path: &str) -> WincentResult<()> {
    remove_frequent_folder_with_verb_current_sta(path, REMOVE_FROM_HOME_VERB)
}

fn remove_pinned_frequent_folder_current_sta(path: &str) -> WincentResult<()> {
    remove_frequent_folder_with_verb_current_sta(path, UNPIN_FROM_HOME_VERB)
}

/// Invokes a Shell verb directly on a folder using Folder.Self pattern
///
/// This function invokes the verb through the Shell `Namespace(path).Self`
/// pattern. It is more efficient than enumerating parent folder items to find
/// the target.
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
    // "pintohome" can act as a toggle. Check in the same STA flow as the verb
    // invocation so an already-pinned folder is never toggled off.
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
/// - `AlreadyExists`: The folder is already pinned. Existing unpinned frequent
///   entries are allowed to proceed through the pin verb.
/// - `SystemError`: COM operation failed (e.g., failed to open folder namespace, invoke verb)
///
/// # See Also
///
/// - [`pin_frequent_folder()`] - Validating native wrapper used by public add APIs
/// - [`invoke_verb_on_self_current_sta()`] - The underlying verb invocation mechanism
fn pin_frequent_folder_native(
    path: &str,
    timeout: std::time::Duration,
    permit: MutationPermit,
) -> WincentResult<()> {
    // Check and invoke in one STA worker to narrow the TOCTOU window.
    // "pintohome" can act as a toggle. The check prevents toggling an
    // existing folder off.
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread_with_permit(
        move || {
            run_pin_frequent_folder_checked(
                &path,
                |candidate| {
                    crate::destlist::frequent_folder_pin_status(candidate)
                        .map(|status| matches!(status, FrequentFolderPinStatus::Pinned))
                },
                invoke_verb_on_self_current_sta,
            )
        },
        timeout,
        permit,
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

fn wait_for_frequent_folder_removed<C, S>(
    path: &str,
    verification_timeout: Duration,
    poll_interval: Duration,
    contains: &mut C,
    sleep: &mut S,
) -> WincentResult<bool>
where
    C: FnMut(&str) -> WincentResult<bool>,
    S: FnMut(Duration),
{
    wait_for_frequent_folder_presence_with(
        path,
        false,
        verification_timeout,
        poll_interval,
        contains,
        sleep,
    )
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

/// Adds a file to the Windows Recent Items list with `SHAddToRecentDocs`.
///
/// The library requires `path` to exist and be a file before calling
/// `SHAddToRecentDocs(SHARD_PATHW, path)`. The Windows function itself is more
/// permissive: experiments accepted missing absolute file-like paths, ignored
/// existing directories for DestList purposes, and returned no status because
/// the API has a `void` return type. Keeping validation here gives the public
/// API deterministic file-only semantics.
///
/// # Windows Behavior Notes
///
/// - A normal add creates or updates the Recent entry and can also add one
///   usage point for the parent folder in Frequent Folders.
/// - Zero-byte, extensionless, read-only, and hidden files are accepted. A file
///   hidden when added is persisted but can be filtered from the UI according
///   to Explorer's hidden-item setting.
/// - Re-adding a path can update its existing entry instead of creating a
///   duplicate. Shell persistence is session-sensitive and asynchronous.
/// - `timeout` limits only how long the caller waits. The Shell operation can
///   still complete after the timeout is reported.
pub(crate) fn add_file_to_recent_native(
    path: &str,
    timeout: Duration,
    permits: Vec<MutationPermit>,
) -> WincentResult<()> {
    validate_path(path, PathType::File)?;
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread_with_permits(
        move || add_file_to_recent_native_current_sta(&path),
        timeout,
        permits,
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
/// - [`crate::manager::QuickAccessManager::remove_item`] - Public removal API
fn remove_recent_file_native(
    path: &str,
    timeout: std::time::Duration,
    permit: MutationPermit,
) -> WincentResult<()> {
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread_with_permit(
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
        permit,
    )
}

/// Pins a folder to the Windows Quick Access Frequent Folders list with COM.
///
/// The folder must exist and be a directory. The implementation performs its
/// pinned-state check and invokes the canonical `pintohome` verb through native
/// Shell COM on one dedicated STA thread.
///
/// Observed Windows 10 22H2 behavior:
///
/// - The result is a pinned Frequent Folder, not an unpinned usage entry.
/// - The new item is appended after all existing pinned items and therefore
///   appears before displayed unpinned items.
/// - A newly pinned item starts with zero usage count and score.
/// - `pintohome` can act like a toggle, so the same STA flow rejects an already
///   pinned path before invoking the verb.
///
/// # Errors
///
/// - `InvalidPath`: Path validation failed (not a directory, doesn't exist)
/// - `AlreadyExists`: The folder is already pinned
/// - `SystemError` or `Timeout`: Native COM failed or the caller stopped waiting
///
/// # See Also
///
/// - [`crate::manager::QuickAccessManager::add_item`] - Public mutation API
/// - [`pin_frequent_folder_native()`] - Native COM implementation
pub(crate) fn pin_frequent_folder(
    path: &str,
    timeout: std::time::Duration,
    permit: MutationPermit,
) -> WincentResult<()> {
    validate_path(path, PathType::Directory)?;
    pin_frequent_folder_native(path, timeout, permit)
}

/****************************************************** Handle Quick Access ******************************************************/

/// Removes a file from Windows Recent Files with a custom COM STA thread timeout.
///
/// Internal timeout-aware variant used by [`crate::manager::QuickAccessManager::remove_item`].
/// Removal is performed only through the native Shell COM path.
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
    permit: MutationPermit,
) -> WincentResult<()> {
    if timeout.is_zero() {
        return Err(WincentError::InvalidArgument(
            "timeout must be greater than zero".to_string(),
        ));
    }
    validate_path(path, PathType::File)?;
    remove_recent_file_native(path, timeout, permit)
}

/// Pins a folder to Windows Quick Access with a custom COM STA thread timeout.
///
/// Internal timeout-aware variant used by [`crate::manager::QuickAccessManager::add_item`].
/// It uses only native COM and rejects zero timeouts before attempting any
/// mutation.
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
    permit: MutationPermit,
) -> WincentResult<()> {
    if timeout.is_zero() {
        return Err(WincentError::InvalidArgument(
            "timeout must be greater than zero".to_string(),
        ));
    }
    pin_frequent_folder(path, timeout, permit)
}

/// Hides a visible unpinned Frequent Folders entry with the native
/// `removefromhome` Shell verb.
pub(crate) fn remove_unpinned_frequent_folder_with_timeout(
    path: &str,
    timeout: std::time::Duration,
    permit: MutationPermit,
) -> WincentResult<()> {
    if timeout.is_zero() {
        return Err(WincentError::InvalidArgument(
            "timeout must be greater than zero".to_string(),
        ));
    }
    validate_path(path, PathType::Directory)?;
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread_with_permit(
        move || remove_unpinned_frequent_folder_current_sta(&path),
        timeout,
        permit,
    )
}

/// Removes a visible pinned Frequent Folders entry with the native
/// `unpinfromhome` Shell verb.
pub(crate) fn remove_pinned_frequent_folder_with_timeout(
    path: &str,
    timeout: std::time::Duration,
    permit: MutationPermit,
) -> WincentResult<()> {
    if timeout.is_zero() {
        return Err(WincentError::InvalidArgument(
            "timeout must be greater than zero".to_string(),
        ));
    }
    validate_path(path, PathType::Directory)?;
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread_with_permit(
        move || remove_pinned_frequent_folder_current_sta(&path),
        timeout,
        permit,
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::cell::{Cell, RefCell};

    #[test]
    fn path_comparison_uses_windows_semantics() {
        assert!(paths_equal("C:\\Users\\Test", "c:/users/test/"));
        assert!(!paths_equal("C:\\Users\\Test", "D:\\Users\\Test"));
    }

    #[test]
    fn pin_check_rejects_existing_folder_without_invoking_verb() {
        let invoked = Cell::new(false);
        let result = run_pin_frequent_folder_checked(
            "C:\\Folder",
            |_| Ok(true),
            |_, _| {
                invoked.set(true);
                Ok(())
            },
        );

        assert!(matches!(result, Err(WincentError::AlreadyExists { .. })));
        assert!(!invoked.get());
    }

    #[test]
    fn pin_check_invokes_pintohome_for_absent_folder() -> WincentResult<()> {
        let actions = RefCell::new(Vec::new());

        run_pin_frequent_folder_checked(
            "C:\\Folder",
            |_| Ok(false),
            |path, verb| {
                actions.borrow_mut().push(format!("{path}:{verb}"));
                Ok(())
            },
        )?;

        assert_eq!(actions.into_inner(), ["C:\\Folder:pintohome"]);
        Ok(())
    }

    #[test]
    fn pin_check_propagates_lookup_error() {
        let invoked = Cell::new(false);
        let result = run_pin_frequent_folder_checked(
            "C:\\Folder",
            |_| Err(WincentError::SystemError("lookup failed".to_string())),
            |_, _| {
                invoked.set(true);
                Ok(())
            },
        );

        assert!(matches!(
            result,
            Err(WincentError::SystemError(ref message)) if message == "lookup failed"
        ));
        assert!(!invoked.get());
    }

    #[test]
    fn removal_wait_stops_when_folder_disappears() -> WincentResult<()> {
        let checks = Cell::new(0);
        let removed = wait_for_frequent_folder_removed(
            "C:\\Folder",
            Duration::from_millis(10),
            Duration::from_millis(1),
            &mut |_| {
                let current = checks.get();
                checks.set(current + 1);
                Ok(current == 0)
            },
            &mut |_| {},
        )?;

        assert!(removed);
        assert!(checks.get() >= 2);
        Ok(())
    }

    #[test]
    fn timeout_aware_operations_reject_zero_timeout() {
        let recent = "C:\\Windows\\notepad.exe";
        let folder = "C:\\Windows";
        let recent_folder = tempfile::tempdir().expect("temp Recent directory");
        let permit = |target| {
            MutationPermit::acquire_in_folder(target, recent_folder.path(), false)
                .expect("test admission")
                .pop()
                .expect("category admission")
        };
        use crate::quick_access_lock::QuickAccessLockTarget::{FrequentFolders, RecentFiles};

        for result in [
            remove_from_recent_files_with_timeout(recent, Duration::ZERO, permit(RecentFiles)),
            add_to_frequent_folders_with_timeout(folder, Duration::ZERO, permit(FrequentFolders)),
            remove_unpinned_frequent_folder_with_timeout(
                folder,
                Duration::ZERO,
                permit(FrequentFolders),
            ),
            remove_pinned_frequent_folder_with_timeout(
                folder,
                Duration::ZERO,
                permit(FrequentFolders),
            ),
            add_file_to_recent_native(
                recent,
                Duration::ZERO,
                vec![permit(RecentFiles), permit(FrequentFolders)],
            ),
        ] {
            assert!(matches!(result, Err(WincentError::InvalidArgument(_))));
        }
    }
}
