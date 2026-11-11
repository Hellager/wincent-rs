//! Standalone Quick Access operations, independent of `QuickAccessManager`.
//!
//! Queries read persisted DestList data. Mutations and refreshes require a
//! nonzero caller-supplied [`Duration`]. The timeout bounds each supported
//! Shell/COM wait, not the entire multi-step workflow or filesystem access.
//! It does not cancel native work already running. Do not automatically retry
//! a timed-out mutation; inspect the resulting state first.
//!
//! Lock functions return owned guards. Keep the guard alive until unlocking.
//! Dropping it or calling `unlock` releases the lock without cleaning up new
//! Recent shortcuts; `unlock_and_cleanup` also removes newly created shortcuts.
//!
//! # Examples
//!
//! ```rust,no_run
//! use std::time::Duration;
//! use wincent::prelude::*;
//!
//! # fn main() -> WincentResult<()> {
//! let timeout = Duration::from_secs(15);
//! add_to_recent_files(r"C:\Work\report.txt", timeout)?;
//! let files = get_recent_files()?;
//! refresh_quick_access(timeout)?;
//!
//! let lock = lock_recent_files()?;
//! // Keep the guard while Recent Files should be locked.
//! let report = lock.unlock_and_cleanup()?;
//! # Ok(())
//! # }
//! ```

use crate::backend::{self, QuickAccessBackend, SystemQuickAccessBackend};
use crate::destlist::FrequentFolderPinStatus;
use crate::empty;
use crate::error::WincentError;
use crate::lock::{MutationPermit, QuickAccessLock, QuickAccessLockTarget, UnlockReport};
use crate::path::{paths_equal, PathType};
use crate::{QuickAccess, WincentResult};
use std::fmt;
use std::path::Path;
use std::time::Duration;

/// Returns Recent Files in the same order as the manager's category query.
///
/// Reads persisted data and applies the current Explorer visibility rules,
/// including the Recent Files result limit. This is not a live window query.
/// A missing Recent Files backing file returns an empty list.
///
/// # Errors
///
/// Returns the underlying path, I/O, registry or DestList parsing error.
pub fn get_recent_files() -> WincentResult<Vec<String>> {
    backend::ui_visible_items(QuickAccess::RecentFiles)
}

/// Returns Frequent Folders, including pinned folders, in Explorer query order.
///
/// Uses the same persisted-data filtering as the manager's category query. A
/// missing Frequent Folders backing file returns an empty list.
///
/// # Errors
///
/// Returns the underlying path, I/O, registry or DestList parsing error.
pub fn get_frequent_folders() -> WincentResult<Vec<String>> {
    backend::ui_visible_items(QuickAccess::FrequentFolders)
}

/// Checks the persisted pin flag for a path in Frequent Folders.
///
/// This is independent of UI visibility and does not require the target to
/// still exist. An unpinned or absent entry, or a missing Frequent Folders
/// backing file, returns `false`. Recent Files data is not read.
///
/// # Errors
///
/// Returns [`WincentError::InvalidPath`] for an empty path, or the underlying
/// I/O or parsing error. Errors other than a missing backing file are not
/// converted to `false`.
pub fn is_folder_pinned<P: AsRef<Path>>(path: P) -> WincentResult<bool> {
    let path = path_string(path.as_ref())?;
    Ok(matches!(
        pin_status(&SystemQuickAccessBackend, &path)?,
        FrequentFolderPinStatus::Pinned
    ))
}

/// Adds an existing file to Recent Files.
///
/// Native file visits may also update Frequent Folders, so a lock on either
/// category denies this operation. Success does not imply immediate query/UI
/// convergence. See the [module-level timeout contract](self).
///
/// # Errors
///
/// Returns [`WincentError::InvalidArgument`] for a zero timeout,
/// [`WincentError::InvalidPath`] for an empty, missing or non-file path,
/// [`WincentError::AlreadyExists`] for a visible duplicate,
/// [`WincentError::QuickAccessLocked`] for a denied mutation, or an underlying error.
pub fn add_to_recent_files<P: AsRef<Path>>(path: P, timeout: Duration) -> WincentResult<()> {
    mutate(
        path.as_ref(),
        timeout,
        Mutation::AddRecent,
        &SystemQuickAccessBackend,
    )
}

/// Pins an existing directory to Frequent Folders.
///
/// An existing unpinned entry may be pinned; an already-pinned entry is
/// rejected instead of invoking a Shell verb that could toggle it off.
/// Success does not imply immediate query/UI convergence.
///
/// # Errors
///
/// Returns [`WincentError::InvalidArgument`] for a zero timeout,
/// [`WincentError::InvalidPath`] for an empty, missing or non-directory path,
/// [`WincentError::AlreadyExists`] when already pinned,
/// [`WincentError::QuickAccessLocked`] for a denied mutation, or an underlying error.
pub fn pin_to_frequent_folders<P: AsRef<Path>>(path: P, timeout: Duration) -> WincentResult<()> {
    mutate(
        path.as_ref(),
        timeout,
        Mutation::PinFolder,
        &SystemQuickAccessBackend,
    )
}

/// Removes an existing file from Recent Files through the native Shell API.
///
/// Success does not imply immediate query/UI convergence.
///
/// # Errors
///
/// Returns [`WincentError::InvalidArgument`] for a zero timeout,
/// [`WincentError::InvalidPath`] for an empty, missing or non-file path,
/// [`WincentError::NotInQuickAccess`] when absent from the category query,
/// [`WincentError::QuickAccessLocked`] for a denied mutation, or an underlying error.
pub fn remove_from_recent_files<P: AsRef<Path>>(path: P, timeout: Duration) -> WincentResult<()> {
    mutate(
        path.as_ref(),
        timeout,
        Mutation::RemoveRecent,
        &SystemQuickAccessBackend,
    )
}

/// Removes a pinned or unpinned existing directory from Frequent Folders.
///
/// Selects the native removal operation from its persisted pin state, just
/// like the manager. This is not an unpin-only API.
///
/// # Errors
///
/// Returns [`WincentError::InvalidArgument`] for a zero timeout,
/// [`WincentError::InvalidPath`] for an empty, missing or non-directory path,
/// [`WincentError::NotInQuickAccess`] when absent from the category query,
/// [`WincentError::QuickAccessLocked`] for a denied mutation, or an underlying error.
pub fn remove_from_frequent_folders<P: AsRef<Path>>(
    path: P,
    timeout: Duration,
) -> WincentResult<()> {
    mutate(
        path.as_ref(),
        timeout,
        Mutation::RemoveFolder,
        &SystemQuickAccessBackend,
    )
}

/// An owned Recent Files lock, released when dropped.
///
/// Does not implement `Clone`. Dropping or calling [`RecentFilesLock::unlock`]
/// only releases handles. Use [`RecentFilesLock::unlock_and_cleanup`] to also
/// remove and report newly created shortcuts.
#[must_use = "keep the guard alive to keep Recent Files locked"]
pub struct RecentFilesLock {
    inner: QuickAccessLock,
}

impl RecentFilesLock {
    /// Releases the lock without scanning or deleting Recent shortcuts.
    pub fn unlock(self) {
        drop(self);
    }

    /// Releases the lock and cleans up new shortcuts from the Recent directory.
    ///
    /// Cleanup is best-effort. Individual deletion failures are returned in
    /// [`UnlockReport::cleanup_failures`].
    ///
    /// # Errors
    ///
    /// Returns an error if the current Recent shortcuts cannot be enumerated.
    pub fn unlock_and_cleanup(self) -> WincentResult<UnlockReport> {
        self.inner.unlock_and_cleanup()
    }
}

impl fmt::Debug for RecentFilesLock {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("RecentFilesLock").finish_non_exhaustive()
    }
}

/// An owned Frequent Folders lock, released when dropped.
///
/// Does not implement `Clone`. Dropping or calling [`FrequentFoldersLock::unlock`]
/// only releases handles. Use [`FrequentFoldersLock::unlock_and_cleanup`] to also
/// remove and report newly created shortcuts.
#[must_use = "keep the guard alive to keep Frequent Folders locked"]
pub struct FrequentFoldersLock {
    inner: QuickAccessLock,
}

impl FrequentFoldersLock {
    /// Releases the lock without scanning or deleting Recent shortcuts.
    pub fn unlock(self) {
        drop(self);
    }

    /// Releases the lock and cleans up new shortcuts from the Recent directory.
    ///
    /// The shortcut delta covers the entire Recent directory, not only folder
    /// shortcuts. Cleanup is best-effort, with individual deletion failures in
    /// [`UnlockReport::cleanup_failures`].
    ///
    /// # Errors
    ///
    /// Returns an error if the current Recent shortcuts cannot be enumerated.
    pub fn unlock_and_cleanup(self) -> WincentResult<UnlockReport> {
        self.inner.unlock_and_cleanup()
    }
}

impl fmt::Debug for FrequentFoldersLock {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("FrequentFoldersLock")
            .finish_non_exhaustive()
    }
}

/// Locks Recent Files until the returned guard is unlocked or dropped.
///
/// Shares admission gates with managers and other wincent processes. Queries
/// remain available. Non-cooperating Shell clients may still create shortcuts.
///
/// # Errors
///
/// Returns an error if a gate or backing file cannot be locked, including
/// overlapping locks or in-flight cooperating mutations.
pub fn lock_recent_files() -> WincentResult<RecentFilesLock> {
    Ok(RecentFilesLock {
        inner: QuickAccessLock::lock_target(QuickAccessLockTarget::RecentFiles)?,
    })
}

/// Locks Frequent Folders until the returned guard is unlocked or dropped.
///
/// Recent Files and Frequent Folders guards may coexist. Adding a Recent file
/// is still blocked because that operation can affect both categories.
///
/// # Errors
///
/// Returns an error if a gate or backing file cannot be locked, including
/// overlapping locks or in-flight cooperating mutations.
pub fn lock_frequent_folders() -> WincentResult<FrequentFoldersLock> {
    Ok(FrequentFoldersLock {
        inner: QuickAccessLock::lock_target(QuickAccessLockTarget::FrequentFolders)?,
    })
}

/// Notifies the Shell and refreshes open Quick Access/Home views.
///
/// Does not wait for prior mutations to persist or verify their results.
///
/// # Errors
///
/// Returns [`WincentError::InvalidArgument`] for a zero timeout, or an
/// underlying refresh/timeout error.
pub fn refresh_quick_access(timeout: Duration) -> WincentResult<()> {
    validate_timeout(timeout)?;
    SystemQuickAccessBackend.refresh(timeout)
}

/// Clears Recent Files history and shortcuts, then refreshes matching views.
///
/// Uses the same scoped cleanup as the manager. A refresh failure may be
/// returned after cleanup has already succeeded.
///
/// # Errors
///
/// Returns [`WincentError::InvalidArgument`] for a zero timeout,
/// [`WincentError::QuickAccessLocked`] for a denied mutation, or an underlying
/// cleanup/refresh error.
pub fn empty_recent_files(timeout: Duration) -> WincentResult<()> {
    empty_category(QuickAccess::RecentFiles, timeout, &SystemQuickAccessBackend)
}

/// Clears recorded Frequent Folders and their shortcuts, then refreshes views.
///
/// User pins are removed; Windows may recreate default pins. A refresh
/// failure may be returned after cleanup has already succeeded.
///
/// # Errors
///
/// Returns [`WincentError::InvalidArgument`] for a zero timeout,
/// [`WincentError::QuickAccessLocked`] for a denied mutation, or an underlying
/// cleanup/refresh error.
pub fn empty_frequent_folders(timeout: Duration) -> WincentResult<()> {
    empty_category(
        QuickAccess::FrequentFolders,
        timeout,
        &SystemQuickAccessBackend,
    )
}

fn validate_timeout(timeout: Duration) -> WincentResult<()> {
    if timeout.is_zero() {
        return Err(WincentError::InvalidArgument(
            "timeout must be greater than zero".into(),
        ));
    }
    Ok(())
}

fn path_string(path: &Path) -> WincentResult<String> {
    if path.as_os_str().is_empty() {
        return Err(WincentError::invalid_path_reason("Empty path provided"));
    }
    Ok(path.to_string_lossy().into_owned())
}

fn pin_status(
    backend: &dyn QuickAccessBackend,
    path: &str,
) -> WincentResult<FrequentFolderPinStatus> {
    backend.frequent_folder_pin_status(path)
}

#[derive(Clone, Copy)]
enum Mutation {
    AddRecent,
    PinFolder,
    RemoveRecent,
    RemoveFolder,
}

fn mutate(
    path: &Path,
    timeout: Duration,
    operation: Mutation,
    backend: &dyn QuickAccessBackend,
) -> WincentResult<()> {
    validate_timeout(timeout)?;
    let path = path_string(path)?;
    let (expected, category, target) = match operation {
        Mutation::AddRecent => (
            PathType::File,
            QuickAccess::RecentFiles,
            QuickAccessLockTarget::All,
        ),
        Mutation::RemoveRecent => (
            PathType::File,
            QuickAccess::RecentFiles,
            QuickAccessLockTarget::RecentFiles,
        ),
        Mutation::PinFolder | Mutation::RemoveFolder => (
            PathType::Directory,
            QuickAccess::FrequentFolders,
            QuickAccessLockTarget::FrequentFolders,
        ),
    };
    if backend.path_type(&path)? != expected {
        return Err(WincentError::invalid_path(
            &path,
            format!("Expected a {expected:?} path"),
        ));
    }
    // Hold admission through preflight too. Native workers retain their own
    // admission handles if the caller times out, as in the manager workflow.
    let _admission = backend
        .uses_system_mutation_gate()
        .then(|| MutationPermit::acquire(target))
        .transpose()?;

    if matches!(operation, Mutation::PinFolder) {
        if matches!(pin_status(backend, &path)?, FrequentFolderPinStatus::Pinned) {
            return Err(WincentError::already_exists(&path, category));
        }
        return backend.add_frequent_folder(&path, timeout);
    }

    let items = backend.get_items(category, timeout)?;
    let present = items.iter().any(|item| paths_equal(item, &path));
    if matches!(operation, Mutation::AddRecent) {
        if present {
            return Err(WincentError::already_exists(&path, category));
        }
        return backend.add_recent_file(&path, timeout);
    }
    if !present {
        return Err(WincentError::not_in_quick_access(&path, category));
    }
    if matches!(operation, Mutation::RemoveRecent) {
        return backend.remove_recent_file(&path, timeout);
    }
    match backend.frequent_folder_pin_status(&path)? {
        FrequentFolderPinStatus::Pinned => backend.remove_pinned_frequent_folder(&path, timeout),
        FrequentFolderPinStatus::Unpinned => {
            backend.remove_unpinned_frequent_folder(&path, timeout)
        }
        FrequentFolderPinStatus::NotFound => {
            Err(WincentError::not_in_quick_access(&path, category))
        }
    }
}

fn empty_category(
    category: QuickAccess,
    timeout: Duration,
    backend: &dyn QuickAccessBackend,
) -> WincentResult<()> {
    validate_timeout(timeout)?;
    let target = match category {
        QuickAccess::RecentFiles => QuickAccessLockTarget::RecentFiles,
        QuickAccess::FrequentFolders => QuickAccessLockTarget::FrequentFolders,
        QuickAccess::All => QuickAccessLockTarget::All,
    };
    let _admission = backend
        .uses_system_mutation_gate()
        .then(|| MutationPermit::acquire(target))
        .transpose()?;
    empty::empty_items_with_backend(category, timeout, backend)
}

#[cfg(test)]
mod tests;
