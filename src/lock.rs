//! Locks Explorer Quick Access backing files and gates cooperating wincent mutations.
//!
//! While a target is locked, its backing file is not writable through new
//! handles, and cooperating wincent instances reject mutations.
//! Explorer operations may still report success, time out, or create Recent
//! shortcuts while the lock is held.
//! Windows may still create Recent-folder `.lnk` shortcuts during the lock.

use crate::error::WincentError;
use crate::path::paths_equal;
use crate::recent_links::recent_lnk_paths;
use crate::windows::{get_windows_recent_folder, os_str_to_wide_null};
use crate::WincentResult;
use std::fs;
use std::io;
use std::path::{Path, PathBuf};
use windows::core::PCWSTR;
use windows::Win32::Foundation::{CloseHandle, GENERIC_READ, HANDLE};
use windows::Win32::Storage::FileSystem::{
    CreateFileW, FILE_ATTRIBUTE_NORMAL, FILE_SHARE_NONE, FILE_SHARE_READ, OPEN_ALWAYS,
    OPEN_EXISTING,
};

pub(crate) const RECENT_FILES_AUTOMATIC_DESTINATION: &str =
    "5f7b5f1e01b83767.automaticDestinations-ms";
pub(crate) const FREQUENT_FOLDERS_AUTOMATIC_DESTINATION: &str =
    "f01b4d95cf55d32a.automaticDestinations-ms";
const RECENT_GATE: &str = ".wincent-recent-files.lock";
const FREQUENT_GATE: &str = ".wincent-frequent-folders.lock";

/// Quick Access backing file set to lock.
///
/// Locking a target rejects cooperating wincent mutations and inhibits writes
/// to the corresponding Explorer backing file while the lock is held.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum QuickAccessLockTarget {
    /// Lock the Recent Files automatic destination file.
    ///
    /// Recent Files mutations through wincent are rejected while this lock is held.
    RecentFiles,
    /// Lock the Frequent Folders automatic destination file.
    ///
    /// Frequent Folders mutations through wincent are rejected while this lock is held.
    FrequentFolders,
    /// Lock both Recent Files and Frequent Folders automatic destination files.
    ///
    /// Mutations through wincent are rejected for both categories while held.
    All,
}

/// Results from cleaning up shortcuts while releasing a Quick Access lock.
#[must_use = "inspect cleanup failures before discarding an unlock report"]
#[derive(Debug)]
#[non_exhaustive]
pub struct UnlockReport {
    cleaned_shortcuts: Vec<PathBuf>,
    cleanup_failures: Vec<ShortcutCleanupFailure>,
}

impl UnlockReport {
    /// Recent `.lnk` paths that were removed or were already absent.
    #[must_use]
    pub fn cleaned_shortcuts(&self) -> &[PathBuf] {
        &self.cleaned_shortcuts
    }

    /// Recent `.lnk` paths that could not be removed.
    #[must_use]
    pub fn cleanup_failures(&self) -> &[ShortcutCleanupFailure] {
        &self.cleanup_failures
    }

    /// Returns `true` when every discovered shortcut was cleaned up.
    #[must_use]
    pub fn is_complete(&self) -> bool {
        self.cleanup_failures.is_empty()
    }
}

/// A Recent `.lnk` file that could not be removed during unlock cleanup.
#[derive(Debug)]
#[non_exhaustive]
pub struct ShortcutCleanupFailure {
    path: PathBuf,
    error: io::Error,
}

impl ShortcutCleanupFailure {
    /// Returns the shortcut path that could not be removed.
    #[must_use]
    pub fn path(&self) -> &Path {
        &self.path
    }

    /// Returns the filesystem error raised while removing the shortcut.
    #[must_use]
    pub fn error(&self) -> &io::Error {
        &self.error
    }
}

/// Holds category admission gates and locks backing files.
///
/// While this guard is alive, Explorer cannot update the locked automatic
/// destination files. Mutations made through cooperating wincent instances
/// return [`WincentError::QuickAccessLocked`] for a locked category. Other
/// Explorer/Shell clients may still report success or time out even when a
/// backing-file write is denied.
///
/// The manager owns this value. Dropping it releases the backing-file locks.
pub(crate) struct QuickAccessLock {
    recent_folder: PathBuf,
    initial_lnk_paths: Vec<PathBuf>,
    /// Held only for their Drop behavior; keeping these handles open is the lock.
    _locks: Vec<LockedFile>,
    _gates: Vec<MutationPermit>,
}

/// An admission handle retained by a Shell worker after caller timeout.
pub(crate) struct MutationPermit {
    handle: HANDLE,
}

impl MutationPermit {
    pub(crate) fn acquire(target: QuickAccessLockTarget) -> WincentResult<Vec<Self>> {
        let recent_folder = PathBuf::from(get_windows_recent_folder()?);
        Self::acquire_in_folder(target, &recent_folder, false)
    }

    pub(crate) fn acquire_one(target: QuickAccessLockTarget) -> WincentResult<Self> {
        Self::acquire(target)?.pop().ok_or_else(|| {
            WincentError::InvalidArgument("Quick Access lock target has no category".into())
        })
    }

    pub(crate) fn acquire_in_folder(
        target: QuickAccessLockTarget,
        recent_folder: &Path,
        exclusive: bool,
    ) -> WincentResult<Vec<Self>> {
        let mut gates = Vec::new();
        for (category, file) in [
            (QuickAccessLockTarget::RecentFiles, RECENT_GATE),
            (QuickAccessLockTarget::FrequentFolders, FREQUENT_GATE),
        ] {
            if target != category && target != QuickAccessLockTarget::All {
                continue;
            }
            let path = recent_folder.join(file);
            let mut wide_path = os_str_to_wide_null(path.as_os_str());
            // SAFETY: The path is null-terminated and lives through CreateFileW.
            // A successful handle is owned and closed by MutationPermit.
            let handle = unsafe {
                CreateFileW(
                    PCWSTR::from_raw(wide_path.as_mut_ptr()),
                    GENERIC_READ.0,
                    if exclusive {
                        FILE_SHARE_NONE
                    } else {
                        FILE_SHARE_READ
                    },
                    None,
                    OPEN_ALWAYS,
                    FILE_ATTRIBUTE_NORMAL,
                    None,
                )
            }
            .map_err(|error| {
                if error.code() == windows::Win32::Foundation::ERROR_SHARING_VIOLATION.into() {
                    if exclusive {
                        WincentError::SystemError(format!(
                            "Quick Access category {category:?} is already locked or has an active mutation"
                        ))
                    } else {
                        WincentError::QuickAccessLocked { qa_type: category }
                    }
                } else {
                    WincentError::SystemError(format!(
                        "Failed to open Quick Access admission file '{}': {error}",
                        path.display()
                    ))
                }
            })?;
            gates.push(Self { handle });
        }
        Ok(gates)
    }
}

// SAFETY: MutationPermit uniquely owns an OS file handle and closes it once in Drop.
// Windows file handles can be closed on a different thread.
unsafe impl Send for MutationPermit {}

impl Drop for MutationPermit {
    fn drop(&mut self) {
        // SAFETY: This owned handle is closed exactly once.
        unsafe {
            let _ = CloseHandle(self.handle);
        }
    }
}

impl QuickAccessLock {
    pub(crate) fn lock_target(target: QuickAccessLockTarget) -> WincentResult<Self> {
        let recent_folder = PathBuf::from(get_windows_recent_folder()?);
        Self::lock_target_in_recent_folder(target, recent_folder)
    }

    pub(crate) fn lock_target_in_recent_folder(
        target: QuickAccessLockTarget,
        recent_folder: PathBuf,
    ) -> WincentResult<Self> {
        let gates = MutationPermit::acquire_in_folder(target, &recent_folder, true)?;
        let automatic_destinations = recent_folder.join("AutomaticDestinations");
        let recent_files_path = automatic_destinations.join(RECENT_FILES_AUTOMATIC_DESTINATION);
        let frequent_folders_path =
            automatic_destinations.join(FREQUENT_FOLDERS_AUTOMATIC_DESTINATION);

        let mut locks = Vec::new();
        if matches!(
            target,
            QuickAccessLockTarget::RecentFiles | QuickAccessLockTarget::All
        ) {
            locks.push(LockedFile::open(&recent_files_path)?);
        }
        if matches!(
            target,
            QuickAccessLockTarget::FrequentFolders | QuickAccessLockTarget::All
        ) {
            locks.push(LockedFile::open(&frequent_folders_path)?);
        }
        let initial_lnk_paths = recent_lnk_paths(&recent_folder)?;

        Ok(Self {
            recent_folder,
            initial_lnk_paths,
            _locks: locks,
            _gates: gates,
        })
    }

    /// Cleans up new Recent shortcuts and releases all held handles.
    pub(crate) fn unlock_and_cleanup(self) -> WincentResult<UnlockReport> {
        let current_lnk_paths = recent_lnk_paths(&self.recent_folder)?;
        let new_lnk_paths = diff_new_paths(&self.initial_lnk_paths, &current_lnk_paths);
        Ok(cleanup_shortcuts(new_lnk_paths))
    }
}

fn cleanup_shortcuts(paths: Vec<PathBuf>) -> UnlockReport {
    cleanup_shortcuts_with(paths, |path| fs::remove_file(path))
}

fn cleanup_shortcuts_with<F>(paths: Vec<PathBuf>, mut remove: F) -> UnlockReport
where
    F: FnMut(&Path) -> io::Result<()>,
{
    let mut cleaned_shortcuts = Vec::new();
    let mut cleanup_failures = Vec::new();

    for path in paths {
        match remove(&path) {
            Ok(()) => cleaned_shortcuts.push(path),
            Err(error) if error.kind() == io::ErrorKind::NotFound => {
                cleaned_shortcuts.push(path);
            }
            Err(error) => cleanup_failures.push(ShortcutCleanupFailure { path, error }),
        }
    }

    UnlockReport {
        cleaned_shortcuts,
        cleanup_failures,
    }
}

struct LockedFile {
    handle: HANDLE,
}

// SAFETY: The handle is an owned process-wide file handle returned by CreateFileW.
// LockedFile has exclusive ownership of the handle, exposes no borrowed or
// thread-affine state, and closes the handle exactly once in Drop. Windows file
// handles may be closed from a thread other than the one that opened them.
unsafe impl Send for LockedFile {}

impl LockedFile {
    fn open(path: &Path) -> WincentResult<Self> {
        let mut wide_path = os_str_to_wide_null(path.as_os_str());
        // SAFETY: `wide_path` is null-terminated and lives until CreateFileW
        // returns. The returned HANDLE is owned by LockedFile and closed in Drop.
        let handle = unsafe {
            CreateFileW(
                PCWSTR::from_raw(wide_path.as_mut_ptr()),
                GENERIC_READ.0,
                FILE_SHARE_READ,
                None,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL,
                None,
            )
        }
        .map_err(|err| {
            WincentError::SystemError(format!("Failed to lock '{}': {err}", path.display()))
        })?;

        Ok(Self { handle })
    }
}

impl Drop for LockedFile {
    fn drop(&mut self) {
        // SAFETY: `handle` was returned by CreateFileW and is closed exactly
        // once when LockedFile is dropped.
        unsafe {
            let _ = CloseHandle(self.handle);
        }
    }
}

fn diff_new_paths(initial: &[PathBuf], current: &[PathBuf]) -> Vec<PathBuf> {
    current
        .iter()
        .filter(|path| {
            !initial
                .iter()
                .any(|initial_path| paths_equal_path(initial_path, path))
        })
        .cloned()
        .collect()
}

fn paths_equal_path(left: &Path, right: &Path) -> bool {
    paths_equal(&left.to_string_lossy(), &right.to_string_lossy())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn quick_access_lock_is_send_but_not_sync() {
        fn assert_send<T: Send>() {}
        trait AmbiguousIfSync<Marker> {}
        impl<T: ?Sized> AmbiguousIfSync<()> for T {}
        struct SyncMarker;
        impl<T: ?Sized + Sync> AmbiguousIfSync<SyncMarker> for T {}
        fn assert_not_sync<T, Marker>()
        where
            T: ?Sized + AmbiguousIfSync<Marker>,
        {
        }

        assert_send::<QuickAccessLock>();
        assert_not_sync::<QuickAccessLock, _>();
    }

    #[test]
    fn lock_target_variants_are_distinct() {
        assert_ne!(
            QuickAccessLockTarget::RecentFiles,
            QuickAccessLockTarget::FrequentFolders
        );
        assert_ne!(
            QuickAccessLockTarget::RecentFiles,
            QuickAccessLockTarget::All
        );
        assert_ne!(
            QuickAccessLockTarget::FrequentFolders,
            QuickAccessLockTarget::All
        );
    }

    #[test]
    fn diff_new_paths_reports_only_paths_absent_from_initial_snapshot() {
        let initial = vec![PathBuf::from("a.lnk"), PathBuf::from("b.lnk")];
        let current = vec![
            PathBuf::from("a.lnk"),
            PathBuf::from("b.lnk"),
            PathBuf::from("c.lnk"),
        ];

        assert_eq!(
            diff_new_paths(&initial, &current),
            vec![PathBuf::from("c.lnk")]
        );
    }

    #[test]
    fn diff_new_paths_uses_windows_path_comparison() {
        let initial = vec![PathBuf::from("C:\\Users\\Me\\Recent\\Example.LNK")];
        let current = vec![
            PathBuf::from("c:/users/me/recent/example.lnk"),
            PathBuf::from("C:\\Users\\Me\\Recent\\New.lnk"),
        ];

        assert_eq!(
            diff_new_paths(&initial, &current),
            vec![PathBuf::from("C:\\Users\\Me\\Recent\\New.lnk")]
        );
    }

    #[test]
    fn cleanup_report_classifies_removed_absent_and_failed_paths_without_file_system_access() {
        let removed = PathBuf::from("removed.lnk");
        let absent = PathBuf::from("absent.lnk");
        let denied = PathBuf::from("denied.lnk");

        let report = cleanup_shortcuts_with(
            vec![removed.clone(), absent.clone(), denied.clone()],
            |path| {
                if path == absent {
                    Err(io::Error::from(io::ErrorKind::NotFound))
                } else if path == denied {
                    Err(io::Error::from(io::ErrorKind::PermissionDenied))
                } else {
                    Ok(())
                }
            },
        );

        assert_eq!(report.cleaned_shortcuts(), &[removed, absent]);
        assert_eq!(report.cleanup_failures().len(), 1);
        assert_eq!(report.cleanup_failures()[0].path(), denied);
        assert!(!report.is_complete());
    }

    #[test]
    fn cleanup_report_is_complete_when_every_path_is_already_absent() {
        let path = PathBuf::from("already-absent.lnk");

        let report = cleanup_shortcuts_with(vec![path.clone()], |_| {
            Err(io::Error::from(io::ErrorKind::NotFound))
        });

        assert_eq!(report.cleaned_shortcuts(), &[path]);
        assert!(report.cleanup_failures().is_empty());
        assert!(report.is_complete());
    }
}
