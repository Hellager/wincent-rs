//! Locks Explorer Quick Access backing files and optionally cleans new shortcuts.
//!
//! While a target is locked, Explorer cannot write the corresponding automatic
//! destination file, so adding or removing items in that Quick Access category
//! is blocked or fails until the [`QuickAccessLock`] is dropped or unlocked.
//! Windows may still create Recent-folder `.lnk` shortcuts during the lock;
//! [`QuickAccessUnlockOptions::cleanup_new_recent_links`] can delete those on
//! unlock.

use crate::error::WincentError;
use crate::recent_links::recent_lnk_paths;
use crate::utils::{get_windows_recent_folder, os_str_to_wide_null, paths_equal};
use crate::WincentResult;
use std::fs;
use std::path::{Path, PathBuf};
use windows::core::PCWSTR;
use windows::Win32::Foundation::{CloseHandle, GENERIC_READ, HANDLE};
use windows::Win32::Storage::FileSystem::{
    CreateFileW, FILE_ATTRIBUTE_NORMAL, FILE_SHARE_READ, OPEN_EXISTING,
};

pub(crate) const RECENT_FILES_AUTOMATIC_DESTINATION: &str =
    "5f7b5f1e01b83767.automaticDestinations-ms";
pub(crate) const FREQUENT_FOLDERS_AUTOMATIC_DESTINATION: &str =
    "f01b4d95cf55d32a.automaticDestinations-ms";

/// Quick Access backing file set to lock.
///
/// Locking a target prevents Explorer from adding or removing items in the
/// corresponding Quick Access category while the lock is held.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum QuickAccessLockTarget {
    /// Lock the Recent Files automatic destination file.
    ///
    /// Recent Files additions and removals are blocked while this lock is held.
    RecentFiles,
    /// Lock the Frequent Folders automatic destination file.
    ///
    /// Frequent Folders additions and removals are blocked while this lock is held.
    FrequentFolders,
    /// Lock both Recent Files and Frequent Folders automatic destination files.
    ///
    /// Additions and removals in both categories are blocked while this lock is held.
    All,
}

/// Options used when unlocking Quick Access backing files.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct QuickAccessUnlockOptions {
    cleanup_new_recent_links: bool,
}

impl QuickAccessUnlockOptions {
    /// Creates default unlock options.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether `.lnk` files created after locking should be deleted during unlock.
    #[must_use]
    pub fn cleanup_new_recent_links_enabled(&self) -> bool {
        self.cleanup_new_recent_links
    }

    /// Sets whether `.lnk` files created after locking should be deleted during unlock.
    #[must_use]
    pub fn with_cleanup_new_recent_links(mut self, cleanup_new_recent_links: bool) -> Self {
        self.cleanup_new_recent_links = cleanup_new_recent_links;
        self
    }

    /// Deletes `.lnk` files that appeared in the Windows Recent folder after locking.
    #[must_use]
    pub fn cleanup_new_recent_links(self) -> Self {
        self.with_cleanup_new_recent_links(true)
    }
}

/// Report returned when unlocking Quick Access backing files.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QuickAccessUnlockReport {
    recent_folder: PathBuf,
    initial_lnk_paths: Vec<PathBuf>,
    current_lnk_paths: Vec<PathBuf>,
    new_lnk_paths: Vec<PathBuf>,
    disappeared_lnk_paths: Vec<PathBuf>,
    deleted_lnk_paths: Vec<PathBuf>,
    failed_lnk_deletions: Vec<QuickAccessUnlockFailure>,
}

/// A `.lnk` file that could not be deleted while unlocking.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QuickAccessUnlockFailure {
    path: PathBuf,
    error: String,
}

impl QuickAccessUnlockFailure {
    #[must_use]
    pub fn path(&self) -> &Path {
        &self.path
    }

    #[must_use]
    pub fn error(&self) -> &str {
        &self.error
    }
}

impl QuickAccessUnlockReport {
    #[must_use]
    pub fn recent_folder(&self) -> &Path {
        &self.recent_folder
    }

    #[must_use]
    pub fn initial_lnk_paths(&self) -> &[PathBuf] {
        &self.initial_lnk_paths
    }

    #[must_use]
    pub fn current_lnk_paths(&self) -> &[PathBuf] {
        &self.current_lnk_paths
    }

    #[must_use]
    pub fn new_lnk_paths(&self) -> &[PathBuf] {
        &self.new_lnk_paths
    }

    #[must_use]
    pub fn disappeared_lnk_paths(&self) -> &[PathBuf] {
        &self.disappeared_lnk_paths
    }

    #[must_use]
    pub fn deleted_lnk_paths(&self) -> &[PathBuf] {
        &self.deleted_lnk_paths
    }

    #[must_use]
    pub fn failed_lnk_deletions(&self) -> &[QuickAccessUnlockFailure] {
        &self.failed_lnk_deletions
    }
}

/// Guard that holds locks on Explorer Quick Access backing files.
///
/// While this guard is alive, Explorer cannot update the locked automatic
/// destination files. Add or remove operations for the locked target are
/// expected to fail or be ignored by Explorer until the guard is dropped or
/// [`QuickAccessLock::unlock`] is called.
///
/// This guard owns process-wide Windows file handles and may be moved to another
/// thread. It does not provide shared concurrent access; transfer ownership when
/// another thread is responsible for unlocking or dropping it.
///
/// Dropping the guard releases the backing-file locks without producing an
/// unlock cleanup report. [`QuickAccessLock::unlock`] consumes the guard and may
/// return an error, but its backing-file handles are still released during
/// destruction. Cleanup of newly-created Recent shortcuts is best-effort and is
/// separate from releasing the backing-file locks.
pub struct QuickAccessLock {
    target: QuickAccessLockTarget,
    recent_folder: PathBuf,
    initial_lnk_paths: Vec<PathBuf>,
    /// Held only for their Drop behavior; keeping these handles open is the lock.
    locks: Vec<LockedFile>,
}

impl QuickAccessLock {
    pub(crate) fn lock() -> WincentResult<Self> {
        Self::lock_target(QuickAccessLockTarget::All)
    }

    pub(crate) fn lock_target(target: QuickAccessLockTarget) -> WincentResult<Self> {
        let recent_folder = PathBuf::from(get_windows_recent_folder()?);
        Self::lock_target_in_recent_folder(target, recent_folder)
    }

    fn lock_target_in_recent_folder(
        target: QuickAccessLockTarget,
        recent_folder: PathBuf,
    ) -> WincentResult<Self> {
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
            target,
            recent_folder,
            initial_lnk_paths,
            locks,
        })
    }

    #[must_use]
    pub fn target(&self) -> QuickAccessLockTarget {
        self.target
    }

    #[must_use]
    pub fn recent_folder(&self) -> &Path {
        &self.recent_folder
    }

    #[must_use]
    pub fn initial_lnk_paths(&self) -> &[PathBuf] {
        &self.initial_lnk_paths
    }

    #[must_use]
    pub fn locked_file_count(&self) -> usize {
        self.locks.len()
    }

    /// Unlocks the backing files and optionally deletes newly-created Recent shortcuts.
    ///
    /// When cleanup is enabled, deletion is best-effort: failures do not abort
    /// the unlock and are returned in [`QuickAccessUnlockReport::failed_lnk_deletions`].
    /// The guard is consumed even when snapshot enumeration returns an error, and
    /// its backing-file handles are released during destruction in all cases.
    pub fn unlock(
        self,
        options: QuickAccessUnlockOptions,
    ) -> WincentResult<QuickAccessUnlockReport> {
        let current_lnk_paths = recent_lnk_paths(&self.recent_folder)?;
        let new_lnk_paths = diff_new_paths(&self.initial_lnk_paths, &current_lnk_paths);
        let disappeared_lnk_paths = diff_new_paths(&current_lnk_paths, &self.initial_lnk_paths);
        let mut deleted_lnk_paths = Vec::new();
        let mut failed_lnk_deletions = Vec::new();

        if options.cleanup_new_recent_links_enabled() {
            for path in &new_lnk_paths {
                match fs::remove_file(path) {
                    Ok(()) => deleted_lnk_paths.push(path.clone()),
                    Err(error) => failed_lnk_deletions.push(QuickAccessUnlockFailure {
                        path: path.clone(),
                        error: error.to_string(),
                    }),
                }
            }
        }

        Ok(QuickAccessUnlockReport {
            recent_folder: self.recent_folder.clone(),
            initial_lnk_paths: self.initial_lnk_paths.clone(),
            current_lnk_paths,
            new_lnk_paths,
            disappeared_lnk_paths,
            deleted_lnk_paths,
            failed_lnk_deletions,
        })
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
    use windows::Win32::Storage::FileSystem::FILE_SHARE_NONE;

    fn create_recent_fixture() -> WincentResult<(tempfile::TempDir, PathBuf, PathBuf, PathBuf)> {
        let temp_dir = tempfile::tempdir().map_err(WincentError::Io)?;
        let recent_folder = temp_dir.path().to_path_buf();
        let automatic_destinations = recent_folder.join("AutomaticDestinations");
        fs::create_dir(&automatic_destinations).map_err(WincentError::Io)?;

        let recent_file = automatic_destinations.join(RECENT_FILES_AUTOMATIC_DESTINATION);
        let frequent_file = automatic_destinations.join(FREQUENT_FOLDERS_AUTOMATIC_DESTINATION);
        fs::write(&recent_file, b"recent").map_err(WincentError::Io)?;
        fs::write(&frequent_file, b"frequent").map_err(WincentError::Io)?;

        Ok((temp_dir, recent_folder, recent_file, frequent_file))
    }

    fn can_open_exclusively(path: &Path) -> bool {
        let mut wide_path = os_str_to_wide_null(path.as_os_str());
        // SAFETY: wide_path is null-terminated and remains alive for the call.
        // A successful handle is owned locally and closed exactly once below.
        let handle = unsafe {
            CreateFileW(
                PCWSTR::from_raw(wide_path.as_mut_ptr()),
                GENERIC_READ.0,
                FILE_SHARE_NONE,
                None,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL,
                None,
            )
        };

        match handle {
            Ok(handle) => {
                // SAFETY: handle was returned successfully by CreateFileW and
                // has not been closed or transferred.
                unsafe {
                    let _ = CloseHandle(handle);
                }
                true
            }
            Err(_) => false,
        }
    }

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
        // Marker inference becomes ambiguous if QuickAccessLock implements Sync.
        assert_not_sync::<QuickAccessLock, _>();
    }

    #[test]
    fn lock_can_be_unlocked_on_another_thread() -> WincentResult<()> {
        let (_temp_dir, recent_folder, recent_file, _) = create_recent_fixture()?;
        let lock = QuickAccessLock::lock_target_in_recent_folder(
            QuickAccessLockTarget::RecentFiles,
            recent_folder,
        )?;
        assert!(!can_open_exclusively(&recent_file));

        std::thread::spawn(move || lock.unlock(QuickAccessUnlockOptions::new()))
            .join()
            .expect("unlock worker panicked")?;

        assert!(can_open_exclusively(&recent_file));
        Ok(())
    }

    #[test]
    fn lock_can_be_dropped_on_another_thread() -> WincentResult<()> {
        let (_temp_dir, recent_folder, recent_file, _) = create_recent_fixture()?;
        let lock = QuickAccessLock::lock_target_in_recent_folder(
            QuickAccessLockTarget::RecentFiles,
            recent_folder,
        )?;
        assert!(!can_open_exclusively(&recent_file));

        std::thread::spawn(move || drop(lock))
            .join()
            .expect("drop worker panicked");

        assert!(can_open_exclusively(&recent_file));
        Ok(())
    }

    #[test]
    fn partial_lock_failure_releases_opened_handles() -> WincentResult<()> {
        let (_temp_dir, recent_folder, recent_file, frequent_file) = create_recent_fixture()?;
        fs::remove_file(frequent_file).map_err(WincentError::Io)?;

        assert!(QuickAccessLock::lock_target_in_recent_folder(
            QuickAccessLockTarget::All,
            recent_folder,
        )
        .is_err());
        assert!(can_open_exclusively(&recent_file));
        Ok(())
    }

    #[test]
    fn unlock_snapshot_failure_releases_opened_handles() -> WincentResult<()> {
        let (_temp_dir, _recent_folder, recent_file, _) = create_recent_fixture()?;
        let lock = QuickAccessLock {
            target: QuickAccessLockTarget::RecentFiles,
            recent_folder: recent_file.join("not-a-directory"),
            initial_lnk_paths: Vec::new(),
            locks: vec![LockedFile::open(&recent_file)?],
        };
        assert!(!can_open_exclusively(&recent_file));

        assert!(lock.unlock(QuickAccessUnlockOptions::new()).is_err());

        assert!(can_open_exclusively(&recent_file));
        Ok(())
    }

    #[test]
    fn cleanup_failure_is_reported_and_releases_backing_handle() -> WincentResult<()> {
        let (_temp_dir, recent_folder, recent_file, _) = create_recent_fixture()?;
        let lock = QuickAccessLock::lock_target_in_recent_folder(
            QuickAccessLockTarget::RecentFiles,
            recent_folder.clone(),
        )?;
        let new_link = recent_folder.join("new.lnk");
        fs::write(&new_link, b"new").map_err(WincentError::Io)?;
        let deletion_blocker = LockedFile::open(&new_link)?;

        let report = lock.unlock(QuickAccessUnlockOptions::new().cleanup_new_recent_links())?;

        assert_eq!(report.failed_lnk_deletions().len(), 1);
        assert_eq!(report.failed_lnk_deletions()[0].path(), new_link);
        assert!(can_open_exclusively(&recent_file));
        drop(deletion_blocker);
        Ok(())
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
    fn unlock_options_default_does_not_cleanup_new_links() {
        assert!(!QuickAccessUnlockOptions::new().cleanup_new_recent_links_enabled());
    }

    #[test]
    fn unlock_options_builder_enables_cleanup_new_links() {
        let options = QuickAccessUnlockOptions::new().cleanup_new_recent_links();

        assert!(options.cleanup_new_recent_links_enabled());
        assert!(!options
            .with_cleanup_new_recent_links(false)
            .cleanup_new_recent_links_enabled());
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
    fn reversed_diff_reports_disappeared_paths() {
        let initial = vec![
            PathBuf::from("a.lnk"),
            PathBuf::from("b.lnk"),
            PathBuf::from("gone.lnk"),
        ];
        let current = vec![PathBuf::from("a.lnk"), PathBuf::from("b.lnk")];

        assert_eq!(
            diff_new_paths(&current, &initial),
            vec![PathBuf::from("gone.lnk")]
        );
    }

    #[test]
    fn reversed_diff_uses_windows_path_comparison() {
        let initial = vec![
            PathBuf::from("C:\\Users\\Me\\Recent\\Example.LNK"),
            PathBuf::from("C:\\Users\\Me\\Recent\\Gone.lnk"),
        ];
        let current = vec![PathBuf::from("c:/users/me/recent/example.lnk")];

        assert_eq!(
            diff_new_paths(&current, &initial),
            vec![PathBuf::from("C:\\Users\\Me\\Recent\\Gone.lnk")]
        );
    }

    #[test]
    fn unlock_failure_accessors_expose_path_and_error() {
        let failure = QuickAccessUnlockFailure {
            path: PathBuf::from("new.lnk"),
            error: "access denied".to_string(),
        };

        assert_eq!(failure.path(), Path::new("new.lnk"));
        assert_eq!(failure.error(), "access denied");
    }

    #[test]
    fn unlock_report_exposes_failed_lnk_deletions() {
        let failure = QuickAccessUnlockFailure {
            path: PathBuf::from("new.lnk"),
            error: "access denied".to_string(),
        };
        let report = QuickAccessUnlockReport {
            recent_folder: PathBuf::from("Recent"),
            initial_lnk_paths: Vec::new(),
            current_lnk_paths: vec![PathBuf::from("new.lnk")],
            new_lnk_paths: vec![PathBuf::from("new.lnk")],
            disappeared_lnk_paths: vec![PathBuf::from("gone.lnk")],
            deleted_lnk_paths: Vec::new(),
            failed_lnk_deletions: vec![failure.clone()],
        };

        assert_eq!(report.disappeared_lnk_paths(), &[PathBuf::from("gone.lnk")]);
        assert_eq!(report.failed_lnk_deletions(), &[failure]);
    }

    #[test]
    fn unlock_reports_disappeared_paths_without_cleanup() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir().map_err(WincentError::Io)?;
        let recent_folder = temp_dir.path().to_path_buf();
        let automatic_destinations = recent_folder.join("AutomaticDestinations");
        fs::create_dir(&automatic_destinations).map_err(WincentError::Io)?;

        let recent_file = automatic_destinations.join(RECENT_FILES_AUTOMATIC_DESTINATION);
        fs::write(&recent_file, b"lock").map_err(WincentError::Io)?;
        let existing = recent_folder.join("existing.lnk");
        let gone = recent_folder.join("gone.lnk");
        let new = recent_folder.join("new.lnk");
        fs::write(&existing, b"existing").map_err(WincentError::Io)?;
        fs::write(&gone, b"gone").map_err(WincentError::Io)?;

        let lock = QuickAccessLock::lock_target_in_recent_folder(
            QuickAccessLockTarget::RecentFiles,
            recent_folder.clone(),
        )?;
        fs::remove_file(&gone).map_err(WincentError::Io)?;
        fs::write(&new, b"new").map_err(WincentError::Io)?;

        let report = lock.unlock(QuickAccessUnlockOptions::new())?;

        assert_eq!(report.new_lnk_paths(), std::slice::from_ref(&new));
        assert_eq!(report.disappeared_lnk_paths(), &[gone]);
        assert!(report.deleted_lnk_paths().is_empty());
        assert!(new.exists());
        Ok(())
    }

    #[test]
    fn unlock_cleanup_deleted_paths_are_orthogonal_to_disappeared_paths() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir().map_err(WincentError::Io)?;
        let recent_folder = temp_dir.path().to_path_buf();
        let automatic_destinations = recent_folder.join("AutomaticDestinations");
        fs::create_dir(&automatic_destinations).map_err(WincentError::Io)?;

        let recent_file = automatic_destinations.join(RECENT_FILES_AUTOMATIC_DESTINATION);
        fs::write(&recent_file, b"lock").map_err(WincentError::Io)?;
        let existing = recent_folder.join("existing.lnk");
        let gone = recent_folder.join("gone.lnk");
        let new = recent_folder.join("new.lnk");
        fs::write(&existing, b"existing").map_err(WincentError::Io)?;
        fs::write(&gone, b"gone").map_err(WincentError::Io)?;

        let lock = QuickAccessLock::lock_target_in_recent_folder(
            QuickAccessLockTarget::RecentFiles,
            recent_folder,
        )?;
        fs::remove_file(&gone).map_err(WincentError::Io)?;
        fs::write(&new, b"new").map_err(WincentError::Io)?;

        let report = lock.unlock(QuickAccessUnlockOptions::new().cleanup_new_recent_links())?;

        assert_eq!(report.disappeared_lnk_paths(), std::slice::from_ref(&gone));
        assert_eq!(report.deleted_lnk_paths(), std::slice::from_ref(&new));
        assert!(!new.exists());
        Ok(())
    }
}
