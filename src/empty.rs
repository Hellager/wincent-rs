//! Windows Quick Access cleanup operations
//!
//! Provides unified interface for clearing Windows Quick Access items including:
//! - Recent files
//! - Frequent folders (Explorer backing file and optional explicit pin cleanup)
//! - Complete Quick Access history
//!
//! Implements multiple cleanup strategies with fallback mechanisms
//!
//! # Key Functionality
//! - Clear recent files using Windows Shell API
//! - Remove frequent folders by deleting Explorer's backing file
//! - Optionally clear visible pinned folders through Explorer shell verbs
//! - Full Quick Access reset capabilities
//! - Atomic operations with proper cleanup sequencing

use crate::{
    backend::{QuickAccessBackend, SystemQuickAccessBackend},
    error::WincentError,
    utils::get_windows_recent_folder,
    QuickAccess, WincentResult,
};
use std::time::Duration;
use windows::Win32::UI::Shell::SHAddToRecentDocs;

/// SHARD_PATHW - with a null data pointer, clears all recent documents.
const SHARD_PATHW: u32 = 0x0000_0003;
const EMPTY_PINNED_FOLDERS_TIMEOUT: Duration = Duration::from_secs(10);

/// Options for clearing Quick Access items.
///
/// Clearing Frequent Folders deletes Explorer's Frequent Folders backing file.
/// On current Windows builds, Explorer may rebuild default folder pins and may
/// remove user-pinned folders even when [`EmptyOptions::also_pinned_folders`]
/// is `false`.
///
/// By default, wincent does not additionally invoke Explorer's unpin verb for
/// visible pinned folders. Use [`EmptyOptions::remove_pinned_folders`] when the
/// caller explicitly wants that extra unpin step too.
///
/// # Examples
///
/// ```rust
/// use wincent::EmptyOptions;
///
/// let options = EmptyOptions::new()
///     .remove_pinned_folders()
///     .refresh_explorer();
///
/// assert!(options.also_pinned_folders());
/// assert!(options.refresh_explorer_enabled());
/// ```
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EmptyOptions {
    /// Also attempt to explicitly unpin visible folders from Quick Access.
    also_pinned_folders: bool,
    /// Refresh open Explorer windows after a successful clear.
    refresh_explorer: bool,
}

impl EmptyOptions {
    /// Creates default clear options.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether visible pinned folders should also be explicitly unpinned.
    #[must_use]
    pub fn also_pinned_folders(&self) -> bool {
        self.also_pinned_folders
    }

    /// Sets whether visible pinned folders should also be explicitly unpinned.
    #[must_use]
    pub fn with_also_pinned_folders(mut self, also_pinned_folders: bool) -> Self {
        self.also_pinned_folders = also_pinned_folders;
        self
    }

    /// Also explicitly unpins visible folders when clearing Frequent Folders or all items.
    #[must_use]
    pub fn remove_pinned_folders(self) -> Self {
        self.with_also_pinned_folders(true)
    }

    /// Whether open Explorer windows should be refreshed after a successful clear.
    #[must_use]
    pub fn refresh_explorer_enabled(&self) -> bool {
        self.refresh_explorer
    }

    /// Sets whether open Explorer windows should be refreshed after a successful clear.
    #[must_use]
    pub fn with_refresh_explorer(mut self, enabled: bool) -> Self {
        self.refresh_explorer = enabled;
        self
    }

    /// Refreshes open Explorer windows after a successful clear.
    #[must_use]
    pub fn refresh_explorer(self) -> Self {
        self.with_refresh_explorer(true)
    }
}

/// Clears the Windows Recent Files list via a dedicated STA thread.
///
/// `timeout` limits how long the caller waits; the underlying Shell operation
/// may still complete after the timeout elapses.
pub(crate) fn empty_recent_files_with_api(timeout: Duration) -> WincentResult<()> {
    crate::com_thread::run_on_sta_thread(
        || {
            // SAFETY: SHARD_PATHW with a null data pointer is a documented
            // Windows API pattern for clearing the recent documents list.
            unsafe { SHAddToRecentDocs(SHARD_PATHW, None) };
            Ok(())
        },
        timeout,
    )
}

/// Clears user folders from Quick Access by removing the Windows jump list file.
pub(crate) fn empty_user_folders_with_jumplist_file() -> WincentResult<()> {
    let recent_folder = get_windows_recent_folder()?;

    let jumplist_file = std::path::Path::new(&recent_folder)
        .join("AutomaticDestinations")
        .join("f01b4d95cf55d32a.automaticDestinations-ms");

    if jumplist_file.exists() {
        std::fs::remove_file(&jumplist_file).map_err(WincentError::Io)?;
    }

    Ok(())
}

fn empty_pinned_folders_from_snapshot<F>(
    paths: &[String],
    mut remove_folder: F,
) -> WincentResult<()>
where
    F: FnMut(&str) -> WincentResult<()>,
{
    let mut first_error = None;

    for path in paths {
        match remove_folder(path) {
            Ok(()) | Err(WincentError::NotInQuickAccess { .. }) => {}
            Err(error) if first_error.is_none() => first_error = Some(error),
            Err(_) => {}
        }
    }

    if let Some(error) = first_error {
        Err(error)
    } else {
        Ok(())
    }
}

/// Attempts to unpin all pinned folders from Quick Access.
///
/// Snapshots every item currently in the Frequent Folders namespace and removes
/// each one through the single-item native mutation path. That path tries
/// `unpinfromhome` first and falls back to the Windows 11 `pintohome` toggle
/// workaround when needed.
#[allow(dead_code)]
pub(crate) fn empty_pinned_folders() -> WincentResult<()> {
    let backend = SystemQuickAccessBackend;
    let paths = backend.get_items(QuickAccess::FrequentFolders, EMPTY_PINNED_FOLDERS_TIMEOUT)?;
    empty_pinned_folders_from_snapshot(&paths, |path| {
        backend.remove_frequent_folder(path, EMPTY_PINNED_FOLDERS_TIMEOUT)
    })
}

/****************************************************** Empty Quick Access ******************************************************/

#[allow(dead_code)]
pub(crate) fn empty_frequent_folders(also_pinned_folders: bool) -> WincentResult<()> {
    let backend = SystemQuickAccessBackend;
    empty_frequent_folders_with_backend(also_pinned_folders, EMPTY_PINNED_FOLDERS_TIMEOUT, &backend)
}

fn empty_frequent_folders_with_backend(
    also_pinned_folders: bool,
    timeout: Duration,
    backend: &dyn QuickAccessBackend,
) -> WincentResult<()> {
    let pinned_snapshot = if also_pinned_folders {
        Some(backend.get_items(QuickAccess::FrequentFolders, timeout)?)
    } else {
        None
    };

    backend.clear_frequent_folders_jumplist()?;
    if let Some(paths) = pinned_snapshot {
        if let Err(source) = empty_pinned_folders_from_snapshot(&paths, |path| {
            backend.remove_frequent_folder(path, timeout)
        }) {
            return Err(partial_frequent_folders_error(source));
        }
    }
    Ok(())
}

fn partial_frequent_folders_error(source: WincentError) -> WincentError {
    WincentError::PartialEmpty {
        recent_files_cleared: false,
        // This means the user-visited frequent-folders jump list was cleared.
        // It does not guarantee pinned folders were removed.
        frequent_folders_cleared: true,
        source: Box::new(source),
    }
}

fn frequent_folders_cleared_from_error(error: &WincentError) -> bool {
    match error {
        WincentError::PartialEmpty {
            frequent_folders_cleared,
            ..
        } => *frequent_folders_cleared,
        _ => false,
    }
}

/// Clears items from Windows Quick Access.
///
/// Always clears the Recent Files list and deletes Explorer's Frequent Folders
/// backing file. Deleting that backing file can cause Explorer to rebuild
/// default folder pins and remove user-pinned folders even when
/// `also_pinned_folders` is `false`.
///
/// Pass `also_pinned_folders: true` to additionally invoke the native
/// single-item mutation path for every item snapshotted from the Frequent
/// Folders namespace.
///
/// # Parameters
///
/// - `also_pinned_folders` - when `true`, snapshots every item in the Frequent
///   Folders namespace before deleting the backing file, then removes those
///   items through the native single-item mutation path.
///   When `false`, wincent does not additionally invoke Explorer's unpin verb
///   for visible pinned folders.
///
/// # Returns
///
/// Returns `Ok(())` if all requested cleanup steps completed successfully.
///
/// # Example
///
/// ```ignore
/// use wincent::{empty::empty_quick_access, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Clear recent files and delete the Frequent Folders backing file.
///     empty_items(QuickAccess::All, EmptyOptions::new())?;
///
///     // Also explicitly unpin visible Frequent Folders items.
///     empty_items(
///         QuickAccess::All,
///         EmptyOptions::new().with_also_pinned_folders(true),
///     )?;
///
///     Ok(())
/// }
/// ```
fn empty_all_items_with_backend(
    options: EmptyOptions,
    timeout: Duration,
    backend: &dyn QuickAccessBackend,
) -> WincentResult<()> {
    if let Err(source) = backend.clear_recent_files(timeout) {
        return Err(WincentError::PartialEmpty {
            recent_files_cleared: false,
            frequent_folders_cleared: false,
            source: Box::new(source),
        });
    }

    if let Err(source) =
        empty_frequent_folders_with_backend(options.also_pinned_folders(), timeout, backend)
    {
        let frequent_folders_cleared = frequent_folders_cleared_from_error(&source);
        return Err(WincentError::PartialEmpty {
            recent_files_cleared: true,
            frequent_folders_cleared,
            source: Box::new(source),
        });
    }

    Ok(())
}

#[allow(dead_code)]
fn empty_all_items(options: EmptyOptions) -> WincentResult<()> {
    let backend = SystemQuickAccessBackend;
    empty_all_items_with_backend(options, EMPTY_PINNED_FOLDERS_TIMEOUT, &backend)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RefreshPolicy {
    PropagateFailure,
    BestEffort,
    Skip,
}

fn refresh_policy_for_result(result: &WincentResult<()>, refresh_explorer: bool) -> RefreshPolicy {
    if !refresh_explorer {
        return RefreshPolicy::Skip;
    }

    match result {
        Ok(()) => RefreshPolicy::PropagateFailure,
        Err(WincentError::PartialEmpty { .. }) => RefreshPolicy::BestEffort,
        Err(_) => RefreshPolicy::Skip,
    }
}

/// Clears items from a specific Quick Access category.
#[allow(dead_code)]
pub(crate) fn empty_items(qa_type: QuickAccess, options: EmptyOptions) -> WincentResult<()> {
    let backend = SystemQuickAccessBackend;
    empty_items_with_backend(qa_type, options, EMPTY_PINNED_FOLDERS_TIMEOUT, &backend)
}

pub(crate) fn empty_items_with_backend(
    qa_type: QuickAccess,
    options: EmptyOptions,
    timeout: Duration,
    backend: &dyn QuickAccessBackend,
) -> WincentResult<()> {
    let result = match qa_type {
        QuickAccess::RecentFiles => backend.clear_recent_files(timeout),
        QuickAccess::FrequentFolders => {
            empty_frequent_folders_with_backend(options.also_pinned_folders(), timeout, backend)
        }
        QuickAccess::All => empty_all_items_with_backend(options, timeout, backend),
    };

    match refresh_policy_for_result(&result, options.refresh_explorer_enabled()) {
        RefreshPolicy::PropagateFailure => backend.refresh_explorer()?,
        RefreshPolicy::BestEffort => {
            // Preserve the cleanup failure while still trying to show any
            // successfully cleared category in Explorer.
            let _ = backend.refresh_explorer();
        }
        RefreshPolicy::Skip => {}
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::handle::{
        add_file_to_recent_native, pin_frequent_folder, remove_from_recent_files,
        unpin_frequent_folder,
    };
    use crate::query::{get_frequent_folders, get_recent_files};
    use crate::test_utils::{cleanup_test_env, create_test_file, setup_test_env};
    use std::path::PathBuf;
    use std::sync::Mutex;
    use std::thread;
    use std::time::Duration;

    struct FakeEmptyBackend {
        items: Vec<String>,
        calls: Mutex<Vec<String>>,
    }

    impl FakeEmptyBackend {
        fn new(items: Vec<String>) -> Self {
            Self {
                items,
                calls: Mutex::new(Vec::new()),
            }
        }

        fn record(&self, call: impl Into<String>) {
            self.calls.lock().unwrap().push(call.into());
        }

        fn calls(&self) -> Vec<String> {
            self.calls.lock().unwrap().clone()
        }
    }

    impl crate::backend::QuickAccessBackend for FakeEmptyBackend {
        fn validate_path(
            &self,
            _path: &str,
            _expected: crate::utils::PathType,
        ) -> WincentResult<()> {
            Ok(())
        }

        fn get_items(
            &self,
            qa_type: QuickAccess,
            _timeout: Duration,
        ) -> WincentResult<Vec<String>> {
            self.record(format!("get_items:{qa_type:?}"));
            Ok(self.items.clone())
        }

        fn add_recent_file(&self, _path: &str, _timeout: Duration) -> WincentResult<()> {
            Ok(())
        }

        fn add_frequent_folder(&self, _path: &str, _timeout: Duration) -> WincentResult<()> {
            Ok(())
        }

        fn remove_recent_file(&self, _path: &str, _timeout: Duration) -> WincentResult<()> {
            Ok(())
        }

        fn remove_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()> {
            self.record(format!(
                "remove_frequent_folder:{path}:{}",
                timeout.as_secs()
            ));
            Ok(())
        }

        fn delete_recent_links_for_target(
            &self,
            _path: &str,
            _timeout: Duration,
        ) -> WincentResult<()> {
            Ok(())
        }

        fn delete_recent_files_backing_data(&self) -> WincentResult<()> {
            Ok(())
        }

        fn clear_recent_files(&self, _timeout: Duration) -> WincentResult<()> {
            self.record("clear_recent_files");
            Ok(())
        }

        fn clear_frequent_folders_jumplist(&self) -> WincentResult<()> {
            self.record("clear_frequent_folders_jumplist");
            Ok(())
        }

        fn refresh_explorer(&self) -> WincentResult<()> {
            self.record("refresh_explorer");
            Ok(())
        }
    }

    // -------------------------------------------------------------------------
    // Integration-test helpers
    // -------------------------------------------------------------------------

    fn path_to_str(path: &std::path::Path) -> WincentResult<&str> {
        path.to_str()
            .ok_or_else(|| WincentError::invalid_path(path, "Invalid path encoding"))
    }

    fn unique_millis() -> WincentResult<u128> {
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|duration| duration.as_millis())
            .map_err(|error| WincentError::SystemError(error.to_string()))
    }

    /// RAII guard that runs best-effort cleanup when dropped.
    ///
    /// Ensures `cleanup_test_env` and optional Quick Access rollbacks run even
    /// when a test panics or an assertion fails, preventing state leakage into
    /// subsequent tests.
    struct TestGuard {
        dir: PathBuf,
        /// Paths to unpin from Frequent Folders on drop (best effort).
        pinned: Vec<String>,
        /// Path to remove from Recent Files on drop (best effort).
        recent_file: Option<String>,
    }

    impl TestGuard {
        fn new(dir: PathBuf) -> Self {
            Self {
                dir,
                pinned: Vec::new(),
                recent_file: None,
            }
        }

        fn with_pinned(mut self, path: &str) -> Self {
            self.pinned.push(path.to_owned());
            self
        }

        fn with_recent_file(mut self, path: &str) -> Self {
            self.recent_file = Some(path.to_owned());
            self
        }
    }

    impl Drop for TestGuard {
        fn drop(&mut self) {
            // Best-effort unpin for every registered path.
            for p in &self.pinned {
                let _ = unpin_frequent_folder(p, Duration::from_secs(10));
            }
            // Best-effort recent-file removal - ignore errors.
            if let Some(ref p) = self.recent_file {
                let _ = remove_from_recent_files(p);
            }
            let _ = cleanup_test_env(&self.dir);
        }
    }

    #[test]
    fn refresh_policy_skip_when_no_refresh_explorer() {
        assert_eq!(
            refresh_policy_for_result(&Ok(()), false),
            RefreshPolicy::Skip
        );
    }

    #[test]
    fn refresh_policy_propagate_on_ok() {
        assert_eq!(
            refresh_policy_for_result(&Ok(()), true),
            RefreshPolicy::PropagateFailure
        );
    }

    #[test]
    fn refresh_policy_best_effort_on_partial_empty() {
        assert_eq!(
            refresh_policy_for_result(
                &Err(WincentError::PartialEmpty {
                    recent_files_cleared: true,
                    frequent_folders_cleared: false,
                    source: Box::new(WincentError::ScriptFailed("failed".to_string())),
                }),
                true,
            ),
            RefreshPolicy::BestEffort
        );
    }

    #[test]
    fn partial_frequent_folders_error_marks_jump_list_progress() {
        let error = partial_frequent_folders_error(WincentError::ScriptFailed("failed".into()));

        match error {
            WincentError::PartialEmpty {
                recent_files_cleared,
                frequent_folders_cleared,
                ..
            } => {
                assert!(!recent_files_cleared);
                assert!(frequent_folders_cleared);
            }
            other => panic!("Expected PartialEmpty, got: {other:?}"),
        }
    }

    #[test]
    fn frequent_folders_cleared_from_error_preserves_partial_progress() {
        let partial = partial_frequent_folders_error(WincentError::ScriptFailed("failed".into()));
        let non_partial = WincentError::ScriptFailed("failed".into());

        assert!(frequent_folders_cleared_from_error(&partial));
        assert!(!frequent_folders_cleared_from_error(&non_partial));
    }

    #[test]
    fn empty_frequent_folders_with_backend_snapshots_before_clearing_jumplist() -> WincentResult<()>
    {
        let backend =
            FakeEmptyBackend::new(vec!["C:\\FolderA".to_string(), "C:\\FolderB".to_string()]);

        empty_items_with_backend(
            QuickAccess::FrequentFolders,
            EmptyOptions::new().remove_pinned_folders(),
            Duration::from_secs(3),
            &backend,
        )?;

        assert_eq!(
            backend.calls(),
            vec![
                "get_items:FrequentFolders".to_string(),
                "clear_frequent_folders_jumplist".to_string(),
                "remove_frequent_folder:C:\\FolderA:3".to_string(),
                "remove_frequent_folder:C:\\FolderB:3".to_string(),
            ]
        );
        Ok(())
    }

    #[test]
    fn empty_frequent_folders_without_pinned_cleanup_does_not_snapshot() -> WincentResult<()> {
        let backend = FakeEmptyBackend::new(vec!["C:\\FolderA".to_string()]);

        empty_items_with_backend(
            QuickAccess::FrequentFolders,
            EmptyOptions::new(),
            Duration::from_secs(3),
            &backend,
        )?;

        assert_eq!(
            backend.calls(),
            vec!["clear_frequent_folders_jumplist".to_string()]
        );
        Ok(())
    }

    #[test]
    fn empty_pinned_snapshot_ignores_items_already_absent() -> WincentResult<()> {
        let paths = vec![
            "C:\\FolderA".to_string(),
            "C:\\FolderB".to_string(),
            "C:\\FolderC".to_string(),
        ];
        let mut calls = Vec::new();

        empty_pinned_folders_from_snapshot(&paths, |path| {
            calls.push(path.to_string());
            if path.ends_with("FolderB") {
                Err(WincentError::not_in_quick_access(
                    path,
                    QuickAccess::FrequentFolders,
                ))
            } else {
                Ok(())
            }
        })?;

        assert_eq!(calls, paths);
        Ok(())
    }

    #[test]
    fn empty_pinned_snapshot_returns_first_error_after_trying_all_items() {
        let paths = vec![
            "C:\\FolderA".to_string(),
            "C:\\FolderB".to_string(),
            "C:\\FolderC".to_string(),
        ];
        let mut calls = Vec::new();

        let error = empty_pinned_folders_from_snapshot(&paths, |path| {
            calls.push(path.to_string());
            if path.ends_with("FolderB") {
                Err(WincentError::SystemError("first failure".to_string()))
            } else if path.ends_with("FolderC") {
                Err(WincentError::SystemError("second failure".to_string()))
            } else {
                Ok(())
            }
        })
        .unwrap_err();

        assert_eq!(calls, paths);
        assert!(matches!(
            error,
            WincentError::SystemError(message) if message == "first failure"
        ));
    }

    #[test]
    fn refresh_policy_skip_on_non_partial_empty_error() {
        assert_eq!(
            refresh_policy_for_result(&Err(WincentError::ScriptFailed("failed".to_string())), true),
            RefreshPolicy::Skip
        );
    }

    /// Poll `get_recent_files()` until the given path appears or retries are
    /// exhausted. Returns `Ok(true)` when found.
    fn wait_until_recent_file_present(path: &str, max_retries: u32) -> WincentResult<bool> {
        for _ in 0..max_retries {
            let files = get_recent_files()?;
            if files.iter().any(|p| p == path) {
                return Ok(true);
            }
            thread::sleep(Duration::from_millis(500));
        }
        Ok(false)
    }

    /// Poll `get_recent_files()` until the list is empty or retries are
    /// exhausted. Returns `Ok(true)` when empty.
    fn wait_until_recent_files_empty(max_retries: u32) -> WincentResult<bool> {
        for _ in 0..max_retries {
            if get_recent_files()?.is_empty() {
                return Ok(true);
            }
            thread::sleep(Duration::from_millis(500));
        }
        Ok(false)
    }

    /// Poll `get_frequent_folders()` until `path` appears or retries are
    /// exhausted. Returns `Ok(true)` when found.
    fn wait_until_folder_present(path: &str, max_retries: u32) -> WincentResult<bool> {
        for _ in 0..max_retries {
            let folders = get_frequent_folders()?;
            if folders.iter().any(|p| p == path) {
                return Ok(true);
            }
            thread::sleep(Duration::from_millis(500));
        }
        Ok(false)
    }

    /// Poll `get_frequent_folders()` until `path` is absent or retries are
    /// exhausted. Returns `Ok(true)` when absent.
    fn wait_until_folder_absent(path: &str, max_retries: u32) -> WincentResult<bool> {
        for _ in 0..max_retries {
            let folders = get_frequent_folders()?;
            if !folders.iter().any(|p| p == path) {
                return Ok(true);
            }
            thread::sleep(Duration::from_millis(500));
        }
        Ok(false)
    }

    // -------------------------------------------------------------------------
    // COM correctness tests (no system-state modification)
    // -------------------------------------------------------------------------

    /// Verifies that `empty_recent_files_with_api()` does NOT return `ComApartmentMismatch`
    /// when called from an MTA thread, since it now routes through `run_on_sta_thread`.
    #[test]
    #[ignore = "Modifies system state - run with: cargo test empty::tests::test_com_apartment_mismatch -- --ignored --nocapture"]
    fn test_com_apartment_mismatch() -> WincentResult<()> {
        use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED};

        unsafe {
            let hr = CoInitializeEx(None, COINIT_MULTITHREADED);
            assert!(hr.is_ok() || hr.0 == 1, "MTA init should succeed");

            let result = empty_recent_files_with_api(Duration::from_secs(10));
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
    fn test_empty_recent_files_zero_timeout_rejected() {
        let result = empty_recent_files_with_api(Duration::ZERO);
        assert!(
            matches!(result, Err(WincentError::InvalidArgument(_))),
            "got: {:?}",
            result
        );
    }

    /// Verifies the calling thread's COM state is unaffected by `empty_recent_files_with_api`
    /// since it now runs on its own STA thread.
    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_com_s_false_reference_counting -- --ignored --nocapture"]
    fn test_com_s_false_reference_counting() -> WincentResult<()> {
        use windows::Win32::System::Com::{
            CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED,
        };

        unsafe {
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            assert_eq!(hr.0, 0, "CoInitializeEx should return S_OK");

            for _ in 0..3 {
                let result = empty_recent_files_with_api(Duration::from_secs(10));
                assert!(result.is_ok(), "should succeed: {:?}", result);
            }

            CoUninitialize();

            // Calling thread fully uninitialised; MTA must succeed with no leaked refs.
            let hr = CoInitializeEx(None, COINIT_MULTITHREADED);
            assert!(
                hr.0 == 0 || hr.0 == 1,
                "Should initialize MTA with no leaked STA references. Got: 0x{:08X}",
                hr.0
            );
            CoUninitialize();
        }

        Ok(())
    }

    // -------------------------------------------------------------------------
    // Integration tests - internal primitives
    // -------------------------------------------------------------------------

    /// Verifies that `empty_recent_files_with_api()` clears the Recent Files list.
    ///
    /// Arrange: add a unique test file.
    /// Act:     call `empty_recent_files_with_api()`.
    /// Assert:  the list becomes empty.
    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_empty_recent_files -- --ignored --nocapture"]
    fn test_empty_recent_files() -> WincentResult<()> {
        let test_dir = setup_test_env()?;

        let timestamp = unique_millis()?;
        let test_file = create_test_file(&test_dir, &format!("empty_test_{timestamp}.txt"), "x")?;
        let test_path = path_to_str(&test_file)?;
        // Guard is created after test_file so it can register the path for cleanup.
        let _guard = TestGuard::new(test_dir.clone()).with_recent_file(test_path);

        add_file_to_recent_native(test_path, Duration::from_secs(10))?;
        assert!(
            wait_until_recent_file_present(test_path, 20)?,
            "Test file should appear in Recent Files before clearing"
        );

        empty_recent_files_with_api(Duration::from_secs(10))?;
        assert!(
            wait_until_recent_files_empty(10)?,
            "Recent Files list should be empty after empty_recent_files_with_api()"
        );

        Ok(())
    }

    /// Verifies that `empty_user_folders_with_jumplist_file()` deletes the
    /// AutomaticDestinations jump list file that tracks user-visited frequent folders.
    ///
    /// Arrange: ensure the jump list file exists (create a stub if absent).
    /// Act:     call `empty_user_folders_with_jumplist_file()`.
    /// Assert:  the jump list file no longer exists on disk.
    ///
    /// This is a file-system-level test. The Frequent Folders Shell namespace is
    /// not queried because pinned entries (the only ones we can create
    /// programmatically) are stored separately and are NOT affected by deleting
    /// this file.
    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_empty_user_folders_deletes_jumplist_file -- --ignored --nocapture"]
    fn test_empty_user_folders_deletes_jumplist_file() -> WincentResult<()> {
        use crate::utils::get_windows_recent_folder;

        /// RAII guard that restores the jump list file to its pre-test state on drop.
        struct JumplistGuard {
            path: std::path::PathBuf,
            /// `Some(bytes)` -> restore by writing the original bytes back.
            /// `None`        -> file did not exist before the test; delete any stub we created.
            original: Option<Vec<u8>>,
        }
        impl Drop for JumplistGuard {
            fn drop(&mut self) {
                match &self.original {
                    Some(bytes) => {
                        let _ = std::fs::write(&self.path, bytes);
                    }
                    None => {
                        let _ = std::fs::remove_file(&self.path);
                    }
                }
            }
        }

        let recent_folder = get_windows_recent_folder()?;
        let jumplist_file = std::path::Path::new(&recent_folder)
            .join("AutomaticDestinations")
            .join("f01b4d95cf55d32a.automaticDestinations-ms");

        // Snapshot pre-test state; the guard will restore it on drop.
        let original_contents: Option<Vec<u8>> = if jumplist_file.exists() {
            Some(std::fs::read(&jumplist_file).map_err(WincentError::Io)?)
        } else {
            None
        };
        let _restore_guard = JumplistGuard {
            path: jumplist_file.clone(),
            original: original_contents,
        };

        // Ensure the file exists before the call (create a stub if needed).
        if !jumplist_file.exists() {
            if let Some(parent) = jumplist_file.parent() {
                std::fs::create_dir_all(parent).map_err(WincentError::Io)?;
            }
            std::fs::write(&jumplist_file, b"stub").map_err(WincentError::Io)?;
        }
        assert!(
            jumplist_file.exists(),
            "Jump list file must exist before the call"
        );

        empty_user_folders_with_jumplist_file()?;

        // Assert the file is absent; the guard restores the original afterwards.
        assert!(
            !jumplist_file.exists(),
            "Jump list file must be deleted by empty_user_folders_with_jumplist_file()"
        );

        Ok(())
    }

    /// Verifies that `empty_pinned_folders()` removes ALL pinned folders.
    ///
    /// Arrange: pin two distinct test directories.
    /// Act:     call `empty_pinned_folders()`.
    /// Assert:  both pinned directories are absent (semantic boundary: "clears all").
    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_empty_pinned_folders_clears_all -- --ignored --nocapture"]
    fn test_empty_pinned_folders_clears_all() -> WincentResult<()> {
        let test_dir = setup_test_env()?;
        let test_path_a = path_to_str(&test_dir)?.to_owned();

        // Create a second distinct directory inside the test tree.
        let dir_b = test_dir.join("sub_b");
        std::fs::create_dir_all(&dir_b).map_err(WincentError::Io)?;
        let test_path_b = path_to_str(&dir_b)?.to_owned();

        let _guard = TestGuard::new(test_dir.clone())
            .with_pinned(&test_path_a)
            .with_pinned(&test_path_b);

        pin_frequent_folder(&test_path_a, Duration::from_secs(10))?;
        pin_frequent_folder(&test_path_b, Duration::from_secs(10))?;

        assert!(
            wait_until_folder_present(&test_path_a, 10)?,
            "Folder A should be pinned before clearing"
        );
        assert!(
            wait_until_folder_present(&test_path_b, 10)?,
            "Folder B should be pinned before clearing"
        );

        empty_pinned_folders()?;
        thread::sleep(Duration::from_secs(1));

        assert!(
            wait_until_folder_absent(&test_path_a, 10)?,
            "Folder A should be absent after empty_pinned_folders()"
        );
        assert!(
            wait_until_folder_absent(&test_path_b, 10)?,
            "Folder B should be absent after empty_pinned_folders()"
        );

        Ok(())
    }

    // -------------------------------------------------------------------------
    // Integration tests - public API semantics
    // -------------------------------------------------------------------------

    /// Smoke-tests that `empty_frequent_folders(false)` returns `Ok(())`.
    ///
    /// Act:    call `empty_frequent_folders(false)`.
    /// Assert: no error is returned.
    ///
    /// This verifies the call path completes without error. The jump list
    /// deletion is covered by `test_empty_user_folders_deletes_jumplist_file`.
    /// The "pinned folder survives" property cannot be asserted because
    /// `pin_frequent_folder` does not guarantee a shell-pinned entry immune to
    /// jump list deletion on all Windows configurations, and Windows
    /// automatically recreates the jump list file after deletion, making a
    /// post-call file-existence check racy.
    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_empty_frequent_folders_false_no_error -- --ignored --nocapture"]
    fn test_empty_frequent_folders_false_no_error() -> WincentResult<()> {
        empty_frequent_folders(false)
    }

    /// Verifies that `empty_frequent_folders(true)` also unpins pinned folders.
    ///
    /// Arrange: pin a test folder.
    /// Act:     call `empty_frequent_folders(true)`.
    /// Assert:  the pinned folder is absent.
    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_empty_frequent_folders_true_removes_pinned -- --ignored --nocapture"]
    fn test_empty_frequent_folders_true_removes_pinned() -> WincentResult<()> {
        let test_dir = setup_test_env()?;
        let test_path = path_to_str(&test_dir)?.to_owned();
        let _guard = TestGuard::new(test_dir.clone()).with_pinned(&test_path);

        pin_frequent_folder(&test_path, Duration::from_secs(10))?;
        assert!(
            wait_until_folder_present(&test_path, 10)?,
            "Test folder should be pinned before the call"
        );

        // true - jump list deleted AND all pinned folders unpinned.
        empty_frequent_folders(true)?;
        thread::sleep(Duration::from_secs(1));

        assert!(
            wait_until_folder_absent(&test_path, 10)?,
            "Pinned folder must be absent after empty_frequent_folders(true)"
        );

        Ok(())
    }

    /// Verifies that the default Quick Access clear path clears Recent Files.
    ///
    /// Arrange: add a test file to Recent Files.
    /// Act:     call `empty_items(QuickAccess::All, EmptyOptions::new())`.
    /// Assert:  Recent Files is empty.
    ///
    /// Note: the jump list file deletion is already covered by
    /// `test_empty_user_folders_deletes_jumplist_file`. It is not asserted here
    /// because Windows automatically recreates the file shortly after deletion,
    /// making a post-call existence check inherently racy. The "pinned folder
    /// survives" property is also not asserted because `pin_frequent_folder`
    /// does not guarantee a shell-pinned entry immune to jump list deletion on
    /// all Windows configurations.
    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_empty_quick_access_default_clears_recent_files -- --ignored --nocapture"]
    fn test_empty_quick_access_default_clears_recent_files() -> WincentResult<()> {
        let test_dir = setup_test_env()?;

        let timestamp = unique_millis()?;
        let test_file = create_test_file(&test_dir, &format!("qa_false_{timestamp}.txt"), "x")?;
        let file_path = path_to_str(&test_file)?;
        let _guard = TestGuard::new(test_dir.clone()).with_recent_file(file_path);

        add_file_to_recent_native(file_path, Duration::from_secs(10))?;
        assert!(
            wait_until_recent_file_present(file_path, 20)?,
            "Test file should appear in Recent Files before clearing"
        );

        empty_items(QuickAccess::All, EmptyOptions::new())?;
        thread::sleep(Duration::from_secs(1));

        assert!(
            wait_until_recent_files_empty(10)?,
            "Recent Files should be empty after the default Quick Access clear"
        );

        Ok(())
    }

    /// Verifies that `empty_quick_access(true)` clears both Recent Files and all
    /// pinned folders.
    ///
    /// Arrange: pin a test folder and add a test file to Recent Files.
    /// Act:     call `empty_quick_access(true)`.
    /// Assert:  Recent Files is empty and the pinned folder is absent.
    #[test]
    #[ignore = "Modifies system state - run with: cargo test test_empty_quick_access_true_clears_all -- --ignored --nocapture"]
    fn test_empty_quick_access_true_clears_all() -> WincentResult<()> {
        let test_dir = setup_test_env()?;
        let test_path = path_to_str(&test_dir)?.to_owned();

        let timestamp = unique_millis()?;
        let test_file = create_test_file(&test_dir, &format!("qa_true_{timestamp}.txt"), "x")?;
        let file_path = path_to_str(&test_file)?;
        let _guard = TestGuard::new(test_dir.clone())
            .with_pinned(&test_path)
            .with_recent_file(file_path);

        add_file_to_recent_native(file_path, Duration::from_secs(10))?;
        pin_frequent_folder(&test_path, Duration::from_secs(10))?;

        assert!(
            wait_until_recent_file_present(file_path, 20)?,
            "Test file should appear in Recent Files before clearing"
        );
        assert!(
            wait_until_folder_present(&test_path, 10)?,
            "Test folder should be pinned before clearing"
        );

        empty_items(
            QuickAccess::All,
            EmptyOptions::new().with_also_pinned_folders(true),
        )?;
        thread::sleep(Duration::from_secs(1));

        assert!(
            wait_until_recent_files_empty(10)?,
            "Recent Files should be empty after empty_quick_access(true)"
        );
        assert!(
            wait_until_folder_absent(&test_path, 10)?,
            "Pinned folder must be absent after empty_quick_access(true)"
        );

        Ok(())
    }
}
