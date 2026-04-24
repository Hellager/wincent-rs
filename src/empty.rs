//! Windows Quick Access cleanup operations
//!
//! Provides unified interface for clearing Windows Quick Access items including:
//! - Recent files
//! - Frequent folders (both pinned and normal)
//! - Complete Quick Access history
//!
//! Implements multiple cleanup strategies with fallback mechanisms
//!
//! # Key Functionality
//! - Clear recent files using Windows Shell API
//! - Remove frequent folders through file system operations
//! - Clear pinned folders via PowerShell scripts
//! - Full Quick Access reset capabilities
//! - Atomic operations with proper cleanup sequencing

use crate::{
    com::{ComGuard, ComInitStatus},
    error::WincentError,
    utils::get_windows_recent_folder,
    WincentResult,
};
use windows::Win32::UI::Shell::SHAddToRecentDocs;

/// Clears the Windows Recent Files list using the Windows Shell API.
///
/// Uses `SHAddToRecentDocs` with a NULL parameter to clear all recent files.
/// This is an asynchronous operation that requires the calling thread's COM
/// context to remain active for Windows Shell to complete the operation.
///
/// # Windows Behavior Notes
///
/// - **Asynchronous Processing**: The list may not be cleared immediately.
///   Windows Shell processes this update asynchronously.
/// - **System-wide Effect**: This clears the Recent Files list for the current user,
///   affecting all applications that use the Windows Recent Items feature.
pub(crate) fn empty_recent_files_with_api() -> WincentResult<()> {
    let _com = ComGuard::try_initialize().map_err(|status| match status {
        ComInitStatus::ApartmentMismatch => WincentError::ComApartmentMismatch(
            "Thread already initialized with incompatible COM apartment model".to_string(),
        ),
        ComInitStatus::OtherError(hr) => WincentError::WindowsApi(hr),
        _ => unreachable!(),
    })?;

    unsafe {
        // Passing None clears all recent docs (SHARD_PATHW with null pv)
        SHAddToRecentDocs(0x0000_0003, None);
    }

    Ok(())
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

/// Unpin all pinned folders from Quick Access using PowerShell commands.
///
/// Iterates every item in the Frequent Folders namespace and invokes
/// `unpinfromhome` on each one. On Windows 10 this reliably removes both
/// user-pinned and system-default entries. On Windows 11, `unpinfromhome`
/// alone may not remove items that require the `pintohome` toggle workaround;
/// this function does **not** apply that workaround, so results may be
/// incomplete on Windows 11.
fn empty_pinned_folders_powershell() -> WincentResult<()> {
    use crate::script_executor::ScriptExecutor;
    use crate::script_strategy::PSScript;

    let start = std::time::Instant::now();
    let script_path = crate::script_storage::ScriptStorage::get_script_path(PSScript::EmptyPinnedFolders)?;
    let output = ScriptExecutor::execute_ps_script(PSScript::EmptyPinnedFolders, None)?;
    let duration = start.elapsed();
    let _ = ScriptExecutor::parse_output_to_strings(
        output,
        PSScript::EmptyPinnedFolders,
        script_path,
        None,
        duration,
    )?;

    Ok(())
}

/// Attempts to unpin all pinned folders from Quick Access.
///
/// Invokes `unpinfromhome` on every item currently in the Frequent Folders
/// namespace via PowerShell. On Windows 10 this reliably removes all pinned
/// entries. On Windows 11, items may require a `pintohome` toggle workaround
/// to be fully removed; this function does **not** apply that workaround, so
/// some entries may persist on Windows 11.
///
/// Currently delegates to the PowerShell fallback. A native COM fast path
/// with proper Win11 compatibility is planned but not yet implemented.
pub(crate) fn empty_pinned_folders() -> WincentResult<()> {
    empty_pinned_folders_powershell()
}

/****************************************************** Empty Quick Access ******************************************************/

/// Clears all items from the Windows Recent Files list.
///
/// # Returns
///
/// Returns `Ok(())` if all recent files were successfully cleared.
///
/// # Example
///
/// ```no_run
/// use wincent::{empty::empty_recent_files, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Clear all recent files
///     empty_recent_files()?;
///     println!("Recent files list has been cleared");
///     Ok(())
/// }
/// ```
pub fn empty_recent_files() -> WincentResult<()> {
    empty_recent_files_with_api()
}

/// Clears the Windows Frequent Folders list from Quick Access.
///
/// Always removes user-visited frequent folders by deleting the jump list file.
/// Pass `also_pinned_folders: true` to additionally attempt to unpin all pinned
/// folders via PowerShell (`unpinfromhome`). On Windows 10 this reliably
/// removes all pinned entries; on Windows 11 some entries may persist because
/// the PowerShell path does not apply the `pintohome` toggle workaround.
///
/// # Parameters
///
/// - `also_pinned_folders` — when `true`, also invokes `unpinfromhome` on
///   every item in the Frequent Folders namespace, attempting to clear pinned
///   entries. Results are reliable on Windows 10; on Windows 11 some items may
///   persist (no `pintohome` toggle workaround is applied).
///   When `false`, only the jump list file is removed; pinned folders are
///   left untouched.
///
/// # Returns
///
/// Returns `Ok(())` if the requested cleanup steps completed successfully.
///
/// # Example
///
/// ```no_run
/// use wincent::{empty::empty_frequent_folders, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Clear only user-visited frequent folders (jump list); pinned folders remain.
///     empty_frequent_folders(false)?;
///
///     // Clear user-visited frequent folders AND unpin all pinned folders.
///     empty_frequent_folders(true)?;
///
///     Ok(())
/// }
/// ```
pub fn empty_frequent_folders(also_pinned_folders: bool) -> WincentResult<()> {
    empty_user_folders_with_jumplist_file()?;
    if also_pinned_folders {
        empty_pinned_folders()?;
    }
    Ok(())
}

/// Clears items from Windows Quick Access.
///
/// Always clears the Recent Files list and user-visited frequent folders
/// (jump list file). Pass `also_pinned_folders: true` to additionally unpin
/// all pinned folders.
///
/// # Parameters
///
/// - `also_pinned_folders` — when `true`, also invokes `unpinfromhome` on
///   every item in the Frequent Folders namespace, attempting to clear pinned
///   entries. Results are reliable on Windows 10; on Windows 11 some items may
///   persist (no `pintohome` toggle workaround is applied).
///   When `false`, pinned folders are left untouched.
///
/// # Returns
///
/// Returns `Ok(())` if all requested cleanup steps completed successfully.
///
/// # Example
///
/// ```no_run
/// use wincent::{empty::empty_quick_access, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Clear recent files and frequent folders; leave pinned folders intact.
///     empty_quick_access(false)?;
///
///     // Clear everything, including pinned folders.
///     empty_quick_access(true)?;
///
///     Ok(())
/// }
/// ```
pub fn empty_quick_access(also_pinned_folders: bool) -> WincentResult<()> {
    empty_recent_files()?;
    empty_frequent_folders(also_pinned_folders)?;
    Ok(())
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
    use std::thread;
    use std::time::Duration;

    // -------------------------------------------------------------------------
    // Integration-test helpers
    // -------------------------------------------------------------------------

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
            Self { dir, pinned: Vec::new(), recent_file: None }
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
            // Best-effort recent-file removal — ignore errors.
            if let Some(ref p) = self.recent_file {
                let _ = remove_from_recent_files(p);
            }
            let _ = cleanup_test_env(&self.dir);
        }
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

    /// Verifies that `empty_recent_files_with_api()` returns `ComApartmentMismatch`
    /// when the calling thread is already initialized as MTA.
    #[test]
    fn test_com_apartment_mismatch() -> WincentResult<()> {
        use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED};

        unsafe {
            let hr = CoInitializeEx(None, COINIT_MULTITHREADED);
            assert!(hr.is_ok() || hr.0 == 1, "MTA init should succeed");

            let result = empty_recent_files_with_api();

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

            CoUninitialize();
        }

        Ok(())
    }

    /// Verifies that `ComGuard` correctly balances `CoInitialize`/`CoUninitialize`
    /// reference counts by cycling through multiple init/uninit rounds and checking
    /// that no STA references leak into a subsequent MTA initialization.
    #[test]
    #[ignore = "Modifies system state"]
    fn test_com_s_false_reference_counting() -> WincentResult<()> {
        use windows::Win32::System::Com::{
            CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED,
        };

        unsafe {
            for cycle in 1..=3 {
                let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
                assert_eq!(hr.0, 0, "Cycle {cycle}: CoInitializeEx should return S_OK");

                let result = empty_recent_files_with_api();
                assert!(result.is_ok(), "Cycle {cycle}: should handle S_FALSE: {:?}", result);

                CoUninitialize();
            }

            // If STA references leaked, this will fail with RPC_E_CHANGED_MODE.
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
    // Integration tests — internal primitives
    // -------------------------------------------------------------------------

    /// Verifies that `empty_recent_files_with_api()` clears the Recent Files list.
    ///
    /// Arrange: add a unique test file.
    /// Act:     call `empty_recent_files_with_api()`.
    /// Assert:  the list becomes empty.
    #[test]
    #[ignore = "Modifies system state"]
    fn test_empty_recent_files() -> WincentResult<()> {
        let test_dir = setup_test_env()?;

        let timestamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_millis();
        let test_file = create_test_file(&test_dir, &format!("empty_test_{timestamp}.txt"), "x")?;
        let test_path = test_file.to_str().unwrap();
        // Guard is created after test_file so it can register the path for cleanup.
        let _guard = TestGuard::new(test_dir.clone()).with_recent_file(test_path);

        add_file_to_recent_native(test_path)?;
        assert!(
            wait_until_recent_file_present(test_path, 20)?,
            "Test file should appear in Recent Files before clearing"
        );

        empty_recent_files_with_api()?;
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
    /// This is a file-system–level test. The Frequent Folders Shell namespace is
    /// not queried because pinned entries (the only ones we can create
    /// programmatically) are stored separately and are NOT affected by deleting
    /// this file.
    #[test]
    #[ignore = "Modifies system state"]
    fn test_empty_user_folders_deletes_jumplist_file() -> WincentResult<()> {
        use crate::utils::get_windows_recent_folder;

        /// RAII guard that restores the jump list file to its pre-test state on drop.
        struct JumplistGuard {
            path: std::path::PathBuf,
            /// `Some(bytes)` → restore by writing the original bytes back.
            /// `None`        → file did not exist before the test; delete any stub we created.
            original: Option<Vec<u8>>,
        }
        impl Drop for JumplistGuard {
            fn drop(&mut self) {
                match &self.original {
                    Some(bytes) => { let _ = std::fs::write(&self.path, bytes); }
                    None        => { let _ = std::fs::remove_file(&self.path); }
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
        assert!(jumplist_file.exists(), "Jump list file must exist before the call");

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
    #[ignore = "Modifies system state"]
    fn test_empty_pinned_folders_clears_all() -> WincentResult<()> {
        let test_dir = setup_test_env()?;
        let test_path_a = test_dir.to_str().unwrap().to_owned();

        // Create a second distinct directory inside the test tree.
        let dir_b = test_dir.join("sub_b");
        std::fs::create_dir_all(&dir_b).map_err(WincentError::Io)?;
        let test_path_b = dir_b.to_str().unwrap().to_owned();

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
    // Integration tests — public API semantics
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
    #[ignore = "Modifies system state"]
    fn test_empty_frequent_folders_false_no_error() -> WincentResult<()> {
        empty_frequent_folders(false)
    }

    /// Verifies that `empty_frequent_folders(true)` also unpins pinned folders.
    ///
    /// Arrange: pin a test folder.
    /// Act:     call `empty_frequent_folders(true)`.
    /// Assert:  the pinned folder is absent.
    #[test]
    #[ignore = "Modifies system state"]
    fn test_empty_frequent_folders_true_removes_pinned() -> WincentResult<()> {
        let test_dir = setup_test_env()?;
        let test_path = test_dir.to_str().unwrap().to_owned();
        let _guard = TestGuard::new(test_dir.clone()).with_pinned(&test_path);

        pin_frequent_folder(&test_path, Duration::from_secs(10))?;
        assert!(
            wait_until_folder_present(&test_path, 10)?,
            "Test folder should be pinned before the call"
        );

        // true → jump list deleted AND all pinned folders unpinned.
        empty_frequent_folders(true)?;
        thread::sleep(Duration::from_secs(1));

        assert!(
            wait_until_folder_absent(&test_path, 10)?,
            "Pinned folder must be absent after empty_frequent_folders(true)"
        );

        Ok(())
    }

    /// Verifies that `empty_quick_access(false)` clears Recent Files.
    ///
    /// Arrange: add a test file to Recent Files.
    /// Act:     call `empty_quick_access(false)`.
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
    #[ignore = "Modifies system state"]
    fn test_empty_quick_access_false_preserves_pinned() -> WincentResult<()> {
        let test_dir = setup_test_env()?;

        let timestamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_millis();
        let test_file = create_test_file(&test_dir, &format!("qa_false_{timestamp}.txt"), "x")?;
        let file_path = test_file.to_str().unwrap();
        let _guard = TestGuard::new(test_dir.clone()).with_recent_file(file_path);

        add_file_to_recent_native(file_path)?;
        assert!(
            wait_until_recent_file_present(file_path, 20)?,
            "Test file should appear in Recent Files before clearing"
        );

        empty_quick_access(false)?;
        thread::sleep(Duration::from_secs(1));

        assert!(
            wait_until_recent_files_empty(10)?,
            "Recent Files should be empty after empty_quick_access(false)"
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
    #[ignore = "Modifies system state"]
    fn test_empty_quick_access_true_clears_all() -> WincentResult<()> {
        let test_dir = setup_test_env()?;
        let test_path = test_dir.to_str().unwrap().to_owned();

        let timestamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_millis();
        let test_file = create_test_file(&test_dir, &format!("qa_true_{timestamp}.txt"), "x")?;
        let file_path = test_file.to_str().unwrap();
        let _guard = TestGuard::new(test_dir.clone())
            .with_pinned(&test_path)
            .with_recent_file(file_path);

        add_file_to_recent_native(file_path)?;
        pin_frequent_folder(&test_path, Duration::from_secs(10))?;

        assert!(
            wait_until_recent_file_present(file_path, 20)?,
            "Test file should appear in Recent Files before clearing"
        );
        assert!(
            wait_until_folder_present(&test_path, 10)?,
            "Test folder should be pinned before clearing"
        );

        empty_quick_access(true)?;
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
