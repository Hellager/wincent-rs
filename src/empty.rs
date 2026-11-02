//! Windows Quick Access cleanup operations
//!
//! Provides unified interface for clearing Windows Quick Access items including:
//! - Recent files
//! - Frequent folders (matching Recent shortcuts and Explorer backing data)
//! - Complete Quick Access history
//!
//! # Key Functionality
//! - Clear recent files using Windows Shell API
//! - Remove frequent folders and their matching Recent shortcuts
//! - Full Quick Access reset capabilities
//! - Atomic operations with proper cleanup sequencing

use crate::{
    backend::QuickAccessBackend,
    destlist::{frequent_folders_dest_path, parse_file},
    error::WincentError,
    recent_links,
    utils::paths_equal,
    QuickAccess, WincentResult,
};
use std::{io::ErrorKind, path::Path, time::Duration};
use windows::Win32::UI::Shell::SHAddToRecentDocs;

/// SHARD_PATHW - with a null data pointer, clears all recent documents.
const SHARD_PATHW: u32 = 0x0000_0003;

/// Clears the Windows Recent Files list via a dedicated STA thread.
///
/// Calls `SHAddToRecentDocs(SHARD_PATHW, NULL)`, which clears the Recent Files
/// entries and streams together with their shortcuts in the Recent directory.
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

/// Clears Frequent Folders by removing matching Recent shortcuts before the DestList.
///
/// Every raw DestList entry is included, including hidden entries. Matching `.lnk`
/// files are deleted first so Explorer cannot use them while rebuilding the list.
/// The DestList file is deleted only after all shortcut cleanup succeeds.
pub(crate) fn clear_frequent_folders_with_destlist_reset(timeout: Duration) -> WincentResult<()> {
    let jumplist_file = frequent_folders_dest_path()?;
    if !jumplist_file.try_exists().map_err(WincentError::Io)? {
        return Ok(());
    }

    let parsed = parse_file(&jumplist_file)?;
    let paths = unique_paths(
        parsed
            .dest_list()
            .entries()
            .iter()
            .map(crate::destlist::DestListEntry::path),
    );

    clear_frequent_folders_data(
        &paths,
        |targets| recent_links::delete_recent_links_for_targets(targets, timeout).map(|_| ()),
        || remove_file_if_present(&jumplist_file),
    )
}

fn unique_paths<'path>(paths: impl IntoIterator<Item = &'path str>) -> Vec<String> {
    let mut unique: Vec<String> = Vec::new();

    for path in paths {
        if path.is_empty() || unique.iter().any(|existing| paths_equal(existing, path)) {
            continue;
        }
        unique.push(path.to_owned());
    }

    unique
}

fn clear_frequent_folders_data<FDeleteLinks, FDeleteDestList>(
    paths: &[String],
    delete_links: FDeleteLinks,
    delete_dest_list: FDeleteDestList,
) -> WincentResult<()>
where
    FDeleteLinks: FnOnce(&[String]) -> WincentResult<()>,
    FDeleteDestList: FnOnce() -> WincentResult<()>,
{
    delete_links(paths)?;
    delete_dest_list()
}

fn remove_file_if_present(path: &Path) -> WincentResult<()> {
    match std::fs::remove_file(path) {
        Ok(()) => Ok(()),
        Err(error) if error.kind() == ErrorKind::NotFound => Ok(()),
        Err(error) => Err(WincentError::Io(error)),
    }
}

/// Clears items from Windows Quick Access.
///
/// Recent Files are cleared through the Shell API. Frequent Folders are parsed
/// from the current DestList; matching Recent shortcuts are deleted before the
/// DestList file so Explorer can rebuild from a clean state. Explorer may then
/// recreate system-default folder pins. A successful clear always refreshes
/// matching Quick Access or Home Explorer windows.
///
/// # Returns
///
/// Returns `Ok(())` if all requested cleanup steps completed successfully.
///
/// # Example
///
/// ```rust,no_run
/// use wincent::prelude::*;
///
/// # fn main() -> WincentResult<()> {
///     let manager = QuickAccessManager::new();
///     manager.empty_items(QuickAccess::All)?;
///     Ok(())
/// # }
/// ```
fn empty_all_items_with_backend(
    timeout: Duration,
    backend: &dyn QuickAccessBackend,
) -> WincentResult<()> {
    let mut recent_files_cleared = false;
    let frequent_folders_cleared;
    let mut first_error = None;

    match backend.clear_recent_files(timeout) {
        Ok(()) => recent_files_cleared = true,
        Err(source) => first_error = Some(source),
    }

    match backend.clear_frequent_folders_jumplist(timeout) {
        Ok(()) => {
            frequent_folders_cleared = true;
        }
        Err(source) => {
            frequent_folders_cleared = false;
            if first_error.is_none() {
                first_error = Some(source);
            }
        }
    }

    if let Some(source) = first_error {
        return Err(WincentError::PartialEmpty {
            recent_files_cleared,
            frequent_folders_cleared,
            source: Box::new(source),
        });
    }

    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RefreshPolicy {
    PropagateFailure,
    BestEffort,
}

fn refresh_policy_for_result(result: &WincentResult<()>) -> Option<RefreshPolicy> {
    match result {
        Ok(()) => Some(RefreshPolicy::PropagateFailure),
        Err(WincentError::PartialEmpty { .. }) => Some(RefreshPolicy::BestEffort),
        Err(_) => None,
    }
}

pub(crate) fn empty_items_with_backend(
    qa_type: QuickAccess,
    timeout: Duration,
    backend: &dyn QuickAccessBackend,
) -> WincentResult<()> {
    let result = match qa_type {
        QuickAccess::RecentFiles => backend.clear_recent_files(timeout),
        QuickAccess::FrequentFolders => backend.clear_frequent_folders_jumplist(timeout),
        QuickAccess::All => empty_all_items_with_backend(timeout, backend),
    };

    match refresh_policy_for_result(&result) {
        Some(RefreshPolicy::PropagateFailure) => backend.refresh(timeout)?,
        Some(RefreshPolicy::BestEffort) => {
            // Preserve the cleanup failure while still trying to show any
            // successfully cleared category in Explorer.
            let _ = backend.refresh(timeout);
        }
        None => {}
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::cell::{Cell, RefCell};
    use std::sync::Mutex;
    use std::time::Duration;

    struct FakeEmptyBackend {
        items: Vec<String>,
        calls: Mutex<Vec<String>>,
        clear_recent_files_error: Option<String>,
        clear_frequent_folders_jumplist_error: Option<String>,
        refresh_explorer_error: Option<String>,
    }

    impl FakeEmptyBackend {
        fn new(items: Vec<String>) -> Self {
            Self {
                items,
                calls: Mutex::new(Vec::new()),
                clear_recent_files_error: None,
                clear_frequent_folders_jumplist_error: None,
                refresh_explorer_error: None,
            }
        }

        fn with_clear_recent_files_error(mut self, message: impl Into<String>) -> Self {
            self.clear_recent_files_error = Some(message.into());
            self
        }

        fn with_clear_frequent_folders_jumplist_error(
            mut self,
            message: impl Into<String>,
        ) -> Self {
            self.clear_frequent_folders_jumplist_error = Some(message.into());
            self
        }

        fn with_refresh_explorer_error(mut self, message: impl Into<String>) -> Self {
            self.refresh_explorer_error = Some(message.into());
            self
        }

        fn record(&self, call: impl Into<String>) {
            self.calls.lock().unwrap().push(call.into());
        }

        fn calls(&self) -> Vec<String> {
            self.calls.lock().unwrap().clone()
        }
    }

    impl crate::backend::QuickAccessBackend for FakeEmptyBackend {
        fn get_items(&self, qa_type: QuickAccess, timeout: Duration) -> WincentResult<Vec<String>> {
            self.record(format!("get_items:{qa_type:?}:{}", timeout.as_secs()));
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

        fn remove_pinned_frequent_folder(
            &self,
            _path: &str,
            _timeout: Duration,
        ) -> WincentResult<()> {
            Ok(())
        }

        fn remove_unpinned_frequent_folder(
            &self,
            _path: &str,
            _timeout: Duration,
        ) -> WincentResult<()> {
            Ok(())
        }

        fn clear_recent_files(&self, _timeout: Duration) -> WincentResult<()> {
            self.record("clear_recent_files");
            if let Some(error) = &self.clear_recent_files_error {
                return Err(WincentError::SystemError(error.clone()));
            }
            Ok(())
        }

        fn clear_frequent_folders_jumplist(&self, _timeout: Duration) -> WincentResult<()> {
            self.record("clear_frequent_folders_jumplist");
            if let Some(error) = &self.clear_frequent_folders_jumplist_error {
                return Err(WincentError::SystemError(error.clone()));
            }
            Ok(())
        }

        fn notify_shell_change(&self) {
            self.record("notify_shell_change");
        }

        fn refresh_shell_views(&self, _timeout: Duration) -> WincentResult<()> {
            self.record("refresh_shell_views");
            if let Some(error) = &self.refresh_explorer_error {
                return Err(WincentError::SystemError(error.clone()));
            }
            Ok(())
        }
    }

    #[test]
    fn unique_paths_uses_windows_path_semantics() {
        let paths = unique_paths(["C:\\Work", "c:/work/", "C:\\Other", "", "C:\\OTHER\\"]);

        assert_eq!(paths, vec!["C:\\Work", "C:\\Other"]);
    }

    #[test]
    fn frequent_cleanup_deletes_links_before_destlist() -> WincentResult<()> {
        let calls = RefCell::new(Vec::new());
        let paths = vec!["C:\\One".to_string(), "C:\\Two".to_string()];

        clear_frequent_folders_data(
            &paths,
            |targets| {
                calls
                    .borrow_mut()
                    .push(format!("links:{}", targets.join("|")));
                Ok(())
            },
            || {
                calls.borrow_mut().push("destlist".to_string());
                Ok(())
            },
        )?;

        assert_eq!(
            calls.into_inner(),
            vec!["links:C:\\One|C:\\Two", "destlist"]
        );
        Ok(())
    }

    #[test]
    fn frequent_cleanup_preserves_destlist_when_link_cleanup_fails() {
        let destlist_deleted = Cell::new(false);
        let paths = vec!["C:\\One".to_string(), "C:\\Two".to_string()];

        let error = clear_frequent_folders_data(
            &paths,
            |_| Err(WincentError::SystemError("link cleanup failed".to_string())),
            || {
                destlist_deleted.set(true);
                Ok(())
            },
        )
        .expect_err("link cleanup should fail");

        assert!(matches!(
            error,
            WincentError::SystemError(ref message) if message == "link cleanup failed"
        ));
        assert!(!destlist_deleted.get());
    }

    #[test]
    fn refresh_policy_propagate_on_ok() {
        assert_eq!(
            refresh_policy_for_result(&Ok(())),
            Some(RefreshPolicy::PropagateFailure)
        );
    }

    #[test]
    fn refresh_policy_best_effort_on_partial_empty() {
        assert_eq!(
            refresh_policy_for_result(&Err(WincentError::PartialEmpty {
                recent_files_cleared: true,
                frequent_folders_cleared: false,
                source: Box::new(WincentError::SystemError("failed".to_string())),
            })),
            Some(RefreshPolicy::BestEffort)
        );
    }

    #[test]
    fn refresh_policy_skips_non_partial_empty_error() {
        assert_eq!(
            refresh_policy_for_result(&Err(WincentError::SystemError("failed".to_string()))),
            None
        );
    }

    fn assert_partial_empty(
        error: WincentError,
        expected_recent_files_cleared: bool,
        expected_frequent_folders_cleared: bool,
        expected_source: &str,
    ) {
        match error {
            WincentError::PartialEmpty {
                recent_files_cleared,
                frequent_folders_cleared,
                source,
            } => {
                assert_eq!(recent_files_cleared, expected_recent_files_cleared);
                assert_eq!(frequent_folders_cleared, expected_frequent_folders_cleared);
                assert!(matches!(
                    *source,
                    WincentError::SystemError(ref message) if message == expected_source
                ));
            }
            other => panic!("Expected PartialEmpty, got: {other:?}"),
        }
    }

    #[test]
    fn empty_all_continues_to_frequent_folders_when_recent_files_fails() {
        let backend =
            FakeEmptyBackend::new(Vec::new()).with_clear_recent_files_error("recent failed");

        let error = empty_items_with_backend(QuickAccess::All, Duration::from_secs(3), &backend)
            .unwrap_err();

        assert_partial_empty(error, false, true, "recent failed");
        assert_eq!(
            backend.calls(),
            vec![
                "clear_recent_files".to_string(),
                "clear_frequent_folders_jumplist".to_string(),
                "notify_shell_change".to_string(),
                "refresh_shell_views".to_string(),
            ]
        );
    }

    #[test]
    fn empty_all_preserves_recent_error_when_both_categories_fail_before_frequent_progress() {
        let backend = FakeEmptyBackend::new(Vec::new())
            .with_clear_recent_files_error("recent failed")
            .with_clear_frequent_folders_jumplist_error("frequent failed");

        let error = empty_items_with_backend(QuickAccess::All, Duration::from_secs(3), &backend)
            .unwrap_err();

        assert_partial_empty(error, false, false, "recent failed");
        assert_eq!(
            backend.calls(),
            vec![
                "clear_recent_files".to_string(),
                "clear_frequent_folders_jumplist".to_string(),
                "notify_shell_change".to_string(),
                "refresh_shell_views".to_string(),
            ]
        );
    }

    #[test]
    fn empty_all_reports_frequent_error_when_recent_files_succeeds() {
        let backend = FakeEmptyBackend::new(Vec::new())
            .with_clear_frequent_folders_jumplist_error("frequent failed");

        let error = empty_items_with_backend(QuickAccess::All, Duration::from_secs(3), &backend)
            .unwrap_err();

        assert_partial_empty(error, true, false, "frequent failed");
        assert_eq!(
            backend.calls(),
            vec![
                "clear_recent_files".to_string(),
                "clear_frequent_folders_jumplist".to_string(),
                "notify_shell_change".to_string(),
                "refresh_shell_views".to_string(),
            ]
        );
    }

    #[test]
    fn empty_all_partial_from_recent_failure_refreshes_best_effort_without_overwriting_error() {
        let backend = FakeEmptyBackend::new(Vec::new())
            .with_clear_recent_files_error("recent failed")
            .with_refresh_explorer_error("refresh failed");

        let error = empty_items_with_backend(QuickAccess::All, Duration::from_secs(3), &backend)
            .unwrap_err();

        assert_partial_empty(error, false, true, "recent failed");
        assert_eq!(
            backend.calls(),
            vec![
                "clear_recent_files".to_string(),
                "clear_frequent_folders_jumplist".to_string(),
                "notify_shell_change".to_string(),
                "refresh_shell_views".to_string(),
            ]
        );
    }

    #[test]
    fn empty_frequent_folders_clears_then_refreshes() -> WincentResult<()> {
        let backend = FakeEmptyBackend::new(vec!["C:\\FolderA".to_string()]);

        empty_items_with_backend(
            QuickAccess::FrequentFolders,
            Duration::from_secs(3),
            &backend,
        )?;

        assert_eq!(
            backend.calls(),
            vec![
                "clear_frequent_folders_jumplist".to_string(),
                "notify_shell_change".to_string(),
                "refresh_shell_views".to_string(),
            ]
        );
        Ok(())
    }

    #[test]
    fn empty_recent_files_propagates_refresh_failure() {
        let backend =
            FakeEmptyBackend::new(Vec::new()).with_refresh_explorer_error("refresh failed");

        let error =
            empty_items_with_backend(QuickAccess::RecentFiles, Duration::from_secs(3), &backend)
                .unwrap_err();

        assert!(matches!(
            error,
            WincentError::SystemError(message) if message == "refresh failed"
        ));
        assert_eq!(
            backend.calls(),
            vec![
                "clear_recent_files".to_string(),
                "notify_shell_change".to_string(),
                "refresh_shell_views".to_string(),
            ]
        );
    }

    #[test]
    fn empty_recent_files_zero_timeout_is_rejected() {
        let result = empty_recent_files_with_api(Duration::ZERO);
        assert!(matches!(result, Err(WincentError::InvalidArgument(_))));
    }
}
