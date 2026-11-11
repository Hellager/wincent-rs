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

use crate::{backend::QuickAccessBackend, error::WincentError, QuickAccess, WincentResult};
use std::time::Duration;

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
}
