//! Batch operations for Windows Quick Access items.

use crate::{
    error::WincentError,
    handle::{
        add_to_frequent_folders_with_timeout, add_to_recent_files_with_options,
        remove_from_frequent_folders_with_timeout, remove_from_recent_files_with_timeout,
        AddRecentFileOptions,
    },
    query,
    script_executor::QuickAccessDataFiles,
    utils::{paths_equal, refresh_explorer_window},
    QuickAccess, WincentResult,
};
use std::time::Duration;

/// Result of a batch operation containing succeeded and failed items.
#[derive(Debug, Default)]
pub struct BatchResult {
    /// Successfully processed items.
    pub succeeded: Vec<String>,
    /// Failed items with error details.
    ///
    /// `AlreadyExists` and `NotInRecent` may come from best-effort preflight
    /// checks. Explorer state can still change between the preflight query and
    /// the shell operation, so callers should treat those errors as a snapshot
    /// of the attempted operation rather than a durable global truth.
    pub failed: Vec<(String, WincentError)>,
}

impl BatchResult {
    /// Returns true if all operations succeeded.
    #[must_use]
    pub fn is_complete_success(&self) -> bool {
        self.failed.is_empty()
    }

    /// Returns true if at least one operation succeeded.
    #[must_use]
    pub fn has_partial_success(&self) -> bool {
        !self.succeeded.is_empty()
    }

    /// Returns the success rate (0.0 to 1.0).
    #[must_use]
    pub fn success_rate(&self) -> f64 {
        let total = self.total();
        if total == 0 {
            return 1.0;
        }
        self.succeeded.len() as f64 / total as f64
    }

    /// Returns the total number of operations.
    #[must_use]
    pub fn total(&self) -> usize {
        self.succeeded.len() + self.failed.len()
    }
}

/// Options for batch operations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BatchOptions {
    /// Timeout for shell operations that support timeout control.
    pub timeout: Duration,
    /// Refresh Recent Files display data once after successful Recent Files add operations.
    ///
    /// Batch mode coalesces this into one post-batch data-file refresh instead of
    /// deleting the Recent Files data file for every individual item.
    ///
    /// This option is intentionally add-only: `remove_items_batch` ignores it.
    /// Removal paths operate on the shell item directly, while deleting backing
    /// data files as a removal refresh would be more destructive than the
    /// requested per-item operation.
    pub force_update: bool,
}

impl Default for BatchOptions {
    fn default() -> Self {
        Self {
            timeout: Duration::from_secs(10),
            force_update: false,
        }
    }
}

fn unsupported_add(qa_type: QuickAccess) -> WincentError {
    WincentError::UnsupportedOperation(format!("Unsupported add operation for {:?}", qa_type))
}

fn unsupported_remove(qa_type: QuickAccess) -> WincentError {
    WincentError::UnsupportedOperation(format!("Unsupported remove operation for {:?}", qa_type))
}

fn get_items(qa_type: QuickAccess) -> WincentResult<Vec<String>> {
    match qa_type {
        QuickAccess::RecentFiles => query::get_recent_files(),
        QuickAccess::FrequentFolders => query::get_frequent_folders(),
        QuickAccess::All => query::get_quick_access_items(),
    }
}

fn check_item_exact(path: &str, qa_type: QuickAccess) -> WincentResult<bool> {
    let items = get_items(qa_type)?;
    Ok(items.iter().any(|item| paths_equal(item, path)))
}

fn add_item(path: &str, qa_type: QuickAccess, options: BatchOptions) -> WincentResult<()> {
    if matches!(qa_type, QuickAccess::All) {
        return Err(unsupported_add(qa_type));
    }

    // This preflight check is best-effort; Explorer state may still change
    // before the shell operation runs.
    if check_item_exact(path, qa_type)? {
        return Err(WincentError::AlreadyExists(path.to_string()));
    }

    match qa_type {
        QuickAccess::RecentFiles => add_to_recent_files_with_options(
            path,
            AddRecentFileOptions {
                // Batch force_update is handled once after all Recent Files adds.
                force_update: false,
            },
        ),
        QuickAccess::FrequentFolders => add_to_frequent_folders_with_timeout(path, options.timeout),
        QuickAccess::All => unreachable!(),
    }
}

fn remove_item(path: &str, qa_type: QuickAccess, options: BatchOptions) -> WincentResult<()> {
    if matches!(qa_type, QuickAccess::All) {
        return Err(unsupported_remove(qa_type));
    }

    // This preflight check is best-effort; Explorer state may still change
    // before the shell operation runs.
    if !check_item_exact(path, qa_type)? {
        return Err(WincentError::NotInRecent(path.to_string()));
    }

    match qa_type {
        QuickAccess::RecentFiles => remove_from_recent_files_with_timeout(path, options.timeout),
        QuickAccess::FrequentFolders => {
            remove_from_frequent_folders_with_timeout(path, options.timeout)
        }
        QuickAccess::All => unreachable!(),
    }
}

fn should_force_update_recent_files(options: BatchOptions, recent_files_succeeded: bool) -> bool {
    options.force_update && recent_files_succeeded
}

/// Adds multiple items to Quick Access, collecting per-item failures.
///
/// Each item performs a best-effort existence preflight before invoking the
/// shell operation. If Explorer changes between those two steps, batch results
/// may report `AlreadyExists` or a later operation error that reflects that
/// race. Successful Recent Files additions can be coalesced into one display
/// refresh with [`BatchOptions::force_update`].
pub(crate) fn add_items_batch(items: &[(String, QuickAccess)], options: BatchOptions) -> BatchResult {
    let mut succeeded = Vec::new();
    let mut failed = Vec::new();
    let mut recent_files_succeeded = false;

    for (path, qa_type) in items {
        match add_item(path, *qa_type, options) {
            Ok(()) => {
                if matches!(qa_type, QuickAccess::RecentFiles) {
                    recent_files_succeeded = true;
                }
                succeeded.push(path.clone());
            }
            Err(error) => failed.push((path.clone(), error)),
        }
    }

    if should_force_update_recent_files(options, recent_files_succeeded) {
        let _ = QuickAccessDataFiles::new().and_then(|data_files| data_files.remove_recent_file());
        let _ = refresh_explorer_window();
    }

    BatchResult { succeeded, failed }
}

/// Removes multiple items from Quick Access, collecting per-item failures.
///
/// Each item performs a best-effort existence preflight before invoking the
/// shell operation. If Explorer changes between those two steps, batch results
/// may report `NotInRecent` or a later operation error that reflects that race.
/// `BatchOptions::force_update` is intentionally ignored for removals because
/// the remove operations target shell items directly and should not delete
/// Recent Files or Frequent Folders backing data as a broad refresh side effect.
pub(crate) fn remove_items_batch(
    items: &[(String, QuickAccess)],
    options: BatchOptions,
) -> BatchResult {
    let mut succeeded = Vec::new();
    let mut failed = Vec::new();

    for (path, qa_type) in items {
        match remove_item(path, *qa_type, options) {
            Ok(()) => succeeded.push(path.clone()),
            Err(error) => failed.push((path.clone(), error)),
        }
    }

    BatchResult { succeeded, failed }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn batch_result_empty_is_complete_success() {
        let result = BatchResult {
            succeeded: vec![],
            failed: vec![],
        };

        assert!(result.is_complete_success());
        assert!(!result.has_partial_success());
        assert_eq!(result.success_rate(), 1.0);
        assert_eq!(result.total(), 0);
    }

    #[test]
    fn batch_result_reports_partial_success() {
        let result = BatchResult {
            succeeded: vec!["file1.txt".to_string(), "file2.txt".to_string()],
            failed: vec![(
                "file3.txt".to_string(),
                WincentError::NotInRecent("file3.txt".to_string()),
            )],
        };

        assert_eq!(result.total(), 3);
        assert!(!result.is_complete_success());
        assert!(result.has_partial_success());
        assert!((result.success_rate() - 0.666).abs() < 0.01);
    }

    #[test]
    fn force_update_requires_recent_files_success() {
        let options = BatchOptions {
            force_update: true,
            ..BatchOptions::default()
        };

        assert!(should_force_update_recent_files(options, true));
        assert!(!should_force_update_recent_files(options, false));
        assert!(!should_force_update_recent_files(
            BatchOptions::default(),
            true
        ));
    }
}
