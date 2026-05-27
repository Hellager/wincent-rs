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
///
/// Batch operations are best-effort: every input item is attempted, and
/// failures are stored as `BatchFailure` values instead of aborting the whole
/// batch.
///
/// When returned by [`crate::manager::QuickAccessManager`] batch methods,
/// failures are grouped by processing phase: path conversion failures detected
/// by the manager are listed first, followed by failures from the underlying
/// shell operations. The failed list is therefore not guaranteed to preserve
/// the original input order.
///
/// # Examples
///
/// ```rust,no_run
/// use wincent::prelude::*;
///
/// let manager = QuickAccessManager::new();
/// let items = [
///     QuickAccessItem::recent_file("C:\\Work\\report.docx"),
///     QuickAccessItem::recent_file("Z:\\missing.txt"),
/// ];
///
/// let result = manager.add_items_batch(&items, BatchOptions::new());
/// if result.has_partial_success() {
///     println!("{} of {} succeeded", result.succeeded().len(), result.total());
/// }
///
/// for failure in result.failed() {
///     eprintln!("{} failed: {}", failure.path(), failure.error());
/// }
/// ```
#[derive(Debug, Default)]
pub struct BatchResult {
    /// Successfully processed items.
    succeeded: Vec<String>,
    /// Failed items with error details.
    ///
    /// `AlreadyExists` and `NotInQuickAccess` may come from best-effort preflight
    /// checks. Explorer state can still change between the preflight query and
    /// the shell operation, so callers should treat those errors as a snapshot
    /// of the attempted operation rather than a durable global truth.
    failed: Vec<BatchFailure>,
}

/// Failed item from a batch operation.
///
/// The error describes the operation as it was attempted. For example,
/// [`WincentError::AlreadyExists`] and [`WincentError::NotInQuickAccess`] can
/// come from best-effort preflight checks, and Explorer state may have changed
/// by the time the caller inspects the result.
#[derive(Debug)]
pub struct BatchFailure {
    path: String,
    error: WincentError,
}

impl BatchFailure {
    pub(crate) fn new(path: String, error: WincentError) -> Self {
        Self { path, error }
    }

    /// Path that failed to process.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Error returned for this path.
    #[must_use]
    pub fn error(&self) -> &WincentError {
        &self.error
    }
}

impl BatchResult {
    pub(crate) fn new(succeeded: Vec<String>, failed: Vec<BatchFailure>) -> Self {
        Self { succeeded, failed }
    }

    /// Successfully processed items.
    #[must_use]
    pub fn succeeded(&self) -> &[String] {
        &self.succeeded
    }

    /// Failed items with error details.
    ///
    /// Results produced through [`crate::manager::QuickAccessManager`] list
    /// manager-side path conversion failures before operation failures. Do not
    /// rely on this slice preserving the original input order.
    #[must_use]
    pub fn failed(&self) -> &[BatchFailure] {
        &self.failed
    }

    /// Consumes the result and returns its raw parts.
    #[must_use]
    pub fn into_parts(self) -> (Vec<String>, Vec<BatchFailure>) {
        (self.succeeded, self.failed)
    }

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
///
/// `BatchOptions` controls batch behavior. Shell operation timeout is supplied
/// by [`crate::manager::QuickAccessManager`] configuration rather than this
/// type.
///
/// # Examples
///
/// ```rust
/// use wincent::BatchOptions;
///
/// let options = BatchOptions::new().refresh_recent_files();
///
/// assert!(options.force_update());
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BatchOptions {
    /// Refresh Recent Files display data once after successful Recent Files add operations.
    ///
    /// Batch mode coalesces this into one post-batch data-file refresh instead of
    /// deleting the Recent Files data file for every individual item.
    ///
    /// This option is intentionally add-only: `remove_items_batch` ignores it.
    /// Removal paths operate on the shell item directly, while deleting backing
    /// data files as a removal refresh would be more destructive than the
    /// requested per-item operation.
    force_update: bool,
}

impl Default for BatchOptions {
    fn default() -> Self {
        Self {
            force_update: false,
        }
    }
}

impl BatchOptions {
    /// Creates default batch options.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether Recent Files display data is refreshed after successful additions.
    #[must_use]
    pub fn force_update(&self) -> bool {
        self.force_update
    }

    /// Sets whether Recent Files display data is refreshed after successful additions.
    #[must_use]
    pub fn with_force_update(mut self, force_update: bool) -> Self {
        self.force_update = force_update;
        self
    }

    /// Refreshes Recent Files display data after successful additions.
    #[must_use]
    pub fn refresh_recent_files(self) -> Self {
        self.with_force_update(true)
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

fn add_item(path: &str, qa_type: QuickAccess, timeout: Duration) -> WincentResult<()> {
    match qa_type {
        QuickAccess::RecentFiles => add_recent_file(path),
        QuickAccess::FrequentFolders => add_frequent_folder(path, timeout),
        unsupported => Err(unsupported_add(unsupported)),
    }
}

fn remove_item(path: &str, qa_type: QuickAccess, timeout: Duration) -> WincentResult<()> {
    match qa_type {
        QuickAccess::RecentFiles => remove_recent_file(path, timeout),
        QuickAccess::FrequentFolders => remove_frequent_folder(path, timeout),
        unsupported => Err(unsupported_remove(unsupported)),
    }
}

fn add_recent_file(path: &str) -> WincentResult<()> {
    ensure_not_present(path, QuickAccess::RecentFiles)?;
    add_to_recent_files_with_options(
        path,
        AddRecentFileOptions {
            // Batch force_update is handled once after all Recent Files adds.
            force_update: false,
        },
    )
}

fn add_frequent_folder(path: &str, timeout: Duration) -> WincentResult<()> {
    ensure_not_present(path, QuickAccess::FrequentFolders)?;
    add_to_frequent_folders_with_timeout(path, timeout)
}

fn remove_recent_file(path: &str, timeout: Duration) -> WincentResult<()> {
    ensure_present(path, QuickAccess::RecentFiles)?;
    remove_from_recent_files_with_timeout(path, timeout)
}

fn remove_frequent_folder(path: &str, timeout: Duration) -> WincentResult<()> {
    ensure_present(path, QuickAccess::FrequentFolders)?;
    remove_from_frequent_folders_with_timeout(path, timeout)
}

fn ensure_not_present(path: &str, qa_type: QuickAccess) -> WincentResult<()> {
    // This preflight check is best-effort; Explorer state may still change
    // before the shell operation runs.
    if check_item_exact(path, qa_type)? {
        return Err(WincentError::already_exists(path, qa_type));
    }

    Ok(())
}

fn ensure_present(path: &str, qa_type: QuickAccess) -> WincentResult<()> {
    // This preflight check is best-effort; Explorer state may still change
    // before the shell operation runs.
    if !check_item_exact(path, qa_type)? {
        return Err(WincentError::not_in_quick_access(path, qa_type));
    }

    Ok(())
}

fn should_force_update_recent_files(options: BatchOptions, recent_files_succeeded: bool) -> bool {
    options.force_update() && recent_files_succeeded
}

/// Adds multiple items to Quick Access, collecting per-item failures.
///
/// Each item performs a best-effort existence preflight before invoking the
/// shell operation. If Explorer changes between those two steps, batch results
/// may report `AlreadyExists` or a later operation error that reflects that
/// race. Successful Recent Files additions can be coalesced into one display
/// refresh with [`BatchOptions::force_update`].
pub(crate) fn add_items_batch(
    items: &[(String, QuickAccess)],
    options: BatchOptions,
    timeout: Duration,
) -> BatchResult {
    let mut succeeded = Vec::new();
    let mut failed = Vec::new();
    let mut recent_files_succeeded = false;

    for (path, qa_type) in items {
        match add_item(path, *qa_type, timeout) {
            Ok(()) => {
                if matches!(qa_type, QuickAccess::RecentFiles) {
                    recent_files_succeeded = true;
                }
                succeeded.push(path.clone());
            }
            Err(error) => failed.push(BatchFailure::new(path.clone(), error)),
        }
    }

    if should_force_update_recent_files(options, recent_files_succeeded) {
        let _ = QuickAccessDataFiles::new().and_then(|data_files| data_files.remove_recent_file());
        let _ = refresh_explorer_window();
    }

    BatchResult::new(succeeded, failed)
}

/// Removes multiple items from Quick Access, collecting per-item failures.
///
/// Each item performs a best-effort existence preflight before invoking the
/// shell operation. If Explorer changes between those two steps, batch results
/// may report `NotInQuickAccess` or a later operation error that reflects that race.
/// `BatchOptions::force_update` is intentionally ignored for removals because
/// the remove operations target shell items directly and should not delete
/// Recent Files or Frequent Folders backing data as a broad refresh side effect.
pub(crate) fn remove_items_batch(
    items: &[(String, QuickAccess)],
    timeout: Duration,
) -> BatchResult {
    let mut succeeded = Vec::new();
    let mut failed = Vec::new();

    for (path, qa_type) in items {
        match remove_item(path, *qa_type, timeout) {
            Ok(()) => succeeded.push(path.clone()),
            Err(error) => failed.push(BatchFailure::new(path.clone(), error)),
        }
    }

    BatchResult::new(succeeded, failed)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn batch_result_empty_is_complete_success() {
        let result = BatchResult::new(vec![], vec![]);

        assert!(result.is_complete_success());
        assert!(!result.has_partial_success());
        assert_eq!(result.success_rate(), 1.0);
        assert_eq!(result.total(), 0);
    }

    #[test]
    fn batch_result_reports_partial_success() {
        let result = BatchResult::new(
            vec!["file1.txt".to_string(), "file2.txt".to_string()],
            vec![BatchFailure::new(
                "file3.txt".to_string(),
                WincentError::NotInQuickAccess {
                    path: "file3.txt".to_string(),
                    qa_type: QuickAccess::RecentFiles,
                },
            )],
        );

        assert_eq!(result.total(), 3);
        assert!(!result.is_complete_success());
        assert!(result.has_partial_success());
        assert!((result.success_rate() - 0.666).abs() < 0.01);
    }

    #[test]
    fn force_update_requires_recent_files_success() {
        let options = BatchOptions::default().with_force_update(true);

        assert!(should_force_update_recent_files(options, true));
        assert!(!should_force_update_recent_files(options, false));
        assert!(!should_force_update_recent_files(
            BatchOptions::default(),
            true
        ));
    }
}
