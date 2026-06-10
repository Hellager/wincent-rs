use crate::{
    empty,
    handle::{
        add_file_to_recent_native, add_to_frequent_folders_with_timeout,
        remove_from_frequent_folders_with_timeout, remove_from_recent_files_with_timeout,
    },
    query, recent_links,
    script_executor::QuickAccessDataFiles,
    utils::{refresh_explorer_window, validate_path, PathType},
    QuickAccess, WincentResult,
};
use std::path::PathBuf;
use std::time::Duration;

pub(crate) trait QuickAccessBackend: Send + Sync {
    fn validate_path(&self, path: &str, expected: PathType) -> WincentResult<()>;

    fn get_items(&self, qa_type: QuickAccess, timeout: Duration) -> WincentResult<Vec<String>>;

    fn add_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()>;
    fn add_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()>;

    fn remove_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()>;
    fn remove_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()>;

    fn delete_recent_links_for_target(&self, path: &str, timeout: Duration) -> WincentResult<()>;

    /// Deletes Explorer's Recent Files backing data so Explorer can rebuild it.
    fn delete_recent_files_backing_data(&self) -> WincentResult<()>;

    fn clear_recent_files(&self, timeout: Duration) -> WincentResult<()>;
    fn clear_frequent_folders_jumplist(&self) -> WincentResult<()>;
    fn refresh_explorer(&self) -> WincentResult<()>;
}

#[derive(Debug, Default)]
pub(crate) struct SystemQuickAccessBackend;

impl QuickAccessBackend for SystemQuickAccessBackend {
    fn validate_path(&self, path: &str, expected: PathType) -> WincentResult<()> {
        validate_path(path, expected)
    }

    fn get_items(&self, qa_type: QuickAccess, timeout: Duration) -> WincentResult<Vec<String>> {
        query::query_recent_with_timeout(qa_type, timeout)
    }

    fn add_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        add_file_to_recent_native(path, timeout)
    }

    fn add_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        add_to_frequent_folders_with_timeout(path, timeout)
    }

    fn remove_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        remove_from_recent_files_with_timeout(path, timeout)
    }

    fn remove_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        remove_from_frequent_folders_with_timeout(path, timeout)
    }

    fn delete_recent_links_for_target(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        recent_links::delete_recent_links_for_target(path, timeout).map(|_: Vec<PathBuf>| ())
    }

    fn delete_recent_files_backing_data(&self) -> WincentResult<()> {
        QuickAccessDataFiles::new()?.remove_recent_file()
    }

    fn clear_recent_files(&self, timeout: Duration) -> WincentResult<()> {
        empty::empty_recent_files_with_api(timeout)
    }

    fn clear_frequent_folders_jumplist(&self) -> WincentResult<()> {
        empty::empty_user_folders_with_jumplist_file()
    }

    fn refresh_explorer(&self) -> WincentResult<()> {
        refresh_explorer_window()
    }
}

// QuickAccessLock is intentionally not routed through QuickAccessBackend in this pass.
// The DI layer currently covers query, mutation, batch, recent-link cleanup, and empty
// workflows. Locking still uses the native backing-file handle path directly.
