use crate::{
    empty,
    handle::{
        add_to_frequent_folders_with_timeout, add_to_recent_files_with_options,
        remove_from_frequent_folders_with_timeout, remove_from_recent_files_with_timeout,
        AddRecentFileOptions,
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

    fn get_items(&self, qa_type: QuickAccess) -> WincentResult<Vec<String>>;

    fn add_recent_file(&self, path: &str) -> WincentResult<()>;
    fn add_recent_file_and_refresh(&self, path: &str) -> WincentResult<()>;
    fn add_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()>;

    fn remove_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()>;
    fn remove_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()>;

    fn delete_recent_links_for_target(&self, path: &str, timeout: Duration) -> WincentResult<()>;

    /// Batch-add display refresh: remove Recent Files backing data and refresh Explorer.
    /// This intentionally keeps the current compound batch refresh semantics.
    fn refresh_recent_files_display(&self) -> WincentResult<()>;

    fn clear_recent_files(&self) -> WincentResult<()>;
    fn clear_frequent_folders_jumplist(&self) -> WincentResult<()>;
    fn refresh_explorer(&self) -> WincentResult<()>;
}

#[derive(Debug, Default)]
pub(crate) struct SystemQuickAccessBackend;

impl QuickAccessBackend for SystemQuickAccessBackend {
    fn validate_path(&self, path: &str, expected: PathType) -> WincentResult<()> {
        validate_path(path, expected)
    }

    fn get_items(&self, qa_type: QuickAccess) -> WincentResult<Vec<String>> {
        match qa_type {
            QuickAccess::RecentFiles => query::get_recent_files(),
            QuickAccess::FrequentFolders => query::get_frequent_folders(),
            QuickAccess::All => query::get_quick_access_items(),
        }
    }

    fn add_recent_file(&self, path: &str) -> WincentResult<()> {
        add_to_recent_files_with_options(
            path,
            AddRecentFileOptions {
                force_update: false,
            },
        )
    }

    fn add_recent_file_and_refresh(&self, path: &str) -> WincentResult<()> {
        add_to_recent_files_with_options(path, AddRecentFileOptions { force_update: true })
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

    fn refresh_recent_files_display(&self) -> WincentResult<()> {
        QuickAccessDataFiles::new()?.remove_recent_file()?;
        refresh_explorer_window()
    }

    fn clear_recent_files(&self) -> WincentResult<()> {
        empty::empty_recent_files()
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
