use crate::{
    destlist::{frequent_folders_dest_path, parse_file, DestListEntry},
    empty,
    handle::{
        add_file_to_recent_native, add_to_frequent_folders_with_timeout,
        remove_from_frequent_folders_with_timeout, remove_from_recent_files_with_timeout,
    },
    query, recent_links,
    script_executor::QuickAccessDataFiles,
    utils::{
        get_windows_recent_folder, refresh_explorer_window_with_timeout, validate_path, PathType,
    },
    QuickAccess, WincentResult,
};
use std::path::{Path, PathBuf};
use std::time::Duration;

pub(crate) trait QuickAccessBackend: Send + Sync {
    fn validate_path(&self, path: &str, expected: PathType) -> WincentResult<()>;

    fn get_items(&self, qa_type: QuickAccess, timeout: Duration) -> WincentResult<Vec<String>>;

    fn add_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()>;
    fn add_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()>;

    fn remove_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()>;
    fn remove_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()>;

    fn delete_recent_links_for_target(&self, path: &str, timeout: Duration) -> WincentResult<()>;
    fn list_recent_lnk_files(&self) -> WincentResult<Vec<PathBuf>>;
    fn delete_lnk_file(&self, path: &Path) -> WincentResult<()>;
    fn resolve_lnk_with_type(
        &self,
        path: &Path,
        timeout: Duration,
    ) -> WincentResult<Option<recent_links::LnkResolution>>;

    /// Deletes Explorer's Recent Files backing data so Explorer can rebuild it.
    fn delete_recent_files_backing_data(&self) -> WincentResult<()>;
    fn delete_frequent_folders_backing_file(&self) -> WincentResult<()>;
    fn wait_for_frequent_folders_rebuild(
        &self,
        timeout: Duration,
    ) -> WincentResult<Vec<DestListEntry>>;

    fn clear_recent_files(&self, timeout: Duration) -> WincentResult<()>;
    fn clear_frequent_folders_jumplist(&self) -> WincentResult<()>;
    fn refresh_explorer(&self, timeout: Duration) -> WincentResult<()>;
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

    fn list_recent_lnk_files(&self) -> WincentResult<Vec<PathBuf>> {
        recent_links::recent_lnk_paths(Path::new(&get_windows_recent_folder()?))
    }

    fn delete_lnk_file(&self, path: &Path) -> WincentResult<()> {
        std::fs::remove_file(path).map_err(crate::error::WincentError::Io)
    }

    fn resolve_lnk_with_type(
        &self,
        path: &Path,
        timeout: Duration,
    ) -> WincentResult<Option<recent_links::LnkResolution>> {
        recent_links::resolve_lnk_with_type(path, timeout)
    }

    fn delete_recent_files_backing_data(&self) -> WincentResult<()> {
        QuickAccessDataFiles::new()?.remove_recent_file()
    }

    fn delete_frequent_folders_backing_file(&self) -> WincentResult<()> {
        match std::fs::remove_file(frequent_folders_dest_path()?) {
            Ok(()) => Ok(()),
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
            Err(error) => Err(crate::error::WincentError::Io(error)),
        }
    }

    fn wait_for_frequent_folders_rebuild(
        &self,
        timeout: Duration,
    ) -> WincentResult<Vec<DestListEntry>> {
        let path = frequent_folders_dest_path()?;
        let started = std::time::Instant::now();
        let mut last_error = None;

        loop {
            if path.exists() {
                match parse_file(&path) {
                    Ok(parsed) => return Ok(parsed.dest_list().entries().to_vec()),
                    Err(error) => last_error = Some(error),
                }
            }

            if started.elapsed() >= timeout {
                return Err(last_error.unwrap_or_else(|| {
                    crate::error::WincentError::Timeout(format!(
                        "Frequent Folders backing file was not rebuilt within {}s",
                        timeout.as_secs_f64()
                    ))
                }));
            }

            std::thread::sleep(Duration::from_millis(100));
        }
    }

    fn clear_recent_files(&self, timeout: Duration) -> WincentResult<()> {
        empty::empty_recent_files_with_api(timeout)
    }

    fn clear_frequent_folders_jumplist(&self) -> WincentResult<()> {
        empty::empty_user_folders_with_jumplist_file()
    }

    fn refresh_explorer(&self, timeout: Duration) -> WincentResult<()> {
        refresh_explorer_window_with_timeout(timeout)
    }
}

// QuickAccessLock is intentionally not routed through QuickAccessBackend in this pass.
// The DI layer currently covers query, mutation, batch, recent-link cleanup, and empty
// workflows. Locking still uses the native backing-file handle path directly.
