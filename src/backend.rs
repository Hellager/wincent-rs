use crate::{
    destlist::{
        frequent_folders_dest_path, parse_file, recent_files_dest_path, DestListEntry,
        DEFAULT_RECENT_FILES_RESULT_LIMIT,
    },
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
    fn add_destlist_entry(
        &self,
        path: &str,
        qa_type: QuickAccess,
        timeout: Duration,
    ) -> WincentResult<()> {
        let _ = (path, qa_type, timeout);
        Err(crate::error::WincentError::UnsupportedOperation(
            "direct DestList addition is not implemented by this backend".to_string(),
        ))
    }

    fn remove_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()>;
    fn remove_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()>;
    fn remove_destlist_entry(&self, _path: &str, _qa_type: QuickAccess) -> WincentResult<()> {
        Err(crate::error::WincentError::UnsupportedOperation(
            "direct DestList removal is not implemented by this backend".to_string(),
        ))
    }
    fn update_destlist_entry(
        &self,
        _path: &str,
        _qa_type: QuickAccess,
        _update: crate::destlist::DestListItemUpdate,
    ) -> WincentResult<()> {
        Err(crate::error::WincentError::UnsupportedOperation(
            "direct DestList update is not implemented by this backend".to_string(),
        ))
    }

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

/// Tries the CFB rule parser first so the returned path order and filtering
/// match the custom CAutoList rule result instead of a live Shell query.
///
/// Recent Files results are additionally truncated to the same default
/// 20-item window the observed Shell query uses. Any parse, I/O, or path
/// resolution error falls back to the caller so the original Shell query
/// path still runs.
fn rule_based_items(qa_type: QuickAccess) -> WincentResult<Vec<String>> {
    let (parsed, apply_limit) = match qa_type {
        QuickAccess::RecentFiles => (parse_file(recent_files_dest_path()?)?, true),
        QuickAccess::FrequentFolders => (parse_file(frequent_folders_dest_path()?)?, false),
        QuickAccess::All => {
            let recent = rule_based_items(QuickAccess::RecentFiles)?;
            let frequent = rule_based_items(QuickAccess::FrequentFolders)?;
            let mut merged = recent;
            for path in frequent {
                if !merged
                    .iter()
                    .any(|existing| crate::utils::paths_equal(existing, &path))
                {
                    merged.push(path);
                }
            }
            return Ok(merged);
        }
    };

    let mut entries = parsed.rule_entries();
    if apply_limit {
        entries.truncate(DEFAULT_RECENT_FILES_RESULT_LIMIT);
    }
    Ok(entries
        .into_iter()
        .map(|entry| entry.path().to_string())
        .collect())
}

#[derive(Debug, Default)]
pub(crate) struct SystemQuickAccessBackend;

impl QuickAccessBackend for SystemQuickAccessBackend {
    fn validate_path(&self, path: &str, expected: PathType) -> WincentResult<()> {
        validate_path(path, expected)
    }

    fn get_items(&self, qa_type: QuickAccess, timeout: Duration) -> WincentResult<Vec<String>> {
        rule_based_items(qa_type).or_else(|_| query::query_recent_with_timeout(qa_type, timeout))
    }

    fn add_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        add_file_to_recent_native(path, timeout)
    }

    fn add_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        add_to_frequent_folders_with_timeout(path, timeout)
    }

    fn add_destlist_entry(
        &self,
        path: &str,
        qa_type: QuickAccess,
        timeout: Duration,
    ) -> WincentResult<()> {
        crate::destlist::add::add_entry_path(qa_type, path, timeout)
    }

    fn remove_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        remove_from_recent_files_with_timeout(path, timeout)
    }

    fn remove_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        remove_from_frequent_folders_with_timeout(path, timeout)
    }

    fn remove_destlist_entry(&self, path: &str, qa_type: QuickAccess) -> WincentResult<()> {
        crate::destlist::remove::remove_entry_path(qa_type, path)
    }

    fn update_destlist_entry(
        &self,
        path: &str,
        qa_type: QuickAccess,
        update: crate::destlist::DestListItemUpdate,
    ) -> WincentResult<()> {
        crate::destlist::update::update_entry_path(qa_type, path, update)
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
