use crate::{
    destlist::{
        frequent_folders_dest_path, parse_file, recent_files_dest_path, AutomaticDestinations,
        DestListEntry, DEFAULT_RECENT_FILES_RESULT_LIMIT,
    },
    empty,
    handle::{
        add_file_to_recent_native, add_to_frequent_folders_with_timeout,
        remove_from_frequent_folders_with_timeout, remove_from_recent_files_with_timeout,
    },
    recent_links,
    script_executor::QuickAccessDataFiles,
    utils::{get_windows_recent_folder, validate_path, PathType},
    QuickAccess, WincentResult,
};
use std::path::{Path, PathBuf};
use std::time::Duration;

pub(crate) trait QuickAccessBackend: Send + Sync {
    fn path_type(&self, path: &str) -> WincentResult<PathType> {
        crate::utils::detect_path_type(path)
    }

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
    fn clear_frequent_folders_jumplist(&self, timeout: Duration) -> WincentResult<()>;
    fn refresh_explorer(&self, timeout: Duration) -> WincentResult<()>;
}

/// Parses the CFB DestList and applies Explorer's Quick Access UI rules.
///
/// Ordering, metadata filtering, deduplication, and the Recent Files result
/// limit come from the recovered CAutoList rules. Hidden-item visibility comes
/// from Explorer's Folder Options setting and the attributes serialized in the
/// numeric Shell Link stream when the item was added.
fn ui_visible_items(qa_type: QuickAccess) -> WincentResult<Vec<String>> {
    let (parsed, apply_limit) = match qa_type {
        QuickAccess::RecentFiles => (parse_file(recent_files_dest_path()?)?, true),
        QuickAccess::FrequentFolders => (parse_file(frequent_folders_dest_path()?)?, false),
        QuickAccess::All => {
            let recent = ui_visible_items(QuickAccess::RecentFiles)?;
            let frequent = ui_visible_items(QuickAccess::FrequentFolders)?;
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

    let show_hidden_items = crate::visible::explorer_shows_hidden_items()?;
    Ok(ui_item_paths(&parsed, apply_limit, show_hidden_items))
}

fn ui_item_paths(
    parsed: &AutomaticDestinations,
    apply_limit: bool,
    show_hidden_items: bool,
) -> Vec<String> {
    let mut entries = parsed.ui_entries(show_hidden_items);
    if apply_limit {
        entries.truncate(DEFAULT_RECENT_FILES_RESULT_LIMIT);
    }
    entries
        .into_iter()
        .map(|entry| entry.path().to_string())
        .collect()
}

#[derive(Debug, Default)]
pub(crate) struct SystemQuickAccessBackend;

impl QuickAccessBackend for SystemQuickAccessBackend {
    fn validate_path(&self, path: &str, expected: PathType) -> WincentResult<()> {
        validate_path(path, expected)
    }

    /// Reads Quick Access items directly from Explorer's on-disk DestList and
    /// applies the recovered Quick Access UI visibility rules.
    fn get_items(&self, qa_type: QuickAccess, _timeout: Duration) -> WincentResult<Vec<String>> {
        ui_visible_items(qa_type)
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

    fn clear_frequent_folders_jumplist(&self, timeout: Duration) -> WincentResult<()> {
        empty::clear_frequent_folders_with_destlist_reset(timeout)
    }

    fn refresh_explorer(&self, timeout: Duration) -> WincentResult<()> {
        crate::explorer_window::refresh_explorer_with_timeout(timeout)
    }
}

// QuickAccessLock is intentionally not routed through QuickAccessBackend in this pass.
// The DI layer currently covers query, mutation, batch, recent-link cleanup, and empty
// workflows. Locking still uses the native backing-file handle path directly.

#[cfg(test)]
mod tests {
    use super::*;
    use crate::destlist::test_support::{automatic_destinations_for_test, destlist_entry_for_test};
    use crate::destlist::DestListKind;

    fn set_hidden_attribute(parsed: &mut AutomaticDestinations, stream_name: &str) {
        let stream = parsed
            .cfb_info
            .streams
            .iter_mut()
            .find(|stream| stream.name == stream_name)
            .expect("test stream");
        stream.data[0x18..0x1c].copy_from_slice(&0x2u32.to_le_bytes());
    }

    #[test]
    fn recent_ui_limit_is_applied_after_hidden_filtering() {
        let entries: Vec<_> = (1..=21)
            .map(|number| {
                destlist_entry_for_test(
                    &format!("C:\\Recent\\item-{number:02}.txt"),
                    number,
                    0,
                    -1,
                    1,
                    None,
                )
            })
            .collect();
        let streams: Vec<_> = (1..=21).map(|number| format!("{number:x}")).collect();
        let mut parsed =
            automatic_destinations_for_test(DestListKind::RecentFiles, entries, &streams);
        set_hidden_attribute(&mut parsed, "1");

        let paths = ui_item_paths(&parsed, true, false);

        assert_eq!(paths.len(), DEFAULT_RECENT_FILES_RESULT_LIMIT);
        assert!(!paths.iter().any(|path| path.ends_with("item-01.txt")));
        assert!(paths.iter().any(|path| path.ends_with("item-21.txt")));
    }

    #[test]
    fn frequent_ui_filter_removes_hidden_pinned_item() {
        let entries = vec![
            destlist_entry_for_test("C:\\Visible", 1, 0, -1, 0, Some(0)),
            destlist_entry_for_test("C:\\Hidden", 2, 0, -1, 0, Some(1)),
        ];
        let streams = vec!["1".to_string(), "2".to_string()];
        let mut parsed =
            automatic_destinations_for_test(DestListKind::FrequentFolders, entries, &streams);
        set_hidden_attribute(&mut parsed, "2");

        assert_eq!(ui_item_paths(&parsed, false, false), vec!["C:\\Visible"]);
        assert_eq!(
            ui_item_paths(&parsed, false, true),
            vec!["C:\\Visible", "C:\\Hidden"]
        );
    }
}
