use crate::{
    destlist::{
        frequent_folders_dest_path, parse_file_if_exists, recent_files_dest_path,
        AutomaticDestinations, DEFAULT_RECENT_FILES_RESULT_LIMIT,
    },
    lock::{MutationPermit, QuickAccessLockTarget},
    path::PathType,
    shell::{
        clear,
        mutation::{
            add_file_to_recent_native, add_to_frequent_folders_with_timeout,
            remove_from_recent_files_with_timeout, remove_pinned_frequent_folder_with_timeout,
            remove_unpinned_frequent_folder_with_timeout,
        },
    },
    QuickAccess, WincentResult,
};
use std::time::Duration;

pub(crate) trait QuickAccessBackend: Send + Sync {
    fn uses_system_mutation_gate(&self) -> bool {
        false
    }

    fn path_type(&self, path: &str) -> WincentResult<PathType> {
        crate::path::detect_path_type(path)
    }

    fn get_items(&self, qa_type: QuickAccess, timeout: Duration) -> WincentResult<Vec<String>>;

    fn add_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()>;
    fn add_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()>;
    fn remove_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()>;
    fn frequent_folder_pin_status(
        &self,
        _path: &str,
    ) -> WincentResult<crate::destlist::FrequentFolderPinStatus> {
        Ok(crate::destlist::FrequentFolderPinStatus::NotFound)
    }
    fn remove_pinned_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()>;
    fn remove_unpinned_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()>;
    fn clear_recent_files(&self, timeout: Duration) -> WincentResult<()>;
    fn clear_frequent_folders_jumplist(&self, timeout: Duration) -> WincentResult<()>;
    fn notify_shell_change(&self);
    fn refresh_shell_views(&self, timeout: Duration) -> WincentResult<()>;

    fn refresh(&self, timeout: Duration) -> WincentResult<()> {
        self.notify_shell_change();
        self.refresh_shell_views(timeout)
    }
}

/// Parses the CFB DestList and applies Explorer's Quick Access UI rules.
///
/// Ordering, metadata filtering, deduplication, and the Recent Files result
/// limit come from the recovered CAutoList rules. Hidden-item visibility comes
/// from Explorer's Folder Options setting and the attributes serialized in the
/// numeric Shell Link stream when the item was added.
pub(crate) fn ui_visible_items(qa_type: QuickAccess) -> WincentResult<Vec<String>> {
    match qa_type {
        QuickAccess::RecentFiles => ui_visible_items_from_path(recent_files_dest_path()?, true),
        QuickAccess::FrequentFolders => {
            ui_visible_items_from_path(frequent_folders_dest_path()?, false)
        }
        QuickAccess::All => {
            let recent = ui_visible_items(QuickAccess::RecentFiles)?;
            let frequent = ui_visible_items(QuickAccess::FrequentFolders)?;
            let mut merged = recent;
            for path in frequent {
                if !merged
                    .iter()
                    .any(|existing| crate::path::paths_equal(existing, &path))
                {
                    merged.push(path);
                }
            }
            Ok(merged)
        }
    }
}

fn ui_visible_items_from_path(
    path: impl AsRef<std::path::Path>,
    apply_limit: bool,
) -> WincentResult<Vec<String>> {
    let Some(parsed) = parse_file_if_exists(path)? else {
        return Ok(Vec::new());
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
    fn uses_system_mutation_gate(&self) -> bool {
        true
    }

    /// Reads Quick Access items directly from Explorer's on-disk DestList and
    /// applies the recovered Quick Access UI visibility rules.
    fn get_items(&self, qa_type: QuickAccess, _timeout: Duration) -> WincentResult<Vec<String>> {
        ui_visible_items(qa_type)
    }

    fn add_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        let permits = MutationPermit::acquire(QuickAccessLockTarget::All)?;
        add_file_to_recent_native(path, timeout, permits)
    }

    fn add_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        let permit = MutationPermit::acquire_one(QuickAccessLockTarget::FrequentFolders)?;
        add_to_frequent_folders_with_timeout(path, timeout, permit)
    }

    fn remove_recent_file(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        let permit = MutationPermit::acquire_one(QuickAccessLockTarget::RecentFiles)?;
        remove_from_recent_files_with_timeout(path, timeout, permit)
    }

    fn frequent_folder_pin_status(
        &self,
        path: &str,
    ) -> WincentResult<crate::destlist::FrequentFolderPinStatus> {
        crate::destlist::frequent_folder_pin_status(path)
    }

    fn remove_pinned_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        let permit = MutationPermit::acquire_one(QuickAccessLockTarget::FrequentFolders)?;
        remove_pinned_frequent_folder_with_timeout(path, timeout, permit)
    }

    fn remove_unpinned_frequent_folder(&self, path: &str, timeout: Duration) -> WincentResult<()> {
        let permit = MutationPermit::acquire_one(QuickAccessLockTarget::FrequentFolders)?;
        remove_unpinned_frequent_folder_with_timeout(path, timeout, permit)
    }

    fn clear_recent_files(&self, timeout: Duration) -> WincentResult<()> {
        let permit = MutationPermit::acquire_one(QuickAccessLockTarget::RecentFiles)?;
        clear::clear_recent_files(timeout, permit)
    }

    fn clear_frequent_folders_jumplist(&self, timeout: Duration) -> WincentResult<()> {
        let _permit = MutationPermit::acquire_one(QuickAccessLockTarget::FrequentFolders)?;
        clear::clear_frequent_folders(timeout)
    }

    fn notify_shell_change(&self) {
        crate::shell::explorer::notify_shell_change();
    }

    fn refresh_shell_views(&self, timeout: Duration) -> WincentResult<()> {
        crate::shell::explorer::refresh_shell_views_with_timeout(timeout)
    }
}

// QuickAccessLock is intentionally not routed through QuickAccessBackend in this pass.
// The DI layer currently covers query, mutation, refresh, and empty workflows.
// Locking still uses the native backing-file handle path directly.

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
