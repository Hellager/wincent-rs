//! Native Quick Access clearing primitives.

use crate::{
    destlist::{frequent_folders_dest_path, parse_file},
    error::WincentError,
    lock::MutationPermit,
    path::paths_equal,
    recent_links, WincentResult,
};
use std::{io::ErrorKind, path::Path, time::Duration};
use windows::Win32::UI::Shell::SHAddToRecentDocs;

/// SHARD_PATHW with a null data pointer clears all recent documents.
const SHARD_PATHW: u32 = 0x0000_0003;

/// Clears the Windows Recent Files list via a dedicated STA thread.
pub(crate) fn clear_recent_files(timeout: Duration, permit: MutationPermit) -> WincentResult<()> {
    super::thread::run_on_sta_thread_with_permit(
        || {
            // SAFETY: SHARD_PATHW with a null data pointer is a documented
            // Windows API pattern for clearing the recent documents list.
            unsafe { SHAddToRecentDocs(SHARD_PATHW, None) };
            Ok(())
        },
        timeout,
        permit,
    )
}

/// Clears Frequent Folders backing data and matching Recent shortcuts.
pub(crate) fn clear_frequent_folders(timeout: Duration) -> WincentResult<()> {
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

#[cfg(test)]
mod tests {
    use super::*;
    use std::cell::RefCell;

    #[test]
    fn unique_paths_ignores_empty_and_equivalent_paths() {
        let paths = ["", "C:\\Work", "c:/work/", "C:\\Other"];

        assert_eq!(unique_paths(paths), vec!["C:\\Work", "C:\\Other"]);
    }

    #[test]
    fn frequent_clear_deletes_shortcuts_before_destlist() -> WincentResult<()> {
        let calls = RefCell::new(Vec::new());
        let paths = vec!["C:\\Work".to_string()];

        clear_frequent_folders_data(
            &paths,
            |targets| {
                assert_eq!(targets, paths);
                calls.borrow_mut().push("shortcuts");
                Ok(())
            },
            || {
                calls.borrow_mut().push("destlist");
                Ok(())
            },
        )?;

        assert_eq!(*calls.borrow(), ["shortcuts", "destlist"]);
        Ok(())
    }

    #[test]
    fn frequent_clear_preserves_destlist_when_shortcut_cleanup_fails() {
        let destlist_deleted = RefCell::new(false);
        let error = clear_frequent_folders_data(
            &["C:\\Work".to_string()],
            |_| Err(WincentError::SystemError("cleanup failed".to_string())),
            || {
                *destlist_deleted.borrow_mut() = true;
                Ok(())
            },
        )
        .unwrap_err();

        assert!(matches!(error, WincentError::SystemError(message) if message == "cleanup failed"));
        assert!(!*destlist_deleted.borrow());
    }
}
