//! Visibility-oriented status queries for Quick Access paths.
//!
//! These helpers answer one question: given a path, where would it appear in a
//! default Quick Access Shell query? The answer is derived from the persisted
//! `.automaticDestinations-ms` files with the same rules used by
//! [`AutomaticDestinations::rule_entries`](crate::destlist::AutomaticDestinations::rule_entries),
//! so it cannot drift from the rest of the DestList API.
//!
//! # Visibility model
//!
//! A path has a Recent Files position when it survives the observed Explorer
//! Recent Files rules: a valid v4 checksum table, no negative or `NaN` score,
//! `reserved_78 == 0`, a loadable numeric Shell Link stream, no earlier
//! duplicate path, and a rank inside
//! [`DEFAULT_RECENT_FILES_RESULT_LIMIT`](crate::destlist::DEFAULT_RECENT_FILES_RESULT_LIMIT).
//!
//! It has a Frequent Folders slot when it is pinned, or when it is an unpinned
//! entry that meets the observed access-count threshold
//! ([`FREQUENT_FOLDERS_MIN_ACCESS_COUNT`](crate::destlist::FREQUENT_FOLDERS_MIN_ACCESS_COUNT))
//! and fits inside
//! [`DEFAULT_FREQUENT_FOLDERS_NORMAL_SLOTS`](crate::destlist::DEFAULT_FREQUENT_FOLDERS_NORMAL_SLOTS)
//! after pinned entries and duplicate paths are removed.
//!
//! Everything else returns `None`, which means the rule-based query would omit
//! the path. It does not distinguish "absent from the backing file" from
//! "present but filtered". Use [`crate::destlist::entries`] on a parsed file for
//! that distinction.
//!
//! The window, slot count, and access-count threshold are the values observed
//! on Windows 10 1809 and 22H2; they are not documented Windows contracts.
//!
//! # Freshness
//!
//! These queries read the backing files on every call. They do not inspect the
//! currently rendered Explorer window. Explorer persists Shell changes
//! asynchronously and an open window may also need a refresh, so the rule-based
//! result and the visible UI can temporarily differ.

use std::io::ErrorKind;
use std::path::PathBuf;

use crate::destlist::parser::DestListRuleSet;
use crate::destlist::{
    frequent_folders_dest_path, parse_file, recent_files_dest_path, AutomaticDestinations,
    DestListEntry, DEFAULT_RECENT_FILES_RESULT_LIMIT,
};
use crate::error::WincentError;
use crate::utils::paths_equal;
use crate::{QuickAccess, WincentResult};

/// Position and pin state in a rule-based Frequent Folders query result.
///
/// # Examples
///
/// ```rust,no_run
/// use wincent::prelude::*;
///
/// # fn main() -> WincentResult<()> {
/// let manager = QuickAccessManager::new();
/// if let Some(slot) = manager.frequent_folder_slot(r"C:\Work")? {
///     println!("position={} pinned={}", slot.position(), slot.is_pinned());
/// }
/// # Ok(())
/// # }
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct FrequentFolderSlot {
    position: usize,
    pinned: bool,
}

impl FrequentFolderSlot {
    /// Zero-based position in the rule-based Frequent Folders result.
    ///
    /// Pinned entries come first in pin order, followed by unpinned entries in
    /// rank order.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Whether the entry is pinned to Quick Access.
    #[must_use]
    pub const fn is_pinned(&self) -> bool {
        self.pinned
    }

    pub(crate) const fn new(position: usize, pinned: bool) -> Self {
        Self { position, pinned }
    }
}

/// Where a path appears in the rule-based Quick Access query results.
///
/// A path omitted from both results is `Default` (both slots are `None`).
///
/// # Examples
///
/// ```rust,no_run
/// use wincent::prelude::*;
///
/// # fn main() -> WincentResult<()> {
/// let manager = QuickAccessManager::new();
/// let status = manager.quick_access_status(r"C:\Work\report.docx")?;
/// if let Some(position) = status.recent_file() {
///     println!("recent file at position {position}");
/// }
/// if let Some(slot) = status.frequent_folder() {
///     println!("frequent folder pinned={}", slot.is_pinned());
/// }
/// # Ok(())
/// # }
/// ```
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct QuickAccessStatus {
    recent_file: Option<usize>,
    frequent_folder: Option<FrequentFolderSlot>,
}

impl QuickAccessStatus {
    /// Zero-based position in the rule-based Recent Files result.
    #[must_use]
    pub const fn recent_file(&self) -> Option<usize> {
        self.recent_file
    }

    /// Position and pin state in the rule-based Frequent Folders result.
    #[must_use]
    pub const fn frequent_folder(&self) -> Option<FrequentFolderSlot> {
        self.frequent_folder
    }

    /// Whether the path appears in either rule-based result.
    #[must_use]
    pub const fn is_visible(&self) -> bool {
        self.recent_file.is_some() || self.frequent_folder.is_some()
    }

    /// Which rule-based results contain the path.
    ///
    /// Returns `None` when both rule-based results omit the path.
    #[must_use]
    pub const fn lists(&self) -> Option<QuickAccess> {
        match (self.recent_file.is_some(), self.frequent_folder.is_some()) {
            (true, true) => Some(QuickAccess::All),
            (true, false) => Some(QuickAccess::RecentFiles),
            (false, true) => Some(QuickAccess::FrequentFolders),
            (false, false) => None,
        }
    }

    pub(crate) const fn new(
        recent_file: Option<usize>,
        frequent_folder: Option<FrequentFolderSlot>,
    ) -> Self {
        Self {
            recent_file,
            frequent_folder,
        }
    }
}

/// Returns the entries from a rule-based Recent Files query.
pub(crate) fn recent_visible(dest: &AutomaticDestinations) -> Vec<DestListEntry> {
    let mut entries = dest.entries_for_rule_set(&DestListRuleSet::recent_files());
    entries.truncate(DEFAULT_RECENT_FILES_RESULT_LIMIT);
    entries
}

/// Returns the entries from a rule-based Frequent Folders query.
pub(crate) fn frequent_visible(dest: &AutomaticDestinations) -> Vec<DestListEntry> {
    dest.entries_for_rule_set(&DestListRuleSet::frequent_folders())
}

/// Returns the position of `path` inside a visible entry list.
pub(crate) fn locate(entries: &[DestListEntry], path: &str) -> Option<usize> {
    entries
        .iter()
        .position(|entry| paths_equal(entry.path(), path))
}

/// Evaluates both lists from already parsed backing files.
pub(crate) fn status_for_parsed(
    recent: Option<&AutomaticDestinations>,
    frequent: Option<&AutomaticDestinations>,
    path: &str,
) -> QuickAccessStatus {
    let recent_position = recent.and_then(|dest| locate(&recent_visible(dest), path));
    let frequent_slot = frequent.and_then(|dest| {
        let visible = frequent_visible(dest);
        let position = locate(&visible, path)?;
        Some(FrequentFolderSlot::new(
            position,
            visible[position].is_pinned(),
        ))
    });

    QuickAccessStatus::new(recent_position, frequent_slot)
}

/// Reads both backing files and evaluates the status of `path`.
pub(crate) fn status_for_path(path: &str) -> WincentResult<QuickAccessStatus> {
    let recent = load_list(recent_files_dest_path()?)?;
    let frequent = load_list(frequent_folders_dest_path()?)?;
    Ok(status_for_parsed(recent.as_ref(), frequent.as_ref(), path))
}

/// Reads the Recent Files backing file and returns the visible position of `path`.
pub(crate) fn recent_position_for_path(path: &str) -> WincentResult<Option<usize>> {
    let Some(dest) = load_list(recent_files_dest_path()?)? else {
        return Ok(None);
    };
    Ok(locate(&recent_visible(&dest), path))
}

/// Reads the Frequent Folders backing file and returns the visible slot of `path`.
pub(crate) fn frequent_slot_for_path(path: &str) -> WincentResult<Option<FrequentFolderSlot>> {
    let Some(dest) = load_list(frequent_folders_dest_path()?)? else {
        return Ok(None);
    };
    let visible = frequent_visible(&dest);
    let Some(position) = locate(&visible, path) else {
        return Ok(None);
    };
    Ok(Some(FrequentFolderSlot::new(
        position,
        visible[position].is_pinned(),
    )))
}

/// Parses one backing file, treating a missing file as an empty list.
fn load_list(path: PathBuf) -> WincentResult<Option<AutomaticDestinations>> {
    match parse_file(&path) {
        Ok(parsed) => Ok(Some(parsed)),
        Err(WincentError::Io(error)) if error.kind() == ErrorKind::NotFound => Ok(None),
        Err(error) => Err(error),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::destlist::test_support::{automatic_destinations_for_test, destlist_entry_for_test};
    use crate::destlist::DestListKind;

    fn parsed(kind: DestListKind, entries: Vec<DestListEntry>) -> AutomaticDestinations {
        let streams: Vec<String> = entries
            .iter()
            .map(|entry| entry.stream_name().to_string())
            .collect();
        automatic_destinations_for_test(kind, entries, &streams)
    }

    fn recent_entries(count: u32) -> Vec<DestListEntry> {
        (0..count)
            .map(|index| {
                destlist_entry_for_test(
                    &format!("C:\\Recent{index}.txt"),
                    index + 1,
                    0,
                    -1,
                    1,
                    None,
                )
            })
            .collect()
    }

    #[test]
    fn recent_visible_truncates_to_the_result_window() {
        let dest = parsed(DestListKind::RecentFiles, recent_entries(25));

        let visible = recent_visible(&dest);

        assert_eq!(visible.len(), DEFAULT_RECENT_FILES_RESULT_LIMIT);
        assert_eq!(locate(&visible, "C:\\Recent0.txt"), Some(0));
        assert_eq!(locate(&visible, "C:\\Recent19.txt"), Some(19));
        assert_eq!(locate(&visible, "C:\\Recent20.txt"), None);
    }

    #[test]
    fn locate_matches_windows_path_variants() {
        let dest = parsed(DestListKind::RecentFiles, recent_entries(3));
        let visible = recent_visible(&dest);

        assert_eq!(locate(&visible, "c:/recent1.txt"), Some(1));
        assert_eq!(locate(&visible, "C:\\RECENT1.TXT"), Some(1));
        assert_eq!(locate(&visible, "C:\\Recent9.txt"), None);
    }

    #[test]
    fn locate_matches_existing_equivalent_paths() -> WincentResult<()> {
        let directory = tempfile::tempdir()?;
        let nested = directory.path().join("nested");
        std::fs::create_dir(&nested)?;
        let target = directory.path().join("report.txt");
        std::fs::write(&target, b"status fixture")?;
        let equivalent = nested.join("..").join("report.txt");
        let target = target.to_string_lossy();
        let equivalent = equivalent.to_string_lossy();
        let entries = vec![destlist_entry_for_test(&target, 1, 0, -1, 1, None)];

        assert_eq!(locate(&entries, &equivalent), Some(0));
        Ok(())
    }

    #[test]
    fn frequent_visible_orders_pinned_first_then_ranked_unpinned() {
        let entries = vec![
            destlist_entry_for_test("C:\\Unpinned", 1, 0, 0, 3, None),
            destlist_entry_for_test("C:\\Pinned", 2, 0, 1, 0, Some(0)),
        ];
        let dest = parsed(DestListKind::FrequentFolders, entries);

        let visible = frequent_visible(&dest);

        assert_eq!(visible.len(), 2);
        assert_eq!(visible[0].path(), "C:\\Pinned");
        assert!(visible[0].is_pinned());
        assert_eq!(visible[1].path(), "C:\\Unpinned");
        assert!(!visible[1].is_pinned());
    }

    #[test]
    fn frequent_visible_drops_unpinned_entries_below_the_access_threshold() {
        let entries = vec![
            destlist_entry_for_test("C:\\Enough", 1, 0, 0, 3, None),
            destlist_entry_for_test("C:\\TooFew", 2, 0, 1, 2, None),
        ];
        let dest = parsed(DestListKind::FrequentFolders, entries);

        let visible = frequent_visible(&dest);

        assert_eq!(locate(&visible, "C:\\Enough"), Some(0));
        assert_eq!(locate(&visible, "C:\\TooFew"), None);
    }

    #[test]
    fn frequent_visible_keeps_pinned_entries_below_the_access_threshold() {
        let entries = vec![destlist_entry_for_test("C:\\Pinned", 1, 0, 0, 0, Some(0))];
        let dest = parsed(DestListKind::FrequentFolders, entries);

        assert_eq!(locate(&frequent_visible(&dest), "C:\\Pinned"), Some(0));
    }

    #[test]
    fn frequent_visible_drops_unpinned_entries_outside_the_slot_window() {
        let entries: Vec<_> = (0..6)
            .map(|index| {
                destlist_entry_for_test(
                    &format!("C:\\Slot{index}"),
                    index + 1,
                    0,
                    index as i32,
                    3,
                    None,
                )
            })
            .collect();
        let dest = parsed(DestListKind::FrequentFolders, entries);

        let visible = frequent_visible(&dest);

        assert_eq!(
            visible.len(),
            usize::try_from(crate::destlist::DEFAULT_FREQUENT_FOLDERS_NORMAL_SLOTS).unwrap()
        );
        assert_eq!(locate(&visible, "C:\\Slot3"), Some(3));
        assert_eq!(locate(&visible, "C:\\Slot4"), None);
    }

    #[test]
    fn frequent_visible_prefers_the_pinned_entry_for_duplicate_paths() {
        let entries = vec![
            destlist_entry_for_test("C:\\Dup", 1, 0, 0, 3, None),
            destlist_entry_for_test("c:/dup", 2, 0, 1, 0, Some(0)),
        ];
        let dest = parsed(DestListKind::FrequentFolders, entries);

        let visible = frequent_visible(&dest);

        assert_eq!(visible.len(), 1);
        assert!(visible[0].is_pinned());
        assert_eq!(locate(&visible, "C:\\Dup"), Some(0));
    }

    #[test]
    fn status_for_parsed_reports_both_lists() {
        let recent = parsed(DestListKind::RecentFiles, recent_entries(3));
        let frequent = parsed(
            DestListKind::FrequentFolders,
            vec![destlist_entry_for_test("C:\\Folder", 9, 0, 0, 0, Some(0))],
        );

        let mixed = status_for_parsed(Some(&recent), Some(&frequent), "c:/recent2.txt");
        assert_eq!(mixed.recent_file(), Some(2));
        assert_eq!(mixed.frequent_folder(), None);

        let folder = status_for_parsed(Some(&recent), Some(&frequent), "C:\\Folder");
        assert_eq!(folder.recent_file(), None);
        let slot = folder.frequent_folder().expect("folder should be visible");
        assert_eq!(slot.position(), 0);
        assert!(slot.is_pinned());

        let absent = status_for_parsed(Some(&recent), Some(&frequent), "C:\\Nope.txt");
        assert!(!absent.is_visible());
    }

    #[test]
    fn status_for_parsed_handles_missing_backing_files() {
        let status = status_for_parsed(None, None, "C:\\Anything");

        assert!(!status.is_visible());
        assert_eq!(status, QuickAccessStatus::default());
    }

    #[test]
    fn lists_reports_the_visible_categories() {
        let recent = QuickAccessStatus::new(Some(1), None);
        let frequent = QuickAccessStatus::new(None, Some(FrequentFolderSlot::new(0, true)));
        let both = QuickAccessStatus::new(Some(0), Some(FrequentFolderSlot::new(0, false)));

        assert_eq!(recent.lists(), Some(QuickAccess::RecentFiles));
        assert_eq!(frequent.lists(), Some(QuickAccess::FrequentFolders));
        assert_eq!(both.lists(), Some(QuickAccess::All));
        assert_eq!(QuickAccessStatus::default().lists(), None);
        assert!(!QuickAccessStatus::default().is_visible());
    }
}
