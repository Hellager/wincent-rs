use std::collections::HashSet;

use super::parser::{DestList, DestListEntry};

/// Returns the visible entries from a DestList, applying version-appropriate filtering.
///
/// - v4: entries with `access_count > 0` are visible.
/// - v6: entries whose `recent_rank` is in `[0, normal_slot_count)` are visible,
///   plus all pinned entries sorted by `pin_order`.
/// - Other versions: all entries returned as-is.
pub fn visible_entries(dest_list: &DestList, normal_slot_count: i32) -> Vec<DestListEntry> {
    match dest_list.version {
        4 => visible_entries_v4(&dest_list.entries),
        6 => visible_entries_v6(&dest_list.entries, normal_slot_count),
        _ => dest_list.entries.clone(),
    }
}

fn visible_entries_v6(entries: &[DestListEntry], normal_slot_count: i32) -> Vec<DestListEntry> {
    let mut pinned: Vec<_> = entries
        .iter()
        .filter(|entry| entry.pin_order.is_some())
        .cloned()
        .collect();
    pinned.sort_by_key(|entry| entry.pin_order.unwrap_or(i32::MAX));

    let mut used_paths: HashSet<String> = pinned
        .iter()
        .map(|entry| entry.path.to_ascii_lowercase())
        .collect();

    let mut normal: Vec<_> = entries
        .iter()
        .filter(|entry| {
            entry.pin_order.is_none()
                && entry.recent_rank >= 0
                && entry.recent_rank < normal_slot_count
        })
        .cloned()
        .collect();
    normal.sort_by_key(|entry| std::cmp::Reverse(entry.recent_rank));
    normal.retain(|entry| used_paths.insert(entry.path.to_ascii_lowercase()));

    pinned.extend(normal);
    pinned
}

fn visible_entries_v4(entries: &[DestListEntry]) -> Vec<DestListEntry> {
    entries
        .iter()
        .filter(|entry| entry.access_count > 0)
        .cloned()
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::destlist::parser::{DestList, DestListEntry};

    fn make_entry(
        id: u64,
        path: &str,
        access_count: u32,
        pin_order: Option<i32>,
        recent_rank: i32,
    ) -> DestListEntry {
        DestListEntry {
            entry_offset: 0,
            entry_id: id,
            stream_name: format!("{id:x}"),
            raw_path: path.to_string(),
            path: path.to_string(),
            pin_order,
            recent_rank,
            access_count,
            last_access_filetime: None,
        }
    }

    fn make_dest_list(version: u32, entries: Vec<DestListEntry>) -> DestList {
        DestList {
            version,
            declared_entry_count: entries.len(),
            pinned_entry_count: 0,
            last_entry_id: 0,
            entries,
        }
    }

    #[test]
    fn v4_all_zero_access_count_returns_empty() {
        let dl = make_dest_list(
            4,
            vec![
                make_entry(1, "C:\\a.txt", 0, None, 0),
                make_entry(2, "C:\\b.txt", 0, None, 1),
            ],
        );
        let result = visible_entries(&dl, 10);
        assert!(result.is_empty());
    }

    #[test]
    fn v4_nonzero_access_count_visible() {
        let dl = make_dest_list(
            4,
            vec![
                make_entry(1, "C:\\a.txt", 0, None, 0),
                make_entry(2, "C:\\b.txt", 3, None, 1),
                make_entry(3, "C:\\c.txt", 1, None, 2),
            ],
        );
        let result = visible_entries(&dl, 10);
        assert_eq!(result.len(), 2);
        assert!(result.iter().any(|e| e.path == "C:\\b.txt"));
        assert!(result.iter().any(|e| e.path == "C:\\c.txt"));
    }

    #[test]
    fn v4_pinned_entries_included_regardless_of_access_count() {
        // In v4, pinned entries still have access_count > 0 in practice,
        // but the filter is purely access_count > 0.
        let dl = make_dest_list(
            4,
            vec![
                make_entry(1, "C:\\pinned.txt", 5, Some(0), 0),
                make_entry(2, "C:\\hidden.txt", 0, None, 1),
            ],
        );
        let result = visible_entries(&dl, 10);
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].path, "C:\\pinned.txt");
    }

    #[test]
    fn v6_pinned_entries_sorted_by_pin_order() {
        let dl = make_dest_list(
            6,
            vec![
                make_entry(1, "C:\\b.txt", 1, Some(2), -1),
                make_entry(2, "C:\\a.txt", 1, Some(0), -1),
                make_entry(3, "C:\\c.txt", 1, Some(1), -1),
            ],
        );
        let result = visible_entries(&dl, 10);
        assert_eq!(result.len(), 3);
        assert_eq!(result[0].path, "C:\\a.txt");
        assert_eq!(result[1].path, "C:\\c.txt");
        assert_eq!(result[2].path, "C:\\b.txt");
    }

    #[test]
    fn v6_recent_rank_filtering() {
        let dl = make_dest_list(
            6,
            vec![
                make_entry(1, "C:\\in_range.txt", 1, None, 2),
                make_entry(2, "C:\\out_of_range.txt", 1, None, 5),
                make_entry(3, "C:\\negative_rank.txt", 1, None, -1),
            ],
        );
        let result = visible_entries(&dl, 4); // normal_slot_count = 4
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].path, "C:\\in_range.txt");
    }
}
