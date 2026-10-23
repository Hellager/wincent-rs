//! Direct DestList v4 entry visibility, pin-state, and ordering updates.

use std::cmp::Ordering;
use std::io::{Cursor, Write};
use std::path::Path;

use crate::destlist::parser::{
    frequent_folders_dest_path, parse_bytes_with_kind, recent_files_dest_path, DestListEntry,
    DestListKind, FREQUENT_FOLDERS_MIN_ACCESS_COUNT,
};
use crate::destlist::{DestListItemUpdate, DestListPosition};
use crate::error::WincentError;
use crate::utils::normalize_path_lightweight;
use crate::{QuickAccess, WincentResult};

const DEST_LIST_HEADER_LEN: usize = 0x20;
const PINNED_COUNT_OFFSET: usize = 0x08;
const PIN_STATUS_OFFSET: usize = 0x6c;
const RANK_OFFSET: usize = 0x70;
const ACCESS_COUNT_OFFSET: usize = 0x74;
const RESERVED_78_OFFSET: usize = 0x78;

#[derive(Clone)]
struct MutableEntry {
    entry_id: u64,
    path_key: String,
    pin_status: i32,
    rank: i32,
    access_count: u32,
    bytes: Vec<u8>,
}

pub(crate) fn update_entry_path(
    qa_type: QuickAccess,
    target_path: &str,
    update: DestListItemUpdate,
) -> WincentResult<()> {
    if update.is_empty() {
        return Err(WincentError::InvalidArgument(
            "direct DestList update must request at least one change".to_string(),
        ));
    }
    if qa_type == QuickAccess::RecentFiles && update.pinned_value().is_some() {
        return Err(WincentError::UnsupportedOperation(
            "direct pin changes are supported only for Frequent Folders".to_string(),
        ));
    }

    let (dest_path, kind) = match qa_type {
        QuickAccess::RecentFiles => (recent_files_dest_path()?, DestListKind::RecentFiles),
        QuickAccess::FrequentFolders => {
            (frequent_folders_dest_path()?, DestListKind::FrequentFolders)
        }
        QuickAccess::All => {
            return Err(WincentError::UnsupportedOperation(
                "QuickAccess::All is not valid for direct DestList updates".to_string(),
            ));
        }
    };

    update_entry_path_in_file(&dest_path, kind, qa_type, target_path, update)
}

fn update_entry_path_in_file(
    dest_path: &Path,
    kind: DestListKind,
    qa_type: QuickAccess,
    target_path: &str,
    update: DestListItemUpdate,
) -> WincentResult<()> {
    let write_guard = super::transaction::DirectWriteGuard::acquire(dest_path)?;
    let snapshot = write_guard.read_snapshot(dest_path)?;
    let original = snapshot.bytes().to_vec();
    let rewritten = rewrite_cfb_bytes(original, kind, qa_type, target_path, update)?;
    super::remove::replace_file_atomically(dest_path, &rewritten, &write_guard, &snapshot)
}

fn rewrite_cfb_bytes(
    original: Vec<u8>,
    kind: DestListKind,
    qa_type: QuickAccess,
    target_path: &str,
    update: DestListItemUpdate,
) -> WincentResult<Vec<u8>> {
    let parsed = parse_bytes_with_kind(original.clone(), kind)?;
    validate_original(
        parsed.dest_list().entries(),
        parsed.dest_list().version(),
        parsed.dest_list().declared_entry_count(),
    )?;

    let old_dest = parsed
        .cfb_info()
        .stream("DestList")
        .ok_or_else(|| WincentError::DestListParse("missing DestList stream".to_string()))?
        .data();
    let mut entries = mutable_entries(parsed.dest_list().entries(), old_dest)?;
    let target_key = normalize_path_lightweight(target_path);
    let matches: Vec<_> = entries
        .iter()
        .filter(|entry| entry.path_key == target_key)
        .map(|entry| entry.entry_id)
        .collect();
    let target_id = match matches.as_slice() {
        [] => return Err(WincentError::not_in_quick_access(target_path, qa_type)),
        [entry_id] => *entry_id,
        _ => {
            return Err(WincentError::DestListParse(format!(
                "multiple raw DestList entries match {target_path}"
            )));
        }
    };
    let target_index = entry_index(&entries, target_id)?;
    let currently_pinned = entries[target_index].pin_status >= 0;
    let currently_visible = read_u32(&entries[target_index].bytes, RESERVED_78_OFFSET)? == 0;
    let finally_pinned = update.pinned_value().unwrap_or(currently_pinned);
    let finally_visible = update.visible_value().unwrap_or(currently_visible);
    if kind == DestListKind::FrequentFolders && finally_pinned && !finally_visible {
        return Err(WincentError::InvalidArgument(
            "a pinned Frequent Folder cannot be hidden; explicitly unpin it when hiding"
                .to_string(),
        ));
    }

    if let Some(pinned) = update.pinned_value() {
        set_frequent_pinned(&mut entries, target_id, pinned)?;
    }
    if let Some(visible) = update.visible_value() {
        set_visible(&mut entries, target_id, kind, visible)?;
    }
    if let Some(position) = update.position_value() {
        set_position(&mut entries, target_id, kind, position)?;
    }

    let dest_bytes = build_destlist(old_dest, &mut entries)?;
    let rewritten = replace_destlist_stream(original, &dest_bytes)?;
    validate_rewritten(
        &rewritten,
        kind,
        parsed.dest_list().entries(),
        target_id,
        update,
    )?;
    Ok(rewritten)
}

fn validate_original(
    entries: &[DestListEntry],
    version: u32,
    declared_entry_count: usize,
) -> WincentResult<()> {
    if version != 4 {
        return Err(WincentError::DestListUnsupportedVersion(version));
    }
    if declared_entry_count != entries.len() {
        return Err(WincentError::DestListParse(format!(
            "cannot rewrite a partial DestList: declared {declared_entry_count}, parsed {}",
            entries.len()
        )));
    }
    if entries.iter().any(|entry| {
        !entry.checksum_valid()
            || matches!(entry.score().partial_cmp(&0.0), Some(Ordering::Less) | None)
    }) {
        return Err(WincentError::DestListParse(
            "cannot rewrite an invalid DestList v4 table".to_string(),
        ));
    }
    Ok(())
}

fn mutable_entries(
    entries: &[DestListEntry],
    dest_bytes: &[u8],
) -> WincentResult<Vec<MutableEntry>> {
    entries
        .iter()
        .map(|entry| {
            let end = entry
                .entry_offset()
                .checked_add(entry.entry_len())
                .ok_or_else(|| {
                    WincentError::DestListParse("DestList entry range overflow".to_string())
                })?;
            let bytes = dest_bytes
                .get(entry.entry_offset()..end)
                .ok_or_else(|| {
                    WincentError::DestListParse(format!(
                        "invalid DestList entry range 0x{:x}..0x{end:x}",
                        entry.entry_offset()
                    ))
                })?
                .to_vec();
            Ok(MutableEntry {
                entry_id: entry.entry_id(),
                path_key: normalize_path_lightweight(entry.path()),
                pin_status: entry.pin_status(),
                rank: entry.recent_rank(),
                access_count: entry.access_count(),
                bytes,
            })
        })
        .collect()
}

fn set_frequent_pinned(
    entries: &mut [MutableEntry],
    target_id: u64,
    pinned: bool,
) -> WincentResult<()> {
    let target_index = entry_index(entries, target_id)?;
    let was_pinned = entries[target_index].pin_status >= 0;
    if was_pinned == pinned {
        return Ok(());
    }

    if pinned {
        let mut pinned_ids = pinned_ids(entries, Some(target_id));
        pinned_ids.push(target_id);
        assign_pin_order(entries, &pinned_ids)?;
        return Ok(());
    }

    entries[target_index].pin_status = -1;
    entries[target_index].access_count = entries[target_index]
        .access_count
        .max(FREQUENT_FOLDERS_MIN_ACCESS_COUNT);
    entries[target_index].rank = next_unpinned_rank(entries, target_id)?;
    let pinned_ids = pinned_ids(entries, Some(target_id));
    assign_pin_order(entries, &pinned_ids)
}

fn set_visible(
    entries: &mut [MutableEntry],
    target_id: u64,
    kind: DestListKind,
    visible: bool,
) -> WincentResult<()> {
    let target_index = entry_index(entries, target_id)?;
    write_u32(
        &mut entries[target_index].bytes,
        RESERVED_78_OFFSET,
        u32::from(!visible),
    )?;
    if visible && kind == DestListKind::FrequentFolders && entries[target_index].pin_status < 0 {
        entries[target_index].access_count = entries[target_index]
            .access_count
            .max(FREQUENT_FOLDERS_MIN_ACCESS_COUNT);
        if entries[target_index].rank < 0 {
            entries[target_index].rank = next_unpinned_rank(entries, target_id)?;
        }
    }
    Ok(())
}

fn set_position(
    entries: &mut Vec<MutableEntry>,
    target_id: u64,
    kind: DestListKind,
    position: DestListPosition,
) -> WincentResult<()> {
    match kind {
        DestListKind::RecentFiles => move_recent_entry(entries, target_id, position),
        DestListKind::FrequentFolders => {
            let target_index = entry_index(entries, target_id)?;
            if entries[target_index].pin_status >= 0 {
                move_pinned_entry(entries, target_id, position)
            } else {
                move_unpinned_entry(entries, target_id, position)
            }
        }
        DestListKind::Unknown => Err(WincentError::UnsupportedOperation(
            "unknown DestList kind cannot be reordered".to_string(),
        )),
    }
}

fn move_recent_entry(
    entries: &mut Vec<MutableEntry>,
    target_id: u64,
    position: DestListPosition,
) -> WincentResult<()> {
    let target_index = entry_index(entries, target_id)?;
    let destination = position_index(position, entries.len())?;
    let target = entries.remove(target_index);
    entries.insert(destination, target);
    Ok(())
}

fn move_pinned_entry(
    entries: &mut [MutableEntry],
    target_id: u64,
    position: DestListPosition,
) -> WincentResult<()> {
    let mut ids = pinned_ids(entries, None);
    move_id(&mut ids, target_id, position)?;
    assign_pin_order(entries, &ids)
}

fn move_unpinned_entry(
    entries: &mut [MutableEntry],
    target_id: u64,
    position: DestListPosition,
) -> WincentResult<()> {
    let mut slots: Vec<_> = entries
        .iter()
        .enumerate()
        .filter(|(_, entry)| {
            entry.pin_status < 0 && (entry.rank >= 0 || entry.entry_id == target_id)
        })
        .map(|(index, _)| index)
        .collect();
    slots.sort_unstable();

    let mut group: Vec<_> = slots.iter().map(|index| entries[*index].clone()).collect();
    group.sort_by(|left, right| {
        rank_sort_key(left)
            .cmp(&rank_sort_key(right))
            .then_with(|| left.entry_id.cmp(&right.entry_id))
    });
    let mut ids: Vec<_> = group.iter().map(|entry| entry.entry_id).collect();
    move_id(&mut ids, target_id, position)?;

    group.sort_by_key(|entry| {
        ids.iter()
            .position(|entry_id| *entry_id == entry.entry_id)
            .unwrap_or(usize::MAX)
    });
    for (rank, entry) in group.iter_mut().enumerate() {
        entry.rank = i32::try_from(rank).map_err(|_| {
            WincentError::DestListParse("Frequent Folder rank overflow".to_string())
        })?;
    }
    for (slot, entry) in slots.into_iter().zip(group) {
        entries[slot] = entry;
    }
    Ok(())
}

fn pinned_ids(entries: &[MutableEntry], excluded: Option<u64>) -> Vec<u64> {
    let mut indexed: Vec<_> = entries
        .iter()
        .enumerate()
        .filter(|(_, entry)| entry.pin_status >= 0 && Some(entry.entry_id) != excluded)
        .collect();
    indexed.sort_by(|(left_index, left), (right_index, right)| {
        left.pin_status
            .cmp(&right.pin_status)
            .then_with(|| left_index.cmp(right_index))
    });
    indexed
        .into_iter()
        .map(|(_, entry)| entry.entry_id)
        .collect()
}

fn assign_pin_order(entries: &mut [MutableEntry], ids: &[u64]) -> WincentResult<()> {
    for (order, entry_id) in ids.iter().copied().enumerate() {
        let index = entry_index(entries, entry_id)?;
        entries[index].pin_status = i32::try_from(order).map_err(|_| {
            WincentError::DestListParse("Frequent Folder pin order overflow".to_string())
        })?;
    }
    Ok(())
}

fn next_unpinned_rank(entries: &[MutableEntry], target_id: u64) -> WincentResult<i32> {
    entries
        .iter()
        .filter(|entry| entry.entry_id != target_id && entry.pin_status < 0 && entry.rank >= 0)
        .map(|entry| entry.rank)
        .max()
        .unwrap_or(-1)
        .checked_add(1)
        .ok_or_else(|| WincentError::DestListParse("Frequent Folder rank overflow".to_string()))
}

fn rank_sort_key(entry: &MutableEntry) -> i32 {
    if entry.rank < 0 {
        i32::MAX
    } else {
        entry.rank
    }
}

fn move_id(ids: &mut Vec<u64>, target_id: u64, position: DestListPosition) -> WincentResult<()> {
    let current = ids
        .iter()
        .position(|entry_id| *entry_id == target_id)
        .ok_or_else(|| {
            WincentError::DestListParse("target entry is not in its order group".to_string())
        })?;
    let destination = position_index(position, ids.len())?;
    let entry_id = ids.remove(current);
    ids.insert(destination, entry_id);
    Ok(())
}

fn position_index(position: DestListPosition, len: usize) -> WincentResult<usize> {
    if len == 0 {
        return Err(WincentError::DestListParse(
            "cannot position an entry in an empty list".to_string(),
        ));
    }
    match position {
        DestListPosition::First => Ok(0),
        DestListPosition::Last => Ok(len - 1),
        DestListPosition::Index(index) if index < len => Ok(index),
        DestListPosition::Index(index) => Err(WincentError::InvalidArgument(format!(
            "DestList position {index} is outside the {len}-entry group"
        ))),
    }
}

fn entry_index(entries: &[MutableEntry], target_id: u64) -> WincentResult<usize> {
    entries
        .iter()
        .position(|entry| entry.entry_id == target_id)
        .ok_or_else(|| {
            WincentError::DestListParse("target entry disappeared during update".to_string())
        })
}

fn build_destlist(old_dest: &[u8], entries: &mut [MutableEntry]) -> WincentResult<Vec<u8>> {
    if old_dest.len() < DEST_LIST_HEADER_LEN {
        return Err(WincentError::DestListParse(
            "DestList stream is shorter than its v4 header".to_string(),
        ));
    }
    let mut result = old_dest[..DEST_LIST_HEADER_LEN].to_vec();
    for entry in entries.iter_mut() {
        write_i32(&mut entry.bytes, PIN_STATUS_OFFSET, entry.pin_status)?;
        write_i32(&mut entry.bytes, RANK_OFFSET, entry.rank)?;
        write_u32(&mut entry.bytes, ACCESS_COUNT_OFFSET, entry.access_count)?;
        let checksum = crc64(&entry.bytes)?;
        write_u64(&mut entry.bytes, 0, checksum)?;
        result.extend_from_slice(&entry.bytes);
    }
    let pinned_count = u32::try_from(entries.iter().filter(|entry| entry.pin_status >= 0).count())
        .map_err(|_| WincentError::DestListParse("pinned entry count overflow".to_string()))?;
    write_u32(&mut result, PINNED_COUNT_OFFSET, pinned_count)?;
    Ok(result)
}

fn replace_destlist_stream(original: Vec<u8>, dest_bytes: &[u8]) -> WincentResult<Vec<u8>> {
    let mut compound = ::cfb::CompoundFile::open(Cursor::new(original)).map_err(|error| {
        WincentError::DestListParse(format!("failed to open writable CFB image: {error}"))
    })?;
    {
        let mut stream = compound.create_stream("/DestList").map_err(|error| {
            WincentError::DestListParse(format!("failed to replace DestList stream: {error}"))
        })?;
        stream.write_all(dest_bytes).map_err(WincentError::Io)?;
    }
    compound.flush().map_err(WincentError::Io)?;
    Ok(compound.into_inner().into_inner())
}

fn validate_rewritten(
    rewritten: &[u8],
    kind: DestListKind,
    original_entries: &[DestListEntry],
    target_id: u64,
    update: DestListItemUpdate,
) -> WincentResult<()> {
    let parsed = parse_bytes_with_kind(rewritten.to_vec(), kind)?;
    let entries = parsed.dest_list().entries();
    validate_original(
        entries,
        parsed.dest_list().version(),
        parsed.dest_list().declared_entry_count(),
    )?;
    if entries.len() != original_entries.len() {
        return Err(WincentError::DestListParse(
            "direct update changed the DestList entry count".to_string(),
        ));
    }
    for original in original_entries {
        let current = entries
            .iter()
            .find(|entry| entry.entry_id() == original.entry_id())
            .ok_or_else(|| {
                WincentError::DestListParse(format!(
                    "direct update lost entry {}",
                    original.entry_id()
                ))
            })?;
        if current.score().to_bits() != original.score().to_bits()
            || current.last_interaction_filetime() != original.last_interaction_filetime()
        {
            return Err(WincentError::DestListParse(
                "direct update changed score or last-interaction time".to_string(),
            ));
        }
    }
    let target = entries
        .iter()
        .find(|entry| entry.entry_id() == target_id)
        .ok_or_else(|| {
            WincentError::DestListParse("updated target entry is missing".to_string())
        })?;
    if let Some(visible) = update.visible_value() {
        let is_visible = target.reserved_78().unwrap_or_default() == 0;
        if is_visible != visible {
            return Err(WincentError::DestListParse(
                "updated target visibility did not persist".to_string(),
            ));
        }
    }
    if let Some(pinned) = update.pinned_value() {
        if target.is_pinned() != pinned {
            return Err(WincentError::DestListParse(
                "updated target pin state did not persist".to_string(),
            ));
        }
    }
    Ok(())
}

fn crc64(entry: &[u8]) -> WincentResult<u64> {
    const POLYNOMIAL: u64 = 0x92c6_4265_d321_39a4;
    let covered = entry.get(..0x70).ok_or_else(|| {
        WincentError::DestListParse("DestList entry is shorter than its CRC span".to_string())
    })?;
    let mut crc = u64::MAX;
    for (index, byte) in covered.iter().copied().enumerate() {
        crc ^= u64::from(if index < 8 { 0 } else { byte });
        for _ in 0..8 {
            crc = if crc & 1 == 0 {
                crc >> 1
            } else {
                (crc >> 1) ^ POLYNOMIAL
            };
        }
    }
    Ok(crc)
}

fn write_i32(data: &mut [u8], offset: usize, value: i32) -> WincentResult<()> {
    write_bytes(data, offset, &value.to_le_bytes(), "i32")
}

fn write_u32(data: &mut [u8], offset: usize, value: u32) -> WincentResult<()> {
    write_bytes(data, offset, &value.to_le_bytes(), "u32")
}

fn write_u64(data: &mut [u8], offset: usize, value: u64) -> WincentResult<()> {
    write_bytes(data, offset, &value.to_le_bytes(), "u64")
}

fn read_u32(data: &[u8], offset: usize) -> WincentResult<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| WincentError::DestListParse("u32 offset overflow".to_string()))?;
    let bytes: [u8; 4] = data
        .get(offset..end)
        .ok_or_else(|| {
            WincentError::DestListParse(format!("missing u32 at DestList offset 0x{offset:x}"))
        })?
        .try_into()
        .map_err(|_| WincentError::DestListParse("invalid u32 width".to_string()))?;
    Ok(u32::from_le_bytes(bytes))
}

fn write_bytes(data: &mut [u8], offset: usize, value: &[u8], name: &str) -> WincentResult<()> {
    let end = offset
        .checked_add(value.len())
        .ok_or_else(|| WincentError::DestListParse(format!("{name} offset overflow")))?;
    let bytes = data.get_mut(offset..end).ok_or_else(|| {
        WincentError::DestListParse(format!("missing {name} at DestList offset 0x{offset:x}"))
    })?;
    bytes.copy_from_slice(value);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[derive(Clone, Copy)]
    struct TestEntry<'a> {
        id: u64,
        path: &'a str,
        score: f32,
        time: u64,
        pin_status: i32,
        rank: i32,
        access_count: u32,
        reserved_78: u32,
    }

    #[test]
    fn recent_update_hides_and_moves_entry_without_changing_score_or_time() -> WincentResult<()> {
        let source = build_test_cfb(&[
            test_entry(1, "C:\\Work\\a.txt", 1.0, 100, -1, -1, 1),
            test_entry(2, "C:\\Work\\b.txt", 2.0, 200, -1, -1, 2),
        ]);
        let update = DestListItemUpdate::new()
            .hide()
            .with_position(DestListPosition::First);

        let rewritten = rewrite_cfb_bytes(
            source,
            DestListKind::RecentFiles,
            QuickAccess::RecentFiles,
            "c:/work/b.txt",
            update,
        )?;
        let parsed = parse_bytes_with_kind(rewritten, DestListKind::RecentFiles)?;
        let entries = parsed.dest_list().entries();

        assert_eq!(entries[0].entry_id(), 2);
        assert_eq!(entries[0].reserved_78(), Some(1));
        assert_eq!(entries[0].score(), 2.0);
        assert_eq!(entries[0].last_interaction_filetime(), Some(200));
        assert!(parsed.cfb_info().stream("2").is_some());
        Ok(())
    }

    #[test]
    fn frequent_update_pins_entry_and_preserves_usage_metadata() -> WincentResult<()> {
        let source = build_test_cfb(&[
            test_entry(1, "C:\\Pinned", 5.0, 100, 0, 3, 4),
            test_entry(2, "C:\\Frequent", 3.0, 200, -1, 0, 3),
        ]);

        let rewritten = rewrite_cfb_bytes(
            source,
            DestListKind::FrequentFolders,
            QuickAccess::FrequentFolders,
            "C:\\Frequent",
            DestListItemUpdate::new().pin(),
        )?;
        let parsed = parse_bytes_with_kind(rewritten, DestListKind::FrequentFolders)?;
        let target = entry_by_id(parsed.dest_list().entries(), 2);

        assert_eq!(parsed.dest_list().pinned_entry_count(), 2);
        assert_eq!(target.pin_order(), Some(1));
        assert_eq!(target.score(), 3.0);
        assert_eq!(target.last_interaction_filetime(), Some(200));
        assert_eq!(target.access_count(), 3);
        Ok(())
    }

    #[test]
    fn frequent_update_unpins_without_deleting_entry_or_stream() -> WincentResult<()> {
        let source = build_test_cfb(&[
            test_entry(1, "C:\\Pinned", 0.0, 0, 0, -1, 0),
            test_entry(2, "C:\\Frequent", 3.0, 200, -1, 0, 3),
        ]);

        let rewritten = rewrite_cfb_bytes(
            source,
            DestListKind::FrequentFolders,
            QuickAccess::FrequentFolders,
            "C:\\Pinned",
            DestListItemUpdate::new().unpin(),
        )?;
        let parsed = parse_bytes_with_kind(rewritten, DestListKind::FrequentFolders)?;
        let target = entry_by_id(parsed.dest_list().entries(), 1);

        assert_eq!(parsed.dest_list().declared_entry_count(), 2);
        assert_eq!(parsed.dest_list().pinned_entry_count(), 0);
        assert!(!target.is_pinned());
        assert_eq!(target.recent_rank(), 1);
        assert_eq!(target.access_count(), FREQUENT_FOLDERS_MIN_ACCESS_COUNT);
        assert_eq!(target.score(), 0.0);
        assert_eq!(target.last_interaction_filetime(), Some(0));
        assert!(parsed.cfb_info().stream("1").is_some());
        Ok(())
    }

    #[test]
    fn frequent_update_rejects_hidden_pinned_final_state() {
        let source = build_test_cfb(&[
            test_entry(1, "C:\\Pinned", 5.0, 100, 0, 3, 4),
            test_entry(2, "C:\\Frequent", 3.0, 200, -1, 0, 3),
        ]);

        let error = rewrite_cfb_bytes(
            source,
            DestListKind::FrequentFolders,
            QuickAccess::FrequentFolders,
            "C:\\Pinned",
            DestListItemUpdate::new().hide(),
        )
        .unwrap_err();

        assert!(matches!(
            error,
            WincentError::InvalidArgument(message)
                if message.contains("explicitly unpin it when hiding")
        ));
    }

    #[test]
    fn frequent_update_can_unpin_and_hide_without_deleting() -> WincentResult<()> {
        let source = build_test_cfb(&[
            test_entry(1, "C:\\Pinned", 5.0, 100, 0, 3, 4),
            test_entry(2, "C:\\Frequent", 3.0, 200, -1, 0, 3),
        ]);

        let rewritten = rewrite_cfb_bytes(
            source,
            DestListKind::FrequentFolders,
            QuickAccess::FrequentFolders,
            "C:\\Pinned",
            DestListItemUpdate::new().unpin().hide(),
        )?;
        let parsed = parse_bytes_with_kind(rewritten, DestListKind::FrequentFolders)?;
        let target = entry_by_id(parsed.dest_list().entries(), 1);

        assert!(!target.is_pinned());
        assert_eq!(target.reserved_78(), Some(1));
        assert!(parsed.cfb_info().stream("1").is_some());
        Ok(())
    }

    #[test]
    fn frequent_unpinned_move_updates_rank_and_candidate_sequence() -> WincentResult<()> {
        let source = build_test_cfb(&[
            test_entry(1, "C:\\Pinned", 5.0, 100, 0, 4, 4),
            test_entry(2, "C:\\First", 3.0, 200, -1, 0, 3),
            test_entry(3, "C:\\Second", 3.0, 300, -1, 1, 3),
        ]);

        let rewritten = rewrite_cfb_bytes(
            source,
            DestListKind::FrequentFolders,
            QuickAccess::FrequentFolders,
            "C:\\Second",
            DestListItemUpdate::new().with_position(DestListPosition::First),
        )?;
        let parsed = parse_bytes_with_kind(rewritten, DestListKind::FrequentFolders)?;
        let entries = parsed.dest_list().entries();

        assert_eq!(entries[1].entry_id(), 3);
        assert_eq!(entries[1].recent_rank(), 0);
        assert_eq!(entries[2].entry_id(), 2);
        assert_eq!(entries[2].recent_rank(), 1);
        assert_eq!(entries[1].score(), 3.0);
        assert_eq!(entries[1].last_interaction_filetime(), Some(300));
        Ok(())
    }

    #[test]
    fn duplicate_paths_are_rejected_as_ambiguous() {
        let source = build_test_cfb(&[
            test_entry(1, "C:\\Work", 3.0, 100, -1, 0, 3),
            test_entry(2, "c:/work/", 3.0, 200, -1, 1, 3),
        ]);

        let error = rewrite_cfb_bytes(
            source,
            DestListKind::FrequentFolders,
            QuickAccess::FrequentFolders,
            "C:\\Work",
            DestListItemUpdate::new().hide(),
        )
        .unwrap_err();

        assert!(
            matches!(error, WincentError::DestListParse(message) if message.contains("multiple"))
        );
    }

    fn test_entry<'a>(
        id: u64,
        path: &'a str,
        score: f32,
        time: u64,
        pin_status: i32,
        rank: i32,
        access_count: u32,
    ) -> TestEntry<'a> {
        TestEntry {
            id,
            path,
            score,
            time,
            pin_status,
            rank,
            access_count,
            reserved_78: 0,
        }
    }

    fn entry_by_id(entries: &[DestListEntry], entry_id: u64) -> &DestListEntry {
        entries
            .iter()
            .find(|entry| entry.entry_id() == entry_id)
            .expect("test entry should exist")
    }

    fn build_test_cfb(entries: &[TestEntry<'_>]) -> Vec<u8> {
        let mut dest_list = vec![0_u8; DEST_LIST_HEADER_LEN];
        write_test_u32(&mut dest_list, 0, 4);
        write_test_u32(&mut dest_list, 0x04, entries.len() as u32);
        write_test_u32(
            &mut dest_list,
            PINNED_COUNT_OFFSET,
            entries.iter().filter(|entry| entry.pin_status >= 0).count() as u32,
        );
        write_test_u32(
            &mut dest_list,
            0x0c,
            entries
                .iter()
                .map(|entry| entry.score)
                .sum::<f32>()
                .to_bits(),
        );
        write_test_u64(
            &mut dest_list,
            0x10,
            entries.last().map_or(0, |entry| entry.id),
        );
        write_test_u64(&mut dest_list, 0x18, 7);
        for entry in entries {
            dest_list.extend_from_slice(&build_test_entry(*entry));
        }

        let cursor = Cursor::new(Vec::new());
        let mut compound = ::cfb::CompoundFile::create_with_version(::cfb::Version::V3, cursor)
            .expect("create test CFB");
        compound
            .create_stream("/DestList")
            .expect("create DestList")
            .write_all(&dest_list)
            .expect("write DestList");
        for entry in entries {
            compound
                .create_stream(format!("/{:x}", entry.id))
                .expect("create numeric stream")
                .write_all(b"test shell link")
                .expect("write numeric stream");
        }
        compound.flush().expect("flush test CFB");
        compound.into_inner().into_inner()
    }

    fn build_test_entry(entry: TestEntry<'_>) -> Vec<u8> {
        let path_bytes: Vec<u8> = entry
            .path
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let mut bytes = vec![0_u8; 0x82 + path_bytes.len() + 4];
        write_test_u64(&mut bytes, 0x58, entry.id);
        write_test_u32(&mut bytes, 0x60, entry.score.to_bits());
        write_test_u64(&mut bytes, 0x64, entry.time);
        write_test_u32(&mut bytes, PIN_STATUS_OFFSET, entry.pin_status as u32);
        write_test_u32(&mut bytes, RANK_OFFSET, entry.rank as u32);
        write_test_u32(&mut bytes, ACCESS_COUNT_OFFSET, entry.access_count);
        write_test_u32(&mut bytes, RESERVED_78_OFFSET, entry.reserved_78);
        let path_chars =
            u16::try_from(entry.path.encode_utf16().count()).expect("test path should fit in u16");
        bytes[0x80..0x82].copy_from_slice(&path_chars.to_le_bytes());
        bytes[0x82..0x82 + path_bytes.len()].copy_from_slice(&path_bytes);
        let checksum = crc64(&bytes).expect("test entry should have CRC span");
        write_test_u64(&mut bytes, 0, checksum);
        bytes
    }

    fn write_test_u32(data: &mut [u8], offset: usize, value: u32) {
        data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    }

    fn write_test_u64(data: &mut [u8], offset: usize, value: u64) {
        data[offset..offset + 8].copy_from_slice(&value.to_le_bytes());
    }
}
