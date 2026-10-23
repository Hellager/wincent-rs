//! Direct removal from Explorer automatic destination CFB files.
//!
//! This is intentionally narrower than the parser: only DestList v4 is
//! writable because its entry layout and removal bookkeeping have been
//! validated against the project's Windows 10 experiments.

use std::fs::{self, OpenOptions};
use std::io::{Cursor, Write};
use std::os::windows::ffi::OsStrExt;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};

use windows::core::PCWSTR;
use windows::Win32::Storage::FileSystem::{ReplaceFileW, REPLACEFILE_WRITE_THROUGH};

use crate::error::WincentError;
use crate::utils::normalize_path_lightweight;
use crate::{QuickAccess, WincentResult};

use super::parser::{
    frequent_folders_dest_path, parse_bytes_with_kind, recent_files_dest_path, DestListEntry,
    DestListKind,
};

const DEST_LIST_HEADER_LEN: usize = 0x20;
const ENTRY_COUNT_OFFSET: usize = 0x04;
const PINNED_COUNT_OFFSET: usize = 0x08;
const SCORE_COUNTER_OFFSET: usize = 0x0c;
const ACTION_COUNT_OFFSET: usize = 0x18;

static TEMP_FILE_COUNTER: AtomicU64 = AtomicU64::new(0);

#[derive(Debug, Clone)]
struct RemovedEntry {
    offset: usize,
    len: usize,
    stream_name: String,
    pinned: bool,
    score: f32,
}

pub(crate) fn remove_entry_path(qa_type: QuickAccess, target_path: &str) -> WincentResult<()> {
    let (dest_path, kind) = match qa_type {
        QuickAccess::RecentFiles => (recent_files_dest_path()?, DestListKind::RecentFiles),
        QuickAccess::FrequentFolders => {
            (frequent_folders_dest_path()?, DestListKind::FrequentFolders)
        }
        QuickAccess::All => {
            return Err(WincentError::UnsupportedOperation(
                "QuickAccess::All is not valid for direct DestList removal".to_string(),
            ));
        }
    };

    remove_entry_path_from_file(&dest_path, kind, qa_type, target_path)
}

fn remove_entry_path_from_file(
    dest_path: &Path,
    kind: DestListKind,
    qa_type: QuickAccess,
    target_path: &str,
) -> WincentResult<()> {
    let write_guard = super::transaction::DirectWriteGuard::acquire(dest_path)?;
    let snapshot = write_guard.read_snapshot(dest_path)?;
    let original = snapshot.bytes().to_vec();
    let rewritten = rewrite_cfb_bytes(original, kind, qa_type, target_path)?;
    replace_file_atomically(dest_path, &rewritten, &write_guard, &snapshot)
}

fn rewrite_cfb_bytes(
    original: Vec<u8>,
    kind: DestListKind,
    qa_type: QuickAccess,
    target_path: &str,
) -> WincentResult<Vec<u8>> {
    let parsed = parse_bytes_with_kind(original.clone(), kind)?;
    let dest_list = parsed.dest_list();
    if dest_list.version() != 4 {
        return Err(WincentError::DestListUnsupportedVersion(
            dest_list.version(),
        ));
    }
    if dest_list.declared_entry_count() != dest_list.entries().len() {
        return Err(WincentError::DestListParse(format!(
            "cannot rewrite a partial DestList: declared {}, parsed {}",
            dest_list.declared_entry_count(),
            dest_list.entries().len()
        )));
    }
    if dest_list.entries().iter().any(|entry| {
        !entry.checksum_valid()
            || matches!(
                entry.score().partial_cmp(&0.0),
                Some(std::cmp::Ordering::Less) | None
            )
    }) {
        return Err(WincentError::DestListParse(
            "cannot rewrite an invalid DestList v4 table".to_string(),
        ));
    }

    let target_key = normalize_path_lightweight(target_path);
    let removed: Vec<RemovedEntry> = dest_list
        .entries()
        .iter()
        .filter(|entry| entry_matches_path(entry, &target_key))
        .map(|entry| RemovedEntry {
            offset: entry.entry_offset(),
            len: entry.entry_len(),
            stream_name: entry.stream_name().to_string(),
            pinned: entry.is_pinned(),
            score: entry.score(),
        })
        .collect();
    if removed.is_empty() {
        return Err(WincentError::not_in_quick_access(target_path, qa_type));
    }

    for entry in dest_list.entries() {
        if !entry_matches_path(entry, &target_key)
            && removed.iter().any(|removed| {
                removed
                    .stream_name
                    .eq_ignore_ascii_case(entry.stream_name())
            })
        {
            return Err(WincentError::DestListParse(format!(
                "numeric stream {} is referenced by both removed and retained entries",
                entry.stream_name()
            )));
        }
    }

    let mut dest_bytes = parsed
        .cfb_info()
        .stream("DestList")
        .ok_or_else(|| WincentError::DestListParse("missing DestList stream".to_string()))?
        .data()
        .to_vec();
    update_header(&mut dest_bytes, dest_list.entries(), &removed, kind)?;
    for entry in removed.iter().rev() {
        let end = entry.offset.checked_add(entry.len).ok_or_else(|| {
            WincentError::DestListParse("DestList entry range overflow".to_string())
        })?;
        if entry.offset < DEST_LIST_HEADER_LEN || end > dest_bytes.len() {
            return Err(WincentError::DestListParse(format!(
                "invalid DestList entry range 0x{:x}..0x{end:x}",
                entry.offset
            )));
        }
        dest_bytes.drain(entry.offset..end);
    }

    let mut compound = ::cfb::CompoundFile::open(Cursor::new(original)).map_err(|error| {
        WincentError::DestListParse(format!("failed to open writable CFB image: {error}"))
    })?;
    {
        let mut stream = compound.create_stream("/DestList").map_err(|error| {
            WincentError::DestListParse(format!("failed to replace DestList stream: {error}"))
        })?;
        stream.write_all(&dest_bytes).map_err(WincentError::Io)?;
    }
    for stream_name in unique_stream_names(&removed) {
        if parsed.cfb_info().stream(&stream_name).is_some() {
            compound
                .remove_stream(format!("/{stream_name}"))
                .map_err(|error| {
                    WincentError::DestListParse(format!(
                        "failed to remove numeric stream {stream_name}: {error}"
                    ))
                })?;
        }
    }
    compound.flush().map_err(WincentError::Io)?;
    let rewritten = compound.into_inner().into_inner();
    validate_rewritten(&rewritten, kind, &target_key, &removed)?;
    Ok(rewritten)
}

fn entry_matches_path(entry: &DestListEntry, target_key: &str) -> bool {
    normalize_path_lightweight(entry.path()) == target_key
        || normalize_path_lightweight(entry.raw_path()) == target_key
}

fn unique_stream_names(entries: &[RemovedEntry]) -> Vec<String> {
    let mut names = Vec::new();
    for entry in entries {
        if !names
            .iter()
            .any(|name: &String| name.eq_ignore_ascii_case(&entry.stream_name))
        {
            names.push(entry.stream_name.clone());
        }
    }
    names
}

fn update_header(
    dest_bytes: &mut [u8],
    entries: &[DestListEntry],
    removed: &[RemovedEntry],
    kind: DestListKind,
) -> WincentResult<()> {
    if dest_bytes.len() < DEST_LIST_HEADER_LEN {
        return Err(WincentError::DestListParse(
            "DestList stream is shorter than its v4 header".to_string(),
        ));
    }

    let entry_count = read_u32_at(dest_bytes, ENTRY_COUNT_OFFSET)?;
    let removed_count = u32::try_from(removed.len()).map_err(|_| {
        WincentError::DestListParse("too many entries selected for removal".to_string())
    })?;
    write_u32_at(
        dest_bytes,
        ENTRY_COUNT_OFFSET,
        entry_count.checked_sub(removed_count).ok_or_else(|| {
            WincentError::DestListParse("DestList entry count underflow".to_string())
        })?,
    )?;

    let pinned_count = read_u32_at(dest_bytes, PINNED_COUNT_OFFSET)?;
    let removed_pinned = u32::try_from(removed.iter().filter(|entry| entry.pinned).count())
        .map_err(|_| WincentError::DestListParse("pinned entry count overflow".to_string()))?;
    write_u32_at(
        dest_bytes,
        PINNED_COUNT_OFFSET,
        pinned_count.checked_sub(removed_pinned).ok_or_else(|| {
            WincentError::DestListParse("DestList pinned count underflow".to_string())
        })?,
    )?;

    let current_counter = f32::from_bits(read_u32_at(dest_bytes, SCORE_COUNTER_OFFSET)?);
    let removed_score: f32 = removed.iter().map(|entry| entry.score).sum();
    let mut new_counter = current_counter - removed_score;
    if new_counter.is_nan() {
        new_counter = entries
            .iter()
            .filter(|entry| {
                !removed.iter().any(|removed| {
                    removed.offset == entry.entry_offset() && removed.len == entry.entry_len()
                })
            })
            .map(DestListEntry::score)
            .sum();
    }
    if new_counter < 0.0 {
        new_counter = 0.0;
    }
    write_u32_at(dest_bytes, SCORE_COUNTER_OFFSET, new_counter.to_bits())?;

    // Observed v4 Recent physical deletion keeps this counter unchanged, while
    // Frequent pinned-entry deletion advances it once per removed entry.
    if kind == DestListKind::FrequentFolders {
        let action_count = read_u64_at(dest_bytes, ACTION_COUNT_OFFSET)?;
        write_u64_at(
            dest_bytes,
            ACTION_COUNT_OFFSET,
            action_count
                .checked_add(u64::from(removed_count))
                .ok_or_else(|| {
                    WincentError::DestListParse("DestList action count overflow".to_string())
                })?,
        )?;
    }
    Ok(())
}

fn validate_rewritten(
    rewritten: &[u8],
    kind: DestListKind,
    target_key: &str,
    removed: &[RemovedEntry],
) -> WincentResult<()> {
    let parsed = parse_bytes_with_kind(rewritten.to_vec(), kind)?;
    if parsed.dest_list().version() != 4
        || parsed.dest_list().declared_entry_count() != parsed.dest_list().entries().len()
        || parsed
            .dest_list()
            .entries()
            .iter()
            .any(|entry| !entry.checksum_valid())
    {
        return Err(WincentError::DestListParse(
            "rewritten DestList failed structural validation".to_string(),
        ));
    }
    if parsed
        .dest_list()
        .entries()
        .iter()
        .any(|entry| entry_matches_path(entry, target_key))
    {
        return Err(WincentError::DestListParse(
            "rewritten DestList still contains the requested path".to_string(),
        ));
    }
    for stream_name in unique_stream_names(removed) {
        if parsed.cfb_info().stream(&stream_name).is_some() {
            return Err(WincentError::DestListParse(format!(
                "rewritten CFB still contains numeric stream {stream_name}"
            )));
        }
    }
    Ok(())
}

pub(super) fn replace_file_atomically(
    path: &Path,
    data: &[u8],
    write_guard: &super::transaction::DirectWriteGuard,
    snapshot: &super::transaction::BackingFileSnapshot,
) -> WincentResult<()> {
    let temp_path = temporary_replacement_path(path)?;
    let result = (|| {
        let mut temp = OpenOptions::new()
            .create_new(true)
            .write(true)
            .open(&temp_path)
            .map_err(WincentError::Io)?;
        temp.write_all(data).map_err(WincentError::Io)?;
        temp.sync_all().map_err(WincentError::Io)?;
        drop(temp);

        let _verification = write_guard.verify_unchanged(path, snapshot)?;

        let path_wide = wide_null(path);
        let temp_wide = wide_null(&temp_path);
        // SAFETY: Both paths are valid, null-terminated UTF-16 buffers that remain alive for
        // the call. The replacement is on the same volume and was fully flushed above.
        unsafe {
            ReplaceFileW(
                PCWSTR::from_raw(path_wide.as_ptr()),
                PCWSTR::from_raw(temp_wide.as_ptr()),
                PCWSTR::null(),
                REPLACEFILE_WRITE_THROUGH,
                None,
                None,
            )
        }
        .map_err(|error| {
            WincentError::SystemError(format!(
                "failed to atomically replace {}: {error}",
                path.display()
            ))
        })
    })();

    if result.is_err() {
        let _ = fs::remove_file(&temp_path);
    }
    result
}

fn temporary_replacement_path(path: &Path) -> WincentResult<PathBuf> {
    let parent = path.parent().ok_or_else(|| {
        WincentError::InvalidArgument(format!(
            "DestList path has no parent directory: {}",
            path.display()
        ))
    })?;
    let name = path.file_name().ok_or_else(|| {
        WincentError::InvalidArgument(format!(
            "DestList path has no file name: {}",
            path.display()
        ))
    })?;
    let counter = TEMP_FILE_COUNTER.fetch_add(1, Ordering::Relaxed);
    Ok(parent.join(format!(
        ".{}.wincent-{}-{counter}.tmp",
        name.to_string_lossy(),
        std::process::id()
    )))
}

fn wide_null(path: &Path) -> Vec<u16> {
    path.as_os_str().encode_wide().chain(Some(0)).collect()
}

fn read_u32_at(data: &[u8], offset: usize) -> WincentResult<u32> {
    let bytes = data.get(offset..offset + 4).ok_or_else(|| {
        WincentError::DestListParse(format!("missing u32 at DestList offset 0x{offset:x}"))
    })?;
    Ok(u32::from_le_bytes(bytes.try_into().expect("4-byte slice")))
}

fn read_u64_at(data: &[u8], offset: usize) -> WincentResult<u64> {
    let bytes = data.get(offset..offset + 8).ok_or_else(|| {
        WincentError::DestListParse(format!("missing u64 at DestList offset 0x{offset:x}"))
    })?;
    Ok(u64::from_le_bytes(bytes.try_into().expect("8-byte slice")))
}

fn write_u32_at(data: &mut [u8], offset: usize, value: u32) -> WincentResult<()> {
    let bytes = data.get_mut(offset..offset + 4).ok_or_else(|| {
        WincentError::DestListParse(format!("missing u32 at DestList offset 0x{offset:x}"))
    })?;
    bytes.copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn write_u64_at(data: &mut [u8], offset: usize, value: u64) -> WincentResult<()> {
    let bytes = data.get_mut(offset..offset + 8).ok_or_else(|| {
        WincentError::DestListParse(format!("missing u64 at DestList offset 0x{offset:x}"))
    })?;
    bytes.copy_from_slice(&value.to_le_bytes());
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Read, Write};

    #[test]
    fn removes_entry_stream_and_updates_v4_header() -> WincentResult<()> {
        let source = build_test_cfb(&[
            (1, "C:\\Work\\keep.txt", 2.0, -1),
            (2, "C:\\Work\\remove.txt", 3.0, 4),
        ]);

        let rewritten = rewrite_cfb_bytes(
            source,
            DestListKind::RecentFiles,
            QuickAccess::RecentFiles,
            "c:/work/remove.txt",
        )?;
        let parsed = parse_bytes_with_kind(rewritten, DestListKind::RecentFiles)?;

        assert_eq!(parsed.dest_list().declared_entry_count(), 1);
        assert_eq!(parsed.dest_list().pinned_entry_count(), 0);
        assert_eq!(parsed.dest_list().header_counter_f32(), 2.0);
        assert_eq!(parsed.dest_list().add_delete_action_count(), 7);
        assert_eq!(parsed.dest_list().entries()[0].path(), "C:\\Work\\keep.txt");
        assert!(parsed.cfb_info().stream("1").is_some());
        assert!(parsed.cfb_info().stream("2").is_none());
        Ok(())
    }

    #[test]
    fn frequent_removal_increments_action_counter() -> WincentResult<()> {
        let source = build_test_cfb(&[(1, "C:\\Work", 0.0, 4)]);

        let rewritten = rewrite_cfb_bytes(
            source,
            DestListKind::FrequentFolders,
            QuickAccess::FrequentFolders,
            "c:/work",
        )?;
        let parsed = parse_bytes_with_kind(rewritten, DestListKind::FrequentFolders)?;

        assert_eq!(parsed.dest_list().declared_entry_count(), 0);
        assert_eq!(parsed.dest_list().pinned_entry_count(), 0);
        assert_eq!(parsed.dest_list().add_delete_action_count(), 8);
        Ok(())
    }

    #[test]
    fn missing_path_does_not_modify_image() {
        let source = build_test_cfb(&[(1, "C:\\Work\\keep.txt", 1.0, -1)]);
        let error = rewrite_cfb_bytes(
            source,
            DestListKind::RecentFiles,
            QuickAccess::RecentFiles,
            "C:\\Work\\missing.txt",
        )
        .unwrap_err();

        assert!(matches!(error, WincentError::NotInQuickAccess { .. }));
    }

    #[test]
    fn rejects_non_v4_destlist() {
        let mut source = build_test_cfb(&[(1, "C:\\Work\\keep.txt", 1.0, -1)]);
        let mut compound = ::cfb::CompoundFile::open(Cursor::new(source)).expect("open test CFB");
        let mut dest_list = Vec::new();
        compound
            .open_stream("/DestList")
            .expect("open DestList")
            .read_to_end(&mut dest_list)
            .expect("read DestList");
        dest_list[0..4].copy_from_slice(&6_u32.to_le_bytes());
        compound
            .create_stream("/DestList")
            .expect("replace DestList")
            .write_all(&dest_list)
            .expect("write DestList");
        compound.flush().expect("flush CFB");
        source = compound.into_inner().into_inner();

        let error = rewrite_cfb_bytes(
            source,
            DestListKind::RecentFiles,
            QuickAccess::RecentFiles,
            "C:\\Work\\keep.txt",
        )
        .unwrap_err();
        assert!(matches!(error, WincentError::DestListUnsupportedVersion(6)));
    }

    fn build_test_cfb(entries: &[(u32, &str, f32, i32)]) -> Vec<u8> {
        let mut dest_list = vec![0_u8; DEST_LIST_HEADER_LEN];
        write_test_u32(&mut dest_list, 0, 4);
        write_test_u32(&mut dest_list, ENTRY_COUNT_OFFSET, entries.len() as u32);
        write_test_u32(
            &mut dest_list,
            PINNED_COUNT_OFFSET,
            entries.iter().filter(|entry| entry.3 >= 0).count() as u32,
        );
        write_test_u32(
            &mut dest_list,
            SCORE_COUNTER_OFFSET,
            entries.iter().map(|entry| entry.2).sum::<f32>().to_bits(),
        );
        let last_id = entries.last().map_or(0, |entry| u64::from(entry.0));
        write_test_u64(&mut dest_list, 0x10, last_id);
        write_test_u64(&mut dest_list, ACTION_COUNT_OFFSET, 7);

        for (id, path, score, pin_status) in entries {
            dest_list.extend_from_slice(&build_test_entry(*id, path, *score, *pin_status));
        }

        let cursor = Cursor::new(Vec::new());
        let mut compound = ::cfb::CompoundFile::create_with_version(::cfb::Version::V3, cursor)
            .expect("create test CFB");
        compound
            .create_stream("/DestList")
            .expect("create DestList")
            .write_all(&dest_list)
            .expect("write DestList");
        for (id, _, _, _) in entries {
            compound
                .create_stream(format!("/{id:x}"))
                .expect("create numeric stream")
                .write_all(b"test shell link")
                .expect("write numeric stream");
        }
        compound.flush().expect("flush test CFB");
        compound.into_inner().into_inner()
    }

    fn build_test_entry(id: u32, path: &str, score: f32, pin_status: i32) -> Vec<u8> {
        let path_bytes: Vec<u8> = path.encode_utf16().flat_map(u16::to_le_bytes).collect();
        let mut entry = vec![0_u8; 0x82 + path_bytes.len() + 4];
        write_test_u64(&mut entry, 0x58, u64::from(id));
        write_test_u32(&mut entry, 0x60, score.to_bits());
        write_test_u32(&mut entry, 0x6c, pin_status as u32);
        write_test_u32(&mut entry, 0x70, u32::MAX);
        write_test_u32(&mut entry, 0x74, 3);
        let path_chars = u16::try_from(path.encode_utf16().count()).expect("test path length");
        entry[0x80..0x82].copy_from_slice(&path_chars.to_le_bytes());
        entry[0x82..0x82 + path_bytes.len()].copy_from_slice(&path_bytes);
        let checksum = test_entry_checksum(&entry);
        write_test_u64(&mut entry, 0, checksum);
        entry
    }

    fn test_entry_checksum(entry: &[u8]) -> u64 {
        const POLYNOMIAL: u64 = 0x92C6_4265_D321_39A4;
        let mut crc = u64::MAX;
        for (index, byte) in entry[..0x70].iter().copied().enumerate() {
            crc ^= u64::from(if index < 8 { 0 } else { byte });
            for _ in 0..8 {
                crc = if crc & 1 == 0 {
                    crc >> 1
                } else {
                    (crc >> 1) ^ POLYNOMIAL
                };
            }
        }
        crc
    }

    fn write_test_u32(data: &mut [u8], offset: usize, value: u32) {
        data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    }

    fn write_test_u64(data: &mut [u8], offset: usize, value: u64) {
        data[offset..offset + 8].copy_from_slice(&value.to_le_bytes());
    }
}
