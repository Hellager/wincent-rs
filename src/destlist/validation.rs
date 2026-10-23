use std::collections::BTreeSet;
use std::path::Path;

use crate::destlist::parser::{AutomaticDestinations, DestListKind};
use crate::error::WincentError;
use crate::WincentResult;

pub(super) fn validate_writer_input(
    parsed: &AutomaticDestinations,
    kind: DestListKind,
) -> WincentResult<()> {
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

    validate_stream_bijection(parsed)?;
    match kind {
        DestListKind::RecentFiles => validate_recent_metadata(parsed),
        DestListKind::FrequentFolders => validate_frequent_metadata(parsed),
        DestListKind::Unknown => Err(WincentError::UnsupportedOperation(
            "unknown DestList kind cannot be rewritten".to_string(),
        )),
    }
}

fn validate_stream_bijection(parsed: &AutomaticDestinations) -> WincentResult<()> {
    let mut numeric_streams = BTreeSet::new();
    for stream in parsed.cfb_info().streams() {
        let Some(name) = canonical_numeric_name(stream.name()) else {
            continue;
        };
        if stream.path().parent() != Some(Path::new("/")) {
            return Err(WincentError::DestListParse(format!(
                "numeric stream {name} is not in the CFB root storage"
            )));
        }
        if !numeric_streams.insert(name.clone()) {
            return Err(WincentError::DestListParse(format!(
                "CFB contains duplicate numeric stream {name}"
            )));
        }
    }

    let mut referenced_streams = BTreeSet::new();
    for entry in parsed.dest_list().entries() {
        let name = entry.stream_name().to_ascii_lowercase();
        if !referenced_streams.insert(name.clone()) {
            return Err(WincentError::DestListParse(format!(
                "numeric stream {name} is referenced by multiple DestList entries"
            )));
        }
        if !numeric_streams.contains(&name) {
            return Err(WincentError::DestListParse(format!(
                "DestList entry {} is missing root numeric stream {name}",
                entry.entry_id()
            )));
        }
    }

    if let Some(orphan) = numeric_streams.difference(&referenced_streams).next() {
        return Err(WincentError::DestListParse(format!(
            "CFB contains orphan numeric stream {orphan}"
        )));
    }
    Ok(())
}

fn canonical_numeric_name(name: &str) -> Option<String> {
    let lower = name.to_ascii_lowercase();
    let value = u32::from_str_radix(&lower, 16).ok()?;
    (format!("{value:x}") == lower).then_some(lower)
}

fn validate_recent_metadata(parsed: &AutomaticDestinations) -> WincentResult<()> {
    let dest_list = parsed.dest_list();
    if dest_list.pinned_entry_count() != 0 {
        return Err(WincentError::DestListParse(
            "Recent Files must declare zero pinned entries".to_string(),
        ));
    }
    for entry in dest_list.entries() {
        if entry.pin_status() != -1 {
            return Err(WincentError::DestListParse(format!(
                "Recent Files entry {} has invalid pin status {}",
                entry.entry_id(),
                entry.pin_status()
            )));
        }
        if entry.recent_rank() != -1 {
            return Err(WincentError::DestListParse(format!(
                "Recent Files entry {} has invalid rank {}",
                entry.entry_id(),
                entry.recent_rank()
            )));
        }
    }
    Ok(())
}

fn validate_frequent_metadata(parsed: &AutomaticDestinations) -> WincentResult<()> {
    let dest_list = parsed.dest_list();
    let entries = dest_list.entries();
    let pinned_count = entries.iter().filter(|entry| entry.is_pinned()).count();
    if usize::try_from(dest_list.pinned_entry_count()).ok() != Some(pinned_count) {
        return Err(WincentError::DestListParse(format!(
            "Frequent Folders pinned count mismatch: declared {}, actual {pinned_count}",
            dest_list.pinned_entry_count()
        )));
    }

    let mut pin_orders = BTreeSet::new();
    let mut ranks = BTreeSet::new();
    for entry in entries {
        if entry.pin_status() < -1 {
            return Err(WincentError::DestListParse(format!(
                "Frequent Folders entry {} has invalid pin status {}",
                entry.entry_id(),
                entry.pin_status()
            )));
        }
        if entry.recent_rank() < 0 || !ranks.insert(entry.recent_rank()) {
            return Err(WincentError::DestListParse(format!(
                "Frequent Folders entry {} has invalid or duplicate rank {}",
                entry.entry_id(),
                entry.recent_rank()
            )));
        }
        if let Some(order) = entry.pin_order() {
            if !pin_orders.insert(order) {
                return Err(WincentError::DestListParse(format!(
                    "Frequent Folders has duplicate pin order {order}"
                )));
            }
            if entry.reserved_78() != Some(0) {
                return Err(WincentError::DestListParse(format!(
                    "pinned Frequent Folders entry {} cannot be hidden",
                    entry.entry_id()
                )));
            }
        }
    }

    validate_contiguous(&ranks, entries.len(), "Frequent Folders rank")?;
    validate_contiguous(&pin_orders, pinned_count, "Frequent Folders pin order")
}

fn validate_contiguous(values: &BTreeSet<i32>, len: usize, label: &str) -> WincentResult<()> {
    for (expected, actual) in (0_i32..).zip(values.iter().copied()).take(len) {
        if expected != actual {
            return Err(WincentError::DestListParse(format!(
                "{label} sequence is not contiguous at {expected}: found {actual}"
            )));
        }
    }
    if values.len() != len {
        return Err(WincentError::DestListParse(format!(
            "{label} sequence has {} values for {len} entries",
            values.len()
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use super::*;
    use crate::destlist::parser::parse_bytes_with_kind;

    #[derive(Clone, Copy)]
    struct TestEntry {
        id: u64,
        pin_status: i32,
        rank: i32,
        reserved_78: u32,
    }

    #[test]
    fn accepts_valid_recent_and_unrelated_stream() -> WincentResult<()> {
        validate(
            &[entry(1, -1, -1), entry(2, -1, -1)],
            &["1", "2", "Keep"],
            DestListKind::RecentFiles,
            None,
        )
    }

    #[test]
    fn accepts_valid_frequent_ordering() -> WincentResult<()> {
        validate(
            &[entry(1, -1, 0), entry(2, 0, 1), entry(3, 1, 2)],
            &["1", "2", "3"],
            DestListKind::FrequentFolders,
            None,
        )
    }

    #[test]
    fn rejects_missing_numeric_stream() {
        assert_error_contains(
            validate(&[entry(1, -1, -1)], &[], DestListKind::RecentFiles, None),
            "missing root numeric stream",
        );
    }

    #[test]
    fn rejects_duplicate_entry_stream_reference() {
        assert_error_contains(
            validate(
                &[entry(1, -1, -1), entry(0x1_0000_0001, -1, -1)],
                &["1"],
                DestListKind::RecentFiles,
                None,
            ),
            "referenced by multiple",
        );
    }

    #[test]
    fn rejects_orphan_numeric_stream() {
        assert_error_contains(
            validate(
                &[entry(1, -1, -1)],
                &["1", "deadbeef"],
                DestListKind::RecentFiles,
                None,
            ),
            "orphan numeric stream deadbeef",
        );
    }

    #[test]
    fn rejects_frequent_pin_order_gap_and_duplicate() {
        for entries in [
            [entry(1, 0, 0), entry(2, 2, 1)],
            [entry(1, 0, 0), entry(2, 0, 1)],
        ] {
            assert_error_contains(
                validate(&entries, &["1", "2"], DestListKind::FrequentFolders, None),
                "pin order",
            );
        }
    }

    #[test]
    fn rejects_frequent_rank_gap_duplicate_and_negative() {
        for entries in [
            [entry(1, -1, 0), entry(2, -1, 2)],
            [entry(1, -1, 0), entry(2, -1, 0)],
            [entry(1, -1, 0), entry(2, -1, -1)],
        ] {
            assert_error_contains(
                validate(&entries, &["1", "2"], DestListKind::FrequentFolders, None),
                "rank",
            );
        }
    }

    #[test]
    fn rejects_frequent_pinned_count_mismatch() {
        assert_error_contains(
            validate(
                &[entry(1, 0, 0)],
                &["1"],
                DestListKind::FrequentFolders,
                Some(0),
            ),
            "pinned count mismatch",
        );
    }

    #[test]
    fn rejects_recent_pin_and_rank_metadata() {
        for invalid in [entry(1, 0, -1), entry(1, -1, 0)] {
            assert!(validate(&[invalid], &["1"], DestListKind::RecentFiles, None,).is_err());
        }
    }

    #[test]
    fn rejects_hidden_pinned_frequent_entry() {
        let mut hidden = entry(1, 0, 0);
        hidden.reserved_78 = 1;
        assert_error_contains(
            validate(&[hidden], &["1"], DestListKind::FrequentFolders, None),
            "cannot be hidden",
        );
    }

    fn entry(id: u64, pin_status: i32, rank: i32) -> TestEntry {
        TestEntry {
            id,
            pin_status,
            rank,
            reserved_78: 0,
        }
    }

    fn validate(
        entries: &[TestEntry],
        streams: &[&str],
        kind: DestListKind,
        pinned_count: Option<u32>,
    ) -> WincentResult<()> {
        let bytes = build_test_cfb(entries, streams, pinned_count);
        let parsed = parse_bytes_with_kind(bytes, kind)?;
        validate_writer_input(&parsed, kind)
    }

    fn assert_error_contains(result: WincentResult<()>, needle: &str) {
        let error = result.expect_err("validation should reject corrupt writer input");
        assert!(
            error.to_string().contains(needle),
            "unexpected error: {error}"
        );
    }

    fn build_test_cfb(
        entries: &[TestEntry],
        streams: &[&str],
        pinned_count: Option<u32>,
    ) -> Vec<u8> {
        let mut dest_list = vec![0_u8; 0x20];
        write_u32(&mut dest_list, 0, 4);
        write_u32(&mut dest_list, 4, entries.len() as u32);
        write_u32(
            &mut dest_list,
            8,
            pinned_count.unwrap_or_else(|| {
                entries.iter().filter(|entry| entry.pin_status >= 0).count() as u32
            }),
        );
        write_u64(
            &mut dest_list,
            0x10,
            entries.last().map_or(0, |entry| entry.id),
        );
        for entry in entries {
            dest_list.extend_from_slice(&build_test_entry(*entry));
        }

        let mut compound =
            ::cfb::CompoundFile::create_with_version(::cfb::Version::V3, Cursor::new(Vec::new()))
                .expect("create test CFB");
        compound
            .create_stream("/DestList")
            .expect("create DestList")
            .write_all(&dest_list)
            .expect("write DestList");
        for stream in streams {
            compound
                .create_stream(format!("/{stream}"))
                .expect("create test stream")
                .write_all(b"test payload")
                .expect("write test stream");
        }
        compound.flush().expect("flush test CFB");
        compound.into_inner().into_inner()
    }

    fn build_test_entry(entry: TestEntry) -> Vec<u8> {
        let path = format!("C:\\Test\\{}.txt", entry.id);
        let path_bytes: Vec<u8> = path.encode_utf16().flat_map(u16::to_le_bytes).collect();
        let mut bytes = vec![0_u8; 0x82 + path_bytes.len() + 4];
        write_u64(&mut bytes, 0x58, entry.id);
        write_u32(&mut bytes, 0x60, 1.0_f32.to_bits());
        write_u32(&mut bytes, 0x6c, entry.pin_status as u32);
        write_u32(&mut bytes, 0x70, entry.rank as u32);
        write_u32(&mut bytes, 0x74, 3);
        write_u32(&mut bytes, 0x78, entry.reserved_78);
        let path_chars = u16::try_from(path.encode_utf16().count()).expect("test path length");
        bytes[0x80..0x82].copy_from_slice(&path_chars.to_le_bytes());
        bytes[0x82..0x82 + path_bytes.len()].copy_from_slice(&path_bytes);
        let checksum = checksum(&bytes);
        write_u64(&mut bytes, 0, checksum);
        bytes
    }

    fn checksum(entry: &[u8]) -> u64 {
        const POLYNOMIAL: u64 = 0x92c6_4265_d321_39a4;
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

    fn write_u32(data: &mut [u8], offset: usize, value: u32) {
        data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    }

    fn write_u64(data: &mut [u8], offset: usize, value: u64) {
        data[offset..offset + 8].copy_from_slice(&value.to_le_bytes());
    }
}
