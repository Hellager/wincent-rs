//! Shared in-memory fixtures for DestList tests.
//!
//! These helpers build parser values without touching the filesystem or
//! requiring a real Compound File. They are shared between the parser tests and
//! the status-query tests so both exercise the same synthetic layouts.

use std::path::PathBuf;

use crate::destlist::parser::{
    AutomaticDestinations, CfbDirectoryEntry, CfbInfo, CfbStream, DestList, DestListEntry,
    DestListKind, PathSource,
};

pub(crate) fn dest_list_for_test(version: u32, entries: Vec<DestListEntry>) -> DestList {
    DestList {
        version,
        declared_entry_count: entries.len(),
        pinned_entry_count: entries.iter().filter(|entry| entry.is_pinned()).count() as u32,
        header_counter_raw: 0,
        header_counter_f32: 0.0,
        last_entry_id: entries.last().map(DestListEntry::entry_id).unwrap_or(0),
        last_entry_number: entries.last().map(DestListEntry::entry_number).unwrap_or(0),
        add_delete_action_count: 0,
        entries,
        diagnostics: Vec::new(),
    }
}

pub(crate) fn automatic_destinations_for_test(
    kind: DestListKind,
    entries: Vec<DestListEntry>,
    stream_names: &[String],
) -> AutomaticDestinations {
    AutomaticDestinations {
        kind,
        cfb_info: CfbInfo {
            major_version: 3,
            sector_size: 512,
            mini_sector_size: 64,
            mini_cutoff_size: 4096,
            directory_entries: stream_names
                .iter()
                .map(|name| CfbDirectoryEntry {
                    name: name.clone(),
                    path: PathBuf::from(format!("/{name}")),
                    object_type: 2,
                    start_sector: None,
                    stream_size: 1,
                })
                .collect(),
            streams: stream_names
                .iter()
                .map(|name| CfbStream {
                    name: name.clone(),
                    path: PathBuf::from(format!("/{name}")),
                    data: {
                        let mut data = vec![0u8; 0x4c];
                        data[0..4].copy_from_slice(&0x4cu32.to_le_bytes());
                        data[4..20].copy_from_slice(&[
                            0x01, 0x14, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00,
                            0x00, 0x00, 0x00, 0x46,
                        ]);
                        data
                    },
                })
                .collect(),
        },
        dest_list: dest_list_for_test(4, entries),
    }
}

pub(crate) fn destlist_entry_for_test(
    path: &str,
    entry_number: u32,
    entry_number_unknown: u32,
    recent_rank: i32,
    access_count: u32,
    pin_order: Option<i32>,
) -> DestListEntry {
    let entry_id = ((entry_number_unknown as u64) << 32) | entry_number as u64;
    DestListEntry {
        entry_offset: 0,
        entry_len: 0,
        mru_position: 0,
        checksum: 0,
        checksum_valid: true,
        entry_id,
        entry_number,
        entry_number_unknown,
        hostname: String::new(),
        volume_droid: String::new(),
        file_droid: String::new(),
        volume_birth_droid: String::new(),
        file_birth_droid: String::new(),
        file_droid_mac: String::new(),
        stream_name: format!("{entry_number:x}"),
        raw_path: path.to_string(),
        path: path.to_string(),
        pin_status: pin_order.unwrap_or(-1),
        pin_order,
        rank: recent_rank,
        recent_rank,
        count: access_count,
        access_count,
        score: 0.0,
        last_access_filetime: Some(100 + entry_number as u64),
        last_interaction_filetime: Some(100 + entry_number as u64),
        sps_size: None,
        reserved_78: None,
        reserved_7c: None,
        path_sources: vec![PathSource::new("test", path)],
        warnings: Vec::new(),
    }
}
