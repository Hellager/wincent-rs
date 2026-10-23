use std::env;
use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use wincent::destlist::{parse_bytes_with_kind, DestListKind};

const ENTRY_FIXED_LEN: usize = 0x82;
const PIN_STATUS_OFFSET: usize = 0x6c;
const RANK_OFFSET: usize = 0x70;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    if args.len() != 5 {
        return Err("usage: cfb_corrupt <input> <output> <recent|frequent> <case>".into());
    }

    let input = Path::new(&args[1]);
    let output = Path::new(&args[2]);
    let kind = match args[3].as_str() {
        "recent" => DestListKind::RecentFiles,
        "frequent" => DestListKind::FrequentFolders,
        other => return Err(format!("unsupported kind: {other}").into()),
    };
    let case = args[4].as_str();
    let original = fs::read(input)?;

    if case == "truncate-cfb" {
        if original.len() < 1024 {
            return Err("CFB is too short to truncate".into());
        }
        fs::write(output, &original[..original.len() - 512])?;
        return Ok(());
    }
    if matches!(case, "bad-fat" | "bad-minifat" | "bad-directory") {
        let mut corrupted = original;
        let offset = match case {
            "bad-directory" => 0x30,
            "bad-minifat" => 0x3c,
            "bad-fat" => 0x4c,
            _ => unreachable!(),
        };
        write_u32(&mut corrupted, offset, 0xffff_ff00)?;
        fs::write(output, corrupted)?;
        return Ok(());
    }

    let parsed = parse_bytes_with_kind(original.clone(), kind)?;
    let entries = parsed.dest_list().entries();
    let first = entries.first().ok_or("DestList has no entries")?;
    let first_offset = first.entry_offset();
    let first_stream = first.stream_name().to_string();
    let second_stream = entries.get(1).map(|entry| entry.stream_name().to_string());
    let mut compound = cfb::CompoundFile::open(Cursor::new(original))?;
    let mut dest = Vec::new();
    compound.open_stream("/DestList")?.read_to_end(&mut dest)?;

    match case {
        "bad-crc" => {
            let byte = dest
                .get_mut(first_offset)
                .ok_or("checksum offset is out of range")?;
            *byte ^= 0x80;
        }
        "count-mismatch" => {
            let count = read_u32(&dest, 4)?
                .checked_add(1)
                .ok_or("entry count overflow")?;
            write_u32(&mut dest, 4, count)?;
        }
        "truncate-destlist" => {
            dest.truncate(first_offset + first.entry_len().saturating_sub(1));
        }
        "missing-stream" => {
            compound.remove_stream(format!("/{first_stream}"))?;
        }
        "orphan-stream" => {
            let mut numeric = Vec::new();
            compound
                .open_stream(format!("/{first_stream}"))?
                .read_to_end(&mut numeric)?;
            compound.create_stream("/deadbeef")?.write_all(&numeric)?;
        }
        "duplicate-path" => {
            let second = entries.get(1).ok_or("duplicate-path needs two entries")?;
            let first_path_len = first
                .entry_len()
                .checked_sub(ENTRY_FIXED_LEN + 4)
                .ok_or("first entry is shorter than the fixed v4 layout")?;
            let second_path_len = second
                .entry_len()
                .checked_sub(ENTRY_FIXED_LEN + 4)
                .ok_or("second entry is shorter than the fixed v4 layout")?;
            if first_path_len != second_path_len {
                return Err(format!(
                    "duplicate-path requires equal UTF-16 byte lengths: {first_path_len} != \
                     {second_path_len}"
                )
                .into());
            }
            let source_start = first_offset + ENTRY_FIXED_LEN;
            let source_end = source_start + first_path_len;
            let destination_start = second.entry_offset() + ENTRY_FIXED_LEN;
            let destination_end = destination_start + second_path_len;
            let path_bytes = dest
                .get(source_start..source_end)
                .ok_or("first path range is out of bounds")?
                .to_vec();
            dest.get_mut(destination_start..destination_end)
                .ok_or("second path range is out of bounds")?
                .copy_from_slice(&path_bytes);
            let checksum = entry_checksum(&dest, second.entry_offset())?;
            write_u64(&mut dest, second.entry_offset(), checksum)?;

            let mut numeric = Vec::new();
            compound
                .open_stream(format!("/{first_stream}"))?
                .read_to_end(&mut numeric)?;
            let second_stream = second_stream
                .as_deref()
                .ok_or("duplicate-path needs a second numeric stream")?;
            compound
                .create_stream(format!("/{second_stream}"))?
                .write_all(&numeric)?;
        }
        "duplicate-stream" => {
            let second = entries.get(1).ok_or("duplicate-stream needs two entries")?;
            let duplicated_id =
                (second.entry_id() & !u64::from(u32::MAX)) | u64::from(first.entry_number());
            write_u64(&mut dest, second.entry_offset() + 0x58, duplicated_id)?;
            let checksum = entry_checksum(&dest, second.entry_offset())?;
            write_u64(&mut dest, second.entry_offset(), checksum)?;
        }
        "illegal-pin-rank" => {
            write_i32(&mut dest, first_offset + PIN_STATUS_OFFSET, 99)?;
            write_i32(&mut dest, first_offset + RANK_OFFSET, -42)?;
            let checksum = entry_checksum(&dest, first_offset)?;
            write_u64(&mut dest, first_offset, checksum)?;
        }
        "pin-gap" => {
            let pinned = entries
                .iter()
                .find(|entry| entry.is_pinned())
                .ok_or("pin-gap needs a pinned entry")?;
            write_i32(&mut dest, pinned.entry_offset() + PIN_STATUS_OFFSET, 99)?;
            let checksum = entry_checksum(&dest, pinned.entry_offset())?;
            write_u64(&mut dest, pinned.entry_offset(), checksum)?;
        }
        "rank-gap" => {
            let invalid_rank = i32::try_from(entries.len())?
                .checked_add(10)
                .ok_or("rank gap overflow")?;
            write_i32(&mut dest, first_offset + RANK_OFFSET, invalid_rank)?;
            let checksum = entry_checksum(&dest, first_offset)?;
            write_u64(&mut dest, first_offset, checksum)?;
        }
        other => return Err(format!("unsupported corruption case: {other}").into()),
    }

    if case != "missing-stream" {
        let mut stream = compound.create_stream("/DestList")?;
        stream.write_all(&dest)?;
    }
    compound.flush()?;
    fs::write(output, compound.into_inner().into_inner())?;
    Ok(())
}

fn entry_checksum(data: &[u8], offset: usize) -> Result<u64, Box<dyn std::error::Error>> {
    const SPAN: usize = 0x70;
    const POLYNOMIAL: u64 = 0x92C6_4265_D321_39A4;

    let prefix = data
        .get(offset..offset + SPAN)
        .ok_or("entry checksum range is out of bounds")?;
    let mut crc = u64::MAX;
    for (index, byte) in prefix.iter().copied().enumerate() {
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

fn read_u32(bytes: &[u8], offset: usize) -> Result<u32, Box<dyn std::error::Error>> {
    let value = bytes
        .get(offset..offset + 4)
        .ok_or("u32 range is out of bounds")?
        .try_into()?;
    Ok(u32::from_le_bytes(value))
}

fn write_u32(
    bytes: &mut [u8],
    offset: usize,
    value: u32,
) -> Result<(), Box<dyn std::error::Error>> {
    bytes
        .get_mut(offset..offset + 4)
        .ok_or("u32 range is out of bounds")?
        .copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn write_i32(
    bytes: &mut [u8],
    offset: usize,
    value: i32,
) -> Result<(), Box<dyn std::error::Error>> {
    bytes
        .get_mut(offset..offset + 4)
        .ok_or("i32 range is out of bounds")?
        .copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn write_u64(
    bytes: &mut [u8],
    offset: usize,
    value: u64,
) -> Result<(), Box<dyn std::error::Error>> {
    bytes
        .get_mut(offset..offset + 8)
        .ok_or("u64 range is out of bounds")?
        .copy_from_slice(&value.to_le_bytes());
    Ok(())
}
