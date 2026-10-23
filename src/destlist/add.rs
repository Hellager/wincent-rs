//! Direct DestList v4 creation for Recent Files and Frequent Folders.

use std::fs;
use std::io::{Cursor, Write};
use std::path::{Path, PathBuf};
use std::time::SystemTime;

use windows::core::{Interface, GUID, PCWSTR};
use windows::Win32::Foundation::{PROPERTYKEY, VARIANT_BOOL};
use windows::Win32::System::Com::StructuredStorage::{
    CreateStreamOnHGlobal, PROPVARIANT, PROPVARIANT_0, PROPVARIANT_0_0, PROPVARIANT_0_0_0,
};
use windows::Win32::System::Com::{
    CoCreateInstance, IPersistFile, IPersistStream, IStream, CLSCTX_INPROC_SERVER, STREAM_SEEK_SET,
};
use windows::Win32::System::Variant::VT_BOOL;
use windows::Win32::UI::Shell::{IShellLinkDataList, IShellLinkW, SHChangeNotify, ShellLink};
use windows::Win32::UI::Shell::{SHCNE_ASSOCCHANGED, SHCNF_IDLIST};
use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;

use crate::com_thread::run_on_sta_thread;
use crate::destlist::parser::{
    parse_bytes_with_kind, AutomaticDestinations, DestListEntry, DestListKind,
};
use crate::error::WincentError;
use crate::recent_links::resolve_lnk_target;
use crate::utils::{get_windows_recent_folder, normalize_path_lightweight, paths_equal};
use crate::{QuickAccess, WincentResult};

const DEST_LIST_HEADER_LEN: usize = 0x20;
const ENTRY_COUNT_OFFSET: usize = 0x04;
const PINNED_COUNT_OFFSET: usize = 0x08;
const SCORE_COUNTER_OFFSET: usize = 0x0c;
const LAST_ENTRY_ID_OFFSET: usize = 0x10;
const ACTION_COUNT_OFFSET: usize = 0x18;
const ENTRY_FIXED_LEN: usize = 0x82;
const RECENT_LINK_FLAGS: u32 = 0x0020_009b;
const DESTLIST_LINK_FLAGS: u32 = 0x00a0_0000;
const PKEY_APP_USER_MODEL_IS_DEST_LIST_LINK: PROPERTYKEY = PROPERTYKEY {
    fmtid: GUID::from_u128(0x9f4c2855_9f79_4b39_a8d0_e1d42de1d5f3),
    pid: 7,
};

#[derive(Debug, Clone)]
struct AddPlan {
    dest_path: PathBuf,
    kind: DestListKind,
    qa_type: QuickAccess,
}

pub(crate) fn add_entry_path(
    qa_type: QuickAccess,
    target_path: &str,
    timeout: std::time::Duration,
) -> WincentResult<()> {
    let plan = match qa_type {
        QuickAccess::RecentFiles => AddPlan {
            dest_path: super::parser::recent_files_dest_path()?,
            kind: DestListKind::RecentFiles,
            qa_type,
        },
        QuickAccess::FrequentFolders => AddPlan {
            dest_path: super::parser::frequent_folders_dest_path()?,
            kind: DestListKind::FrequentFolders,
            qa_type,
        },
        QuickAccess::All => {
            return Err(WincentError::UnsupportedOperation(
                "QuickAccess::All is not valid for direct DestList addition".to_string(),
            ));
        }
    };

    let write_guard = super::transaction::DirectWriteGuard::acquire(&plan.dest_path)?;
    let target = Path::new(target_path);
    if !target.exists() {
        return Err(WincentError::invalid_path(
            target_path,
            "Path does not exist",
        ));
    }
    let is_dir = target.is_dir();
    if plan.qa_type == QuickAccess::RecentFiles && !target.is_file() {
        return Err(WincentError::invalid_path(target_path, "Not a valid file"));
    }
    if plan.qa_type == QuickAccess::FrequentFolders && !is_dir {
        return Err(WincentError::invalid_path(
            target_path,
            "Not a valid directory",
        ));
    }

    let snapshot = write_guard.read_snapshot(&plan.dest_path)?;
    let original = snapshot.bytes().to_vec();
    let parsed = parse_bytes_with_kind(original.clone(), plan.kind)?;
    validate_original(&parsed)?;
    let target_key = normalize_path_lightweight(target_path);
    if parsed.dest_list().entries().iter().any(|entry| {
        normalize_path_lightweight(entry.path()) == target_key
            || normalize_path_lightweight(entry.raw_path()) == target_key
    }) {
        return Err(WincentError::already_exists(target_path, qa_type));
    }

    let entry_id = next_entry_id(&parsed)?;
    let numeric_stream = create_destlist_shell_link(target_path, timeout)?;
    let old_dest = parsed
        .cfb_info()
        .stream("DestList")
        .ok_or_else(|| WincentError::DestListParse("missing DestList stream".to_string()))?
        .data()
        .to_vec();
    let (dest_bytes, stream_name) = build_destlist_v4(
        parsed.dest_list().entries(),
        &old_dest,
        target_path,
        &numeric_stream,
        plan.kind,
        entry_id,
    )?;
    let rewritten = replace_cfb_streams(original, &dest_bytes, &stream_name, &numeric_stream)?;
    validate_rewritten(&rewritten, plan.kind, target_path, &stream_name)?;
    super::remove::replace_file_atomically(&plan.dest_path, &rewritten, &write_guard, &snapshot)?;
    let live = fs::read(&plan.dest_path).map_err(WincentError::Io)?;
    validate_rewritten(&live, plan.kind, target_path, &stream_name)?;

    let link_result: WincentResult<()> = (|| {
        let recent_folder = PathBuf::from(get_windows_recent_folder()?);
        if plan.kind == DestListKind::RecentFiles {
            let file_name = target.file_stem().ok_or_else(|| {
                WincentError::InvalidArgument("target has no file stem".to_string())
            })?;
            let file_link =
                choose_link_path(&recent_folder, &file_name.to_string_lossy(), target_path)?;
            let parent = target.parent().ok_or_else(|| {
                WincentError::InvalidArgument("target has no parent directory".to_string())
            })?;
            let parent_name = parent.file_name().ok_or_else(|| {
                WincentError::InvalidArgument("target parent has no name".to_string())
            })?;
            let parent_link = choose_link_path(
                &recent_folder,
                &parent_name.to_string_lossy(),
                parent.to_string_lossy().as_ref(),
            )?;
            save_file_shell_link(
                &file_link,
                target_path,
                parent.to_string_lossy().as_ref(),
                timeout,
            )?;
            save_file_shell_link(&parent_link, parent.to_string_lossy().as_ref(), "", timeout)?;
        } else {
            let folder_name = target.file_name().ok_or_else(|| {
                WincentError::InvalidArgument("target directory has no name".to_string())
            })?;
            let folder_link =
                choose_link_path(&recent_folder, &folder_name.to_string_lossy(), target_path)?;
            save_file_shell_link(&folder_link, target_path, "", timeout)?;
        }
        Ok(())
    })();
    link_result.map_err(|source| {
        WincentError::post_mutation_failure(
            target_path,
            qa_type,
            crate::error::QuickAccessPostMutationStep::CreateRecentLinks,
            source,
        )
    })?;
    unsafe {
        // SAFETY: SHChangeNotify accepts null item pointers for ASSOCCHANGED and does not
        // retain them after the synchronous call.
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
    }
    Ok(())
}

fn choose_link_path(recent_folder: &Path, stem: &str, target: &str) -> WincentResult<PathBuf> {
    for suffix in 1..=u32::MAX {
        let name = if suffix == 1 {
            format!("{stem}.lnk")
        } else {
            format!("{stem} ({suffix}).lnk")
        };
        let candidate = recent_folder.join(name);
        if !candidate.exists() {
            return Ok(candidate);
        }
        if resolve_lnk_target(&candidate, std::time::Duration::ZERO)?
            .as_deref()
            .is_some_and(|existing| paths_equal(existing, target))
        {
            return Ok(candidate);
        }
    }
    Err(WincentError::SystemError(format!(
        "could not allocate a Recent shortcut name for {target}"
    )))
}

fn validate_original(parsed: &AutomaticDestinations) -> WincentResult<()> {
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
    Ok(())
}

fn next_entry_id(parsed: &AutomaticDestinations) -> WincentResult<u64> {
    let mut id = parsed
        .dest_list()
        .last_entry_id()
        .max(
            parsed
                .dest_list()
                .entries()
                .iter()
                .map(DestListEntry::entry_id)
                .max()
                .unwrap_or_default(),
        )
        .checked_add(1)
        .ok_or_else(|| WincentError::DestListParse("entry id overflow".to_string()))?;
    while parsed.cfb_info().stream(&format!("{id:x}")).is_some() {
        id = id
            .checked_add(1)
            .ok_or_else(|| WincentError::DestListParse("entry id overflow".to_string()))?;
    }
    Ok(id)
}

fn validate_rewritten(
    rewritten: &[u8],
    kind: DestListKind,
    target_path: &str,
    stream_name: &str,
) -> WincentResult<()> {
    let parsed = parse_bytes_with_kind(rewritten.to_vec(), kind)?;
    validate_original(&parsed)?;
    if parsed.cfb_info().stream(stream_name).is_none() {
        return Err(WincentError::DestListParse(format!(
            "rewritten CFB is missing numeric stream {stream_name}"
        )));
    }
    let target_key = normalize_path_lightweight(target_path);
    if !parsed.rule_entries().iter().any(|entry| {
        normalize_path_lightweight(entry.path()) == target_key
            || normalize_path_lightweight(entry.raw_path()) == target_key
    }) {
        return Err(WincentError::DestListParse(
            "rewritten DestList entry is not visible under v4 rules".to_string(),
        ));
    }
    Ok(())
}

fn build_destlist_v4(
    entries: &[DestListEntry],
    old_dest: &[u8],
    target: &str,
    numeric_stream: &[u8],
    kind: DestListKind,
    entry_id: u64,
) -> WincentResult<(Vec<u8>, String)> {
    if old_dest.len() < DEST_LIST_HEADER_LEN {
        return Err(WincentError::DestListParse(
            "invalid DestList v4 header".to_string(),
        ));
    }
    let version = read_u32(old_dest, 0)?;
    if version != 4 {
        return Err(WincentError::DestListUnsupportedVersion(version));
    }
    let stream_name = format!("{entry_id:x}");
    let path_bytes: Vec<u8> = target.encode_utf16().flat_map(u16::to_le_bytes).collect();
    let mut entry = vec![0u8; ENTRY_FIXED_LEN + path_bytes.len() + 4];
    let tracker = tracker_offset(numeric_stream).ok_or_else(|| {
        WincentError::DestListParse("numeric Shell Link lacks TrackerDataBlock".to_string())
    })?;
    for (src, dst) in [
        (0x20, 0x08),
        (0x30, 0x18),
        (0x40, 0x28),
        (0x50, 0x38),
        (0x10, 0x48),
    ] {
        entry[dst..dst + 16].copy_from_slice(
            numeric_stream
                .get(tracker + src..tracker + src + 16)
                .ok_or_else(|| {
                    WincentError::DestListParse("invalid TrackerDataBlock".to_string())
                })?,
        );
    }
    write_u64(&mut entry, 0x58, entry_id)?;
    let score = if kind == DestListKind::FrequentFolders {
        3.0_f32
    } else {
        1.0_f32
    };
    write_u32(&mut entry, 0x60, score.to_bits())?;
    write_u64(&mut entry, 0x64, filetime_now())?;
    write_i32(&mut entry, 0x6c, -1)?;
    write_i32(
        &mut entry,
        0x70,
        if kind == DestListKind::FrequentFolders {
            0
        } else {
            -1
        },
    )?;
    write_u32(
        &mut entry,
        0x74,
        if kind == DestListKind::FrequentFolders {
            3
        } else {
            1
        },
    )?;
    write_u32(&mut entry, 0x78, 0)?;
    write_u32(&mut entry, 0x7c, 0)?;
    let path_chars = u16::try_from(path_bytes.len() / 2)
        .map_err(|_| WincentError::DestListParse("path is too long".to_string()))?;
    entry[0x80..0x82].copy_from_slice(&path_chars.to_le_bytes());
    entry[0x82..0x82 + path_bytes.len()].copy_from_slice(&path_bytes);
    let checksum = crc64(&entry);
    write_u64(&mut entry, 0, checksum)?;

    let mut adjusted = old_dest.to_vec();
    if kind == DestListKind::FrequentFolders {
        for existing in entries.iter().filter(|existing| !existing.is_pinned()) {
            let rank = existing.recent_rank();
            if rank >= 0 {
                write_i32(
                    &mut adjusted,
                    existing.entry_offset() + 0x70,
                    rank.checked_add(1).ok_or_else(|| {
                        WincentError::DestListParse("Frequent Folder rank overflow".to_string())
                    })?,
                )?;
            }
        }
    }

    let mut result = Vec::with_capacity(adjusted.len() + entry.len());
    result.extend_from_slice(&adjusted[..DEST_LIST_HEADER_LEN]);
    result.extend_from_slice(&entry);
    result.extend_from_slice(&adjusted[DEST_LIST_HEADER_LEN..]);
    write_u32(
        &mut result,
        ENTRY_COUNT_OFFSET,
        read_u32(old_dest, ENTRY_COUNT_OFFSET)?
            .checked_add(1)
            .ok_or_else(|| WincentError::DestListParse("entry count overflow".to_string()))?,
    )?;
    let old_counter = f32::from_bits(read_u32(old_dest, SCORE_COUNTER_OFFSET)?);
    write_u32(
        &mut result,
        SCORE_COUNTER_OFFSET,
        (old_counter + score).to_bits(),
    )?;
    write_u64(&mut result, LAST_ENTRY_ID_OFFSET, entry_id)?;
    let action = read_u64(old_dest, ACTION_COUNT_OFFSET)?;
    let action_increment = if kind == DestListKind::RecentFiles {
        2
    } else {
        1
    };
    write_u64(
        &mut result,
        ACTION_COUNT_OFFSET,
        action
            .checked_add(action_increment)
            .ok_or_else(|| WincentError::DestListParse("action count overflow".to_string()))?,
    )?;
    if kind == DestListKind::FrequentFolders {
        let pinned = read_u32(old_dest, PINNED_COUNT_OFFSET)?;
        write_u32(&mut result, PINNED_COUNT_OFFSET, pinned)?;
    }
    Ok((result, stream_name))
}

fn replace_cfb_streams(
    original: Vec<u8>,
    dest_bytes: &[u8],
    stream_name: &str,
    numeric: &[u8],
) -> WincentResult<Vec<u8>> {
    let mut compound = ::cfb::CompoundFile::open(Cursor::new(original))
        .map_err(|e| WincentError::DestListParse(format!("failed to open CFB: {e}")))?;
    {
        let mut stream = compound
            .create_stream("/DestList")
            .map_err(|e| WincentError::DestListParse(format!("failed to replace DestList: {e}")))?;
        stream.write_all(dest_bytes).map_err(WincentError::Io)?;
    }
    {
        let mut stream = compound
            .create_stream(format!("/{stream_name}"))
            .map_err(|e| {
                WincentError::DestListParse(format!("failed to create numeric stream: {e}"))
            })?;
        stream.write_all(numeric).map_err(WincentError::Io)?;
    }
    compound.flush().map_err(WincentError::Io)?;
    Ok(compound.into_inner().into_inner())
}

fn create_destlist_shell_link(
    target: &str,
    timeout: std::time::Duration,
) -> WincentResult<Vec<u8>> {
    let target = target.to_string();
    run_on_sta_thread(
        move || {
            unsafe {
                // SAFETY: COM is initialized by run_on_sta_thread; CLSID and interface type are valid
                // Shell Link identifiers and the returned object owns its COM reference.
                let link: IShellLinkW = CoCreateInstance(&ShellLink, None, CLSCTX_INPROC_SERVER)
                    .map_err(|e| {
                        WincentError::SystemError(format!("CoCreateInstance ShellLink failed: {e}"))
                    })?;
                let target_wide = wide(&target);
                link.SetPath(PCWSTR::from_raw(target_wide.as_ptr()))
                    .map_err(|e| {
                        WincentError::SystemError(format!("IShellLinkW::SetPath failed: {e}"))
                    })?;
                link.SetShowCmd(SW_SHOWNORMAL).map_err(|e| {
                    WincentError::SystemError(format!("IShellLinkW::SetShowCmd failed: {e}"))
                })?;
                let data_list: IShellLinkDataList = link.cast().map_err(|e| {
                    WincentError::SystemError(format!("IShellLinkDataList cast failed: {e}"))
                })?;
                let flags = data_list
                    .GetFlags()
                    .map_err(|e| WincentError::SystemError(format!("GetFlags failed: {e}")))?;
                data_list
                    .SetFlags(flags | DESTLIST_LINK_FLAGS)
                    .map_err(|e| WincentError::SystemError(format!("SetFlags failed: {e}")))?;
                let store: windows::Win32::UI::Shell::PropertiesSystem::IPropertyStore =
                    link.cast().map_err(|e| {
                        WincentError::SystemError(format!("IPropertyStore cast failed: {e}"))
                    })?;
                let value = PROPVARIANT {
                    Anonymous: PROPVARIANT_0 {
                        Anonymous: std::mem::ManuallyDrop::new(PROPVARIANT_0_0 {
                            vt: VT_BOOL,
                            wReserved1: 0,
                            wReserved2: 0,
                            wReserved3: 0,
                            Anonymous: PROPVARIANT_0_0_0 {
                                boolVal: VARIANT_BOOL(-1),
                            },
                        }),
                    },
                };
                store
                    .SetValue(&PKEY_APP_USER_MODEL_IS_DEST_LIST_LINK, &value)
                    .map_err(|e| {
                        WincentError::SystemError(format!("IPropertyStore::SetValue failed: {e}"))
                    })?;
                store.Commit().map_err(|e| {
                    WincentError::SystemError(format!("IPropertyStore::Commit failed: {e}"))
                })?;
                let stream: IStream =
                    CreateStreamOnHGlobal(windows::Win32::Foundation::HGLOBAL::default(), true)
                        .map_err(|e| {
                            WincentError::SystemError(format!("CreateStreamOnHGlobal failed: {e}"))
                        })?;
                let persist: IPersistStream = link.cast().map_err(|e| {
                    WincentError::SystemError(format!("IPersistStream cast failed: {e}"))
                })?;
                persist.Save(&stream, true).map_err(|e| {
                    WincentError::SystemError(format!("IPersistStream::Save failed: {e}"))
                })?;
                stream
                    .Seek(0, STREAM_SEEK_SET, None)
                    .map_err(|e| WincentError::SystemError(format!("IStream::Seek failed: {e}")))?;
                let mut stat = std::mem::zeroed();
                stream
                    .Stat(&mut stat, windows::Win32::System::Com::STATFLAG_NONAME)
                    .map_err(|e| WincentError::SystemError(format!("IStream::Stat failed: {e}")))?;
                let len = usize::try_from(stat.cbSize).map_err(|_| {
                    WincentError::DestListParse("Shell Link stream too large".to_string())
                })?;
                let mut bytes = vec![0u8; len];
                let mut read = 0u32;
                let hr = stream.Read(
                    bytes.as_mut_ptr().cast(),
                    bytes.len() as u32,
                    Some(&mut read),
                );
                hr.ok()
                    .map_err(|e| WincentError::SystemError(format!("IStream::Read failed: {e}")))?;
                bytes.truncate(read as usize);
                Ok(bytes)
            }
        },
        timeout,
    )
}

fn save_file_shell_link(
    link_path: &Path,
    target: &str,
    working_directory: &str,
    timeout: std::time::Duration,
) -> WincentResult<()> {
    let link_path = link_path.to_path_buf();
    let target = target.to_string();
    let working_directory = working_directory.to_string();
    run_on_sta_thread(
        move || unsafe {
            // SAFETY: COM is initialized by run_on_sta_thread and all PCWSTR buffers outlive calls.
            let link: IShellLinkW = CoCreateInstance(&ShellLink, None, CLSCTX_INPROC_SERVER)
                .map_err(|e| {
                    WincentError::SystemError(format!("CoCreateInstance ShellLink failed: {e}"))
                })?;
            let target_wide = wide(&target);
            link.SetPath(PCWSTR::from_raw(target_wide.as_ptr()))
                .map_err(|e| WincentError::SystemError(format!("SetPath failed: {e}")))?;
            if !working_directory.is_empty() {
                let dir_wide = wide(&working_directory);
                link.SetWorkingDirectory(PCWSTR::from_raw(dir_wide.as_ptr()))
                    .map_err(|e| {
                        WincentError::SystemError(format!("SetWorkingDirectory failed: {e}"))
                    })?;
            }
            let data_list: IShellLinkDataList = link.cast().map_err(|e| {
                WincentError::SystemError(format!("IShellLinkDataList cast failed: {e}"))
            })?;
            let flags = data_list
                .GetFlags()
                .map_err(|e| WincentError::SystemError(format!("GetFlags failed: {e}")))?;
            data_list
                .SetFlags(flags | RECENT_LINK_FLAGS)
                .map_err(|e| WincentError::SystemError(format!("SetFlags failed: {e}")))?;
            let persist: IPersistFile = link
                .cast()
                .map_err(|e| WincentError::SystemError(format!("IPersistFile cast failed: {e}")))?;
            let path_wide = wide(&link_path.to_string_lossy());
            persist
                .Save(PCWSTR::from_raw(path_wide.as_ptr()), true)
                .map_err(|e| {
                    WincentError::SystemError(format!("IPersistFile::Save failed: {e}"))
                })?;
            Ok(())
        },
        timeout,
    )
}

fn tracker_offset(data: &[u8]) -> Option<usize> {
    if data.len() < 0x4c {
        return None;
    }
    let flags = u32::from_le_bytes(data[0x14..0x18].try_into().ok()?);
    let mut offset = 0x4cusize;
    if flags & 1 != 0 {
        offset = offset.checked_add(
            2 + usize::from(u16::from_le_bytes(
                data.get(offset..offset + 2)?.try_into().ok()?,
            )),
        )?;
    }
    if flags & 2 != 0 {
        offset = offset.checked_add(
            usize::try_from(u32::from_le_bytes(
                data.get(offset..offset + 4)?.try_into().ok()?,
            ))
            .ok()?,
        )?;
    }
    for flag in [4u32, 8, 16, 32, 64] {
        if flags & flag != 0 {
            let count = usize::from(u16::from_le_bytes(
                data.get(offset..offset + 2)?.try_into().ok()?,
            ));
            offset = offset.checked_add(2 + count * if flags & 0x80 != 0 { 2 } else { 1 })?;
        }
    }
    while offset.checked_add(8)? <= data.len() {
        let size = usize::try_from(u32::from_le_bytes(
            data[offset..offset + 4].try_into().ok()?,
        ))
        .ok()?;
        if size == 0 {
            break;
        }
        if u32::from_le_bytes(data[offset + 4..offset + 8].try_into().ok()?) == 0xa000_0003 {
            return Some(offset);
        }
        offset = offset.checked_add(size)?;
    }
    None
}

fn crc64(entry: &[u8]) -> u64 {
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

fn filetime_now() -> u64 {
    const EPOCH: u64 = 116_444_736_000_000_000;
    let duration = SystemTime::now()
        .duration_since(SystemTime::UNIX_EPOCH)
        .unwrap_or_default();
    EPOCH + duration.as_secs() * 10_000_000 + u64::from(duration.subsec_nanos() / 100)
}

fn wide(value: &str) -> Vec<u16> {
    value.encode_utf16().chain(Some(0)).collect()
}

fn read_u32(data: &[u8], offset: usize) -> WincentResult<u32> {
    data.get(offset..offset + 4)
        .ok_or_else(|| WincentError::DestListParse(format!("missing u32 at 0x{offset:x}")))
        .map(|b| u32::from_le_bytes(b.try_into().expect("4-byte slice")))
}
fn read_u64(data: &[u8], offset: usize) -> WincentResult<u64> {
    data.get(offset..offset + 8)
        .ok_or_else(|| WincentError::DestListParse(format!("missing u64 at 0x{offset:x}")))
        .map(|b| u64::from_le_bytes(b.try_into().expect("8-byte slice")))
}
fn write_u32(data: &mut [u8], offset: usize, value: u32) -> WincentResult<()> {
    data.get_mut(offset..offset + 4)
        .ok_or_else(|| WincentError::DestListParse(format!("missing u32 at 0x{offset:x}")))
        .map(|b| b.copy_from_slice(&value.to_le_bytes()))
}
fn write_u64(data: &mut [u8], offset: usize, value: u64) -> WincentResult<()> {
    data.get_mut(offset..offset + 8)
        .ok_or_else(|| WincentError::DestListParse(format!("missing u64 at 0x{offset:x}")))
        .map(|b| b.copy_from_slice(&value.to_le_bytes()))
}
fn write_i32(data: &mut [u8], offset: usize, value: i32) -> WincentResult<()> {
    write_u32(data, offset, value as u32)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Read;

    #[test]
    fn builds_recent_v4_entry_and_header() -> WincentResult<()> {
        let old = empty_destlist(2.5, 7);
        let numeric = shell_link_with_tracker();

        let (dest, stream_name) = build_destlist_v4(
            &[],
            &old,
            "C:\\Work\\report.txt",
            &numeric,
            DestListKind::RecentFiles,
            4,
        )?;

        assert_eq!(stream_name, "4");
        assert_eq!(read_u32(&dest, ENTRY_COUNT_OFFSET)?, 1);
        assert_eq!(f32::from_bits(read_u32(&dest, SCORE_COUNTER_OFFSET)?), 3.5);
        assert_eq!(read_u64(&dest, LAST_ENTRY_ID_OFFSET)?, 4);
        assert_eq!(read_u64(&dest, ACTION_COUNT_OFFSET)?, 9);
        assert_eq!(
            read_u32(&dest, DEST_LIST_HEADER_LEN + 0x60)?,
            1.0_f32.to_bits()
        );
        assert_eq!(read_u32(&dest, DEST_LIST_HEADER_LEN + 0x6c)?, u32::MAX);
        assert_eq!(read_u32(&dest, DEST_LIST_HEADER_LEN + 0x70)?, u32::MAX);
        assert_eq!(read_u32(&dest, DEST_LIST_HEADER_LEN + 0x74)?, 1);
        assert_eq!(
            read_u64(&dest, DEST_LIST_HEADER_LEN)?,
            crc64(&dest[DEST_LIST_HEADER_LEN..])
        );
        Ok(())
    }

    #[test]
    fn builds_visible_unpinned_frequent_v4_entry() -> WincentResult<()> {
        let old = empty_destlist(10.0, 20);
        let numeric = shell_link_with_tracker();

        let (dest, _) = build_destlist_v4(
            &[],
            &old,
            "C:\\Work\\Folder",
            &numeric,
            DestListKind::FrequentFolders,
            8,
        )?;

        assert_eq!(read_u32(&dest, ENTRY_COUNT_OFFSET)?, 1);
        assert_eq!(read_u32(&dest, PINNED_COUNT_OFFSET)?, 0);
        assert_eq!(f32::from_bits(read_u32(&dest, SCORE_COUNTER_OFFSET)?), 13.0);
        assert_eq!(read_u64(&dest, ACTION_COUNT_OFFSET)?, 21);
        assert_eq!(
            read_u32(&dest, DEST_LIST_HEADER_LEN + 0x60)?,
            3.0_f32.to_bits()
        );
        assert_eq!(read_u32(&dest, DEST_LIST_HEADER_LEN + 0x6c)?, u32::MAX);
        assert_eq!(read_u32(&dest, DEST_LIST_HEADER_LEN + 0x70)?, 0);
        assert_eq!(read_u32(&dest, DEST_LIST_HEADER_LEN + 0x74)?, 3);
        Ok(())
    }

    #[test]
    fn rejects_shell_link_without_tracker_block() {
        let error = build_destlist_v4(
            &[],
            &empty_destlist(0.0, 0),
            "C:\\Work\\report.txt",
            &[0; 0x4c],
            DestListKind::RecentFiles,
            1,
        )
        .unwrap_err();

        assert!(
            matches!(error, WincentError::DestListParse(message) if message.contains("TrackerDataBlock"))
        );
    }

    #[test]
    fn complete_rewrite_contains_all_streams_before_commit() -> WincentResult<()> {
        let old_dest = empty_destlist(2.0, 4);
        let new_dest = empty_destlist(3.0, 6);
        let numeric = b"numeric-stream";
        let original = {
            let cursor = Cursor::new(Vec::new());
            let mut compound = cfb::CompoundFile::create_with_version(cfb::Version::V3, cursor)
                .expect("create test CFB");
            compound
                .create_stream("/DestList")
                .expect("create DestList")
                .write_all(&old_dest)
                .expect("write old DestList");
            compound
                .create_stream("/Keep")
                .expect("create unrelated stream")
                .write_all(b"keep")
                .expect("write unrelated stream");
            compound.flush().expect("flush test CFB");
            compound.into_inner().into_inner()
        };
        let rewritten = replace_cfb_streams(original, &new_dest, "4", numeric)?;

        let mut compound =
            cfb::CompoundFile::open(Cursor::new(rewritten)).expect("open rewritten test CFB");
        let mut actual_dest = Vec::new();
        compound
            .open_stream("/DestList")
            .expect("open DestList")
            .read_to_end(&mut actual_dest)
            .expect("read DestList");
        let mut actual_numeric = Vec::new();
        compound
            .open_stream("/4")
            .expect("open numeric stream")
            .read_to_end(&mut actual_numeric)
            .expect("read numeric stream");
        let mut unrelated = Vec::new();
        compound
            .open_stream("/Keep")
            .expect("open unrelated stream")
            .read_to_end(&mut unrelated)
            .expect("read unrelated stream");

        assert_eq!(actual_dest, new_dest);
        assert_eq!(actual_numeric, numeric);
        assert_eq!(unrelated, b"keep");
        Ok(())
    }

    fn empty_destlist(counter: f32, action_count: u64) -> Vec<u8> {
        let mut dest = vec![0u8; DEST_LIST_HEADER_LEN];
        write_u32(&mut dest, 0, 4).expect("version field");
        write_u32(&mut dest, SCORE_COUNTER_OFFSET, counter.to_bits()).expect("counter field");
        write_u64(&mut dest, ACTION_COUNT_OFFSET, action_count).expect("action field");
        dest
    }

    fn shell_link_with_tracker() -> Vec<u8> {
        let mut link = vec![0u8; 0x4c + 0x60];
        link[0..4].copy_from_slice(&0x4c_u32.to_le_bytes());
        link[0x4c..0x50].copy_from_slice(&0x60_u32.to_le_bytes());
        link[0x50..0x54].copy_from_slice(&0xa000_0003_u32.to_le_bytes());
        for (index, byte) in link[0x5c..0xac].iter_mut().enumerate() {
            *byte = u8::try_from(index + 1).expect("test tracker byte");
        }
        link
    }
}
