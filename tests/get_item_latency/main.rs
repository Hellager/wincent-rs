//! VM-only latency test for Quick Access list retrieval.
//!
//! Measures two independent retrieval paths for the same on-disk fixture:
//! - live Shell namespace enumeration (STA COM), and
//! - direct `.automaticDestinations-ms` CFB parsing.
//!
//! The fixture contains 120 Recent Files raw entries (20 visible) and 100
//! Frequent Folders raw entries (46 pinned + 4 visible unpinned + 50 hidden
//! unpinned). It is then reduced to 8 visible Recent Files and 40 visible
//! Frequent Folders for the "under limit" scenarios.
//!
//! This test mutates real Quick Access state and restarts Explorer. It is
//! ignored by default and additionally gated behind `WINCENT_GET_ITEM_LATENCY=1`;
//! run it only inside the dedicated Windows test VM.

#![cfg(windows)]

use std::collections::BTreeMap;
use std::env;
use std::fs;
use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::mpsc;
use std::time::{Duration, Instant};

use wincent::destlist::{
    frequent_folders_dest_path, parse_file, recent_files_dest_path, DestListEntry,
};
use wincent::prelude::*;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER,
    COINIT_APARTMENTTHREADED,
};
use windows::Win32::System::Variant::VARIANT;
use windows::Win32::UI::Shell::{Folder, FolderItems, IShellDispatch, Shell};

const RECENT_RAW: usize = 120;
const RECENT_VISIBLE: usize = 20;
const RECENT_UNDER_HIDE: usize = 12; // leaves 8 visible
const FREQUENT_RAW: usize = 100;
const FREQUENT_PINNED: usize = 46;
const FREQUENT_VISIBLE_UNPINNED: usize = 4; // 46 + 4 = 50 visible
const FREQUENT_UNDER_HIDE_PINNED: usize = 10; // leaves 40 visible

const QUERY_TIMEOUT: Duration = Duration::from_secs(30);
const RECENT_NAMESPACE: &str = "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}";
const FREQUENT_NAMESPACE: &str = "shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Scenario {
    id: &'static str,
    kind: QuickAccess,
    expected_shell: usize,
    expected_raw: usize,
}

#[derive(Debug)]
struct Sample {
    scenario: &'static str,
    method: &'static str,
    iteration: usize,
    elapsed_ms: f64,
    observed_count: usize,
}

#[derive(Debug, Clone, Copy)]
struct Stats {
    count: usize,
    mean_ms: f64,
    median_ms: f64,
    p95_ms: f64,
    min_ms: f64,
    max_ms: f64,
    stdev_ms: f64,
}

impl Stats {
    fn from_values(values: &[f64]) -> Option<Self> {
        if values.is_empty() {
            return None;
        }
        let mut sorted = values.to_vec();
        sorted.sort_by(f64::total_cmp);
        let count = sorted.len();
        let mean = sorted.iter().sum::<f64>() / count as f64;
        let median = if count % 2 == 1 {
            sorted[count / 2]
        } else {
            (sorted[count / 2 - 1] + sorted[count / 2]) / 2.0
        };
        let rank = ((count as f64 * 0.95).ceil() as usize).clamp(1, count);
        let variance = if count > 1 {
            sorted
                .iter()
                .map(|value| {
                    let delta = value - mean;
                    delta * delta
                })
                .sum::<f64>()
                / (count - 1) as f64
        } else {
            0.0
        };
        Some(Self {
            count,
            mean_ms: mean,
            median_ms: median,
            p95_ms: sorted[rank - 1],
            min_ms: sorted[0],
            max_ms: sorted[count - 1],
            stdev_ms: variance.sqrt(),
        })
    }
}

#[test]
#[ignore = "writes real Quick Access state; run inside the Windows test VM with WINCENT_GET_ITEM_LATENCY=1"]
fn get_item_latency() {
    if env::var("WINCENT_GET_ITEM_LATENCY").as_deref() != Ok("1") {
        eprintln!(
            "get_item_latency skipped: set WINCENT_GET_ITEM_LATENCY=1 and run this test only \
             inside the dedicated Windows test VM"
        );
        return;
    }

    let out_path = PathBuf::from(
        env::var("WINCENT_GET_ITEM_LATENCY_OUT")
            .expect("WINCENT_GET_ITEM_LATENCY_OUT must point at the result JSON"),
    );
    let root = PathBuf::from(
        env::var("WINCENT_GET_ITEM_LATENCY_ROOT")
            .expect("WINCENT_GET_ITEM_LATENCY_ROOT must point at the fixture root"),
    );
    let iterations = parse_env_usize("WINCENT_GET_ITEM_LATENCY_ITERS", 20);
    let warmup = parse_env_usize("WINCENT_GET_ITEM_LATENCY_WARMUP", 2);
    assert!(iterations > 0, "iterations must be > 0");

    if root.exists() {
        fs::remove_dir_all(&root).expect("failed to clean fixture root");
    }
    fs::create_dir_all(&root).expect("failed to create fixture root");

    let manager = QuickAccessManager::builder()
        .timeout(Duration::from_secs(60))
        .build();

    println!(
        "get_item_latency: root={} iters={iterations} warmup={warmup}",
        root.display()
    );

    build_fixture(&manager, &root).unwrap_or_else(|error| panic!("fixture build failed: {error}"));
    validate_fixture(&manager).unwrap_or_else(|error| panic!("fixture validation failed: {error}"));

    start_explorer().unwrap_or_else(|error| panic!("Explorer start failed: {error}"));
    wait_for_shell(QuickAccess::RecentFiles, RECENT_VISIBLE)
        .unwrap_or_else(|error| panic!("Recent Files shell count mismatch: {error}"));
    wait_for_shell(
        QuickAccess::FrequentFolders,
        FREQUENT_PINNED + FREQUENT_VISIBLE_UNPINNED,
    )
    .unwrap_or_else(|error| panic!("Frequent Folders shell count mismatch: {error}"));
    check_frequent_distribution(&manager, FREQUENT_PINNED)
        .unwrap_or_else(|error| panic!("post-restart fixture validation failed: {error}"));

    let full_recent = Scenario {
        id: "recent_full",
        kind: QuickAccess::RecentFiles,
        expected_shell: RECENT_VISIBLE,
        expected_raw: RECENT_RAW,
    };
    let full_frequent = Scenario {
        id: "frequent_full",
        kind: QuickAccess::FrequentFolders,
        expected_shell: FREQUENT_PINNED + FREQUENT_VISIBLE_UNPINNED,
        expected_raw: FREQUENT_RAW,
    };
    let mut samples = Vec::new();
    measure_scenario(&full_recent, iterations, warmup, &mut samples)
        .unwrap_or_else(|error| panic!("{} measurement failed: {error}", full_recent.id));
    measure_scenario(&full_frequent, iterations, warmup, &mut samples)
        .unwrap_or_else(|error| panic!("{} measurement failed: {error}", full_frequent.id));

    reduce_fixture_for_under(&manager, &root)
        .unwrap_or_else(|error| panic!("fixture reduction failed: {error}"));
    start_explorer().unwrap_or_else(|error| panic!("second Explorer start failed: {error}"));
    wait_for_shell(QuickAccess::RecentFiles, RECENT_VISIBLE - RECENT_UNDER_HIDE)
        .unwrap_or_else(|error| panic!("reduced Recent Files shell count mismatch: {error}"));
    wait_for_shell(
        QuickAccess::FrequentFolders,
        FREQUENT_PINNED + FREQUENT_VISIBLE_UNPINNED - FREQUENT_UNDER_HIDE_PINNED,
    )
    .unwrap_or_else(|error| panic!("reduced Frequent Folders shell count mismatch: {error}"));
    check_frequent_distribution(&manager, FREQUENT_PINNED - FREQUENT_UNDER_HIDE_PINNED)
        .unwrap_or_else(|error| panic!("post-reduction fixture validation failed: {error}"));

    let under_recent = Scenario {
        id: "recent_under",
        kind: QuickAccess::RecentFiles,
        expected_shell: RECENT_VISIBLE - RECENT_UNDER_HIDE,
        expected_raw: RECENT_RAW,
    };
    let under_frequent = Scenario {
        id: "frequent_under",
        kind: QuickAccess::FrequentFolders,
        expected_shell: FREQUENT_PINNED + FREQUENT_VISIBLE_UNPINNED - FREQUENT_UNDER_HIDE_PINNED,
        expected_raw: FREQUENT_RAW,
    };
    measure_scenario(&under_recent, iterations, warmup, &mut samples)
        .unwrap_or_else(|error| panic!("{} measurement failed: {error}", under_recent.id));
    measure_scenario(&under_frequent, iterations, warmup, &mut samples)
        .unwrap_or_else(|error| panic!("{} measurement failed: {error}", under_frequent.id));

    let scenarios = [full_recent, full_frequent, under_recent, under_frequent];
    let stats = summarize(&samples, &scenarios);
    print_summary(&stats);
    write_result_json(&out_path, iterations, warmup, &scenarios, &samples, &stats)
        .expect("failed to write get_item_latency result JSON");
    write_timings_csv(&out_path.with_file_name("timings.csv"), &samples)
        .expect("failed to write get_item_latency timings CSV");
}

fn parse_env_usize(name: &str, default: usize) -> usize {
    match env::var(name) {
        Ok(value) => value
            .parse()
            .unwrap_or_else(|_| panic!("{name} must be an integer, got {value}")),
        Err(_) => default,
    }
}

fn recent_target(root: &Path, index: usize) -> PathBuf {
    root.join(format!("recent-{index:03}.txt"))
}

fn frequent_target(root: &Path, index: usize) -> PathBuf {
    root.join(format!("frequent-{index:03}"))
}

fn raw_entries(
    manager: &QuickAccessManager,
    kind: QuickAccess,
) -> Result<Vec<DestListEntry>, String> {
    match kind {
        QuickAccess::RecentFiles => manager
            .get_recent_files_metadata()
            .map_err(|error| error.to_string()),
        QuickAccess::FrequentFolders => manager
            .get_frequent_folders_metadata()
            .map_err(|error| error.to_string()),
        QuickAccess::All => Err("All is not a single list".to_string()),
        _ => Err("unsupported Quick Access category".to_string()),
    }
}

fn write_empty_cfb(kind: QuickAccess) -> Result<(), String> {
    let path = source_path(kind)?;
    let mut header = vec![0u8; 0x20];
    header[0..4].copy_from_slice(&4u32.to_le_bytes());
    // Entry count, pinned count, score counter, last entry id and action count
    // stay zero. A valid empty v4 header is enough for the direct writer.
    let cursor = Cursor::new(Vec::new());
    let mut compound = cfb::CompoundFile::create_with_version(cfb::Version::V3, cursor)
        .map_err(|error| format!("create empty CFB: {error}"))?;
    compound
        .create_stream("/DestList")
        .map_err(|error| format!("create DestList stream: {error}"))?
        .write_all(&header)
        .map_err(|error| format!("write DestList header: {error}"))?;
    compound
        .flush()
        .map_err(|error| format!("flush empty CFB: {error}"))?;
    let bytes = compound.into_inner().into_inner();
    fs::write(&path, bytes).map_err(|error| format!("write {}: {error}", path.display()))
}

fn build_fixture(manager: &QuickAccessManager, root: &Path) -> Result<(), String> {
    // Start from valid empty v4 files instead of rewriting the user's baseline
    // CFB: the baseline is not guaranteed to satisfy the direct writer's
    // contiguous-rank invariants.
    stop_explorer()?;
    write_empty_cfb(QuickAccess::RecentFiles)?;
    write_empty_cfb(QuickAccess::FrequentFolders)?;

    for index in 0..RECENT_RAW {
        let path = recent_target(root, index);
        fs::write(&path, b"wincent get_item latency fixture\n")
            .map_err(|error| format!("create {}: {error}", path.display()))?;
        add_entry_checked(manager, &path, QuickAccess::RecentFiles)?;
    }
    for index in 0..FREQUENT_RAW {
        let path = frequent_target(root, index);
        fs::create_dir_all(&path).map_err(|error| format!("create {}: {error}", path.display()))?;
        add_entry_checked(manager, &path, QuickAccess::FrequentFolders)?;
    }

    // 100 hidden Recent Files; 20 remain visible and form the 20-item window.
    for index in RECENT_VISIBLE..RECENT_RAW {
        hide_entry(
            manager,
            &recent_target(root, index),
            QuickAccess::RecentFiles,
        )?;
    }

    // 46 pinned, 4 visible unpinned, 50 hidden unpinned.
    for index in 0..FREQUENT_PINNED {
        pin_entry(manager, &frequent_target(root, index))?;
    }
    for index in (FREQUENT_PINNED + FREQUENT_VISIBLE_UNPINNED)..FREQUENT_RAW {
        hide_entry(
            manager,
            &frequent_target(root, index),
            QuickAccess::FrequentFolders,
        )?;
    }
    patch_frequent_access_counts()?;
    Ok(())
}

fn patch_frequent_access_counts() -> Result<(), String> {
    let path = source_path(QuickAccess::FrequentFolders)?;
    let parsed = parse_file(&path).map_err(|error| error.to_string())?;
    let mut targets: Vec<(usize, usize, u32)> = Vec::new();
    for entry in parsed.cfb_entries() {
        let name = Path::new(entry.path())
            .file_name()
            .and_then(|value| value.to_str())
            .ok_or_else(|| format!("unexpected fixture path {}", entry.path()))?;
        let index: usize = name
            .strip_prefix("frequent-")
            .ok_or_else(|| format!("unexpected fixture name {name}"))?
            .parse()
            .map_err(|error| format!("parse fixture index {name}: {error}"))?;
        let count = if index < FREQUENT_PINNED {
            0
        } else if index < FREQUENT_PINNED + FREQUENT_VISIBLE_UNPINNED {
            3
        } else if index < 70 {
            1
        } else if index < 90 {
            5
        } else {
            3
        };
        targets.push((entry.entry_offset(), entry.entry_len(), count));
    }

    let original = fs::read(&path).map_err(|error| format!("read {}: {error}", path.display()))?;
    let mut compound = cfb::CompoundFile::open(Cursor::new(original))
        .map_err(|error| format!("open frequent CFB: {error}"))?;
    {
        let mut stream = compound
            .open_stream("/DestList")
            .map_err(|error| format!("open DestList: {error}"))?;
        for (offset, length, count) in targets {
            let mut entry = vec![0u8; length];
            stream
                .seek(SeekFrom::Start(offset as u64))
                .and_then(|_| stream.read_exact(&mut entry))
                .map_err(|error| format!("read DestList entry at {offset:#x}: {error}"))?;
            let count_offset = 0x74;
            entry[count_offset..count_offset + 4].copy_from_slice(&count.to_le_bytes());
            let checksum = destlist_entry_crc64(&entry);
            entry[0..8].copy_from_slice(&checksum.to_le_bytes());
            stream
                .seek(SeekFrom::Start(offset as u64))
                .and_then(|_| stream.write_all(&entry))
                .map_err(|error| format!("write DestList entry at {offset:#x}: {error}"))?;
        }
    }
    compound
        .flush()
        .map_err(|error| format!("flush frequent CFB: {error}"))?;
    let bytes = compound.into_inner().into_inner();
    fs::write(&path, bytes).map_err(|error| format!("write {}: {error}", path.display()))?;

    let parsed = parse_file(&path).map_err(|error| error.to_string())?;
    if parsed
        .cfb_entries()
        .iter()
        .any(|entry| !entry.checksum_valid())
    {
        return Err("patched frequent CFB has invalid checksums".to_string());
    }
    Ok(())
}

fn destlist_entry_crc64(entry: &[u8]) -> u64 {
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

fn find_entry(
    manager: &QuickAccessManager,
    kind: QuickAccess,
    path: &Path,
) -> Result<Option<DestListEntry>, String> {
    let wanted = path.to_string_lossy();
    Ok(raw_entries(manager, kind)?
        .into_iter()
        .find(|entry| entry.path().eq_ignore_ascii_case(wanted.as_ref())))
}

fn add_entry_checked(
    manager: &QuickAccessManager,
    path: &Path,
    kind: QuickAccess,
) -> Result<(), String> {
    match manager.add_item_direct(path, kind) {
        Ok(()) => Ok(()),
        Err(error) => {
            // With Explorer stopped the CFB write can succeed while the
            // post-mutation refresh reports a failure. Trust the raw entry.
            if find_entry(manager, kind, path)?.is_some() {
                Ok(())
            } else {
                Err(format!("add {}: {error}", path.display()))
            }
        }
    }
}

fn hide_entry(manager: &QuickAccessManager, path: &Path, kind: QuickAccess) -> Result<(), String> {
    let update = DestListItemUpdate::new().hide();
    match manager.update_item_direct(path, kind, update) {
        Ok(()) => Ok(()),
        Err(error) => {
            let hidden = find_entry(manager, kind, path)?
                .is_some_and(|entry| entry.reserved_78().unwrap_or_default() != 0);
            if hidden {
                Ok(())
            } else {
                Err(format!("hide {}: {error}", path.display()))
            }
        }
    }
}

fn pin_entry(manager: &QuickAccessManager, path: &Path) -> Result<(), String> {
    let update = DestListItemUpdate::new().pin();
    match manager.update_item_direct(path, QuickAccess::FrequentFolders, update) {
        Ok(()) => Ok(()),
        Err(error) => {
            let pinned = find_entry(manager, QuickAccess::FrequentFolders, path)?
                .is_some_and(|entry| entry.pin_status() >= 0);
            if pinned {
                Ok(())
            } else {
                Err(format!("pin {}: {error}", path.display()))
            }
        }
    }
}

fn unpin_and_hide_entry(manager: &QuickAccessManager, path: &Path) -> Result<(), String> {
    let update = DestListItemUpdate::new().unpin().hide();
    match manager.update_item_direct(path, QuickAccess::FrequentFolders, update) {
        Ok(()) => Ok(()),
        Err(error) => {
            let hidden =
                find_entry(manager, QuickAccess::FrequentFolders, path)?.is_some_and(|entry| {
                    entry.pin_status() < 0 && entry.reserved_78().unwrap_or_default() != 0
                });
            if hidden {
                Ok(())
            } else {
                Err(format!("unpin+hide {}: {error}", path.display()))
            }
        }
    }
}

fn validate_fixture(manager: &QuickAccessManager) -> Result<(), String> {
    let recent = raw_entries(manager, QuickAccess::RecentFiles)?;
    if recent.len() != RECENT_RAW {
        return Err(format!("recent raw count {} != {RECENT_RAW}", recent.len()));
    }
    let recent_hidden = recent
        .iter()
        .filter(|entry| entry.reserved_78().unwrap_or_default() != 0)
        .count();
    if recent_hidden != RECENT_RAW - RECENT_VISIBLE {
        return Err(format!(
            "recent hidden count {recent_hidden} != {}",
            RECENT_RAW - RECENT_VISIBLE
        ));
    }

    let frequent = raw_entries(manager, QuickAccess::FrequentFolders)?;
    if frequent.len() != FREQUENT_RAW {
        return Err(format!(
            "frequent raw count {} != {FREQUENT_RAW}",
            frequent.len()
        ));
    }
    let pinned = frequent
        .iter()
        .filter(|entry| entry.pin_status() >= 0)
        .count();
    if pinned != FREQUENT_PINNED {
        return Err(format!("pinned count {pinned} != {FREQUENT_PINNED}"));
    }
    let counts: BTreeMap<u32, usize> = frequent.iter().fold(BTreeMap::new(), |mut map, entry| {
        *map.entry(entry.access_count()).or_default() += 1;
        map
    });
    println!("fixture access_count distribution: {counts:?}");
    for expected_count in [0, 1, 5] {
        if !counts.contains_key(&expected_count) {
            return Err(format!(
                "fixture access_count distribution {counts:?} lacks {expected_count}"
            ));
        }
    }
    if frequent
        .iter()
        .any(|entry| entry.pin_status() >= 0 && entry.access_count() != 0)
    {
        return Err("pinned fixture entries must have access_count 0".to_string());
    }

    let recent_parsed = parse_file(recent_files_dest_path().map_err(|error| error.to_string())?)
        .map_err(|error| error.to_string())?;
    let recent_rule = recent_parsed.rule_entries().len();
    if recent_rule != RECENT_VISIBLE {
        return Err(format!(
            "recent rule count {recent_rule} != {RECENT_VISIBLE}"
        ));
    }
    let frequent_parsed =
        parse_file(frequent_folders_dest_path().map_err(|error| error.to_string())?)
            .map_err(|error| error.to_string())?;
    let frequent_rule = frequent_parsed.rule_entries().len();
    let expected_visible = FREQUENT_PINNED + FREQUENT_VISIBLE_UNPINNED;
    if frequent_rule != expected_visible {
        return Err(format!(
            "frequent rule count {frequent_rule} != {expected_visible}"
        ));
    }
    println!(
        "fixture: recent raw={} rule={} hidden={}; frequent raw={} rule={} pinned={} counts={:?}",
        recent.len(),
        recent_rule,
        recent_hidden,
        frequent.len(),
        frequent_rule,
        pinned,
        counts
    );
    Ok(())
}

fn check_frequent_distribution(
    manager: &QuickAccessManager,
    expected_pinned: usize,
) -> Result<(), String> {
    let frequent = raw_entries(manager, QuickAccess::FrequentFolders)?;
    if frequent.len() != FREQUENT_RAW {
        return Err(format!(
            "frequent raw count {} != {FREQUENT_RAW}",
            frequent.len()
        ));
    }
    let pinned = frequent
        .iter()
        .filter(|entry| entry.pin_status() >= 0)
        .count();
    if pinned != expected_pinned {
        return Err(format!("pinned count {pinned} != {expected_pinned}"));
    }
    let counts: BTreeMap<u32, usize> = frequent.iter().fold(BTreeMap::new(), |mut map, entry| {
        *map.entry(entry.access_count()).or_default() += 1;
        map
    });
    println!("post-restart access_count distribution: {counts:?}");
    for expected_count in [0, 1, 5] {
        if !counts.contains_key(&expected_count) {
            return Err(format!(
                "post-restart access_count distribution {counts:?} lacks {expected_count}"
            ));
        }
    }
    Ok(())
}

fn reduce_fixture_for_under(manager: &QuickAccessManager, root: &Path) -> Result<(), String> {
    stop_explorer()?;
    for index in 0..RECENT_UNDER_HIDE {
        hide_entry(
            manager,
            &recent_target(root, index),
            QuickAccess::RecentFiles,
        )?;
    }
    for index in 0..FREQUENT_UNDER_HIDE_PINNED {
        unpin_and_hide_entry(manager, &frequent_target(root, index))?;
    }
    Ok(())
}

fn stop_explorer() -> Result<(), String> {
    let windir = env::var("WINDIR").unwrap_or_else(|_| "C:\\Windows".to_string());
    let taskkill = format!("{windir}\\System32\\taskkill.exe");
    let _ = Command::new(&taskkill)
        .args(["/F", "/IM", "explorer.exe"])
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status();
    std::thread::sleep(Duration::from_millis(1500));
    Ok(())
}

fn start_explorer() -> Result<(), String> {
    let windir = env::var("WINDIR").unwrap_or_else(|_| "C:\\Windows".to_string());
    // Explorer is long-lived: never let it inherit the benchmark log handles,
    // or the caller's redirected stdout/stderr can stay open forever.
    Command::new(format!("{windir}\\explorer.exe"))
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .map_err(|error| format!("start explorer: {error}"))?;
    std::thread::sleep(Duration::from_millis(2500));
    Ok(())
}

fn wait_for_shell(kind: QuickAccess, expected: usize) -> Result<(), String> {
    let mut last = 0;
    let mut last_error = None;
    for _ in 0..40 {
        match shell_query(kind) {
            Ok(items) => {
                last = items.len();
                if last == expected {
                    println!("shell count for {kind:?}: {last}");
                    return Ok(());
                }
            }
            Err(error) => last_error = Some(error),
        }
        std::thread::sleep(Duration::from_millis(500));
    }
    Err(format!(
        "expected {expected} shell items for {kind:?}, last={last}, error={last_error:?}"
    ))
}

fn measure_scenario(
    scenario: &Scenario,
    iterations: usize,
    warmup: usize,
    samples: &mut Vec<Sample>,
) -> Result<(), String> {
    let source_path = source_path(scenario.kind)?;
    println!(
        "scenario {}: expected shell={} raw={}",
        scenario.id, scenario.expected_shell, scenario.expected_raw
    );

    // Interleave the two methods so neither always runs in the same cache state.
    for index in 0..(warmup + iterations) {
        let timed = index >= warmup;
        let iteration = index + 1 - warmup;
        if index % 2 == 0 {
            let (elapsed_ms, observed_count) = shell_sample(scenario)?;
            if timed {
                samples.push(Sample {
                    scenario: scenario.id,
                    method: "shell_query",
                    iteration,
                    elapsed_ms,
                    observed_count,
                });
            }
            let (elapsed_ms, observed_count) = cfb_sample(scenario, &source_path)?;
            if timed {
                samples.push(Sample {
                    scenario: scenario.id,
                    method: "cfb_parse",
                    iteration,
                    elapsed_ms,
                    observed_count,
                });
            }
        } else {
            let (elapsed_ms, observed_count) = cfb_sample(scenario, &source_path)?;
            if timed {
                samples.push(Sample {
                    scenario: scenario.id,
                    method: "cfb_parse",
                    iteration,
                    elapsed_ms,
                    observed_count,
                });
            }
            let (elapsed_ms, observed_count) = shell_sample(scenario)?;
            if timed {
                samples.push(Sample {
                    scenario: scenario.id,
                    method: "shell_query",
                    iteration,
                    elapsed_ms,
                    observed_count,
                });
            }
        }
    }
    Ok(())
}

fn shell_sample(scenario: &Scenario) -> Result<(f64, usize), String> {
    let started = Instant::now();
    let items = shell_query(scenario.kind)?;
    let elapsed_ms = started.elapsed().as_secs_f64() * 1000.0;
    if items.len() != scenario.expected_shell {
        return Err(format!(
            "{} shell count {} != {}",
            scenario.id,
            items.len(),
            scenario.expected_shell
        ));
    }
    Ok((elapsed_ms, items.len()))
}

fn cfb_sample(scenario: &Scenario, path: &Path) -> Result<(f64, usize), String> {
    let started = Instant::now();
    let parsed = parse_file(path).map_err(|error| error.to_string())?;
    let count = parsed.cfb_entries().len();
    let elapsed_ms = started.elapsed().as_secs_f64() * 1000.0;
    if count != scenario.expected_raw {
        return Err(format!(
            "{} raw count {} != {}",
            scenario.id, count, scenario.expected_raw
        ));
    }
    Ok((elapsed_ms, count))
}

fn source_path(kind: QuickAccess) -> Result<PathBuf, String> {
    match kind {
        QuickAccess::RecentFiles => recent_files_dest_path().map_err(|error| error.to_string()),
        QuickAccess::FrequentFolders => {
            frequent_folders_dest_path().map_err(|error| error.to_string())
        }
        QuickAccess::All => Err("All has no single backing file".to_string()),
        _ => Err("unsupported Quick Access category".to_string()),
    }
}

fn shell_query(kind: QuickAccess) -> Result<Vec<String>, String> {
    let (namespace, files_only) = match kind {
        QuickAccess::RecentFiles => (RECENT_NAMESPACE, true),
        QuickAccess::FrequentFolders => (FREQUENT_NAMESPACE, false),
        QuickAccess::All => return Err("All is not supported by this benchmark".to_string()),
        _ => return Err("unsupported Quick Access category".to_string()),
    };
    let (tx, rx) = mpsc::channel();
    std::thread::spawn(move || {
        let _ = tx.send(shell_query_sta(namespace, files_only));
    });
    match rx.recv_timeout(QUERY_TIMEOUT) {
        Ok(result) => result,
        Err(mpsc::RecvTimeoutError::Timeout) => Err(format!(
            "Shell query timed out after {}s",
            QUERY_TIMEOUT.as_secs()
        )),
        Err(mpsc::RecvTimeoutError::Disconnected) => {
            Err("Shell query worker disconnected".to_string())
        }
    }
}

fn shell_query_sta(namespace: &str, files_only: bool) -> Result<Vec<String>, String> {
    // SAFETY: Called on the dedicated worker thread; a successful
    // CoInitializeEx is balanced by CoUninitialize before the thread exits.
    let initialized = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).is_ok() };
    if !initialized {
        return Err("CoInitializeEx failed".to_string());
    }
    let result = (|| -> Result<Vec<String>, String> {
        // SAFETY: COM is initialized on this thread; the created Shell
        // interface and namespace collection are used only on this thread.
        let shell: IShellDispatch =
            unsafe { CoCreateInstance(&Shell, None, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER) }
                .map_err(|error| format!("CoCreateInstance Shell: {error}"))?;
        let variant = VARIANT::from(namespace);
        // SAFETY: `shell` and the namespace VARIANT are live for the call.
        let folder: Folder = unsafe {
            shell
                .NameSpace(&variant)
                .map_err(|error| format!("open {namespace}: {error}"))?
        };
        // SAFETY: `folder` is a live Shell namespace on this STA thread.
        let items: FolderItems = unsafe {
            folder
                .Items()
                .map_err(|error| format!("enumerate {namespace}: {error}"))?
        };
        // SAFETY: `items` is a live collection on this thread.
        let count = unsafe {
            items
                .Count()
                .map_err(|error| format!("item count {namespace}: {error}"))?
        };
        let mut paths = Vec::with_capacity(count.max(0) as usize);
        for index in 0..count {
            let index_variant = VARIANT::from(index);
            // SAFETY: index is within Count(); the VARIANT lives for the call.
            let item = match unsafe { items.Item(&index_variant) } {
                Ok(item) => item,
                Err(_) => continue,
            };
            // SAFETY: `item` is a live FolderItem on this thread.
            let is_folder = unsafe { item.IsFolder().map(bool::from).unwrap_or(false) };
            if files_only == is_folder {
                continue;
            }
            // SAFETY: `item` is live; Path returns an owned BSTR.
            if let Ok(path) = unsafe { item.Path() } {
                let value = path.to_string();
                if !value.trim().is_empty() {
                    paths.push(value);
                }
            }
        }
        Ok(paths)
    })();
    // SAFETY: balances the successful CoInitializeEx above on the same thread.
    unsafe { CoUninitialize() };
    result
}

fn summarize(
    samples: &[Sample],
    scenarios: &[Scenario],
) -> BTreeMap<&'static str, BTreeMap<&'static str, Stats>> {
    let mut map = BTreeMap::new();
    for scenario in scenarios {
        let mut by_method = BTreeMap::new();
        for method in ["shell_query", "cfb_parse"] {
            let values: Vec<f64> = samples
                .iter()
                .filter(|sample| sample.scenario == scenario.id && sample.method == method)
                .map(|sample| sample.elapsed_ms)
                .collect();
            if let Some(stats) = Stats::from_values(&values) {
                by_method.insert(method, stats);
            }
        }
        map.insert(scenario.id, by_method);
    }
    map
}

fn print_summary(stats: &BTreeMap<&'static str, BTreeMap<&'static str, Stats>>) {
    println!(
        "{:<16} {:<12} {:>4} {:>11} {:>11} {:>11} {:>11}",
        "scenario", "method", "n", "mean_ms", "median_ms", "p95_ms", "max_ms"
    );
    for (scenario, methods) in stats {
        for (method, stats) in methods {
            println!(
                "{scenario:<16} {method:<12} {:>4} {:>11.3} {:>11.3} {:>11.3} {:>11.3}",
                stats.count, stats.mean_ms, stats.median_ms, stats.p95_ms, stats.max_ms
            );
        }
    }
}

fn write_result_json(
    path: &Path,
    iterations: usize,
    warmup: usize,
    scenarios: &[Scenario],
    samples: &[Sample],
    stats: &BTreeMap<&'static str, BTreeMap<&'static str, Stats>>,
) -> std::io::Result<()> {
    let mut json = String::new();
    json.push_str("{\n");
    json.push_str("  \"test\": \"get_item_latency\",\n");
    json.push_str(&format!("  \"iterations\": {iterations},\n"));
    json.push_str(&format!("  \"warmup\": {warmup},\n"));
    json.push_str(&format!(
        "  \"fixture\": {{\"recent_raw\": {RECENT_RAW}, \"recent_visible\": {RECENT_VISIBLE}, \
         \"frequent_raw\": {FREQUENT_RAW}, \"frequent_pinned\": {FREQUENT_PINNED}, \
         \"frequent_visible_unpinned\": {FREQUENT_VISIBLE_UNPINNED}}},\n"
    ));
    json.push_str("  \"scenarios\": {\n");
    let entries: Vec<String> = scenarios
        .iter()
        .filter_map(|scenario| {
            let methods = stats.get(scenario.id)?;
            let method_entries: Vec<String> = methods
                .iter()
                .map(|(method, stats)| {
                    format!(
                        "      {}: {{\"count\": {}, \"mean_ms\": {:.6}, \"median_ms\": {:.6}, \
                         \"p95_ms\": {:.6}, \"min_ms\": {:.6}, \"max_ms\": {:.6}, \
                         \"stdev_ms\": {:.6}}}",
                        json_escape(method),
                        stats.count,
                        stats.mean_ms,
                        stats.median_ms,
                        stats.p95_ms,
                        stats.min_ms,
                        stats.max_ms,
                        stats.stdev_ms
                    )
                })
                .collect();
            Some(format!(
                "    {}: {{\"kind\": {}, \"expected_shell\": {}, \"expected_raw\": {},\n{} \n    }}",
                json_escape(scenario.id),
                json_escape(if scenario.kind == QuickAccess::RecentFiles {
                    "recent"
                } else {
                    "frequent"
                }),
                scenario.expected_shell,
                scenario.expected_raw,
                method_entries.join(",\n")
            ))
        })
        .collect();
    json.push_str(&entries.join(",\n"));
    json.push_str("\n  },\n  \"samples\": [\n");
    let rows: Vec<String> = samples
        .iter()
        .map(|sample| {
            format!(
                "    {{\"scenario\": {}, \"method\": {}, \"iteration\": {}, \"elapsed_ms\": {:.6}, \
                 \"observed_count\": {}}}",
                json_escape(sample.scenario),
                json_escape(sample.method),
                sample.iteration,
                sample.elapsed_ms,
                sample.observed_count
            )
        })
        .collect();
    json.push_str(&rows.join(",\n"));
    json.push_str("\n  ]\n}\n");
    fs::write(path, json)
}

fn write_timings_csv(path: &Path, samples: &[Sample]) -> std::io::Result<()> {
    let mut csv = String::from("scenario,method,iteration,elapsed_ms,observed_count\r\n");
    for sample in samples {
        csv.push_str(&format!(
            "{},{},{},{:.6},{}\r\n",
            sample.scenario,
            sample.method,
            sample.iteration,
            sample.elapsed_ms,
            sample.observed_count
        ));
    }
    fs::write(path, csv)
}

fn json_escape(value: &str) -> String {
    let mut out = String::with_capacity(value.len() + 2);
    out.push('"');
    for character in value.chars() {
        match character {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            character if (character as u32) < 0x20 => {
                out.push_str(&format!("\\u{:04x}", character as u32));
            }
            character => out.push(character),
        }
    }
    out.push('"');
    out
}
