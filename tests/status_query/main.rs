//! VM-only validation for Quick Access path status queries.
//!
//! Validates the public status API against the live Quick Access state of the
//! dedicated Windows test VM:
//!
//! - the Recent Files and Frequent Folders positions reported by
//!   `QuickAccessManager` must match the Shell namespace enumation order;
//! - every raw entry that Explorer filters out must report no position;
//! - Windows path variants must resolve to the same position;
//! - a Shell mutation must be observable through the status API (add/pin and
//!   remove/unpin round trips);
//! - parse, metadata and query costs are measured for the caching analysis.
//!
//! The test mutates Quick Access and is therefore ignored by default and gated
//! behind `WINCENT_STATUS_QUERY=1`. The guest orchestrator backs up and restores
//! both backing files and the Recent `.lnk` set.

#![cfg(windows)]

use std::collections::BTreeMap;
use std::env;
use std::fs;
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant};

use wincent::destlist::{
    frequent_folders_dest_path, parse_file, recent_files_dest_path,
    DEFAULT_RECENT_FILES_RESULT_LIMIT,
};
use wincent::prelude::*;

const QUERY_TIMEOUT: Duration = Duration::from_secs(30);
const PARSE_ITERATIONS: usize = 100;
const METADATA_ITERATIONS: usize = 200;
const STATUS_ITERATIONS: usize = 50;
const MUTATION_TIMEOUT: Duration = Duration::from_secs(20);
const FIXTURE_TIMEOUT: Duration = Duration::from_secs(120);
const FIXTURE_RECENT_FILES: usize = 25;
const FIXTURE_UNPINNED_FOLDERS: usize = 2;
const MUTATION_POLL: Duration = Duration::from_millis(250);

#[derive(Debug, Clone, Copy)]
struct Stats {
    count: usize,
    mean_ms: f64,
    median_ms: f64,
    p95_ms: f64,
    min_ms: f64,
    max_ms: f64,
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
        let p95_index = ((count as f64) * 0.95).ceil() as usize;
        let p95 = sorted[(p95_index.max(1) - 1).min(count - 1)];

        Some(Self {
            count,
            mean_ms: mean,
            median_ms: median,
            p95_ms: p95,
            min_ms: sorted[0],
            max_ms: sorted[count - 1],
        })
    }
}

#[derive(Default)]
struct Checks {
    entries: Vec<(String, bool, String)>,
    failures: usize,
}

impl Checks {
    fn check(&mut self, id: impl Into<String>, ok: bool, detail: impl Into<String>) {
        if !ok {
            self.failures += 1;
        }
        self.entries.push((id.into(), ok, detail.into()));
    }
}

fn normalize(path: &str) -> String {
    let mut value = path.to_lowercase().replace('/', "\\");
    if value.len() > 3 && value.ends_with('\\') {
        value.pop();
    }
    value
}

fn same_path(left: &str, right: &str) -> bool {
    normalize(left) == normalize(right)
}

fn measure<F: FnMut()>(iterations: usize, mut body: F) -> Option<Stats> {
    let mut samples = Vec::with_capacity(iterations);
    for _ in 0..iterations {
        let start = Instant::now();
        body();
        samples.push(start.elapsed().as_secs_f64() * 1000.0);
    }
    Stats::from_values(&samples)
}

fn wait_until<F: FnMut() -> bool>(timeout: Duration, mut predicate: F) -> bool {
    let deadline = Instant::now() + timeout;
    loop {
        if predicate() {
            return true;
        }
        if Instant::now() >= deadline {
            return false;
        }
        std::thread::sleep(MUTATION_POLL);
    }
}

fn json_escape(value: &str) -> String {
    let mut out = String::with_capacity(value.len() + 2);
    for character in value.chars() {
        match character {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            other => out.push(other),
        }
    }
    out
}

fn build_recent_fixture(manager: &QuickAccessManager, root: &Path, count: usize) -> Vec<PathBuf> {
    let directory = root.join("recent-fixture");
    let _ = fs::create_dir_all(&directory);
    let mut added = Vec::new();
    for index in 0..count {
        let path = directory.join(format!("fixture-{index:02}.txt"));
        if fs::write(&path, b"wincent status fixture").is_err() {
            continue;
        }
        if manager
            .add_item(&path, QuickAccess::RecentFiles, AddOptions::new())
            .is_ok()
        {
            added.push(path);
        }
    }
    added
}

fn build_frequent_fixture(manager: &QuickAccessManager, root: &Path, count: usize) -> Vec<PathBuf> {
    let mut added = Vec::new();
    for index in 0..count {
        let directory = root.join(format!("frequent-fixture-{index:02}"));
        if fs::create_dir_all(&directory).is_err() {
            continue;
        }
        if manager
            .add_item_direct(&directory, QuickAccess::FrequentFolders)
            .is_ok()
        {
            added.push(directory);
        }
    }
    added
}

#[test]
#[ignore = "VM-only: validates status queries against live Quick Access state"]
fn status_query() {
    if env::var("WINCENT_STATUS_QUERY").as_deref() != Ok("1") {
        eprintln!("WINCENT_STATUS_QUERY is not set; skipping");
        return;
    }

    let run_id = env::var("WINCENT_STATUS_QUERY_RUN_ID").unwrap_or_else(|_| "local".to_string());
    let root = PathBuf::from(
        env::var("WINCENT_STATUS_QUERY_ROOT")
            .unwrap_or_else(|_| "C:\\WincentTrack\\status-targets".to_string()),
    );
    let out_path = PathBuf::from(
        env::var("WINCENT_STATUS_QUERY_OUT")
            .unwrap_or_else(|_| "C:\\WincentTrack\\status-result.json".to_string()),
    );
    let rows_path = PathBuf::from(
        env::var("WINCENT_STATUS_QUERY_ROWS")
            .unwrap_or_else(|_| "C:\\WincentTrack\\status-timings.csv".to_string()),
    );
    let mutate = env::var("WINCENT_STATUS_QUERY_MUTATE").as_deref() == Ok("1");

    fs::create_dir_all(&root).expect("create target root");

    let manager = QuickAccessManager::builder()
        .try_timeout(QUERY_TIMEOUT)
        .expect("timeout")
        .build();

    let mut checks = Checks::default();
    let mut timings: BTreeMap<String, Stats> = BTreeMap::new();
    let mut summary: BTreeMap<String, String> = BTreeMap::new();
    let mut mutation_log: Vec<String> = Vec::new();

    let recent_path = recent_files_dest_path().expect("recent dest path");
    let frequent_path = frequent_folders_dest_path().expect("frequent dest path");

    // ---- fixture build ---------------------------------------------------
    // The stock 22H2 profile can have an empty Recent list and no unpinned
    // Frequent entry, which would make the positional checks vacuous. Build a
    // deterministic population through the public API so the result window and
    // the pinned/unpinned mix are actually exercised. The guest orchestrator
    // restores both backing files afterwards.
    if mutate {
        let recent_fixture = build_recent_fixture(&manager, &root, FIXTURE_RECENT_FILES);
        mutation_log.push(format!(
            "recent_fixture_added={} of {}",
            recent_fixture.len(),
            FIXTURE_RECENT_FILES
        ));
        let recent_ready = wait_until(FIXTURE_TIMEOUT, || {
            parse_file(&recent_path)
                .map(|parsed| parsed.cfb_entries().len() >= recent_fixture.len())
                .unwrap_or(false)
        });
        mutation_log.push(format!("recent_fixture_persisted={recent_ready}"));

        let frequent_fixture = build_frequent_fixture(&manager, &root, FIXTURE_UNPINNED_FOLDERS);
        mutation_log.push(format!(
            "frequent_fixture_added={} of {}",
            frequent_fixture.len(),
            FIXTURE_UNPINNED_FOLDERS
        ));
        let frequent_ready = wait_until(FIXTURE_TIMEOUT, || {
            parse_file(&frequent_path)
                .map(|parsed| {
                    parsed
                        .cfb_entries()
                        .iter()
                        .filter(|entry| !entry.is_pinned())
                        .count()
                        >= FIXTURE_UNPINNED_FOLDERS
                })
                .unwrap_or(false)
        });
        mutation_log.push(format!("frequent_fixture_persisted={frequent_ready}"));
    }

    // ---- baseline reads -------------------------------------------------
    let recent = parse_file(&recent_path).expect("parse recent backing file");
    let frequent = parse_file(&frequent_path).expect("parse frequent backing file");

    let recent_visible: Vec<String> = recent
        .shell_entries(DEFAULT_RECENT_FILES_RESULT_LIMIT)
        .iter()
        .map(|entry| entry.path().to_string())
        .collect();
    let recent_raw: Vec<String> = recent
        .cfb_entries()
        .iter()
        .map(|entry| entry.path().to_string())
        .collect();
    let frequent_visible: Vec<(String, bool)> = frequent
        .rule_entries()
        .iter()
        .map(|entry| (entry.path().to_string(), entry.is_pinned()))
        .collect();
    let frequent_raw: Vec<String> = frequent
        .cfb_entries()
        .iter()
        .map(|entry| entry.path().to_string())
        .collect();

    summary.insert("recent_raw".into(), recent_raw.len().to_string());
    summary.insert(
        "recent_rule_visible".into(),
        recent_visible.len().to_string(),
    );
    summary.insert("frequent_raw".into(), frequent_raw.len().to_string());
    summary.insert(
        "frequent_rule_visible".into(),
        frequent_visible.len().to_string(),
    );

    // ---- Shell oracle ---------------------------------------------------
    let shell_recent = manager
        .get_items(QuickAccess::RecentFiles)
        .expect("shell recent query");
    let shell_frequent = manager
        .get_items(QuickAccess::FrequentFolders)
        .expect("shell frequent query");
    summary.insert("shell_recent".into(), shell_recent.len().to_string());
    summary.insert("shell_frequent".into(), shell_frequent.len().to_string());

    checks.check(
        "shell_recent_matches_rule_window",
        shell_recent.len() == recent_visible.len()
            && shell_recent
                .iter()
                .zip(&recent_visible)
                .all(|(left, right)| same_path(left, right)),
        format!(
            "shell={:?} rule={:?}",
            shell_recent
                .iter()
                .map(|value| value.as_str())
                .collect::<Vec<_>>(),
            recent_visible
        ),
    );
    checks.check(
        "shell_frequent_matches_rule_set",
        shell_frequent.len() == frequent_visible.len()
            && shell_frequent
                .iter()
                .zip(&frequent_visible)
                .all(|(left, right)| same_path(left, &right.0)),
        format!(
            "shell={:?} rule={:?}",
            shell_frequent
                .iter()
                .map(|value| value.as_str())
                .collect::<Vec<_>>(),
            frequent_visible
                .iter()
                .map(|value| value.0.as_str())
                .collect::<Vec<_>>()
        ),
    );

    // ---- Recent positions ----------------------------------------------
    for (expected, path) in recent_visible.iter().enumerate() {
        let observed = manager.recent_file_position(path).expect("recent status");
        checks.check(
            format!("recent_position_{expected}"),
            observed == Some(expected),
            format!("path={path} expected={expected} observed={observed:?}"),
        );

        let status = manager.quick_access_status(path).expect("aggregate status");
        checks.check(
            format!("recent_aggregate_{expected}"),
            status.recent_file() == Some(expected) && status.frequent_folder().is_none(),
            format!("path={path} status={status:?}"),
        );
    }

    let mut recent_filtered = 0usize;
    for path in &recent_raw {
        if recent_visible
            .iter()
            .any(|visible| same_path(visible, path))
        {
            continue;
        }
        let observed = manager.recent_file_position(path).expect("recent status");
        checks.check(
            format!("recent_outside_window_{recent_filtered}"),
            observed.is_none(),
            format!("path={path} observed={observed:?}"),
        );
        recent_filtered += 1;
    }
    summary.insert("recent_filtered".into(), recent_filtered.to_string());

    // ---- Frequent positions --------------------------------------------
    for (expected, (path, pinned)) in frequent_visible.iter().enumerate() {
        let observed = manager.frequent_folder_slot(path).expect("frequent status");
        let matches = observed
            .as_ref()
            .is_some_and(|slot| slot.position() == expected && slot.is_pinned() == *pinned);
        checks.check(
            format!("frequent_slot_{expected}"),
            matches,
            format!("path={path} expected=({expected},{pinned}) observed={observed:?}"),
        );

        let status = manager.quick_access_status(path).expect("aggregate status");
        let aggregate_matches = status
            .frequent_folder()
            .is_some_and(|slot| slot.position() == expected && slot.is_pinned() == *pinned);
        checks.check(
            format!("frequent_aggregate_{expected}"),
            aggregate_matches,
            format!("path={path} status={status:?}"),
        );
    }

    let mut frequent_filtered = 0usize;
    for path in &frequent_raw {
        if frequent_visible
            .iter()
            .any(|visible| same_path(&visible.0, path))
        {
            continue;
        }
        let observed = manager.frequent_folder_slot(path).expect("frequent status");
        checks.check(
            format!("frequent_outside_slots_{frequent_filtered}"),
            observed.is_none(),
            format!("path={path} observed={observed:?}"),
        );
        frequent_filtered += 1;
    }
    summary.insert("frequent_filtered".into(), frequent_filtered.to_string());

    // ---- Path variants ---------------------------------------------------
    if let Some(first) = recent_visible.first() {
        let forward = first.replace('\\', "/");
        let lower = forward.to_lowercase();
        let trailing = format!("{first}\\");
        checks.check(
            "recent_variant_forward_slashes",
            manager.recent_file_position(&forward).expect("variant") == Some(0),
            format!("path={forward}"),
        );
        checks.check(
            "recent_variant_lowercase",
            manager.recent_file_position(&lower).expect("variant") == Some(0),
            format!("path={lower}"),
        );
        checks.check(
            "recent_variant_trailing_slash",
            manager.recent_file_position(&trailing).expect("variant") == Some(0),
            format!("path={trailing}"),
        );
    }

    // ---- Absent paths ----------------------------------------------------
    let absent_path = root.join("definitely-not-in-quick-access.txt");
    let absent = absent_path.to_string_lossy().to_string();
    let absent_status = manager.quick_access_status(&absent).expect("absent status");
    checks.check(
        "absent_path_not_visible",
        !absent_status.is_visible() && absent_status.lists().is_none(),
        format!("path={absent} status={absent_status:?}"),
    );
    checks.check(
        "absent_path_recent_none",
        manager
            .recent_file_position(&absent)
            .expect("absent")
            .is_none(),
        format!("path={absent}"),
    );
    checks.check(
        "absent_path_frequent_none",
        manager
            .frequent_folder_slot(&absent)
            .expect("absent")
            .is_none(),
        format!("path={absent}"),
    );

    // ---- Empty path errors ----------------------------------------------
    checks.check(
        "empty_path_rejected",
        matches!(
            manager.quick_access_status(""),
            Err(WincentError::InvalidPath(_))
        ) && matches!(
            manager.recent_file_position(""),
            Err(WincentError::InvalidPath(_))
        ) && matches!(
            manager.frequent_folder_slot(""),
            Err(WincentError::InvalidPath(_))
        ),
        "empty path must return InvalidPath",
    );

    // ---- Timings ---------------------------------------------------------
    if let Some(stats) = measure(PARSE_ITERATIONS, || {
        let _ = parse_file(&recent_path).expect("parse recent");
    }) {
        timings.insert("parse_recent".into(), stats);
    }
    if let Some(stats) = measure(PARSE_ITERATIONS, || {
        let _ = parse_file(&frequent_path).expect("parse frequent");
    }) {
        timings.insert("parse_frequent".into(), stats);
    }
    if let Some(stats) = measure(METADATA_ITERATIONS, || {
        let _ = fs::metadata(&recent_path).expect("metadata recent");
    }) {
        timings.insert("metadata_recent".into(), stats);
    }
    if let Some(sample) = recent_visible.first() {
        if let Some(stats) = measure(STATUS_ITERATIONS, || {
            let _ = manager.quick_access_status(sample).expect("status");
        }) {
            timings.insert("status_aggregate".into(), stats);
        }
        if let Some(stats) = measure(STATUS_ITERATIONS, || {
            let _ = manager.recent_file_position(sample).expect("recent");
        }) {
            timings.insert("status_recent_only".into(), stats);
        }
    }
    if let Some(sample) = frequent_visible.first() {
        if let Some(stats) = measure(STATUS_ITERATIONS, || {
            let _ = manager.frequent_folder_slot(&sample.0).expect("frequent");
        }) {
            timings.insert("status_frequent_only".into(), stats);
        }
    }

    // ---- Live mutation round trip ---------------------------------------
    if mutate {
        let recent_target = root.join("status-recent-target.txt");
        fs::write(&recent_target, b"wincent status query\n").expect("write recent target");

        manager
            .add_item(&recent_target, QuickAccess::RecentFiles, AddOptions::new())
            .expect("add recent target");
        let add_started = Instant::now();
        let recent_became_visible = wait_until(MUTATION_TIMEOUT, || {
            manager
                .recent_file_position(&recent_target)
                .map(|position| position.is_some())
                .unwrap_or(false)
        });
        let recent_visible_latency = add_started.elapsed().as_secs_f64() * 1000.0;
        mutation_log.push(format!(
            "recent_add_visible={recent_became_visible} latency_ms={recent_visible_latency:.3}"
        ));
        checks.check(
            "mutation_recent_add_becomes_visible",
            recent_became_visible,
            format!("latency_ms={recent_visible_latency:.3}"),
        );

        manager
            .remove_item(&recent_target, QuickAccess::RecentFiles)
            .expect("remove recent target");
        let remove_started = Instant::now();
        let recent_became_hidden = wait_until(MUTATION_TIMEOUT, || {
            manager
                .recent_file_position(&recent_target)
                .map(|position| position.is_none())
                .unwrap_or(false)
        });
        mutation_log.push(format!(
            "recent_remove_hidden={recent_became_hidden} latency_ms={:.3}",
            remove_started.elapsed().as_secs_f64() * 1000.0
        ));
        checks.check(
            "mutation_recent_remove_becomes_hidden",
            recent_became_hidden,
            "removed recent file must stop being visible",
        );

        let folder_target = root.join("status-frequent-target");
        fs::create_dir_all(&folder_target).expect("create frequent target");
        manager
            .pin_frequent_folder(&folder_target)
            .expect("pin frequent target");
        let pin_started = Instant::now();
        let pinned_visible = wait_until(MUTATION_TIMEOUT, || {
            manager
                .frequent_folder_slot(&folder_target)
                .map(|slot| slot.is_some_and(|slot| slot.is_pinned()))
                .unwrap_or(false)
        });
        mutation_log.push(format!(
            "frequent_pin_visible={pinned_visible} latency_ms={:.3}",
            pin_started.elapsed().as_secs_f64() * 1000.0
        ));
        checks.check(
            "mutation_frequent_pin_becomes_visible",
            pinned_visible,
            "pinned folder must report a pinned slot",
        );

        manager
            .remove_item(&folder_target, QuickAccess::FrequentFolders)
            .expect("remove frequent target");
        let unpin_started = Instant::now();
        let unpinned_hidden = wait_until(MUTATION_TIMEOUT, || {
            manager
                .frequent_folder_slot(&folder_target)
                .map(|slot| slot.is_none())
                .unwrap_or(false)
        });
        mutation_log.push(format!(
            "frequent_remove_hidden={unpinned_hidden} latency_ms={:.3}",
            unpin_started.elapsed().as_secs_f64() * 1000.0
        ));
        checks.check(
            "mutation_frequent_remove_becomes_hidden",
            unpinned_hidden,
            "removed frequent folder must stop being visible",
        );
    }

    // ---- Write artifacts -------------------------------------------------
    let mut json = String::new();
    json.push_str("{\n");
    json.push_str(&format!("  \"run_id\": \"{}\",\n", json_escape(&run_id)));
    json.push_str(&format!("  \"failures\": {},\n", checks.failures));
    json.push_str(&format!("  \"checks\": {},\n", checks.entries.len()));
    json.push_str("  \"summary\": {\n");
    let summary_lines: Vec<String> = summary
        .iter()
        .map(|(key, value)| format!("    \"{}\": {}", json_escape(key), value))
        .collect();
    json.push_str(&summary_lines.join(",\n"));
    json.push_str("\n  },\n");
    json.push_str("  \"timings\": {\n");
    let timing_lines: Vec<String> = timings
        .iter()
        .map(|(key, stats)| {
            format!(
                "    \"{}\": {{\"count\":{},\"mean_ms\":{:.4},\"median_ms\":{:.4},\"p95_ms\":{:.4},\"min_ms\":{:.4},\"max_ms\":{:.4}}}",
                json_escape(key),
                stats.count,
                stats.mean_ms,
                stats.median_ms,
                stats.p95_ms,
                stats.min_ms,
                stats.max_ms
            )
        })
        .collect();
    json.push_str(&timing_lines.join(",\n"));
    json.push_str("\n  },\n");
    json.push_str("  \"mutation\": [\n    ");
    json.push_str(
        &mutation_log
            .iter()
            .map(|line| format!("\"{}\"", json_escape(line)))
            .collect::<Vec<_>>()
            .join(",\n    "),
    );
    json.push_str("\n  ],\n");
    json.push_str("  \"check_results\": [\n");
    let check_lines: Vec<String> = checks
        .entries
        .iter()
        .map(|(id, ok, detail)| {
            format!(
                "    {{\"id\":\"{}\",\"ok\":{},\"detail\":\"{}\"}}",
                json_escape(id),
                ok,
                json_escape(detail)
            )
        })
        .collect();
    json.push_str(&check_lines.join(",\n"));
    json.push_str("\n  ]\n}\n");

    fs::write(&out_path, json).expect("write result json");

    let mut rows = String::from("metric,count,mean_ms,median_ms,p95_ms,min_ms,max_ms\n");
    for (name, stats) in &timings {
        rows.push_str(&format!(
            "{},{},{:.4},{:.4},{:.4},{:.4},{:.4}\n",
            name,
            stats.count,
            stats.mean_ms,
            stats.median_ms,
            stats.p95_ms,
            stats.min_ms,
            stats.max_ms
        ));
    }
    fs::write(&rows_path, rows).expect("write timing rows");

    if checks.failures > 0 {
        let failed: Vec<String> = checks
            .entries
            .iter()
            .filter(|(_, ok, _)| !ok)
            .map(|(id, _, detail)| format!("{id}: {detail}"))
            .collect();
        panic!(
            "{} status checks failed:\n{}",
            checks.failures,
            failed.join("\n")
        );
    }
}
