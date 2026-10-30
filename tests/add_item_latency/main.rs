//! VM-only average-latency test for `QuickAccessManager::add_item`.
//!
//! This integration test is not meant to run on a developer workstation: it
//! creates real Quick Access entries. It is ignored by default and additionally
//! gated behind `WINCENT_ADD_LATENCY=1`, so a plain `cargo test` (or
//! `cargo test -- --ignored`) never mutates the host.
//!
//! The VM driver `tests/add_item_latency/run_add_item_latency_vm.ps1` builds
//! this test binary, uploads it to the guest and runs it from the guest
//! terminal. See `tests/add_item_latency/vm_add_item_latency_guest.ps1` for the
//! full baseline/restore flow.
//!
//! Environment contract:
//!
//! | Variable | Meaning |
//! | --- | --- |
//! | `WINCENT_ADD_LATENCY` | Must be `1`; any other value skips the test. |
//! | `WINCENT_ADD_LATENCY_OUT` | Required output JSON path. |
//! | `WINCENT_ADD_LATENCY_ROOT` | Required scratch/targets directory. |
//! | `WINCENT_ADD_LATENCY_ITERS` | Timed iterations per series (default 20). |
//! | `WINCENT_ADD_LATENCY_WARMUP` | Warm-up iterations per series (default 2). |
//! | `WINCENT_ADD_LATENCY_SERIES` | Comma list, default `recent,frequent`. |
//!
//! Supported series identifiers are `recent` and `frequent`.

#![cfg(windows)]

use std::collections::BTreeMap;
use std::env;
use std::fs;
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

use wincent::prelude::*;

/// Maximum time a single Shell add may take before the manager returns a
/// timeout error. The measurement should fail instead of hanging the VM driver.
const ADD_TIMEOUT: Duration = Duration::from_secs(30);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Series {
    id: &'static str,
    kind: QuickAccess,
}

impl Series {
    fn parse_all(spec: &str) -> Result<Vec<Self>, String> {
        let mut series = Vec::new();
        for token in spec.split(',') {
            let token = token.trim();
            if token.is_empty() {
                continue;
            }
            let parsed = match token {
                "recent" => Self {
                    id: "recent",
                    kind: QuickAccess::RecentFiles,
                },
                "frequent" => Self {
                    id: "frequent",
                    kind: QuickAccess::FrequentFolders,
                },
                other => return Err(format!("unknown add-latency series: {other}")),
            };
            if !series.contains(&parsed) {
                series.push(parsed);
            }
        }
        if series.is_empty() {
            return Err("add-latency series list is empty".to_string());
        }
        Ok(series)
    }

    fn is_recent(self) -> bool {
        self.kind == QuickAccess::RecentFiles
    }
}

#[derive(Debug)]
struct Sample {
    series: &'static str,
    iteration: usize,
    target: String,
    elapsed_ms: f64,
    error: Option<String>,
}

#[derive(Debug)]
struct AddedTarget {
    recent: bool,
    path: String,
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
    fn from_samples(samples: &[f64]) -> Option<Self> {
        if samples.is_empty() {
            return None;
        }
        let mut sorted = samples.to_vec();
        sorted.sort_by(f64::total_cmp);
        let count = sorted.len();
        let sum: f64 = sorted.iter().sum();
        let mean = sum / count as f64;
        let median = if count % 2 == 1 {
            sorted[count / 2]
        } else {
            (sorted[count / 2 - 1] + sorted[count / 2]) / 2.0
        };
        // Nearest-rank p95: ceil(0.95 * count), clamped to 1..=count.
        let rank = ((count as f64 * 0.95).ceil() as usize).clamp(1, count);
        let p95 = sorted[rank - 1];
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
            p95_ms: p95,
            min_ms: sorted[0],
            max_ms: sorted[count - 1],
            stdev_ms: variance.sqrt(),
        })
    }
}

#[test]
#[ignore = "writes real Quick Access state; run inside the Windows test VM with WINCENT_ADD_LATENCY=1"]
fn add_item_shell_latency() {
    if env::var("WINCENT_ADD_LATENCY").as_deref() != Ok("1") {
        eprintln!(
            "add_item_shell_latency skipped: set WINCENT_ADD_LATENCY=1 and run this test only \
             inside the dedicated Windows test VM"
        );
        return;
    }

    let out_path = PathBuf::from(
        env::var("WINCENT_ADD_LATENCY_OUT")
            .expect("WINCENT_ADD_LATENCY_OUT must point at the result JSON path"),
    );
    let root = PathBuf::from(
        env::var("WINCENT_ADD_LATENCY_ROOT")
            .expect("WINCENT_ADD_LATENCY_ROOT must point at the scratch directory"),
    );
    let iterations = parse_env_usize("WINCENT_ADD_LATENCY_ITERS", 20);
    let warmup = parse_env_usize("WINCENT_ADD_LATENCY_WARMUP", 2);
    let series_spec =
        env::var("WINCENT_ADD_LATENCY_SERIES").unwrap_or_else(|_| "recent,frequent".to_string());
    let series = Series::parse_all(&series_spec).expect("invalid WINCENT_ADD_LATENCY_SERIES");
    assert!(iterations > 0, "WINCENT_ADD_LATENCY_ITERS must be > 0");

    fs::create_dir_all(&root).expect("failed to create scratch root");
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock before UNIX epoch")
        .as_millis();
    let manager = QuickAccessManager::builder().timeout(ADD_TIMEOUT).build();

    println!(
        "add_item latency: root={} iters={iterations} warmup={warmup} series={}",
        root.display(),
        series
            .iter()
            .map(|item| item.id)
            .collect::<Vec<_>>()
            .join(",")
    );

    let mut samples: Vec<Sample> = Vec::new();
    let mut added: Vec<AddedTarget> = Vec::new();
    let mut warmup_added = 0usize;
    let mut warmup_failed = 0usize;

    // Warm-up pass: same code path, unique throw-away targets, not recorded.
    for current in &series {
        for index in 0..warmup {
            let target = make_target(&root, "warmup", current, index, nonce, iterations);
            match manager.add_item(&target) {
                Ok(()) => {
                    warmup_added += 1;
                    added.push(AddedTarget {
                        recent: current.is_recent(),
                        path: target,
                    });
                }
                Err(error) => {
                    warmup_failed += 1;
                    eprintln!("warm-up add failed for {target}: {error}");
                }
            }
        }
    }

    for iteration in 0..iterations {
        // Rotate the series order every round so that one series does not
        // always observe the same cold/warm cache position.
        let offset = iteration % series.len();
        for position in 0..series.len() {
            let current = series[(position + offset) % series.len()];
            let target = make_target(&root, "timed", &current, iteration, nonce, iterations);
            let started = Instant::now();
            let result = manager.add_item(&target);
            let elapsed_ms = started.elapsed().as_secs_f64() * 1000.0;
            match result {
                Ok(()) => {
                    println!(
                        "{:>3} {:16} ok   {:>9.3} ms  {}",
                        iteration + 1,
                        current.id,
                        elapsed_ms,
                        target
                    );
                    samples.push(Sample {
                        series: current.id,
                        iteration: iteration + 1,
                        target: target.clone(),
                        elapsed_ms,
                        error: None,
                    });
                    added.push(AddedTarget {
                        recent: current.is_recent(),
                        path: target,
                    });
                }
                Err(error) => {
                    let message = error.to_string();
                    eprintln!(
                        "{:>3} {:16} FAIL {:>9.3} ms  {}: {}",
                        iteration + 1,
                        current.id,
                        elapsed_ms,
                        target,
                        message
                    );
                    samples.push(Sample {
                        series: current.id,
                        iteration: iteration + 1,
                        target,
                        elapsed_ms,
                        error: Some(message),
                    });
                }
            }
        }
    }

    // Independent raw check: read the DestList backing files directly and count
    // how many of the targets this run added are present. Query errors are
    // reported, not fatal, because this is diagnostic data only.
    let mut verification: BTreeMap<&'static str, (usize, usize, Option<String>)> = BTreeMap::new();
    for kind_name in ["recent", "frequent"] {
        let recent = kind_name == "recent";
        let expected = added.iter().filter(|item| item.recent == recent).count();
        if expected == 0 {
            continue;
        }
        let parsed = if recent {
            manager.get_recent_files_metadata()
        } else {
            manager.get_frequent_folders_metadata()
        };
        match parsed {
            Ok(entries) => {
                let matched = entries
                    .iter()
                    .filter(|entry| {
                        added.iter().any(|item| {
                            item.recent == recent
                                && entry.path().eq_ignore_ascii_case(item.path.as_str())
                        })
                    })
                    .count();
                verification.insert(kind_name, (expected, matched, None));
            }
            Err(error) => {
                verification.insert(kind_name, (expected, 0, Some(error.to_string())));
            }
        }
    }

    let stats_by_series: BTreeMap<&'static str, Stats> = series
        .iter()
        .filter_map(|current| {
            let values: Vec<f64> = samples
                .iter()
                .filter(|sample| sample.series == current.id && sample.error.is_none())
                .map(|sample| sample.elapsed_ms)
                .collect();
            Stats::from_samples(&values).map(|stats| (current.id, stats))
        })
        .collect();

    let stdout_table = render_table(&stats_by_series);
    println!("{stdout_table}");
    println!(
        "warm-up adds: ok={warmup_added} failed={warmup_failed}; verification: {:?}",
        verification
    );

    let failure_count = samples
        .iter()
        .filter(|sample| sample.error.is_some())
        .count();
    write_result_json(
        &out_path,
        &root,
        iterations,
        warmup,
        &series,
        &samples,
        &stats_by_series,
        &verification,
        warmup_added,
        warmup_failed,
    )
    .expect("failed to write add-latency result JSON");
    write_timings_csv(&out_path.with_file_name("timings.csv"), &samples)
        .expect("failed to write add-latency timings CSV");

    assert_eq!(
        failure_count,
        0,
        "{failure_count} of {} timed add_item calls failed; see {}",
        samples.len(),
        out_path.display()
    );
}

fn parse_env_usize(name: &str, default: usize) -> usize {
    match env::var(name) {
        Ok(value) => value
            .parse()
            .unwrap_or_else(|_| panic!("{name} must be a positive integer, got {value}")),
        Err(_) => default,
    }
}

fn make_target(
    root: &Path,
    phase: &str,
    series: &Series,
    index: usize,
    nonce: u128,
    iterations: usize,
) -> String {
    let width = iterations.max(1).to_string().len().max(3);
    let name = format!(
        "{phase}-{}-{index:0width$}-{nonce}",
        series.id,
        width = width
    );
    let path = if series.is_recent() {
        let file = root.join(format!("{name}.txt"));
        fs::write(&file, b"wincent add_item latency target\n")
            .expect("failed to create recent target file");
        file
    } else {
        let folder = root.join(name);
        fs::create_dir_all(&folder).expect("failed to create frequent target folder");
        folder
    };
    path.to_string_lossy().into_owned()
}

fn render_table(stats_by_series: &BTreeMap<&'static str, Stats>) -> String {
    let mut table = String::from(
        "series                 n     mean_ms    median_ms      p95_ms      min_ms      max_ms  \
         stdev_ms\n",
    );
    for (series, stats) in stats_by_series {
        table.push_str(&format!(
            "{series:<16} {:>6} {:>11.3} {:>11.3} {:>11.3} {:>11.3} {:>11.3} {:>10.3}\n",
            stats.count,
            stats.mean_ms,
            stats.median_ms,
            stats.p95_ms,
            stats.min_ms,
            stats.max_ms,
            stats.stdev_ms
        ));
    }
    table
}

#[allow(clippy::too_many_arguments)]
fn write_result_json(
    path: &Path,
    root: &Path,
    iterations: usize,
    warmup: usize,
    series: &[Series],
    samples: &[Sample],
    stats_by_series: &BTreeMap<&'static str, Stats>,
    verification: &BTreeMap<&'static str, (usize, usize, Option<String>)>,
    warmup_added: usize,
    warmup_failed: usize,
) -> std::io::Result<()> {
    let failure_count = samples
        .iter()
        .filter(|sample| sample.error.is_some())
        .count();
    let success_count = samples.len() - failure_count;
    let mut json = String::new();
    json.push_str("{\n");
    json.push_str("  \"test\": \"add_item_shell_latency\",\n");
    json.push_str(&format!("  \"iterations\": {iterations},\n"));
    json.push_str(&format!("  \"warmup\": {warmup},\n"));
    json.push_str(&format!(
        "  \"series_config\": [{}],\n",
        series
            .iter()
            .map(|item| format!(
                "{{\"id\": {}, \"kind\": {}}}",
                json_escape(item.id),
                json_escape(if item.is_recent() {
                    "recent"
                } else {
                    "frequent"
                })
            ))
            .collect::<Vec<_>>()
            .join(", ")
    ));
    json.push_str(&format!("  \"success_count\": {success_count},\n"));
    json.push_str(&format!("  \"failure_count\": {failure_count},\n"));
    json.push_str(&format!("  \"warmup_added\": {warmup_added},\n"));
    json.push_str(&format!("  \"warmup_failed\": {warmup_failed},\n"));
    json.push_str(&format!(
        "  \"targets_root\": {},\n",
        json_escape(&root.to_string_lossy())
    ));
    json.push_str("  \"stats\": {");
    if stats_by_series.is_empty() {
        json.push('}');
    } else {
        json.push('\n');
        let entries: Vec<String> = stats_by_series
            .iter()
            .map(|(series, stats)| {
                format!(
                    "    {}: {{\n      \"count\": {},\n      \"mean_ms\": {:.6},\n      \
                     \"median_ms\": {:.6},\n      \"p95_ms\": {:.6},\n      \"min_ms\": {:.6},\n      \
                     \"max_ms\": {:.6},\n      \"stdev_ms\": {:.6}\n    }}",
                    json_escape(series),
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
        json.push_str(&entries.join(",\n"));
        json.push_str("\n  },\n");
    }
    json.push_str("  \"verification\": {");
    if verification.is_empty() {
        json.push('}');
    } else {
        json.push('\n');
        let entries: Vec<String> = verification
            .iter()
            .map(|(kind, (expected, matched, error))| {
                format!(
                    "    {}: {{\"expected_adds\": {}, \"matched_raw_entries\": {}, \"error\": {}}}",
                    json_escape(kind),
                    expected,
                    matched,
                    error
                        .as_deref()
                        .map_or_else(|| "null".to_string(), json_escape)
                )
            })
            .collect();
        json.push_str(&entries.join(",\n"));
        json.push_str("\n  },\n");
    }
    json.push_str("  \"samples\": [\n");
    let rows: Vec<String> = samples
        .iter()
        .map(|sample| {
            format!(
                "    {{\"series\": {}, \"iteration\": {}, \"target\": {}, \"elapsed_ms\": {:.6}, \
                 \"ok\": {}, \"error\": {}}}",
                json_escape(sample.series),
                sample.iteration,
                json_escape(&sample.target),
                sample.elapsed_ms,
                sample.error.is_none(),
                sample
                    .error
                    .as_deref()
                    .map_or_else(|| "null".to_string(), json_escape)
            )
        })
        .collect();
    json.push_str(&rows.join(",\n"));
    json.push_str("\n  ]\n}\n");
    fs::write(path, json)
}

fn write_timings_csv(path: &Path, samples: &[Sample]) -> std::io::Result<()> {
    let mut csv = String::from("series,iteration,target,ok,elapsed_ms,error\r\n");
    for sample in samples {
        csv.push_str(&format!(
            "{},{},{},{},{:.6},{}\r\n",
            csv_escape(sample.series),
            sample.iteration,
            csv_escape(&sample.target),
            sample.error.is_none(),
            sample.elapsed_ms,
            sample.error.as_deref().map_or_else(String::new, csv_escape)
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

fn csv_escape(value: &str) -> String {
    if value.contains([',', '"', '\r', '\n']) {
        format!("\"{}\"", value.replace('"', "\"\""))
    } else {
        value.to_string()
    }
}
