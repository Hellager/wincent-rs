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

const NATIVE_RECENT_ITEMS: usize = 20;
const NATIVE_FREQUENT_ITEMS: usize = 2;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum TestPlatform {
    Win10,
    Win11,
}

impl TestPlatform {
    fn from_env() -> Self {
        match env::var("WINCENT_TEST_PLATFORM").as_deref() {
            Ok("win10") => Self::Win10,
            Ok("win11") => Self::Win11,
            Ok(value) => panic!("WINCENT_TEST_PLATFORM must be win10 or win11, got {value}"),
            Err(_) => panic!("WINCENT_TEST_PLATFORM must be set by the VM guest harness"),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Win10 => "win10",
            Self::Win11 => "win11",
        }
    }
}

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

    let platform = TestPlatform::from_env();
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

    let mut samples = Vec::new();
    let (fixture_mode, scenarios) = run_native(&manager, &root, iterations, warmup, &mut samples)
        .unwrap_or_else(|error| panic!("fixture or measurement failed: {error}"));
    let stats = summarize(&samples, &scenarios);
    print_summary(&stats);
    write_result_json(
        &out_path,
        platform,
        fixture_mode,
        (iterations, warmup),
        &scenarios,
        &samples,
        &stats,
    )
    .expect("failed to write get_item_latency result JSON");
    write_timings_csv(&out_path.with_file_name("timings.csv"), &samples)
        .expect("failed to write get_item_latency timings CSV");
}

fn run_native(
    manager: &QuickAccessManager,
    root: &Path,
    iterations: usize,
    warmup: usize,
    samples: &mut Vec<Sample>,
) -> Result<(&'static str, Vec<Scenario>), String> {
    for index in 0..NATIVE_RECENT_ITEMS {
        let path = recent_target(root, index);
        fs::write(&path, b"wincent native get_item latency fixture\n")
            .map_err(|error| format!("create {}: {error}", path.display()))?;
        manager
            .add_item(&path)
            .map_err(|error| format!("add native recent {}: {error}", path.display()))?;
    }
    for index in 0..NATIVE_FREQUENT_ITEMS {
        let path = frequent_target(root, index);
        fs::create_dir_all(&path).map_err(|error| format!("create {}: {error}", path.display()))?;
        manager
            .add_item(&path)
            .map_err(|error| format!("pin native frequent {}: {error}", path.display()))?;
    }

    stop_explorer()?;
    start_explorer()?;
    let mut scenarios = Vec::new();
    if let Some(recent) = native_scenario(manager, "recent_native", QuickAccess::RecentFiles)? {
        measure_scenario(&recent, iterations, warmup, samples)?;
        scenarios.push(recent);
    }
    if let Some(frequent) =
        native_scenario(manager, "frequent_native", QuickAccess::FrequentFolders)?
    {
        measure_scenario(&frequent, iterations, warmup, samples)?;
        scenarios.push(frequent);
    }
    if scenarios.is_empty() {
        return Err("native Quick Access lists are both unavailable".to_string());
    }
    Ok(("native", scenarios))
}

fn native_scenario(
    manager: &QuickAccessManager,
    id: &'static str,
    kind: QuickAccess,
) -> Result<Option<Scenario>, String> {
    let expected_raw = raw_entries(manager, kind)?.len();
    let expected_shell = shell_query(kind)?.len();
    if expected_raw == 0 || expected_shell == 0 {
        println!("native {kind:?} unavailable: raw={expected_raw} shell={expected_shell}");
        return Ok(None);
    }
    Ok(Some(Scenario {
        id,
        kind,
        expected_shell,
        expected_raw,
    }))
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
    _manager: &QuickAccessManager,
    kind: QuickAccess,
) -> Result<Vec<DestListEntry>, String> {
    let parsed = parse_file(source_path(kind)?).map_err(|error| error.to_string())?;
    Ok(wincent::destlist::entries(parsed.dest_list()))
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
    platform: TestPlatform,
    fixture_mode: &str,
    run_counts: (usize, usize),
    scenarios: &[Scenario],
    samples: &[Sample],
    stats: &BTreeMap<&'static str, BTreeMap<&'static str, Stats>>,
) -> std::io::Result<()> {
    let (iterations, warmup) = run_counts;
    let mut json = String::new();
    json.push_str("{\n");
    json.push_str("  \"test\": \"get_item_latency\",\n");
    json.push_str(&format!(
        "  \"platform\": {},\n",
        json_escape(platform.as_str())
    ));
    json.push_str(&format!(
        "  \"fixture_mode\": {},\n",
        json_escape(fixture_mode)
    ));
    let recent_available = scenarios
        .iter()
        .any(|scenario| scenario.kind == QuickAccess::RecentFiles);
    let frequent_available = scenarios
        .iter()
        .any(|scenario| scenario.kind == QuickAccess::FrequentFolders);
    json.push_str(&format!(
        "  \"fixture_capabilities\": {{\"recent\": {}, \"frequent\": {}}},\n",
        recent_available, frequent_available
    ));
    json.push_str(&format!("  \"iterations\": {iterations},\n"));
    json.push_str(&format!("  \"warmup\": {warmup},\n"));
    json.push_str(&format!(
        "  \"fixture\": {{\"requested_recent\": {}, \"requested_frequent\": {}}},\n",
        NATIVE_RECENT_ITEMS, NATIVE_FREQUENT_ITEMS
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
