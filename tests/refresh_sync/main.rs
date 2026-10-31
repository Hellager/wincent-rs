//! VM-only experiment: which Explorer refresh path makes the open UI catch
//! up with a directly replaced Quick Access backing file.
//!
//! The test builds a small Recent Files fixture (5 visible items), then builds
//! a second state (the same 5 items plus one marker). While Quick Access is
//! open, it swaps only the `.automaticDestinations-ms` bytes to the second
//! state, with no notification and no refresh, then compares:
//!
//! 1. current program refresh (`QuickAccessManager::refresh_explorer`);
//! 2. an explicit `SHChangeNotify(SHCNE_ASSOCCHANGED)` followed by the same refresh.
//!
//! Phase markers are written for the host to capture screenshots; refresh
//! durations and Shell/rule/file observations are written to JSON.
//!
//! Ignored by default and gated behind `WINCENT_REFRESH_SYNC=1`.

#![cfg(windows)]

use std::env;
use std::ffi::OsStr;
use std::fs;
use std::io::Write;
use std::os::windows::ffi::OsStrExt;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::mpsc;
use std::time::{Duration, Instant};

use wincent::destlist::{parse_file, recent_files_dest_path};
use wincent::prelude::*;
use windows::core::{Interface, PCWSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::Storage::FileSystem::{ReplaceFileW, REPLACE_FILE_FLAGS};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER,
    COINIT_APARTMENTTHREADED,
};
use windows::Win32::System::Variant::VARIANT;
use windows::Win32::UI::Shell::{
    Folder, FolderItems, IShellWindows, IWebBrowser2, SHChangeNotify, ShellWindows,
    SHCNE_ASSOCCHANGED, SHCNF_IDLIST,
};
use windows::Win32::UI::WindowsAndMessaging::{SetForegroundWindow, ShowWindow, SW_MAXIMIZE};

const QUICK_ACCESS_NAMESPACE: &str = "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}";
const FIXTURE_ITEMS: usize = 5;
const MARKER_STEM: &str = "refresh-sync-marker";
const SHELL_TIMEOUT: Duration = Duration::from_secs(30);

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

#[test]
#[ignore = "replaces real Quick Access backing data; run inside the test VM with WINCENT_REFRESH_SYNC=1"]
fn refresh_sync() {
    if env::var("WINCENT_REFRESH_SYNC").as_deref() != Ok("1") {
        eprintln!(
            "refresh_sync skipped: set WINCENT_REFRESH_SYNC=1 and run only inside the \
             dedicated Windows test VM"
        );
        return;
    }

    let platform = TestPlatform::from_env();
    let out_path = PathBuf::from(
        env::var("WINCENT_REFRESH_SYNC_OUT")
            .expect("WINCENT_REFRESH_SYNC_OUT must point at the result JSON"),
    );
    let root = PathBuf::from(
        env::var("WINCENT_REFRESH_SYNC_ROOT")
            .expect("WINCENT_REFRESH_SYNC_ROOT must point at the fixture root"),
    );
    let phase_sleep_ms = env::var("WINCENT_REFRESH_SYNC_PHASE_SLEEP_MS")
        .ok()
        .and_then(|value| value.parse().ok())
        .unwrap_or(3000_u64);

    if root.exists() {
        fs::remove_dir_all(&root).expect("failed to clean fixture root");
    }
    fs::create_dir_all(root.join("phases")).expect("failed to create fixture root");
    let manager = QuickAccessManager::builder()
        .timeout(Duration::from_secs(60))
        .build();

    let marker_path = root.join(format!("{MARKER_STEM}.txt"));
    let mut fixture =
        build_fixture(&manager, &root, &marker_path, platform).expect("fixture build failed");
    if platform == TestPlatform::Win11 && (fixture.a_raw == 0 || fixture.b_raw <= fixture.a_raw) {
        let skipped = format!(
            concat!(
                "{{\n",
                "  \"test\": \"refresh_sync\",\n",
                "  \"platform\": \"win11\",\n",
                "  \"status\": \"skipped\",\n",
                "  \"reason\": \"native_recent_tracking_unavailable\",\n",
                "  \"fixture_mode\": \"native_v6_snapshots\",\n",
                "  \"fixture\": {{\"items\": {items}, \"a_raw\": {a_raw}, \"b_raw\": {b_raw}}}\n",
                "}}\n"
            ),
            items = FIXTURE_ITEMS,
            a_raw = fixture.a_raw,
            b_raw = fixture.b_raw,
        );
        fs::write(&out_path, skipped).expect("write skipped refresh-sync result");
        phase(&root, "80-complete", 100);
        return;
    }
    let dest = recent_files_dest_path().expect("recent DestFiles path");
    start_explorer_quick_access().expect("start Explorer Quick Access");
    if platform == TestPlatform::Win11 {
        fixture.a_visible = wait_shell_stable(20).expect("native baseline did not stabilize");
    } else {
        wait_shell_count(fixture.a_visible, 20).expect("baseline Shell count not reached");
    }

    let baseline = observe(&manager, &marker_path);
    println!("baseline: {baseline:?}");
    phase(&root, "10-baseline", phase_sleep_ms);

    // Trial 1: replace A -> B with no notification and no refresh.
    replace_cfb(&dest, &fixture.b_bytes).expect("replace CFB with B");
    std::thread::sleep(Duration::from_millis(800));
    let before_current = observe(&manager, &marker_path);
    println!("after B replacement (no refresh): {before_current:?}");
    phase(&root, "20-b-before-refresh", phase_sleep_ms);

    let current_refresh_started = Instant::now();
    manager.refresh_explorer().expect("current refresh failed");
    let current_refresh_ms = current_refresh_started.elapsed().as_secs_f64() * 1000.0;
    std::thread::sleep(Duration::from_millis(1200));
    let after_current = observe(&manager, &marker_path);
    let windows_after_current = quick_access_window_count().unwrap_or(0);
    println!(
        "after current refresh: {after_current:?}, quick_access_windows={windows_after_current}"
    );
    phase(&root, "30-after-current-refresh", phase_sleep_ms);

    // Reset to A and reset the visible window for the second trial.
    stop_explorer().expect("stop Explorer before reset");
    fs::write(&dest, &fixture.a_bytes).expect("restore A CFB");
    start_explorer_quick_access().expect("restart Explorer for trial 2");
    wait_shell_count(fixture.a_visible, 20).expect("trial 2 baseline Shell count not reached");
    phase(&root, "40-reset-a", 2000);

    // Trial 2: replace A -> B, then notify + refresh.
    replace_cfb(&dest, &fixture.b_bytes).expect("replace CFB with B for trial 2");
    std::thread::sleep(Duration::from_millis(800));
    let before_notify = observe(&manager, &marker_path);
    println!("after B replacement (before notify): {before_notify:?}");
    phase(&root, "50-b-before-notify", phase_sleep_ms);

    let notify_started = Instant::now();
    // SAFETY: SHCNE_ASSOCCHANGED with SHCNF_IDLIST does not consume item
    // pointers; both pointers are intentionally null.
    unsafe {
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
    }
    let notify_us = notify_started.elapsed().as_secs_f64() * 1_000_000.0;
    let notify_refresh_started = Instant::now();
    manager.refresh_explorer().expect("notify refresh failed");
    let notify_refresh_ms = notify_refresh_started.elapsed().as_secs_f64() * 1000.0;
    std::thread::sleep(Duration::from_millis(1200));
    let after_notify = observe(&manager, &marker_path);
    let windows_after_notify = quick_access_window_count().unwrap_or(0);
    println!(
        "after SHChangeNotify + refresh: {after_notify:?}, quick_access_windows={windows_after_notify}"
    );
    phase(&root, "60-after-notify-refresh", phase_sleep_ms);

    // Repeated current-refresh cost after the UI is already settled.
    let mut refresh_samples = Vec::new();
    for _ in 0..5 {
        let started = Instant::now();
        manager.refresh_explorer().expect("repeat refresh failed");
        refresh_samples.push(started.elapsed().as_secs_f64() * 1000.0);
    }

    // Trial 3: isolate SHChangeNotify without the refresh.
    stop_explorer().expect("stop Explorer before trial 3 reset");
    fs::write(&dest, &fixture.a_bytes).expect("restore A before trial 3");
    start_explorer_quick_access().expect("restart Explorer for trial 3");
    wait_shell_count(fixture.a_visible, 20).expect("trial 3 baseline Shell count not reached");
    phase(&root, "65-reset-a-2", 2000);

    replace_cfb(&dest, &fixture.b_bytes).expect("replace CFB with B for trial 3");
    std::thread::sleep(Duration::from_millis(800));
    let before_notify_only = observe(&manager, &marker_path);
    phase(&root, "70-b-before-notify-only", phase_sleep_ms);

    let notify_only_started = Instant::now();
    // SAFETY: same null-pointer ASSOCCHANGED notification as above.
    unsafe {
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
    }
    let notify_only_us = notify_only_started.elapsed().as_secs_f64() * 1_000_000.0;
    std::thread::sleep(Duration::from_millis(1500));
    let after_notify_only = observe(&manager, &marker_path);
    phase(&root, "75-after-notify-only", phase_sleep_ms);

    let notify_only_refresh_started = Instant::now();
    manager
        .refresh_explorer()
        .expect("notify-only follow-up refresh failed");
    let notify_only_refresh_ms = notify_only_refresh_started.elapsed().as_secs_f64() * 1000.0;
    std::thread::sleep(Duration::from_millis(1200));
    let after_notify_only_refresh = observe(&manager, &marker_path);
    phase(&root, "78-after-notify-only-refresh", phase_sleep_ms);

    let result = format!(
        concat!(
            "{{\n",
            "  \"test\": \"refresh_sync\",\n",
            "  \"platform\": {platform},\n",
            "  \"status\": \"completed\",\n",
            "  \"fixture_mode\": {fixture_mode},\n",
            "  \"fixture\": {{\"items\": {fixture_items}, \"marker\": {marker}, \"a_raw\": {a_raw}, \"b_raw\": {b_raw}, \"a_visible\": {a_visible}}},\n",
            "  \"observations\": {{\n",
            "    \"baseline\": {baseline},\n",
            "    \"after_b_without_refresh\": {before_current},\n",
            "    \"after_current_refresh\": {after_current},\n",
            "    \"quick_access_windows_after_current\": {windows_after_current},\n",
            "    \"after_b_before_notify\": {before_notify},\n",
            "    \"after_notify_refresh\": {after_notify},\n",
            "    \"quick_access_windows_after_notify\": {windows_after_notify},\n",
            "    \"after_b_before_notify_only\": {before_notify_only},\n",
            "    \"after_notify_only\": {after_notify_only},\n",
            "    \"after_notify_only_refresh\": {after_notify_only_refresh}\n",
            "  }},\n",
            "  \"timings\": {{\n",
            "    \"current_refresh_ms\": {current_refresh_ms:.6},\n",
            "    \"shchange_notify_us\": {notify_us:.3},\n",
            "    \"notify_refresh_ms\": {notify_refresh_ms:.6},\n",
            "    \"notify_plus_refresh_ms\": {notify_total:.6},\n",
            "    \"notify_only_us\": {notify_only_us:.3},\n",
            "    \"notify_only_refresh_ms\": {notify_only_refresh_ms:.6},\n",
            "    \"repeat_refresh_ms\": [{repeat}]\n",
            "  }}\n",
            "}}\n"
        ),
        marker = json_escape(&marker_path.to_string_lossy()),
        platform = json_escape(platform.as_str()),
        fixture_mode = json_escape(fixture.mode),
        fixture_items = FIXTURE_ITEMS,
        a_raw = fixture.a_raw,
        b_raw = fixture.b_raw,
        a_visible = fixture.a_visible,
        baseline = baseline.json(),
        before_current = before_current.json(),
        after_current = after_current.json(),
        windows_after_current = windows_after_current,
        before_notify = before_notify.json(),
        after_notify = after_notify.json(),
        windows_after_notify = windows_after_notify,
        before_notify_only = before_notify_only.json(),
        after_notify_only = after_notify_only.json(),
        after_notify_only_refresh = after_notify_only_refresh.json(),
        current_refresh_ms = current_refresh_ms,
        notify_us = notify_us,
        notify_refresh_ms = notify_refresh_ms,
        notify_total = notify_us / 1000.0 + notify_refresh_ms,
        notify_only_us = notify_only_us,
        notify_only_refresh_ms = notify_only_refresh_ms,
        repeat = refresh_samples
            .iter()
            .map(|value| format!("{value:.6}"))
            .collect::<Vec<_>>()
            .join(", ")
    );
    fs::write(&out_path, result).expect("write refresh-sync result");
    phase(&root, "80-complete", 500);
}

struct Fixture {
    a_bytes: Vec<u8>,
    b_bytes: Vec<u8>,
    a_raw: usize,
    b_raw: usize,
    a_visible: usize,
    mode: &'static str,
}

#[derive(Debug)]
struct Observation {
    shell_count: usize,
    shell_has_marker: bool,
    rule_has_marker: bool,
    file_has_marker: bool,
    file_raw: usize,
}

impl Observation {
    fn json(&self) -> String {
        format!(
            "{{\"shell_count\": {}, \"shell_has_marker\": {}, \"rule_has_marker\": {}, \"file_has_marker\": {}, \"file_raw\": {}}}",
            self.shell_count,
            self.shell_has_marker,
            self.rule_has_marker,
            self.file_has_marker,
            self.file_raw
        )
    }
}

fn observe(manager: &QuickAccessManager, marker: &Path) -> Observation {
    let marker = marker.to_string_lossy();
    let shell = shell_query().unwrap_or_default();
    let rule = manager
        .get_items(QuickAccess::RecentFiles)
        .unwrap_or_default();
    let path = recent_files_dest_path().expect("recent path");
    let parsed = parse_file(&path).expect("parse recent CFB");
    Observation {
        shell_count: shell.len(),
        shell_has_marker: shell
            .iter()
            .any(|item| item.eq_ignore_ascii_case(marker.as_ref())),
        rule_has_marker: rule
            .iter()
            .any(|item| item.eq_ignore_ascii_case(marker.as_ref())),
        file_has_marker: parsed
            .cfb_entries()
            .iter()
            .any(|entry| entry.path().eq_ignore_ascii_case(marker.as_ref())),
        file_raw: parsed.cfb_entries().len(),
    }
}

fn build_fixture(
    manager: &QuickAccessManager,
    root: &Path,
    marker: &Path,
    platform: TestPlatform,
) -> Result<Fixture, String> {
    if platform == TestPlatform::Win10 {
        stop_explorer()?;
        write_empty_recent_cfb()?;
    }

    for index in 0..FIXTURE_ITEMS {
        let path = root.join(format!("refresh-item-{index:02}.txt"));
        fs::write(&path, b"wincent refresh-sync fixture\n")
            .map_err(|error| format!("create {}: {error}", path.display()))?;
        add_fixture_item(manager, &path, platform)?;
    }
    if platform == TestPlatform::Win11 {
        manager
            .refresh_explorer()
            .map_err(|error| format!("refresh native A: {error}"))?;
        std::thread::sleep(Duration::from_millis(800));
    }
    let a_bytes = fs::read(recent_files_dest_path().map_err(|error| error.to_string())?)
        .map_err(|error| format!("read A: {error}"))?;
    let a_raw = raw_count()?;
    let a_visible = visible_count()?;

    fs::write(marker, b"wincent refresh-sync marker\n")
        .map_err(|error| format!("create marker: {error}"))?;
    add_fixture_item(manager, marker, platform)?;
    if platform == TestPlatform::Win11 {
        manager
            .refresh_explorer()
            .map_err(|error| format!("refresh native B: {error}"))?;
        std::thread::sleep(Duration::from_millis(800));
    }
    let b_bytes = fs::read(recent_files_dest_path().map_err(|error| error.to_string())?)
        .map_err(|error| format!("read B: {error}"))?;
    let b_raw = raw_count()?;
    if platform == TestPlatform::Win10 && (a_raw != FIXTURE_ITEMS || b_raw != FIXTURE_ITEMS + 1) {
        return Err(format!(
            "unexpected fixture raw counts: A={a_raw}, B={b_raw}"
        ));
    }

    // Remove the marker's standalone Recent shortcut so it cannot influence
    // the baseline UI/Shell state; the embedded numeric stream is enough.
    remove_marker_links()?;

    stop_explorer()?;
    // Leave A as the initial on-disk state.
    fs::write(
        recent_files_dest_path().map_err(|error| error.to_string())?,
        &a_bytes,
    )
    .map_err(|error| format!("restore A: {error}"))?;
    Ok(Fixture {
        a_bytes,
        b_bytes,
        a_raw,
        b_raw,
        a_visible,
        mode: if platform == TestPlatform::Win10 {
            "synthetic_v4"
        } else {
            "native_v6_snapshots"
        },
    })
}

fn visible_count() -> Result<usize, String> {
    Ok(shell_query()?.len())
}

fn add_fixture_item(
    manager: &QuickAccessManager,
    path: &Path,
    platform: TestPlatform,
) -> Result<(), String> {
    match platform {
        TestPlatform::Win10 => add_direct(path),
        TestPlatform::Win11 => manager
            .add_item(path)
            .map_err(|error| format!("add native item {}: {error}", path.display())),
    }
}

fn raw_count() -> Result<usize, String> {
    let path = recent_files_dest_path().map_err(|error| error.to_string())?;
    let parsed = parse_file(&path).map_err(|error| error.to_string())?;
    Ok(parsed.cfb_entries().len())
}

fn write_empty_recent_cfb() -> Result<(), String> {
    let path = recent_files_dest_path().map_err(|error| error.to_string())?;
    let mut header = vec![0u8; 0x20];
    header[0..4].copy_from_slice(&4u32.to_le_bytes());
    let mut compound =
        cfb::CompoundFile::create_with_version(cfb::Version::V3, std::io::Cursor::new(Vec::new()))
            .map_err(|error| format!("create empty CFB: {error}"))?;
    compound
        .create_stream("/DestList")
        .map_err(|error| format!("create DestList: {error}"))?
        .write_all(&header)
        .map_err(|error| format!("write DestList: {error}"))?;
    compound
        .flush()
        .map_err(|error| format!("flush empty CFB: {error}"))?;
    fs::write(&path, compound.into_inner().into_inner())
        .map_err(|error| format!("write {}: {error}", path.display()))
}

fn add_direct(path: &Path) -> Result<(), String> {
    let manager = QuickAccessManager::builder()
        .timeout(Duration::from_secs(60))
        .build();
    if manager
        .add_item_direct(path, QuickAccess::RecentFiles)
        .is_ok()
    {
        return Ok(());
    }
    // Explorer is stopped during setup, so the post-write refresh can fail
    // while the direct write itself succeeded.
    let entries = manager
        .get_recent_files_metadata()
        .map_err(|error| error.to_string())?;
    if entries
        .iter()
        .any(|entry| entry.path().eq_ignore_ascii_case(&path.to_string_lossy()))
    {
        Ok(())
    } else {
        Err(format!("direct add failed for {}", path.display()))
    }
}

fn remove_marker_links() -> Result<(), String> {
    let recent = PathBuf::from(env::var("APPDATA").map_err(|error| error.to_string())?)
        .join("Microsoft\\Windows\\Recent");
    for entry in fs::read_dir(&recent).map_err(|error| format!("read Recent: {error}"))? {
        let path = entry.map_err(|error| error.to_string())?.path();
        if path.extension() == Some(OsStr::new("lnk"))
            && path
                .file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with(MARKER_STEM))
        {
            fs::remove_file(&path).map_err(|error| format!("remove link: {error}"))?;
        }
    }
    Ok(())
}

fn replace_cfb(dest: &Path, bytes: &[u8]) -> Result<(), String> {
    let temp = dest.with_extension("refresh-sync-tmp");
    fs::write(&temp, bytes).map_err(|error| format!("write temp CFB: {error}"))?;
    let dest_wide: Vec<u16> = dest.as_os_str().encode_wide().chain(Some(0)).collect();
    let temp_wide: Vec<u16> = temp.as_os_str().encode_wide().chain(Some(0)).collect();
    let result = unsafe {
        ReplaceFileW(
            PCWSTR(dest_wide.as_ptr()),
            PCWSTR(temp_wide.as_ptr()),
            PCWSTR::null(),
            REPLACE_FILE_FLAGS(0),
            None,
            None,
        )
    };
    if let Err(error) = result {
        let _ = fs::remove_file(&temp);
        fs::write(dest, bytes).map_err(|fallback| {
            format!("ReplaceFileW failed ({error}); write failed: {fallback}")
        })?;
    }
    Ok(())
}

fn stop_explorer() -> Result<(), String> {
    let windir = env::var("WINDIR").unwrap_or_else(|_| "C:\\Windows".to_string());
    let _ = Command::new(format!("{windir}\\System32\\taskkill.exe"))
        .args(["/F", "/IM", "explorer.exe"])
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status();
    std::thread::sleep(Duration::from_millis(1500));
    Ok(())
}

fn start_explorer_quick_access() -> Result<(), String> {
    stop_explorer()?;
    let windir = env::var("WINDIR").unwrap_or_else(|_| "C:\\Windows".to_string());
    Command::new(format!("{windir}\\explorer.exe"))
        .arg(QUICK_ACCESS_NAMESPACE)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .map_err(|error| format!("start explorer: {error}"))?;
    std::thread::sleep(Duration::from_millis(3000));
    if let Err(error) = focus_quick_access_window() {
        eprintln!("focus Quick Access window failed: {error}");
    }
    std::thread::sleep(Duration::from_millis(1000));
    Ok(())
}

fn focus_quick_access_window() -> Result<(), String> {
    let (tx, rx) = mpsc::channel();
    std::thread::spawn(move || {
        let _ = tx.send(focus_quick_access_sta());
    });
    rx.recv_timeout(SHELL_TIMEOUT)
        .map_err(|error| format!("focus timeout: {error}"))?
}

fn focus_quick_access_sta() -> Result<(), String> {
    // SAFETY: dedicated thread; successful init is balanced below.
    if !unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).is_ok() } {
        return Err("CoInitializeEx failed".to_string());
    }
    let result = (|| -> Result<(), String> {
        // SAFETY: COM is initialized; ShellWindows is a local-server COM class.
        let windows: IShellWindows =
            unsafe { CoCreateInstance(&ShellWindows, None, CLSCTX_LOCAL_SERVER) }
                .map_err(|error| format!("ShellWindows: {error}"))?;
        // SAFETY: live collection.
        let count = unsafe { windows.Count() }.map_err(|error| format!("Count: {error}"))?;
        for index in 0..count {
            let dispatch = match unsafe { windows.Item(&VARIANT::from(index)) } {
                Ok(dispatch) => dispatch,
                Err(_) => continue,
            };
            let browser = match dispatch.cast::<IWebBrowser2>() {
                Ok(browser) => browser,
                Err(_) => continue,
            };
            let url = unsafe { browser.LocationURL() }
                .map(|value| value.to_string())
                .unwrap_or_default();
            let name = unsafe { browser.LocationName() }
                .map(|value| value.to_string())
                .unwrap_or_default();
            if !is_quick_access_location(&name, &url) {
                continue;
            }
            let hwnd = unsafe { browser.HWND() }.map_err(|error| format!("HWND: {error}"))?;
            let hwnd = HWND(hwnd.0 as *mut _);
            // SAFETY: hwnd belongs to the Explorer window found above.
            unsafe {
                let _ = ShowWindow(hwnd, SW_MAXIMIZE);
                let _ = SetForegroundWindow(hwnd);
            }
            return Ok(());
        }
        Err("Quick Access window not found".to_string())
    })();
    // SAFETY: balances the successful CoInitializeEx above.
    unsafe { CoUninitialize() };
    result
}

fn quick_access_window_count() -> Result<usize, String> {
    let (tx, rx) = mpsc::channel();
    std::thread::spawn(move || {
        let _ = tx.send(quick_access_window_count_sta());
    });
    rx.recv_timeout(SHELL_TIMEOUT)
        .map_err(|error| format!("window count timeout: {error}"))?
}

fn quick_access_window_count_sta() -> Result<usize, String> {
    // SAFETY: dedicated thread; successful init is balanced below.
    if !unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).is_ok() } {
        return Err("CoInitializeEx failed".to_string());
    }
    let result = (|| -> Result<usize, String> {
        // SAFETY: COM is initialized; ShellWindows is a local-server COM class.
        let windows: IShellWindows =
            unsafe { CoCreateInstance(&ShellWindows, None, CLSCTX_LOCAL_SERVER) }
                .map_err(|error| format!("ShellWindows: {error}"))?;
        // SAFETY: live collection.
        let count = unsafe { windows.Count() }.map_err(|error| format!("Count: {error}"))?;
        let mut matching = 0;
        for index in 0..count {
            let dispatch = match unsafe { windows.Item(&VARIANT::from(index)) } {
                Ok(dispatch) => dispatch,
                Err(_) => continue,
            };
            let browser = match dispatch.cast::<IWebBrowser2>() {
                Ok(browser) => browser,
                Err(_) => continue,
            };
            let url = unsafe { browser.LocationURL() }
                .map(|value| value.to_string())
                .unwrap_or_default();
            let name = unsafe { browser.LocationName() }
                .map(|value| value.to_string())
                .unwrap_or_default();
            if is_quick_access_location(&name, &url) {
                matching += 1;
            }
        }
        Ok(matching)
    })();
    // SAFETY: balances the successful CoInitializeEx above.
    unsafe { CoUninitialize() };
    result
}

fn is_quick_access_location(name: &str, url: &str) -> bool {
    let lower = url.to_ascii_lowercase();
    lower.contains("679f85cb-0220-4080-b29b-5540cc05aab6")
        || lower.contains("f874310e-b6b7-47dc-bc84-b9e6b38f5903")
        || url.is_empty() && !name.is_empty()
}

fn wait_shell_count(expected: usize, attempts: usize) -> Result<(), String> {
    let mut last = 0;
    for _ in 0..attempts {
        if let Ok(items) = shell_query() {
            last = items.len();
            if last == expected {
                return Ok(());
            }
        }
        std::thread::sleep(Duration::from_millis(500));
    }
    Err(format!("expected {expected} Shell items, last {last}"))
}

fn wait_shell_stable(attempts: usize) -> Result<usize, String> {
    let mut last = None;
    let mut stable = 0;
    for _ in 0..attempts {
        if let Ok(items) = shell_query() {
            let count = items.len();
            if Some(count) == last {
                stable += 1;
                if stable >= 2 {
                    return Ok(count);
                }
            } else {
                last = Some(count);
                stable = 0;
            }
        }
        std::thread::sleep(Duration::from_millis(500));
    }
    Err(format!("Shell count did not stabilize, last={last:?}"))
}

fn shell_query() -> Result<Vec<String>, String> {
    let (tx, rx) = mpsc::channel();
    std::thread::spawn(move || {
        let _ = tx.send(shell_query_sta());
    });
    rx.recv_timeout(SHELL_TIMEOUT)
        .map_err(|error| format!("Shell query timeout: {error}"))?
}

fn shell_query_sta() -> Result<Vec<String>, String> {
    // SAFETY: dedicated worker; successful init is balanced below.
    if !unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).is_ok() } {
        return Err("CoInitializeEx failed".to_string());
    }
    let result = (|| -> Result<Vec<String>, String> {
        // SAFETY: COM is initialized; Shell.Application is an in-proc/local server.
        let shell: windows::Win32::UI::Shell::IShellDispatch = unsafe {
            CoCreateInstance(
                &windows::Win32::UI::Shell::Shell,
                None,
                CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
            )
        }
        .map_err(|error| format!("Shell: {error}"))?;
        let namespace = VARIANT::from(QUICK_ACCESS_NAMESPACE);
        // SAFETY: shell and namespace are live for the call.
        let folder: Folder = unsafe {
            shell
                .NameSpace(&namespace)
                .map_err(|error| format!("NameSpace: {error}"))?
        };
        // SAFETY: folder is a live Shell namespace.
        let items: FolderItems =
            unsafe { folder.Items() }.map_err(|error| format!("Items: {error}"))?;
        // SAFETY: live collection.
        let count = unsafe { items.Count() }.map_err(|error| format!("Count: {error}"))?;
        let mut paths = Vec::new();
        for index in 0..count {
            let variant = VARIANT::from(index);
            let item = match unsafe { items.Item(&variant) } {
                Ok(item) => item,
                Err(_) => continue,
            };
            let is_folder = unsafe { item.IsFolder().map(bool::from).unwrap_or(false) };
            if is_folder {
                continue;
            }
            if let Ok(path) = unsafe { item.Path() } {
                let value = path.to_string();
                if !value.trim().is_empty() {
                    paths.push(value);
                }
            }
        }
        Ok(paths)
    })();
    // SAFETY: balances the successful CoInitializeEx above.
    unsafe { CoUninitialize() };
    result
}

fn phase(root: &Path, name: &str, sleep_ms: u64) {
    let phases = root.join("phases");
    let path = phases.join(format!("{name}.txt"));
    fs::write(&path, "ready\n").expect("write phase marker");
    println!("phase {name}");
    std::thread::sleep(Duration::from_millis(sleep_ms));
}

fn json_escape(value: &str) -> String {
    let mut out = String::from("\"");
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
