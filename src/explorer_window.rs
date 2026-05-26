use crate::{error::WincentError, WincentResult};
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use std::path::{Path, PathBuf};
use std::thread;
use std::time::Duration;
use windows::core::{Interface, VARIANT};
use windows::Win32::Foundation::HANDLE;
use windows::Win32::System::Com::{
    CoCreateInstance, CoTaskMemFree, IDispatch, IServiceProvider, CLSCTX_LOCAL_SERVER,
};
use windows::Win32::System::Ole::READYSTATE_COMPLETE;
use windows::Win32::UI::Shell::{
    FOLDERID_Desktop, IShellBrowser, IShellWindows, IWebBrowser2, SHGetKnownFolderPath,
    SID_STopLevelBrowser, ShellWindows, KNOWN_FOLDER_FLAG,
};

const QUICK_ACCESS_NAMESPACE: &str = "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}";
const HOME_NAMESPACE: &str = "shell:::{f874310e-b6b7-47dc-bc84-b9e6b38f5903}";
const BROWSER_READY_POLL_INTERVAL: Duration = Duration::from_millis(200);

#[allow(dead_code)]
#[derive(Debug)]
pub(crate) struct NavigationCycleResult {
    pub(crate) original: ExplorerLocation,
    pub(crate) after_desktop: ExplorerLocation,
    pub(crate) restored: ExplorerLocation,
}

pub(crate) fn refresh_explorer_shell_views() -> WincentResult<()> {
    crate::com_thread::run_on_sta_thread(
        refresh_explorer_shell_views_on_sta,
        Duration::from_secs(10),
    )
}

#[allow(dead_code)]
pub(crate) fn navigate_recent_access_window_to_desktop_and_back(
) -> WincentResult<Option<NavigationCycleResult>> {
    crate::com_thread::run_on_sta_thread(
        navigate_recent_access_window_to_desktop_and_back_on_sta,
        Duration::from_secs(20),
    )
}

fn refresh_explorer_shell_views_on_sta() -> WincentResult<()> {
    let (shell_windows, count) = shell_windows_with_count()?;

    let recent_result = refresh_recent_access_shell_views(&shell_windows, count)?;
    if recent_result.matched > 0 {
        if recent_result.refreshed > 0 {
            return Ok(());
        }

        return Err(WincentError::SystemError(format!(
            "Found {} Quick Access/Home Explorer windows but failed to refresh any",
            recent_result.matched
        )));
    }

    let explorer_result = refresh_all_explorer_shell_views(&shell_windows, count)?;
    if explorer_result.matched > 0 && explorer_result.refreshed == 0 {
        return Err(WincentError::SystemError(format!(
            "Found {} Explorer windows but failed to refresh any",
            explorer_result.matched
        )));
    }

    Ok(())
}

fn shell_windows_with_count() -> WincentResult<(IShellWindows, i32)> {
    let shell_windows: IShellWindows = unsafe {
        CoCreateInstance(&ShellWindows, None, CLSCTX_LOCAL_SERVER)
    }
    .map_err(|e| WincentError::SystemError(format!("Failed to create ShellWindows: {}", e)))?;

    let count = unsafe { shell_windows.Count() }
        .map_err(|e| WincentError::SystemError(format!("Failed to get window count: {}", e)))?;

    Ok((shell_windows, count))
}

#[derive(Debug, Default)]
struct RefreshResult {
    matched: usize,
    refreshed: usize,
}

#[derive(Debug, Clone)]
pub(crate) struct ExplorerLocation {
    pub(crate) location_name: String,
    pub(crate) location_url: String,
}

fn navigate_recent_access_window_to_desktop_and_back_on_sta(
) -> WincentResult<Option<NavigationCycleResult>> {
    let (shell_windows, count) = shell_windows_with_count()?;
    let Some(web_browser) = find_recent_access_web_browser(&shell_windows, count)? else {
        return Ok(None);
    };

    let original = browser_location(&web_browser);
    let desktop = desktop_path()?;
    let desktop_url = file_url_from_path(&desktop);
    navigate_to_url(&web_browser, &desktop_url)?;
    let after_desktop = browser_location(&web_browser);

    navigate_back_to_location(&web_browser, &original)?;
    let restored = browser_location(&web_browser);

    Ok(Some(NavigationCycleResult {
        original,
        after_desktop,
        restored,
    }))
}

fn find_recent_access_web_browser(
    shell_windows: &IShellWindows,
    count: i32,
) -> WincentResult<Option<IWebBrowser2>> {
    for index in 0..count {
        let dispatch = match unsafe { shell_windows.Item(&index.into()) } {
            Ok(dispatch) => dispatch,
            Err(_) => continue,
        };

        let web_browser = match dispatch.cast::<IWebBrowser2>() {
            Ok(web_browser) => web_browser,
            Err(_) => continue,
        };

        let location = browser_location(&web_browser);
        if is_recent_access_location(&location) {
            return Ok(Some(web_browser));
        }
    }

    Ok(None)
}

fn refresh_recent_access_shell_views(
    shell_windows: &IShellWindows,
    count: i32,
) -> WincentResult<RefreshResult> {
    refresh_matching_shell_views(shell_windows, count, is_recent_access_location)
}

fn refresh_all_explorer_shell_views(
    shell_windows: &IShellWindows,
    count: i32,
) -> WincentResult<RefreshResult> {
    refresh_matching_shell_views(shell_windows, count, is_probable_explorer_location)
}

fn refresh_matching_shell_views(
    shell_windows: &IShellWindows,
    count: i32,
    matches_location: fn(&ExplorerLocation) -> bool,
) -> WincentResult<RefreshResult> {
    let mut result = RefreshResult::default();

    for index in 0..count {
        let dispatch = match unsafe { shell_windows.Item(&index.into()) } {
            Ok(dispatch) => dispatch,
            Err(_) => continue,
        };

        let web_browser = match dispatch.cast::<IWebBrowser2>() {
            Ok(web_browser) => web_browser,
            Err(_) => continue,
        };

        let location = browser_location(&web_browser);
        if !matches_location(&location) {
            continue;
        }

        result.matched += 1;
        if refresh_shell_view(&dispatch).is_ok() {
            result.refreshed += 1;
        }
    }

    Ok(result)
}

fn browser_location(web_browser: &IWebBrowser2) -> ExplorerLocation {
    let location_name = unsafe {
        web_browser
            .LocationName()
            .map(|value| value.to_string())
            .unwrap_or_else(|_| String::new())
    };
    let location_url = unsafe {
        web_browser
            .LocationURL()
            .map(|value| value.to_string())
            .unwrap_or_else(|_| String::new())
    };

    ExplorerLocation {
        location_name,
        location_url,
    }
}

fn navigate_to_url(web_browser: &IWebBrowser2, url: &str) -> WincentResult<()> {
    let target = VARIANT::from(url);
    let empty = VARIANT::default();
    unsafe {
        web_browser
            .Navigate2(
                &target,
                Some(&empty),
                Some(&empty),
                Some(&empty),
                Some(&empty),
            )
            .map_err(|e| {
                WincentError::SystemError(format!("Failed to navigate Explorer to {}: {}", url, e))
            })?;
    }
    wait_for_browser_ready(web_browser, Duration::from_secs(5));
    Ok(())
}

fn navigate_back_to_location(
    web_browser: &IWebBrowser2,
    original: &ExplorerLocation,
) -> WincentResult<()> {
    if !original.location_url.is_empty() {
        return navigate_to_url(web_browser, &original.location_url);
    }

    let mut errors = Vec::new();
    for candidate in [QUICK_ACCESS_NAMESPACE, HOME_NAMESPACE] {
        match navigate_to_url(web_browser, candidate) {
            Ok(()) => {
                let current = browser_location(web_browser);
                if is_recent_access_location(&current) {
                    return Ok(());
                }
            }
            Err(error) => errors.push(error.to_string()),
        }
    }

    Err(WincentError::SystemError(format!(
        "Failed to navigate Explorer back to Quick Access/Home: {}",
        errors.join("; ")
    )))
}

fn wait_for_browser_ready(web_browser: &IWebBrowser2, timeout: Duration) {
    let started = std::time::Instant::now();
    let mut polls = 0_u32;
    let mut last_busy = None;
    let mut last_ready = None;

    while started.elapsed() < timeout {
        let busy = unsafe {
            web_browser
                .Busy()
                .map(|value| value.0 != 0)
                .unwrap_or(false)
        };
        let ready = unsafe {
            web_browser
                .ReadyState()
                .map(|value| value == READYSTATE_COMPLETE)
                .unwrap_or(true)
        };
        polls += 1;
        last_busy = Some(busy);
        last_ready = Some(ready);

        if !busy && ready {
            return;
        }

        thread::sleep(BROWSER_READY_POLL_INTERVAL);
    }

    eprintln!(
        "Warning: Explorer browser did not report ready within {:.3}s after {} polls (last_busy={:?}, last_ready={:?})",
        timeout.as_secs_f64(),
        polls,
        last_busy,
        last_ready
    );
}

fn desktop_path() -> WincentResult<PathBuf> {
    let path = unsafe {
        SHGetKnownFolderPath(
            &FOLDERID_Desktop,
            KNOWN_FOLDER_FLAG(0x00),
            HANDLE(std::ptr::null_mut()),
        )
    }
    .map_err(|e| {
        WincentError::SystemError(format!("Failed to get Desktop known folder path: {}", e))
    })?;

    let desktop = unsafe {
        let wide = OsString::from_wide(path.as_wide());
        CoTaskMemFree(Some(path.as_ptr() as _));
        PathBuf::from(wide)
    };

    Ok(desktop)
}

fn file_url_from_path(path: &Path) -> String {
    let mut value = path.to_string_lossy().replace('\\', "/");
    if value.len() >= 2 && value.as_bytes()[1] == b':' {
        value = format!("/{value}");
    }
    format!("file://{value}")
}

fn refresh_shell_view(dispatch: &IDispatch) -> WincentResult<()> {
    let service_provider: IServiceProvider = dispatch.cast().map_err(|e| {
        WincentError::SystemError(format!(
            "Failed to cast Explorer dispatch to IServiceProvider: {}",
            e
        ))
    })?;
    let browser: IShellBrowser = unsafe {
        service_provider
            .QueryService(&SID_STopLevelBrowser)
            .map_err(|e| {
                WincentError::SystemError(format!("Failed to query top-level Shell browser: {}", e))
            })?
    };
    let view = unsafe {
        browser.QueryActiveShellView().map_err(|e| {
            WincentError::SystemError(format!("Failed to query active Shell view: {}", e))
        })?
    };

    unsafe {
        view.Refresh()
            .map_err(|e| WincentError::SystemError(format!("Failed to refresh Shell view: {}", e)))
    }
}

fn is_recent_access_location(location: &ExplorerLocation) -> bool {
    // Match only stable shell namespace URLs/GUIDs. Explorer display names are
    // localized and may be empty, so name-based matching is intentionally not
    // used; callers fall back to broader Explorer refresh when URL matching is
    // unavailable.
    is_home_or_recent_url(&location.location_url)
}

fn is_probable_explorer_location(location: &ExplorerLocation) -> bool {
    let lower = location.location_url.to_ascii_lowercase();
    lower.starts_with("file:///")
        || lower.starts_with("shell:::")
        || lower.starts_with("::{")
        || lower.starts_with("ms-shell")
        || location.location_url.is_empty() && !location.location_name.is_empty()
}

fn is_home_or_recent_url(url: &str) -> bool {
    let lower = url.to_ascii_lowercase();
    lower.contains("f874310e-b6b7-47dc-bc84-b9e6b38f5903")
        || lower.contains("679f85cb-0220-4080-b29b-5540cc05aab6")
}

#[cfg(test)]
fn list_detected_explorer_locations() -> WincentResult<Vec<ExplorerLocation>> {
    crate::com_thread::run_on_sta_thread(
        || {
            let (shell_windows, count) = shell_windows_with_count()?;
            let mut locations = Vec::new();

            for index in 0..count {
                let dispatch = match unsafe { shell_windows.Item(&index.into()) } {
                    Ok(dispatch) => dispatch,
                    Err(_) => continue,
                };
                let web_browser = match dispatch.cast::<IWebBrowser2>() {
                    Ok(web_browser) => web_browser,
                    Err(_) => continue,
                };
                let location = browser_location(&web_browser);
                if is_probable_explorer_location(&location) {
                    locations.push(location);
                }
            }

            Ok(locations)
        },
        Duration::from_secs(10),
    )
}

#[cfg(test)]
fn refresh_detected_recent_access_locations() -> WincentResult<RefreshResult> {
    crate::com_thread::run_on_sta_thread(
        || {
            let (shell_windows, count) = shell_windows_with_count()?;
            refresh_recent_access_shell_views(&shell_windows, count)
        },
        Duration::from_secs(10),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_recent_access_location_detection() {
        assert!(is_recent_access_location(&ExplorerLocation {
            location_name: String::new(),
            location_url: "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}".to_string(),
        }));
        assert!(!is_recent_access_location(&ExplorerLocation {
            location_name: "Quick Access".to_string(),
            location_url: String::new(),
        }));
        assert!(!is_recent_access_location(&ExplorerLocation {
            location_name: "\u{5feb}\u{901f}\u{8bbf}\u{95ee}".to_string(),
            location_url: String::new(),
        }));
        assert!(!is_recent_access_location(&ExplorerLocation {
            location_name: "Documents".to_string(),
            location_url: "file:///C:/Users/example/Documents".to_string(),
        }));
    }

    #[test]
    fn test_probable_explorer_location_detection() {
        assert!(is_probable_explorer_location(&ExplorerLocation {
            location_name: "Documents".to_string(),
            location_url: "file:///C:/Users/example/Documents".to_string(),
        }));
        assert!(is_probable_explorer_location(&ExplorerLocation {
            location_name: "Home".to_string(),
            location_url: String::new(),
        }));
        assert!(!is_probable_explorer_location(&ExplorerLocation {
            location_name: String::new(),
            location_url: "https://example.com".to_string(),
        }));
    }

    #[test]
    #[ignore = "Requires an interactive desktop session with Explorer windows"]
    fn test_list_detected_explorer_windows() -> WincentResult<()> {
        let locations = list_detected_explorer_locations()?;
        println!("Detected Explorer windows: {}", locations.len());
        for (index, location) in locations.iter().enumerate() {
            println!(
                "#{index}: name=\"{}\" url=\"{}\" recent_access={}",
                location.location_name,
                location.location_url,
                is_recent_access_location(location)
            );
        }
        Ok(())
    }

    #[test]
    #[ignore = "Requires an interactive desktop session with a Quick Access/Home Explorer window"]
    fn test_refresh_detected_recent_access_windows() -> WincentResult<()> {
        let result = refresh_detected_recent_access_locations()?;
        println!(
            "Detected Quick Access/Home Explorer windows: {}, refreshed: {}",
            result.matched, result.refreshed
        );

        if result.matched > 0 {
            assert!(
                result.refreshed > 0,
                "detected Quick Access/Home windows but refreshed none"
            );
        }

        Ok(())
    }

    #[test]
    #[ignore = "Requires an interactive desktop session with a Quick Access/Home Explorer window"]
    fn test_navigate_recent_access_window_to_desktop_and_back() -> WincentResult<()> {
        let Some(result) = navigate_recent_access_window_to_desktop_and_back()? else {
            println!("No Quick Access/Home Explorer window detected; nothing to navigate");
            return Ok(());
        };

        println!(
            "Original: name=\"{}\" url=\"{}\"",
            result.original.location_name, result.original.location_url
        );
        println!(
            "After desktop: name=\"{}\" url=\"{}\"",
            result.after_desktop.location_name, result.after_desktop.location_url
        );
        println!(
            "Restored: name=\"{}\" url=\"{}\"",
            result.restored.location_name, result.restored.location_url
        );

        assert!(
            result
                .after_desktop
                .location_url
                .to_ascii_lowercase()
                .starts_with("file:///"),
            "expected Explorer to navigate to a file URL for Desktop"
        );
        assert!(
            is_recent_access_location(&result.restored),
            "expected Explorer to navigate back to Quick Access/Home"
        );

        Ok(())
    }
}
