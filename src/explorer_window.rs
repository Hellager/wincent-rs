use crate::{error::WincentError, WincentResult};
use std::time::Duration;
use windows::core::Interface;
use windows::Win32::System::Com::{
    CoCreateInstance, IDispatch, IServiceProvider, CLSCTX_LOCAL_SERVER,
};
use windows::Win32::UI::Shell::{
    IShellBrowser, IShellWindows, IWebBrowser2, SID_STopLevelBrowser, ShellWindows,
};

pub(crate) fn refresh_explorer_shell_views() -> WincentResult<()> {
    crate::com_thread::run_on_sta_thread(
        refresh_explorer_shell_views_on_sta,
        Duration::from_secs(10),
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

#[derive(Debug)]
struct ExplorerLocation {
    location_name: String,
    location_url: String,
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
    is_home_or_recent_url(&location.location_url)
        || is_empty_url_recent_access_name(&location.location_name)
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

fn is_empty_url_recent_access_name(name: &str) -> bool {
    let normalized = name.trim().to_ascii_lowercase();
    normalized == "quick access"
        || normalized == "home"
        || name.trim() == "\u{5feb}\u{901f}\u{8bbf}\u{95ee}"
        || name.trim() == "\u{4e3b}\u{9875}"
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
        assert!(is_recent_access_location(&ExplorerLocation {
            location_name: "Quick Access".to_string(),
            location_url: String::new(),
        }));
        assert!(is_recent_access_location(&ExplorerLocation {
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
}
