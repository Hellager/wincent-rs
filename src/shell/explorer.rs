use crate::{error::WincentError, WincentResult};
use std::time::Duration;
use windows::core::Interface;
use windows::Win32::System::Com::{
    CoCreateInstance, IDispatch, IServiceProvider, CLSCTX_LOCAL_SERVER,
};
use windows::Win32::UI::Shell::{
    IShellBrowser, IShellWindows, IWebBrowser2, SHChangeNotify, SID_STopLevelBrowser, ShellWindows,
    SHCNE_ASSOCCHANGED, SHCNF_IDLIST,
};

pub(crate) fn notify_shell_change() {
    // SAFETY: SHCNE_ASSOCCHANGED with SHCNF_IDLIST accepts null item pointers.
    unsafe {
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
    }
}

#[cfg(test)]
pub(crate) fn refresh_quick_access(timeout: Duration) -> WincentResult<()> {
    notify_shell_change();
    refresh_shell_views_with_timeout(timeout)
}

pub(crate) fn refresh_shell_views_with_timeout(timeout: Duration) -> WincentResult<()> {
    crate::shell::thread::run_on_sta_thread(refresh_shell_views_on_sta, timeout)
}

fn refresh_shell_views_on_sta() -> WincentResult<()> {
    let (shell_windows, count) = shell_windows_with_count()?;

    let recent_result = refresh_recent_access_shell_views(&shell_windows, count)?;
    if recent_result.matched > 0 && recent_result.refreshed == 0 {
        return Err(WincentError::SystemError(format!(
            "Found {} Quick Access/Home Explorer windows but failed to refresh any",
            recent_result.matched
        )));
    }

    Ok(())
}

fn shell_windows_with_count() -> WincentResult<(IShellWindows, i32)> {
    // SAFETY: Called from a COM-initialized STA worker. ShellWindows is a
    // local-server COM class and the returned interface is owned by the wrapper.
    let shell_windows: IShellWindows = unsafe {
        CoCreateInstance(&ShellWindows, None, CLSCTX_LOCAL_SERVER)
    }
    .map_err(|e| WincentError::SystemError(format!("Failed to create ShellWindows: {}", e)))?;

    // SAFETY: `shell_windows` came from a successful CoCreateInstance call and
    // remains alive for the duration of the Count call.
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

fn refresh_recent_access_shell_views(
    shell_windows: &IShellWindows,
    count: i32,
) -> WincentResult<RefreshResult> {
    refresh_matching_shell_views(shell_windows, count, is_recent_access_location)
}

fn refresh_matching_shell_views(
    shell_windows: &IShellWindows,
    count: i32,
    matches_location: fn(&ExplorerLocation) -> bool,
) -> WincentResult<RefreshResult> {
    let mut result = RefreshResult::default();

    for index in 0..count {
        // SAFETY: `shell_windows` is a live COM collection from ShellWindows;
        // failures for individual windows are handled by skipping them.
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
    // SAFETY: `web_browser` is a live IWebBrowser2 interface obtained from
    // ShellWindows; errors are converted to empty strings for best-effort detection.
    let location_name = unsafe {
        web_browser
            .LocationName()
            .map(|value| value.to_string())
            .unwrap_or_else(|_| String::new())
    };
    // SAFETY: Same live IWebBrowser2 interface as above; LocationURL is
    // best-effort and failures are intentionally treated as unknown.
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
    // SAFETY: QueryService is called on the IServiceProvider interface obtained
    // from the Explorer dispatch and returns the active top-level Shell browser.
    let browser: IShellBrowser = unsafe {
        service_provider
            .QueryService(&SID_STopLevelBrowser)
            .map_err(|e| {
                WincentError::SystemError(format!("Failed to query top-level Shell browser: {}", e))
            })?
    };
    // SAFETY: `browser` is a live IShellBrowser returned by QueryService.
    let view = unsafe {
        browser.QueryActiveShellView().map_err(|e| {
            WincentError::SystemError(format!("Failed to query active Shell view: {}", e))
        })?
    };

    // SAFETY: `view` is the active Shell view returned by IShellBrowser and is
    // valid for this immediate Refresh call.
    unsafe {
        view.Refresh()
            .map_err(|e| WincentError::SystemError(format!("Failed to refresh Shell view: {}", e)))
    }
}

fn is_recent_access_location(location: &ExplorerLocation) -> bool {
    // Home/Quick Access display names are localized. On some Windows builds,
    // ShellWindows also reports their LocationURL as empty, so a non-empty
    // display name with no URL is the language-neutral fallback.
    is_home_or_recent_url(&location.location_url)
        || location.location_url.is_empty() && !location.location_name.is_empty()
}

#[cfg(test)]
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
            location_name: String::new(),
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
}
