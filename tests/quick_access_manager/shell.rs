//! Independent Shell oracle. Never routes through wincent's DestList query APIs.

use std::sync::mpsc;
use std::time::Duration;

use wincent::{error::WincentError, QuickAccess, WincentResult};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER,
    COINIT_APARTMENTTHREADED,
};
use windows::Win32::System::Variant::VARIANT;
use windows::Win32::UI::Shell::{IShellDispatch, Shell};

pub fn query(category: QuickAccess) -> WincentResult<Vec<String>> {
    let (send, receive) = mpsc::channel();
    std::thread::spawn(move || {
        let _ = send.send(query_sta(category));
    });
    receive
        .recv_timeout(Duration::from_secs(15))
        .map_err(|error| WincentError::Timeout(format!("independent Shell oracle: {error}")))?
}

fn query_sta(category: QuickAccess) -> WincentResult<Vec<String>> {
    // SAFETY: This is a fresh dedicated thread with no previous COM apartment.
    unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()? };
    let result = (|| -> WincentResult<Vec<String>> {
        // SAFETY: COM is initialized on this thread; interfaces stay on this thread
        // and are released before CoUninitialize. All VARIANTs outlive their calls.
        unsafe {
            let shell: IShellDispatch =
                CoCreateInstance(&Shell, None, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER)?;
            let namespace = if category == QuickAccess::FrequentFolders {
                "shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}"
            } else {
                "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}"
            };
            let folder = shell.NameSpace(&VARIANT::from(namespace))?;
            let items = folder.Items()?;
            let mut paths = Vec::new();
            for index in 0..items.Count()? {
                let item = items.Item(&VARIANT::from(index))?;
                if category == QuickAccess::RecentFiles && item.IsFolder()?.as_bool() {
                    continue;
                }
                let path = item.Path()?.to_string();
                if !path.is_empty() {
                    paths.push(path);
                }
            }
            Ok(paths)
        }
    })();
    // SAFETY: Balances successful CoInitializeEx; all COM interfaces are dropped.
    unsafe { CoUninitialize() };
    result
}
