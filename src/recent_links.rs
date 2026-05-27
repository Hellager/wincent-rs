//! Helpers for cleaning Windows Recent folder shortcuts.

use crate::error::WincentError;
use crate::utils::{get_windows_recent_folder, paths_equal};
use crate::WincentResult;
use std::fs;
use std::os::windows::ffi::OsStrExt;
use std::path::{Path, PathBuf};
use std::time::Duration;
use windows::core::{Interface, PCWSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::System::Com::{
    CoCreateInstance, IPersistFile, CLSCTX_INPROC_SERVER, STGM_READ,
};
use windows::Win32::UI::Shell::{IShellLinkW, ShellLink, SLGP_RAWPATH, SLR_NOUPDATE, SLR_NO_UI};

/// Deletes Recent-folder `.lnk` files whose resolved target equals `target`.
pub(crate) fn delete_recent_links_for_target(
    target: &str,
    timeout: Duration,
) -> WincentResult<Vec<PathBuf>> {
    let recent_folder = PathBuf::from(get_windows_recent_folder()?);
    delete_recent_links_for_target_in(&recent_folder, target, timeout, resolve_lnk_target)
}

fn delete_recent_links_for_target_in<F>(
    recent_folder: &Path,
    target: &str,
    timeout: Duration,
    mut resolver: F,
) -> WincentResult<Vec<PathBuf>>
where
    F: FnMut(&Path, Duration) -> Option<String>,
{
    let mut deleted = Vec::new();

    for entry in fs::read_dir(recent_folder).map_err(WincentError::Io)? {
        let entry = entry.map_err(WincentError::Io)?;
        let path = entry.path();
        let file_type = entry.file_type().map_err(WincentError::Io)?;
        if !file_type.is_file() || !is_lnk_file(&path) {
            continue;
        }

        let Some(resolved_target) = resolver(&path, timeout) else {
            continue;
        };

        if paths_equal(&resolved_target, target) {
            fs::remove_file(&path).map_err(WincentError::Io)?;
            deleted.push(path);
        }
    }

    Ok(deleted)
}

fn resolve_lnk_target(lnk_path: &Path, timeout: Duration) -> Option<String> {
    let path = lnk_path.to_path_buf();
    crate::com_thread::run_on_sta_thread(move || resolve_lnk_target_on_sta(&path), timeout).ok()
}

fn resolve_lnk_target_on_sta(lnk_path: &Path) -> WincentResult<String> {
    let mut link_path = os_str_to_wide_null(lnk_path.as_os_str());

    let shell_link: IShellLinkW = unsafe {
        CoCreateInstance(&ShellLink, None, CLSCTX_INPROC_SERVER)
    }
    .map_err(|err| WincentError::SystemError(format!("Failed to create ShellLink: {err}")))?;
    let persist_file: IPersistFile = shell_link
        .cast()
        .map_err(|err| WincentError::SystemError(format!("Failed to cast ShellLink: {err}")))?;

    unsafe {
        persist_file
            .Load(PCWSTR::from_raw(link_path.as_mut_ptr()), STGM_READ)
            .map_err(|err| {
                WincentError::SystemError(format!(
                    "Failed to load shortcut '{}': {err}",
                    lnk_path.display()
                ))
            })?;
        shell_link
            .Resolve(
                HWND(std::ptr::null_mut()),
                (SLR_NO_UI | SLR_NOUPDATE).0 as u32,
            )
            .map_err(|err| {
                WincentError::SystemError(format!(
                    "Failed to resolve shortcut '{}': {err}",
                    lnk_path.display()
                ))
            })?;
    }

    let mut target = vec![0u16; 32768];
    unsafe {
        shell_link
            .GetPath(&mut target, std::ptr::null_mut(), SLGP_RAWPATH.0 as u32)
            .map_err(|err| {
                WincentError::SystemError(format!(
                    "Failed to read shortcut target '{}': {err}",
                    lnk_path.display()
                ))
            })?;
    }

    let end = target
        .iter()
        .position(|&ch| ch == 0)
        .unwrap_or(target.len());
    if end == 0 {
        return Err(WincentError::SystemError(format!(
            "Shortcut '{}' does not contain a filesystem target path",
            lnk_path.display()
        )));
    }

    Ok(String::from_utf16_lossy(&target[..end]))
}

fn is_lnk_file(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.eq_ignore_ascii_case("lnk"))
        .unwrap_or(false)
}

fn os_str_to_wide_null(value: &std::ffi::OsStr) -> Vec<u16> {
    value.encode_wide().chain(std::iter::once(0)).collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;
    use tempfile::tempdir;

    #[test]
    fn lnk_extension_detection_is_case_insensitive() {
        assert!(is_lnk_file(Path::new("example.lnk")));
        assert!(is_lnk_file(Path::new("example.LNK")));
        assert!(!is_lnk_file(Path::new("example.txt")));
        assert!(!is_lnk_file(Path::new("example")));
    }

    #[test]
    fn deep_clean_deletes_only_matching_recent_links() -> WincentResult<()> {
        let dir = tempdir().map_err(WincentError::Io)?;
        let matching = dir.path().join("matching.lnk");
        let other = dir.path().join("other.lnk");
        let broken = dir.path().join("broken.lnk");
        let non_lnk = dir.path().join("matching.txt");
        fs::write(&matching, b"matching").map_err(WincentError::Io)?;
        fs::write(&other, b"other").map_err(WincentError::Io)?;
        fs::write(&broken, b"broken").map_err(WincentError::Io)?;
        fs::write(&non_lnk, b"not a shortcut").map_err(WincentError::Io)?;

        let mut targets = HashMap::new();
        targets.insert(matching.clone(), "C:\\Work\\Report.docx".to_string());
        targets.insert(other.clone(), "C:\\Work\\Other.docx".to_string());

        let deleted = delete_recent_links_for_target_in(
            dir.path(),
            "c:/work/report.docx",
            Duration::from_secs(1),
            |path, _| targets.get(path).cloned(),
        )?;

        assert_eq!(deleted, vec![matching.clone()]);
        assert!(!matching.exists());
        assert!(other.exists());
        assert!(broken.exists());
        assert!(non_lnk.exists());
        Ok(())
    }
}
