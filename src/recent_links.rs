//! Helpers for cleaning Windows Recent folder shortcuts.

use crate::error::WincentError;
use crate::utils::{get_windows_recent_folder, os_str_to_wide_null, paths_equal};
use crate::WincentResult;
use std::fs;
use std::path::{Path, PathBuf};
use std::time::Duration;
use windows::core::{Interface, PCWSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::Storage::FileSystem::{FILE_ATTRIBUTE_DIRECTORY, WIN32_FIND_DATAW};
use windows::Win32::System::Com::{
    CoCreateInstance, IPersistFile, CLSCTX_INPROC_SERVER, STGM_READ,
};
use windows::Win32::UI::Shell::{IShellLinkW, ShellLink, SLGP_RAWPATH, SLR_NOUPDATE, SLR_NO_UI};

/// Resolved shortcut target plus best-effort target type information.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct LnkResolution {
    pub(crate) path: String,
    pub(crate) is_dir: Option<bool>,
}

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
    F: FnMut(&Path, Duration) -> WincentResult<Option<String>>,
{
    let mut deleted = Vec::new();

    for entry in fs::read_dir(recent_folder).map_err(WincentError::Io)? {
        let entry = entry.map_err(WincentError::Io)?;
        let path = entry.path();
        let file_type = entry.file_type().map_err(WincentError::Io)?;
        if !file_type.is_file() || !is_lnk_file(&path) {
            continue;
        }

        let Some(resolved_target) = resolver(&path, timeout)? else {
            continue;
        };

        if paths_equal(&resolved_target, target) {
            fs::remove_file(&path).map_err(WincentError::Io)?;
            deleted.push(path);
        }
    }

    Ok(deleted)
}

pub(crate) fn recent_lnk_paths(recent_folder: &Path) -> WincentResult<Vec<PathBuf>> {
    let mut paths = Vec::new();

    for entry in fs::read_dir(recent_folder).map_err(WincentError::Io)? {
        let entry = entry.map_err(WincentError::Io)?;
        let path = entry.path();
        let file_type = entry.file_type().map_err(WincentError::Io)?;
        if file_type.is_file() && is_lnk_file(&path) {
            paths.push(path);
        }
    }

    paths.sort();
    Ok(paths)
}

pub(crate) fn resolve_lnk_target(
    lnk_path: &Path,
    timeout: Duration,
) -> WincentResult<Option<String>> {
    let path = lnk_path.to_path_buf();
    crate::com_thread::run_on_sta_thread(move || resolve_lnk_target_on_sta(&path), timeout)
}

fn resolve_lnk_target_on_sta(lnk_path: &Path) -> WincentResult<Option<String>> {
    Ok(resolve_lnk_with_type_on_sta(lnk_path)?.map(|resolution| resolution.path))
}

pub(crate) fn resolve_lnk_with_type(
    lnk_path: &Path,
    timeout: Duration,
) -> WincentResult<Option<LnkResolution>> {
    let path = lnk_path.to_path_buf();
    crate::com_thread::run_on_sta_thread(move || resolve_lnk_with_type_on_sta(&path), timeout)
}

fn resolve_lnk_with_type_on_sta(lnk_path: &Path) -> WincentResult<Option<LnkResolution>> {
    let mut link_path = os_str_to_wide_null(lnk_path.as_os_str());

    let shell_link: IShellLinkW = unsafe {
        CoCreateInstance(&ShellLink, None, CLSCTX_INPROC_SERVER)
    }
    .map_err(|err| WincentError::SystemError(format!("Failed to create ShellLink: {err}")))?;
    let persist_file: IPersistFile = shell_link
        .cast()
        .map_err(|err| WincentError::SystemError(format!("Failed to cast ShellLink: {err}")))?;

    unsafe {
        if persist_file
            .Load(PCWSTR::from_raw(link_path.as_mut_ptr()), STGM_READ)
            .is_err()
        {
            return Ok(None);
        }
        if shell_link
            .Resolve(
                HWND(std::ptr::null_mut()),
                (SLR_NO_UI | SLR_NOUPDATE).0 as u32,
            )
            .is_err()
        {
            return Ok(None);
        }
    }

    let mut target = vec![0u16; 32768];
    let mut find_data = WIN32_FIND_DATAW::default();
    unsafe {
        if shell_link
            .GetPath(&mut target, &mut find_data, SLGP_RAWPATH.0 as u32)
            .is_err()
        {
            return Ok(None);
        }
    }

    let end = target
        .iter()
        .position(|&ch| ch == 0)
        .unwrap_or(target.len());
    if end == 0 {
        return Ok(None);
    }

    let path = String::from_utf16_lossy(&target[..end]);
    let is_dir = lnk_target_is_dir(&path, &find_data);

    Ok(Some(LnkResolution { path, is_dir }))
}

fn lnk_target_is_dir(path: &str, find_data: &WIN32_FIND_DATAW) -> Option<bool> {
    let attributes = find_data.dwFileAttributes;
    if attributes != 0 {
        return Some((attributes & FILE_ATTRIBUTE_DIRECTORY.0) != 0);
    }

    // When Shell Link does not populate WIN32_FIND_DATAW (for example an
    // unavailable UNC/network target), filesystem metadata may fail with
    // access, connectivity, or other errors. Treat those as unknown so restore
    // flows apply their "unknown target" deletion policy.
    match fs::metadata(path) {
        Ok(metadata) => Some(metadata.is_dir()),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => None,
        Err(_) => None,
    }
}

pub(crate) fn is_lnk_file(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.eq_ignore_ascii_case("lnk"))
        .unwrap_or(false)
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
            |path, _| Ok(targets.get(path).cloned()),
        )?;

        assert_eq!(deleted, vec![matching.clone()]);
        assert!(!matching.exists());
        assert!(other.exists());
        assert!(broken.exists());
        assert!(non_lnk.exists());
        Ok(())
    }

    #[test]
    fn deep_clean_propagates_resolver_errors() -> WincentResult<()> {
        let dir = tempdir().map_err(WincentError::Io)?;
        let link = dir.path().join("matching.lnk");
        fs::write(&link, b"matching").map_err(WincentError::Io)?;

        let error = delete_recent_links_for_target_in(
            dir.path(),
            "c:/work/report.docx",
            Duration::from_secs(1),
            |_, _| Err(WincentError::Timeout("resolver timed out".to_string())),
        )
        .unwrap_err();

        assert!(link.exists());
        assert!(matches!(
            error,
            WincentError::Timeout(message) if message == "resolver timed out"
        ));
        Ok(())
    }

    #[test]
    fn resolve_lnk_target_preserves_sta_thread_errors() {
        let error = resolve_lnk_target(Path::new("missing.lnk"), Duration::ZERO).unwrap_err();

        assert!(matches!(error, WincentError::InvalidArgument(_)));
    }

    #[test]
    fn recent_lnk_paths_returns_only_lnk_files_sorted() -> WincentResult<()> {
        let dir = tempdir().map_err(WincentError::Io)?;
        let a = dir.path().join("a.lnk");
        let b = dir.path().join("b.LNK");
        let txt = dir.path().join("c.txt");
        fs::write(&b, b"b").map_err(WincentError::Io)?;
        fs::write(&txt, b"txt").map_err(WincentError::Io)?;
        fs::write(&a, b"a").map_err(WincentError::Io)?;

        assert_eq!(recent_lnk_paths(dir.path())?, vec![a, b]);
        Ok(())
    }
}
