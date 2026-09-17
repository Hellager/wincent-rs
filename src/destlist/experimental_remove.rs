//! Experimental removal by forcing Explorer to rebuild `.automaticDestinations-ms`.
//!
//! This module implements an observed, risky strategy:
//! 1. Delete the matching Explorer `.automaticDestinations-ms` backing file.
//! 2. Delete matching `.lnk` files from the Windows Recent folder.
//! 3. Wait briefly for Explorer to rebuild it.
//! 4. Parse the rebuilt file and verify the target entries are gone.
//!
//! # Experimental and risky
//!
//! This touches Shell-maintained files directly. Windows may change this behavior,
//! Explorer may rebuild asynchronously, and deleting the backing file can temporarily
//! affect Quick Access state. It may also delete matching `.lnk` files from the
//! Windows Recent folder. Callers should treat this as best-effort experimental
//! functionality with a weaker compatibility contract than the stable parser and
//! query APIs, and keep their own backups when the data matters. The operation is
//! not transactional; if the process is interrupted after deletion starts, no rollback
//! is attempted.
//!
//! # Observed limitation
//!
//! Deleting a matching `.lnk` from the Windows Recent folder and then deleting the
//! Recent Files `.automaticDestinations-ms` file does not reliably remove the matching
//! entry after Explorer rebuilds the DestList. A Home or Quick Access navigation can
//! recreate the backing file, but Explorer's in-memory cache or other Shell state may
//! write the removed entry back into the rebuilt file. Therefore, a successful rebuild
//! is not evidence that a targeted removal succeeded.
//!
//! Do not use this module for selective Recent Files removal. At present, the supported
//! recovery path is to reset the affected section to its default state with
//! [`crate::QuickAccessManager::restore_recent_files_defaults`] and
//! [`crate::RestoreDefaultsOptions`], or reset all Quick Access sections with
//! [`crate::QuickAccessManager::restore_defaults`] and [`crate::QuickAccess::All`].

use std::fs;
use std::path::{Path, PathBuf};
use std::thread;
use std::time::{Duration, Instant};

use crate::error::WincentError;
use crate::recent_links::{is_lnk_file, resolve_lnk_target};
use crate::utils::{get_windows_recent_folder, normalize_path_lightweight};
use crate::WincentResult;

use super::parser::{
    frequent_folders_dest_path, parse_file, recent_files_dest_path, AutomaticDestinations,
    DestListEntry,
};

const SHORTCUT_RESOLVE_TIMEOUT: Duration = Duration::from_secs(10);
const REBUILD_POLL_TIMEOUT: Duration = Duration::from_secs(5);

/// Explorer automatic destination file family to modify.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub(crate) enum AutomaticDestinationsKind {
    /// Recent Files automatic destination.
    RecentFiles,
    /// Frequent Folders automatic destination.
    FrequentFolders,
}

/// Options for the experimental remove-and-rebuild flow.
///
/// The delay is only the initial grace period after deleting Explorer's backing
/// file. The implementation still polls for a rebuilt file afterwards.
///
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ExperimentalRemoveOptions {
    /// Initial grace delay after deleting the `.automaticDestinations-ms` file
    /// before polling for the rebuilt file.
    ///
    /// This is not the total rebuild timeout. After this delay, the implementation
    /// still polls for the rebuilt destination file for up to 5 seconds. Increase
    /// this value on slow or busy systems to give Explorer a quieter rebuild window
    /// before parsing begins.
    rebuild_delay: Duration,
}

impl Default for ExperimentalRemoveOptions {
    fn default() -> Self {
        Self {
            rebuild_delay: Duration::from_millis(500),
        }
    }
}

impl ExperimentalRemoveOptions {
    /// Creates default experimental remove options.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Initial grace delay after deleting the `.automaticDestinations-ms` file.
    #[must_use]
    pub fn rebuild_delay(&self) -> Duration {
        self.rebuild_delay
    }

    /// Sets the initial grace delay after deleting the `.automaticDestinations-ms` file.
    #[must_use]
    pub fn with_rebuild_delay(mut self, rebuild_delay: Duration) -> Self {
        self.rebuild_delay = rebuild_delay;
        self
    }
}

/// Result of the experimental remove-and-rebuild flow.
///
/// A successful function call means the delete-and-rebuild sequence completed
/// without an immediate API error. Check [`ExperimentalRemoveReport::success`]
/// to learn whether the requested entries were absent after Explorer rebuilt
/// the backing file. Failures after the backing file was deleted are captured in
/// [`ExperimentalRemoveReport::post_delete_error`] so callers can inspect the
/// operation's partial progress.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ExperimentalRemoveReport {
    /// `.lnk` files deleted from the Recent folder.
    deleted_lnk_paths: Vec<PathBuf>,
    /// Requested target paths that had no matching `.lnk` file.
    missing_lnk_target_paths: Vec<String>,
    /// Whether the backing automatic destination file was deleted.
    dest_deleted: bool,
    /// Whether Explorer rebuilt the automatic destination file during polling.
    rebuilt: bool,
    /// Time spent waiting until the rebuilt file could be parsed.
    rebuild_parse_elapsed: Option<Duration>,
    /// Last parse error observed while waiting for the rebuilt file.
    rebuild_parse_error: Option<String>,
    /// Error observed after the backing destination file was deleted.
    post_delete_error: Option<String>,
    /// Matching paths still present after Explorer rebuilt the file.
    remaining_paths_after_rebuild: Vec<String>,
    /// Whether all requested entries were absent after rebuild.
    success: bool,
}

impl ExperimentalRemoveReport {
    /// `.lnk` files deleted from the Recent folder.
    #[must_use]
    pub fn deleted_lnk_paths(&self) -> &[PathBuf] {
        &self.deleted_lnk_paths
    }

    /// Requested target paths that had no matching `.lnk` file.
    #[must_use]
    pub fn missing_lnk_target_paths(&self) -> &[String] {
        &self.missing_lnk_target_paths
    }

    /// Whether the backing automatic destination file was deleted.
    #[must_use]
    pub fn dest_deleted(&self) -> bool {
        self.dest_deleted
    }

    /// Whether Explorer rebuilt the automatic destination file during polling.
    #[must_use]
    pub fn rebuilt(&self) -> bool {
        self.rebuilt
    }

    /// Time spent waiting until the rebuilt file could be parsed.
    #[must_use]
    pub fn rebuild_parse_elapsed(&self) -> Option<Duration> {
        self.rebuild_parse_elapsed
    }

    /// Last parse error observed while waiting for the rebuilt file.
    #[must_use]
    pub fn rebuild_parse_error(&self) -> Option<&str> {
        self.rebuild_parse_error.as_deref()
    }

    /// Error observed after the backing destination file was deleted.
    ///
    /// This is `Some` when shortcut cleanup, Explorer refresh, or rebuild
    /// checking failed after the operation had already deleted the backing file.
    /// A value of `None` does not guarantee success: `rebuilt() == false` with
    /// `post_delete_error() == None` means Explorer did not rebuild the backing
    /// file before the polling timeout.
    #[must_use]
    pub fn post_delete_error(&self) -> Option<&str> {
        self.post_delete_error.as_deref()
    }

    /// Matching paths still present after Explorer rebuilt the file.
    #[must_use]
    pub fn remaining_paths_after_rebuild(&self) -> &[String] {
        &self.remaining_paths_after_rebuild
    }

    /// Whether all requested entries were absent after rebuild.
    #[must_use]
    pub fn success(&self) -> bool {
        self.success
    }
}

/// Experimentally removes DestList entries by deleting their Recent `.lnk`
/// files, deleting the backing `.automaticDestinations-ms` file, and checking
/// Explorer's rebuilt file.
///
/// # Experimental and risky
///
/// This function deletes Shell-maintained files. It should be used only when the
/// caller accepts that Windows may rebuild Quick Access state asynchronously or
/// differently across versions. The delete sequence is ordered to remove the
/// backing `.automaticDestinations-ms` file before deleting matching `.lnk`
/// files, avoiding a state where `.lnk` files are deleted but the backing file
/// deletion fails. This still is not atomic and has no rollback if the process
/// is interrupted after deletion starts.
///
/// # Errors
///
/// Returns [`WincentError::InvalidArgument`] when `target_paths` is empty.
/// Returns I/O errors when the backing destination file cannot be removed.
/// Returns DestList parse errors if the backing file cannot be parsed before
/// deletion. After the backing file is deleted, shortcut cleanup, Explorer
/// refresh, and rebuild-check errors are captured in the returned report's
/// [`ExperimentalRemoveReport::post_delete_error`] instead of being returned as
/// `Err`.
///
/// `post_delete_error() == None` and `rebuilt() == false` means Explorer did
/// not rebuild the backing file before the polling timeout. It is distinct from
/// a post-delete operation error, and callers should inspect `success()`,
/// `rebuilt()`, `post_delete_error()`, and `remaining_paths_after_rebuild()`
/// together.
///
pub(crate) fn experimental_remove_entry_paths_by_rebuild<P: AsRef<Path>>(
    kind: AutomaticDestinationsKind,
    target_paths: &[P],
    options: ExperimentalRemoveOptions,
) -> WincentResult<ExperimentalRemoveReport> {
    experimental_remove_entry_paths_by_rebuild_with_resolver(
        kind,
        target_paths,
        options,
        resolve_lnk_target,
    )
}

fn experimental_remove_entry_paths_by_rebuild_with_resolver<P, F>(
    kind: AutomaticDestinationsKind,
    target_paths: &[P],
    options: ExperimentalRemoveOptions,
    resolver: F,
) -> WincentResult<ExperimentalRemoveReport>
where
    P: AsRef<Path>,
    F: FnMut(&Path, Duration) -> WincentResult<Option<String>>,
{
    if target_paths.is_empty() {
        return Err(WincentError::InvalidArgument(
            "target_paths must not be empty".to_string(),
        ));
    }

    let requested_paths: Vec<String> = target_paths
        .iter()
        .map(|path| path.as_ref().to_string_lossy().to_string())
        .collect();
    let recent_folder = PathBuf::from(get_windows_recent_folder()?);
    let dest_path = dest_path_for_kind(kind)?;

    parse_file_with_retries(&dest_path, Duration::from_secs(2))?;

    fs::remove_file(&dest_path).map_err(WincentError::Io)?;

    let base = ExperimentalRemoveBase {
        recent_folder,
        dest_path,
        requested_paths,
    };

    Ok(complete_after_destination_deleted(
        base,
        options,
        resolver,
        crate::utils::refresh_explorer_window,
        wait_for_rebuilt_dest,
    ))
}

/// Variant that accepts parsed DestList entries directly.
///
/// # Experimental and risky
///
/// This has the same risks as [`experimental_remove_entry_paths_by_rebuild`].
///
/// # Errors
///
/// Returns the same errors as [`experimental_remove_entry_paths_by_rebuild`].
pub(crate) fn experimental_remove_entries_by_rebuild(
    kind: AutomaticDestinationsKind,
    entries: &[DestListEntry],
    options: ExperimentalRemoveOptions,
) -> WincentResult<ExperimentalRemoveReport> {
    let paths: Vec<PathBuf> = entries
        .iter()
        .map(|entry| PathBuf::from(entry.path()))
        .collect();
    experimental_remove_entry_paths_by_rebuild(kind, &paths, options)
}

#[derive(Debug)]
struct ExperimentalRemoveBase {
    recent_folder: PathBuf,
    dest_path: PathBuf,
    requested_paths: Vec<String>,
}

impl ExperimentalRemoveBase {
    #[allow(clippy::too_many_arguments)]
    fn report(
        self,
        deleted_lnk_paths: Vec<PathBuf>,
        missing_lnk_target_paths: Vec<String>,
        rebuilt: bool,
        rebuild_parse_elapsed: Option<Duration>,
        rebuild_parse_error: Option<String>,
        post_delete_error: Option<String>,
        remaining_paths_after_rebuild: Vec<String>,
    ) -> ExperimentalRemoveReport {
        let success =
            post_delete_error.is_none() && rebuilt && remaining_paths_after_rebuild.is_empty();

        ExperimentalRemoveReport {
            deleted_lnk_paths,
            missing_lnk_target_paths,
            dest_deleted: true,
            rebuilt,
            rebuild_parse_elapsed,
            rebuild_parse_error,
            post_delete_error,
            remaining_paths_after_rebuild,
            success,
        }
    }
}

fn complete_after_destination_deleted<FResolve, FRefresh, FWait>(
    base: ExperimentalRemoveBase,
    options: ExperimentalRemoveOptions,
    resolver: FResolve,
    refresh_explorer: FRefresh,
    wait_for_rebuild: FWait,
) -> ExperimentalRemoveReport
where
    FResolve: FnMut(&Path, Duration) -> WincentResult<Option<String>>,
    FRefresh: FnOnce() -> WincentResult<()>,
    FWait: FnOnce(&Path, &[String], Duration) -> WincentResult<RebuiltDestWait>,
{
    let deleted_links = match delete_matching_recent_links(
        &base.recent_folder,
        &base.requested_paths,
        SHORTCUT_RESOLVE_TIMEOUT,
        resolver,
    ) {
        Ok(deleted_links) => deleted_links,
        Err(error) => {
            return base.report(
                deleted_lnk_paths(&error.deleted),
                Vec::new(),
                false,
                None,
                None,
                Some(error.source.to_string()),
                Vec::new(),
            );
        }
    };

    let missing_lnk_target_paths = missing_lnk_target_paths(&base.requested_paths, &deleted_links);
    let deleted_lnk_paths = deleted_lnk_paths(&deleted_links);

    if let Err(error) = refresh_explorer() {
        return base.report(
            deleted_lnk_paths,
            missing_lnk_target_paths,
            false,
            None,
            None,
            Some(error.to_string()),
            Vec::new(),
        );
    }

    thread::sleep(options.rebuild_delay());

    let check = match wait_for_rebuild(&base.dest_path, &base.requested_paths, REBUILD_POLL_TIMEOUT)
    {
        Ok(check) => check,
        Err(error) => {
            return base.report(
                deleted_lnk_paths,
                missing_lnk_target_paths,
                false,
                None,
                None,
                Some(error.to_string()),
                Vec::new(),
            );
        }
    };

    match check.rebuilt {
        Some(rebuilt) => base.report(
            deleted_lnk_paths,
            missing_lnk_target_paths,
            true,
            Some(rebuilt.elapsed),
            check.last_parse_error,
            None,
            rebuilt.remaining_paths,
        ),
        None => base.report(
            deleted_lnk_paths,
            missing_lnk_target_paths,
            false,
            None,
            check.last_parse_error,
            None,
            Vec::new(),
        ),
    }
}

fn parse_file_with_retries(path: &Path, timeout: Duration) -> WincentResult<AutomaticDestinations> {
    let started = Instant::now();

    loop {
        match parse_file(path) {
            Ok(parsed) => return Ok(parsed),
            Err(error) => {
                if started.elapsed() >= timeout {
                    return Err(error);
                }
                thread::sleep(Duration::from_millis(100));
            }
        }
    }
}

fn wait_for_rebuilt_dest(
    dest_path: &Path,
    requested_paths: &[String],
    timeout: Duration,
) -> WincentResult<RebuiltDestWait> {
    let started = Instant::now();
    let mut last_parse_error = None;

    loop {
        if dest_path.exists() {
            match parse_file(dest_path) {
                Ok(parsed) => {
                    return Ok(RebuiltDestWait {
                        rebuilt: Some(RebuiltDestCheck {
                            elapsed: started.elapsed(),
                            remaining_paths: matching_dest_paths(
                                parsed.dest_list().entries(),
                                requested_paths,
                            ),
                        }),
                        last_parse_error,
                    });
                }
                Err(error) => {
                    last_parse_error = Some(error.to_string());
                }
            }
        }

        if started.elapsed() >= timeout {
            return Ok(RebuiltDestWait {
                rebuilt: None,
                last_parse_error,
            });
        }

        thread::sleep(Duration::from_millis(100));
    }
}

#[derive(Debug)]
struct RebuiltDestWait {
    rebuilt: Option<RebuiltDestCheck>,
    last_parse_error: Option<String>,
}

#[derive(Debug)]
struct RebuiltDestCheck {
    elapsed: Duration,
    remaining_paths: Vec<String>,
}

fn dest_path_for_kind(kind: AutomaticDestinationsKind) -> WincentResult<PathBuf> {
    match kind {
        AutomaticDestinationsKind::RecentFiles => recent_files_dest_path(),
        AutomaticDestinationsKind::FrequentFolders => frequent_folders_dest_path(),
    }
}

fn matching_dest_paths(entries: &[DestListEntry], target_paths: &[String]) -> Vec<String> {
    entries
        .iter()
        .filter(|entry| {
            target_paths
                .iter()
                .any(|target| paths_equal_without_io(entry.path(), target))
        })
        .map(|entry| entry.path().to_string())
        .collect()
}

fn paths_equal_without_io(left: &str, right: &str) -> bool {
    normalize_path_lightweight(left) == normalize_path_lightweight(right)
}

fn missing_lnk_target_paths(
    requested_paths: &[String],
    deleted_links: &[DeletedRecentLink],
) -> Vec<String> {
    requested_paths
        .iter()
        .filter(|target| {
            !deleted_links
                .iter()
                .any(|link| paths_equal_without_io(&link.target_path, target))
        })
        .cloned()
        .collect()
}

fn deleted_lnk_paths(deleted_links: &[DeletedRecentLink]) -> Vec<PathBuf> {
    deleted_links
        .iter()
        .map(|link| link.lnk_path.clone())
        .collect()
}

fn delete_matching_recent_links<F>(
    recent_folder: &Path,
    target_paths: &[String],
    timeout: Duration,
    resolver: F,
) -> Result<Vec<DeletedRecentLink>, RecentLinkCleanupError>
where
    F: FnMut(&Path, Duration) -> WincentResult<Option<String>>,
{
    delete_matching_recent_links_with_remove(
        recent_folder,
        target_paths,
        timeout,
        resolver,
        |path| fs::remove_file(path).map_err(WincentError::Io),
    )
}

fn delete_matching_recent_links_with_remove<FResolve, FRemove>(
    recent_folder: &Path,
    target_paths: &[String],
    timeout: Duration,
    mut resolver: FResolve,
    mut remove_file: FRemove,
) -> Result<Vec<DeletedRecentLink>, RecentLinkCleanupError>
where
    FResolve: FnMut(&Path, Duration) -> WincentResult<Option<String>>,
    FRemove: FnMut(&Path) -> WincentResult<()>,
{
    let mut deleted = Vec::new();

    let entries = match fs::read_dir(recent_folder).map_err(WincentError::Io) {
        Ok(entries) => entries,
        Err(source) => {
            return Err(RecentLinkCleanupError { deleted, source });
        }
    };

    for entry in entries {
        let entry = match entry.map_err(WincentError::Io) {
            Ok(entry) => entry,
            Err(source) => {
                return Err(RecentLinkCleanupError { deleted, source });
            }
        };
        let path = entry.path();
        let file_type = match entry.file_type().map_err(WincentError::Io) {
            Ok(file_type) => file_type,
            Err(source) => {
                return Err(RecentLinkCleanupError { deleted, source });
            }
        };
        if !file_type.is_file() || !is_lnk_file(&path) {
            continue;
        }

        let target = match resolver(&path, timeout) {
            Ok(Some(target)) => target,
            Ok(None) => continue,
            Err(source) => {
                return Err(RecentLinkCleanupError { deleted, source });
            }
        };

        if !target_paths
            .iter()
            .any(|requested| paths_equal_without_io(&target, requested))
        {
            continue;
        }

        if let Err(source) = remove_file(&path) {
            return Err(RecentLinkCleanupError { deleted, source });
        }
        deleted.push(DeletedRecentLink {
            lnk_path: path,
            target_path: target,
        });
    }

    Ok(deleted)
}

#[derive(Debug)]
struct DeletedRecentLink {
    lnk_path: PathBuf,
    target_path: String,
}

#[derive(Debug)]
struct RecentLinkCleanupError {
    deleted: Vec<DeletedRecentLink>,
    source: WincentError,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::manager::QuickAccessManager;
    use crate::test_utils::{cleanup_test_env, create_test_file, setup_test_env};
    use crate::{AddOptions, QuickAccess, RemoveOptions};
    use std::env;
    use std::io::{self, Write};
    use std::process::Command;
    use tempfile::tempdir;
    use windows::core::{w, HSTRING};
    use windows::Win32::Foundation::RECT;
    use windows::Win32::System::Com::{
        CoCreateInstance, CoTaskMemFree, CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER,
    };
    use windows::Win32::UI::Shell::Common::IObjectArray;
    use windows::Win32::UI::Shell::{
        ApplicationDocumentLists, ExplorerBrowser, FOLDERID_Desktop, IApplicationDocumentLists,
        IExplorerBrowser, SHChangeNotify, SHGetKnownFolderIDList, SHParseDisplayName,
        ADLT_FREQUENT, ADLT_RECENT, EBO_NOBORDER, EBO_NOPERSISTVIEWSTATE, EBO_NOTRAVELLOG,
        FOLDERSETTINGS, FVM_DETAILS, KF_FLAG_DEFAULT, SBSP_ABSOLUTE, SHCNE_ASSOCCHANGED,
        SHCNF_IDLIST,
    };
    use windows::Win32::UI::WindowsAndMessaging::{
        CreateWindowExW, DestroyWindow, DispatchMessageW, IsWindowVisible, PeekMessageW,
        TranslateMessage, MSG, PM_REMOVE, WINDOW_EX_STYLE, WS_POPUP,
    };

    const HIDDEN_EXPLORER_BROWSER_HOME_NAMESPACE: &str =
        "shell:::{f874310e-b6b7-47dc-bc84-b9e6b38f5903}";
    const HIDDEN_EXPLORER_BROWSER_QUICK_ACCESS_NAMESPACE: &str =
        "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}";

    fn report_base_for_tests(recent_folder: &Path) -> ExperimentalRemoveBase {
        ExperimentalRemoveBase {
            recent_folder: recent_folder.to_path_buf(),
            dest_path: recent_folder.join("automaticDestinations-ms"),
            requested_paths: vec!["C:\\Work\\Report.docx".to_string()],
        }
    }

    fn parse_dest_for_destructive_test(dest_path: &Path) -> WincentResult<AutomaticDestinations> {
        match parse_file_with_retries(dest_path, Duration::from_secs(2)) {
            Ok(parsed) => Ok(parsed),
            Err(error) => {
                println!(
                    "initial parse failed: {error}; refreshing Quick Access windows and retrying"
                );
                crate::utils::refresh_explorer_window()?;
                thread::sleep(Duration::from_secs(1));
                parse_file_with_retries(dest_path, Duration::from_secs(5))
            }
        }
    }

    #[test]
    fn experimental_remove_rejects_empty_targets() {
        let result = experimental_remove_entry_paths_by_rebuild::<PathBuf>(
            AutomaticDestinationsKind::RecentFiles,
            &[],
            ExperimentalRemoveOptions::default(),
        );

        assert!(matches!(result, Err(WincentError::InvalidArgument(_))));
    }

    #[test]
    fn frequent_folders_kind_uses_frequent_dest_path() -> WincentResult<()> {
        assert_eq!(
            dest_path_for_kind(AutomaticDestinationsKind::FrequentFolders)?,
            frequent_folders_dest_path()?
        );
        Ok(())
    }

    #[test]
    fn default_rebuild_delay_allows_explorer_grace_period() {
        assert_eq!(
            ExperimentalRemoveOptions::default().rebuild_delay(),
            Duration::from_millis(500)
        );
    }

    #[test]
    fn lnk_extension_detection_is_case_insensitive() {
        assert!(is_lnk_file(Path::new("example.lnk")));
        assert!(is_lnk_file(Path::new("example.LNK")));
        assert!(!is_lnk_file(Path::new("example.txt")));
    }

    #[test]
    fn post_delete_refresh_error_returns_unsuccessful_report() -> WincentResult<()> {
        let dir = tempdir().map_err(WincentError::Io)?;
        let matching = dir.path().join("matching.lnk");
        let other = dir.path().join("other.lnk");
        fs::write(&matching, b"matching").map_err(WincentError::Io)?;
        fs::write(&other, b"other").map_err(WincentError::Io)?;

        let matching_for_resolver = matching.clone();
        let report = complete_after_destination_deleted(
            report_base_for_tests(dir.path()),
            ExperimentalRemoveOptions::new().with_rebuild_delay(Duration::ZERO),
            move |path, timeout| {
                assert_eq!(timeout, SHORTCUT_RESOLVE_TIMEOUT);
                if path == matching_for_resolver {
                    Ok(Some("c:/work/report.docx".to_string()))
                } else {
                    Ok(Some("C:\\Work\\Other.docx".to_string()))
                }
            },
            || Err(WincentError::SystemError("refresh failed".to_string())),
            |_, _, _| panic!("wait should not run after refresh failure"),
        );

        assert!(report.dest_deleted());
        assert!(!report.success());
        assert!(!report.rebuilt());
        assert_eq!(
            report.post_delete_error(),
            Some("System error: refresh failed")
        );
        assert_eq!(report.deleted_lnk_paths(), std::slice::from_ref(&matching));
        assert!(report.missing_lnk_target_paths().is_empty());
        assert!(!matching.exists());
        assert!(other.exists());
        Ok(())
    }

    #[test]
    fn post_delete_resolver_error_returns_unsuccessful_report() -> WincentResult<()> {
        let dir = tempdir().map_err(WincentError::Io)?;
        let shortcut = dir.path().join("shortcut.lnk");
        fs::write(&shortcut, b"shortcut").map_err(WincentError::Io)?;

        let report = complete_after_destination_deleted(
            report_base_for_tests(dir.path()),
            ExperimentalRemoveOptions::new().with_rebuild_delay(Duration::ZERO),
            |_, _| Err(WincentError::Timeout("resolver timed out".to_string())),
            || panic!("refresh should not run after resolver failure"),
            |_, _, _| panic!("wait should not run after resolver failure"),
        );

        assert!(report.dest_deleted());
        assert!(!report.success());
        assert_eq!(
            report.post_delete_error(),
            Some("Operation timed out: resolver timed out")
        );
        assert!(report.deleted_lnk_paths().is_empty());
        assert!(report.missing_lnk_target_paths().is_empty());
        assert!(shortcut.exists());
        Ok(())
    }

    #[test]
    fn rebuild_timeout_without_operation_error_has_no_post_delete_error() -> WincentResult<()> {
        let dir = tempdir().map_err(WincentError::Io)?;

        let report = complete_after_destination_deleted(
            report_base_for_tests(dir.path()),
            ExperimentalRemoveOptions::new().with_rebuild_delay(Duration::ZERO),
            |_, _| Ok(None),
            || Ok(()),
            |_, _, timeout| {
                assert_eq!(timeout, REBUILD_POLL_TIMEOUT);
                Ok(RebuiltDestWait {
                    rebuilt: None,
                    last_parse_error: Some("still rebuilding".to_string()),
                })
            },
        );

        assert!(report.dest_deleted());
        assert!(!report.success());
        assert!(!report.rebuilt());
        assert_eq!(report.post_delete_error(), None);
        assert_eq!(report.rebuild_parse_error(), Some("still rebuilding"));
        assert!(report.remaining_paths_after_rebuild().is_empty());
        Ok(())
    }

    #[test]
    fn delete_matching_recent_links_deletes_only_matching_shortcuts() -> WincentResult<()> {
        let dir = tempdir().map_err(WincentError::Io)?;
        let matching = dir.path().join("matching.lnk");
        let other = dir.path().join("other.lnk");
        let non_lnk = dir.path().join("matching.txt");
        fs::write(&matching, b"matching").map_err(WincentError::Io)?;
        fs::write(&other, b"other").map_err(WincentError::Io)?;
        fs::write(&non_lnk, b"not a shortcut").map_err(WincentError::Io)?;

        let matching_for_resolver = matching.clone();
        let deleted = delete_matching_recent_links(
            dir.path(),
            &["c:/work/report.docx".to_string()],
            Duration::from_secs(7),
            move |path, timeout| {
                assert_eq!(timeout, Duration::from_secs(7));
                if path == matching_for_resolver {
                    Ok(Some("C:\\Work\\Report.docx".to_string()))
                } else {
                    Ok(Some("C:\\Work\\Other.docx".to_string()))
                }
            },
        )
        .map_err(|error| error.source)?;

        assert_eq!(deleted_lnk_paths(&deleted), vec![matching.clone()]);
        assert!(!matching.exists());
        assert!(other.exists());
        assert!(non_lnk.exists());
        Ok(())
    }

    #[test]
    fn delete_matching_recent_links_skips_unresolved_shortcuts() -> WincentResult<()> {
        let dir = tempdir().map_err(WincentError::Io)?;
        let broken = dir.path().join("broken.lnk");
        fs::write(&broken, b"broken").map_err(WincentError::Io)?;

        let deleted = delete_matching_recent_links(
            dir.path(),
            &["C:\\Work\\Report.docx".to_string()],
            Duration::from_secs(7),
            |_, _| Ok(None),
        )
        .map_err(|error| error.source)?;

        assert!(deleted.is_empty());
        assert!(broken.exists());
        Ok(())
    }

    #[test]
    fn delete_matching_recent_links_reports_remove_file_errors() -> WincentResult<()> {
        let dir = tempdir().map_err(WincentError::Io)?;
        let shortcut = dir.path().join("shortcut.lnk");
        fs::write(&shortcut, b"shortcut").map_err(WincentError::Io)?;

        let error = delete_matching_recent_links_with_remove(
            dir.path(),
            &["C:\\Work\\Report.docx".to_string()],
            Duration::from_secs(7),
            |_, _| Ok(Some("C:\\Work\\Report.docx".to_string())),
            |_| {
                Err(WincentError::SystemError(
                    "shortcut delete failed".to_string(),
                ))
            },
        )
        .unwrap_err();

        assert!(error.deleted.is_empty());
        assert_eq!(
            error.source.to_string(),
            "System error: shortcut delete failed"
        );
        assert!(shortcut.exists());
        Ok(())
    }

    #[test]
    #[ignore = "Destructive integration test; deletes a real Recent .lnk and the Explorer recent-files automaticDestinations file"]
    fn experimental_remove_last_recent_file_entry_rebuilds_without_entry() -> WincentResult<()> {
        let dest_path = recent_files_dest_path()?;
        let parsed = parse_dest_for_destructive_test(&dest_path)?;
        let Some(entry) = parsed.dest_list().entries().last().cloned() else {
            println!("No entries found in {}", dest_path.display());
            return Ok(());
        };

        println!("recent_files_dest_path={}", dest_path.display());
        println!(
            "removing last entry: id={:#x} stream={} path={}",
            entry.entry_id(),
            entry.stream_name(),
            entry.path()
        );

        let report = experimental_remove_entries_by_rebuild(
            AutomaticDestinationsKind::RecentFiles,
            std::slice::from_ref(&entry),
            ExperimentalRemoveOptions::default(),
        )?;

        println!("deleted_lnk_paths={:?}", report.deleted_lnk_paths());
        println!("dest_deleted={}", report.dest_deleted());
        println!("rebuilt={}", report.rebuilt());
        println!("rebuild_parse_elapsed={:?}", report.rebuild_parse_elapsed());
        println!("rebuild_parse_error={:?}", report.rebuild_parse_error());
        println!("post_delete_error={:?}", report.post_delete_error());
        println!(
            "remaining_paths_after_rebuild={:?}",
            report.remaining_paths_after_rebuild()
        );

        assert!(report.rebuilt(), "recent-files dest file was not rebuilt");
        assert!(
            report.remaining_paths_after_rebuild().is_empty(),
            "rebuilt recent-files dest still contains the removed entry"
        );
        assert!(report.success(), "experimental removal did not succeed");

        Ok(())
    }

    #[test]
    #[ignore = "Destructive integration test; set WINCENT_DESTLIST_REBUILD_TEST_KEYWORD and WINCENT_DESTLIST_REBUILD_TEST_CONFIRM=DELETE"]
    fn experimental_keyword_removal_rebuilds_without_affecting_unrelated_entries(
    ) -> WincentResult<()> {
        require_destructive_test_confirmation()?;
        let keyword = required_environment_variable("WINCENT_DESTLIST_REBUILD_TEST_KEYWORD")?;
        let kinds = configured_dest_kinds()?;
        let mut snapshots = Vec::new();

        for kind in kinds {
            if let Some(snapshot) = keyword_dest_list_snapshot(kind, &keyword)? {
                snapshots.push(snapshot);
            }
        }
        if snapshots.is_empty() {
            return Err(WincentError::InvalidArgument(format!(
                "no selected DestList entries contain keyword {keyword:?}"
            )));
        }

        for snapshot in snapshots {
            remove_keyword_entries_and_verify(snapshot)?;
        }

        Ok(())
    }

    #[test]
    #[ignore = "Destructive integration test; set WINCENT_DESTLIST_HOME_REBUILD_TEST_KEYWORD and WINCENT_DESTLIST_REBUILD_TEST_CONFIRM=DELETE"]
    fn experimental_recent_keyword_removal_rebuilds_via_hidden_home_navigation() -> WincentResult<()>
    {
        require_destructive_test_confirmation()?;
        let keyword = required_environment_variable("WINCENT_DESTLIST_HOME_REBUILD_TEST_KEYWORD")?;

        // Preflight the trigger before deleting recent shortcuts or the DestList.
        browse_home_with_hidden_explorer_browser()?;

        let snapshot =
            keyword_dest_list_snapshot(AutomaticDestinationsKind::RecentFiles, &keyword)?
                .ok_or_else(|| {
                    WincentError::InvalidArgument(format!(
                        "no recent DestList entries contain keyword {keyword:?}"
                    ))
                })?;
        remove_recent_keyword_entries_via_hidden_home_navigation(snapshot)
    }

    fn require_destructive_test_confirmation() -> WincentResult<()> {
        match env::var("WINCENT_DESTLIST_REBUILD_TEST_CONFIRM").as_deref() {
            Ok("DELETE") => Ok(()),
            _ => Err(WincentError::InvalidArgument(
                "set WINCENT_DESTLIST_REBUILD_TEST_CONFIRM=DELETE to run this destructive test"
                    .to_string(),
            )),
        }
    }

    fn required_environment_variable(name: &str) -> WincentResult<String> {
        let value = env::var(name).map_err(|_| {
            WincentError::InvalidArgument(format!("set {name} to a non-empty search keyword"))
        })?;

        if value.trim().is_empty() {
            return Err(WincentError::InvalidArgument(format!(
                "{name} must not be empty"
            )));
        }

        Ok(value)
    }

    fn configured_dest_kinds() -> WincentResult<Vec<AutomaticDestinationsKind>> {
        match env::var("WINCENT_DESTLIST_REBUILD_TEST_KIND")
            .unwrap_or_else(|_| "all".to_string())
            .as_str()
        {
            "recent" => Ok(vec![AutomaticDestinationsKind::RecentFiles]),
            "frequent" => Ok(vec![AutomaticDestinationsKind::FrequentFolders]),
            "all" => Ok(vec![
                AutomaticDestinationsKind::RecentFiles,
                AutomaticDestinationsKind::FrequentFolders,
            ]),
            value => Err(WincentError::InvalidArgument(format!(
                "WINCENT_DESTLIST_REBUILD_TEST_KIND must be recent, frequent, or all; got {value}"
            ))),
        }
    }

    fn kind_name(kind: AutomaticDestinationsKind) -> &'static str {
        match kind {
            AutomaticDestinationsKind::RecentFiles => "recent",
            AutomaticDestinationsKind::FrequentFolders => "frequent",
        }
    }

    fn dest_list_path_contains_keyword(path: &str, keyword: &str) -> bool {
        path.to_lowercase().contains(&keyword.to_lowercase())
    }

    #[test]
    fn dest_list_keyword_matching_is_case_insensitive() {
        assert!(dest_list_path_contains_keyword(
            "C:\\Temp\\Report.TXT",
            "temp"
        ));
        assert!(dest_list_path_contains_keyword(
            "C:\\Temp\\Report.TXT",
            "REPORT"
        ));
        assert!(dest_list_path_contains_keyword("C:\\Temp\\Report.TXT", ""));
        assert!(!dest_list_path_contains_keyword(
            "C:\\Temp\\Report.TXT",
            "archive"
        ));
    }

    struct KeywordDestListSnapshot {
        kind: AutomaticDestinationsKind,
        dest_path: PathBuf,
        matching_paths: Vec<String>,
        unrelated_paths: Vec<String>,
    }

    fn keyword_dest_list_snapshot(
        kind: AutomaticDestinationsKind,
        keyword: &str,
    ) -> WincentResult<Option<KeywordDestListSnapshot>> {
        let dest_path = dest_path_for_kind(kind)?;
        let initial = parse_dest_for_destructive_test(&dest_path)?;
        let matching_paths: Vec<String> = initial
            .dest_list()
            .entries()
            .iter()
            .filter(|entry| dest_list_path_contains_keyword(entry.path(), keyword))
            .map(|entry| entry.path().to_string())
            .collect();
        let unrelated_paths: Vec<String> = initial
            .dest_list()
            .entries()
            .iter()
            .filter(|entry| !dest_list_path_contains_keyword(entry.path(), keyword))
            .map(|entry| entry.path().to_string())
            .collect();

        if matching_paths.is_empty() {
            println!("{} keyword matches=0; skipping", kind_name(kind));
            return Ok(None);
        }
        if unrelated_paths.is_empty() {
            return Err(WincentError::InvalidArgument(format!(
                "keyword {keyword:?} would remove every {} DestList entry",
                kind_name(kind)
            )));
        }

        Ok(Some(KeywordDestListSnapshot {
            kind,
            dest_path,
            matching_paths,
            unrelated_paths,
        }))
    }

    fn remove_keyword_entries_and_verify(snapshot: KeywordDestListSnapshot) -> WincentResult<()> {
        let KeywordDestListSnapshot {
            kind,
            dest_path,
            matching_paths,
            unrelated_paths,
        } = snapshot;

        println!(
            "{} keyword matches={} unrelated={}",
            kind_name(kind),
            matching_paths.len(),
            unrelated_paths.len()
        );

        let recent_folder = PathBuf::from(get_windows_recent_folder()?);
        let deleted_links = delete_matching_recent_links(
            &recent_folder,
            &matching_paths,
            SHORTCUT_RESOLVE_TIMEOUT,
            resolve_lnk_target,
        )
        .map_err(|error| error.source)?;
        for deleted_link in &deleted_links {
            assert!(
                !deleted_link.lnk_path.exists(),
                "matching shortcut was not deleted: {}",
                deleted_link.lnk_path.display()
            );
        }

        fs::remove_file(&dest_path).map_err(WincentError::Io)?;
        crate::utils::refresh_explorer_window()?;
        thread::sleep(ExperimentalRemoveOptions::default().rebuild_delay());

        let rebuilt = wait_for_rebuilt_dest(&dest_path, &matching_paths, REBUILD_POLL_TIMEOUT)?;
        let rebuilt_check = rebuilt.rebuilt.ok_or_else(|| {
            WincentError::Timeout(format!(
                "{} DestList was not rebuilt within {}s",
                kind_name(kind),
                REBUILD_POLL_TIMEOUT.as_secs()
            ))
        })?;
        assert!(
            rebuilt_check.remaining_paths.is_empty(),
            "rebuilt {} DestList still contains keyword matches: {:?}",
            kind_name(kind),
            rebuilt_check.remaining_paths
        );

        let rebuilt = parse_file(&dest_path)?;
        let remaining_matches = matching_dest_paths(rebuilt.dest_list().entries(), &matching_paths);
        assert!(
            remaining_matches.is_empty(),
            "reparsed {} DestList still contains keyword matches: {remaining_matches:?}",
            kind_name(kind)
        );

        let missing_unrelated = missing_dest_paths(rebuilt.dest_list().entries(), &unrelated_paths);
        assert!(
            missing_unrelated.is_empty(),
            "reparsed {} DestList lost unrelated entries: {missing_unrelated:?}",
            kind_name(kind)
        );

        Ok(())
    }

    fn remove_recent_keyword_entries_via_hidden_home_navigation(
        snapshot: KeywordDestListSnapshot,
    ) -> WincentResult<()> {
        let KeywordDestListSnapshot {
            kind,
            dest_path,
            matching_paths,
            unrelated_paths,
        } = snapshot;
        assert_eq!(
            kind,
            AutomaticDestinationsKind::RecentFiles,
            "hidden Home navigation rebuilds the Recent DestList only"
        );

        println!(
            "recent hidden Home keyword matches={} unrelated={}",
            matching_paths.len(),
            unrelated_paths.len()
        );
        let expected_links = required_recent_links_for_target_paths(&matching_paths)?;
        let recent_folder = PathBuf::from(get_windows_recent_folder()?);
        let deleted_links = delete_matching_recent_links(
            &recent_folder,
            &matching_paths,
            SHORTCUT_RESOLVE_TIMEOUT,
            resolve_lnk_target,
        )
        .map_err(|error| error.source)?;
        let mut deleted_paths = deleted_lnk_paths(&deleted_links);
        deleted_paths.sort();
        deleted_paths.dedup();
        assert_eq!(
            deleted_paths, expected_links,
            "deleted Recent shortcuts differ from the preflight matches"
        );
        for deleted_path in &deleted_paths {
            assert!(
                !deleted_path.exists(),
                "matching Recent shortcut was not deleted: {}",
                deleted_path.display()
            );
            println!(
                "deleted matching Recent shortcut {}",
                deleted_path.display()
            );
        }
        for target_path in &matching_paths {
            let remaining_links = crate::find_windows_recent_links_for_target(target_path)?;
            assert!(
                remaining_links.is_empty(),
                "Recent shortcuts still match {target_path:?} after deletion: {remaining_links:?}"
            );
        }
        println!(
            "verified {} matching Recent shortcuts deleted; deleting DestList {}",
            deleted_paths.len(),
            dest_path.display()
        );

        fs::remove_file(&dest_path).map_err(WincentError::Io)?;
        assert!(
            !dest_path.exists(),
            "Recent DestList still exists after deletion: {}",
            dest_path.display()
        );
        println!("deleted Recent DestList {}", dest_path.display());
        pause_for_private_trace()?;
        browse_home_with_hidden_explorer_browser()?;

        let rebuilt = wait_for_rebuilt_dest(&dest_path, &matching_paths, REBUILD_POLL_TIMEOUT)?;
        let rebuilt_check = rebuilt.rebuilt.ok_or_else(|| {
            WincentError::Timeout(format!(
                "hidden Home navigation did not rebuild Recent DestList within {:?}; \
                 last parse error: {:?}",
                REBUILD_POLL_TIMEOUT, rebuilt.last_parse_error
            ))
        })?;
        assert!(
            rebuilt_check.remaining_paths.is_empty(),
            "rebuilt Recent DestList still contains keyword matches: {:?}",
            rebuilt_check.remaining_paths
        );

        let rebuilt = parse_file(&dest_path)?;
        let remaining_matches = matching_dest_paths(rebuilt.dest_list().entries(), &matching_paths);
        assert!(
            remaining_matches.is_empty(),
            "reparsed Recent DestList still contains keyword matches: {remaining_matches:?}"
        );
        let missing_unrelated = missing_dest_paths(rebuilt.dest_list().entries(), &unrelated_paths);
        println!(
            "reparsed Recent DestList missing unrelated entries count={} paths={missing_unrelated:?}",
            missing_unrelated.len()
        );

        Ok(())
    }

    fn pause_for_private_trace() -> WincentResult<()> {
        let Some(value) = env::var_os("WINCENT_DESTLIST_PRIVATE_TRACE_PAUSE_MS") else {
            return Ok(());
        };
        let Ok(milliseconds) = value.to_string_lossy().parse::<u64>() else {
            println!("ignoring invalid WINCENT_DESTLIST_PRIVATE_TRACE_PAUSE_MS={value:?}");
            return Ok(());
        };
        println!(
            "PRIVATE_TRACE_PAUSE pid={} milliseconds={milliseconds}",
            std::process::id()
        );
        io::stdout().flush().map_err(WincentError::Io)?;
        thread::sleep(Duration::from_millis(milliseconds));
        println!("PRIVATE_TRACE_RESUME pid={}", std::process::id());
        io::stdout().flush().map_err(WincentError::Io)
    }

    fn required_recent_links_for_target_paths(
        target_paths: &[String],
    ) -> WincentResult<Vec<PathBuf>> {
        let mut matching_links = Vec::new();
        let mut paths_without_links = Vec::new();

        for target_path in target_paths {
            let links = crate::find_windows_recent_links_for_target(target_path)?;
            println!(
                "Recent shortcut preflight target={target_path:?} matches={}",
                links.len()
            );
            for link in &links {
                println!("  matching shortcut: {}", link.display());
            }
            if links.is_empty() {
                paths_without_links.push(target_path.clone());
            }
            matching_links.extend(links);
        }

        if !paths_without_links.is_empty() {
            return Err(WincentError::InvalidArgument(format!(
                "matching DestList paths have no corresponding Recent shortcut; \
                 stopped before DestList deletion: {paths_without_links:?}"
            )));
        }

        matching_links.sort();
        matching_links.dedup();
        Ok(matching_links)
    }

    #[test]
    #[ignore = "Destructive integration test; set WINCENT_DESTLIST_REBUILD_TEST_CONFIRM=DELETE"]
    fn home_navigation_rebuilds_recent_destlist() -> WincentResult<()> {
        require_destructive_test_confirmation()?;

        // Preflight before deleting the backing file so an unavailable Home window leaves the
        // existing DestList untouched.
        let Some(preflight) =
            crate::explorer_window::navigate_recent_access_window_to_desktop_and_back()?
        else {
            return Err(WincentError::SystemError(
                "open an Explorer window at Home or Quick Access before running this test"
                    .to_string(),
            ));
        };
        println!(
            "Home navigation preflight completed: {:?} -> {:?} -> {:?}",
            preflight.original, preflight.after_desktop, preflight.restored
        );

        let dest_path = recent_files_dest_path()?;
        let baseline = parse_dest_for_destructive_test(&dest_path)?;
        println!(
            "recent baseline entries={} path={}",
            baseline.dest_list().entries().len(),
            dest_path.display()
        );
        delete_dest_file_for_exploration(&dest_path)?;

        let Some(navigation) =
            crate::explorer_window::navigate_recent_access_window_to_desktop_and_back()?
        else {
            return Err(WincentError::SystemError(
                "the Home or Quick Access Explorer window disappeared after the preflight"
                    .to_string(),
            ));
        };
        println!(
            "Home navigation verification completed: {:?} -> {:?} -> {:?}",
            navigation.original, navigation.after_desktop, navigation.restored
        );

        let rebuilt = wait_for_rebuilt_dest(&dest_path, &[], REBUILD_POLL_TIMEOUT)?;
        let check = rebuilt.rebuilt.ok_or_else(|| {
            WincentError::Timeout(format!(
                "Home navigation did not rebuild {} within {:?}; last parse error: {:?}",
                dest_path.display(),
                REBUILD_POLL_TIMEOUT,
                rebuilt.last_parse_error
            ))
        })?;
        println!(
            "Home navigation rebuilt {} after {:?}",
            dest_path.display(),
            check.elapsed
        );

        Ok(())
    }

    #[test]
    #[ignore = "Destructive integration test; set WINCENT_DESTLIST_REBUILD_TEST_CONFIRM=DELETE"]
    fn hidden_explorer_browser_home_navigation_rebuilds_recent_destlist() -> WincentResult<()> {
        require_destructive_test_confirmation()?;

        browse_home_with_hidden_explorer_browser()?;
        println!("hidden IExplorerBrowser Home navigation preflight completed");

        let dest_path = recent_files_dest_path()?;
        let baseline = parse_dest_for_destructive_test(&dest_path)?;
        println!(
            "recent baseline entries={} path={}",
            baseline.dest_list().entries().len(),
            dest_path.display()
        );
        delete_dest_file_for_exploration(&dest_path)?;

        browse_home_with_hidden_explorer_browser()?;
        println!("hidden IExplorerBrowser Home navigation verification completed");

        let rebuilt = wait_for_rebuilt_dest(&dest_path, &[], REBUILD_POLL_TIMEOUT)?;
        let check = rebuilt.rebuilt.ok_or_else(|| {
            WincentError::Timeout(format!(
                "hidden IExplorerBrowser Home navigation did not rebuild {} within {:?}; \
                 last parse error: {:?}",
                dest_path.display(),
                REBUILD_POLL_TIMEOUT,
                rebuilt.last_parse_error
            ))
        })?;
        println!(
            "hidden IExplorerBrowser Home navigation rebuilt {} after {:?}",
            dest_path.display(),
            check.elapsed
        );

        Ok(())
    }

    fn browse_home_with_hidden_explorer_browser() -> WincentResult<()> {
        crate::com_thread::run_on_sta_thread(
            || {
                // SAFETY: The STATIC class is a system window class. WS_POPUP without
                // WS_VISIBLE creates a hidden top-level host that is destroyed below.
                let host = unsafe {
                    CreateWindowExW(
                        WINDOW_EX_STYLE::default(),
                        w!("STATIC"),
                        w!("wincent-hidden-explorer-browser"),
                        WS_POPUP,
                        0,
                        0,
                        1,
                        1,
                        None,
                        None,
                        None,
                        None,
                    )
                }
                .map_err(|error| {
                    WincentError::SystemError(format!(
                        "failed to create hidden IExplorerBrowser host: {error}"
                    ))
                })?;

                // SAFETY: host was just created on this thread and remains valid until
                // DestroyWindow below.
                let browse_result = if unsafe { IsWindowVisible(host).as_bool() } {
                    Err(WincentError::SystemError(
                        "hidden IExplorerBrowser host was unexpectedly visible".to_string(),
                    ))
                } else {
                    browse_home_in_hidden_explorer_browser(host)
                };
                // SAFETY: host was created on this STA thread and has not been destroyed.
                let destroy_result = unsafe { DestroyWindow(host) }.map_err(|error| {
                    WincentError::SystemError(format!(
                        "failed to destroy hidden IExplorerBrowser host: {error}"
                    ))
                });

                match (browse_result, destroy_result) {
                    (Ok(()), Ok(())) => Ok(()),
                    (Err(error), Ok(())) | (Ok(()), Err(error)) => Err(error),
                    (Err(browse_error), Err(destroy_error)) => {
                        Err(WincentError::SystemError(format!(
                            "hidden IExplorerBrowser failed: {browse_error}; \
                             host cleanup also failed: {destroy_error}"
                        )))
                    }
                }
            },
            Duration::from_secs(10),
        )
    }

    fn browse_home_in_hidden_explorer_browser(
        host: windows::Win32::Foundation::HWND,
    ) -> WincentResult<()> {
        // SAFETY: This runs on the dedicated COM-initialized STA worker. The returned interface
        // is owned by the windows crate wrapper and is destroyed before its hidden host window.
        let browser: IExplorerBrowser =
            unsafe { CoCreateInstance(&ExplorerBrowser, None, CLSCTX_INPROC_SERVER) }.map_err(
                |error| {
                    WincentError::SystemError(format!("failed to create IExplorerBrowser: {error}"))
                },
            )?;
        let settings = FOLDERSETTINGS {
            ViewMode: FVM_DETAILS.0 as u32,
            fFlags: 0,
        };
        let rect = RECT {
            left: 0,
            top: 0,
            right: 1,
            bottom: 1,
        };

        // SAFETY: host is a valid hidden window on this STA thread; settings and rect remain
        // alive for the duration of the Initialize call.
        unsafe { browser.Initialize(host, &rect, Some(&settings)) }.map_err(|error| {
            WincentError::SystemError(format!("IExplorerBrowser::Initialize failed: {error}"))
        })?;

        let browse_result = browse_hidden_explorer_browser_to_home(&browser, host);
        // SAFETY: browser was initialized successfully above and Destroy is called once before
        // its host window is destroyed.
        let destroy_result = unsafe { browser.Destroy() }.map_err(|error| {
            WincentError::SystemError(format!("IExplorerBrowser::Destroy failed: {error}"))
        });

        match (browse_result, destroy_result) {
            (Ok(()), Ok(())) => Ok(()),
            (Err(error), Ok(())) | (Ok(()), Err(error)) => Err(error),
            (Err(browse_error), Err(destroy_error)) => Err(WincentError::SystemError(format!(
                "IExplorerBrowser browse failed: {browse_error}; destroy also failed: \
                 {destroy_error}"
            ))),
        }
    }

    fn browse_hidden_explorer_browser_to_home(
        browser: &IExplorerBrowser,
        host: windows::Win32::Foundation::HWND,
    ) -> WincentResult<()> {
        // SAFETY: browser is a valid initialized IExplorerBrowser. These options prevent
        // persistent view state and travel-log changes while leaving the hidden view navigable.
        unsafe { browser.SetOptions(EBO_NOBORDER | EBO_NOPERSISTVIEWSTATE | EBO_NOTRAVELLOG) }
            .map_err(|error| {
                WincentError::SystemError(format!("IExplorerBrowser::SetOptions failed: {error}"))
            })?;

        browse_hidden_explorer_browser_to_desktop(browser)?;
        pump_hidden_explorer_browser_messages(Duration::from_millis(200));

        let mut failures = Vec::new();
        for (name, namespace) in [
            (
                "Quick Access",
                HIDDEN_EXPLORER_BROWSER_QUICK_ACCESS_NAMESPACE,
            ),
            ("Home", HIDDEN_EXPLORER_BROWSER_HOME_NAMESPACE),
        ] {
            match browse_hidden_explorer_browser_to_namespace(browser, name, namespace) {
                Ok(()) => {
                    println!("hidden IExplorerBrowser navigated Desktop -> {name}");
                    pump_hidden_explorer_browser_messages(Duration::from_secs(1));
                    return ensure_hidden_explorer_browser_host_is_invisible(host);
                }
                Err(error) => failures.push(format!("{name}: {error}")),
            }
        }

        Err(WincentError::SystemError(format!(
            "hidden IExplorerBrowser could not navigate from Desktop to Quick Access or Home: {}",
            failures.join("; ")
        )))
    }

    fn browse_hidden_explorer_browser_to_desktop(browser: &IExplorerBrowser) -> WincentResult<()> {
        // SAFETY: the returned PIDL is allocated by Shell and released by
        // browse_hidden_explorer_browser_to_owned_pidl below.
        let desktop_pidl =
            unsafe { SHGetKnownFolderIDList(&FOLDERID_Desktop, KF_FLAG_DEFAULT.0 as u32, None) }
                .map_err(|error| {
                    WincentError::SystemError(format!(
                        "failed to get Desktop PIDL for hidden IExplorerBrowser: {error}"
                    ))
                })?;
        browse_hidden_explorer_browser_to_owned_pidl(browser, desktop_pidl, "Desktop")
    }

    fn browse_hidden_explorer_browser_to_namespace(
        browser: &IExplorerBrowser,
        name: &str,
        namespace: &str,
    ) -> WincentResult<()> {
        let mut namespace_pidl = std::ptr::null_mut();
        let namespace = HSTRING::from(namespace);
        // SAFETY: namespace remains alive through the call, namespace_pidl is valid writable
        // storage, and the successful PIDL is released by the owned-PIDL helper below.
        unsafe { SHParseDisplayName(&namespace, None, &mut namespace_pidl, 0, None) }.map_err(
            |error| {
                WincentError::SystemError(format!(
                    "failed to parse {name} PIDL for hidden IExplorerBrowser: {error}"
                ))
            },
        )?;
        browse_hidden_explorer_browser_to_owned_pidl(browser, namespace_pidl, name)
    }

    fn browse_hidden_explorer_browser_to_owned_pidl(
        browser: &IExplorerBrowser,
        pidl: *mut windows::Win32::UI::Shell::Common::ITEMIDLIST,
        name: &str,
    ) -> WincentResult<()> {
        // SAFETY: pidl is allocated by a Shell API and remains valid until it is released below.
        let browse_result =
            unsafe { browser.BrowseToIDList(pidl, SBSP_ABSOLUTE) }.map_err(|error| {
                WincentError::SystemError(format!(
                    "IExplorerBrowser::BrowseToIDList({name}) failed: {error}"
                ))
            });
        // SAFETY: pidl was allocated by a Shell API and is released exactly once here.
        unsafe {
            CoTaskMemFree(Some(pidl.cast()));
        }
        browse_result
    }

    fn ensure_hidden_explorer_browser_host_is_invisible(
        host: windows::Win32::Foundation::HWND,
    ) -> WincentResult<()> {
        // SAFETY: host remains valid and owned by this thread until the caller destroys it.
        if unsafe { IsWindowVisible(host).as_bool() } {
            return Err(WincentError::SystemError(
                "hidden IExplorerBrowser host became visible while browsing Home".to_string(),
            ));
        }
        Ok(())
    }

    fn pump_hidden_explorer_browser_messages(duration: Duration) {
        let started = Instant::now();
        while started.elapsed() < duration {
            let mut message = MSG::default();
            // SAFETY: message is valid writable storage. This STA thread owns the queue being
            // pumped and removes only messages from its own queue.
            while unsafe { PeekMessageW(&mut message, None, 0, 0, PM_REMOVE).as_bool() } {
                // SAFETY: message was obtained from PeekMessageW and is dispatched on the same
                // owning thread.
                unsafe {
                    let _ = TranslateMessage(&message);
                    let _ = DispatchMessageW(&message);
                }
            }
            thread::sleep(Duration::from_millis(20));
        }
    }

    #[test]
    #[ignore = "Destructive integration test; set WINCENT_DESTLIST_REBUILD_TEST_CONFIRM=DELETE"]
    fn explore_rebuild_triggers_then_add_item_rebuilds_destlist() -> WincentResult<()> {
        require_destructive_test_confirmation()?;

        for kind in configured_exploration_kinds()? {
            explore_rebuild_triggers(kind)?;
        }

        Ok(())
    }

    fn configured_exploration_kinds() -> WincentResult<Vec<AutomaticDestinationsKind>> {
        match env::var("WINCENT_DESTLIST_REBUILD_EXPLORATION_KIND")
            .unwrap_or_else(|_| "recent".to_string())
            .as_str()
        {
            "recent" => Ok(vec![AutomaticDestinationsKind::RecentFiles]),
            "frequent" => Ok(vec![AutomaticDestinationsKind::FrequentFolders]),
            "all" => Ok(vec![
                AutomaticDestinationsKind::RecentFiles,
                AutomaticDestinationsKind::FrequentFolders,
            ]),
            value => Err(WincentError::InvalidArgument(format!(
                "WINCENT_DESTLIST_REBUILD_EXPLORATION_KIND must be recent, frequent, or all; \
                 got {value}"
            ))),
        }
    }

    fn explore_rebuild_triggers(kind: AutomaticDestinationsKind) -> WincentResult<()> {
        let dest_path = dest_path_for_kind(kind)?;
        let baseline = parse_dest_for_destructive_test(&dest_path)?;
        println!(
            "{} baseline entries={} path={}",
            kind_name(kind),
            baseline.dest_list().entries().len(),
            dest_path.display()
        );

        explore_rebuild_trigger(kind, "SHChangeNotify", notify_shell_association_change)?;
        if kind == AutomaticDestinationsKind::RecentFiles {
            explore_rebuild_trigger(
                kind,
                "hidden IExplorerBrowser Desktop -> Quick Access navigation",
                browse_home_with_hidden_explorer_browser,
            )?;
        } else {
            println!(
                "{}: hidden IExplorerBrowser navigation skipped; it targets the Recent DestList",
                kind_name(kind)
            );
        }
        explore_rebuild_trigger(kind, "IApplicationDocumentLists::GetList", || {
            read_application_document_list(kind)
        })?;
        explore_rebuild_trigger(kind, "Explorer restart", restart_explorer_for_exploration)?;

        let seed_directory = setup_test_env()?;
        let seed_path = create_seed_path(kind, &seed_directory)?;
        let result = add_item_and_verify_rebuild(kind, &seed_path);
        let cleanup = cleanup_test_env(&seed_directory);
        match (result, cleanup) {
            (Ok(()), Ok(())) => Ok(()),
            (Err(error), Ok(())) | (Ok(()), Err(error)) => Err(error),
            (Err(operation_error), Err(cleanup_error)) => Err(WincentError::SystemError(format!(
                "{} add_item experiment failed: {operation_error}; seed cleanup also failed: \
                     {cleanup_error}",
                kind_name(kind)
            ))),
        }
    }

    fn explore_rebuild_trigger<F>(
        kind: AutomaticDestinationsKind,
        trigger_name: &str,
        trigger: F,
    ) -> WincentResult<()>
    where
        F: FnOnce() -> WincentResult<()>,
    {
        let dest_path = dest_path_for_kind(kind)?;
        delete_dest_file_for_exploration(&dest_path)?;
        println!("{}: running {trigger_name}", kind_name(kind));

        if let Err(error) = trigger() {
            println!(
                "{}: {trigger_name} could not run: {error}; recording as exploratory failure",
                kind_name(kind)
            );
            return Ok(());
        }

        let rebuilt = wait_for_rebuilt_dest(&dest_path, &[], REBUILD_POLL_TIMEOUT)?;
        match rebuilt.rebuilt {
            Some(check) => println!(
                "{}: {trigger_name} rebuilt {} after {:?}",
                kind_name(kind),
                dest_path.display(),
                check.elapsed
            ),
            None => println!(
                "{}: {trigger_name} did not rebuild within {:?}; last parse error: {:?}",
                kind_name(kind),
                REBUILD_POLL_TIMEOUT,
                rebuilt.last_parse_error
            ),
        }

        Ok(())
    }

    fn delete_dest_file_for_exploration(dest_path: &Path) -> WincentResult<()> {
        match fs::remove_file(dest_path) {
            Ok(()) => {
                println!("deleted DestList {}", dest_path.display());
                Ok(())
            }
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {
                println!("DestList was already absent: {}", dest_path.display());
                Ok(())
            }
            Err(error) => Err(WincentError::Io(error)),
        }
    }

    fn notify_shell_association_change() -> WincentResult<()> {
        // SAFETY: SHCNE_ASSOCCHANGED with SHCNF_IDLIST does not consume item pointers.
        unsafe {
            SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
        }
        Ok(())
    }

    fn read_application_document_list(kind: AutomaticDestinationsKind) -> WincentResult<()> {
        crate::com_thread::run_on_sta_thread(
            move || {
                // SAFETY: The dedicated worker initialized COM in STA mode and the wrapper owns
                // the returned interface pointer.
                let document_lists: IApplicationDocumentLists = unsafe {
                    CoCreateInstance(
                        &ApplicationDocumentLists,
                        None,
                        CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
                    )
                }
                .map_err(|error| {
                    WincentError::SystemError(format!(
                        "failed to create IApplicationDocumentLists: {error}"
                    ))
                })?;

                let app_id = HSTRING::from("Microsoft.Windows.Explorer");
                // SAFETY: document_lists is a valid COM interface and app_id remains alive for
                // the duration of this call.
                unsafe { document_lists.SetAppID(&app_id) }.map_err(|error| {
                    WincentError::SystemError(format!(
                        "IApplicationDocumentLists::SetAppID failed: {error}"
                    ))
                })?;

                let list_type = match kind {
                    AutomaticDestinationsKind::RecentFiles => ADLT_RECENT,
                    AutomaticDestinationsKind::FrequentFolders => ADLT_FREQUENT,
                };
                // SAFETY: document_lists is valid, and the returned object array is owned by
                // the windows crate wrapper.
                let entries: IObjectArray = unsafe { document_lists.GetList(list_type, 20) }
                    .map_err(|error| {
                        WincentError::SystemError(format!(
                            "IApplicationDocumentLists::GetList failed: {error}"
                        ))
                    })?;
                // SAFETY: entries is a valid IObjectArray returned by GetList.
                let count = unsafe { entries.GetCount() }.map_err(|error| {
                    WincentError::SystemError(format!(
                        "IApplicationDocumentLists result count failed: {error}"
                    ))
                })?;
                println!(
                    "{} IApplicationDocumentLists items={count}",
                    kind_name(kind)
                );
                Ok(())
            },
            Duration::from_secs(10),
        )
    }

    fn restart_explorer_for_exploration() -> WincentResult<()> {
        let status = Command::new("taskkill")
            .args(["/F", "/IM", "explorer.exe"])
            .status()
            .map_err(WincentError::Io)?;
        if !status.success() {
            return Err(WincentError::SystemError(format!(
                "taskkill explorer.exe exited with {status}"
            )));
        }

        Command::new("explorer.exe")
            .spawn()
            .map_err(WincentError::Io)?;
        thread::sleep(Duration::from_secs(1));
        Ok(())
    }

    fn create_seed_path(
        kind: AutomaticDestinationsKind,
        seed_directory: &Path,
    ) -> WincentResult<PathBuf> {
        match kind {
            AutomaticDestinationsKind::RecentFiles => {
                create_test_file(seed_directory, "destlist-rebuild-seed.txt", "wincent")
            }
            AutomaticDestinationsKind::FrequentFolders => {
                let seed_path = seed_directory.join("destlist-rebuild-seed-folder");
                fs::create_dir(&seed_path).map_err(WincentError::Io)?;
                Ok(seed_path)
            }
        }
    }

    fn add_item_and_verify_rebuild(
        kind: AutomaticDestinationsKind,
        seed_path: &Path,
    ) -> WincentResult<()> {
        let dest_path = dest_path_for_kind(kind)?;
        delete_dest_file_for_exploration(&dest_path)?;

        let qa_type = match kind {
            AutomaticDestinationsKind::RecentFiles => QuickAccess::RecentFiles,
            AutomaticDestinationsKind::FrequentFolders => QuickAccess::FrequentFolders,
        };
        let manager = QuickAccessManager::new();
        manager.add_item(seed_path, qa_type, AddOptions::new())?;

        let seed_path_string = seed_path.to_string_lossy().into_owned();
        let rebuild_result = wait_for_rebuilt_dest(
            &dest_path,
            std::slice::from_ref(&seed_path_string),
            REBUILD_POLL_TIMEOUT,
        );
        remove_added_seed(&manager, seed_path, qa_type)?;
        let rebuilt = rebuild_result?;
        let check = rebuilt.rebuilt.ok_or_else(|| {
            WincentError::Timeout(format!(
                "{} add_item did not rebuild {} within {:?}",
                kind_name(kind),
                dest_path.display(),
                REBUILD_POLL_TIMEOUT
            ))
        })?;
        assert!(
            !check.remaining_paths.is_empty(),
            "{} add_item rebuilt DestList but seed path was absent: {}",
            kind_name(kind),
            seed_path.display()
        );
        println!(
            "{} add_item rebuilt {} after {:?} and recorded seed {}",
            kind_name(kind),
            dest_path.display(),
            check.elapsed,
            seed_path.display()
        );

        Ok(())
    }

    fn remove_added_seed(
        manager: &QuickAccessManager,
        seed_path: &Path,
        qa_type: QuickAccess,
    ) -> WincentResult<()> {
        let removal_result = manager.remove_item_with_options(
            seed_path,
            qa_type,
            RemoveOptions::new()
                .deep_clean_recent_links()
                .refresh_explorer(),
        );
        let link_cleanup_result = remove_seed_recent_links(seed_path);

        match (removal_result, link_cleanup_result) {
            (Ok(()), Ok(())) => Ok(()),
            (Err(WincentError::NotInQuickAccess { .. }), Ok(())) => {
                println!(
                    "seed was not queryable through the missing DestList; \
                     removed any matching Recent shortcuts directly"
                );
                Ok(())
            }
            (Err(removal_error), Ok(())) => Err(removal_error),
            (Ok(()), Err(link_error)) => Err(link_error),
            (Err(removal_error), Err(link_error)) => Err(WincentError::SystemError(format!(
                "seed cleanup failed: manager removal: {removal_error}; \
                 Recent shortcut cleanup: {link_error}"
            ))),
        }
    }

    fn remove_seed_recent_links(seed_path: &Path) -> WincentResult<()> {
        let seed_path = seed_path.to_string_lossy();
        for lnk_path in crate::find_windows_recent_links_for_target(&seed_path)? {
            fs::remove_file(&lnk_path).map_err(WincentError::Io)?;
            println!("deleted seed Recent shortcut {}", lnk_path.display());
        }
        Ok(())
    }

    fn missing_dest_paths(entries: &[DestListEntry], expected_paths: &[String]) -> Vec<String> {
        expected_paths
            .iter()
            .filter(|expected| {
                !entries
                    .iter()
                    .any(|entry| paths_equal_without_io(entry.path(), expected))
            })
            .cloned()
            .collect()
    }
}
