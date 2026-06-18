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

use std::fs;
use std::path::{Path, PathBuf};
use std::thread;
use std::time::{Duration, Instant};

use crate::error::WincentError;
use crate::recent_links::{is_lnk_file, resolve_lnk_target};
use crate::utils::{get_windows_recent_folder, paths_equal};
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
pub enum AutomaticDestinationsKind {
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
/// # Examples
///
/// ```rust
/// use std::time::Duration;
/// use wincent::destlist::ExperimentalRemoveOptions;
///
/// let options = ExperimentalRemoveOptions::new()
///     .with_rebuild_delay(Duration::from_secs(1));
///
/// assert_eq!(options.rebuild_delay(), Duration::from_secs(1));
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExperimentalRemoveOptions {
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
pub struct ExperimentalRemoveReport {
    /// Automatic destination kind that was processed.
    kind: AutomaticDestinationsKind,
    /// Current user's Windows Recent folder.
    recent_folder: PathBuf,
    /// `.automaticDestinations-ms` file that was deleted and then monitored.
    dest_path: PathBuf,
    /// Target paths requested by the caller.
    requested_paths: Vec<String>,
    /// DestList paths that matched before deletion started.
    matching_paths_before: Vec<String>,
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
    /// Automatic destination kind that was processed.
    #[must_use]
    pub fn kind(&self) -> AutomaticDestinationsKind {
        self.kind
    }

    /// Current user's Windows Recent folder.
    #[must_use]
    pub fn recent_folder(&self) -> &Path {
        &self.recent_folder
    }

    /// `.automaticDestinations-ms` file that was deleted and then monitored.
    #[must_use]
    pub fn dest_path(&self) -> &Path {
        &self.dest_path
    }

    /// Target paths requested by the caller.
    #[must_use]
    pub fn requested_paths(&self) -> &[String] {
        &self.requested_paths
    }

    /// DestList paths that matched before deletion started.
    #[must_use]
    pub fn matching_paths_before(&self) -> &[String] {
        &self.matching_paths_before
    }

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
/// # Examples
///
/// ```rust,no_run
/// # fn main() -> wincent::WincentResult<()> {
/// use wincent::destlist::{
///     experimental_remove_entry_paths_by_rebuild, AutomaticDestinationsKind,
///     ExperimentalRemoveOptions,
/// };
///
/// let report = experimental_remove_entry_paths_by_rebuild(
///     AutomaticDestinationsKind::RecentFiles,
///     &["C:\\Work\\old-report.docx"],
///     ExperimentalRemoveOptions::new(),
/// )?;
///
/// if !report.success() {
///     eprintln!("remaining entries: {:?}", report.remaining_paths_after_rebuild());
/// }
/// Ok(())
/// # }
/// ```
pub fn experimental_remove_entry_paths_by_rebuild<P: AsRef<Path>>(
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

    let before = parse_file_with_retries(&dest_path, Duration::from_secs(2))?;
    let matching_paths_before = matching_dest_paths(before.dest_list().entries(), &requested_paths);

    fs::remove_file(&dest_path).map_err(WincentError::Io)?;

    let base = ExperimentalRemoveBase {
        kind,
        recent_folder,
        dest_path,
        requested_paths,
        matching_paths_before,
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
pub fn experimental_remove_entries_by_rebuild(
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
    kind: AutomaticDestinationsKind,
    recent_folder: PathBuf,
    dest_path: PathBuf,
    requested_paths: Vec<String>,
    matching_paths_before: Vec<String>,
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
            kind: self.kind,
            recent_folder: self.recent_folder,
            dest_path: self.dest_path,
            requested_paths: self.requested_paths,
            matching_paths_before: self.matching_paths_before,
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
                .any(|target| paths_equal(entry.path(), target))
        })
        .map(|entry| entry.path().to_string())
        .collect()
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
                .any(|link| paths_equal(&link.target_path, target))
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
            .any(|requested| paths_equal(&target, requested))
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
    use tempfile::tempdir;

    fn report_base_for_tests(recent_folder: &Path) -> ExperimentalRemoveBase {
        ExperimentalRemoveBase {
            kind: AutomaticDestinationsKind::RecentFiles,
            recent_folder: recent_folder.to_path_buf(),
            dest_path: recent_folder.join("automaticDestinations-ms"),
            requested_paths: vec!["C:\\Work\\Report.docx".to_string()],
            matching_paths_before: vec!["C:\\Work\\Report.docx".to_string()],
        }
    }

    fn parse_recent_files_dest_for_destructive_test(
        dest_path: &Path,
    ) -> WincentResult<AutomaticDestinations> {
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
        let parsed = parse_recent_files_dest_for_destructive_test(&dest_path)?;
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
}
