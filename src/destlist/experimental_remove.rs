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
//! affect Quick Access state. Callers should treat this as best-effort experimental
//! functionality and keep their own backups when the data matters. The operation is
//! not transactional; if the process is interrupted after deletion starts, no rollback
//! is attempted.

use std::fs;
use std::path::{Path, PathBuf};
use std::thread;
use std::time::{Duration, Instant};

use crate::error::WincentError;
use crate::utils::{get_windows_recent_folder, paths_equal};
use crate::WincentResult;

use super::parser::{
    frequent_folders_dest_path, parse_file, parse_lnk_local_path, recent_files_dest_path,
    AutomaticDestinations, DestListEntry,
};

/// Explorer automatic destination file family to modify.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AutomaticDestinationsKind {
    /// Recent Files automatic destination.
    RecentFiles,
    /// Frequent Folders automatic destination.
    FrequentFolders,
}

/// Options for the experimental remove-and-rebuild flow.
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
pub fn experimental_remove_entry_paths_by_rebuild<P: AsRef<Path>>(
    kind: AutomaticDestinationsKind,
    target_paths: &[P],
    options: ExperimentalRemoveOptions,
) -> WincentResult<ExperimentalRemoveReport> {
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
    let dest_deleted = true;

    let deleted_links = delete_matching_recent_links(&recent_folder, &requested_paths)?;
    let missing_lnk_target_paths = requested_paths
        .iter()
        .filter(|target| {
            !deleted_links
                .iter()
                .any(|link| paths_equal(&link.target_path, target))
        })
        .cloned()
        .collect();
    let deleted_lnk_paths = deleted_links
        .iter()
        .map(|link| link.lnk_path.clone())
        .collect();

    crate::utils::refresh_explorer_window()?;
    thread::sleep(options.rebuild_delay());

    let check = wait_for_rebuilt_dest(&dest_path, &requested_paths, Duration::from_secs(5))?;
    let (rebuilt, rebuild_parse_elapsed, rebuild_parse_error, remaining_paths_after_rebuild) =
        match check.rebuilt {
            Some(rebuilt) => (
                true,
                Some(rebuilt.elapsed),
                check.last_parse_error,
                rebuilt.remaining_paths,
            ),
            None => (false, None, check.last_parse_error, Vec::new()),
        };

    let success = rebuilt && remaining_paths_after_rebuild.is_empty();

    Ok(ExperimentalRemoveReport {
        kind,
        recent_folder,
        dest_path,
        requested_paths,
        matching_paths_before,
        deleted_lnk_paths,
        missing_lnk_target_paths,
        dest_deleted,
        rebuilt,
        rebuild_parse_elapsed,
        rebuild_parse_error,
        remaining_paths_after_rebuild,
        success,
    })
}

/// Variant that accepts parsed DestList entries directly.
///
/// # Experimental and risky
///
/// This has the same risks as [`experimental_remove_entry_paths_by_rebuild`].
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

fn delete_matching_recent_links(
    recent_folder: &Path,
    target_paths: &[String],
) -> WincentResult<Vec<DeletedRecentLink>> {
    let mut deleted = Vec::new();

    for entry in fs::read_dir(recent_folder).map_err(WincentError::Io)? {
        let entry = entry.map_err(WincentError::Io)?;
        let path = entry.path();
        if !is_lnk_file(&path) {
            continue;
        }

        let Some(target) = lnk_target_path(&path) else {
            continue;
        };

        if target_paths
            .iter()
            .any(|requested| paths_equal(&target, requested))
        {
            fs::remove_file(&path).map_err(WincentError::Io)?;
            deleted.push(DeletedRecentLink {
                lnk_path: path,
                target_path: target,
            });
        }
    }

    Ok(deleted)
}

#[derive(Debug)]
struct DeletedRecentLink {
    lnk_path: PathBuf,
    target_path: String,
}

fn lnk_target_path(lnk_path: &Path) -> Option<String> {
    let data = fs::read(lnk_path).ok()?;
    parse_lnk_local_path(&data)
}

fn is_lnk_file(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.eq_ignore_ascii_case("lnk"))
        .unwrap_or(false)
}

#[cfg(test)]
mod tests {
    use super::*;

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
