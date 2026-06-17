//! Restore Windows Quick Access categories to their system-default state.

use crate::{
    backend::QuickAccessBackend,
    destlist::{parser::starts_with_knownfolder, DestListEntry},
    error::WincentError,
    recent_links::LnkResolution,
    QuickAccess, WincentResult,
};
use std::path::{Path, PathBuf};
use std::thread;
use std::time::Duration;

const DEFAULT_LNK_RESOLVE_TIMEOUT: Duration = Duration::from_secs(10);
const DEFAULT_CLEAR_TIMEOUT: Duration = Duration::from_secs(10);
const DEFAULT_REBUILD_DELAY: Duration = Duration::from_millis(500);
const DEFAULT_REBUILD_POLL_TIMEOUT: Duration = Duration::from_secs(5);

/// Options for restoring Quick Access categories to their system defaults.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RestoreDefaultsOptions {
    refresh_explorer: bool,
    lnk_resolve_timeout: Duration,
    clear_timeout: Duration,
    rebuild_delay: Duration,
    rebuild_poll_timeout: Duration,
}

impl Default for RestoreDefaultsOptions {
    fn default() -> Self {
        Self {
            refresh_explorer: true,
            lnk_resolve_timeout: DEFAULT_LNK_RESOLVE_TIMEOUT,
            clear_timeout: DEFAULT_CLEAR_TIMEOUT,
            rebuild_delay: DEFAULT_REBUILD_DELAY,
            rebuild_poll_timeout: DEFAULT_REBUILD_POLL_TIMEOUT,
        }
    }
}

impl RestoreDefaultsOptions {
    /// Creates default restore options.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether Explorer should be refreshed to trigger or show rebuilt state.
    #[must_use]
    pub fn refresh_explorer_enabled(&self) -> bool {
        self.refresh_explorer
    }

    /// Sets whether Explorer should be refreshed.
    #[must_use]
    pub fn with_refresh_explorer(mut self, enabled: bool) -> Self {
        self.refresh_explorer = enabled;
        self
    }

    /// Enables Explorer refresh.
    #[must_use]
    pub fn refresh_explorer(self) -> Self {
        self.with_refresh_explorer(true)
    }

    /// Timeout for resolving `.lnk` shortcut targets.
    #[must_use]
    pub fn lnk_resolve_timeout(&self) -> Duration {
        self.lnk_resolve_timeout
    }

    /// Sets timeout for resolving `.lnk` shortcut targets.
    #[must_use]
    pub fn with_lnk_resolve_timeout(mut self, timeout: Duration) -> Self {
        self.lnk_resolve_timeout = timeout;
        self
    }

    /// Timeout for clearing Recent Files through the Shell API.
    #[must_use]
    pub fn clear_timeout(&self) -> Duration {
        self.clear_timeout
    }

    /// Sets timeout for clearing Recent Files through the Shell API.
    #[must_use]
    pub fn with_clear_timeout(mut self, timeout: Duration) -> Self {
        self.clear_timeout = timeout;
        self
    }

    /// Initial delay before polling for a rebuilt Frequent Folders backing file.
    #[must_use]
    pub fn rebuild_delay(&self) -> Duration {
        self.rebuild_delay
    }

    /// Sets initial delay before polling for a rebuilt Frequent Folders backing file.
    #[must_use]
    pub fn with_rebuild_delay(mut self, delay: Duration) -> Self {
        self.rebuild_delay = delay;
        self
    }

    /// Poll timeout for a rebuilt Frequent Folders backing file.
    #[must_use]
    pub fn rebuild_poll_timeout(&self) -> Duration {
        self.rebuild_poll_timeout
    }

    /// Sets poll timeout for a rebuilt Frequent Folders backing file.
    #[must_use]
    pub fn with_rebuild_poll_timeout(mut self, timeout: Duration) -> Self {
        self.rebuild_poll_timeout = timeout;
        self
    }
}

/// Combined restore report for one or more Quick Access categories.
#[derive(Debug)]
pub struct RestoreDefaultsReport {
    recent: Option<RecentRestoreReport>,
    frequent: Option<FrequentRestoreReport>,
}

impl RestoreDefaultsReport {
    pub(crate) fn recent(report: RecentRestoreReport) -> Self {
        Self {
            recent: Some(report),
            frequent: None,
        }
    }

    pub(crate) fn frequent(report: FrequentRestoreReport) -> Self {
        Self {
            recent: None,
            frequent: Some(report),
        }
    }

    pub(crate) fn all(
        recent: RecentRestoreReport,
        frequent: Option<FrequentRestoreReport>,
    ) -> Self {
        Self {
            recent: Some(recent),
            frequent,
        }
    }

    /// Recent Files restore report, if that category was requested.
    #[must_use]
    pub fn recent_report(&self) -> Option<&RecentRestoreReport> {
        self.recent.as_ref()
    }

    /// Frequent Folders restore report, if that category was requested.
    #[must_use]
    pub fn frequent_report(&self) -> Option<&FrequentRestoreReport> {
        self.frequent.as_ref()
    }

    /// Whether every requested restore category succeeded.
    #[must_use]
    pub fn success(&self) -> bool {
        self.recent
            .as_ref()
            .map_or(true, RecentRestoreReport::success)
            && self
                .frequent
                .as_ref()
                .map_or(true, FrequentRestoreReport::success)
    }
}

/// Restore report for Recent Files.
#[derive(Debug)]
pub struct RecentRestoreReport {
    deleted_lnk_paths: Vec<PathBuf>,
    recent_files_cleared: bool,
    error: Option<WincentError>,
}

impl RecentRestoreReport {
    fn new(
        deleted_lnk_paths: Vec<PathBuf>,
        recent_files_cleared: bool,
        error: Option<WincentError>,
    ) -> Self {
        Self {
            deleted_lnk_paths,
            recent_files_cleared,
            error,
        }
    }

    /// `.lnk` files deleted from the Windows Recent folder.
    #[must_use]
    pub fn deleted_lnk_paths(&self) -> &[PathBuf] {
        &self.deleted_lnk_paths
    }

    /// Whether the Shell Recent Files clear operation completed.
    #[must_use]
    pub fn recent_files_cleared(&self) -> bool {
        self.recent_files_cleared
    }

    /// Restore failure captured in the report, if any.
    #[must_use]
    pub fn error(&self) -> Option<&WincentError> {
        self.error.as_ref()
    }

    /// Whether Recent Files were restored successfully.
    #[must_use]
    pub fn success(&self) -> bool {
        self.recent_files_cleared && self.error.is_none()
    }
}

/// Restore report for Frequent Folders.
#[derive(Debug)]
pub struct FrequentRestoreReport {
    deleted_lnk_paths: Vec<PathBuf>,
    backing_file_deleted: bool,
    rebuilt: bool,
    non_default_raw_paths: Vec<String>,
    raw_path_remove_report: Option<FrequentRawPathRemoveReport>,
    error: Option<WincentError>,
}

impl FrequentRestoreReport {
    fn new(
        deleted_lnk_paths: Vec<PathBuf>,
        backing_file_deleted: bool,
        rebuilt: bool,
        non_default_raw_paths: Vec<String>,
        raw_path_remove_report: Option<FrequentRawPathRemoveReport>,
        error: Option<WincentError>,
    ) -> Self {
        Self {
            deleted_lnk_paths,
            backing_file_deleted,
            rebuilt,
            non_default_raw_paths,
            raw_path_remove_report,
            error,
        }
    }

    /// `.lnk` files deleted from the Windows Recent folder.
    #[must_use]
    pub fn deleted_lnk_paths(&self) -> &[PathBuf] {
        &self.deleted_lnk_paths
    }

    /// Whether the Frequent Folders backing file was deleted or already absent.
    #[must_use]
    pub fn backing_file_deleted(&self) -> bool {
        self.backing_file_deleted
    }

    /// Whether a rebuilt Frequent Folders backing file was detected and parsed.
    #[must_use]
    pub fn rebuilt(&self) -> bool {
        self.rebuilt
    }

    /// Raw paths that were not system-default `knownfolder:` entries.
    #[must_use]
    pub fn non_default_raw_paths(&self) -> &[String] {
        &self.non_default_raw_paths
    }

    /// Report from the raw-path remove cleanup, if it was needed.
    #[must_use]
    pub fn raw_path_remove_report(&self) -> Option<&FrequentRawPathRemoveReport> {
        self.raw_path_remove_report.as_ref()
    }

    /// Restore failure captured in the report, if any.
    #[must_use]
    pub fn error(&self) -> Option<&WincentError> {
        self.error.as_ref()
    }

    /// Whether Frequent Folders were restored successfully.
    #[must_use]
    pub fn success(&self) -> bool {
        self.backing_file_deleted
            && self.rebuilt
            && self.error.is_none()
            && self.raw_path_remove_report.as_ref().map_or(
                self.non_default_raw_paths.is_empty(),
                FrequentRawPathRemoveReport::success,
            )
    }
}

/// Report for the Frequent Folders raw-path cleanup pass.
#[derive(Debug)]
pub struct FrequentRawPathRemoveReport {
    requested_raw_paths: Vec<String>,
    backing_file_deleted: bool,
    rebuilt: bool,
    remaining_non_default_raw_paths: Vec<String>,
    error: Option<WincentError>,
}

impl FrequentRawPathRemoveReport {
    fn new(
        requested_raw_paths: Vec<String>,
        backing_file_deleted: bool,
        rebuilt: bool,
        remaining_non_default_raw_paths: Vec<String>,
        error: Option<WincentError>,
    ) -> Self {
        Self {
            requested_raw_paths,
            backing_file_deleted,
            rebuilt,
            remaining_non_default_raw_paths,
            error,
        }
    }

    /// Non-default raw paths requested for cleanup.
    #[must_use]
    pub fn requested_raw_paths(&self) -> &[String] {
        &self.requested_raw_paths
    }

    /// Whether the backing file was deleted or already absent.
    #[must_use]
    pub fn backing_file_deleted(&self) -> bool {
        self.backing_file_deleted
    }

    /// Whether a rebuilt backing file was detected and parsed.
    #[must_use]
    pub fn rebuilt(&self) -> bool {
        self.rebuilt
    }

    /// Non-default raw paths still present after the single cleanup cycle.
    #[must_use]
    pub fn remaining_non_default_raw_paths(&self) -> &[String] {
        &self.remaining_non_default_raw_paths
    }

    /// Cleanup failure captured in the report, if any.
    #[must_use]
    pub fn error(&self) -> Option<&WincentError> {
        self.error.as_ref()
    }

    /// Whether the single raw-path cleanup cycle removed all non-default entries.
    #[must_use]
    pub fn success(&self) -> bool {
        self.backing_file_deleted
            && self.rebuilt
            && self.remaining_non_default_raw_paths.is_empty()
            && self.error.is_none()
    }
}

enum RestoreTarget {
    RecentFiles,
    FrequentFolders,
}

pub(crate) fn restore_defaults(
    qa_type: QuickAccess,
    options: RestoreDefaultsOptions,
    backend: &dyn QuickAccessBackend,
) -> WincentResult<RestoreDefaultsReport> {
    match qa_type {
        QuickAccess::RecentFiles => Ok(RestoreDefaultsReport::recent(
            restore_recent_files_defaults(options, backend)?,
        )),
        QuickAccess::FrequentFolders => Ok(RestoreDefaultsReport::frequent(
            restore_frequent_folders_defaults(options, backend)?,
        )),
        QuickAccess::All => {
            let recent = restore_recent_files_defaults(options, backend)?;
            // The restore workflows capture operational failures in their
            // reports. `?` is reserved for future hard API failures that would
            // prevent a report from being produced at all.
            let frequent = Some(restore_frequent_folders_defaults(options, backend)?);
            Ok(RestoreDefaultsReport::all(recent, frequent))
        }
    }
}

pub(crate) fn restore_recent_files_defaults(
    options: RestoreDefaultsOptions,
    backend: &dyn QuickAccessBackend,
) -> WincentResult<RecentRestoreReport> {
    restore_recent_files_defaults_with(
        options,
        || backend.list_recent_lnk_files(),
        |path, timeout| backend.resolve_lnk_with_type(path, timeout),
        |path| backend.delete_lnk_file(path),
        |timeout| backend.clear_recent_files(timeout),
    )
}

pub(crate) fn restore_frequent_folders_defaults(
    options: RestoreDefaultsOptions,
    backend: &dyn QuickAccessBackend,
) -> WincentResult<FrequentRestoreReport> {
    restore_frequent_folders_defaults_with(
        options,
        || backend.list_recent_lnk_files(),
        |path, timeout| backend.resolve_lnk_with_type(path, timeout),
        |path| backend.delete_lnk_file(path),
        || backend.delete_frequent_folders_backing_file(),
        || {
            if options.refresh_explorer_enabled() {
                backend.refresh_explorer()
            } else {
                Ok(())
            }
        },
        |timeout| backend.wait_for_frequent_folders_rebuild(timeout),
    )
}

fn restore_recent_files_defaults_with<FList, FResolve, FDelete, FClear>(
    options: RestoreDefaultsOptions,
    list_lnk_files: FList,
    resolve_lnk: FResolve,
    delete_lnk: FDelete,
    clear_recent_files: FClear,
) -> WincentResult<RecentRestoreReport>
where
    FList: FnOnce() -> WincentResult<Vec<PathBuf>>,
    FResolve: FnMut(&Path, Duration) -> WincentResult<Option<LnkResolution>>,
    FDelete: FnMut(&Path) -> WincentResult<()>,
    FClear: FnOnce(Duration) -> WincentResult<()>,
{
    let lnk_cleanup = delete_matching_lnk_files(
        RestoreTarget::RecentFiles,
        options.lnk_resolve_timeout(),
        list_lnk_files,
        resolve_lnk,
        delete_lnk,
    );

    if let Some(error) = lnk_cleanup.error {
        return Ok(RecentRestoreReport::new(
            lnk_cleanup.deleted_lnk_paths,
            false,
            Some(error),
        ));
    }

    match clear_recent_files(options.clear_timeout()) {
        Ok(()) => Ok(RecentRestoreReport::new(
            lnk_cleanup.deleted_lnk_paths,
            true,
            None,
        )),
        Err(error) => Ok(RecentRestoreReport::new(
            lnk_cleanup.deleted_lnk_paths,
            false,
            Some(error),
        )),
    }
}

fn restore_frequent_folders_defaults_with<FList, FResolve, FDeleteLnk, FDeleteDest, FRefresh, FWait>(
    options: RestoreDefaultsOptions,
    list_lnk_files: FList,
    resolve_lnk: FResolve,
    delete_lnk: FDeleteLnk,
    mut delete_backing_file: FDeleteDest,
    mut refresh_explorer: FRefresh,
    mut wait_for_rebuild: FWait,
) -> WincentResult<FrequentRestoreReport>
where
    FList: FnOnce() -> WincentResult<Vec<PathBuf>>,
    FResolve: FnMut(&Path, Duration) -> WincentResult<Option<LnkResolution>>,
    FDeleteLnk: FnMut(&Path) -> WincentResult<()>,
    FDeleteDest: FnMut() -> WincentResult<()>,
    FRefresh: FnMut() -> WincentResult<()>,
    FWait: FnMut(Duration) -> WincentResult<Vec<DestListEntry>>,
{
    let lnk_cleanup = delete_matching_lnk_files(
        RestoreTarget::FrequentFolders,
        options.lnk_resolve_timeout(),
        list_lnk_files,
        resolve_lnk,
        delete_lnk,
    );

    if let Some(error) = lnk_cleanup.error {
        return Ok(FrequentRestoreReport::new(
            lnk_cleanup.deleted_lnk_paths,
            false,
            false,
            Vec::new(),
            None,
            Some(error),
        ));
    }

    let backing_file_deleted = match delete_backing_file() {
        Ok(()) => true,
        Err(error) => {
            return Ok(FrequentRestoreReport::new(
                lnk_cleanup.deleted_lnk_paths,
                false,
                false,
                Vec::new(),
                None,
                Some(error),
            ));
        }
    };

    if let Err(error) = refresh_explorer() {
        return Ok(FrequentRestoreReport::new(
            lnk_cleanup.deleted_lnk_paths,
            backing_file_deleted,
            false,
            Vec::new(),
            None,
            Some(error),
        ));
    }

    thread::sleep(options.rebuild_delay());
    let rebuilt = match wait_for_rebuild(options.rebuild_poll_timeout()) {
        Ok(rebuilt) => rebuilt,
        Err(error) => {
            return Ok(FrequentRestoreReport::new(
                lnk_cleanup.deleted_lnk_paths,
                backing_file_deleted,
                false,
                Vec::new(),
                None,
                Some(error),
            ));
        }
    };

    let non_default_raw_paths = non_default_raw_paths(&rebuilt);
    if non_default_raw_paths.is_empty() {
        return Ok(FrequentRestoreReport::new(
            lnk_cleanup.deleted_lnk_paths,
            backing_file_deleted,
            true,
            non_default_raw_paths,
            None,
            None,
        ));
    }

    let raw_path_remove_report = frequent_raw_path_remove_with(
        options,
        &rebuilt,
        &mut delete_backing_file,
        &mut refresh_explorer,
        &mut wait_for_rebuild,
    );
    let final_non_default = raw_path_remove_report
        .remaining_non_default_raw_paths
        .clone();

    Ok(FrequentRestoreReport::new(
        lnk_cleanup.deleted_lnk_paths,
        backing_file_deleted,
        true,
        final_non_default,
        Some(raw_path_remove_report),
        None,
    ))
}

fn frequent_raw_path_remove_with<FDelete, FRefresh, FWait>(
    options: RestoreDefaultsOptions,
    entries: &[DestListEntry],
    mut delete_backing_file: FDelete,
    mut refresh_explorer: FRefresh,
    mut wait_for_rebuild: FWait,
) -> FrequentRawPathRemoveReport
where
    FDelete: FnMut() -> WincentResult<()>,
    FRefresh: FnMut() -> WincentResult<()>,
    FWait: FnMut(Duration) -> WincentResult<Vec<DestListEntry>>,
{
    let requested_raw_paths = non_default_raw_paths(entries);
    if requested_raw_paths.is_empty() {
        return FrequentRawPathRemoveReport::new(Vec::new(), false, false, Vec::new(), None);
    }

    let backing_file_deleted = match delete_backing_file() {
        Ok(()) => true,
        Err(error) => {
            return FrequentRawPathRemoveReport::new(
                requested_raw_paths,
                false,
                false,
                Vec::new(),
                Some(error),
            );
        }
    };

    if let Err(error) = refresh_explorer() {
        return FrequentRawPathRemoveReport::new(
            requested_raw_paths,
            backing_file_deleted,
            false,
            Vec::new(),
            Some(error),
        );
    }

    thread::sleep(options.rebuild_delay());
    match wait_for_rebuild(options.rebuild_poll_timeout()) {
        Ok(rebuilt) => {
            let remaining = non_default_raw_paths(&rebuilt);
            FrequentRawPathRemoveReport::new(
                requested_raw_paths,
                backing_file_deleted,
                true,
                remaining,
                None,
            )
        }
        Err(error) => FrequentRawPathRemoveReport::new(
            requested_raw_paths,
            backing_file_deleted,
            false,
            Vec::new(),
            Some(error),
        ),
    }
}

struct LnkCleanup {
    deleted_lnk_paths: Vec<PathBuf>,
    error: Option<WincentError>,
}

fn delete_matching_lnk_files<FList, FResolve, FDelete>(
    target: RestoreTarget,
    timeout: Duration,
    list_lnk_files: FList,
    mut resolve_lnk: FResolve,
    mut delete_lnk: FDelete,
) -> LnkCleanup
where
    FList: FnOnce() -> WincentResult<Vec<PathBuf>>,
    FResolve: FnMut(&Path, Duration) -> WincentResult<Option<LnkResolution>>,
    FDelete: FnMut(&Path) -> WincentResult<()>,
{
    let mut deleted = Vec::new();

    let lnk_paths = match list_lnk_files() {
        Ok(paths) => paths,
        Err(error) => {
            return LnkCleanup {
                deleted_lnk_paths: deleted,
                error: Some(error),
            };
        }
    };

    for lnk_path in lnk_paths {
        let resolution = match resolve_lnk(&lnk_path, timeout) {
            Ok(resolution) => resolution,
            Err(error) => {
                return LnkCleanup {
                    deleted_lnk_paths: deleted,
                    error: Some(error),
                };
            }
        };
        if should_delete_lnk_for_restore(&target, resolution.as_ref()) {
            if let Err(error) = delete_lnk(&lnk_path) {
                return LnkCleanup {
                    deleted_lnk_paths: deleted,
                    error: Some(error),
                };
            }
            deleted.push(lnk_path);
        }
    }

    LnkCleanup {
        deleted_lnk_paths: deleted,
        error: None,
    }
}

fn should_delete_lnk_for_restore(
    target: &RestoreTarget,
    resolution: Option<&LnkResolution>,
) -> bool {
    match (target, resolution.and_then(|resolution| resolution.is_dir)) {
        (RestoreTarget::RecentFiles, Some(true)) => false,
        (RestoreTarget::FrequentFolders, Some(false)) => false,
        _ => true,
    }
}

fn non_default_raw_paths(entries: &[DestListEntry]) -> Vec<String> {
    entries
        .iter()
        .filter(|entry| !starts_with_knownfolder(entry.raw_path()))
        .map(|entry| entry.raw_path().to_string())
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::destlist::Diagnostic;
    use std::cell::RefCell;

    fn lnk(path: &str, is_dir: Option<bool>) -> LnkResolution {
        LnkResolution {
            path: path.to_string(),
            is_dir,
        }
    }

    fn entry(raw_path: &str) -> DestListEntry {
        DestListEntry {
            entry_offset: 0,
            entry_len: 0,
            mru_position: 0,
            checksum: 0,
            entry_id: 0,
            entry_number: 0,
            entry_number_unknown: 0,
            hostname: String::new(),
            volume_droid: String::new(),
            file_droid: String::new(),
            volume_birth_droid: String::new(),
            file_birth_droid: String::new(),
            file_droid_mac: String::new(),
            stream_name: String::new(),
            raw_path: raw_path.to_string(),
            path: raw_path.to_string(),
            pin_status: -1,
            pin_order: None,
            rank: 0,
            recent_rank: 0,
            count: 0,
            access_count: 0,
            score: 0.0,
            last_access_filetime: None,
            last_interaction_filetime: None,
            sps_size: None,
            reserved_78: None,
            reserved_7c: None,
            path_sources: Vec::new(),
            warnings: Vec::<Diagnostic>::new(),
        }
    }

    #[test]
    fn recent_restore_deletes_files_unknown_missing_and_unresolved_but_keeps_dirs(
    ) -> WincentResult<()> {
        let paths = vec![
            PathBuf::from("file.lnk"),
            PathBuf::from("dir.lnk"),
            PathBuf::from("unknown.lnk"),
            PathBuf::from("unresolved.lnk"),
        ];
        let deleted = RefCell::new(Vec::new());

        let report = restore_recent_files_defaults_with(
            RestoreDefaultsOptions::new(),
            || Ok(paths.clone()),
            |path, _| {
                Ok(match path.to_string_lossy().as_ref() {
                    "file.lnk" => Some(lnk("C:\\file.txt", Some(false))),
                    "dir.lnk" => Some(lnk("C:\\dir", Some(true))),
                    "unknown.lnk" => Some(lnk("C:\\missing", None)),
                    _ => None,
                })
            },
            |path| {
                deleted.borrow_mut().push(path.to_path_buf());
                Ok(())
            },
            |timeout| {
                assert_eq!(timeout, Duration::from_secs(10));
                Ok(())
            },
        )?;

        assert!(report.success());
        assert_eq!(
            deleted.into_inner(),
            vec![
                PathBuf::from("file.lnk"),
                PathBuf::from("unknown.lnk"),
                PathBuf::from("unresolved.lnk"),
            ]
        );
        Ok(())
    }

    #[test]
    fn frequent_restore_deletes_dirs_unknown_missing_and_unresolved_but_keeps_files(
    ) -> WincentResult<()> {
        let paths = vec![
            PathBuf::from("file.lnk"),
            PathBuf::from("dir.lnk"),
            PathBuf::from("unknown.lnk"),
            PathBuf::from("unresolved.lnk"),
        ];
        let deleted = RefCell::new(Vec::new());

        let report = restore_frequent_folders_defaults_with(
            RestoreDefaultsOptions::new().with_rebuild_delay(Duration::ZERO),
            || Ok(paths.clone()),
            |path, _| {
                Ok(match path.to_string_lossy().as_ref() {
                    "file.lnk" => Some(lnk("C:\\file.txt", Some(false))),
                    "dir.lnk" => Some(lnk("C:\\dir", Some(true))),
                    "unknown.lnk" => Some(lnk("C:\\missing", None)),
                    _ => None,
                })
            },
            |path| {
                deleted.borrow_mut().push(path.to_path_buf());
                Ok(())
            },
            || Ok(()),
            || Ok(()),
            |_| Ok(vec![entry("knownfolder:{guid}")]),
        )?;

        assert!(report.success());
        assert_eq!(
            deleted.into_inner(),
            vec![
                PathBuf::from("dir.lnk"),
                PathBuf::from("unknown.lnk"),
                PathBuf::from("unresolved.lnk"),
            ]
        );
        Ok(())
    }

    #[test]
    fn frequent_restore_reports_rebuild_failure() -> WincentResult<()> {
        let report = restore_frequent_folders_defaults_with(
            RestoreDefaultsOptions::new().with_rebuild_delay(Duration::ZERO),
            || Ok(Vec::new()),
            |_, _| Ok(None),
            |_| Ok(()),
            || Ok(()),
            || Ok(()),
            |_| Err(WincentError::Timeout("no rebuild".to_string())),
        )?;

        assert!(!report.success());
        assert!(report.backing_file_deleted());
        assert!(!report.rebuilt());
        assert!(matches!(report.error(), Some(WincentError::Timeout(_))));
        Ok(())
    }

    #[test]
    fn frequent_restore_invokes_raw_path_remove_for_non_default_entries() -> WincentResult<()> {
        let raw_remove_calls = RefCell::new(0usize);

        let report = restore_frequent_folders_defaults_with(
            RestoreDefaultsOptions::new().with_rebuild_delay(Duration::ZERO),
            || Ok(Vec::new()),
            |_, _| Ok(None),
            |_| Ok(()),
            || Ok(()),
            || Ok(()),
            |_| {
                let mut calls = raw_remove_calls.borrow_mut();
                *calls += 1;
                if *calls == 1 {
                    Ok(vec![entry("C:\\Projects"), entry("knownfolder:{guid}")])
                } else {
                    Ok(vec![entry("knownfolder:{guid}")])
                }
            },
        )?;

        assert!(report.success());
        assert_eq!(*raw_remove_calls.borrow(), 2);
        assert!(report.raw_path_remove_report().is_some());
        Ok(())
    }

    #[test]
    fn restore_frequent_folders_defaults_success_records_sequence() -> WincentResult<()> {
        let calls = RefCell::new(Vec::<&'static str>::new());

        let report = restore_frequent_folders_defaults_with(
            RestoreDefaultsOptions::new().with_rebuild_delay(Duration::ZERO),
            || {
                calls.borrow_mut().push("list");
                Ok(Vec::new())
            },
            |_, _| Ok(None),
            |_| Ok(()),
            || {
                calls.borrow_mut().push("delete_backing");
                Ok(())
            },
            || {
                calls.borrow_mut().push("refresh");
                Ok(())
            },
            |_| {
                calls.borrow_mut().push("wait_rebuild");
                Ok(vec![entry("knownfolder:{guid}")])
            },
        )?;

        assert!(report.success());
        assert!(report.rebuilt());
        assert!(report.backing_file_deleted());
        assert!(report.non_default_raw_paths().is_empty());
        assert_eq!(
            *calls.borrow(),
            vec!["list", "delete_backing", "refresh", "wait_rebuild"]
        );
        Ok(())
    }

    #[test]
    fn raw_path_remove_runs_one_cycle_and_reports_remaining_non_default() {
        let report = frequent_raw_path_remove_with(
            RestoreDefaultsOptions::new().with_rebuild_delay(Duration::ZERO),
            &[entry("C:\\Projects")],
            || Ok(()),
            || Ok(()),
            |_| Ok(vec![entry("C:\\Projects")]),
        );

        assert!(!report.success());
        assert_eq!(report.requested_raw_paths(), &["C:\\Projects".to_string()]);
        assert_eq!(
            report.remaining_non_default_raw_paths(),
            &["C:\\Projects".to_string()]
        );
    }
}
