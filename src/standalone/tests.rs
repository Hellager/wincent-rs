use super::*;
use std::fs;
use std::fs::OpenOptions;
use std::io::ErrorKind;
use std::os::windows::fs::OpenOptionsExt;
use std::path::PathBuf;
use std::sync::Mutex;

const TIMEOUT: Duration = Duration::from_millis(1379);
const PATH: &str = r"C:\Standalone\item";

struct Backend {
    kind: PathType,
    items: Vec<String>,
    pin: FrequentFolderPinStatus,
    query_error: Option<ErrorKind>,
    pin_error: Option<ErrorKind>,
    fail_mutation: bool,
    calls: Mutex<Vec<(&'static str, Duration)>>,
}

impl Backend {
    fn new(kind: PathType) -> Self {
        Self {
            kind,
            items: Vec::new(),
            pin: FrequentFolderPinStatus::NotFound,
            query_error: None,
            pin_error: None,
            fail_mutation: false,
            calls: Mutex::new(Vec::new()),
        }
    }

    fn record(&self, name: &'static str, timeout: Duration) -> WincentResult<()> {
        self.calls
            .lock()
            .expect("calls mutex")
            .push((name, timeout));
        if self.fail_mutation {
            return Err(WincentError::Timeout("injected timeout".into()));
        }
        Ok(())
    }

    fn calls(&self) -> Vec<(&'static str, Duration)> {
        self.calls.lock().expect("calls mutex").clone()
    }
}

impl QuickAccessBackend for Backend {
    fn path_type(&self, _path: &str) -> WincentResult<PathType> {
        Ok(self.kind)
    }

    fn get_items(&self, category: QuickAccess, timeout: Duration) -> WincentResult<Vec<String>> {
        self.calls.lock().expect("calls mutex").push((
            if category == QuickAccess::RecentFiles {
                "query_recent"
            } else {
                "query_frequent"
            },
            timeout,
        ));
        if let Some(kind) = self.query_error {
            return Err(std::io::Error::from(kind).into());
        }
        Ok(self.items.clone())
    }

    fn frequent_folder_pin_status(&self, _path: &str) -> WincentResult<FrequentFolderPinStatus> {
        if let Some(kind) = self.pin_error {
            return Err(std::io::Error::from(kind).into());
        }
        Ok(self.pin)
    }

    fn add_recent_file(&self, _path: &str, timeout: Duration) -> WincentResult<()> {
        self.record("add_recent", timeout)
    }

    fn add_frequent_folder(&self, _path: &str, timeout: Duration) -> WincentResult<()> {
        self.record("pin", timeout)
    }

    fn remove_recent_file(&self, _path: &str, timeout: Duration) -> WincentResult<()> {
        self.record("remove_recent", timeout)
    }

    fn remove_pinned_frequent_folder(&self, _path: &str, timeout: Duration) -> WincentResult<()> {
        self.record("remove_pinned", timeout)
    }

    fn remove_unpinned_frequent_folder(&self, _path: &str, timeout: Duration) -> WincentResult<()> {
        self.record("remove_unpinned", timeout)
    }

    fn clear_recent_files(&self, timeout: Duration) -> WincentResult<()> {
        self.record("clear_recent", timeout)
    }

    fn clear_frequent_folders_jumplist(&self, timeout: Duration) -> WincentResult<()> {
        self.record("clear_frequent", timeout)
    }

    fn notify_shell_change(&self) {
        self.calls
            .lock()
            .expect("calls mutex")
            .push(("notify", Duration::ZERO));
    }

    fn refresh_shell_views(&self, timeout: Duration) -> WincentResult<()> {
        self.record("refresh", timeout)
    }
}

#[test]
fn all_com_entry_points_reject_zero_before_system_access() {
    for result in [
        add_to_recent_files("", Duration::ZERO),
        pin_to_frequent_folders("", Duration::ZERO),
        remove_from_recent_files("", Duration::ZERO),
        remove_from_frequent_folders("", Duration::ZERO),
        empty_recent_files(Duration::ZERO),
        empty_frequent_folders(Duration::ZERO),
        refresh_quick_access(Duration::ZERO),
    ] {
        assert!(matches!(result, Err(WincentError::InvalidArgument(_))));
    }
    assert!(matches!(
        is_folder_pinned(""),
        Err(WincentError::InvalidPath(_))
    ));
}

#[test]
fn mutation_routes_forward_the_exact_timeout() -> WincentResult<()> {
    for (operation, kind, pin, expected) in [
        (
            Mutation::AddRecent,
            PathType::File,
            FrequentFolderPinStatus::NotFound,
            "add_recent",
        ),
        (
            Mutation::PinFolder,
            PathType::Directory,
            FrequentFolderPinStatus::Unpinned,
            "pin",
        ),
        (
            Mutation::RemoveRecent,
            PathType::File,
            FrequentFolderPinStatus::NotFound,
            "remove_recent",
        ),
        (
            Mutation::RemoveFolder,
            PathType::Directory,
            FrequentFolderPinStatus::Pinned,
            "remove_pinned",
        ),
        (
            Mutation::RemoveFolder,
            PathType::Directory,
            FrequentFolderPinStatus::Unpinned,
            "remove_unpinned",
        ),
    ] {
        let mut backend = Backend::new(kind);
        backend.pin = pin;
        if !matches!(operation, Mutation::AddRecent) {
            backend.items.push(PATH.into());
        }
        mutate(Path::new(PATH), TIMEOUT, operation, &backend)?;
        assert_eq!(backend.calls().last(), Some(&(expected, TIMEOUT)));
    }
    Ok(())
}

#[test]
fn rejects_wrong_path_categories_before_query_or_mutation() {
    for (operation, wrong_kind) in [
        (Mutation::AddRecent, PathType::Directory),
        (Mutation::RemoveRecent, PathType::Directory),
        (Mutation::PinFolder, PathType::File),
        (Mutation::RemoveFolder, PathType::File),
    ] {
        let backend = Backend::new(wrong_kind);
        let result = mutate(Path::new(PATH), TIMEOUT, operation, &backend);
        assert!(matches!(result, Err(WincentError::InvalidPath(_))));
        assert!(backend.calls().is_empty());
    }
}

#[test]
fn pin_uses_only_pin_state_and_allows_existing_unpinned_entries() -> WincentResult<()> {
    let mut backend = Backend::new(PathType::Directory);
    backend.items.push(PATH.into());
    backend.query_error = Some(ErrorKind::PermissionDenied);
    for pin in [
        FrequentFolderPinStatus::Unpinned,
        FrequentFolderPinStatus::NotFound,
    ] {
        backend.pin = pin;
        mutate(Path::new(PATH), TIMEOUT, Mutation::PinFolder, &backend)?;
    }
    backend.pin = FrequentFolderPinStatus::Pinned;
    let result = mutate(Path::new(PATH), TIMEOUT, Mutation::PinFolder, &backend);
    assert!(matches!(
        result,
        Err(WincentError::AlreadyExists {
            qa_type: QuickAccess::FrequentFolders,
            ..
        })
    ));
    assert_eq!(backend.calls(), vec![("pin", TIMEOUT); 2]);
    Ok(())
}

#[test]
fn pin_status_uses_domain_absence_and_propagates_backend_errors() -> WincentResult<()> {
    let mut backend = Backend::new(PathType::Directory);
    assert_eq!(
        pin_status(&backend, PATH)?,
        FrequentFolderPinStatus::NotFound
    );
    backend.pin_error = Some(ErrorKind::PermissionDenied);
    assert!(matches!(
        pin_status(&backend, PATH),
        Err(WincentError::Io(_))
    ));
    Ok(())
}

#[test]
fn duplicate_and_missing_membership_keep_category_errors() {
    let mut backend = Backend::new(PathType::File);
    backend.items.push("c:/standalone/ITEM".into());
    assert!(matches!(
        mutate(Path::new(PATH), TIMEOUT, Mutation::AddRecent, &backend),
        Err(WincentError::AlreadyExists {
            qa_type: QuickAccess::RecentFiles,
            ..
        })
    ));
    backend.items.clear();
    assert!(matches!(
        mutate(Path::new(PATH), TIMEOUT, Mutation::RemoveRecent, &backend),
        Err(WincentError::NotInQuickAccess {
            qa_type: QuickAccess::RecentFiles,
            ..
        })
    ));
    backend.kind = PathType::Directory;
    backend.items.push(PATH.into());
    assert!(matches!(
        mutate(Path::new(PATH), TIMEOUT, Mutation::RemoveFolder, &backend),
        Err(WincentError::NotInQuickAccess {
            qa_type: QuickAccess::FrequentFolders,
            ..
        })
    ));
}

#[test]
fn adding_allows_an_empty_category_but_not_query_errors() -> WincentResult<()> {
    let mut backend = Backend::new(PathType::File);
    mutate(Path::new(PATH), TIMEOUT, Mutation::AddRecent, &backend)?;
    backend.query_error = Some(ErrorKind::PermissionDenied);
    assert!(matches!(
        mutate(Path::new(PATH), TIMEOUT, Mutation::AddRecent, &backend),
        Err(WincentError::Io(_))
    ));
    assert_eq!(
        backend
            .calls()
            .iter()
            .filter(|(name, _)| *name == "add_recent")
            .count(),
        1
    );
    Ok(())
}

#[test]
fn mutations_do_not_retry_a_timeout() {
    let mut backend = Backend::new(PathType::File);
    backend.fail_mutation = true;
    assert!(matches!(
        mutate(Path::new(PATH), TIMEOUT, Mutation::AddRecent, &backend),
        Err(WincentError::Timeout(_))
    ));
    assert_eq!(
        backend.calls(),
        vec![("query_recent", TIMEOUT), ("add_recent", TIMEOUT)]
    );
}

#[test]
fn scoped_clears_forward_timeout_and_refresh_only_after_success() -> WincentResult<()> {
    for (category, expected) in [
        (QuickAccess::RecentFiles, "clear_recent"),
        (QuickAccess::FrequentFolders, "clear_frequent"),
    ] {
        let mut backend = Backend::new(PathType::File);
        empty_category(category, TIMEOUT, &backend)?;
        assert_eq!(
            backend.calls(),
            vec![
                (expected, TIMEOUT),
                ("notify", Duration::ZERO),
                ("refresh", TIMEOUT)
            ]
        );
        backend.calls.lock().expect("calls mutex").clear();
        backend.fail_mutation = true;
        assert!(matches!(
            empty_category(category, TIMEOUT, &backend),
            Err(WincentError::Timeout(_))
        ));
        assert_eq!(backend.calls(), vec![(expected, TIMEOUT)]);
    }
    Ok(())
}

fn lock_fixture() -> WincentResult<(tempfile::TempDir, PathBuf)> {
    let directory = tempfile::tempdir()?;
    let root = directory.path().to_path_buf();
    let destinations = root.join("AutomaticDestinations");
    fs::create_dir(&destinations)?;
    for name in [
        crate::lock::RECENT_FILES_AUTOMATIC_DESTINATION,
        crate::lock::FREQUENT_FOLDERS_AUTOMATIC_DESTINATION,
    ] {
        fs::write(destinations.join(name), b"lock fixture")?;
    }
    Ok((directory, root))
}

#[test]
fn typed_guards_can_coexist_and_unlock_without_deleting_shortcuts() -> WincentResult<()> {
    let (_directory, root) = lock_fixture()?;
    let recent = RecentFilesLock {
        inner: QuickAccessLock::lock_target_in_recent_folder(
            QuickAccessLockTarget::RecentFiles,
            root.clone(),
        )?,
    };
    let frequent = FrequentFoldersLock {
        inner: QuickAccessLock::lock_target_in_recent_folder(
            QuickAccessLockTarget::FrequentFolders,
            root.clone(),
        )?,
    };
    let link = root.join("keep-on-drop.lnk");
    fs::write(&link, b"fixture")?;
    recent.unlock();
    frequent.unlock();
    assert!(link.exists());
    let all = QuickAccessLock::lock_target_in_recent_folder(QuickAccessLockTarget::All, root)?;
    drop(all);
    Ok(())
}

#[test]
fn both_guard_types_can_clean_the_entire_recent_shortcut_delta() -> WincentResult<()> {
    for category in [
        QuickAccessLockTarget::RecentFiles,
        QuickAccessLockTarget::FrequentFolders,
    ] {
        let (_directory, root) = lock_fixture()?;
        let inner = QuickAccessLock::lock_target_in_recent_folder(category, root.clone())?;
        let link = root.join("new-shortcut.lnk");
        fs::write(&link, b"fixture")?;
        let report = if category == QuickAccessLockTarget::RecentFiles {
            RecentFilesLock { inner }.unlock_and_cleanup()?
        } else {
            FrequentFoldersLock { inner }.unlock_and_cleanup()?
        };
        assert_eq!(report.cleaned_shortcuts(), std::slice::from_ref(&link));
        assert!(report.is_complete());
        assert!(!link.exists());
        drop(QuickAccessLock::lock_target_in_recent_folder(
            category, root,
        )?);
    }
    Ok(())
}

#[test]
fn unlock_cleanup_failure_still_releases_both_guard_types() -> WincentResult<()> {
    for category in [
        QuickAccessLockTarget::RecentFiles,
        QuickAccessLockTarget::FrequentFolders,
    ] {
        let (_directory, root) = lock_fixture()?;
        let inner = QuickAccessLock::lock_target_in_recent_folder(category, root.clone())?;
        let link = root.join("deny-delete.lnk");
        fs::write(&link, b"fixture")?;
        let deny_delete = OpenOptions::new().read(true).share_mode(1).open(&link)?;
        let report = if category == QuickAccessLockTarget::RecentFiles {
            RecentFilesLock { inner }.unlock_and_cleanup()?
        } else {
            FrequentFoldersLock { inner }.unlock_and_cleanup()?
        };
        assert!(report.cleaned_shortcuts().is_empty());
        assert_eq!(report.cleanup_failures().len(), 1);
        assert_eq!(report.cleanup_failures()[0].path(), link);
        assert!(!report.is_complete());
        drop(QuickAccessLock::lock_target_in_recent_folder(
            category, root,
        )?);
        drop(deny_delete);
    }
    Ok(())
}
