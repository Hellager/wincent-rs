//! Standalone API coverage inside the same guarded VM session as the manager suite.

use super::{membership, normalized, same_paths, shell, wait_for, Checks, TestResult, WAIT};
use std::env;
use std::fs::{self, File, OpenOptions};
use std::os::windows::fs::OpenOptionsExt;
use std::path::Path;
use std::process::{Command, Stdio};
use std::time::{Duration, Instant};
use wincent::prelude::*;

fn invalid_inputs(checks: &mut Checks, file: &Path, folder: &Path) -> TestResult {
    for (name, result) in [
        ("add", add_to_recent_files(file, Duration::ZERO)),
        ("pin", pin_to_frequent_folders(folder, Duration::ZERO)),
        (
            "remove_file",
            remove_from_recent_files(file, Duration::ZERO),
        ),
        (
            "remove_folder",
            remove_from_frequent_folders(folder, Duration::ZERO),
        ),
        ("empty_recent", empty_recent_files(Duration::ZERO)),
        ("empty_frequent", empty_frequent_folders(Duration::ZERO)),
        ("refresh", refresh_quick_access(Duration::ZERO)),
    ] {
        checks.check(
            &format!("standalone.zero_timeout.{name}"),
            matches!(result, Err(WincentError::InvalidArgument(_))),
            result,
        )?;
    }
    for (name, result) in [
        ("add_folder", add_to_recent_files(folder, WAIT)),
        ("pin_file", pin_to_frequent_folders(file, WAIT)),
        (
            "remove_folder_as_file",
            remove_from_recent_files(folder, WAIT),
        ),
        (
            "remove_file_as_folder",
            remove_from_frequent_folders(file, WAIT),
        ),
        ("add_empty", add_to_recent_files("", WAIT)),
        ("pin_empty", pin_to_frequent_folders("", WAIT)),
        ("remove_file_empty", remove_from_recent_files("", WAIT)),
        (
            "remove_folder_empty",
            remove_from_frequent_folders("", WAIT),
        ),
    ] {
        checks.check(
            &format!("standalone.invalid.{name}"),
            matches!(result, Err(WincentError::InvalidPath(_))),
            result,
        )?;
    }
    let missing = folder.join("missing");
    for (name, result) in [
        ("add", add_to_recent_files(&missing, WAIT)),
        ("pin", pin_to_frequent_folders(&missing, WAIT)),
        ("remove_file", remove_from_recent_files(&missing, WAIT)),
        (
            "remove_folder",
            remove_from_frequent_folders(&missing, WAIT),
        ),
    ] {
        checks.check(
            &format!("standalone.missing.{name}"),
            matches!(result, Err(WincentError::InvalidPath(_))),
            result,
        )?;
    }
    let empty = is_folder_pinned("");
    checks.check(
        "standalone.pin_status.empty",
        matches!(empty, Err(WincentError::InvalidPath(_))),
        empty,
    )?;
    checks.check(
        "standalone.pin_status.missing",
        !is_folder_pinned(missing)?,
        "absent path does not require filesystem existence",
    )?;
    Ok(())
}

fn compare_queries(checks: &mut Checks, manager: &QuickAccessManager) -> TestResult {
    // Use the crate-root exports too, not only prelude exports.
    for (category, actual) in [
        (QuickAccess::RecentFiles, wincent::get_recent_files()?),
        (
            QuickAccess::FrequentFolders,
            wincent::get_frequent_folders()?,
        ),
    ] {
        checks.check(
            &format!("standalone.query.manager.{category:?}"),
            same_paths(&actual, &manager.get_items(category)?),
            &actual,
        )?;
        let oracle = wait_for(|| Ok(same_paths(&actual, &shell::query(category)?)));
        checks.require(
            &format!("standalone.query.shell.{category:?}"),
            oracle.is_ok(),
            oracle,
        )?;
    }
    Ok(())
}

fn child_probe(checks: &mut Checks, root: &Path, category: QuickAccess) -> TestResult {
    let log = File::create(root.join(format!("standalone-child-{category:?}.log")))?;
    let mut child = Command::new(env::current_exe()?)
        .args(["--ignored", "--exact", "lock_probe", "--nocapture"])
        .env("WINCENT_QAM_LOCK", format!("{category:?}"))
        .stdout(Stdio::from(log.try_clone()?))
        .stderr(Stdio::from(log))
        .spawn()?;
    let start = Instant::now();
    let success = loop {
        if let Some(status) = child.try_wait()? {
            break status.success();
        }
        if start.elapsed() >= WAIT {
            child.kill()?;
            child.wait()?;
            break false;
        }
        std::thread::sleep(Duration::from_millis(100));
    };
    checks.require(
        &format!("standalone.lock.child.{category:?}"),
        success,
        "both manager and standalone mutations must be blocked in another process",
    )
}

fn while_locked(
    checks: &mut Checks,
    manager: &QuickAccessManager,
    root: &Path,
    category: QuickAccess,
    file: &Path,
    folder: &Path,
) -> TestResult {
    let probe = root.join("targets").join("lock-probe.txt");
    for (name, result) in [
        ("add_file", add_to_recent_files(&probe, WAIT)),
        ("manager_add", manager.add_item(&probe)),
        (
            "remove",
            if category == QuickAccess::RecentFiles {
                remove_from_recent_files(file, WAIT)
            } else {
                remove_from_frequent_folders(folder, WAIT)
            },
        ),
        (
            "clear",
            if category == QuickAccess::RecentFiles {
                empty_recent_files(WAIT)
            } else {
                empty_frequent_folders(WAIT)
            },
        ),
    ] {
        checks.check(
            &format!("standalone.lock.blocked.{category:?}.{name}"),
            matches!(result, Err(WincentError::QuickAccessLocked { .. })),
            result,
        )?;
    }
    if category == QuickAccess::FrequentFolders {
        let result = pin_to_frequent_folders(root.join("targets").join("lock-probe-folder"), WAIT);
        checks.check(
            "standalone.lock.blocked.pin",
            matches!(result, Err(WincentError::QuickAccessLocked { .. })),
            result,
        )?;
    }
    let overlap = manager.lock(category);
    checks.require(
        &format!("standalone.lock.manager_overlap.{category:?}"),
        overlap.is_err(),
        overlap,
    )?;
    checks.check(
        &format!("standalone.lock.query.{category:?}"),
        !get_recent_files()?.is_empty() && is_folder_pinned(folder)?,
        "queries work while locked",
    )?;
    child_probe(checks, root, category)
}

fn locks(
    checks: &mut Checks,
    manager: &QuickAccessManager,
    root: &Path,
    file: &Path,
    folder: &Path,
) -> TestResult {
    let backing = recent_files_dest_path()?;
    let recent = backing
        .parent()
        .and_then(Path::parent)
        .ok_or("Recent parent")?;
    for category in [QuickAccess::RecentFiles, QuickAccess::FrequentFolders] {
        let link = recent.join(format!(
            "standalone-{category:?}-{}.lnk",
            std::process::id()
        ));
        let links = if category == QuickAccess::RecentFiles {
            let guard = lock_recent_files()?;
            let repeated = lock_recent_files();
            checks.require(
                "standalone.lock.recent_overlap",
                repeated.is_err(),
                &repeated,
            )?;
            while_locked(checks, manager, root, category, file, folder)?;
            fs::write(&link, b"new shortcut while independently locked")?;
            guard.unlock_and_cleanup()?
        } else {
            let guard = lock_frequent_folders()?;
            let repeated = lock_frequent_folders();
            checks.require(
                "standalone.lock.frequent_overlap",
                repeated.is_err(),
                &repeated,
            )?;
            while_locked(checks, manager, root, category, file, folder)?;
            fs::write(&link, b"new shortcut while independently locked")?;
            guard.unlock_and_cleanup()?
        };
        checks.check(
            &format!("standalone.unlock.delta.{category:?}"),
            links.cleaned_shortcuts().iter().any(|value| value == &link)
                && links.is_complete()
                && !link.exists(),
            &links,
        )?;

        let kept = recent.join(format!(
            "standalone-keep-{category:?}-{}.lnk",
            std::process::id()
        ));
        if category == QuickAccess::RecentFiles {
            let guard = lock_recent_files()?;
            fs::write(&kept, b"plain unlock must not clean shortcut delta")?;
            guard.unlock();
        } else {
            let guard = lock_frequent_folders()?;
            fs::write(&kept, b"plain unlock must not clean shortcut delta")?;
            guard.unlock();
        }
        checks.check(
            &format!("standalone.unlock.keeps_delta.{category:?}"),
            kept.exists(),
            &kept,
        )?;
        fs::remove_file(&kept)?;

        let denied = recent.join(format!(
            "standalone-deny-delete-{category:?}-{}.lnk",
            std::process::id()
        ));
        let report = if category == QuickAccess::RecentFiles {
            let guard = lock_recent_files()?;
            fs::write(&denied, b"delete sharing failure")?;
            let deny_delete = OpenOptions::new().read(true).share_mode(1).open(&denied)?;
            let report = guard.unlock_and_cleanup()?;
            manager.lock(category)?;
            manager.unlock()?;
            drop(deny_delete);
            report
        } else {
            let guard = lock_frequent_folders()?;
            fs::write(&denied, b"delete sharing failure")?;
            let deny_delete = OpenOptions::new().read(true).share_mode(1).open(&denied)?;
            let report = guard.unlock_and_cleanup()?;
            manager.lock(category)?;
            manager.unlock()?;
            drop(deny_delete);
            report
        };
        checks.check(
            &format!("standalone.unlock.failure_releases.{category:?}"),
            report.cleaned_shortcuts().is_empty()
                && report.cleanup_failures().len() == 1
                && report.cleanup_failures()[0].path() == denied
                && !report.is_complete(),
            &report,
        )?;
        fs::remove_file(&denied)?;

        manager.lock(category)?;
        let overlap = if category == QuickAccess::RecentFiles {
            lock_recent_files().map(drop)
        } else {
            lock_frequent_folders().map(drop)
        };
        checks.check(
            &format!("standalone.lock.manager_owns.{category:?}"),
            overlap.is_err(),
            overlap,
        )?;
        manager.unlock()?;
    }
    let recent = lock_recent_files()?;
    let frequent = lock_frequent_folders()?;
    drop(recent);
    manager.lock(QuickAccess::RecentFiles)?;
    manager.unlock()?;
    drop(frequent);
    manager.lock(QuickAccess::All)?;
    manager.unlock()?;
    checks.check(
        "standalone.lock.disjoint_drop_releases",
        true,
        "typed guards coexist; drop permits manager reacquisition",
    )
}

fn clears(checks: &mut Checks, manager: &QuickAccessManager, targets: &Path) -> TestResult {
    for category in [QuickAccess::RecentFiles, QuickAccess::FrequentFolders] {
        empty_recent_files(WAIT)?;
        empty_frequent_folders(WAIT)?;
        let file = targets.join(format!("standalone-clear-{category:?}.txt"));
        let folder = targets.join(format!("standalone-clear-{category:?}"));
        fs::write(&file, b"scoped standalone clear")?;
        fs::create_dir(&folder)?;
        add_to_recent_files(&file, WAIT)?;
        pin_to_frequent_folders(&folder, WAIT)?;
        membership(
            checks,
            manager,
            &format!("standalone.clear.seed_file.{category:?}"),
            &file,
            QuickAccess::RecentFiles,
            true,
        )?;
        membership(
            checks,
            manager,
            &format!("standalone.clear.seed_folder.{category:?}"),
            &folder,
            QuickAccess::FrequentFolders,
            true,
        )?;
        if category == QuickAccess::RecentFiles {
            empty_recent_files(WAIT)?;
        } else {
            empty_frequent_folders(WAIT)?;
        }
        for (path, queried, expected) in [
            (
                &file,
                QuickAccess::RecentFiles,
                category == QuickAccess::FrequentFolders,
            ),
            (
                &folder,
                QuickAccess::FrequentFolders,
                category == QuickAccess::RecentFiles,
            ),
        ] {
            let result = wait_for(|| {
                let status = manager.get_item_status(path)?;
                let present = if queried == QuickAccess::RecentFiles {
                    status.recent_file().is_some()
                } else {
                    status.frequent_folder().is_some()
                };
                let oracle = shell::query(queried)?;
                let shell_present = oracle
                    .iter()
                    .any(|value| normalized(value) == normalized(&path.to_string_lossy()));
                Ok(present == expected && shell_present == expected)
            });
            checks.require(
                &format!("standalone.clear.scope.{category:?}.{queried:?}"),
                result.is_ok(),
                result,
            )?;
        }
        if category == QuickAccess::RecentFiles {
            empty_recent_files(WAIT)?;
        } else {
            empty_frequent_folders(WAIT)?;
        }
        checks.check(&format!("standalone.clear.repeat.{category:?}"), true, "Ok")?;
    }
    refresh_quick_access(WAIT)?;
    checks.check(
        "standalone.refresh.cleared",
        true,
        "API result, not pixel validation",
    )
}

pub(super) fn run(checks: &mut Checks, manager: &QuickAccessManager, root: &Path) -> TestResult {
    empty_recent_files(WAIT)?;
    empty_frequent_folders(WAIT)?;
    let targets = root.join("targets");
    let pin_parent = targets.join("standalone pin \u{76ee}\u{5f55}");
    let remove_parent = targets.join("standalone remove \u{76ee}\u{5f55}");
    fs::create_dir(&pin_parent)?;
    fs::create_dir(&remove_parent)?;
    let mut files = Vec::new();
    for visit in 0..3 {
        if visit != 0 {
            println!("standalone fixture: waiting 66 seconds for distinct native usage");
            std::thread::sleep(Duration::from_secs(66));
        }
        for parent in [&pin_parent, &remove_parent] {
            let path = parent.join(format!("standalone \u{6587}\u{4ef6}-{visit}.txt"));
            fs::write(&path, b"standalone native fixture")?;
            add_to_recent_files(&path, WAIT)?;
            files.push(path);
        }
    }
    for (index, path) in files.iter().enumerate() {
        membership(
            checks,
            manager,
            &format!("standalone.add.visible.{index}"),
            path,
            QuickAccess::RecentFiles,
            true,
        )?;
    }
    for (name, parent) in [("pin", &pin_parent), ("remove", &remove_parent)] {
        let unpinned = wait_for(|| {
            Ok(manager
                .get_item_status(parent)?
                .frequent_folder()
                .is_some_and(|slot| !slot.is_pinned())
                && !is_folder_pinned(parent)?)
        });
        checks.require(
            &format!("standalone.fixture.unpinned.{name}"),
            unpinned.is_ok(),
            unpinned,
        )?;
    }
    let file = files.last().ok_or("standalone files")?;
    invalid_inputs(checks, file, &pin_parent)?;
    compare_queries(checks, manager)?;
    pin_to_frequent_folders(&pin_parent, WAIT)?;
    let pinned = wait_for(|| is_folder_pinned(&pin_parent));
    checks.require("standalone.pin.existing_unpinned", pinned.is_ok(), pinned)?;
    for (index, path) in [
        pin_parent.to_string_lossy().to_uppercase(),
        pin_parent.to_string_lossy().replace('\\', "/"),
        format!("{}\\.", pin_parent.display()),
        pin_parent.canonicalize()?.to_string_lossy().into_owned(),
    ]
    .iter()
    .enumerate()
    {
        checks.check(
            &format!("standalone.pin.path_variant.{index}"),
            is_folder_pinned(path)?,
            path,
        )?;
    }
    for (name, result) in [
        ("recent", add_to_recent_files(file, WAIT)),
        ("pin", pin_to_frequent_folders(&pin_parent, WAIT)),
    ] {
        checks.check(
            &format!("standalone.duplicate.{name}"),
            matches!(result, Err(WincentError::AlreadyExists { .. })),
            result,
        )?;
    }
    refresh_quick_access(WAIT)?;
    checks.check("standalone.refresh.populated", true, "Ok")?;
    locks(checks, manager, root, file, &pin_parent)?;
    remove_from_frequent_folders(&remove_parent, WAIT)?;
    membership(
        checks,
        manager,
        "standalone.remove.unpinned",
        &remove_parent,
        QuickAccess::FrequentFolders,
        false,
    )?;
    remove_from_frequent_folders(&pin_parent, WAIT)?;
    membership(
        checks,
        manager,
        "standalone.remove.pinned",
        &pin_parent,
        QuickAccess::FrequentFolders,
        false,
    )?;
    checks.check(
        "standalone.pin_status.removed",
        !is_folder_pinned(&pin_parent)?,
        &pin_parent,
    )?;
    remove_from_recent_files(file, WAIT)?;
    membership(
        checks,
        manager,
        "standalone.remove.recent",
        file,
        QuickAccess::RecentFiles,
        false,
    )?;
    for (name, result) in [
        ("recent", remove_from_recent_files(file, WAIT)),
        ("frequent", remove_from_frequent_folders(&pin_parent, WAIT)),
    ] {
        checks.check(
            &format!("standalone.remove.absent.{name}"),
            matches!(result, Err(WincentError::NotInQuickAccess { .. })),
            result,
        )?;
    }
    clears(checks, manager, &targets)
}
