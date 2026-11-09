//! Destructive, VM-only acceptance tests for the public manager facade.
//!
//! Run through run_quick_access_manager_vm.ps1, never cargo test -- --ignored
//! on a workstation. The guest harness owns backup, restoration and JSON reports.

#![cfg(windows)]

mod shell;

use std::env;
use std::fs::{self, File, OpenOptions};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::time::{Duration, Instant};

use wincent::prelude::*;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const WAIT: Duration = Duration::from_secs(30);
const CATEGORIES: [QuickAccess; 3] = [
    QuickAccess::RecentFiles,
    QuickAccess::FrequentFolders,
    QuickAccess::All,
];

struct Checks {
    output: File,
    failures: usize,
    count: usize,
}

impl Checks {
    fn new(root: &Path) -> TestResult<Self> {
        let mut output = File::create(root.join("checks.tsv"))?;
        writeln!(output, "id\tok\tdetail")?;
        Ok(Self {
            output,
            failures: 0,
            count: 0,
        })
    }

    fn check(&mut self, id: &str, ok: bool, detail: impl std::fmt::Debug) -> TestResult {
        let detail = format!("{detail:?}").replace(['\t', '\r', '\n'], " ");
        writeln!(
            self.output,
            "{id}\t{ok}\t\"{}\"",
            detail.replace('"', "\"\"")
        )?;
        self.output.flush()?;
        self.count += 1;
        self.failures += usize::from(!ok);
        println!("{} {id}: {detail}", if ok { "PASS" } else { "FAIL" });
        Ok(())
    }

    fn require(&mut self, id: &str, ok: bool, detail: impl std::fmt::Debug) -> TestResult {
        self.check(id, ok, detail)?;
        if !ok {
            return Err(format!("prerequisite failed: {id}; dependent phases not run").into());
        }
        Ok(())
    }
}

fn guarded_root() -> TestResult<PathBuf> {
    let run_id = env::var("WINCENT_QAM_RUN_ID")?;
    let guest = env::var("WINCENT_QAM_GUEST")?;
    if run_id.is_empty()
        || env::var("WINCENT_QAM_VM_ACK")? != run_id
        || env::var("COMPUTERNAME")? != guest
    {
        return Err("refusing to run without the guest harness acknowledgement".into());
    }
    let root = PathBuf::from(env::var("WINCENT_QAM_ROOT")?);
    let expected = Path::new(r"C:\WincentTrack").join(format!("qam-{run_id}"));
    if root != expected || !root.join("backup-ready.txt").is_file() {
        return Err("VM run directory or backup-ready marker is invalid".into());
    }
    Ok(root)
}

fn normalized(path: &str) -> String {
    path.replace('/', "\\")
        .trim_end_matches('\\')
        .to_lowercase()
}

fn same_paths(left: &[String], right: &[String]) -> bool {
    left.len() == right.len()
        && left
            .iter()
            .zip(right)
            .all(|(a, b)| normalized(a) == normalized(b))
}

fn wait_for(mut predicate: impl FnMut() -> WincentResult<bool>) -> TestResult<String> {
    let started = Instant::now();
    loop {
        let observation = predicate();
        if matches!(observation, Ok(true)) {
            return Ok(format!(
                "converged after {} ms",
                started.elapsed().as_millis()
            ));
        }
        if started.elapsed() >= WAIT {
            return Err(format!("did not converge within {WAIT:?}; last={observation:?}").into());
        }
        std::thread::sleep(Duration::from_millis(250));
    }
}

fn membership(
    checks: &mut Checks,
    manager: &QuickAccessManager,
    id: &str,
    path: &Path,
    category: QuickAccess,
    expected: bool,
) -> TestResult {
    let result = wait_for(|| Ok(manager.contains_path(path, category)? == expected));
    checks.require(id, result.is_ok(), &result)
}

fn snapshot(root: &Path, name: &str, manager: &QuickAccessManager) -> TestResult {
    let directory = root.join(name);
    fs::create_dir(&directory)?;
    for (label, path) in [
        ("recent", recent_files_dest_path()?),
        ("frequent", frequent_folders_dest_path()?),
    ] {
        if path.exists() {
            fs::copy(&path, directory.join(format!("{label}.cfb")))?;
        }
    }
    for category in CATEGORIES {
        fs::write(
            directory.join(format!("{category:?}.txt")),
            format!("{:#?}", manager.get_items(category)),
        )?;
    }
    Ok(())
}

fn constructors(checks: &mut Checks) -> TestResult {
    let default_timeout = Duration::from_secs(10);
    for (id, manager) in [
        ("new", QuickAccessManager::new()),
        ("default", QuickAccessManager::default()),
        ("builder", QuickAccessManager::builder().build()),
        (
            "builder_default",
            QuickAccessManagerBuilder::default().build(),
        ),
    ] {
        checks.check(
            &format!("constructor.{id}"),
            manager.timeout() == default_timeout,
            manager,
        )?;
    }
    let builder = QuickAccessManager::builder().timeout(WAIT);
    let manager = builder.clone().try_timeout(WAIT)?.build();
    checks.check("constructor.timeout", manager.timeout() == WAIT, &manager)?;
    checks.check(
        "constructor.clone",
        manager.clone().timeout() == WAIT,
        &builder,
    )?;
    let zero = QuickAccessManager::builder().try_timeout(Duration::ZERO);
    checks.check(
        "constructor.try_zero",
        matches!(zero, Err(WincentError::InvalidArgument(_))),
        zero,
    )?;
    let panic = std::panic::catch_unwind(|| {
        let _ = QuickAccessManager::builder().timeout(Duration::ZERO);
    });
    checks.check(
        "constructor.zero_panics",
        panic.is_err(),
        "documented panic",
    )?;
    Ok(())
}

fn visibility(checks: &mut Checks, manager: &QuickAccessManager) -> TestResult {
    for target in [
        VisibilityTarget::RecentFiles,
        VisibilityTarget::FrequentFolders,
        VisibilityTarget::StartRecommended,
        VisibilityTarget::All,
    ] {
        for visible in [false, true] {
            manager.set_visible(target, visible)?;
            checks.check(
                &format!("visibility.{target:?}.{visible}"),
                manager.is_visible(target)? == visible,
                "configured registry visibility, not rendered UI",
            )?;
        }
    }
    for start in [false, true] {
        manager.set_visible(VisibilityTarget::StartRecommended, start)?;
        for all in [false, true] {
            manager.set_visible(VisibilityTarget::All, all)?;
            checks.check(
                &format!("visibility.all_excludes_start.{start}.{all}"),
                manager.is_visible(VisibilityTarget::StartRecommended)? == start
                    && manager.is_visible(VisibilityTarget::RecentFiles)? == all
                    && manager.is_visible(VisibilityTarget::FrequentFolders)? == all,
                "All changes only the two Explorer settings",
            )?;
        }
    }
    manager.set_visible(VisibilityTarget::RecentFiles, false)?;
    checks.check(
        "visibility.all_requires_both",
        !manager.is_visible(VisibilityTarget::All)?,
        "mixed visibility",
    )?;
    manager.set_visible(VisibilityTarget::StartRecommended, true)?;
    manager.set_visible(VisibilityTarget::All, true)?;
    Ok(())
}

fn invalid_inputs(checks: &mut Checks, manager: &QuickAccessManager, targets: &Path) -> TestResult {
    let missing = targets.join("does-not-exist.txt");
    for (name, path) in [("empty", Path::new("")), ("missing", missing.as_path())] {
        let add = manager.add_item(path);
        let remove = manager.remove_item(path);
        checks.check(
            &format!("errors.add_{name}"),
            matches!(add, Err(WincentError::InvalidPath(_))),
            add,
        )?;
        checks.check(
            &format!("errors.remove_{name}"),
            matches!(remove, Err(WincentError::InvalidPath(_))),
            remove,
        )?;
    }
    for category in CATEGORIES {
        let result = manager.contains_path("", category);
        checks.check(
            &format!("errors.contains_empty.{category:?}"),
            matches!(result, Err(WincentError::InvalidPath(_))),
            result,
        )?;
        checks.check(
            &format!("errors.contains_missing.{category:?}"),
            !manager.contains_path(&missing, category)?,
            &missing,
        )?;
    }
    let status = manager.get_item_status("");
    checks.check(
        "errors.status_empty",
        matches!(status, Err(WincentError::InvalidPath(_))),
        status,
    )?;
    checks.check(
        "errors.status_missing",
        manager.get_item_status(&missing)? == QuickAccessStatus::default(),
        &missing,
    )?;
    for (name, directory, category) in [
        ("absent.txt", false, QuickAccess::RecentFiles),
        ("absent-folder", true, QuickAccess::FrequentFolders),
    ] {
        let path = targets.join(name);
        if directory {
            fs::create_dir(&path)?;
        } else {
            fs::write(&path, b"not in Quick Access")?;
        }
        let result = manager.remove_item(&path);
        let correct = matches!(
            &result,
            Err(WincentError::NotInQuickAccess { path: actual, qa_type, .. })
                if Path::new(actual) == path && *qa_type == category
        );
        checks.check(&format!("errors.remove_absent.{name}"), correct, result)?;
    }
    Ok(())
}

fn fixture(
    checks: &mut Checks,
    manager: &QuickAccessManager,
    targets: &Path,
) -> TestResult<(Vec<PathBuf>, Vec<PathBuf>)> {
    let recent_directory = targets.join("Recent fixture");
    fs::create_dir(&recent_directory)?;
    let mut files = Vec::new();
    for index in 0..25 {
        if index == 1 || index == 2 {
            // Native Frequent usage coalesces visits within roughly one minute.
            // See EMPTY_ITEMS_VM_VALIDATION_20261029_CN.md: three distinct files,
            // 66 seconds apart, create a visible unpinned parent without CFB edits.
            println!("fixture: waiting 66 seconds for native Frequent usage visit {index}");
            std::thread::sleep(Duration::from_secs(66));
        }
        let path = recent_directory.join(format!("QAM-\u{4e2d}\u{6587}-{index:02}.txt"));
        fs::write(&path, b"native QuickAccessManager fixture")?;
        manager.add_item(&path)?;
        files.push(path);
    }
    let persisted = wait_for(|| {
        let parsed = parse_dest_file(recent_files_dest_path()?)?;
        Ok(files.iter().all(|path| {
            parsed
                .cfb_entries()
                .iter()
                .any(|entry| normalized(entry.path()) == normalized(&path.to_string_lossy()))
        }))
    });
    checks.require("fixture.recent_25_persisted", persisted.is_ok(), persisted)?;
    let mut folders = Vec::new();
    for name in ["Pinned folder A", "Pinned \u{4e2d}\u{6587} B"] {
        let path = targets.join(name);
        fs::create_dir(&path)?;
        manager.add_item(&path)?;
        membership(
            checks,
            manager,
            &format!("fixture.{name}"),
            &path,
            QuickAccess::FrequentFolders,
            true,
        )?;
        checks.require(
            &format!("fixture.pinned.{name}"),
            manager
                .get_item_status(&path)?
                .frequent_folder()
                .is_some_and(|slot| slot.is_pinned()),
            &path,
        )?;
        folders.push(path);
    }
    Ok((files, folders))
}

fn queries(checks: &mut Checks, manager: &QuickAccessManager, root: &Path) -> TestResult {
    let recent = manager.get_items(QuickAccess::RecentFiles)?;
    let frequent = manager.get_items(QuickAccess::FrequentFolders)?;
    let mut merged = recent.clone();
    for item in &frequent {
        if !merged.iter().any(|p| normalized(p) == normalized(item)) {
            merged.push(item.clone());
        }
    }
    checks.check(
        "query.all_recent_first_deduplicated",
        same_paths(&manager.get_items(QuickAccess::All)?, &merged),
        &merged,
    )?;
    for category in CATEGORIES {
        let items = manager.get_items(category)?;
        let paths = manager.get_item_paths(category)?;
        checks.check(
            &format!("query.pathbuf.{category:?}"),
            paths == items.iter().map(PathBuf::from).collect::<Vec<_>>(),
            &items,
        )?;
        for keyword in ["QaM", "", "no-such-wincent-keyword-999999"] {
            let expected: Vec<_> = items
                .iter()
                .filter(|item| item.to_lowercase().contains(&keyword.to_lowercase()))
                .cloned()
                .collect();
            checks.check(
                &format!("query.find.{category:?}.{keyword}"),
                manager.find_items_containing(keyword, category)? == expected,
                &expected,
            )?;
            checks.check(
                &format!("query.keyword.{category:?}.{keyword}"),
                manager.contains_keyword(keyword, category)? == !expected.is_empty(),
                &expected,
            )?;
        }
    }
    for (category, actual) in [
        (QuickAccess::RecentFiles, &recent),
        (QuickAccess::FrequentFolders, &frequent),
    ] {
        // This is an independent live COM enumeration, not another DestList query.
        let mut last = String::new();
        let oracle = wait_for(|| {
            let items = shell::query(category)?;
            last = items.join("\n");
            Ok(same_paths(actual, &items))
        });
        fs::write(root.join(format!("shell-{category:?}.txt")), &last)?;
        checks.require(
            &format!("query.shell_order.{category:?}"),
            oracle.is_ok(),
            (oracle, actual, last),
        )?;
    }
    for (category, file, visible) in [
        (QuickAccess::RecentFiles, recent_files_dest_path()?, &recent),
        (
            QuickAccess::FrequentFolders,
            frequent_folders_dest_path()?,
            &frequent,
        ),
    ] {
        let parsed = parse_dest_file(file)?;
        fs::write(
            root.join(format!("raw-{category:?}.txt")),
            format!("{:#?}", parsed.dest_list()),
        )?;
        for (index, entry) in parsed.cfb_entries().iter().enumerate() {
            let expected = visible
                .iter()
                .position(|item| normalized(item) == normalized(entry.path()));
            let status = manager.get_item_status(entry.path())?;
            let (position, pinned) = if category == QuickAccess::RecentFiles {
                (status.recent_file(), true)
            } else {
                (
                    status.frequent_folder().map(|slot| slot.position()),
                    status
                        .frequent_folder()
                        .is_none_or(|slot| slot.is_pinned() == entry.is_pinned()),
                )
            };
            checks.check(
                &format!("status.raw.{category:?}.{index}"),
                position == expected && pinned,
                (entry.path(), expected, status),
            )?;
        }
    }
    checks.check(
        "query.recent_window",
        recent.len() == DEFAULT_RECENT_FILES_RESULT_LIMIT,
        (
            "25 native fixtures; validate current library window against live Shell",
            recent.len(),
        ),
    )?;
    Ok(())
}

fn path_and_duplicate_checks(
    checks: &mut Checks,
    manager: &QuickAccessManager,
    path: &Path,
    category: QuickAccess,
) -> TestResult {
    let text = path.to_string_lossy();
    let status = manager.get_item_status(path)?;
    let mut variants = vec![
        text.to_uppercase(),
        text.replace('\\', "/"),
        path.canonicalize()?.to_string_lossy().into_owned(),
    ];
    if path.is_dir() {
        variants.push(format!("{text}\\"));
        variants.push(format!("{text}\\."));
    }
    for (index, variant) in variants.iter().enumerate() {
        checks.check(
            &format!("path.variant.{category:?}.{index}"),
            manager.contains_path(variant, category)?
                && manager.get_item_status(variant)? == status,
            variant,
        )?;
    }
    checks.check(
        &format!("path.not_substring.{category:?}"),
        !manager.contains_path("QAM", category)?,
        "exact membership is not keyword matching",
    )?;
    let duplicate = manager.add_item(&variants[0]);
    checks.check(
        &format!("errors.duplicate.{category:?}"),
        matches!(&duplicate, Err(WincentError::AlreadyExists { qa_type, .. })
            if *qa_type == category),
        duplicate,
    )?;
    checks.check(
        &format!("status.lists.{category:?}"),
        status.is_visible() && status.lists() == Some(category),
        status,
    )?;
    Ok(())
}

fn lock_checks(
    checks: &mut Checks,
    manager: &QuickAccessManager,
    root: &Path,
    file: &Path,
    folder: &Path,
) -> TestResult {
    let probe_file = root.join("targets").join("lock-probe.txt");
    let probe_folder = root.join("targets").join("lock-probe-folder");
    fs::write(&probe_file, b"blocked add")?;
    fs::create_dir(&probe_folder)?;
    for category in CATEGORIES {
        manager.lock(category)?;
        let clone = manager.clone();
        let repeated = clone.lock(QuickAccess::All);
        checks.check(
            &format!("lock.clone_already_locked.{category:?}"),
            matches!(repeated, Err(WincentError::ManagerAlreadyLocked)),
            repeated,
        )?;
        checks.check(
            &format!("lock.query_allowed.{category:?}"),
            manager.contains_path(file, QuickAccess::RecentFiles)?
                && clone.contains_path(folder, QuickAccess::FrequentFolders)?,
            "queries remain available while locked",
        )?;
        let other = QuickAccessManager::new();
        let add = other.add_item(&probe_file);
        checks.check(
            &format!("lock.add_file.{category:?}"),
            matches!(add, Err(WincentError::QuickAccessLocked { .. })),
            add,
        )?;
        let target = if category == QuickAccess::FrequentFolders {
            folder
        } else {
            file
        };
        let remove = other.remove_item(target);
        checks.check(
            &format!("lock.remove.{category:?}"),
            matches!(remove, Err(WincentError::QuickAccessLocked { .. })),
            remove,
        )?;
        for clear in CATEGORIES
            .into_iter()
            .filter(|clear| *clear == category || *clear == QuickAccess::All)
        {
            let result = other.empty_items(clear);
            checks.check(
                &format!("lock.clear.{category:?}.{clear:?}"),
                matches!(result, Err(WincentError::QuickAccessLocked { .. })),
                result,
            )?;
        }
        if category != QuickAccess::RecentFiles {
            let add = other.add_item(&probe_folder);
            checks.check(
                &format!("lock.add_folder.{category:?}"),
                matches!(add, Err(WincentError::QuickAccessLocked { .. })),
                add,
            )?;
        }
        let backing = if category == QuickAccess::FrequentFolders {
            frequent_folders_dest_path()?
        } else {
            recent_files_dest_path()?
        };
        let write = OpenOptions::new().write(true).open(backing);
        checks.check(
            &format!("lock.backing_write_denied.{category:?}"),
            write.as_ref().err().and_then(std::io::Error::raw_os_error) == Some(32),
            write,
        )?;
        let log = File::create(root.join(format!("lock-child-{category:?}.log")))?;
        let mut child = Command::new(env::current_exe()?)
            .args(["--ignored", "--exact", "lock_probe", "--nocapture"])
            .env("WINCENT_QAM_LOCK", format!("{category:?}"))
            .stdout(Stdio::from(log.try_clone()?))
            .stderr(Stdio::from(log))
            .spawn()?;
        let start = Instant::now();
        let child_ok = loop {
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
            &format!("lock.cross_process.{category:?}"),
            child_ok,
            "separate process must reject overlapping mutation",
        )?;
        let unlocked = clone.unlock()?;
        checks.check(
            &format!("lock.unlock.{category:?}"),
            unlocked.iter().all(|path| !Path::new(path).exists()),
            unlocked,
        )?;
        let again = manager.unlock();
        checks.check(
            &format!("lock.double_unlock.{category:?}"),
            matches!(again, Err(WincentError::ManagerNotLocked)),
            again,
        )?;
    }
    let owner = QuickAccessManager::new();
    owner.lock(QuickAccess::RecentFiles)?;
    let survivor = owner.clone();
    drop(owner);
    let blocked = manager.add_item(&probe_file);
    checks.check(
        "lock.clone_retains_lock",
        matches!(blocked, Err(WincentError::QuickAccessLocked { .. })),
        blocked,
    )?;
    let separate = QuickAccessManager::new();
    separate.lock(QuickAccess::FrequentFolders)?;
    checks.check("lock.disjoint_categories", true, "independent managers")?;
    separate.unlock()?;
    drop(survivor);
    manager.lock(QuickAccess::All)?;
    let recent = recent_files_dest_path()?;
    let recent_dir = recent
        .parent()
        .and_then(Path::parent)
        .ok_or("Recent parent")?;
    let link = recent_dir.join(format!("wincent-qam-{}.lnk", std::process::id()));
    // A new shortcut need not be a valid Shell link: unlock tracks new .lnk paths.
    fs::write(&link, b"VM-only unlock path tracking fixture")?;
    let links = manager.unlock()?;
    checks.check(
        "lock.last_clone_drop_releases_and_unlock_reports_links",
        links.iter().any(|value| Path::new(value) == link) && !link.exists(),
        links,
    )?;
    Ok(())
}

fn mutations_and_clear(
    checks: &mut Checks,
    manager: &QuickAccessManager,
    root: &Path,
) -> TestResult {
    manager.empty_items(QuickAccess::All)?;
    let targets = root.join("targets");
    let file = targets.join("mutation \u{6587}\u{4ef6}.txt");
    let folder = targets.join("mutation \u{76ee}\u{5f55}");
    fs::write(&file, b"native mutation round trip")?;
    fs::create_dir(&folder)?;
    for (path, category) in [
        (&file, QuickAccess::RecentFiles),
        (&folder, QuickAccess::FrequentFolders),
    ] {
        manager.add_item(path)?;
        membership(
            checks,
            manager,
            &format!("mutation.add.{category:?}"),
            path,
            category,
            true,
        )?;
        manager.refresh()?;
        checks.check(
            &format!("refresh.populated.{category:?}"),
            true,
            "API returned Ok",
        )?;
        manager.remove_item(path)?;
        membership(
            checks,
            manager,
            &format!("mutation.remove.{category:?}"),
            path,
            category,
            false,
        )?;
        let status = manager.get_item_status(path)?;
        checks.check(
            &format!("mutation.status_removed.{category:?}"),
            !status.is_visible(),
            status,
        )?;
    }
    // Each clear category starts with populated file AND folder fixtures.
    for clear in CATEGORIES {
        // A removed Recent path may remain blocked in Shell's cached state.
        // Independent cases must not reuse that path, even after clearing history.
        let file = targets.join(format!("clear-{clear:?}.txt"));
        let folder = targets.join(format!("clear-{clear:?}-folder"));
        fs::write(&file, b"fresh clear fixture")?;
        fs::create_dir(&folder)?;
        manager.empty_items(QuickAccess::All)?;
        manager.add_item(&file)?;
        manager.add_item(&folder)?;
        membership(
            checks,
            manager,
            &format!("clear.seed_file.{clear:?}"),
            &file,
            QuickAccess::RecentFiles,
            true,
        )?;
        membership(
            checks,
            manager,
            &format!("clear.seed_folder.{clear:?}"),
            &folder,
            QuickAccess::FrequentFolders,
            true,
        )?;
        manager.empty_items(clear)?;
        let cleared = wait_for(|| {
            let file_visible = manager.get_item_status(&file)?.recent_file().is_some();
            let folder_visible = manager
                .get_item_status(&folder)?
                .frequent_folder()
                .is_some();
            Ok(file_visible == (clear == QuickAccess::FrequentFolders)
                && folder_visible == (clear == QuickAccess::RecentFiles))
        });
        checks.require(
            &format!("clear.scoped_status.{clear:?}"),
            cleared.is_ok(),
            cleared,
        )?;
        for (path, category, expected) in [
            (
                &file,
                QuickAccess::RecentFiles,
                clear == QuickAccess::FrequentFolders,
            ),
            (
                &folder,
                QuickAccess::FrequentFolders,
                clear == QuickAccess::RecentFiles,
            ),
        ] {
            let oracle = wait_for(|| {
                let items = shell::query(category)?;
                Ok(items
                    .iter()
                    .any(|item| normalized(item) == normalized(&path.to_string_lossy()))
                    == expected)
            });
            checks.require(
                &format!("clear.shell.{clear:?}.{category:?}"),
                oracle.is_ok(),
                oracle,
            )?;
        }
        manager.empty_items(clear)?;
        checks.check(
            &format!("clear.repeat.{clear:?}"),
            true,
            "idempotent API result",
        )?;
    }
    manager.refresh()?;
    checks.check(
        "refresh.empty",
        true,
        "API returned Ok, not pixel-level proof",
    )?;
    snapshot(root, "after-clear", manager)?;
    Ok(())
}

fn suite(checks: &mut Checks, root: &Path) -> TestResult {
    constructors(checks)?;
    let manager = QuickAccessManager::builder().try_timeout(WAIT)?.build();
    visibility(checks, &manager)?;
    manager.empty_items(QuickAccess::All)?;
    let targets = root.join("targets");
    fs::create_dir(&targets)?;
    let (files, folders) = fixture(checks, &manager, &targets)?;
    invalid_inputs(checks, &manager, &targets)?;
    queries(checks, &manager, root)?;
    snapshot(root, "fixture", &manager)?;
    let visible_file = files.last().ok_or("no fixture files")?;
    path_and_duplicate_checks(checks, &manager, visible_file, QuickAccess::RecentFiles)?;
    path_and_duplicate_checks(checks, &manager, &folders[0], QuickAccess::FrequentFolders)?;
    lock_checks(checks, &manager, root, visible_file, &folders[0])?;
    // Native file visits also exercise the unpinned parent-folder remove path.
    let parent = visible_file.parent().ok_or("fixture parent")?;
    let unpinned = wait_for(|| {
        Ok(manager
            .get_item_status(parent)?
            .frequent_folder()
            .is_some_and(|slot| !slot.is_pinned()))
    });
    checks.require("fixture.native_unpinned_parent", unpinned.is_ok(), unpinned)?;
    manager.remove_item(parent)?;
    membership(
        checks,
        &manager,
        "mutation.remove_unpinned",
        parent,
        QuickAccess::FrequentFolders,
        false,
    )?;
    mutations_and_clear(checks, &manager, root)?;
    Ok(())
}

#[test]
#[ignore = "destructive VM-only suite; use the dedicated PowerShell VM runner"]
fn quick_access_manager() -> TestResult {
    let root = guarded_root()?;
    let mut checks = Checks::new(&root)?;
    let result =
        std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| suite(&mut checks, &root)));
    let completed = matches!(&result, Ok(Ok(())));
    checks.check("suite.completed", completed, &result)?;
    if checks.failures != 0 {
        return Err(format!("{} / {} checks failed", checks.failures, checks.count).into());
    }
    Ok(())
}

#[test]
#[ignore = "VM-only child process used by quick_access_manager"]
fn lock_probe() -> TestResult {
    let root = guarded_root()?;
    let category = match env::var("WINCENT_QAM_LOCK")?.as_str() {
        "RecentFiles" => QuickAccess::RecentFiles,
        "FrequentFolders" => QuickAccess::FrequentFolders,
        "All" => QuickAccess::All,
        other => return Err(format!("unknown lock probe category: {other}").into()),
    };
    let manager = QuickAccessManager::new();
    for result in [
        manager.add_item(root.join("targets").join("lock-probe.txt")),
        manager.empty_items(category),
        manager.empty_items(QuickAccess::All),
    ] {
        if !matches!(result, Err(WincentError::QuickAccessLocked { .. })) {
            return Err(format!("cross-process lock not enforced: {result:?}").into());
        }
    }
    Ok(())
}
