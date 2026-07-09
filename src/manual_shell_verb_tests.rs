use crate::destlist::FrequentFolderPinStatus;
use crate::error::WincentError;
use crate::manager::QuickAccessManager;
use crate::query::{
    folder_items, item_path, query_recent, shell_folder, FREQUENT_FOLDERS_NAMESPACE,
};
use crate::utils::{is_windows_11_or_later, paths_equal};
use crate::{AddOptions, QuickAccess, WincentResult};
use std::path::Path;
use std::sync::{Mutex, MutexGuard};
use std::thread;
use std::time::Duration;
use windows::core::Interface;
use windows::Win32::System::Variant::VARIANT;
use windows::Win32::UI::Shell::Folder3;

const COM_TIMEOUT: Duration = Duration::from_secs(10);
const WAIT_RETRIES: usize = 20;
const WAIT_INTERVAL: Duration = Duration::from_millis(500);

static QUICK_ACCESS_TEST_LOCK: Mutex<()> = Mutex::new(());

fn given_pinned_folder() -> &'static str {
    r"G:\Github\wincent-rs\temp"
}

fn given_unpinned_folder() -> &'static str {
    r"G:\Github\scourgify\src-tauri\icons\dark"
}

fn lock_quick_access_tests() -> MutexGuard<'static, ()> {
    QUICK_ACCESS_TEST_LOCK
        .lock()
        .expect("manual Quick Access test lock was poisoned")
}

fn skip_unless_windows_10() -> WincentResult<bool> {
    if is_windows_11_or_later()? {
        println!("Skipping Win10-only manual test on Win11+");
        return Ok(false);
    }

    Ok(true)
}

fn skip_unless_windows_11_or_later() -> WincentResult<bool> {
    if !is_windows_11_or_later()? {
        println!("Skipping Win11+-only manual test on Win10");
        return Ok(false);
    }

    Ok(true)
}

fn assert_existing_directory(path: &str) {
    assert!(
        Path::new(path).is_dir(),
        "manual diagnostic path must be an existing directory: {path}"
    );
}

fn manager() -> QuickAccessManager {
    QuickAccessManager::new()
}

fn cleanup_frequent_folder(path: &str) {
    let _ = manager().remove_item(path, QuickAccess::FrequentFolders);
}

fn wait_for_frequent_presence(path: &str, expected: bool) -> WincentResult<bool> {
    for _ in 0..WAIT_RETRIES {
        let folders = query_recent(QuickAccess::FrequentFolders)?;
        let exists = folders.iter().any(|item| paths_equal(item, path));
        if exists == expected {
            return Ok(true);
        }

        thread::sleep(WAIT_INTERVAL);
    }

    Ok(false)
}

fn wait_for_pin_status(path: &str, expected: FrequentFolderPinStatus) -> WincentResult<bool> {
    for _ in 0..WAIT_RETRIES {
        if manager().frequent_folder_pin_status(path)? == expected {
            return Ok(true);
        }

        thread::sleep(WAIT_INTERVAL);
    }

    Ok(false)
}

fn prepare_pinned_folder() -> WincentResult<&'static str> {
    let path = given_pinned_folder();
    assert_existing_directory(path);
    cleanup_frequent_folder(path);
    manager().add_item(path, QuickAccess::FrequentFolders, AddOptions::new())?;

    assert!(
        wait_for_pin_status(path, FrequentFolderPinStatus::Pinned)?,
        "expected {path} to become pinned in Frequent Folders"
    );

    Ok(path)
}

fn require_unpinned_frequent_folder() -> WincentResult<&'static str> {
    let path = given_unpinned_folder();
    assert_existing_directory(path);

    assert_eq!(
        manager().frequent_folder_pin_status(path)?,
        FrequentFolderPinStatus::Unpinned,
        "manual test precondition failed: {path} must already be an unpinned Frequent Folders entry"
    );

    Ok(path)
}

fn invoke_pintohome(path: &str) -> WincentResult<()> {
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread(
        move || {
            let folder = shell_folder(&path)?;

            // SAFETY: This runs on a COM-initialized STA worker. `folder` is a live
            // Shell Folder interface and the verb VARIANT lives through InvokeVerb.
            unsafe {
                let folder3: Folder3 = folder.cast().map_err(|error| {
                    WincentError::SystemError(format!(
                        "Failed to cast to Folder3 for {path}: {error}"
                    ))
                })?;
                let self_item = folder3.Self_().map_err(|error| {
                    WincentError::SystemError(format!("Failed to get Self for {path}: {error}"))
                })?;
                let verb = VARIANT::from("pintohome");
                self_item.InvokeVerb(&verb).map_err(|error| {
                    WincentError::SystemError(format!(
                        "Failed to invoke pintohome on {path}: {error}"
                    ))
                })?;
            }

            Ok(())
        },
        COM_TIMEOUT,
    )
}

fn invoke_unpinfromhome(path: &str) -> WincentResult<bool> {
    let path = path.to_owned();
    crate::com_thread::run_on_sta_thread(
        move || {
            let folder = shell_folder(FREQUENT_FOLDERS_NAMESPACE)?;
            let items = folder_items(&folder)?;

            // SAFETY: `items` is a live FolderItems collection on this STA worker.
            let count = unsafe {
                items.Count().map_err(|error| {
                    WincentError::SystemError(format!("Failed to get item count: {error}"))
                })?
            };

            for index in 0..count {
                let index_variant = VARIANT::from(index);
                // SAFETY: The collection is live for this loop and the index VARIANT
                // lives until Item returns. Per-item failures are skipped.
                let item = unsafe {
                    match items.Item(&index_variant) {
                        Ok(item) => item,
                        Err(_) => continue,
                    }
                };

                if item_path(&item).is_some_and(|item_path| paths_equal(&item_path, &path)) {
                    let verb = VARIANT::from("unpinfromhome");
                    // SAFETY: `item` is a live FolderItem and the verb VARIANT
                    // remains alive for the synchronous InvokeVerb call.
                    unsafe {
                        item.InvokeVerb(&verb).map_err(|error| {
                            WincentError::SystemError(format!(
                                "Failed to invoke unpinfromhome on {path}: {error}"
                            ))
                        })?;
                    }
                    return Ok(true);
                }
            }

            Ok(false)
        },
        COM_TIMEOUT,
    )
}

#[test]
#[ignore = "Manual Win10 diagnostic; mutates Frequent Folders. Run: cargo test win10_unpinfromhome_removes_pinned_frequent_folder -- --ignored --nocapture"]
fn win10_unpinfromhome_removes_pinned_frequent_folder() -> WincentResult<()> {
    let _lock = lock_quick_access_tests();
    if !skip_unless_windows_10()? {
        return Ok(());
    }

    let path = prepare_pinned_folder()?;
    assert!(
        invoke_unpinfromhome(path)?,
        "target was not found in Frequent Folders: {path}"
    );
    assert!(
        wait_for_frequent_presence(path, false)?,
        "unpinfromhome did not remove pinned Frequent Folders entry on Win10: {path}"
    );

    Ok(())
}

#[test]
#[ignore = "Manual Win10 diagnostic; requires given_unpinned_folder() to already be unpinned frequent. Run: cargo test win10_remove_removes_unpinned_frequent_folder -- --ignored --nocapture"]
fn win10_remove_removes_unpinned_frequent_folder() -> WincentResult<()> {
    let _lock = lock_quick_access_tests();
    if !skip_unless_windows_10()? {
        return Ok(());
    }

    let path = require_unpinned_frequent_folder()?;
    manager().remove_item(path, QuickAccess::FrequentFolders)?;
    assert!(
        wait_for_frequent_presence(path, false)?,
        "remove did not remove unpinned Frequent Folders entry on Win10: {path}"
    );

    Ok(())
}

#[test]
#[ignore = "Manual Win10 diagnostic; mutates Frequent Folders. Run: cargo test win10_pintohome_is_not_frequent_folder_toggle -- --ignored --nocapture"]
fn win10_pintohome_is_not_frequent_folder_toggle() -> WincentResult<()> {
    let _lock = lock_quick_access_tests();
    if !skip_unless_windows_10()? {
        return Ok(());
    }

    let path = given_unpinned_folder();
    assert_existing_directory(path);
    cleanup_frequent_folder(path);

    invoke_pintohome(path)?;
    assert!(
        wait_for_frequent_presence(path, true)?,
        "first pintohome did not add Frequent Folders entry on Win10: {path}"
    );

    invoke_pintohome(path)?;
    let still_present = wait_for_frequent_presence(path, true)?;
    cleanup_frequent_folder(path);

    assert!(
        still_present,
        "second pintohome removed {path} from Frequent Folders on Win10, so it behaved as a toggle"
    );

    Ok(())
}

#[test]
#[ignore = "Manual Win11+ diagnostic; mutates Frequent Folders. Run: cargo test win11_unpinfromhome_removes_pinned_frequent_folder -- --ignored --nocapture"]
fn win11_unpinfromhome_removes_pinned_frequent_folder() -> WincentResult<()> {
    let _lock = lock_quick_access_tests();
    if !skip_unless_windows_11_or_later()? {
        return Ok(());
    }

    let path = prepare_pinned_folder()?;
    assert!(
        invoke_unpinfromhome(path)?,
        "target was not found in Frequent Folders: {path}"
    );
    assert!(
        wait_for_frequent_presence(path, false)?,
        "unpinfromhome did not remove pinned Frequent Folders entry on Win11+: {path}"
    );

    Ok(())
}

#[test]
#[ignore = "Manual Win11+ diagnostic; requires given_unpinned_folder() to already be unpinned frequent. Run: cargo test win11_remove_removes_unpinned_frequent_folder -- --ignored --nocapture"]
fn win11_remove_removes_unpinned_frequent_folder() -> WincentResult<()> {
    let _lock = lock_quick_access_tests();
    if !skip_unless_windows_11_or_later()? {
        return Ok(());
    }

    let path = require_unpinned_frequent_folder()?;
    manager().remove_item(path, QuickAccess::FrequentFolders)?;
    assert!(
        wait_for_frequent_presence(path, false)?,
        "remove did not remove unpinned Frequent Folders entry on Win11+: {path}"
    );

    Ok(())
}

#[test]
#[ignore = "Manual Win11+ diagnostic; mutates Frequent Folders. Run: cargo test win11_pintohome_is_frequent_folder_toggle -- --ignored --nocapture"]
fn win11_pintohome_is_frequent_folder_toggle() -> WincentResult<()> {
    let _lock = lock_quick_access_tests();
    if !skip_unless_windows_11_or_later()? {
        return Ok(());
    }

    let path = given_unpinned_folder();
    assert_existing_directory(path);
    cleanup_frequent_folder(path);

    invoke_pintohome(path)?;
    assert!(
        wait_for_frequent_presence(path, true)?,
        "first pintohome did not add Frequent Folders entry on Win11+: {path}"
    );

    invoke_pintohome(path)?;
    assert!(
        wait_for_frequent_presence(path, false)?,
        "second pintohome left {path} in Frequent Folders on Win11+, so it did not behave as a toggle"
    );

    Ok(())
}

// Manual Win11+ shell verb notes
//
// Version detection in `utils.rs` uses RtlGetVersion().dwBuildNumber and treats
// build >= 22000 as Win11+. On the host used for these diagnostics, registry
// data reported DisplayVersion=25H2, CurrentBuildNumber=26200, UBR=8655
// (10.0.26200.8655). `ProductName` still reported "Windows 10 Enterprise",
// so build number is the reliable classification signal here.
//
// Observed Win11+ Frequent Folders results:
// - win11_unpinfromhome_removes_pinned_frequent_folder:
//   `unpinfromhome` invoked on the matching item in the Frequent Folders
//   namespace successfully removed `given_pinned_folder()`.
// - win11_remove_removes_unpinned_frequent_folder:
//   QuickAccessManager::remove_item(..., QuickAccess::FrequentFolders)
//   successfully removed an unpinned frequent entry after it had been added.
// - win11_pintohome_is_frequent_folder_toggle:
//   the first `pintohome` added the folder to Frequent Folders and the second
//   `pintohome` removed it, confirming that Win11+ treats `pintohome` as a
//   toggle for Frequent Folders state.
//
// Practical guidance for Win11+:
// - `unpinfromhome` is valid for removing a pinned Frequent Folders item when
//   the target item is found through the Frequent Folders namespace.
// - `pintohome` is valid for adding a folder only after an existence check;
//   blindly invoking it on an existing item can remove the item instead.
// - Unpinned frequent entries need special handling. The production remove
//   state machine should continue to verify presence after each verb and may
//   need to pin/toggle before the item disappears.
// - Do not trust InvokeVerb success alone. Explorer updates this state
//   asynchronously, so tests and production code should poll the namespace or
//   DestList-backed status after mutations.
//
// Observed Win10 Frequent Folders results:
// - win10_unpinfromhome_removes_pinned_frequent_folder:
//   `unpinfromhome` invoked on the matching pinned item in the Frequent Folders
//   namespace successfully removed `given_pinned_folder()`.
// - win10_remove_removes_unpinned_frequent_folder:
//   QuickAccessManager::remove_item(..., QuickAccess::FrequentFolders)
//   successfully removed an unpinned frequent entry.
// - win10_pintohome_is_not_frequent_folder_toggle:
//   the second `pintohome` removed the unpinned frequent entry, so this host's
//   Win10 shell also behaved as a Frequent Folders toggle in that scenario.
//
// Practical guidance for Win10:
// - `unpinfromhome` is valid for removing a pinned Frequent Folders item when
//   the target item is found through the Frequent Folders namespace.
// - The public remove path successfully handles unpinned frequent entries on
//   the tested Win10 host.
// - `pintohome` behavior is not safe to model as add-only on all Win10 hosts;
//   keep the production path version-aware but verification-driven.
