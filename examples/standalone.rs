//! A standalone application: query, add, remove, lock and hide Quick Access items.
//!
//! Run `cargo run --example standalone` in a disposable VM with a snapshot.
//! This example changes history and visibility; it restores visibility settings,
//! but does not back up history. It never clears existing Quick Access categories.
//! A timeout does not cancel native work; inspect state before retrying.

use std::fs;
use std::path::Path;
use std::thread;
use std::time::{Duration, Instant};
use wincent::error::WincentError;
use wincent::standalone::*;
use wincent::visible::{is_visible, set_visible, VisibilityTarget};
use wincent::WincentResult;

fn main() -> WincentResult<()> {
    let timeout = Duration::from_secs(15);
    // Use a unique folder outside the system Temp directory, which Shell may filter.
    let folder = tempfile::Builder::new()
        .prefix("wincent-example-")
        .tempdir_in(std::env::current_dir()?)?;
    let file = folder.path().join("report.txt");
    fs::write(&file, b"Standalone Quick Access example")?;

    let recent_visible = is_visible(VisibilityTarget::RecentFiles)?;
    let frequent_visible = is_visible(VisibilityTarget::FrequentFolders)?;
    let result = (|| -> WincentResult<()> {
        set_visible(VisibilityTarget::All, true)?;
        println!("Recent files: {:?}", get_recent_files()?);
        println!("Frequent folders: {:?}", get_frequent_folders()?);

        add_to_recent_files(&file, timeout)?;
        pin_to_frequent_folders(folder.path(), timeout)?;

        // Shell success is not a persistence barrier. Poll queries, not mutations.
        let started = Instant::now();
        loop {
            let listed = get_recent_files()?
                .iter()
                .any(|path| Path::new(path) == file);
            if listed && is_folder_pinned(folder.path())? {
                break;
            }
            if started.elapsed() >= timeout {
                return Err(WincentError::Timeout(
                    "Waiting for example items to persist".into(),
                ));
            }
            thread::sleep(Duration::from_millis(100));
        }
        println!(
            "Example folder pinned: {}",
            is_folder_pinned(folder.path())?
        );

        remove_from_recent_files(&file, timeout)?;
        remove_from_frequent_folders(folder.path(), timeout)?;
        println!("Example items removed from Quick Access.");

        // Guards keep both categories locked while queries remain available.
        let recent_lock = lock_recent_files()?;
        let frequent_lock = lock_frequent_folders()?;
        println!("Recent files while locked: {:?}", get_recent_files()?);
        println!(
            "Frequent folders while locked: {:?}",
            get_frequent_folders()?
        );
        println!("Removed shortcuts: {:?}", unlock_recent_files(recent_lock)?);
        println!(
            "Removed shortcuts: {:?}",
            unlock_frequent_folders(frequent_lock)?
        );

        set_visible(VisibilityTarget::All, false)?;
        refresh_quick_access(timeout)?;
        println!(
            "Quick Access categories visible: {}",
            is_visible(VisibilityTarget::All)?
        );
        Ok(())
    })();

    // Attempt both restorations even if an earlier operation failed.
    let restore_recent = set_visible(VisibilityTarget::RecentFiles, recent_visible);
    let restore_frequent = set_visible(VisibilityTarget::FrequentFolders, frequent_visible);
    result?;
    restore_recent?;
    restore_frequent?;
    refresh_quick_access(timeout)?;
    println!("Original visibility settings restored.");
    Ok(())
}
