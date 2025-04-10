use std::{thread, time::Duration};
use tempfile::Builder;
use wincent::{
    handle::{add_to_frequent_folders, remove_from_frequent_folders},
    query::is_in_frequent_folders,
    WincentResult,
};

fn main() -> WincentResult<()> {
    // Create temporary folder
    let temp_dir = Builder::new().prefix("wincent-test-").tempdir()?;
    let dir_path = temp_dir.path().to_str().unwrap();

    println!("Working with temporary folder: {}", dir_path);

    // Pin folder to frequent folders
    println!("Pinning folder to Quick Access...");
    add_to_frequent_folders(dir_path)?;

    // Wait for Windows to update
    thread::sleep(Duration::from_millis(500));

    // Verify if folder has been pinned
    if is_in_frequent_folders(dir_path)? {
        println!("Folder successfully pinned to Quick Access");
    } else {
        println!("Failed to pin folder to Quick Access");
        return Ok(());
    }

    // Unpin folder from frequent folders
    println!("Unpinning folder from Quick Access...");
    remove_from_frequent_folders(dir_path)?;

    // Wait for Windows to update
    thread::sleep(Duration::from_millis(500));

    // Verify if folder has been unpinned
    if !is_in_frequent_folders(dir_path)? {
        println!("Folder successfully unpinned from Quick Access");
    } else {
        println!("Failed to unpin folder from Quick Access");
    }

    // Temporary folder will be automatically deleted when temp_dir goes out of scope
    Ok(())
}
