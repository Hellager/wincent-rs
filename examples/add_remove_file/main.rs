use std::io::Write;
use std::{thread, time::Duration};
use tempfile::Builder;
use wincent::{
    add_to_recent_files, check_script_feasible, fix_script_feasible, is_in_recent_files,
    remove_from_recent_files, WincentResult,
};

fn main() -> WincentResult<()> {
    // Check and ensure script execution feasibility
    if !check_script_feasible()? {
        println!("Fixing script execution policy...");
        fix_script_feasible()?;
    }

    // Create temporary file
    let temp_file = Builder::new()
        .prefix("wincent-test-")
        .suffix(".txt")
        .tempfile()?;

    // Write some test content
    writeln!(
        temp_file.as_file(),
        "This is a test file for Quick Access operations"
    )?;
    let file_path = temp_file.path().to_str().unwrap();

    println!("Working with temporary file: {}", file_path);

    // Add file to recent items
    println!("Adding file to Quick Access...");
    add_to_recent_files(file_path)?;

    // Wait for Windows to update
    thread::sleep(Duration::from_millis(500));

    // Verify if file has been added
    if is_in_recent_files(file_path)? {
        println!("File successfully added to Quick Access");
    } else {
        println!("Failed to add file to Quick Access");
        return Ok(());
    }

    // Remove file from recent items
    println!("Removing file from Quick Access...");
    remove_from_recent_files(file_path)?;

    // Wait for Windows to update
    thread::sleep(Duration::from_millis(500));

    // Verify if file has been removed
    if !is_in_recent_files(file_path)? {
        println!("File successfully removed from Quick Access");
    } else {
        println!("Failed to remove file from Quick Access");
    }

    // Temporary file will be automatically deleted when temp_file goes out of scope
    Ok(())
}
