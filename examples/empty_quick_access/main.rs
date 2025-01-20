use wincent::{
    empty::{empty_frequent_folders, empty_quick_access, empty_recent_files},
    WincentResult,
};

fn main() -> WincentResult<()> {
    // Example 1: Clear only recent files
    println!("Clearing recent files...");
    empty_recent_files()?;
    println!("Recent files cleared successfully");

    // Example 2: Clear frequent folders (both pinned and normal)
    println!("\nClearing frequent folders...");
    empty_frequent_folders()?;
    println!("Frequent folders cleared successfully");

    // Example 3: Clear everything in Quick Access
    println!("\nClearing entire Quick Access...");
    empty_quick_access()?;
    println!("Quick Access cleared successfully");

    // Example 4: Selective clearing with error handling
    println!("\nDemonstrating error handling...");
    match empty_recent_files() {
        Ok(_) => println!("Recent files cleared"),
        Err(e) => println!("Failed to clear recent files: {}", e),
    }

    Ok(())
}
