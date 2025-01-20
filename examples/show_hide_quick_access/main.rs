use wincent::{
    visible::{
        is_frequent_folders_visible, is_recent_files_visiable, set_frequent_folders_visiable,
        set_recent_files_visiable,
    },
    WincentResult,
};

fn print_visibility_status() -> WincentResult<()> {
    println!(
        "Recent Files: {}",
        if is_recent_files_visiable()? {
            "Visible"
        } else {
            "Hidden"
        }
    );
    println!(
        "Frequent Folders: {}",
        if is_frequent_folders_visible()? {
            "Visible"
        } else {
            "Hidden"
        }
    );
    Ok(())
}

fn main() -> WincentResult<()> {
    // Show initial status
    println!("Initial visibility status:");
    print_visibility_status()?;

    // Save initial state
    let initial_recent = is_recent_files_visiable()?;
    let initial_folders = is_frequent_folders_visible()?;

    // Hide all sections
    println!("Hiding all sections...");
    set_recent_files_visiable(false)?;
    set_frequent_folders_visiable(false)?;

    println!("Status after hiding:");
    print_visibility_status()?;

    // Show all sections
    println!("Showing all sections...");
    set_recent_files_visiable(true)?;
    set_frequent_folders_visiable(true)?;

    println!("Status after showing:");
    print_visibility_status()?;

    // Set different visibility
    println!("Setting different visibility...");
    set_recent_files_visiable(false)?;
    set_frequent_folders_visiable(true)?;

    println!("Status after mixed settings:");
    print_visibility_status()?;

    // Restore initial state
    println!("Restoring initial visibility...");
    set_recent_files_visiable(initial_recent)?;
    set_frequent_folders_visiable(initial_folders)?;

    println!("Final status (restored to initial):");
    print_visibility_status()?;

    Ok(())
}
