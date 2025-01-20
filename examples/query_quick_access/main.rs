use wincent::{
    feasible::{check_script_feasible, fix_script_feasible},
    query::{
        get_frequent_folders, get_quick_access_items, get_recent_files, is_in_frequent_folders,
        is_in_quick_access, is_in_recent_files,
    },
    WincentResult,
};

fn print_items(title: &str, items: &[String]) {
    println!("\n=== {} ===", title);
    if items.is_empty() {
        println!("No items found");
    } else {
        for (idx, item) in items.iter().enumerate() {
            println!("{}. {}", idx + 1, item);
        }
    }
    println!("=== End of {} ===\n", title);
}

fn main() -> WincentResult<()> {
    // Check and ensure script execution feasibility
    if !check_script_feasible()? {
        println!("Fixing script execution policy...");
        fix_script_feasible()?;
    }

    // Get all Quick Access items
    println!("Querying Quick Access items...");
    let all_items = get_quick_access_items()?;
    print_items("All Quick Access Items", &all_items);

    // Get recently used files
    let recent_files = get_recent_files()?;
    print_items("Recent Files", &recent_files);

    // Get frequent folders
    let frequent_folders = get_frequent_folders()?;
    print_items("Frequent Folders", &frequent_folders);

    // Search for specific keywords
    let keywords = ["Documents", "Downloads", "Desktop"];
    println!("\nSearching for specific keywords...");

    for keyword in keywords {
        println!("\nChecking for keyword: '{}'", keyword);

        if is_in_quick_access(keyword)? {
            println!("'{}' found in Quick Access", keyword);

            if is_in_recent_files(keyword)? {
                println!("'{}' found in Recent Files", keyword);
            }

            if is_in_frequent_folders(keyword)? {
                println!("'{}' found in Frequent Folders", keyword);
            }
        } else {
            println!("'{}' not found in Quick Access", keyword);
        }
    }

    Ok(())
}
