use std::io::{self, Write};
use std::path::Path;
use std::time::Duration;
use tokio::time::sleep;
use wincent::{
    error::WincentError,
    predule::{QuickAccess, QuickAccessManager, WincentResult},
};

// Console color codes
const GREEN: &str = "\x1b[32m";
const RED: &str = "\x1b[31m";
const YELLOW: &str = "\x1b[33m";
const BLUE: &str = "\x1b[34m";
const RESET: &str = "\x1b[0m";
const BOLD: &str = "\x1b[1m";

// Console symbols
const CHECK_MARK: &str = "✓";
const CROSS_MARK: &str = "✗";
// const ARROW: &str = "→";
const SPINNER_CHARS: [&str; 4] = ["⠋", "⠙", "⠹", "⠸"];

struct ConsoleSpinner {
    message: String,
    current: usize,
}

impl ConsoleSpinner {
    fn new(message: &str) -> Self {
        Self {
            message: message.to_string(),
            current: 0,
        }
    }

    fn spin(&mut self) {
        print!("\r{} {} ", SPINNER_CHARS[self.current], self.message);
        io::stdout().flush().unwrap();
        self.current = (self.current + 1) % SPINNER_CHARS.len();
    }

    fn complete(&self, success: bool, message: &str) {
        let symbol = if success {
            format!("{}{}{}", GREEN, CHECK_MARK, RESET)
        } else {
            format!("{}{}{}", RED, CROSS_MARK, RESET)
        };

        println!("\r{} {}", symbol, message);
    }
}

// Display welcome screen
fn show_welcome() {
    println!("{}", RESET);
    println!(
        "{}{}╔══════════════════════════════════════════════════╗{}",
        BLUE, BOLD, RESET
    );
    println!(
        "{}{}║                                                  ║{}",
        BLUE, BOLD, RESET
    );
    println!(
        "{}{}║   __      __.__                      __          ║{}",
        BLUE, BOLD, RESET
    );
    println!(
        "{}{}║  /  \\    /  \\__| ____   ____   _____/  |_        ║{}",
        BLUE, BOLD, RESET
    );
    println!(
        "{}{}║  \\   \\/\\/   /  |/    \\_/ ___\\_/ __ \\   __\\       ║{}",
        BLUE, BOLD, RESET
    );
    println!(
        "{}{}║   \\        /|  |   |  \\  \\___\\  ___/|  |         ║{}",
        BLUE, BOLD, RESET
    );
    println!(
        "{}{}║    \\__/\\  / |__|___|  /\\___  >\\___  >__|         ║{}",
        BLUE, BOLD, RESET
    );
    println!(
        "{}{}║         \\/          \\/     \\/     \\/              ║{}",
        BLUE, BOLD, RESET
    );
    println!(
        "{}{}║                                                  ║{}",
        BLUE, BOLD, RESET
    );
    println!(
        "{}{}║           Windows Quick Access Manager           ║{}",
        BLUE, BOLD, RESET
    );
    println!(
        "{}{}║                                                  ║{}",
        BLUE, BOLD, RESET
    );
    println!(
        "{}{}╚══════════════════════════════════════════════════╝{}",
        BLUE, BOLD, RESET
    );
    println!();
}

// Display main menu
fn show_main_menu() {
    println!("\n{}{}Select Operation:{}", YELLOW, BOLD, RESET);
    println!("{}1. Check Execution Policy Status", BLUE);
    println!("{}2. Manage Quick Access Items", BLUE);
    println!("{}3. View Quick Access Items", BLUE);
    println!("{}4. Clear Quick Access Items", BLUE);
    println!("{}0. Exit Program{}", BLUE, RESET);
    print!("\n{}Enter choice [0-4]: {}", YELLOW, RESET);
    io::stdout().flush().unwrap();
}

// Display item management submenu
fn show_item_management_menu() {
    println!("\n{}{}Manage Quick Access Items:{}", YELLOW, BOLD, RESET);
    println!("{}1. Add file to Recent Files", BLUE);
    println!("{}2. Pin folder to Frequent Folders", BLUE);
    println!("{}3. Remove file from Recent Files", BLUE);
    println!("{}4. Unpin folder from Frequent Folders", BLUE);
    println!("{}0. Return to main menu{}", BLUE, RESET);
    print!("\n{}Enter choice [0-4]: {}", YELLOW, RESET);
    io::stdout().flush().unwrap();
}

// Display query submenu
fn show_query_menu() {
    println!("\n{}{}List Quick Access Items:{}", YELLOW, BOLD, RESET);
    println!("{}1. View Recent Files", BLUE);
    println!("{}2. View Frequent Folders", BLUE);
    println!("{}3. View All Quick Access Items", BLUE);
    println!("{}0. Return to main menu{}", BLUE, RESET);
    print!("\n{}Enter choice [0-3]: {}", YELLOW, RESET);
    io::stdout().flush().unwrap();
}

// Display clear submenu
fn show_empty_menu() {
    println!("\n{}{}Clear Quick Access Items:{}", YELLOW, BOLD, RESET);
    println!("{}1. Clear Recent Files", BLUE);
    println!("{}2. Clear Frequent Folders", BLUE);
    println!("{}3. Clear All Quick Access Items", BLUE);
    println!("{}0. Return to main menu{}", BLUE, RESET);
    print!("\n{}Enter choice [0-3]: {}", YELLOW, RESET);
    io::stdout().flush().unwrap();
}

// Read user input
fn read_input() -> String {
    let mut input = String::new();
    io::stdin()
        .read_line(&mut input)
        .expect("Failed to read input");
    input.trim().to_string()
}

// Read path input with prompt
fn read_path_input(prompt: &str) -> String {
    print!("{}{}: {}", YELLOW, prompt, RESET);
    io::stdout().flush().unwrap();
    read_input()
}

// Wait for any key press
fn wait_for_key() {
    println!("\n{}Precess any button to continue...{}", YELLOW, RESET);
    let _ = read_input();
}

// Check execution policy status
async fn check_feasibility(manager: &QuickAccessManager) -> WincentResult<()> {
    let mut spinner = ConsoleSpinner::new("Checking execution policy status...");

    for _ in 0..10 {
        spinner.spin();
        sleep(Duration::from_millis(100)).await;
    }

    let (query_feasible, handle_feasible) = manager.check_feasible().await;

    if query_feasible && handle_feasible {
        spinner.complete(true, "All operations are allowed");
    } else {
        spinner.complete(
            false,
            "Some operations may be restricted, please check system settings",
        );
    }

    Ok(())
}

// Add file to Recent Files
async fn add_file_to_recent(manager: &QuickAccessManager) -> WincentResult<()> {
    let path = read_path_input("Enter file path to add");

    if path.is_empty() {
        println!("{}Path cannot be empty{}", RED, RESET);
        return Ok(());
    }

    if !Path::new(&path).exists() {
        println!("{}File not found: {}{}", RED, path, RESET);
        return Ok(());
    }

    let mut spinner = ConsoleSpinner::new("Adding file to Recent Files...");

    for _ in 0..10 {
        spinner.spin();
        sleep(Duration::from_millis(100)).await;
    }

    match manager.add_item(&path, QuickAccess::RecentFiles, false).await {
        Ok(_) => {
            spinner.complete(true, &format!("Successfully added file: {}", path));
            Ok(())
        }
        Err(e) => {
            spinner.complete(false, &format!("Failed to add file: {}", e));
            Err(e)
        }
    }
}

// Pin folder to Frequent Folders
async fn pin_folder_to_frequent(manager: &QuickAccessManager) -> WincentResult<()> {
    let path = read_path_input("Enter folder path to pin");

    if path.is_empty() {
        println!("{}Path cannot be empty{}", RED, RESET);
        return Ok(());
    }

    if !Path::new(&path).exists() {
        println!("{}Folder not found: {}{}", RED, path, RESET);
        return Ok(());
    }

    let mut spinner = ConsoleSpinner::new("Pinning folder to Frequent Folders...");

    for _ in 0..10 {
        spinner.spin();
        sleep(Duration::from_millis(100)).await;
    }

    match manager.add_item(&path, QuickAccess::FrequentFolders, false).await {
        Ok(_) => {
            spinner.complete(true, &format!("Successfully pinned folder: {}", path));
            Ok(())
        }
        Err(e) => {
            spinner.complete(false, &format!("Failed to pin folder: {}", e));
            Err(e)
        }
    }
}

// Remove file from Recent Files
async fn remove_file_from_recent(manager: &QuickAccessManager) -> WincentResult<()> {
    let path = read_path_input("Enter file path to remove");

    if path.is_empty() {
        println!("{}Path cannot be empty{}", RED, RESET);
        return Ok(());
    }

    let mut spinner = ConsoleSpinner::new("Removing file from Recent Files...");

    for _ in 0..10 {
        spinner.spin();
        sleep(Duration::from_millis(100)).await;
    }

    match manager.remove_item(&path, QuickAccess::RecentFiles).await {
        Ok(_) => {
            spinner.complete(true, &format!("Successfully removed file: {}", path));
            Ok(())
        }
        Err(e) => {
            spinner.complete(false, &format!("Failed to remove file: {}", e));
            Err(e)
        }
    }
}

// Unpin folder from Frequent Folders
async fn unpin_folder_from_frequent(manager: &QuickAccessManager) -> WincentResult<()> {
    let path = read_path_input("Enter folder path to unpin");

    if path.is_empty() {
        println!("{}Path cannot be empty{}", RED, RESET);
        return Ok(());
    }

    let mut spinner = ConsoleSpinner::new("Unpinning folder from Frequent Folders...");

    for _ in 0..10 {
        spinner.spin();
        sleep(Duration::from_millis(100)).await;
    }

    match manager
        .remove_item(&path, QuickAccess::FrequentFolders)
        .await
    {
        Ok(_) => {
            spinner.complete(true, &format!("Successfully unpinned folder: {}", path));
            Ok(())
        }
        Err(e) => {
            spinner.complete(false, &format!("Failed to unpin folder: {}", e));
            Err(e)
        }
    }
}

// Query and display items
async fn query_and_display_items(
    manager: &QuickAccessManager,
    qa_type: QuickAccess,
) -> WincentResult<()> {
    let type_name = match qa_type {
        QuickAccess::RecentFiles => "Recent Files",
        QuickAccess::FrequentFolders => "Frequent Folders",
        QuickAccess::All => "All Quick Access Items",
    };

    let mut spinner = ConsoleSpinner::new(&format!("Querying {}...", type_name));

    for _ in 0..10 {
        spinner.spin();
        sleep(Duration::from_millis(100)).await;
    }

    match manager.get_items(qa_type).await {
        Ok(items) => {
            spinner.complete(true, &format!("Successfully retrieved {} list", type_name));

            println!(
                "\n{}{}{} ({} items):{}",
                YELLOW,
                BOLD,
                type_name,
                items.len(),
                RESET
            );

            if items.is_empty() {
                println!("{}List is empty{}", YELLOW, RESET);
            } else {
                for (i, item) in items.iter().enumerate() {
                    println!("{}{}. {}{}", BLUE, i + 1, item, RESET);
                }
            }

            Ok(())
        }
        Err(e) => {
            spinner.complete(false, &format!("Failed to query {}: {}", type_name, e));
            Err(e)
        }
    }
}

// Clear items
async fn empty_items(manager: &QuickAccessManager, qa_type: QuickAccess) -> WincentResult<()> {
    let type_name = match qa_type {
        QuickAccess::RecentFiles => "Recent Files",
        QuickAccess::FrequentFolders => "Frequent Folders",
        QuickAccess::All => "All Quick Access Items",
    };

    print!("{}Confirm to clear {}? (y/n): {}", YELLOW, type_name, RESET);
    io::stdout().flush().unwrap();

    let confirm = read_input().to_lowercase();

    if confirm != "y" && confirm != "yes" {
        println!("{}Operation cancelled{}", YELLOW, RESET);
        return Ok(());
    }

    let mut spinner = ConsoleSpinner::new(&format!("Clearing {}...", type_name));

    for _ in 0..10 {
        spinner.spin();
        sleep(Duration::from_millis(100)).await;
    }

    match manager.empty_items(qa_type, false).await {
        Ok(_) => {
            spinner.complete(true, &format!("Successfully cleared {}", type_name));
            Ok(())
        }
        Err(e) => {
            spinner.complete(false, &format!("Failed to clear {}: {}", type_name, e));
            Err(e)
        }
    }
}

// Handle item management menu
async fn handle_item_management(manager: &QuickAccessManager) -> WincentResult<()> {
    loop {
        show_item_management_menu();

        let choice = read_input();

        match choice.as_str() {
            "1" => {
                if let Err(e) = add_file_to_recent(manager).await {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
                wait_for_key();
            }
            "2" => {
                if let Err(e) = pin_folder_to_frequent(manager).await {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
                wait_for_key();
            }
            "3" => {
                if let Err(e) = remove_file_from_recent(manager).await {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
                wait_for_key();
            }
            "4" => {
                if let Err(e) = unpin_folder_from_frequent(manager).await {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
                wait_for_key();
            }
            "0" => break,
            _ => {
                println!("{}Invalid choice, please try again{}", RED, RESET);
                wait_for_key();
            }
        }
    }

    Ok(())
}

// Handle query menu
async fn handle_query_menu(manager: &QuickAccessManager) -> WincentResult<()> {
    loop {
        show_query_menu();

        let choice = read_input();

        match choice.as_str() {
            "1" => {
                if let Err(e) = query_and_display_items(manager, QuickAccess::RecentFiles).await {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
                wait_for_key();
            }
            "2" => {
                if let Err(e) = query_and_display_items(manager, QuickAccess::FrequentFolders).await
                {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
                wait_for_key();
            }
            "3" => {
                if let Err(e) = query_and_display_items(manager, QuickAccess::All).await {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
                wait_for_key();
            }
            "0" => break,
            _ => {
                println!("{}Invalid choice, please try again{}", RED, RESET);
                wait_for_key();
            }
        }
    }

    Ok(())
}

// Handle clear menu
async fn handle_empty_menu(manager: &QuickAccessManager) -> WincentResult<()> {
    loop {
        show_empty_menu();

        let choice = read_input();

        match choice.as_str() {
            "1" => {
                if let Err(e) = empty_items(manager, QuickAccess::RecentFiles).await {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
                wait_for_key();
            }
            "2" => {
                if let Err(e) = empty_items(manager, QuickAccess::FrequentFolders).await {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
                wait_for_key();
            }
            "3" => {
                if let Err(e) = empty_items(manager, QuickAccess::All).await {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
                wait_for_key();
            }
            "0" => break,
            _ => {
                println!("{}Invalid choice, please try again{}", RED, RESET);
                wait_for_key();
            }
        }
    }

    Ok(())
}

// Main function
#[tokio::main]
async fn main() -> Result<(), WincentError> {
    // Create QuickAccessManager instance
    let manager = QuickAccessManager::new().await?;

    show_welcome();

    loop {
        show_main_menu();

        let choice = read_input();

        match choice.as_str() {
            "1" => {
                if let Err(e) = check_feasibility(&manager).await {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
                wait_for_key();
            }
            "2" => {
                if let Err(e) = handle_item_management(&manager).await {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
            }
            "3" => {
                if let Err(e) = handle_query_menu(&manager).await {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
            }
            "4" => {
                if let Err(e) = handle_empty_menu(&manager).await {
                    println!("{}Error: {}{}", RED, e, RESET);
                }
            }
            "0" => {
                println!("\n{}{}Exiting program...{}", YELLOW, BOLD, RESET);

                let mut spinner = ConsoleSpinner::new("Cleaning up resources");

                for _ in 0..5 {
                    spinner.spin();
                    sleep(Duration::from_millis(200)).await;
                }

                spinner.complete(true, "Thanks for using Windows Quick Access Manager");
                break;
            }
            _ => {
                println!("{}Invalid option, please try again{}", RED, RESET);
                wait_for_key();
            }
        }
    }

    Ok(())
}
