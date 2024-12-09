extern crate wincent;

use std::io::{self, Write, Error, ErrorKind};
use log::debug;
use wincent::{
    WincentError, QuickAccess, 
    check_feasible, fix_feasible, check_fix_feasible,
    is_in_quick_access, add_to_frequent_folders, remove_from_frequent_folders
};

fn ask_user(prompt: String) -> bool {
    let mut input = String::new();
    print!("{} ", prompt);
    io::stdout().flush().unwrap(); // Ensure the prompt is printed before reading input
    io::stdin().read_line(&mut input).expect("Failed to read line");
    let input = input.trim().to_lowercase();
    input == "y" || input == "yes"
}

#[tokio::main]
async fn main() -> Result<(), WincentError> {
    let _ = env_logger::builder()
        .target(env_logger::Target::Stdout)
        .filter_level(log::LevelFilter::Trace)
        .is_test(true)
        .try_init();

    let is_feasible = check_feasible()?;
    if !is_feasible{
        fix_feasible()?;

        if check_fix_feasible()? {
            debug!("fix script feasible success!");
        } else {
            debug!("fix acript feasible failed!!!");
            return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
        }
    }

    let current_dir = std::env::current_dir().expect("Failed to get current directory");
    let current_dir_str = current_dir.to_str().expect("Failed to convert path to string");

    if is_in_quick_access(vec![current_dir_str], Some(QuickAccess::FrequentFolders)).await? {
        if ask_user(format!("Do you want to remove the current folder '{}' from Quick Access? (y/n)", current_dir_str)) {
            remove_from_frequent_folders(current_dir_str).await?;
        }
    } else {
        if ask_user(format!("Do you want to pin the current folder '{}' to Quick Access? (y/n)", current_dir_str)) {
            add_to_frequent_folders(current_dir_str).await?;
        }        
    }

    Ok(())
}