extern crate wincent;

use wincent::{WincentError, get_recent_files, get_frequent_folders, get_quick_access_items};
use log::debug;

fn main() -> Result<(), WincentError> {
    let _ = env_logger::builder()
        .target(env_logger::Target::Stdout)
        .filter_level(log::LevelFilter::Trace)
        .is_test(true)
        .try_init();

    let recent_files: Vec<String> = get_recent_files()?;
    let frequent_folders: Vec<String> = get_frequent_folders()?;
    let quick_access: Vec<String> = get_quick_access_items()?;

    debug!("recent files");
    for (idx, item) in recent_files.iter().enumerate() {
        debug!("{}. {}", idx, item);
    }
    debug!("\n\n");

    debug!("frequent folders");
    for (idx, item) in frequent_folders.iter().enumerate() {
        debug!("{}. {}", idx, item);
    }
    debug!("\n\n");

    debug!("quick access items");
    for (idx, item) in quick_access.iter().enumerate() {
        debug!("{}. {}", idx, item);
    }

    Ok(())
}