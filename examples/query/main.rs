extern crate wincent;

use wincent::{
    check_feasible, fix_feasible, get_frequent_folders, get_quick_access_items, get_recent_files, error::WincentError
};
use log::debug;
use std::io::{Error, ErrorKind};

fn main() -> Result<(), WincentError> {
    let _ = env_logger::builder()
        .target(env_logger::Target::Stdout)
        .filter_level(log::LevelFilter::Trace)
        .is_test(true)
        .try_init();

    let is_feasible = check_feasible()?;
    debug!("is feasible: {}", is_feasible);
    if !is_feasible{
        debug!("script not feasible, try fix");
        if !fix_feasible()? {
            debug!("fix acript feasible failed!!!");
            return Err(WincentError::Io(Error::from(ErrorKind::PermissionDenied)));            
        }
    }

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