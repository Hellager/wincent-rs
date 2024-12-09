extern crate wincent;

use wincent::{
    check_feasible, fix_feasible, get_frequent_folders, get_quick_access_items, get_recent_files, WincentError
};
use log::debug;
use std::io::{Error, ErrorKind};

#[tokio::main]
async fn main() -> Result<(), WincentError> {
    let _ = env_logger::builder()
        .target(env_logger::Target::Stdout)
        .filter_level(log::LevelFilter::Trace)
        .is_test(true)
        .try_init();

    let is_feasible = check_feasible()?;
    debug!("is feasible: {}", is_feasible);
    if !is_feasible{
        debug!("script not feasible, try fix");
        fix_feasible()?;

        if check_feasible()? {
            debug!("fix script feasible success!");
        } else {
            debug!("fix acript feasible failed!!!");
            return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
        }
    }

    let recent_files: Vec<String> = get_recent_files().await?;
    let frequent_folders: Vec<String> = get_frequent_folders().await?;
    let quick_access: Vec<String> = get_quick_access_items().await?;

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