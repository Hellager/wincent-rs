//! Quick Access visibility controls backed by Explorer registry settings.
//!
//! Windows stores the "Show recently used files" and "Show frequently used
//! folders" options under the current user's Explorer registry key.
//!
//! ## Behavior notes
//!
//! This module changes only the current user's `ShowRecent` and `ShowFrequent`
//! registry values. It does not invoke the Explorer Folder Options UI and does
//! not deliberately clear Quick Access history.
//!
//! Observed Windows behavior:
//!
//! - Setting Frequent Folders hidden through this module hides unpinned frequent
//!   folders, while pinned folders remain visible.
//! - Changing Frequent Folders visibility through Explorer's Folder Options UI
//!   can clear all unpinned frequent folders. Showing them again may restore
//!   Windows default system pinned folders.
//! - Setting Recent Files hidden through this module hides recently used files.
//! - Changing Recent Files visibility through Explorer's Folder Options UI can
//!   clear all recent file entries.

#![cfg(feature = "visible")]
use crate::{QuickAccess, WincentResult};
use winreg::{enums::HKEY_CURRENT_USER, RegKey};

const EXPLORER_KEY: &str = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer";
const SHOW_FREQUENT_VALUE: &str = "ShowFrequent";
const SHOW_RECENT_VALUE: &str = "ShowRecent";
const VISIBLE_DWORD: u32 = 1;
const HIDDEN_DWORD: u32 = 0;

fn open_quick_access_reg_key() -> WincentResult<Option<RegKey>> {
    let hkcu = RegKey::predef(HKEY_CURRENT_USER);

    match hkcu.open_subkey(EXPLORER_KEY) {
        Ok(key) => Ok(Some(key)),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(None),
        Err(error) => Err(error.into()),
    }
}

fn create_quick_access_reg_key() -> WincentResult<RegKey> {
    let hkcu = RegKey::predef(HKEY_CURRENT_USER);
    let (key, _) = hkcu.create_subkey(EXPLORER_KEY)?;
    Ok(key)
}

fn registry_value_for(qa_type: QuickAccess) -> Option<&'static str> {
    match qa_type {
        QuickAccess::RecentFiles => Some(SHOW_RECENT_VALUE),
        QuickAccess::FrequentFolders => Some(SHOW_FREQUENT_VALUE),
        QuickAccess::All => None,
    }
}

fn read_visibility_value(reg_key: &RegKey, value_name: &str) -> WincentResult<bool> {
    match reg_key.get_value::<u32, _>(value_name) {
        Ok(value) => Ok(value != HIDDEN_DWORD),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(true),
        Err(error) => Err(error.into()),
    }
}

fn write_visibility_value(reg_key: &RegKey, value_name: &str, visible: bool) -> WincentResult<()> {
    let value = if visible { VISIBLE_DWORD } else { HIDDEN_DWORD };

    reg_key.set_value(value_name, &value)?;
    Ok(())
}

/// Checks whether a Quick Access section is visible in Explorer.
///
/// `QuickAccess::All` returns `true` only when both Recent Files and Frequent
/// Folders are visible.
///
/// Missing registry values are treated as visible, matching Explorer's default
/// behavior for a new user profile.
///
/// # Errors
///
/// Returns registry I/O errors if the current user's Explorer settings cannot
/// be read.
///
/// # Examples
///
/// ```rust,no_run
/// use wincent::prelude::*;
///
/// # fn main() -> WincentResult<()> {
/// let visible = is_visible(QuickAccess::RecentFiles)?;
/// println!("Recent Files visible: {visible}");
/// # Ok(())
/// # }
/// ```
pub fn is_visible(qa_type: QuickAccess) -> WincentResult<bool> {
    let Some(reg_key) = open_quick_access_reg_key()? else {
        return Ok(true);
    };

    if let Some(value_name) = registry_value_for(qa_type) {
        return read_visibility_value(&reg_key, value_name);
    }

    Ok(read_visibility_value(&reg_key, SHOW_RECENT_VALUE)?
        && read_visibility_value(&reg_key, SHOW_FREQUENT_VALUE)?)
}

/// Sets whether a Quick Access section is visible in Explorer.
///
/// Passing `QuickAccess::All` applies the same value to both Recent Files and
/// Frequent Folders.
///
/// This updates registry values for the current user. It does not clear Quick
/// Access history or invoke Explorer's Folder Options UI.
///
/// # Errors
///
/// Returns registry I/O errors if the Explorer settings key cannot be created
/// or updated.
pub fn set_visible(qa_type: QuickAccess, visible: bool) -> WincentResult<()> {
    let reg_key = create_quick_access_reg_key()?;

    if let Some(value_name) = registry_value_for(qa_type) {
        return write_visibility_value(&reg_key, value_name, visible);
    }

    write_visibility_value(&reg_key, SHOW_RECENT_VALUE, visible)?;
    write_visibility_value(&reg_key, SHOW_FREQUENT_VALUE, visible)
}

/// Checks whether Recent Files are visible in Windows Quick Access.
///
/// # Errors
///
/// Returns registry I/O errors if Explorer visibility settings cannot be read.
pub fn is_recent_files_visible() -> WincentResult<bool> {
    is_visible(QuickAccess::RecentFiles)
}

/// Checks whether Frequent Folders are visible in Windows Quick Access.
///
/// # Errors
///
/// Returns registry I/O errors if Explorer visibility settings cannot be read.
pub fn is_frequent_folders_visible() -> WincentResult<bool> {
    is_visible(QuickAccess::FrequentFolders)
}

/// Sets whether Recent Files are visible in Windows Quick Access.
///
/// # Errors
///
/// Returns registry I/O errors if Explorer visibility settings cannot be
/// created or updated.
pub fn set_recent_files_visible(visible: bool) -> WincentResult<()> {
    set_visible(QuickAccess::RecentFiles, visible)
}

/// Sets whether Frequent Folders are visible in Windows Quick Access.
///
/// # Errors
///
/// Returns registry I/O errors if Explorer visibility settings cannot be
/// created or updated.
pub fn set_frequent_folders_visible(visible: bool) -> WincentResult<()> {
    set_visible(QuickAccess::FrequentFolders, visible)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn registry_value_names_match_quick_access_sections() {
        assert_eq!(
            registry_value_for(QuickAccess::RecentFiles),
            Some(SHOW_RECENT_VALUE)
        );
        assert_eq!(
            registry_value_for(QuickAccess::FrequentFolders),
            Some(SHOW_FREQUENT_VALUE)
        );
        assert_eq!(registry_value_for(QuickAccess::All), None);
    }

    #[test]
    #[ignore = "Reads and writes the current user's Explorer registry settings — run with: cargo test --features visible recent_files_visibility_round_trip -- --ignored --nocapture"]
    fn recent_files_visibility_round_trip() -> WincentResult<()> {
        let initial = is_recent_files_visible()?;

        set_recent_files_visible(!initial)?;
        assert_eq!(is_recent_files_visible()?, !initial);

        set_recent_files_visible(initial)?;
        assert_eq!(is_recent_files_visible()?, initial);

        Ok(())
    }

    #[test]
    #[ignore = "Reads and writes the current user's Explorer registry settings — run with: cargo test --features visible frequent_folders_visibility_round_trip -- --ignored --nocapture"]
    fn frequent_folders_visibility_round_trip() -> WincentResult<()> {
        let initial = is_frequent_folders_visible()?;

        set_frequent_folders_visible(!initial)?;
        assert_eq!(is_frequent_folders_visible()?, !initial);

        set_frequent_folders_visible(initial)?;
        assert_eq!(is_frequent_folders_visible()?, initial);

        Ok(())
    }
}
