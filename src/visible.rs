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
//!
//! The Start menu Recommended section on Windows 11 can be controlled with the
//! current user's Start document-tracking value. This module exposes it as a
//! distinct API instead of mixing it with Quick Access categories. MDM settings,
//! policy values, Windows edition, and Explorer version can override or delay
//! the visible effect of the current-user value.
//!
//! On some Windows 11 builds, hiding Start Recommended by setting
//! `Start_TrackDocs = 0` can also make Explorer Recent Files invisible. In that
//! state, setting `ShowRecent = 1` alone may not make Recent Files visible
//! again; restore Start Recommended first, then use `ShowRecent` for Recent
//! Files visibility.

use crate::{QuickAccess, WincentResult};
use winreg::{enums::HKEY_CURRENT_USER, RegKey};

const EXPLORER_KEY: &str = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer";
const EXPLORER_ADVANCED_KEY: &str =
    "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced";
const SHOW_FREQUENT_VALUE: &str = "ShowFrequent";
const SHOW_RECENT_VALUE: &str = "ShowRecent";
const START_TRACK_DOCS_VALUE: &str = "Start_TrackDocs";
const HIDDEN_ITEMS_VALUE: &str = "Hidden";
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

fn open_explorer_advanced_reg_key() -> WincentResult<Option<RegKey>> {
    let hkcu = RegKey::predef(HKEY_CURRENT_USER);

    match hkcu.open_subkey(EXPLORER_ADVANCED_KEY) {
        Ok(key) => Ok(Some(key)),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(None),
        Err(error) => Err(error.into()),
    }
}

fn create_explorer_advanced_reg_key() -> WincentResult<RegKey> {
    let hkcu = RegKey::predef(HKEY_CURRENT_USER);
    let (key, _) = hkcu.create_subkey(EXPLORER_ADVANCED_KEY)?;
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

fn read_start_recommended_visibility_value(value: Option<u32>) -> bool {
    value != Some(HIDDEN_DWORD)
}

fn read_start_recommended_visibility(reg_key: &RegKey) -> WincentResult<bool> {
    match reg_key.get_value::<u32, _>(START_TRACK_DOCS_VALUE) {
        Ok(value) => Ok(read_start_recommended_visibility_value(Some(value))),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(true),
        Err(error) => Err(error.into()),
    }
}

fn read_hidden_items_visibility_value(value: Option<u32>) -> bool {
    value == Some(VISIBLE_DWORD)
}

pub(crate) fn explorer_shows_hidden_items() -> WincentResult<bool> {
    let Some(reg_key) = open_explorer_advanced_reg_key()? else {
        return Ok(false);
    };

    match reg_key.get_value::<u32, _>(HIDDEN_ITEMS_VALUE) {
        Ok(value) => Ok(read_hidden_items_visibility_value(Some(value))),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(false),
        Err(error) => Err(error.into()),
    }
}

fn start_track_docs_value_for_visible(visible: bool) -> u32 {
    if visible {
        VISIBLE_DWORD
    } else {
        HIDDEN_DWORD
    }
}

fn write_start_recommended_visibility(reg_key: &RegKey, visible: bool) -> WincentResult<()> {
    let value = start_track_docs_value_for_visible(visible);

    reg_key.set_value(START_TRACK_DOCS_VALUE, &value)?;
    Ok(())
}

/// Checks whether a Quick Access section is visible in Explorer.
///
/// `QuickAccess::All` returns `true` only when both Recent Files and Frequent
/// Folders are visible.
///
/// Missing registry values are treated as visible, matching Explorer's default
/// behavior for a new user profile.
/// This reads configured registry values, not the current rendered UI state.
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
/// When hiding `QuickAccess::FrequentFolders` through this registry-backed API,
/// Explorer hides unpinned frequent folders, while pinned folders remain
/// visible. Passing `QuickAccess::All` with `visible` set to `false` has the
/// same Frequent Folders behavior.
///
/// This updates registry values for the current user. It does not clear Quick
/// Access history, invoke Explorer's Folder Options UI, or refresh Explorer.
/// Success means the requested values were written, not that the UI has
/// already changed. `QuickAccess::All` writes two values in sequence; if the
/// second write fails, the first may already have changed.
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

/// Checks the setting for recently used files in Windows Recommended items.
///
/// This reads the current-user value
/// `HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_TrackDocs`.
/// This value controls whether recently used files appear in Windows
/// Recommended items. Missing values are treated as visible. MDM settings,
/// policy values, Windows edition, or Explorer version can override the
/// effective UI state; this does not inspect the rendered Start menu.
///
/// # Errors
///
/// Returns registry I/O errors if the current-user Explorer Advanced key cannot
/// be read.
pub fn is_start_recommended_section_visible() -> WincentResult<bool> {
    let Some(reg_key) = open_explorer_advanced_reg_key()? else {
        return Ok(true);
    };

    read_start_recommended_visibility(&reg_key)
}

/// Sets whether recently used files appear in Windows Recommended items.
///
/// This writes the current-user value
/// `HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_TrackDocs`.
/// Passing `true` writes `1`; passing `false` writes `0`. The value is not
/// removed so repeated calls are deterministic. The value controls whether
/// recently used files appear in Windows Recommended items.
///
/// On some Windows 11 builds, setting this to `false` can also make Explorer
/// Recent Files invisible. In that state, setting `ShowRecent = 1` alone may
/// not make Recent Files visible again; set Start Recommended visible first,
/// then control Recent Files with `ShowRecent`.
///
/// Explorer may require a refresh, restart, sign-out, or supported Windows 11
/// edition/build before the UI reflects the change. This operation does not
/// refresh Explorer or the Start menu. Success means the registry write
/// completed, not that the UI has already changed.
///
/// # Errors
///
/// Returns registry I/O errors if the current-user Explorer Advanced key cannot
/// be created or updated.
pub fn set_start_recommended_section_visible(visible: bool) -> WincentResult<()> {
    let reg_key = create_explorer_advanced_reg_key()?;
    write_start_recommended_visibility(&reg_key, visible)
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
    fn start_recommended_registry_names_match_explorer_advanced_setting() {
        assert_eq!(
            EXPLORER_ADVANCED_KEY,
            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced"
        );
        assert_eq!(START_TRACK_DOCS_VALUE, "Start_TrackDocs");
    }

    #[test]
    fn missing_start_track_docs_value_is_visible() {
        assert!(read_start_recommended_visibility_value(None));
    }

    #[test]
    fn enabled_start_track_docs_value_is_visible() {
        assert!(read_start_recommended_visibility_value(Some(1)));
    }

    #[test]
    fn disabled_start_track_docs_value_is_hidden() {
        assert!(!read_start_recommended_visibility_value(Some(0)));
    }

    #[test]
    fn explorer_hidden_items_value_matches_folder_options() {
        assert!(read_hidden_items_visibility_value(Some(1)));
        assert!(!read_hidden_items_visibility_value(Some(2)));
        assert!(!read_hidden_items_visibility_value(Some(0)));
        assert!(!read_hidden_items_visibility_value(None));
    }

    #[test]
    fn start_recommended_visibility_writes_start_track_docs_values() {
        assert_eq!(start_track_docs_value_for_visible(true), 1);
        assert_eq!(start_track_docs_value_for_visible(false), 0);
    }

    #[test]
    #[ignore = "Reads and writes the current user's Explorer registry settings - run with: cargo test recent_files_visibility_round_trip -- --ignored --nocapture"]
    fn recent_files_visibility_round_trip() -> WincentResult<()> {
        let initial = is_visible(QuickAccess::RecentFiles)?;

        set_visible(QuickAccess::RecentFiles, !initial)?;
        assert_eq!(is_visible(QuickAccess::RecentFiles)?, !initial);

        set_visible(QuickAccess::RecentFiles, initial)?;
        assert_eq!(is_visible(QuickAccess::RecentFiles)?, initial);

        Ok(())
    }

    #[test]
    #[ignore = "Reads and writes the current user's Explorer registry settings - run with: cargo test frequent_folders_visibility_round_trip -- --ignored --nocapture"]
    fn frequent_folders_visibility_round_trip() -> WincentResult<()> {
        let initial = is_visible(QuickAccess::FrequentFolders)?;

        set_visible(QuickAccess::FrequentFolders, !initial)?;
        assert_eq!(is_visible(QuickAccess::FrequentFolders)?, !initial);

        set_visible(QuickAccess::FrequentFolders, initial)?;
        assert_eq!(is_visible(QuickAccess::FrequentFolders)?, initial);

        Ok(())
    }
}
