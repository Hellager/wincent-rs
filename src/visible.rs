//! Quick Access visibility controls backed by Explorer registry settings.
//!
//! Windows stores the "Show recently used files" and "Show frequently used
//! folders" options under the current user's Explorer registry key.
//!
//! ## Behavior notes
//!
//! This module changes only the current user's visibility settings. It does
//! not invoke the Explorer Folder Options UI or deliberately clear Quick
//! Access history.
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
//! current user's Start document-tracking value. MDM settings, policy values,
//! Windows edition, and Explorer version can override or delay the visible
//! effect of the current-user value.
//!
//! On some Windows 11 builds, hiding Start Recommended by setting
//! `Start_TrackDocs = 0` can also make Explorer Recent Files invisible. In that
//! state, setting `ShowRecent = 1` alone may not make Recent Files visible
//! again; restore Start Recommended first, then use `ShowRecent` for Recent
//! Files visibility.

use crate::WincentResult;
use winreg::{enums::HKEY_CURRENT_USER, RegKey};

/// Visibility settings for Quick Access and Start Recommended.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum VisibilityTarget {
    /// Recently used files in Explorer.
    RecentFiles,
    /// Frequently used folders in Explorer.
    FrequentFolders,
    /// Recently used files in Windows Recommended items. Policies or Windows
    /// versions may override the visible effect of this setting.
    StartRecommended,
    /// Both Explorer Recent Files and Frequent Folders, excluding Start Recommended.
    All,
}

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

fn registry_value_for(target: VisibilityTarget) -> Option<&'static str> {
    match target {
        VisibilityTarget::RecentFiles => Some(SHOW_RECENT_VALUE),
        VisibilityTarget::FrequentFolders => Some(SHOW_FREQUENT_VALUE),
        VisibilityTarget::StartRecommended | VisibilityTarget::All => None,
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

/// Checks the configured visibility of an Explorer section or Start Recommended.
///
/// `VisibilityTarget::All` returns `true` only when both Explorer sections are
/// visible; it does not include Start Recommended.
///
/// Missing registry values are treated as visible.
/// This reads configured registry values, not the current rendered UI state.
///
/// # Errors
///
/// Returns registry I/O errors if the current user's visibility settings cannot
/// be read.
///
/// # Examples
///
/// ```rust,no_run
/// use wincent::prelude::*;
///
/// # fn main() -> WincentResult<()> {
/// let visible = is_visible(VisibilityTarget::RecentFiles)?;
/// println!("Recent Files visible: {visible}");
/// # Ok(())
/// # }
/// ```
pub fn is_visible(target: VisibilityTarget) -> WincentResult<bool> {
    if target == VisibilityTarget::StartRecommended {
        let Some(reg_key) = open_explorer_advanced_reg_key()? else {
            return Ok(true);
        };
        return read_start_recommended_visibility(&reg_key);
    }

    let Some(reg_key) = open_quick_access_reg_key()? else {
        return Ok(true);
    };

    if let Some(value_name) = registry_value_for(target) {
        return read_visibility_value(&reg_key, value_name);
    }

    Ok(read_visibility_value(&reg_key, SHOW_RECENT_VALUE)?
        && read_visibility_value(&reg_key, SHOW_FREQUENT_VALUE)?)
}

/// Sets the visibility of an Explorer section or Start Recommended.
///
/// `VisibilityTarget::All` applies the same value to Explorer Recent Files and
/// Frequent Folders, but not Start Recommended.
///
/// When hiding `VisibilityTarget::FrequentFolders` through this registry-backed API,
/// Explorer hides unpinned frequent folders, while pinned folders remain
/// visible. Passing `VisibilityTarget::All` with `visible` set to `false` has the
/// same Frequent Folders behavior.
///
/// This updates registry values for the current user without refreshing Explorer
/// or the Start menu. Success means the requested values were written, not
/// that the UI has already changed. `VisibilityTarget::All` writes two values
/// in sequence; if the second write fails, the first may already have changed.
/// On some Windows 11 builds, hiding Start Recommended can also hide Explorer
/// Recent Files; show it again before relying on the Recent Files setting.
///
/// # Errors
///
/// Returns registry I/O errors if the relevant settings key cannot be created
/// or updated.
pub fn set_visible(target: VisibilityTarget, visible: bool) -> WincentResult<()> {
    if target == VisibilityTarget::StartRecommended {
        let reg_key = create_explorer_advanced_reg_key()?;
        return write_start_recommended_visibility(&reg_key, visible);
    }

    let reg_key = create_quick_access_reg_key()?;

    if let Some(value_name) = registry_value_for(target) {
        return write_visibility_value(&reg_key, value_name, visible);
    }

    write_visibility_value(&reg_key, SHOW_RECENT_VALUE, visible)?;
    write_visibility_value(&reg_key, SHOW_FREQUENT_VALUE, visible)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn registry_value_names_match_quick_access_sections() {
        assert_eq!(
            registry_value_for(VisibilityTarget::RecentFiles),
            Some(SHOW_RECENT_VALUE)
        );
        assert_eq!(
            registry_value_for(VisibilityTarget::FrequentFolders),
            Some(SHOW_FREQUENT_VALUE)
        );
        assert_eq!(registry_value_for(VisibilityTarget::All), None);
        assert_eq!(registry_value_for(VisibilityTarget::StartRecommended), None);
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
        let initial = is_visible(VisibilityTarget::RecentFiles)?;

        set_visible(VisibilityTarget::RecentFiles, !initial)?;
        assert_eq!(is_visible(VisibilityTarget::RecentFiles)?, !initial);

        set_visible(VisibilityTarget::RecentFiles, initial)?;
        assert_eq!(is_visible(VisibilityTarget::RecentFiles)?, initial);

        Ok(())
    }

    #[test]
    #[ignore = "Reads and writes the current user's Explorer registry settings - run with: cargo test frequent_folders_visibility_round_trip -- --ignored --nocapture"]
    fn frequent_folders_visibility_round_trip() -> WincentResult<()> {
        let initial = is_visible(VisibilityTarget::FrequentFolders)?;

        set_visible(VisibilityTarget::FrequentFolders, !initial)?;
        assert_eq!(is_visible(VisibilityTarget::FrequentFolders)?, !initial);

        set_visible(VisibilityTarget::FrequentFolders, initial)?;
        assert_eq!(is_visible(VisibilityTarget::FrequentFolders)?, initial);

        Ok(())
    }

    #[test]
    #[ignore = "Reads and writes the current user's Explorer registry settings"]
    fn start_recommended_visibility_round_trip() -> WincentResult<()> {
        let initial = is_visible(VisibilityTarget::StartRecommended)?;
        set_visible(VisibilityTarget::StartRecommended, !initial)?;
        assert_eq!(is_visible(VisibilityTarget::StartRecommended)?, !initial);
        set_visible(VisibilityTarget::StartRecommended, initial)?;
        Ok(())
    }
}
