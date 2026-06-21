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

use crate::{QuickAccess, WincentResult};
use winreg::{enums::HKEY_CURRENT_USER, RegKey};

const EXPLORER_KEY: &str = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer";
const EXPLORER_ADVANCED_KEY: &str =
    "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced";
const SHOW_FREQUENT_VALUE: &str = "ShowFrequent";
const SHOW_RECENT_VALUE: &str = "ShowRecent";
const START_TRACK_DOCS_VALUE: &str = "Start_TrackDocs";
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

fn open_start_recommended_reg_key() -> WincentResult<Option<RegKey>> {
    let hkcu = RegKey::predef(HKEY_CURRENT_USER);

    match hkcu.open_subkey(EXPLORER_ADVANCED_KEY) {
        Ok(key) => Ok(Some(key)),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(None),
        Err(error) => Err(error.into()),
    }
}

fn create_start_recommended_reg_key() -> WincentResult<RegKey> {
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
/// When hiding Frequent Folders through this registry-backed API, Explorer hides
/// unpinned frequent folders, while pinned folders remain visible.
///
/// # Errors
///
/// Returns registry I/O errors if Explorer visibility settings cannot be
/// created or updated.
pub fn set_frequent_folders_visible(visible: bool) -> WincentResult<()> {
    set_visible(QuickAccess::FrequentFolders, visible)
}

/// Checks whether the Windows 11 Start menu Recommended section is visible.
///
/// This reads the current-user value
/// `HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_TrackDocs`.
/// Missing values are treated as visible. MDM settings, policy values, Windows
/// edition, or Explorer version can override the effective UI state.
///
/// # Errors
///
/// Returns registry I/O errors if the current-user Explorer Advanced key cannot
/// be read.
pub fn is_start_recommended_section_visible() -> WincentResult<bool> {
    let Some(reg_key) = open_start_recommended_reg_key()? else {
        return Ok(true);
    };

    read_start_recommended_visibility(&reg_key)
}

/// Sets whether the Windows 11 Start menu Recommended section is visible.
///
/// This writes the current-user value
/// `HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_TrackDocs`.
/// Passing `true` writes `1`; passing `false` writes `0`. The value is not
/// removed so repeated calls are deterministic. This matches the Registry
/// Editor method that disables Start document tracking to clear the Recommended
/// section.
///
/// Explorer may require a refresh, restart, sign-out, or supported Windows 11
/// edition/build before the UI reflects the change.
///
/// # Errors
///
/// Returns registry I/O errors if the current-user Explorer Advanced key cannot
/// be created or updated.
pub fn set_start_recommended_section_visible(visible: bool) -> WincentResult<()> {
    let reg_key = create_start_recommended_reg_key()?;
    write_start_recommended_visibility(&reg_key, visible)
}

/// Shows the Windows 11 Start menu Recommended section.
///
/// # Errors
///
/// Returns registry I/O errors if the current-user Explorer Advanced key cannot
/// be created or updated.
pub fn show_start_recommended_section() -> WincentResult<()> {
    set_start_recommended_section_visible(true)
}

/// Hides the Windows 11 Start menu Recommended section.
///
/// # Errors
///
/// Returns registry I/O errors if the current-user Explorer Advanced key cannot
/// be created or updated.
pub fn hide_start_recommended_section() -> WincentResult<()> {
    set_start_recommended_section_visible(false)
}

/// Options for visibility write operations.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct VisibilityOptions {
    refresh_explorer: bool,
}

impl VisibilityOptions {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    #[must_use]
    pub fn refresh_explorer_enabled(&self) -> bool {
        self.refresh_explorer
    }

    #[must_use]
    pub fn refresh_explorer(self) -> Self {
        Self {
            refresh_explorer: true,
        }
    }

    #[must_use]
    pub fn with_refresh_explorer(self, enabled: bool) -> Self {
        Self {
            refresh_explorer: enabled,
        }
    }
}

/// Sets whether a Quick Access section is visible, with optional Explorer refresh.
///
/// When hiding `QuickAccess::FrequentFolders` through this registry-backed API,
/// Explorer hides unpinned frequent folders, while pinned folders remain
/// visible. Passing `QuickAccess::All` with `visible` set to `false` has the
/// same Frequent Folders behavior.
///
/// If `options.refresh_explorer_enabled()` is true, calls `refresh_explorer_window()`
/// after the registry write. Registry write is NOT rolled back if refresh fails.
pub fn set_visible_with_options(
    qa_type: QuickAccess,
    visible: bool,
    options: VisibilityOptions,
) -> WincentResult<()> {
    set_visible_with_options_inner(
        qa_type,
        visible,
        options,
        set_visible,
        crate::utils::refresh_explorer_window,
    )
}

fn set_visible_with_options_inner(
    qa_type: QuickAccess,
    visible: bool,
    options: VisibilityOptions,
    write: impl FnOnce(QuickAccess, bool) -> WincentResult<()>,
    refresh: impl FnOnce() -> WincentResult<()>,
) -> WincentResult<()> {
    write(qa_type, visible)?;
    if options.refresh_explorer_enabled() {
        refresh()?;
    }
    Ok(())
}

/// Sets whether Recent Files are visible, with optional Explorer refresh.
pub fn set_recent_files_visible_with_options(
    visible: bool,
    options: VisibilityOptions,
) -> WincentResult<()> {
    set_visible_with_options(QuickAccess::RecentFiles, visible, options)
}

/// Sets whether Frequent Folders are visible, with optional Explorer refresh.
///
/// When hiding Frequent Folders through this registry-backed API, Explorer hides
/// unpinned frequent folders, while pinned folders remain visible.
pub fn set_frequent_folders_visible_with_options(
    visible: bool,
    options: VisibilityOptions,
) -> WincentResult<()> {
    set_visible_with_options(QuickAccess::FrequentFolders, visible, options)
}

/// Sets whether the Windows 11 Start menu Recommended section is visible, with
/// optional Explorer refresh.
///
/// If `options.refresh_explorer_enabled()` is true, calls
/// `refresh_explorer_window()` after the registry write. Registry write is NOT
/// rolled back if refresh fails.
pub fn set_start_recommended_section_visible_with_options(
    visible: bool,
    options: VisibilityOptions,
) -> WincentResult<()> {
    set_start_recommended_section_visible_with_options_inner(
        visible,
        options,
        set_start_recommended_section_visible,
        crate::utils::refresh_explorer_window,
    )
}

fn set_start_recommended_section_visible_with_options_inner(
    visible: bool,
    options: VisibilityOptions,
    write: impl FnOnce(bool) -> WincentResult<()>,
    refresh: impl FnOnce() -> WincentResult<()>,
) -> WincentResult<()> {
    write(visible)?;
    if options.refresh_explorer_enabled() {
        refresh()?;
    }
    Ok(())
}

/// Shows the Windows 11 Start menu Recommended section, with optional Explorer refresh.
pub fn show_start_recommended_section_with_options(
    options: VisibilityOptions,
) -> WincentResult<()> {
    set_start_recommended_section_visible_with_options(true, options)
}

/// Hides the Windows 11 Start menu Recommended section, with optional Explorer refresh.
pub fn hide_start_recommended_section_with_options(
    options: VisibilityOptions,
) -> WincentResult<()> {
    set_start_recommended_section_visible_with_options(false, options)
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
    fn start_recommended_visibility_writes_start_track_docs_values() {
        assert_eq!(start_track_docs_value_for_visible(true), 1);
        assert_eq!(start_track_docs_value_for_visible(false), 0);
    }

    #[test]
    #[ignore = "Reads and writes the current user's Explorer registry settings - run with: cargo test recent_files_visibility_round_trip -- --ignored --nocapture"]
    fn recent_files_visibility_round_trip() -> WincentResult<()> {
        let initial = is_recent_files_visible()?;

        set_recent_files_visible(!initial)?;
        assert_eq!(is_recent_files_visible()?, !initial);

        set_recent_files_visible(initial)?;
        assert_eq!(is_recent_files_visible()?, initial);

        Ok(())
    }

    #[test]
    #[ignore = "Reads and writes the current user's Explorer registry settings - run with: cargo test frequent_folders_visibility_round_trip -- --ignored --nocapture"]
    fn frequent_folders_visibility_round_trip() -> WincentResult<()> {
        let initial = is_frequent_folders_visible()?;

        set_frequent_folders_visible(!initial)?;
        assert_eq!(is_frequent_folders_visible()?, !initial);

        set_frequent_folders_visible(initial)?;
        assert_eq!(is_frequent_folders_visible()?, initial);

        Ok(())
    }

    #[test]
    fn visibility_options_default_does_not_refresh() {
        assert!(!VisibilityOptions::new().refresh_explorer_enabled());
    }

    #[test]
    fn visibility_options_builder_enables_refresh() {
        assert!(VisibilityOptions::new()
            .refresh_explorer()
            .refresh_explorer_enabled());
    }

    #[test]
    fn visibility_options_with_refresh_explorer_true_enables() {
        assert!(VisibilityOptions::new()
            .with_refresh_explorer(true)
            .refresh_explorer_enabled());
    }

    #[test]
    fn visibility_options_with_refresh_explorer_false_disables() {
        assert!(!VisibilityOptions::new()
            .refresh_explorer()
            .with_refresh_explorer(false)
            .refresh_explorer_enabled());
    }

    #[test]
    fn set_visible_inner_no_refresh_does_not_call_refresh() {
        use std::cell::Cell;

        let write_called = Cell::new(false);
        let refresh_called = Cell::new(false);

        let result = set_visible_with_options_inner(
            QuickAccess::RecentFiles,
            true,
            VisibilityOptions::new(),
            |qa_type, visible| {
                assert_eq!(qa_type, QuickAccess::RecentFiles);
                assert!(visible);
                write_called.set(true);
                Ok(())
            },
            || {
                refresh_called.set(true);
                Ok(())
            },
        );

        assert!(result.is_ok());
        assert!(write_called.get(), "registry writer must be called");
        assert!(
            !refresh_called.get(),
            "refresh must not be called when option is disabled"
        );
    }

    #[test]
    fn set_visible_inner_with_refresh_calls_refresh() {
        use std::cell::Cell;

        let write_called = Cell::new(false);
        let refresh_called = Cell::new(false);

        let result = set_visible_with_options_inner(
            QuickAccess::RecentFiles,
            true,
            VisibilityOptions::new().refresh_explorer(),
            |qa_type, visible| {
                assert_eq!(qa_type, QuickAccess::RecentFiles);
                assert!(visible);
                write_called.set(true);
                Ok(())
            },
            || {
                refresh_called.set(true);
                Ok(())
            },
        );

        assert!(result.is_ok());
        assert!(write_called.get(), "registry writer must be called");
        assert!(
            refresh_called.get(),
            "refresh must be called when option is enabled"
        );
    }

    #[test]
    fn set_visible_inner_refresh_error_propagates() {
        use crate::error::WincentError;

        let expected = WincentError::SystemError("sentinel".into());
        let result = set_visible_with_options_inner(
            QuickAccess::RecentFiles,
            true,
            VisibilityOptions::new().refresh_explorer(),
            |_, _| Ok(()),
            || Err(expected),
        );

        match result {
            Err(WincentError::SystemError(message)) => assert_eq!(message, "sentinel"),
            other => panic!("expected refresh sentinel error, got {other:?}"),
        }
    }

    #[test]
    fn set_visible_inner_write_error_skips_refresh() {
        use crate::error::WincentError;
        use std::cell::Cell;

        let refresh_called = Cell::new(false);
        let expected = WincentError::SystemError("write sentinel".into());
        let result = set_visible_with_options_inner(
            QuickAccess::RecentFiles,
            true,
            VisibilityOptions::new().refresh_explorer(),
            |_, _| Err(expected),
            || {
                refresh_called.set(true);
                Ok(())
            },
        );

        match result {
            Err(WincentError::SystemError(message)) => assert_eq!(message, "write sentinel"),
            other => panic!("expected writer sentinel error, got {other:?}"),
        }
        assert!(
            !refresh_called.get(),
            "refresh must not be called when registry writer fails"
        );
    }

    #[test]
    fn set_start_recommended_inner_no_refresh_does_not_call_refresh() {
        use std::cell::Cell;

        let write_called = Cell::new(false);
        let refresh_called = Cell::new(false);

        let result = set_start_recommended_section_visible_with_options_inner(
            true,
            VisibilityOptions::new(),
            |visible| {
                assert!(visible);
                write_called.set(true);
                Ok(())
            },
            || {
                refresh_called.set(true);
                Ok(())
            },
        );

        assert!(result.is_ok());
        assert!(write_called.get(), "registry writer must be called");
        assert!(
            !refresh_called.get(),
            "refresh must not be called when option is disabled"
        );
    }

    #[test]
    fn set_start_recommended_inner_with_refresh_calls_refresh() {
        use std::cell::Cell;

        let write_called = Cell::new(false);
        let refresh_called = Cell::new(false);

        let result = set_start_recommended_section_visible_with_options_inner(
            false,
            VisibilityOptions::new().refresh_explorer(),
            |visible| {
                assert!(!visible);
                write_called.set(true);
                Ok(())
            },
            || {
                refresh_called.set(true);
                Ok(())
            },
        );

        assert!(result.is_ok());
        assert!(write_called.get(), "registry writer must be called");
        assert!(
            refresh_called.get(),
            "refresh must be called when option is enabled"
        );
    }

    #[test]
    fn set_start_recommended_inner_refresh_error_propagates() {
        use crate::error::WincentError;

        let expected = WincentError::SystemError("sentinel".into());
        let result = set_start_recommended_section_visible_with_options_inner(
            true,
            VisibilityOptions::new().refresh_explorer(),
            |_| Ok(()),
            || Err(expected),
        );

        match result {
            Err(WincentError::SystemError(message)) => assert_eq!(message, "sentinel"),
            other => panic!("expected refresh sentinel error, got {other:?}"),
        }
    }

    #[test]
    fn set_start_recommended_inner_write_error_skips_refresh() {
        use crate::error::WincentError;
        use std::cell::Cell;

        let refresh_called = Cell::new(false);
        let expected = WincentError::SystemError("write sentinel".into());
        let result = set_start_recommended_section_visible_with_options_inner(
            true,
            VisibilityOptions::new().refresh_explorer(),
            |_| Err(expected),
            || {
                refresh_called.set(true);
                Ok(())
            },
        );

        match result {
            Err(WincentError::SystemError(message)) => assert_eq!(message, "write sentinel"),
            other => panic!("expected writer sentinel error, got {other:?}"),
        }
        assert!(
            !refresh_called.get(),
            "refresh must not be called when registry writer fails"
        );
    }
}
