//! Check and set Windows Quick Access visibility.
//!
//! ## Example
//!
//! ```no_run
//! use wincent::{
//!     visible::{
//!         is_frequent_folders_visible, is_recent_files_visiable, set_frequent_folders_visiable,
//!         set_recent_files_visiable,
//!     },
//!     WincentResult,
//! };
//!
//! fn print_visibility_status() -> WincentResult<()> {
//!     println!(
//!         "Recent Files: {}",
//!         if is_recent_files_visiable()? {
//!             "Visible"
//!         } else {
//!             "Hidden"
//!         }
//!     );
//!     println!(
//!         "Frequent Folders: {}",
//!         if is_frequent_folders_visible()? {
//!             "Visible"
//!         } else {
//!             "Hidden"
//!         }
//!     );
//!     Ok(())
//! }
//!
//! fn main() -> WincentResult<()> {
//!     // Show initial status
//!     println!("Initial visibility status:");
//!     print_visibility_status()?;
//!
//!     // Save initial state
//!     let initial_recent = is_recent_files_visiable()?;
//!     let initial_folders = is_frequent_folders_visible()?;
//!
//!     // Hide all sections
//!     println!("Hiding all sections...");
//!     set_recent_files_visiable(false)?;
//!     set_frequent_folders_visiable(false)?;
//!
//!     println!("Status after hiding:");
//!     print_visibility_status()?;
//!
//!     // Show all sections
//!     println!("Showing all sections...");
//!     set_recent_files_visiable(true)?;
//!     set_frequent_folders_visiable(true)?;
//!
//!     println!("Status after showing:");
//!     print_visibility_status()?;
//!
//!     // Set different visibility
//!     println!("Setting different visibility...");
//!     set_recent_files_visiable(false)?;
//!     set_frequent_folders_visiable(true)?;
//!
//!     println!("Status after mixed settings:");
//!     print_visibility_status()?;
//!
//!     // Restore initial state
//!     println!("Restoring initial visibility...");
//!     set_recent_files_visiable(initial_recent)?;
//!     set_frequent_folders_visiable(initial_folders)?;
//!
//!     println!("Final status (restored to initial):");
//!     print_visibility_status()?;
//!
//!     Ok(())
//! }
//! ```

use crate::{error::WincentError, QuickAccess, WincentResult};

/// Retrieves the registry key for Quick Access settings.
fn get_quick_access_reg() -> WincentResult<winreg::RegKey> {
    use winreg::enums::*;
    use winreg::RegKey;

    let hklm = RegKey::predef(HKEY_CURRENT_USER);
    hklm.create_subkey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer")
        .map(|(key, _)| key)
        .map_err(WincentError::Io)
}

/// Checks and fixes the Quick Access registry settings.
fn check_fix_quick_acess_reg() -> WincentResult<()> {
    let reg_key = get_quick_access_reg()?;

    let values_to_check = vec!["ShowFrequent", "ShowRecent"];
    for value_name in &values_to_check {
        match reg_key.get_value::<u32, _>(value_name) {
            Ok(_) => {
                // do nothing
            }
            Err(_) => {
                reg_key
                    .set_value(value_name, &u32::from(1_u8))
                    .map_err(WincentError::Io)?;
            }
        }
    }

    Ok(())
}

/// Checks the visibility of a Quick Access item based on registry settings.
pub(crate) fn is_visialbe_with_registry(target: crate::QuickAccess) -> WincentResult<bool> {
    let reg_key = get_quick_access_reg()?;
    check_fix_quick_acess_reg()?;
    let reg_value = match target {
        crate::QuickAccess::FrequentFolders => "ShowFrequent",
        crate::QuickAccess::RecentFiles => "ShowRecent",
        crate::QuickAccess::All => "ShowRecent",
    };

    let visibility: u32 = reg_key.get_value(reg_value).map_err(WincentError::Io)?;
    Ok(visibility != 0)
}

/// Sets the visibility of a Quick Access item in the registry.
pub(crate) fn set_visiable_with_registry(
    target: crate::QuickAccess,
    visiable: bool,
) -> WincentResult<()> {
    let reg_key = get_quick_access_reg()?;
    check_fix_quick_acess_reg()?;
    let reg_value = match target {
        crate::QuickAccess::FrequentFolders => "ShowFrequent",
        crate::QuickAccess::RecentFiles => "ShowRecent",
        crate::QuickAccess::All => "ShowRecent",
    };

    reg_key
        .set_value(reg_value, &u32::from(visiable))
        .map_err(WincentError::Io)?;

    Ok(())
}

/****************************************************** Quick Access Visiablity ******************************************************/

/// Checks if Quick Access visibility settings can be modified.
///
/// # Returns
///
/// Returns `true` if Quick Access visibility can be controlled.
///
/// # Example
///
/// ```no_run
/// use wincent::{
///     visible::{is_recent_files_visiable, set_recent_files_visiable},
///     error::WincentError,
/// };
///
/// fn main() -> Result<(), WincentError> {
///     let is_visible = is_recent_files_visiable()?;
///     if !is_visible {
///         set_recent_files_visiable(true)?;
///     }
///     Ok(())
/// }
/// ```
pub fn is_recent_files_visiable() -> WincentResult<bool> {
    is_visialbe_with_registry(QuickAccess::RecentFiles)
}

/// Checks if frequent folders are visible in Windows Quick Access.
///
/// # Returns
///
/// Returns `true` if frequent folders are visible, `false` if they are hidden.
///
/// # Example
///
/// ```no_run
/// use wincent::{
///     visible::{is_frequent_folders_visible, set_frequent_folders_visiable},
///     error::WincentError,
/// };
///
/// fn main() -> Result<(), WincentError> {
///     let is_visible = is_frequent_folders_visible()?;
///     println!("Frequent folders are {}", if is_visible { "visible" } else { "hidden" });
///     
///     // Ensure frequent folders are visible
///     if !is_visible {
///         set_frequent_folders_visiable(true)?;
///     }
///     Ok(())
/// }
/// ```
pub fn is_frequent_folders_visible() -> WincentResult<bool> {
    is_visialbe_with_registry(QuickAccess::FrequentFolders)
}

/// Sets the visibility of Quick Access recent files.
///
/// # Arguments
///
/// * `is_visiable` - Whether recent files should be visible
///
/// # Example
///
/// ```no_run
/// use wincent::{visible::set_recent_files_visiable, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Hide recent files in Quick Access
///     set_recent_files_visiable(false)?;
///     Ok(())
/// }
/// ```
pub fn set_recent_files_visiable(is_visiable: bool) -> WincentResult<()> {
    set_visiable_with_registry(QuickAccess::RecentFiles, is_visiable)
}

/// Sets the visibility of frequent folders in Windows Quick Access.
///
/// # Arguments
///
/// * `is_visiable` - `true` to show frequent folders, `false` to hide them
///
/// # Returns
///
/// Returns `Ok(())` if the visibility was successfully changed.
///
/// # Example
///
/// ```no_run
/// use wincent::{visible::set_frequent_folders_visiable, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Hide frequent folders in Quick Access
///     set_frequent_folders_visiable(false)?;
///     println!("Frequent folders are now hidden");
///     Ok(())
/// }
/// ```
pub fn set_frequent_folders_visiable(is_visiable: bool) -> WincentResult<()> {
    set_visiable_with_registry(QuickAccess::FrequentFolders, is_visiable)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::QuickAccess;

    #[test]
    #[ignore]
    fn test_recent_files_visibility() -> WincentResult<()> {
        let initial_state = is_visialbe_with_registry(QuickAccess::RecentFiles)?;

        set_visiable_with_registry(QuickAccess::RecentFiles, !initial_state)?;
        let changed_state = is_visialbe_with_registry(QuickAccess::RecentFiles)?;
        assert_eq!(
            changed_state, !initial_state,
            "Visibility should be changed"
        );

        set_visiable_with_registry(QuickAccess::RecentFiles, initial_state)?;
        let final_state = is_visialbe_with_registry(QuickAccess::RecentFiles)?;
        assert_eq!(
            final_state, initial_state,
            "Should restore to initial state"
        );

        Ok(())
    }

    #[test]
    #[ignore]
    fn test_frequent_folders_visibility() -> WincentResult<()> {
        let initial_state = is_visialbe_with_registry(QuickAccess::FrequentFolders)?;

        set_visiable_with_registry(QuickAccess::FrequentFolders, !initial_state)?;
        let changed_state = is_visialbe_with_registry(QuickAccess::FrequentFolders)?;
        assert_eq!(
            changed_state, !initial_state,
            "Visibility should be changed"
        );

        set_visiable_with_registry(QuickAccess::FrequentFolders, initial_state)?;
        let final_state = is_visialbe_with_registry(QuickAccess::FrequentFolders)?;
        assert_eq!(
            final_state, initial_state,
            "Should restore to initial state"
        );

        Ok(())
    }
}
