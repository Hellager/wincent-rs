use crate::error::{WincentResult, WincentError};

/// Retrieves the registry key for Quick Access settings.
///
/// This function opens the registry key associated with Quick Access settings in Windows Explorer.
/// It specifically accesses the `HKEY_CURRENT_USER` hive to retrieve the relevant subkey.
///
/// # Returns
///
/// Returns a `WincentResult<winreg::RegKey>`, which contains the requested registry key.
/// If the operation is successful, it returns `Ok(reg_key)`. If there is an error accessing the registry,
/// it returns `WincentError::Io`.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let quick_access_key = get_quick_access_reg()?;
///     // Use quick_access_key for further operations
///     Ok(())
/// }
/// ```
fn get_quick_access_reg() -> WincentResult<winreg::RegKey> {
    use winreg::enums::*;
    use winreg::RegKey;

    let hklm = RegKey::predef(HKEY_CURRENT_USER);
    hklm.open_subkey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer")
        .map_err(WincentError::Io)
}

/// Checks and fixes the Quick Access registry settings.
///
/// This function retrieves the Quick Access registry key and checks for the presence of specific values:
/// "ShowFrequent" and "ShowRecent". If any of these values are missing, it sets them to `1` (enabled).
///
/// # Returns
///
/// Returns a `WincentResult<()>`. If the operation is successful, it returns `Ok(())`.
/// If there is an error accessing or modifying the registry, it returns `WincentError::Io`.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     check_fix_quick_access_reg()?;
///     println!("Quick Access registry settings checked and fixed if necessary.");
///     Ok(())
/// }
/// ```
fn check_fix_quick_acess_reg() -> WincentResult<()> {
    let reg_key = get_quick_access_reg()?;

    let values_to_check = vec!["ShowFrequent", "ShowRecent"];
    for value_name in &values_to_check {
        match reg_key.get_value::<u32, _>(value_name) {
            Ok(_) => {
                // do nothing
            },
            Err(_) => {
                reg_key.set_value(value_name, &u32::from(1 as u8)).map_err(|e| WincentError::Io(e))?;
            }
        }
    }

    Ok(())
}

/// Checks the visibility of a Quick Access item based on registry settings.
///
/// This function retrieves the Quick Access registry key and checks if the specified target
/// (Frequent Folders, Recent Files, or All) is visible based on the corresponding registry value.
/// It ensures that the registry settings are fixed if necessary before checking the visibility.
///
/// # Parameters
///
/// - `target`: An enum value of type `QuickAccess` that specifies which item to check for visibility.
///   - `QuickAccess::FrequentFolders`: Checks visibility for Frequent Folders.
///   - `QuickAccess::RecentFiles`: Checks visibility for Recent Files.
///   - `QuickAccess::All`: Checks visibility for Recent Files (as a fallback).
///
/// # Returns
///
/// Returns a `WincentResult<bool>`, which indicates whether the specified Quick Access item is visible.
/// If the operation is successful, it returns `Ok(true)` if visible, or `Ok(false)` if not visible.
/// If there is an error accessing the registry, it returns `WincentError::Io`.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let is_visible = is_visible_with_registry(QuickAccess::FrequentFolders)?;
///     if is_visible {
///         println!("Frequent Folders are visible.");
///     } else {
///         println!("Frequent Folders are not visible.");
///     }
///     Ok(())
/// }
/// ```
pub(crate) fn is_visialbe_with_registry(target: crate::QuickAccess) -> WincentResult<bool> {
    let reg_key = get_quick_access_reg()?;
    let _ = check_fix_quick_acess_reg()?;
    let reg_value = match target {
        crate::QuickAccess::FrequentFolders => "ShowFrequent",
        crate::QuickAccess::RecentFiles => "ShowRecent",
        crate::QuickAccess::All => "ShowRecent",
    };

    let val = reg_key.get_raw_value(reg_value).map_err(WincentError::Io)?;
    // raw reg value vec will contains '\n' between characters, needs to filter
    let filtered_vec: Vec<u8> = val.bytes.into_iter().filter(|&x| x != 0).collect();

    let visibility = u32::from_ne_bytes(filtered_vec[0..4].try_into().map_err(|e| WincentError::ArrayConversion(e))?);
    
    Ok(visibility != 0)
}

/// Sets the visibility of a Quick Access item in the registry.
///
/// This function updates the registry settings to set the specified target (Frequent Folders,
/// Recent Files, or All) to be visible or not based on the provided `visible` parameter.
/// It ensures that the registry settings are fixed if necessary before making the update.
///
/// # Parameters
///
/// - `target`: An enum value of type `QuickAccess` that specifies which item to set the visibility for.
///   - `QuickAccess::FrequentFolders`: Sets visibility for Frequent Folders.
///   - `QuickAccess::RecentFiles`: Sets visibility for Recent Files.
///   - `QuickAccess::All`: Sets visibility for Recent Files (as a fallback).
/// - `visible`: A boolean indicating whether the specified Quick Access item should be visible (`true`) or not (`false`).
///
/// # Returns
///
/// Returns a `WincentResult<()>`. If the operation is successful, it returns `Ok(())`.
/// If there is an error accessing or modifying the registry, it returns `WincentError::Io`.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     set_visible_with_registry(QuickAccess::FrequentFolders, true)?;
///     println!("Frequent Folders visibility set successfully.");
///     Ok(())
/// }
/// ```
pub(crate) fn set_visiable_with_registry(target: crate::QuickAccess, visiable: bool) -> WincentResult<()> {
    let reg_key = get_quick_access_reg()?;
    let _ = check_fix_quick_acess_reg()?;
    let reg_value = match target {
        crate::QuickAccess::FrequentFolders => "ShowFrequent",
        crate::QuickAccess::RecentFiles => "ShowRecent",
        crate::QuickAccess::All => "ShowRecent",
    };

    reg_key.set_value(reg_value, &u32::from(visiable)).map_err(WincentError::Io)?;

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::error::WincentResult;
    use crate::QuickAccess;

    #[ignore]
    #[test_log::test]
    fn test_recent_files_visibility() -> WincentResult<()> {
        let initial_state = is_visialbe_with_registry(QuickAccess::RecentFiles)?;
        println!("Initial recent files visibility: {}", initial_state);

        set_visiable_with_registry(QuickAccess::RecentFiles, !initial_state)?;
        let changed_state = is_visialbe_with_registry(QuickAccess::RecentFiles)?;
        assert_eq!(changed_state, !initial_state, "Visibility should be changed");

        set_visiable_with_registry(QuickAccess::RecentFiles, initial_state)?;
        let final_state = is_visialbe_with_registry(QuickAccess::RecentFiles)?;
        assert_eq!(final_state, initial_state, "Should restore to initial state");

        Ok(())
    }

    #[ignore]
    #[test_log::test]
    fn test_frequent_folders_visibility() -> WincentResult<()> {
        let initial_state = is_visialbe_with_registry(QuickAccess::FrequentFolders)?;
        println!("Initial frequent folders visibility: {}", initial_state);

        set_visiable_with_registry(QuickAccess::FrequentFolders, !initial_state)?;
        let changed_state = is_visialbe_with_registry(QuickAccess::FrequentFolders)?;
        assert_eq!(changed_state, !initial_state, "Visibility should be changed");

        set_visiable_with_registry(QuickAccess::FrequentFolders, initial_state)?;
        let final_state = is_visialbe_with_registry(QuickAccess::FrequentFolders)?;
        assert_eq!(final_state, initial_state, "Should restore to initial state");

        Ok(())
    }
}
