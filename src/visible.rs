use crate::WincentError;

/// Retrieves the registry key for Quick Access settings in Windows.
///
/// This function attempts to open the registry key located at
/// `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer`.
/// If successful, it returns the corresponding `RegKey`. If it fails,
/// it returns a `WincentError` indicating the type of error encountered.
///
/// # Errors
///
/// This function can return the following errors:
/// - `WincentError::IoError`: If the registry key cannot be found or if there is an
///   I/O error while attempting to open the key.
///
/// # Examples
///
/// ```
/// match get_quick_access_reg() {
///     Ok(key) => println!("Successfully retrieved Quick Access registry key."),
///     Err(e) => eprintln!("Failed to retrieve Quick Access registry key: {:?}", e),
/// }
/// ```
fn get_quick_access_reg() -> Result<winreg::RegKey, WincentError> {
    use winreg::enums::*;
    use winreg::RegKey;

    let hklm = RegKey::predef(HKEY_CURRENT_USER);
    hklm.open_subkey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer")
        .map_err(WincentError::IoError)
}

/// Checks and fixes the Quick Access registry settings.
///
/// This function checks the values of the "ShowFrequent" and "ShowRecent" registry keys
/// in the Quick Access registry key. If either of these values is missing, it sets them
/// to 1 (enabled).
/// 
/// # Errors
///
/// This function can return the following errors:
/// - `WincentError::IoError`: If the registry key cannot be found or if there is an
///   I/O error while attempting to open the key.
///
/// # Example
/// ```rust
/// use your_crate::WincentError;
///
/// fn main() -> Result<(), WincentError> {
/// match check_fix_quick_acess_reg() {
///     Ok(_) => println!("Successfully check/fix Quick Access registry key values."),
///     Err(e) => eprintln!("Failed to fix Quick Access registry key: {:?}", e),
/// }
/// ```
fn check_fix_quick_acess_reg() -> Result<(), WincentError> {
    let reg_key = get_quick_access_reg()?;

    let values_to_check = vec!["ShowFrequent", "ShowRecent"];
    for value_name in &values_to_check {
        match reg_key.get_value::<u32, _>(value_name) {
            Ok(_) => {
                // do nothing
            },
            Err(_) => {
                reg_key.set_value(value_name, &u32::from(1 as u8)).map_err(|e| WincentError::IoError(e))?;
            }
        }
    }

    Ok(())
}

/// Checks if the specified Quick Access target is visible in Windows.
///
/// This function retrieves the visibility setting for a given Quick Access target
/// (Frequent Folders, Recent Files, or All) from the Windows registry. It returns
/// `Ok(true)` if the target is visible, `Ok(false)` if it is not, or an error
/// if the operation fails.
///
/// # Parameters
///
/// - `target`: A `QuickAccess` enum value representing the target to check visibility for.
///
/// # Errors
///
/// This function can return the following errors:
/// - `WincentError::IoError`: If there is an I/O error while accessing the registry.
/// - `WincentError::ConvertError`: If there is an error converting the registry value.
///
/// # Examples
///
/// ```
/// match is_visible(QuickAccess::FrequentFolders) {
///     Ok(visible) => println!("Frequent Folders visibility: {}", visible),
///     Err(e) => eprintln!("Failed to check visibility: {:?}", e),
/// }
/// ```
pub(crate) fn is_visialbe_with_registry(target: crate::QuickAccess) -> Result<bool, WincentError> {
    let reg_key = get_quick_access_reg()?;
    let _ = check_fix_quick_acess_reg()?;
    let reg_value = match target {
        crate::QuickAccess::FrequentFolders => "ShowFrequent",
        crate::QuickAccess::RecentFiles => "ShowRecent",
        crate::QuickAccess::All => "ShowRecent",
    };

    let val = reg_key.get_raw_value(reg_value).map_err(WincentError::IoError)?;
    // raw reg value vec will contains '\n' between characters, needs to filter
    let filtered_vec: Vec<u8> = val.bytes.into_iter().filter(|&x| x != 0).collect();

    let visibility = u32::from_ne_bytes(filtered_vec[0..4].try_into().map_err(|e| WincentError::ConvertError(e))?);
    
    Ok(visibility != 0)
}

/// Sets the visibility of the specified Quick Access target in Windows.
///
/// This function updates the visibility setting for a given Quick Access target
/// (Frequent Folders, Recent Files, or All) in the Windows registry. It takes a
/// boolean value indicating whether the target should be visible (`true`) or not
/// (`false`). If the operation is successful, it returns `Ok(())`. If it fails,
/// it returns a `WincentError` indicating the type of error encountered.
///
/// # Parameters
///
/// - `target`: A `QuickAccess` enum value representing the target to set visibility for.
/// - `visible`: A boolean indicating the desired visibility state.
///
/// # Errors
///
/// This function can return the following errors:
/// - `WincentError::IoError`: If there is an I/O error while accessing the registry.
///
/// # Examples
///
/// ```
/// match set_visible(QuickAccess::FrequentFolders, true) {
///     Ok(_) => println!("Successfully set visibility for Frequent Folders."),
///     Err(e) => eprintln!("Failed to set visibility: {:?}", e),
/// }
/// ```
pub(crate) fn set_visiable_with_registry(target: crate::QuickAccess, visiable: bool) -> Result<(), WincentError> {
    let reg_key = get_quick_access_reg()?;
    let _ = check_fix_quick_acess_reg()?;
    let reg_value = match target {
        crate::QuickAccess::FrequentFolders => "ShowFrequent",
        crate::QuickAccess::RecentFiles => "ShowRecent",
        crate::QuickAccess::All => "ShowRecent",
    };

    reg_key.set_value(reg_value, &u32::from(visiable)).map_err(WincentError::IoError)?;

    Ok(())
}

#[cfg(test)]
mod visible_test {
    use crate::utils::init_test_logger;
    use super::*;
    
    #[test]
    fn test_check_visiable() {
        init_test_logger();

        let iff = is_visialbe_with_registry(crate::QuickAccess::RecentFiles).unwrap();
        let irf = is_visialbe_with_registry(crate::QuickAccess::FrequentFolders).unwrap();

        assert_eq!(iff, true);
        assert_eq!(irf, true);
    }

    #[ignore]
    #[test]
    fn test_restore_visiable() {
        // !!! Change the visiablity about quick access may lead to unexpected result about layout !!!
        init_test_logger();

        let ori_iff = is_visialbe_with_registry(crate::QuickAccess::RecentFiles).unwrap();
        let ori_irf = is_visialbe_with_registry(crate::QuickAccess::FrequentFolders).unwrap();

        set_visiable_with_registry(crate::QuickAccess::RecentFiles, !ori_iff).unwrap();
        set_visiable_with_registry(crate::QuickAccess::FrequentFolders, !ori_irf).unwrap();
    
        let new_iff = is_visialbe_with_registry(crate::QuickAccess::RecentFiles).unwrap();
        let new_irf = is_visialbe_with_registry(crate::QuickAccess::FrequentFolders).unwrap();

        assert_eq!(ori_iff, !new_iff);
        assert_eq!(ori_irf, !new_irf);

        set_visiable_with_registry(crate::QuickAccess::RecentFiles, ori_iff).unwrap();
        set_visiable_with_registry(crate::QuickAccess::FrequentFolders, ori_irf).unwrap();
    }
}
