use crate::{WincentResult, error::WincentError};

/// Retrieves the registry key for Quick Access settings.
fn get_quick_access_reg() -> WincentResult<winreg::RegKey> {
    use winreg::enums::*;
    use winreg::RegKey;

    let hklm = RegKey::predef(HKEY_CURRENT_USER);
    hklm.open_subkey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer")
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
            },
            Err(_) => {
                reg_key.set_value(value_name, &u32::from(1 as u8)).map_err(|e| WincentError::Io(e))?;
            }
        }
    }

    Ok(())
}

/// Checks the visibility of a Quick Access item based on registry settings.
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
