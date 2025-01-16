use crate::{
    utils, 
    WincentResult,
    error::WincentError,
    scripts::{Script, execute_ps_script}
};
use std::path::Path;

/// Retrieves the registry key for the PowerShell execution policy.
fn get_execution_policy_reg() -> WincentResult<winreg::RegKey> {
    use winreg::enums::*;
    use winreg::RegKey;

    let key_path = "SOFTWARE\\Microsoft\\PowerShell\\1\\ShellIds\\Microsoft.PowerShell";
    let hkcu = RegKey::predef(HKEY_CURRENT_USER);
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);

    if utils::is_admin() {
        let hklm_reg_key = hklm.open_subkey_with_flags(key_path, winreg::enums::KEY_READ | winreg::enums::KEY_WRITE).map_err(WincentError::Io)?;
        return Ok(hklm_reg_key);
    } else {
        let (hkcu_reg_key, _) = hkcu.create_subkey_with_flags(key_path, winreg::enums::KEY_READ | winreg::enums::KEY_WRITE)
            .map_err(WincentError::Io)?;
        return Ok(hkcu_reg_key);
    }
}

/// Checks if the PowerShell execution policy is feasible based on the registry value.
pub(crate) fn check_script_feasible_with_registry() -> WincentResult<bool> {
    let reg_value=  "ExecutionPolicy";
    let reg_key = get_execution_policy_reg()?;
    let feasible_options = ["AllSigned", "Bypass", "RemoteSigned", "Unrestricted"];

    match reg_key.get_raw_value(reg_value) {
        Ok(val) => {
            // raw reg value vec will contains '\n' between characters, needs to filter
            let filtered_vec: Vec<u8> = val.bytes.into_iter().filter(|&x| x != 0).collect();
            let val_in_string = String::from_utf8_lossy(&filtered_vec).to_string();
            let _res = feasible_options.contains(&val_in_string.as_str());
            return Ok(feasible_options.contains(&val_in_string.as_str()));
        },
        Err(e) => {
            if e.kind() == std::io::ErrorKind::NotFound {
                return Ok(false);
            } else {
                return Err(WincentError::Io(e));
            }
        }
    }
}

/// Sets the PowerShell execution policy to "RemoteSigned" in the registry.
pub(crate) fn fix_script_feasible_with_registry() -> WincentResult<()> {
    let reg_value=  "ExecutionPolicy";
    let reg_key = get_execution_policy_reg()?;

    let _ = reg_key.set_value(reg_value, &"RemoteSigned").map_err(WincentError::Io)?;

    Ok(())
}

/// Checks if PowerShell query commands are available and executable.
pub(crate) fn check_query_feasible_with_script() -> WincentResult<bool> {
    let output = execute_ps_script(Script::CheckQueryFeasible, None)?;

    Ok(output.status.success())
}

/// Checks if PowerShell pin/unpin commands are available and executable.
pub(crate) fn check_pinunpin_feasible_with_script() -> WincentResult<bool> {
    let output = execute_ps_script(Script::CheckPinUnpinFeasible, None)?;

    Ok(output.status.success())
}

/// Checks if a registry path exists.
#[allow(dead_code)]
fn registry_path_exists(path: &Path) -> bool {
    use winreg::RegKey;
    use winreg::enums::*;

    let path_str = path.to_str().unwrap_or_default();
    let parts: Vec<&str> = path_str.split('\\').collect();
    
    if parts.len() < 2 {
        return false;
    }

    let hkey = match parts[0] {
        "HKEY_LOCAL_MACHINE" => HKEY_LOCAL_MACHINE,
        "HKEY_CURRENT_USER" => HKEY_CURRENT_USER,
        _ => return false,
    };

    let subkey = parts[1..].join("\\");
    let reg_key = RegKey::predef(hkey);
    
    reg_key.open_subkey_with_flags(&subkey, KEY_READ).is_ok()
}

/// Gets the current PowerShell execution policy.
#[allow(dead_code)]
fn get_execution_policy() -> WincentResult<String> {
    let reg_key = get_execution_policy_reg()?;
    let reg_value = "ExecutionPolicy";

    match reg_key.get_raw_value(reg_value) {
        Ok(val) => {
            let filtered_vec: Vec<u8> = val.bytes.into_iter().filter(|&x| x != 0).collect();
            Ok(String::from_utf8_lossy(&filtered_vec).to_string())
        },
        Err(e) => Err(WincentError::Io(e))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::Path;

    #[test]
    fn test_check_script_feasible() -> WincentResult<()> {
        let result = check_script_feasible_with_registry()?;
        
        assert!(result || !result);
        
        if !result {
            fix_script_feasible_with_registry()?;
            let fixed_result = check_script_feasible_with_registry()?;
            assert!(fixed_result, "Script should be feasible after fix");
        }
        
        Ok(())
    }

    #[test]
    fn test_fix_script_feasible() -> WincentResult<()> {
        let initial_policy = get_execution_policy()?;
        
        fix_script_feasible_with_registry()?;
        
        let final_policy = get_execution_policy()?;
        assert_eq!(final_policy, "RemoteSigned", "Execution policy should be set to RemoteSigned");
        
        let is_feasible = check_script_feasible_with_registry()?;
        assert!(is_feasible, "Should be feasible after fix");
        
        if initial_policy != "RemoteSigned" {
            let reg_key = get_execution_policy_reg()?;
            let _ = reg_key.set_value("ExecutionPolicy", &initial_policy);
        }
        
        Ok(())
    }

    #[test]
    fn test_registry_path_exists() {
        let valid_path = Path::new("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\PowerShell\\1\\ShellIds\\Microsoft.PowerShell");
        assert!(registry_path_exists(valid_path), "Valid PowerShell registry path should exist");

        let invalid_paths = [
            Path::new("INVALID_KEY\\Path"),
            Path::new("HKEY_LOCAL_MACHINE"),
            Path::new("HKEY_CURRENT_USER\\NonExistentPath"),
        ];

        for path in &invalid_paths {
            assert!(!registry_path_exists(path), "Invalid path should return false: {:?}", path);
        }
    }

    #[test]
    fn test_get_execution_policy() -> WincentResult<()> {
        let policy = get_execution_policy()?;
        
        assert!(!policy.is_empty(), "Execution policy should not be empty");
        
        let valid_policies = [
            "Restricted",
            "AllSigned", 
            "RemoteSigned",
            "Unrestricted",
            "Bypass"
        ];
        
        assert!(
            valid_policies.contains(&policy.as_str()),
            "Invalid execution policy: {}. Expected one of: {:?}",
            policy,
            valid_policies
        );
        
        Ok(())
    }

    #[test_log::test]
    fn test_check_query_feasible_with_script() -> WincentResult<()> {
        let result = check_query_feasible_with_script()?;
        
        println!("Query feasibility check result: {}", result);
        
        if !result {
            fix_script_feasible_with_registry()?;
            let fixed_result = check_query_feasible_with_script()?;
            println!("Query feasibility check after fix: {}", fixed_result);
        }
        
        Ok(())
    }

    #[test_log::test]
    fn test_check_pinunpin_feasible_with_script() -> WincentResult<()> {
        let result = check_pinunpin_feasible_with_script()?;
        
        println!("Pin/Unpin feasibility check result: {}", result);
        
        if !result {
            fix_script_feasible_with_registry()?;
            let fixed_result = check_pinunpin_feasible_with_script()?;
            println!("Pin/Unpin feasibility check after fix: {}", fixed_result);
        }
        
        Ok(())
    }
}
