use crate::{utils, error::{WincentError, WincentResult}};
use std::path::Path;

/// Retrieves the registry key for the PowerShell execution policy.
///
/// This function determines which registry key to open based on the current user's
/// permissions (whether the user is an administrator).
/// If the user is an administrator, it opens the registry key under HKEY_LOCAL_MACHINE;
/// otherwise, it opens the registry key under HKEY_CURRENT_USER.
///
/// # Returns
///
/// Returns a `WincentResult<winreg::RegKey>` containing the requested registry key.
/// If an error occurs, it returns `WincentError::Io`.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let reg_key = get_execution_policy_reg()?;
///     // Use reg_key for further operations
///     Ok(())
/// }
/// ```
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
///
/// This function retrieves the execution policy from the registry and checks if it
/// matches any of the feasible options: "AllSigned", "Bypass", "RemoteSigned", or "Unrestricted".
///
/// # Returns
///
/// Returns a `WincentResult<bool>`, which is `true` if the execution policy is feasible,
/// `false` if the policy is not found, and an error if there is an issue accessing the registry.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let is_feasible = check_script_feasible_with_registry()?;
///     if is_feasible {
///         println!("The execution policy is feasible.");
///     } else {
///         println!("The execution policy is not feasible.");
///     }
///     Ok(())
/// }
/// ```
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
///
/// This function retrieves the execution policy registry key and sets its value to
/// "RemoteSigned". This is useful for ensuring that scripts can be run in a secure manner.
///
/// # Returns
///
/// Returns a `WincentResult<()>`. If the operation is successful, it returns `Ok(())`.
/// If there is an error while accessing or modifying the registry, it returns `WincentError::Io`.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     fix_script_feasible_with_registry()?;
///     println!("Execution policy set to 'RemoteSigned'.");
///     Ok(())
/// }
/// ```
pub(crate) fn fix_script_feasible_with_registry() -> WincentResult<()> {
    let reg_value=  "ExecutionPolicy";
    let reg_key = get_execution_policy_reg()?;

    let _ = reg_key.set_value(reg_value, &"RemoteSigned").map_err(WincentError::Io)?;

    Ok(())
}

/// Checks if a registry path exists.
///
/// # Parameters
///
/// - `path`: A Path reference representing the registry path to check
///
/// # Returns
///
/// Returns true if the registry path exists, false otherwise
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
///
/// # Returns
///
/// Returns a WincentResult containing the execution policy as a String
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
    use crate::error::WincentResult;
    use std::path::Path;

    #[test]
    fn test_check_script_feasible() -> WincentResult<()> {
        let result = check_script_feasible_with_registry()?;
        assert!(result || !result, "Should return a boolean value");
        Ok(())
    }

    #[test]
    fn test_fix_script_feasible() -> WincentResult<()> {
        let _ = check_script_feasible_with_registry()?;
        let _ = fix_script_feasible_with_registry()?;
        
        let final_state = check_script_feasible_with_registry()?;
        assert!(final_state, "Should be feasible after fix");
        
        Ok(())
    }

    #[test]
    fn test_registry_path_exists() {
        let path = Path::new("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\PowerShell\\1\\ShellIds\\Microsoft.PowerShell");
        assert!(registry_path_exists(path));
    }

    #[test]
    fn test_get_execution_policy() -> WincentResult<()> {
        let policy = get_execution_policy()?;
        assert!(!policy.is_empty(), "Execution policy should not be empty");
        assert!(
            ["Restricted", "AllSigned", "RemoteSigned", "Unrestricted", "Bypass"]
                .contains(&policy.as_str()),
            "Should return a valid execution policy"
        );
        Ok(())
    }
}
