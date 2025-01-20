use crate::{
    error::WincentError,
    scripts::{execute_ps_script, Script},
    utils, WincentResult,
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
        let hklm_reg_key = hklm
            .open_subkey_with_flags(key_path, winreg::enums::KEY_READ | winreg::enums::KEY_WRITE)
            .map_err(WincentError::Io)?;
        Ok(hklm_reg_key)
    } else {
        let (hkcu_reg_key, _) = hkcu
            .create_subkey_with_flags(key_path, winreg::enums::KEY_READ | winreg::enums::KEY_WRITE)
            .map_err(WincentError::Io)?;
        Ok(hkcu_reg_key)
    }
}

/// Checks if the PowerShell execution policy is feasible based on the registry value.
pub(crate) fn check_script_feasible_with_registry() -> WincentResult<bool> {
    let reg_value = "ExecutionPolicy";
    let reg_key = get_execution_policy_reg()?;
    let feasible_options = ["AllSigned", "Bypass", "RemoteSigned", "Unrestricted"];

    match reg_key.get_raw_value(reg_value) {
        Ok(val) => {
            // raw reg value vec will contains '\n' between characters, needs to filter
            let filtered_vec: Vec<u8> = val.bytes.into_iter().filter(|&x| x != 0).collect();
            let val_in_string = String::from_utf8_lossy(&filtered_vec).to_string();
            let _res = feasible_options.contains(&val_in_string.as_str());
            Ok(feasible_options.contains(&val_in_string.as_str()))
        }
        Err(e) => {
            if e.kind() == std::io::ErrorKind::NotFound {
                Ok(false)
            } else {
                Err(WincentError::Io(e))
            }
        }
    }
}

/// Sets the PowerShell execution policy to "RemoteSigned" in the registry.
pub(crate) fn fix_script_feasible_with_registry() -> WincentResult<()> {
    let reg_value = "ExecutionPolicy";
    let reg_key = get_execution_policy_reg()?;

    reg_key
        .set_value(reg_value, &"RemoteSigned")
        .map_err(WincentError::Io)?;

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
    use winreg::enums::*;
    use winreg::RegKey;

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
        }
        Err(e) => Err(WincentError::Io(e)),
    }
}

/****************************************************** Feature Feasible ******************************************************/

/// Checks if PowerShell script execution is feasible on the current system.
///
/// # Returns
///
/// Returns `true` if script execution is allowed, `false` otherwise.
///
/// # Example
///
/// ```rust
/// use wincent::{feasible::check_script_feasible, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     if check_script_feasible()? {
///         println!("PowerShell scripts can be executed");
///     } else {
///         println!("PowerShell script execution is restricted");
///     }
///     Ok(())
/// }
/// ```
pub fn check_script_feasible() -> WincentResult<bool> {
    check_script_feasible_with_registry()
}

/// Fixes PowerShell script execution policy to allow script execution.
///
/// # Example
///
/// ```rust
/// use wincent::{
///     feasible::{check_script_feasible, fix_script_feasible},
///     error::WincentError
/// };
///
/// fn main() -> Result<(), WincentError> {
///     if !check_script_feasible()? {
///         fix_script_feasible()?;
///         assert!(check_script_feasible()?);
///     }
///     Ok(())
/// }
/// ```
pub fn fix_script_feasible() -> WincentResult<()> {
    fix_script_feasible_with_registry()
}

/// Checks if Quick Access query operations are feasible on the current system.
///
/// # Returns
///
/// Returns `true` if Quick Access query operations are supported, `false` otherwise.
///
/// # Example
///
/// ```rust
/// use wincent::{feasible::check_query_feasible, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     if check_query_feasible()? {
///         println!("Quick Access query operations are supported");
///     } else {
///         println!("Quick Access query operations are not supported");
///     }
///     Ok(())
/// }
/// ```
pub fn check_query_feasible() -> WincentResult<bool> {
    check_query_feasible_with_script()
}

/// Checks if pin/unpin operations are feasible on the current system.
///
/// # Returns
///
/// Returns `true` if pin/unpin operations are supported, `false` otherwise.
///
/// # Example
///
/// ```no_run
/// use wincent::{feasible::check_pinunpin_feasible, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     if check_pinunpin_feasible()? {
///         println!("Pin/unpin operations are supported");
///     } else {
///         println!("Pin/unpin operations are not supported");
///     }
///     Ok(())
/// }
/// ```
pub fn check_pinunpin_feasible() -> WincentResult<bool> {
    check_pinunpin_feasible_with_script()
}

/// Checks if all Quick Access operations are feasible on the current system.
///
/// # Returns
///
/// Returns `true` only if all operations are supported, `false` otherwise.
///
/// # Example
///
/// ```no_run
/// use wincent::{feasible::{check_feasible, fix_feasible}, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     if !check_feasible()? {
///         println!("Some Quick Access operations are not supported");
///         // Try to fix the issues
///         if fix_feasible()? {
///             println!("Successfully enabled Quick Access operations");
///         }
///     }
///     Ok(())
/// }
/// ```
pub fn check_feasible() -> WincentResult<bool> {
    // First check script execution policy
    if !check_script_feasible()? {
        return Ok(false);
    }

    // Then check both operations
    let query_ok = check_query_feasible()?;
    let pinunpin_ok = check_pinunpin_feasible()?;

    Ok(query_ok && pinunpin_ok)
}

/// Attempts to fix Quick Access operation feasibility issues.
///
/// # Returns
///
/// Returns `true` if all operations are successfully enabled, `false` otherwise.
///
/// # Example
///
/// ```no_run
/// use wincent::{feasible::fix_feasible, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     match fix_feasible()? {
///         true => println!("Successfully enabled Quick Access operations"),
///         false => println!("Failed to enable some Quick Access operations")
///     }
///     Ok(())
/// }
/// ```
pub fn fix_feasible() -> WincentResult<bool> {
    fix_script_feasible()?;
    check_feasible()
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
        assert_eq!(
            final_policy, "RemoteSigned",
            "Execution policy should be set to RemoteSigned"
        );

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
        assert!(
            registry_path_exists(valid_path),
            "Valid PowerShell registry path should exist"
        );

        let invalid_paths = [
            Path::new("INVALID_KEY\\Path"),
            Path::new("HKEY_LOCAL_MACHINE"),
            Path::new("HKEY_CURRENT_USER\\NonExistentPath"),
        ];

        for path in &invalid_paths {
            assert!(
                !registry_path_exists(path),
                "Invalid path should return false: {:?}",
                path
            );
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
            "Bypass",
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
    #[ignore]
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
