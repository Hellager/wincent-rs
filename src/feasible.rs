use crate::{utils, WincentError};

/// Retrieves the execution policy registry key for PowerShell.
///
/// This function attempts to access the PowerShell execution policy
/// registry key located at `SOFTWARE\\Microsoft\\PowerShell\\1\\ShellIds\\Microsoft.PowerShell`.
/// It checks if the current user has administrative privileges to determine
/// whether to access the key in the `HKEY_LOCAL_MACHINE` (HKLM) hive or
/// the `HKEY_CURRENT_USER` (HKCU) hive.
///
/// If the user is an administrator, it opens the registry key in HKLM with
/// read and write access. If the user is not an administrator, it creates
/// the registry key in HKCU with the same access rights.
///
/// # Errors
///
/// This function returns a `WincentError` if there is an issue opening or
/// creating the registry key, such as insufficient permissions or other I/O errors.
///
/// # Returns
///
/// Returns a `Result<winreg::RegKey, WincentError>`, where `Ok` contains
/// the opened or created registry key, and `Err` contains the error information.
///
/// # Example
///
/// ```rust
/// match get_execution_policy_reg() {
///     Ok(reg_key) => {
///         // Use the registry key
///     },
///     Err(e) => {
///         eprintln!("Failed to get execution policy registry key: {:?}", e);
///     }
/// }
/// ```
fn get_execution_policy_reg() -> Result<winreg::RegKey, WincentError> {
    use winreg::enums::*;
    use winreg::RegKey;

    let key_path = "SOFTWARE\\Microsoft\\PowerShell\\1\\ShellIds\\Microsoft.PowerShell";
    let hkcu = RegKey::predef(HKEY_CURRENT_USER);
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);

    if utils::is_admin() {
        let hklm_reg_key = hklm.open_subkey_with_flags(key_path, winreg::enums::KEY_READ | winreg::enums::KEY_WRITE).map_err(WincentError::IoError)?;
        return Ok(hklm_reg_key);
    } else {
        let (hkcu_reg_key, _) = hkcu.create_subkey_with_flags(key_path, winreg::enums::KEY_READ | winreg::enums::KEY_WRITE)
            .map_err(WincentError::IoError)?;
        return Ok(hkcu_reg_key);
    }
}

/// Checks if the current PowerShell execution policy is feasible.
///
/// This function retrieves the execution policy from the registry and
/// checks if it matches one of the predefined feasible options:
/// "AllSigned", "Bypass", "RemoteSigned", or "Unrestricted".
///
/// It first calls `get_execution_policy_reg` to obtain the appropriate
/// registry key. If the registry value for "ExecutionPolicy" is found,
/// it filters out any null bytes and converts the value to a string.
/// The function then checks if the resulting execution policy is one of
/// the feasible options.
///
/// # Errors
///
/// This function returns a `WincentError` if there is an issue accessing
/// the registry or if an I/O error occurs. If the registry value is not
/// found, it returns `Ok(false`.
///
/// # Returns
///
/// Returns a `Result<bool, WincentError>`, where `Ok(true)` indicates
/// that the execution policy is feasible, `Ok(false)` indicates that
/// it is not feasible or not found, and `Err` contains the error information.
///
/// # Example
///
/// ```rust
/// match check_feasible() {
///     Ok(is_feasible) => {
///         if is_feasible {
///             println!("The execution policy is feasible.");
///         } else {
///             println!("The execution policy is not feasible.");
///         }
///     },
///     Err(e) => {
///         eprintln!("Error checking execution policy: {:?}", e);
///     }
/// }
/// ```
pub(crate) fn check_script_feasible_with_registry() -> Result<bool, WincentError> {
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
                return Err(WincentError::IoError(e));
            }
        }
    }
}

/// Sets the PowerShell execution policy to "RemoteSigned".
///
/// This function attempts to update the execution policy in the registry
/// to "RemoteSigned" for the current user or the local machine, depending
/// on the user's administrative privileges. It first retrieves the appropriate
/// registry key by calling `get_execution_policy_reg`. Then, it sets the
/// value of the "ExecutionPolicy" registry entry to "RemoteSigned".
///
/// # Errors
///
/// This function returns a `WincentError` if there is an issue accessing
/// the registry or if an I/O error occurs while setting the registry value.
///
/// # Returns
///
/// Returns a `Result<(), WincentError>`, where `Ok(())` indicates that the
/// execution policy was successfully set, and `Err` contains the error information.
///
/// # Example
///
/// ```rust
/// match fix_feasible() {
///     Ok(()) => {
///         println!("Execution policy successfully set to 'RemoteSigned'.");
///     },
///     Err(e) => {
///         eprintln!("Error setting execution policy: {:?}", e);
///     }
/// }
/// ```
pub(crate) fn fix_script_feasible_with_registry() -> Result<(), WincentError> {
    let reg_value=  "ExecutionPolicy";
    let reg_key = get_execution_policy_reg()?;

    let _ = reg_key.set_value(reg_value, &"RemoteSigned").map_err(WincentError::IoError)?;

    Ok(())
}

#[cfg(test)]
mod feasible_test {
    use utils::init_test_logger;
    use log::{debug, error};
    use super::*;
    
    #[test]
    fn test_feasible() -> Result<(), WincentError> {
        init_test_logger();
        match check_script_feasible_with_registry() {
            Ok(is_feasible) => {
                if !is_feasible {
                    debug!("wincent not feasible, try fix");
                    if let Err(e) = fix_script_feasible_with_registry() {
                        error!("Failed to fix feasibility: {:?}", e);
                    }
                }
    
                match check_script_feasible_with_registry() {
                    Ok(double_check) => {
                        if double_check {
                            assert!(true);
                            debug!("wincent fix feasible success");                            
                        } else {
                            assert!(false);
                            error!("wincent fix feasible failed");
                        }
                    },
                    Err(e) => {
                        panic!("Error during double check: {:?}", e);
                    },
                }
    
                Ok(())
            },
            Err(e) => Err(e),
        }
    }
}
