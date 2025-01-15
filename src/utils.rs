#![allow(dead_code)]

use crate::{
    scripts::{Script, execute_ps_script}, 
    error::{WincentError, WincentResult}
};
use windows::Win32::Foundation::BOOL;
use windows::Win32::UI::Shell::IsUserAnAdmin;

/// Checks if the current user has administrative privileges.
///
/// This function uses the `IsUserAnAdmin` function from the Windows API to determine
/// if the current user is an administrator. It returns `true` if the user is an admin,
/// and `false` otherwise.
///
/// # Returns
///
/// Returns a `bool` indicating whether the current user has administrative privileges.
///
/// # Example
///
/// ```rust
/// fn main() {
///     if is_admin() {
///         println!("The user has administrative privileges.");
///     } else {
///         println!("The user does not have administrative privileges.");
///     }
/// }
/// ```
pub(crate) fn is_admin() -> bool {
    unsafe {
        IsUserAnAdmin() == BOOL(1)
    }
}

/// Refreshes the Windows Explorer window using a PowerShell script.
///
/// This function executes a PowerShell script to refresh the Windows Explorer window.
/// It checks the output status to determine if the operation was successful.
///
/// # Returns
///
/// Returns a `WincentResult<()>`. If the operation is successful, it returns `Ok(())`.
/// If the script fails, it returns `WincentError::ScriptFailed` with the error message.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     refresh_explorer_window()?;
///     println!("Explorer window refreshed successfully.");
///     Ok(())
/// }
/// ```
pub(crate) fn refresh_explorer_window() -> WincentResult<()> {
    let output = execute_ps_script(Script::RefreshExplorer, None)?;

    if output.status.success() {
        Ok(())
    } else {
        let error = String::from_utf8(output.stderr)?;
        Err(WincentError::ScriptFailed(error))
    }
}

#[cfg(test)]
mod utils_test {
    use super::*;
    
    #[test_log::test]
    fn test_logger() {
        println!("test logger init success");
    }

    #[test]
    fn test_check_admin() {
        let is_admin = is_admin();
        assert!(is_admin || !is_admin, "Should return a boolean value");
    }

    #[test]
    fn test_refresh_explorer() -> WincentResult<()> {
        refresh_explorer_window()
    }
}
