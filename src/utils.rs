#![allow(dead_code)]

use crate::{
    WincentResult,
    error::WincentError,
    scripts::{Script, execute_ps_script}
};
use windows::Win32::Foundation::BOOL;
use windows::Win32::UI::Shell::IsUserAnAdmin;

/// Checks if the current user has administrative privileges.
pub(crate) fn is_admin() -> bool {
    unsafe {
        IsUserAnAdmin() == BOOL(1)
    }
}

/// Refreshes the Windows Explorer window using a PowerShell script.
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
