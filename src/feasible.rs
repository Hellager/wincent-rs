//! System capability checks and feature enablement
//!
//! Provides runtime verification of system capabilities required for Quick Access operations,
//! with automatic remediation for common configuration issues.
//!
//! # Key Functionality
//! - PowerShell command availability checks
//! - Query operation capability verification
//! - Pin/Unpin operation system support detection
//! - Execution policy validation
//! - Automatic remediation of configuration issues
//!
//! # Usage Flow
//! 1. Check operation feasibility before execution
//! 2. Attempt automatic remediation if supported
//! 3. Provide fallback strategies when unavailable

use crate::{script_executor::ScriptExecutor, script_strategy::PSScript, WincentResult};

/// Checks if PowerShell query commands are available and executable.
pub(crate) fn check_query_feasible_powershell() -> WincentResult<bool> {
    let output = ScriptExecutor::execute_ps_script(PSScript::CheckQueryFeasible, None)?;

    Ok(output.status.success())
}

/// Checks if PowerShell pin/unpin commands are available and executable.
pub(crate) fn check_pinunpin_feasible_powershell() -> WincentResult<bool> {
    let output = ScriptExecutor::execute_ps_script(PSScript::CheckPinUnpinFeasible, None)?;

    Ok(output.status.success())
}

/// Checks if pin/unpin operations are feasible using native COM API
pub(crate) fn check_pinunpin_feasible_native() -> WincentResult<bool> {
    use crate::query::ComGuard;

    // Try to initialize COM
    match ComGuard::initialize() {
        Ok(_) => Ok(true),
        Err(_) => Ok(false),
    }
}

/****************************************************** Feature Feasible ******************************************************/

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
    check_query_feasible_powershell()
}

/// Checks if pin/unpin operations are feasible on the current system.
///
/// This function attempts to check using native COM API first, falling back to PowerShell if COM fails.
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
    // Try native COM first (fast path)
    check_pinunpin_feasible_native().or_else(|_| {
        // Fallback to PowerShell if COM fails (compatibility)
        check_pinunpin_feasible_powershell()
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test_log::test]
    fn test_check_query_feasible_powershell() -> WincentResult<()> {
        let result = check_query_feasible_powershell()?;

        println!("Query feasibility check result: {}", result);

        Ok(())
    }

    #[test_log::test]
    #[ignore = "Modifies system state"]
    fn test_check_pinunpin_feasible_powershell() -> WincentResult<()> {
        let result = check_pinunpin_feasible_powershell()?;

        println!("Pin/Unpin feasibility check result: {}", result);

        Ok(())
    }

    #[test_log::test]
    fn test_check_pinunpin_feasible_native() -> WincentResult<()> {
        let result = check_pinunpin_feasible_native()?;

        println!("Pin/Unpin native feasibility check result: {}", result);
        assert!(result, "Native COM API should be available on Windows");

        Ok(())
    }
}
