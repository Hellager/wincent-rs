//! Check and fix functions execution feasibility.
//!
//! ## Example
//!
//! ```no_run
//! use wincent::{
//!     feasible::{check_feasible, fix_feasible},
//!     error::WincentError,
//! };
//!
//! fn main() -> Result<(), WincentError> {
//!     if !check_feasible()? {
//!         println!("Some Quick Access operations are not supported");
//!         // Try to fix the issues
//!         if fix_feasible()? {
//!             println!("Successfully enabled Quick Access operations");
//!         }
//!     }
//!     Ok(())
//! }
//! ```

use crate::{
    script_executor::ScriptExecutor,
    script_strategy::PSScript,
    WincentResult,
};

/// Checks if PowerShell query commands are available and executable.
pub(crate) fn check_query_feasible_with_script() -> WincentResult<bool> {
    let output = ScriptExecutor::execute_ps_script(PSScript::CheckQueryFeasible, None)?;

    Ok(output.status.success())
}

/// Checks if PowerShell pin/unpin commands are available and executable.
pub(crate) fn check_pinunpin_feasible_with_script() -> WincentResult<bool> {
    let output = ScriptExecutor::execute_ps_script(PSScript::CheckPinUnpinFeasible, None)?;

    Ok(output.status.success())
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test_log::test]
    fn test_check_query_feasible_with_script() -> WincentResult<()> {
        let result = check_query_feasible_with_script()?;

        println!("Query feasibility check result: {}", result);

        Ok(())
    }

    #[test_log::test]
    #[ignore = "Modifies system state"]
    fn test_check_pinunpin_feasible_with_script() -> WincentResult<()> {
        let result = check_pinunpin_feasible_with_script()?;

        println!("Pin/Unpin feasibility check result: {}", result);

        Ok(())
    }
}
