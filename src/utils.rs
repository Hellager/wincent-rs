#![allow(dead_code)]

use crate::{check_feasible, WincentError};
use windows::Win32::Foundation::BOOL;
use windows::Win32::UI::Shell::IsUserAnAdmin;

/// Initializes the logger for testing purposes.
///
/// This function configures the logger to output log messages to standard output (stdout).
/// It sets the log level filter to `Trace`, which means all log messages at this level
/// and above will be displayed. The logger is marked as being in a test context, which
/// can affect its behavior in certain logging frameworks.
///
/// This function is intended to be called in test setups to ensure that log messages
/// are visible during test execution, aiding in debugging and verification of test
/// outcomes.
///
/// # Example
///
/// ```
/// #[cfg(test)]
/// mod tests {
///     use super::*;
///
///     #[test]
///     fn test_logging() {
///         init_test_logger();
///         log::trace!("This is a trace message.");
///         // Additional test code...
///     }
/// }
/// ```
pub(crate) fn init_test_logger() {
    let _ = env_logger::builder()
    .target(env_logger::Target::Stdout)
    .filter_level(log::LevelFilter::Trace)
    .is_test(true)
    .try_init();
}

/// Checks if the current user has administrative privileges.
///
/// This function utilizes the Windows API to determine if the
/// calling user is an administrator. It calls the `IsUserAnAdmin`
/// function from the `windows` crate, which returns a boolean
/// value indicating whether the user has admin rights.
///
/// # Safety
///
/// This function is marked as `unsafe` because it directly calls
/// an external function from the Windows API. The caller must ensure
/// that the environment is appropriate for this call, as improper
/// usage could lead to undefined behavior.
///
/// # Returns
///
/// Returns `true` if the current user is an administrator, and
/// `false` otherwise.
///
/// # Example
///
/// ```rust
/// if is_admin() {
///     println!("User has administrative privileges.");
/// } else {
///     println!("User does not have administrative privileges.");
/// }
/// ```
pub(crate) fn is_admin() -> bool {
    unsafe {
        IsUserAnAdmin() == BOOL(1)
    }
}

/// Refreshes all open Windows Explorer windows asynchronously.
///
/// This function constructs and executes a PowerShell script that refreshes all
/// currently open Windows Explorer windows. It uses the `powershell_script` crate
/// to build and run the PowerShell script in a non-interactive manner.
///
/// # Errors
///
/// This function returns a `Result<(), WincentError>`, which can be:
/// - `Ok(())` if the operation was successful.
/// - `Err(WincentError::ScriptError)` if there was an error executing the PowerShell script.
/// - `Err(WincentError::TimeoutError)` if the operation timed out.
/// - `Err(WincentError::ExecuteError)` if there was an error during execution.
///
/// # Example
///
/// ```rust
/// match refresh_explorer_window().await {
///     Ok(()) => println!("Explorer windows refreshed successfully."),
///     Err(e) => eprintln!("Failed to refresh explorer windows: {:?}", e),
/// }
/// ```
///
/// # Notes
///
/// The PowerShell script executed by this function sets the output encoding to UTF-8,
/// creates a Shell.Application COM object, retrieves all open windows, and refreshes each one.
/// The script is run in a blocking manner using `tokio::task::spawn_blocking`, and a timeout
/// is applied to ensure that the operation does not hang indefinitely.
pub(crate) async fn refresh_explorer_window() -> Result<(), WincentError> {
    use powershell_script::PsScriptBuilder;
    use std::io::{Error, ErrorKind};

    if !check_feasible()? {
        return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
    }

    const SCRIPT: &str = r#"
        $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
        $shellApplication = New-Object -ComObject Shell.Application;
        $windows = $shellApplication.Windows();
        $windows | ForEach-Object { $_.Refresh() }
    "#;

    let ps = PsScriptBuilder::new()
        .no_profile(true)
        .non_interactive(true)
        .hidden(false)
        .print_commands(false)
        .build();

    let handle = tokio::task::spawn_blocking(move || {
        ps.run(SCRIPT)
            .map(|_| ())
            .map_err(WincentError::ScriptError)
    });

    tokio::time::timeout(tokio::time::Duration::from_secs(crate::SCRIPT_TIMEOUT), handle)
        .await
        .map_err(WincentError::TimeoutError)?
        .map_err(WincentError::ExecuteError)?
}

#[cfg(test)]
mod utils_test {
    use log::debug;
    use super::*;
    
    #[test]
    fn test_logger() {
        init_test_logger();
        debug!("test logger init success");
    }

    #[test]
    fn test_check_admin() {
        init_test_logger();
        let is_admin = is_admin();
        debug!("has admin priveledge: {}", is_admin);
    }

    #[tokio::test]
    async fn test_refresh_explorer() -> Result<(), WincentError> {
        refresh_explorer_window().await
    }
}
