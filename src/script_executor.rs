//! PowerShell script execution and output parsing helpers.

use crate::error::WincentError;
use crate::script_storage::ScriptStorage;
use crate::script_strategy::PSScript;
use crate::utils::get_windows_recent_folder;
use crate::WincentResult;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::{Command, Output};
use std::time::Duration;

const RECENT_FILES_AUTOMATIC_DESTINATION: &str = "5f7b5f1e01b83767.automaticDestinations-ms";

/// PowerShell script executor.
pub(crate) struct ScriptExecutor;

impl ScriptExecutor {
    /// Executes a PowerShell script synchronously.
    pub fn execute_ps_script(
        script_type: PSScript,
        parameter: Option<&str>,
    ) -> WincentResult<Output> {
        let script_path = match parameter {
            Some(param) => ScriptStorage::get_dynamic_script_path(script_type, param)?,
            None => ScriptStorage::get_script_path(script_type)?,
        };

        let start = std::time::Instant::now();

        Command::new("powershell")
            .args([
                // Process-scoped policy override for generated temp scripts.
                // This does not change machine or user policy, but environments
                // that forbid Bypass may still reject script execution.
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                script_path.to_str().ok_or_else(|| {
                    WincentError::InvalidPath("Failed to convert script path".to_string())
                })?,
            ])
            .output()
            .map_err(|e| {
                use crate::error::{PowerShellError, PowerShellErrorKind};

                let io_error = Some(e.kind());
                let os_error = e.raw_os_error();
                let kind = if let Some(5) = os_error {
                    PowerShellErrorKind::AccessDenied
                } else if e.kind() == std::io::ErrorKind::PermissionDenied {
                    PowerShellErrorKind::AccessDenied
                } else {
                    PowerShellErrorKind::ExecutionFailed
                };

                WincentError::PowerShellExecution(PowerShellError {
                    kind,
                    operation: script_type.operation(),
                    exit_code: None,
                    stdout: String::new(),
                    stderr: e.to_string(),
                    script_path: script_path.clone(),
                    parameters: parameter.map(|s| s.to_string()),
                    duration: Some(start.elapsed()),
                    io_error,
                    os_error,
                })
            })
    }

    /// Parses script output into a string collection.
    pub fn parse_output_to_strings(
        output: Output,
        script_type: PSScript,
        script_path: PathBuf,
        parameters: Option<String>,
        duration: Duration,
    ) -> WincentResult<Vec<String>> {
        if !output.status.success() {
            use crate::error::PowerShellError;
            let stderr = String::from_utf8_lossy(&output.stderr).to_string();
            let stdout = String::from_utf8_lossy(&output.stdout).to_string();
            let kind = PowerShellError::infer_kind_from_stderr(&stderr);

            return Err(WincentError::PowerShellExecution(PowerShellError {
                kind,
                operation: script_type.operation(),
                exit_code: output.status.code(),
                stdout,
                stderr,
                script_path,
                parameters,
                duration: Some(duration),
                io_error: None,
                os_error: None,
            }));
        }

        let stdout = String::from_utf8_lossy(&output.stdout);
        let lines = stdout
            .lines()
            .map(|s| s.trim().to_string())
            .filter(|s| !s.is_empty())
            .collect();

        Ok(lines)
    }
}

/// Windows Quick Access data file information.
pub(crate) struct QuickAccessDataFiles {
    recent_files_path: PathBuf,
}

impl QuickAccessDataFiles {
    /// Retrieves Quick Access data file paths.
    pub fn new() -> WincentResult<Self> {
        let recent_folder = get_windows_recent_folder()?;
        let automatic_dest_dir = Path::new(&recent_folder).join("AutomaticDestinations");

        Ok(Self {
            // Explorer's Recent Files automatic destination AppID hash on
            // supported Windows 10/11 builds. This is an implementation detail
            // of the shell and may need updating if Windows changes the AppID.
            recent_files_path: automatic_dest_dir.join(RECENT_FILES_AUTOMATIC_DESTINATION),
        })
    }

    fn remove_file(&self, path: &Path) -> WincentResult<()> {
        match fs::remove_file(path) {
            Ok(()) => Ok(()),
            Err(e) if e.kind() == std::io::ErrorKind::NotFound => Ok(()),
            Err(e) => Err(WincentError::Io(e)),
        }
    }

    /// Removes the Recent Files data file, ignoring missing-file errors.
    pub fn remove_recent_file(&self) -> WincentResult<()> {
        self.remove_file(&self.recent_files_path)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::error::PowerShellOperation;
    use std::os::windows::process::ExitStatusExt;

    #[test]
    fn test_output_parsing() {
        let output = Output {
            status: std::process::ExitStatus::from_raw(0),
            stdout: "Line1\nLine2\n\nLine3".as_bytes().to_vec(),
            stderr: Vec::new(),
        };

        let result = ScriptExecutor::parse_output_to_strings(
            output,
            PSScript::QueryQuickAccess,
            PathBuf::from("test.ps1"),
            None,
            Duration::from_millis(100),
        )
        .unwrap();
        assert_eq!(result.len(), 3);
        assert_eq!(result[0], "Line1");
        assert_eq!(result[1], "Line2");
        assert_eq!(result[2], "Line3");
    }

    #[test]
    fn test_error_output_handling() {
        let output = Output {
            status: std::process::ExitStatus::from_raw(1),
            stdout: Vec::new(),
            stderr: "Error message".as_bytes().to_vec(),
        };

        let result = ScriptExecutor::parse_output_to_strings(
            output,
            PSScript::QueryQuickAccess,
            PathBuf::from("test.ps1"),
            None,
            Duration::from_millis(100),
        );
        assert!(result.is_err());
        if let Err(WincentError::PowerShellExecution(ps_err)) = result {
            assert_eq!(ps_err.stderr, "Error message");
            assert_eq!(ps_err.exit_code, Some(1));
            assert_eq!(ps_err.operation, PowerShellOperation::QueryQuickAccess);
            assert_eq!(
                ps_err.kind,
                crate::error::PowerShellErrorKind::ExecutionFailed
            );
        } else {
            panic!("Expected PowerShellExecution error");
        }
    }

    #[test]
    fn test_non_zero_exit_produces_transient_retryable_error() {
        let output = Output {
            status: std::process::ExitStatus::from_raw(1),
            stdout: Vec::new(),
            stderr: b"Shell namespace file is locked by another process".to_vec(),
        };

        let result = ScriptExecutor::parse_output_to_strings(
            output,
            PSScript::QueryQuickAccess,
            PathBuf::from("test.ps1"),
            None,
            Duration::from_millis(10),
        );

        assert!(result.is_err(), "Non-zero exit should produce an error");
        if let Err(WincentError::PowerShellExecution(err)) = result {
            assert!(
                err.is_transient(),
                "Error with 'locked' in stderr should be transient, got kind={:?} stderr={:?}",
                err.kind,
                err.stderr
            );
        } else {
            panic!("Expected PowerShellExecution error");
        }
    }
}
