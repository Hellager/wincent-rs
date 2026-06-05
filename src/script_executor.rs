//! PowerShell script execution and output parsing helpers.

use crate::error::WincentError;
use crate::script_storage::ScriptStorage;
use crate::script_strategy::PSScript;
use crate::utils::get_windows_recent_folder;
use crate::WincentResult;
use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::{Command, Output, Stdio};
use std::thread;
use std::time::{Duration, Instant};

const RECENT_FILES_AUTOMATIC_DESTINATION: &str = "5f7b5f1e01b83767.automaticDestinations-ms";
const DEFAULT_SCRIPT_TIMEOUT: Duration = Duration::from_secs(10);

/// PowerShell script executor.
pub(crate) struct ScriptExecutor;

impl ScriptExecutor {
    /// Executes a PowerShell script synchronously.
    pub fn execute_ps_script(
        script_type: PSScript,
        parameter: Option<&str>,
    ) -> WincentResult<Output> {
        Self::execute_ps_script_with_timeout(script_type, parameter, DEFAULT_SCRIPT_TIMEOUT)
    }

    /// Executes a PowerShell script synchronously with timeout protection.
    pub fn execute_ps_script_with_timeout(
        script_type: PSScript,
        parameter: Option<&str>,
        timeout: Duration,
    ) -> WincentResult<Output> {
        if timeout.is_zero() {
            return Err(WincentError::InvalidArgument(
                "timeout must be greater than zero".to_string(),
            ));
        }

        let script_path = match parameter {
            Some(param) => ScriptStorage::get_dynamic_script_path(script_type, param)?,
            None => ScriptStorage::get_script_path(script_type)?,
        };

        let start = Instant::now();

        let mut child = Command::new("powershell")
            .args([
                // Process-scoped policy override for generated temp scripts.
                // This does not change machine or user policy, but environments
                // that forbid Bypass may still reject script execution.
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                script_path.to_str().ok_or_else(|| {
                    WincentError::invalid_path_reason("Failed to convert script path")
                })?,
            ])
            .stdout(Stdio::piped())
            .stderr(Stdio::piped())
            .spawn()
            .map_err(|e| {
                Self::powershell_io_error(script_type, &script_path, parameter, start.elapsed(), e)
            })?;

        let stdout = child.stdout.take().ok_or_else(|| {
            WincentError::SystemError("Failed to capture PowerShell stdout".to_string())
        })?;
        let stderr = child.stderr.take().ok_or_else(|| {
            WincentError::SystemError("Failed to capture PowerShell stderr".to_string())
        })?;

        let stdout_reader = thread::spawn(move || {
            let mut reader = stdout;
            let mut buffer = Vec::new();
            let _ = reader.read_to_end(&mut buffer);
            buffer
        });
        let stderr_reader = thread::spawn(move || {
            let mut reader = stderr;
            let mut buffer = Vec::new();
            let _ = reader.read_to_end(&mut buffer);
            buffer
        });

        loop {
            if let Some(status) = child.try_wait().map_err(|e| {
                Self::powershell_io_error(script_type, &script_path, parameter, start.elapsed(), e)
            })? {
                let stdout = stdout_reader.join().unwrap_or_default();
                let stderr = stderr_reader.join().unwrap_or_default();
                return Ok(Output {
                    status,
                    stdout,
                    stderr,
                });
            }

            if start.elapsed() >= timeout {
                let _ = child.kill();
                let _ = child.wait();
                let stdout = stdout_reader.join().unwrap_or_default();
                let stderr = stderr_reader.join().unwrap_or_default();
                return Err(Self::powershell_timeout_error(
                    script_type,
                    &script_path,
                    parameter,
                    start.elapsed(),
                    stdout,
                    stderr,
                    timeout,
                ));
            }

            thread::sleep(Duration::from_millis(10));
        }
    }

    fn powershell_io_error(
        script_type: PSScript,
        script_path: &Path,
        parameter: Option<&str>,
        duration: Duration,
        error: std::io::Error,
    ) -> WincentError {
        use crate::error::{PowerShellError, PowerShellErrorKind};

        let io_error = Some(error.kind());
        let os_error = error.raw_os_error();
        let kind = if let Some(5) = os_error {
            PowerShellErrorKind::AccessDenied
        } else if error.kind() == std::io::ErrorKind::PermissionDenied {
            PowerShellErrorKind::AccessDenied
        } else {
            PowerShellErrorKind::ProcessFailed
        };

        let mut builder = PowerShellError::builder(script_type.operation())
            .kind(kind)
            .exit_code(None)
            .stderr(error.to_string())
            .script_path(script_path)
            .duration(duration);

        if let Some(parameter) = parameter {
            builder = builder.parameters(parameter);
        }
        if let Some(io_error) = io_error {
            builder = builder.io_error(io_error);
        }
        if let Some(os_error) = os_error {
            builder = builder.os_error(os_error);
        }

        WincentError::PowerShellExecution(Box::new(builder.build()))
    }

    fn powershell_timeout_error(
        script_type: PSScript,
        script_path: &Path,
        parameter: Option<&str>,
        duration: Duration,
        stdout: Vec<u8>,
        stderr: Vec<u8>,
        timeout: Duration,
    ) -> WincentError {
        use crate::error::{PowerShellError, PowerShellErrorKind};

        let mut stderr_text = String::from_utf8_lossy(&stderr).to_string();
        if !stderr_text.is_empty() {
            stderr_text.push('\n');
        }
        stderr_text.push_str(&format!(
            "PowerShell execution timed out after {}s",
            timeout.as_secs_f64()
        ));

        let mut builder = PowerShellError::builder(script_type.operation())
            .kind(PowerShellErrorKind::Timeout)
            .exit_code(None)
            .stdout(String::from_utf8_lossy(&stdout).to_string())
            .stderr(stderr_text)
            .script_path(script_path)
            .duration(duration);

        if let Some(parameter) = parameter {
            builder = builder.parameters(parameter);
        }

        WincentError::PowerShellExecution(Box::new(builder.build()))
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

            let mut error = PowerShellError::builder(script_type.operation())
                .kind(kind)
                .exit_code(output.status.code())
                .stdout(stdout)
                .stderr(stderr)
                .script_path(script_path)
                .duration(duration);

            if let Some(parameters) = parameters {
                error = error.parameters(parameters);
            }

            return Err(WincentError::PowerShellExecution(Box::new(error.build())));
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
            assert_eq!(ps_err.raw_stderr(), "Error message");
            assert_eq!(ps_err.exit_code(), Some(1));
            assert_eq!(ps_err.operation(), PowerShellOperation::QueryQuickAccess);
            assert_eq!(
                ps_err.kind(),
                crate::error::PowerShellErrorKind::ProcessFailed
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
                err.kind(),
                err.raw_stderr()
            );
        } else {
            panic!("Expected PowerShellExecution error");
        }
    }

    #[test]
    fn test_execute_ps_script_with_timeout_returns_timeout_error() {
        let result = ScriptExecutor::execute_ps_script_with_timeout(
            PSScript::QueryQuickAccess,
            None,
            Duration::from_nanos(1),
        );

        match result {
            Err(WincentError::PowerShellExecution(err)) => {
                assert_eq!(err.kind(), crate::error::PowerShellErrorKind::Timeout);
                assert!(err.is_transient());
                assert_eq!(err.operation(), PowerShellOperation::QueryQuickAccess);
            }
            other => panic!("Expected PowerShell timeout error, got: {:?}", other),
        }
    }
}
