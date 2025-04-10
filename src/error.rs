//! Error handling and error type definitions
//!
//! Provides centralized error handling for Windows Quick Access operations,
//! with unified error types and automatic error conversions.
//!
//! # Key Features
//! - Cross-domain error classification (I/O, system, scripting)
//! - Automatic error conversion from common Rust types
//! - Detailed error context preservation
//! - Async-task error propagation support
//! - Comprehensive test utilities
//!
//! # Error Handling Guide
//! 1. Use `WincentResult<T>` as return type for fallible operations
//! 2. Leverage `?` operator for automatic error conversion
//! 3. Add context to errors using `.map_err` before returning
//! 4. Match specific error variants for recovery attempts

use thiserror::Error;

#[derive(Error, Debug)]
pub enum WincentError {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("UTF-8 conversion error: {0}")]
    Utf8(#[from] std::string::FromUtf8Error),

    #[error("PowerShell execution failed: {0}")]
    PowerShellExecution(String),

    #[error("Invalid path: {0}")]
    InvalidPath(String),

    #[error("Operation not supported: {0}")]
    UnsupportedOperation(String),

    #[error("System error: {0}")]
    SystemError(String),

    #[error("Array conversion error: {0}")]
    ArrayConversion(#[from] std::array::TryFromSliceError),

    #[error("Script failed error: {0}")]
    ScriptFailed(String),

    #[error("Unknown quick access type: {0}")]
    UnknownQuickAccessType(u32),

    #[error("Unknown script method: {0}")]
    UnknownScriptMethod(u32),

    #[error("Missing function parameter")]
    MissingParemeter,

    #[error("Windows API error: {0}")]
    WindowsApi(i32),

    #[error("Script strategy not found: {0}")]
    ScriptStrategyNotFound(String),

    #[error("Async execution error: {0}")]
    AsyncExecution(String),

    #[error("Operation timed out: {0}")]
    Timeout(String),
}

impl From<windows::core::Error> for WincentError {
    fn from(err: windows::core::Error) -> Self {
        WincentError::WindowsApi(err.code().0)
    }
}

impl From<tokio::task::JoinError> for WincentError {
    fn from(err: tokio::task::JoinError) -> Self {
        WincentError::AsyncExecution(err.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::WincentResult;
    use std::io::{Error, ErrorKind};

    #[test]
    fn test_error_conversions() {
        let io_error = Error::new(ErrorKind::NotFound, "file not found");
        let wincent_error = WincentError::from(io_error);
        assert!(matches!(wincent_error, WincentError::Io(_)));

        let missing_param = WincentError::MissingParemeter;
        assert!(format!("{}", missing_param).contains("Missing function parameter"));

        let invalid_path = WincentError::InvalidPath("test/path".to_string());
        assert!(format!("{}", invalid_path).contains("test/path"));

        let ps_error = WincentError::PowerShellExecution("access denied".to_string());
        assert!(format!("{}", ps_error).contains("access denied"));
    }

    #[test]
    fn test_result_type() {
        let success: WincentResult<()> = Ok(());
        assert!(success.is_ok());

        let failure: WincentResult<()> = Err(WincentError::MissingParemeter);
        assert!(failure.is_err());
    }
}
