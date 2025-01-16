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
}

pub type WincentResult<T> = Result<T, WincentError>; 

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Error, ErrorKind};

    #[test]
    fn test_error_conversions() {
        let io_error = Error::new(ErrorKind::NotFound, "file not found");
        let wincent_error = WincentError::from(io_error);
        assert!(matches!(wincent_error, WincentError::Io(_)));

        let missing_param = WincentError::MissingParemeter;
        assert!(format!("{}", missing_param).contains("Missing parameter"));

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
