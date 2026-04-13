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
//! - Localized error handling support
//!
//! # Error Handling Guide
//! 1. Use `WincentResult<T>` as return type for fallible operations
//! 2. Leverage `?` operator for automatic error conversion
//! 3. Add context to errors using `.map_err` before returning
//! 4. Match specific error variants for recovery attempts
//!
//! # Handling Localized Errors
//!
//! The library provides built-in classification for English error messages.
//! For other languages, you can use custom classifiers to handle localized
//! PowerShell error messages.
//!
//! ## Approach 1: Using `reclassify_with()`
//!
//! Reclassify errors after they occur:
//!
//! ```rust
//! use wincent::prelude::*;
//! use wincent::error::{PowerShellErrorKind, WincentError};
//!
//! fn classify_chinese_error(stderr: &str) -> Option<PowerShellErrorKind> {
//!     let stderr_lower = stderr.to_lowercase();
//!
//!     if stderr_lower.contains("拒绝访问") {
//!         return Some(PowerShellErrorKind::AccessDenied);
//!     }
//!     if stderr_lower.contains("执行策略") {
//!         return Some(PowerShellErrorKind::ExecutionPolicy);
//!     }
//!     if stderr_lower.contains("无法识别") {
//!         return Some(PowerShellErrorKind::CmdletNotFound);
//!     }
//!     if stderr_lower.contains("超时") {
//!         return Some(PowerShellErrorKind::Timeout);
//!     }
//!
//!     None
//! }
//!
//! # async fn example() -> WincentResult<()> {
//! # let manager = QuickAccessManager::new().await?;
//! match manager.add_item("C:\\test", QuickAccess::FrequentFolders, false).await {
//!     Ok(_) => println!("Success"),
//!     Err(WincentError::PowerShellExecution(err)) => {
//!         let err = err.reclassify_with(classify_chinese_error);
//!         if err.is_access_denied() {
//!             println!("需要管理员权限");
//!         }
//!     }
//!     Err(e) => println!("Other error: {}", e),
//! }
//! # Ok(())
//! # }
//! ```
//!
//! ## Approach 2: Using `stderr_contains()`
//!
//! Directly check stderr content for quick pattern matching:
//!
//! ```rust
//! # use wincent::error::WincentError;
//! # fn example(result: Result<(), WincentError>) {
//! match result {
//!     Err(WincentError::PowerShellExecution(err)) => {
//!         if err.stderr_contains("拒绝访问") || err.stderr_contains("access denied") {
//!             println!("Access denied in any language");
//!         }
//!     }
//!     _ => {}
//! }
//! # }
//! ```
//!
//! ## Approach 3: Using `raw_stderr()`
//!
//! Access raw stderr for complex custom analysis:
//!
//! ```rust
//! # use wincent::error::WincentError;
//! # fn example(result: Result<(), WincentError>) {
//! match result {
//!     Err(WincentError::PowerShellExecution(err)) => {
//!         let stderr = err.raw_stderr();
//!         // Perform custom regex matching, ML classification, etc.
//!         if stderr.contains("某个特定错误") {
//!             // Handle specific error
//!         }
//!     }
//!     _ => {}
//! }
//! # }
//! ```
//!
//! ## Common Language Patterns
//!
//! ### Chinese (中文)
//! ```rust
//! # use wincent::error::PowerShellErrorKind;
//! fn classify_chinese_error(stderr: &str) -> Option<PowerShellErrorKind> {
//!     let stderr_lower = stderr.to_lowercase();
//!     if stderr_lower.contains("拒绝访问") {
//!         return Some(PowerShellErrorKind::AccessDenied);
//!     }
//!     if stderr_lower.contains("执行策略") {
//!         return Some(PowerShellErrorKind::ExecutionPolicy);
//!     }
//!     if stderr_lower.contains("无法识别") {
//!         return Some(PowerShellErrorKind::CmdletNotFound);
//!     }
//!     if stderr_lower.contains("超时") {
//!         return Some(PowerShellErrorKind::Timeout);
//!     }
//!     None
//! }
//! ```
//!
//! ### French (Français)
//! ```rust
//! # use wincent::error::PowerShellErrorKind;
//! fn classify_french_error(stderr: &str) -> Option<PowerShellErrorKind> {
//!     let stderr_lower = stderr.to_lowercase();
//!     if stderr_lower.contains("accès refusé") || stderr_lower.contains("accès interdit") {
//!         return Some(PowerShellErrorKind::AccessDenied);
//!     }
//!     if stderr_lower.contains("stratégie d'exécution") {
//!         return Some(PowerShellErrorKind::ExecutionPolicy);
//!     }
//!     if stderr_lower.contains("n'est pas reconnu") {
//!         return Some(PowerShellErrorKind::CmdletNotFound);
//!     }
//!     if stderr_lower.contains("délai d'attente") || stderr_lower.contains("timeout") {
//!         return Some(PowerShellErrorKind::Timeout);
//!     }
//!     None
//! }
//! ```
//!
//! ### Russian (Русский)
//! ```rust
//! # use wincent::error::PowerShellErrorKind;
//! fn classify_russian_error(stderr: &str) -> Option<PowerShellErrorKind> {
//!     let stderr_lower = stderr.to_lowercase();
//!     if stderr_lower.contains("отказано в доступе") || stderr_lower.contains("доступ запрещен") {
//!         return Some(PowerShellErrorKind::AccessDenied);
//!     }
//!     if stderr_lower.contains("политика выполнения") {
//!         return Some(PowerShellErrorKind::ExecutionPolicy);
//!     }
//!     if stderr_lower.contains("не распознается") {
//!         return Some(PowerShellErrorKind::CmdletNotFound);
//!     }
//!     if stderr_lower.contains("превышено время ожидания") {
//!         return Some(PowerShellErrorKind::Timeout);
//!     }
//!     None
//! }
//! ```
//!
//! ## Note on OS-Level Errors
//!
//! For I/O errors (file access, permissions), the library automatically uses
//! `os_error` and `io_error` fields which are language-independent. These are
//! checked first by methods like `is_access_denied()`, so localized I/O errors
//! are handled automatically without custom classifiers.

use crate::script_strategy::PSScript;
use std::path::PathBuf;
use std::time::Duration;
use thiserror::Error;

/// PowerShell error classification
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PowerShellErrorKind {
    /// Script execution failed
    ExecutionFailed,
    /// Operation timed out
    Timeout,
    /// Access denied error
    AccessDenied,
    /// Execution policy error
    ExecutionPolicy,
    /// Cmdlet not found
    CmdletNotFound,
}

/// Detailed PowerShell script execution error with diagnostic information
#[derive(Debug, Clone)]
pub struct PowerShellError {
    /// Error classification
    pub kind: PowerShellErrorKind,

    /// Type of script that failed
    pub script_type: PSScript,

    /// Exit code from PowerShell process
    pub exit_code: Option<i32>,

    /// Standard output (may contain diagnostic info)
    pub stdout: String,

    /// Standard error output
    pub stderr: String,

    /// Path to the script that was executed
    pub script_path: PathBuf,

    /// Parameters passed to the script
    pub parameters: Option<String>,

    /// Execution duration (for performance diagnosis)
    pub duration: Option<Duration>,

    /// IO error kind (if available)
    pub io_error: Option<std::io::ErrorKind>,

    /// OS error code (if available)
    pub os_error: Option<i32>,
}

impl PowerShellError {
    /// Infers error kind from stderr content (English only)
    ///
    /// For localized error messages, use `classify_with()` with a custom classifier
    /// or use `reclassify_with()` to reclassify an existing error.
    pub fn infer_kind_from_stderr(stderr: &str) -> PowerShellErrorKind {
        let stderr_lower = stderr.to_lowercase();

        // Check for access denied (highest priority for common errors)
        if (stderr_lower.contains("access") && stderr_lower.contains("denied"))
            || stderr_lower.contains("unauthorizedaccessexception") {
            return PowerShellErrorKind::AccessDenied;
        }

        // Check for execution policy
        if stderr_lower.contains("execution policy")
            || stderr_lower.contains("executionpolicy") {
            return PowerShellErrorKind::ExecutionPolicy;
        }

        // Check for cmdlet not found
        if stderr_lower.contains("not recognized as")
            || stderr_lower.contains("commandnotfoundexception") {
            return PowerShellErrorKind::CmdletNotFound;
        }

        // Check for timeout (match both "timeout" and "timed out")
        if stderr_lower.contains("timeout") || stderr_lower.contains("timed out") {
            return PowerShellErrorKind::Timeout;
        }

        // Default to execution failed
        PowerShellErrorKind::ExecutionFailed
    }

    /// Classifies error with optional custom classifier
    ///
    /// Tries the custom classifier first, then falls back to built-in English classifier.
    ///
    /// # Example
    /// ```rust
    /// use wincent::error::{PowerShellError, PowerShellErrorKind};
    ///
    /// // Chinese error classifier
    /// let classifier = |stderr: &str| {
    ///     let stderr_lower = stderr.to_lowercase();
    ///     if stderr_lower.contains("拒绝访问") {
    ///         return Some(PowerShellErrorKind::AccessDenied);
    ///     }
    ///     if stderr_lower.contains("执行策略") {
    ///         return Some(PowerShellErrorKind::ExecutionPolicy);
    ///     }
    ///     None
    /// };
    ///
    /// let kind = PowerShellError::classify_with("拒绝访问", Some(&classifier));
    /// assert_eq!(kind, PowerShellErrorKind::AccessDenied);
    /// ```
    pub fn classify_with(
        stderr: &str,
        custom_classifier: Option<&dyn Fn(&str) -> Option<PowerShellErrorKind>>,
    ) -> PowerShellErrorKind {
        // Try custom classifier first
        if let Some(classifier) = custom_classifier {
            if let Some(kind) = classifier(stderr) {
                return kind;
            }
        }

        // Fallback to built-in English classifier
        Self::infer_kind_from_stderr(stderr)
    }

    /// Returns the raw stderr for custom analysis
    ///
    /// # Example
    /// ```rust
    /// # use wincent::error::PowerShellError;
    /// # let err = PowerShellError {
    /// #     kind: wincent::error::PowerShellErrorKind::ExecutionFailed,
    /// #     script_type: wincent::script_strategy::PSScript::QueryQuickAccess,
    /// #     exit_code: Some(1),
    /// #     stdout: String::new(),
    /// #     stderr: "拒绝访问".to_string(),
    /// #     script_path: std::path::PathBuf::from("test.ps1"),
    /// #     parameters: None,
    /// #     duration: None,
    /// #     io_error: None,
    /// #     os_error: None,
    /// # };
    /// let stderr = err.raw_stderr();
    /// if stderr.contains("拒绝访问") {
    ///     println!("Chinese access denied error");
    /// }
    /// ```
    pub fn raw_stderr(&self) -> &str {
        &self.stderr
    }

    /// Returns the raw stdout for custom analysis
    pub fn raw_stdout(&self) -> &str {
        &self.stdout
    }

    /// Checks if stderr contains a specific pattern (case-insensitive)
    ///
    /// # Example
    /// ```rust
    /// # use wincent::error::PowerShellError;
    /// # let err = PowerShellError {
    /// #     kind: wincent::error::PowerShellErrorKind::ExecutionFailed,
    /// #     script_type: wincent::script_strategy::PSScript::QueryQuickAccess,
    /// #     exit_code: Some(1),
    /// #     stdout: String::new(),
    /// #     stderr: "拒绝访问".to_string(),
    /// #     script_path: std::path::PathBuf::from("test.ps1"),
    /// #     parameters: None,
    /// #     duration: None,
    /// #     io_error: None,
    /// #     os_error: None,
    /// # };
    /// if err.stderr_contains("拒绝访问") {
    ///     println!("Access denied in Chinese");
    /// }
    /// ```
    pub fn stderr_contains(&self, pattern: &str) -> bool {
        self.stderr.to_lowercase().contains(&pattern.to_lowercase())
    }

    /// Creates a new PowerShellError with a specific kind
    ///
    /// Useful for reclassifying errors based on custom logic.
    ///
    /// # Important: State Consistency
    ///
    /// This method will NOT change the kind if the error has strong evidence
    /// from OS-level or I/O-level error codes (`os_error` or `io_error`).
    /// This prevents state inconsistency where `is_access_denied()` and
    /// `is_timeout()` could both return true for the same error.
    ///
    /// If you need to override this protection, you must manually create
    /// a new `PowerShellError` struct.
    ///
    /// # Example
    /// ```rust
    /// # use wincent::error::{PowerShellError, PowerShellErrorKind};
    /// # let err = PowerShellError {
    /// #     kind: PowerShellErrorKind::ExecutionFailed,
    /// #     script_type: wincent::script_strategy::PSScript::QueryQuickAccess,
    /// #     exit_code: Some(1),
    /// #     stdout: String::new(),
    /// #     stderr: "拒绝访问".to_string(),
    /// #     script_path: std::path::PathBuf::from("test.ps1"),
    /// #     parameters: None,
    /// #     duration: None,
    /// #     io_error: None,
    /// #     os_error: None,
    /// # };
    /// let err = err.with_kind(PowerShellErrorKind::AccessDenied);
    /// assert!(err.is_access_denied());
    /// ```
    ///
    /// # Example: Protection Against Inconsistency
    /// ```rust
    /// # use wincent::error::{PowerShellError, PowerShellErrorKind};
    /// # let mut err = PowerShellError {
    /// #     kind: PowerShellErrorKind::AccessDenied,
    /// #     script_type: wincent::script_strategy::PSScript::QueryQuickAccess,
    /// #     exit_code: Some(1),
    /// #     stdout: String::new(),
    /// #     stderr: "Access denied".to_string(),
    /// #     script_path: std::path::PathBuf::from("test.ps1"),
    /// #     parameters: None,
    /// #     duration: None,
    /// #     io_error: None,
    /// #     os_error: Some(5), // Windows ERROR_ACCESS_DENIED
    /// # };
    /// // This will NOT change the kind because os_error is present
    /// let err = err.with_kind(PowerShellErrorKind::Timeout);
    /// assert_eq!(err.kind, PowerShellErrorKind::AccessDenied); // Unchanged
    /// assert!(err.is_access_denied()); // Still true
    /// assert!(!err.is_timeout()); // False - no inconsistency
    /// ```
    pub fn with_kind(mut self, kind: PowerShellErrorKind) -> Self {
        // Only allow reclassification if there's no strong evidence from OS/IO errors
        if self.os_error.is_none() && self.io_error.is_none() {
            self.kind = kind;
        }
        // If os_error or io_error exists, silently ignore the reclassification
        // to maintain state consistency
        self
    }

    /// Reclassifies the error using a custom classifier
    ///
    /// # Important: State Consistency
    ///
    /// This method will NOT change the kind if the error has strong evidence
    /// from OS-level or I/O-level error codes (`os_error` or `io_error`).
    /// This prevents state inconsistency where multiple `is_*()` methods
    /// could return true for the same error.
    ///
    /// The classifier will still be called, but its result will be ignored
    /// if `os_error` or `io_error` is present.
    ///
    /// # Example
    /// ```rust
    /// # use wincent::error::{PowerShellError, PowerShellErrorKind};
    /// # let err = PowerShellError {
    /// #     kind: PowerShellErrorKind::ExecutionFailed,
    /// #     script_type: wincent::script_strategy::PSScript::QueryQuickAccess,
    /// #     exit_code: Some(1),
    /// #     stdout: String::new(),
    /// #     stderr: "拒绝访问".to_string(),
    /// #     script_path: std::path::PathBuf::from("test.ps1"),
    /// #     parameters: None,
    /// #     duration: None,
    /// #     io_error: None,
    /// #     os_error: None,
    /// # };
    /// let classifier = |stderr: &str| {
    ///     if stderr.contains("拒绝访问") {
    ///         Some(PowerShellErrorKind::AccessDenied)
    ///     } else {
    ///         None
    ///     }
    /// };
    ///
    /// let err = err.reclassify_with(classifier);
    /// assert!(err.is_access_denied());
    /// ```
    ///
    /// # Example: Protection Against Inconsistency
    /// ```rust
    /// # use wincent::error::{PowerShellError, PowerShellErrorKind};
    /// # let mut err = PowerShellError {
    /// #     kind: PowerShellErrorKind::AccessDenied,
    /// #     script_type: wincent::script_strategy::PSScript::QueryQuickAccess,
    /// #     exit_code: Some(1),
    /// #     stdout: String::new(),
    /// #     stderr: "timeout error".to_string(),
    /// #     script_path: std::path::PathBuf::from("test.ps1"),
    /// #     parameters: None,
    /// #     duration: None,
    /// #     io_error: Some(std::io::ErrorKind::PermissionDenied),
    /// #     os_error: None,
    /// # };
    /// let classifier = |_stderr: &str| Some(PowerShellErrorKind::Timeout);
    ///
    /// // Reclassification is ignored because io_error is present
    /// let err = err.reclassify_with(classifier);
    /// assert_eq!(err.kind, PowerShellErrorKind::AccessDenied); // Unchanged
    /// assert!(err.is_access_denied()); // Still true
    /// assert!(!err.is_timeout()); // False - no inconsistency
    /// ```
    pub fn reclassify_with<F>(mut self, classifier: F) -> Self
    where
        F: Fn(&str) -> Option<PowerShellErrorKind>,
    {
        // Always call the classifier to allow side effects (logging, metrics, etc.)
        let classification_result = classifier(&self.stderr);

        // Only apply the result if there's no strong evidence from OS/IO errors
        if self.os_error.is_none() && self.io_error.is_none() {
            if let Some(new_kind) = classification_result {
                self.kind = new_kind;
            }
        }
        // If os_error or io_error exists, ignore the classification result
        // to maintain state consistency
        self
    }

    /// Normalizes stderr for case-insensitive matching
    fn normalized_stderr(&self) -> String {
        self.stderr.to_lowercase().trim().to_string()
    }

    /// Checks if error is due to access denied
    pub fn is_access_denied(&self) -> bool {
        // Priority 1: Check OS error code
        if let Some(os_err) = self.os_error {
            // Windows ERROR_ACCESS_DENIED = 5
            if os_err == 5 {
                return true;
            }
        }

        // Priority 2: Check IO error kind
        if let Some(io_kind) = self.io_error {
            if io_kind == std::io::ErrorKind::PermissionDenied {
                return true;
            }
        }

        // Priority 3: Check error kind
        if self.kind == PowerShellErrorKind::AccessDenied {
            return true;
        }

        // Fallback: Text matching (case-insensitive, English only)
        // Note: Localized errors should be caught by os_error/io_error checks above
        let stderr_lower = self.normalized_stderr();
        stderr_lower.contains("access") && stderr_lower.contains("denied")
            || stderr_lower.contains("unauthorizedaccessexception")
    }

    /// Checks if error is due to execution policy
    pub fn is_execution_policy_error(&self) -> bool {
        // If os_error or io_error exists, they take precedence
        // This prevents state inconsistency where multiple is_*() methods return true
        if self.os_error.is_some() || self.io_error.is_some() {
            return false;
        }

        // Priority 1: Check error kind
        if self.kind == PowerShellErrorKind::ExecutionPolicy {
            return true;
        }

        // Fallback: Text matching (case-insensitive)
        let stderr_lower = self.normalized_stderr();
        stderr_lower.contains("execution policy")
            || stderr_lower.contains("executionpolicy")
    }

    /// Checks if error is due to timeout
    pub fn is_timeout(&self) -> bool {
        // If os_error or io_error exists, they take precedence
        // This prevents state inconsistency where multiple is_*() methods return true
        if self.os_error.is_some() || self.io_error.is_some() {
            return false;
        }

        // Priority 1: Check error kind
        if self.kind == PowerShellErrorKind::Timeout {
            return true;
        }

        // Fallback: Text matching (case-insensitive, English only)
        // Note: Timeout errors should be caught by kind check above
        let stderr_lower = self.normalized_stderr();
        stderr_lower.contains("timeout")
    }

    /// Checks if error is due to missing cmdlet
    pub fn is_cmdlet_not_found(&self) -> bool {
        // If os_error or io_error exists, they take precedence
        // This prevents state inconsistency where multiple is_*() methods return true
        if self.os_error.is_some() || self.io_error.is_some() {
            return false;
        }

        // Priority 1: Check error kind
        if self.kind == PowerShellErrorKind::CmdletNotFound {
            return true;
        }

        // Fallback: Text matching (case-insensitive)
        let stderr_lower = self.normalized_stderr();
        stderr_lower.contains("not recognized as")
            || stderr_lower.contains("commandnotfoundexception")
    }

    /// Provides user-friendly fix suggestions
    pub fn suggest_fix(&self) -> Option<String> {
        if self.is_access_denied() {
            return Some("Try running as administrator or check file permissions.".to_string());
        }

        if self.is_execution_policy_error() {
            return Some(
                "PowerShell execution policy is blocking scripts. \
                 Run: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser"
                    .to_string(),
            );
        }

        if self.is_cmdlet_not_found() {
            return Some(
                "Required PowerShell cmdlet is not available. Check PowerShell version."
                    .to_string(),
            );
        }

        None
    }

    /// Categorizes error as transient (retryable) or permanent
    pub fn is_transient(&self) -> bool {
        // Timeout, network issues, temporary locks are transient
        self.is_timeout()
            || self.normalized_stderr().contains("locked")
            || self.normalized_stderr().contains("in use")
    }
}

impl std::fmt::Display for PowerShellError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "PowerShell script failed: {:?}\n\
             Exit code: {}\n\
             Script: {}\n\
             Parameters: {}\n\
             Error: {}",
            self.script_type,
            self.exit_code
                .map(|c| c.to_string())
                .unwrap_or_else(|| "unknown".to_string()),
            self.script_path.display(),
            self.parameters.as_deref().unwrap_or("none"),
            self.stderr
        )?;

        if let Some(suggestion) = self.suggest_fix() {
            write!(f, "\n\nSuggestion: {}", suggestion)?;
        }

        Ok(())
    }
}

#[derive(Error, Debug)]
pub enum WincentError {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("UTF-8 conversion error: {0}")]
    Utf8(#[from] std::string::FromUtf8Error),

    #[error("PowerShell execution failed: {0}")]
    PowerShellExecution(PowerShellError),

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
    MissingParameter,

    #[error("Windows API error: {0}")]
    WindowsApi(i32),

    #[error("Script strategy not found: {0}")]
    ScriptStrategyNotFound(String),

    #[error("Async execution error: {0}")]
    AsyncExecution(String),

    #[error("Operation timed out: {0}")]
    Timeout(String),

    #[error(
        "Quick Access clear partially succeeded (recent_files_cleared: {recent_files_cleared}, frequent_folders_cleared: {frequent_folders_cleared}): {source}"
    )]
    PartialEmpty {
        recent_files_cleared: bool,
        frequent_folders_cleared: bool,
        #[source]
        source: Box<WincentError>,
    },

    #[error("Item already exists in Quick Access: {0}")]
    AlreadyExists(String),

    #[error("Item not found in Quick Access: {0}")]
    NotInRecent(String),
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

        let missing_param = WincentError::MissingParameter;
        assert!(format!("{}", missing_param).contains("Missing function parameter"));

        let invalid_path = WincentError::InvalidPath("test/path".to_string());
        assert!(format!("{}", invalid_path).contains("test/path"));

        let ps_error = WincentError::PowerShellExecution(PowerShellError {
            kind: PowerShellErrorKind::ExecutionFailed,
            script_type: PSScript::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: "access denied".to_string(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        });
        assert!(format!("{}", ps_error).contains("access denied"));
    }

    #[test]
    fn test_result_type() {
        let success: WincentResult<()> = Ok(());
        assert!(success.is_ok());

        let failure: WincentResult<()> = Err(WincentError::MissingParameter);
        assert!(failure.is_err());

        let partial_empty = WincentError::PartialEmpty {
            recent_files_cleared: true,
            frequent_folders_cleared: false,
            source: Box::new(WincentError::ScriptFailed("failed to clear pinned folders".into())),
        };
        let rendered = format!("{}", partial_empty);
        assert!(rendered.contains("recent_files_cleared: true"));
        assert!(rendered.contains("frequent_folders_cleared: false"));
        assert!(rendered.contains("failed to clear pinned folders"));
    }

    #[test]
    fn test_powershell_error_is_access_denied() {
        let err = PowerShellError {
            kind: PowerShellErrorKind::ExecutionFailed,
            script_type: PSScript::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: "Access to the path is denied.".to_string(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        };
        assert!(err.is_access_denied());

        let err2 = PowerShellError {
            kind: PowerShellErrorKind::ExecutionFailed,
            script_type: PSScript::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: "UnauthorizedAccessException: Access denied".to_string(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        };
        assert!(err2.is_access_denied());
    }

    #[test]
    fn test_powershell_error_is_execution_policy() {
        let err = PowerShellError {
            kind: PowerShellErrorKind::ExecutionFailed,
            script_type: PSScript::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: "execution policy does not allow this script".to_string(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        };
        assert!(err.is_execution_policy_error());

        let err2 = PowerShellError {
            kind: PowerShellErrorKind::ExecutionFailed,
            script_type: PSScript::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: "Set-ExecutionPolicy required".to_string(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        };
        assert!(err2.is_execution_policy_error());
    }

    #[test]
    fn test_powershell_error_is_cmdlet_not_found() {
        let err = PowerShellError {
            kind: PowerShellErrorKind::ExecutionFailed,
            script_type: PSScript::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: "Get-Item is not recognized as a cmdlet".to_string(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        };
        assert!(err.is_cmdlet_not_found());

        let err2 = PowerShellError {
            kind: PowerShellErrorKind::ExecutionFailed,
            script_type: PSScript::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: "CommandNotFoundException: The term 'Get-Item' is not recognized".to_string(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        };
        assert!(err2.is_cmdlet_not_found());
    }

    #[test]
    fn test_powershell_error_suggest_fix() {
        let err = PowerShellError {
            kind: PowerShellErrorKind::ExecutionFailed,
            script_type: PSScript::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: "UnauthorizedAccessException".to_string(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        };
        let suggestion = err.suggest_fix();
        assert!(suggestion.is_some());
        assert!(suggestion.unwrap().contains("administrator"));

        let err2 = PowerShellError {
            kind: PowerShellErrorKind::ExecutionFailed,
            script_type: PSScript::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: "execution policy blocks this script".to_string(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        };
        let suggestion2 = err2.suggest_fix();
        assert!(suggestion2.is_some());
        assert!(suggestion2.unwrap().contains("Set-ExecutionPolicy"));

        let err3 = PowerShellError {
            kind: PowerShellErrorKind::ExecutionFailed,
            script_type: PSScript::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: "CommandNotFoundException".to_string(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        };
        let suggestion3 = err3.suggest_fix();
        assert!(suggestion3.is_some());
        assert!(suggestion3.unwrap().contains("PowerShell version"));
    }

    #[test]
    fn test_powershell_error_is_transient() {
        let err = PowerShellError {
            kind: PowerShellErrorKind::ExecutionFailed,
            script_type: PSScript::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: "The file is locked by another process".to_string(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        };
        assert!(err.is_transient());

        let err2 = PowerShellError {
            kind: PowerShellErrorKind::ExecutionFailed,
            script_type: PSScript::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: "The file is in use".to_string(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        };
        assert!(err2.is_transient());

        let err3 = PowerShellError {
            kind: PowerShellErrorKind::ExecutionFailed,
            script_type: PSScript::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: "Access denied".to_string(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        };
        assert!(!err3.is_transient());
    }

    #[test]
    fn test_powershell_error_display() {
        let err = PowerShellError {
            kind: PowerShellErrorKind::AccessDenied,
            script_type: PSScript::PinToFrequentFolder,
            exit_code: Some(1),
            stdout: "Some output".to_string(),
            stderr: "Access denied".to_string(),
            script_path: PathBuf::from("C:\\scripts\\test.ps1"),
            parameters: Some("C:\\test\\folder".to_string()),
            duration: Some(Duration::from_millis(150)),
            io_error: None,
            os_error: Some(5),
        };

        let display = format!("{}", err);
        assert!(display.contains("PinToFrequentFolder"));
        assert!(display.contains("Exit code: 1"));
        assert!(display.contains("C:\\scripts\\test.ps1"));
        assert!(display.contains("C:\\test\\folder"));
        assert!(display.contains("Access denied"));
    }
}
