//! Error handling and error type definitions
//!
//! Provides centralized error handling for Windows Quick Access operations,
//! with unified error types and automatic error conversions.
//!
//! # Key Features
//! - Cross-domain error classification (I/O, system, scripting)
//! - Automatic error conversion from common Rust types
//! - Detailed error context preservation
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
//! ## Approach 1: Using `classification_with()`
//!
//! Classify errors after they occur:
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
//! # fn example() -> WincentResult<()> {
//! # let manager = QuickAccessManager::new();
//! match manager.add_item("C:\\test", QuickAccess::FrequentFolders, AddOptions::new()) {
//!     Ok(_) => println!("Success"),
//!     Err(WincentError::PowerShellExecution(err)) => {
//!         let kind = err.classification_with(classify_chinese_error);
//!         if kind == PowerShellErrorKind::AccessDenied {
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

use crate::QuickAccess;
use std::path::{Path, PathBuf};
use std::time::Duration;
use thiserror::Error;

type PowerShellClassifier = dyn Fn(&str) -> Option<PowerShellErrorKind>;

/// User-facing operation associated with a PowerShell failure.
///
/// This describes the Quick Access operation that failed without exposing the
/// internal script generation strategy used to implement it.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PowerShellOperation {
    /// Refresh open Explorer windows.
    RefreshExplorer,
    /// Query both Recent Files and Frequent Folders.
    QueryQuickAccess,
    /// Query Recent Files.
    QueryRecentFiles,
    /// Query Frequent Folders.
    QueryFrequentFolders,
    /// Add a file to Recent Files.
    AddRecentFile,
    /// Remove a file from Recent Files.
    RemoveRecentFile,
    /// Pin a folder to Frequent Folders.
    PinFrequentFolder,
    /// Unpin a folder from Frequent Folders.
    UnpinFrequentFolder,
    /// Unpin every pinned folder.
    EmptyPinnedFolders,
    /// Check whether query operations are available.
    CheckQueryFeasible,
    /// Check whether pin and unpin operations are available.
    CheckPinUnpinFeasible,
}

/// PowerShell error classification
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum PowerShellErrorKind {
    /// PowerShell process failed without a more specific classification.
    ProcessFailed,
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
    kind: PowerShellErrorKind,

    /// Quick Access operation that failed.
    operation: PowerShellOperation,

    /// Exit code from PowerShell process
    exit_code: Option<i32>,

    /// Standard output (may contain diagnostic info)
    stdout: String,

    /// Standard error output
    stderr: String,

    /// Path to the script that was executed
    script_path: PathBuf,

    /// Parameters passed to the script
    parameters: Option<String>,

    /// Execution duration (for performance diagnosis)
    duration: Option<Duration>,

    /// IO error kind (if available)
    io_error: Option<std::io::ErrorKind>,

    /// OS error code (if available)
    os_error: Option<i32>,
}

/// Crate-private builder for readable internal PowerShell error construction.
///
/// This keeps production fields private while avoiding long positional
/// constructors in tests and internal call sites.
#[derive(Debug, Clone)]
pub(crate) struct PowerShellErrorBuilder {
    kind: PowerShellErrorKind,
    operation: PowerShellOperation,
    exit_code: Option<i32>,
    stdout: String,
    stderr: String,
    script_path: PathBuf,
    parameters: Option<String>,
    duration: Option<Duration>,
    io_error: Option<std::io::ErrorKind>,
    os_error: Option<i32>,
}

impl Default for PowerShellErrorBuilder {
    fn default() -> Self {
        Self {
            kind: PowerShellErrorKind::ProcessFailed,
            operation: PowerShellOperation::QueryQuickAccess,
            exit_code: Some(1),
            stdout: String::new(),
            stderr: String::new(),
            script_path: PathBuf::from("test.ps1"),
            parameters: None,
            duration: None,
            io_error: None,
            os_error: None,
        }
    }
}

impl PowerShellErrorBuilder {
    pub(crate) fn new(operation: PowerShellOperation) -> Self {
        Self {
            operation,
            ..Self::default()
        }
    }

    pub(crate) fn kind(mut self, kind: PowerShellErrorKind) -> Self {
        self.kind = kind;
        self
    }

    pub(crate) fn exit_code(mut self, exit_code: Option<i32>) -> Self {
        self.exit_code = exit_code;
        self
    }

    pub(crate) fn stdout(mut self, stdout: impl Into<String>) -> Self {
        self.stdout = stdout.into();
        self
    }

    pub(crate) fn stderr(mut self, stderr: impl Into<String>) -> Self {
        self.stderr = stderr.into();
        self
    }

    pub(crate) fn script_path(mut self, script_path: impl Into<PathBuf>) -> Self {
        self.script_path = script_path.into();
        self
    }

    pub(crate) fn parameters(mut self, parameters: impl Into<String>) -> Self {
        self.parameters = Some(parameters.into());
        self
    }

    pub(crate) fn duration(mut self, duration: Duration) -> Self {
        self.duration = Some(duration);
        self
    }

    pub(crate) fn io_error(mut self, io_error: std::io::ErrorKind) -> Self {
        self.io_error = Some(io_error);
        self
    }

    pub(crate) fn os_error(mut self, os_error: i32) -> Self {
        self.os_error = Some(os_error);
        self
    }

    pub(crate) fn build(self) -> PowerShellError {
        PowerShellError::new(self)
    }
}

/// Structured invalid path details.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct InvalidPathError {
    path: Option<PathBuf>,
    reason: String,
}

impl InvalidPathError {
    /// Creates an invalid path error with a path and reason.
    #[must_use]
    pub fn new<P: Into<PathBuf>, S: Into<String>>(path: P, reason: S) -> Self {
        Self {
            path: Some(path.into()),
            reason: reason.into(),
        }
    }

    /// Creates an invalid path error with only a reason.
    #[must_use]
    pub fn reason<S: Into<String>>(reason: S) -> Self {
        Self {
            path: None,
            reason: reason.into(),
        }
    }

    /// Returns the invalid path, if one was recorded.
    #[must_use]
    pub fn path(&self) -> Option<&Path> {
        self.path.as_deref()
    }

    /// Returns the reason the path is invalid.
    #[must_use]
    pub fn reason_text(&self) -> &str {
        &self.reason
    }
}

impl std::fmt::Display for InvalidPathError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match &self.path {
            Some(path) => write!(f, "{}: {}", self.reason, path.display()),
            None => f.write_str(&self.reason),
        }
    }
}

impl PowerShellError {
    pub(crate) fn builder(operation: PowerShellOperation) -> PowerShellErrorBuilder {
        PowerShellErrorBuilder::new(operation)
    }

    pub(crate) fn new(builder: PowerShellErrorBuilder) -> Self {
        Self {
            kind: builder.kind,
            operation: builder.operation,
            exit_code: builder.exit_code,
            stdout: builder.stdout,
            stderr: builder.stderr,
            script_path: builder.script_path,
            parameters: builder.parameters,
            duration: builder.duration,
            io_error: builder.io_error,
            os_error: builder.os_error,
        }
    }

    /// Returns the error classification.
    #[must_use]
    pub fn kind(&self) -> PowerShellErrorKind {
        self.kind.clone()
    }

    /// Returns the Quick Access operation that failed.
    #[must_use]
    pub fn operation(&self) -> PowerShellOperation {
        self.operation
    }

    /// Returns the PowerShell process exit code, if one was reported.
    #[must_use]
    pub fn exit_code(&self) -> Option<i32> {
        self.exit_code
    }

    /// Returns the script path that was executed.
    #[must_use]
    pub fn script_path(&self) -> &std::path::Path {
        &self.script_path
    }

    /// Returns the parameters passed to the script, if recorded.
    #[must_use]
    pub fn parameters(&self) -> Option<&str> {
        self.parameters.as_deref()
    }

    /// Returns the recorded execution duration, if available.
    #[must_use]
    pub fn duration(&self) -> Option<Duration> {
        self.duration
    }

    /// Returns the underlying I/O error kind, if available.
    #[must_use]
    pub fn io_error(&self) -> Option<std::io::ErrorKind> {
        self.io_error
    }

    /// Returns the underlying OS error code, if available.
    #[must_use]
    pub fn os_error(&self) -> Option<i32> {
        self.os_error
    }

    /// Infers error kind from stderr content (English only)
    ///
    /// For localized error messages, use `classify_with()` with a custom classifier
    /// or use `classification_with()` to classify an existing error.
    pub fn infer_kind_from_stderr(stderr: &str) -> PowerShellErrorKind {
        let stderr_lower = stderr.to_lowercase();

        // Check for access denied (highest priority for common errors)
        if (stderr_lower.contains("access") && stderr_lower.contains("denied"))
            || stderr_lower.contains("unauthorizedaccessexception")
        {
            return PowerShellErrorKind::AccessDenied;
        }

        // Check for execution policy
        if stderr_lower.contains("execution policy") || stderr_lower.contains("executionpolicy") {
            return PowerShellErrorKind::ExecutionPolicy;
        }

        // Check for cmdlet not found
        if stderr_lower.contains("not recognized as")
            || stderr_lower.contains("commandnotfoundexception")
        {
            return PowerShellErrorKind::CmdletNotFound;
        }

        // Check for timeout (match both "timeout" and "timed out")
        if stderr_lower.contains("timeout") || stderr_lower.contains("timed out") {
            return PowerShellErrorKind::Timeout;
        }

        // Default to a generic process failure.
        PowerShellErrorKind::ProcessFailed
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
        custom_classifier: Option<&PowerShellClassifier>,
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
    /// ```rust,no_run
    /// # use wincent::prelude::*;
    /// # use wincent::error::WincentError;
    /// # fn example() -> Result<(), WincentError> {
    /// let manager = QuickAccessManager::new();
    /// let Err(WincentError::PowerShellExecution(err)) =
    ///     manager.get_items(QuickAccess::FrequentFolders)
    /// else {
    ///     return Ok(());
    /// };
    ///
    /// let stderr = err.raw_stderr();
    /// if stderr.contains("拒绝访问") {
    ///     println!("Chinese access denied error");
    /// }
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
    pub fn raw_stderr(&self) -> &str {
        &self.stderr
    }

    /// Returns the raw stdout for custom analysis
    #[must_use]
    pub fn raw_stdout(&self) -> &str {
        &self.stdout
    }

    /// Checks if stderr contains a specific pattern (case-insensitive)
    ///
    /// # Example
    /// ```rust,no_run
    /// # use wincent::prelude::*;
    /// # use wincent::error::WincentError;
    /// # fn example() -> Result<(), WincentError> {
    /// let manager = QuickAccessManager::new();
    /// let Err(WincentError::PowerShellExecution(err)) =
    ///     manager.get_items(QuickAccess::FrequentFolders)
    /// else {
    ///     return Ok(());
    /// };
    ///
    /// if err.stderr_contains("拒绝访问") {
    ///     println!("Access denied in Chinese");
    /// }
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
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
    /// If you need to override this protection, run the operation again and
    /// classify the new error from its raw context.
    ///
    /// # Example
    /// ```rust,no_run
    /// # use wincent::prelude::*;
    /// # use wincent::error::{PowerShellErrorKind, WincentError};
    /// # fn example() -> Result<(), WincentError> {
    /// let manager = QuickAccessManager::new();
    /// let Err(WincentError::PowerShellExecution(err)) =
    ///     manager.get_items(QuickAccess::FrequentFolders)
    /// else {
    ///     return Ok(());
    /// };
    ///
    /// let err = err.with_kind(PowerShellErrorKind::AccessDenied);
    /// if err.is_access_denied() {
    ///     println!("access denied");
    /// }
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Example: Protection Against Inconsistency
    /// ```rust,no_run
    /// # use wincent::prelude::*;
    /// # use wincent::error::{PowerShellErrorKind, WincentError};
    /// # fn example() -> Result<(), WincentError> {
    /// let manager = QuickAccessManager::new();
    /// let Err(WincentError::PowerShellExecution(err)) =
    ///     manager.get_items(QuickAccess::FrequentFolders)
    /// else {
    ///     return Ok(());
    /// };
    /// let original_kind = err.kind();
    ///
    /// // This will NOT change the kind because os_error is present
    /// let err = err.with_kind(PowerShellErrorKind::Timeout);
    /// assert_eq!(err.kind(), original_kind);
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
    #[deprecated(
        since = "0.1.2",
        note = "use classification_with to classify without mutating the error"
    )]
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
    /// ```rust,no_run
    /// # use wincent::prelude::*;
    /// # use wincent::error::{PowerShellErrorKind, WincentError};
    /// # fn example() -> Result<(), WincentError> {
    /// let manager = QuickAccessManager::new();
    /// let Err(WincentError::PowerShellExecution(err)) =
    ///     manager.get_items(QuickAccess::FrequentFolders)
    /// else {
    ///     return Ok(());
    /// };
    ///
    /// let classifier = |stderr: &str| {
    ///     if stderr.contains("拒绝访问") {
    ///         Some(PowerShellErrorKind::AccessDenied)
    ///     } else {
    ///         None
    ///     }
    /// };
    ///
    /// let err = err.reclassify_with(classifier);
    /// if err.is_access_denied() {
    ///     println!("access denied");
    /// }
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Example: Protection Against Inconsistency
    /// ```rust,no_run
    /// # use wincent::prelude::*;
    /// # use wincent::error::{PowerShellErrorKind, WincentError};
    /// # fn example() -> Result<(), WincentError> {
    /// let manager = QuickAccessManager::new();
    /// let Err(WincentError::PowerShellExecution(err)) =
    ///     manager.get_items(QuickAccess::FrequentFolders)
    /// else {
    ///     return Ok(());
    /// };
    /// let original_kind = err.kind();
    ///
    /// let classifier = |_stderr: &str| Some(PowerShellErrorKind::Timeout);
    ///
    /// // Reclassification is ignored because io_error is present
    /// let err = err.reclassify_with(classifier);
    /// assert_eq!(err.kind(), original_kind);
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
    #[deprecated(
        since = "0.1.2",
        note = "use classification_with to classify without mutating the error"
    )]
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

    /// Classifies this error using a custom classifier without modifying it.
    #[must_use]
    pub fn classification_with<F>(&self, classifier: F) -> PowerShellErrorKind
    where
        F: Fn(&str) -> Option<PowerShellErrorKind>,
    {
        if self.os_error.is_some() || self.io_error.is_some() {
            return self.kind();
        }

        classifier(&self.stderr).unwrap_or_else(|| self.kind())
    }

    /// Normalizes stderr for case-insensitive matching
    fn normalized_stderr(&self) -> String {
        self.stderr.to_lowercase().trim().to_string()
    }

    /// Checks if error is due to access denied
    #[must_use]
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
    #[must_use]
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
        stderr_lower.contains("execution policy") || stderr_lower.contains("executionpolicy")
    }

    /// Checks if error is due to timeout
    #[must_use]
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
    #[must_use]
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
    #[must_use]
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
    #[must_use]
    pub fn is_transient(&self) -> bool {
        // Timeout, network issues, temporary locks are transient
        self.is_timeout()
            || self.normalized_stderr().contains("locked")
            || self.normalized_stderr().contains("in use")
            || self.normalized_stderr().contains("temporarily unavailable")
    }
}

impl std::fmt::Display for PowerShellError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "PowerShell {:?} failed\n\
             Exit code: {}\n\
             Script: {}\n\
             Parameters: {}\n\
             Error: {}",
            self.operation,
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

/// Post-mutation step that failed after the requested Quick Access mutation succeeded.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum QuickAccessPostMutationStep {
    /// Failed to delete Explorer's Recent Files backing data after adding a recent file.
    DeleteRecentFilesBackingData,
    /// Failed to refresh open Explorer windows after a successful mutation.
    RefreshExplorer,
}

/// Error type returned by wincent operations.
#[derive(Error, Debug)]
#[non_exhaustive]
pub enum WincentError {
    /// Filesystem or process I/O failed.
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// UTF-8 decoding failed.
    #[error("UTF-8 conversion error: {0}")]
    Utf8(#[from] std::string::FromUtf8Error),

    /// A generated PowerShell script failed or could not be started.
    #[error("PowerShell execution failed: {0}")]
    PowerShellExecution(Box<PowerShellError>),

    /// A supplied path is empty, malformed, missing, or has the wrong type.
    #[error("Invalid path: {0}")]
    InvalidPath(InvalidPathError),

    /// A function argument is outside the supported range.
    #[error("Invalid argument: {0}")]
    InvalidArgument(String),

    /// The requested operation is not supported by this implementation.
    #[error("Operation not supported: {0}")]
    UnsupportedOperation(String),

    /// A non-I/O system operation failed.
    #[error("System error: {0}")]
    SystemError(String),

    /// Fixed-size array conversion failed while parsing binary data.
    #[error("Array conversion error: {0}")]
    ArrayConversion(#[from] std::array::TryFromSliceError),

    /// A generated script reported failure.
    #[error("Script failed error: {0}")]
    ScriptFailed(String),

    /// A numeric Quick Access category value is unknown.
    #[error("Unknown quick access type: {0}")]
    UnknownQuickAccessType(u32),

    /// A numeric script method value is unknown.
    #[error("Unknown script method: {0}")]
    UnknownScriptMethod(u32),

    /// A required script or API parameter was not provided.
    #[error("Missing function parameter")]
    MissingParameter,

    /// A Windows API call returned a failing HRESULT or status code.
    #[error("Windows API error: {0}")]
    WindowsApi(i32),

    /// No script generation strategy exists for the requested script.
    #[error("Script strategy not found: {0}")]
    ScriptStrategyNotFound(String),

    /// An operation exceeded its caller-provided timeout.
    #[error("Operation timed out: {0}")]
    Timeout(String),

    /// A clear operation completed some categories before a later category failed.
    #[error(
        "Quick Access clear partially succeeded (recent_files_cleared: {recent_files_cleared}, frequent_folders_cleared: {frequent_folders_cleared}): {source}"
    )]
    #[non_exhaustive]
    PartialEmpty {
        /// Whether Recent Files were cleared before a later cleanup step failed.
        recent_files_cleared: bool,
        /// Whether Frequent Folders cleanup made user-visible progress before failure.
        ///
        /// For Frequent Folders this is intentionally coarse-grained: `true`
        /// means the user-visited jump list was cleared. Pinned folder cleanup
        /// may still have failed and be reported by `source`.
        frequent_folders_cleared: bool,
        /// The underlying error that prevented the full clear from completing.
        #[source]
        source: Box<WincentError>,
    },

    /// A Quick Access mutation succeeded, but a post-mutation display update failed.
    ///
    /// This means the requested add or remove operation already completed. The
    /// error describes a follow-up step, such as deleting Explorer's Recent Files
    /// backing data or refreshing open Explorer windows. Callers should avoid
    /// blindly retrying the original mutation when they receive this error.
    #[error(
        "Quick Access post-mutation step {step:?} failed for {qa_type:?} item {path}: {source}"
    )]
    #[non_exhaustive]
    PostMutationFailure {
        /// Path whose mutation succeeded before the post-mutation step failed.
        path: String,
        /// Quick Access category that was mutated.
        qa_type: QuickAccess,
        /// Post-mutation step that failed.
        step: QuickAccessPostMutationStep,
        /// The underlying error from the failed post-mutation step.
        #[source]
        source: Box<WincentError>,
    },

    /// The item is already present in the requested Quick Access category.
    #[error("Item already exists in {qa_type:?}: {path}")]
    #[non_exhaustive]
    AlreadyExists {
        /// Path that was already present.
        path: String,
        /// Quick Access category where the item was found.
        qa_type: QuickAccess,
    },

    /// The item is not present in the requested Quick Access category.
    #[error("Item not found in {qa_type:?}: {path}")]
    #[non_exhaustive]
    NotInQuickAccess {
        /// Path that was not found.
        path: String,
        /// Quick Access category where the item was expected.
        qa_type: QuickAccess,
    },

    /// COM is already initialized on the current thread with an incompatible apartment model.
    #[error("COM apartment model mismatch: {0}")]
    ComApartmentMismatch(String),

    /// A DestList or Compound File Binary parse operation failed.
    #[error("DestList parse error: {0}")]
    DestListParse(String),

    /// The DestList file uses a version that this parser does not support.
    #[error("Unsupported DestList version: {0}")]
    DestListUnsupportedVersion(u32),
}

impl From<windows::core::Error> for WincentError {
    fn from(err: windows::core::Error) -> Self {
        WincentError::WindowsApi(err.code().0)
    }
}

impl WincentError {
    pub(crate) fn invalid_path_reason(reason: impl Into<String>) -> Self {
        Self::InvalidPath(InvalidPathError::reason(reason))
    }

    pub(crate) fn invalid_path(path: impl Into<PathBuf>, reason: impl Into<String>) -> Self {
        Self::InvalidPath(InvalidPathError::new(path, reason))
    }

    pub(crate) fn already_exists(path: impl Into<String>, qa_type: QuickAccess) -> Self {
        Self::AlreadyExists {
            path: path.into(),
            qa_type,
        }
    }

    pub(crate) fn not_in_quick_access(path: impl Into<String>, qa_type: QuickAccess) -> Self {
        Self::NotInQuickAccess {
            path: path.into(),
            qa_type,
        }
    }

    pub(crate) fn post_mutation_failure(
        path: impl Into<String>,
        qa_type: QuickAccess,
        step: QuickAccessPostMutationStep,
        source: WincentError,
    ) -> Self {
        Self::PostMutationFailure {
            path: path.into(),
            qa_type,
            step,
            source: Box::new(source),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::WincentResult;
    use std::io::{Error, ErrorKind};

    fn ps_error_with_stderr(stderr: impl Into<String>) -> PowerShellError {
        PowerShellError::builder(PowerShellOperation::QueryQuickAccess)
            .stderr(stderr)
            .build()
    }

    #[test]
    fn test_error_conversions() {
        let io_error = Error::new(ErrorKind::NotFound, "file not found");
        let wincent_error = WincentError::from(io_error);
        assert!(matches!(wincent_error, WincentError::Io(_)));

        let missing_param = WincentError::MissingParameter;
        assert!(format!("{}", missing_param).contains("Missing function parameter"));

        let invalid_path = WincentError::invalid_path("test/path", "test path");
        assert!(format!("{}", invalid_path).contains("test/path"));

        let ps_error =
            WincentError::PowerShellExecution(Box::new(ps_error_with_stderr("access denied")));
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
            source: Box::new(WincentError::ScriptFailed(
                "failed to clear pinned folders".into(),
            )),
        };
        let rendered = format!("{}", partial_empty);
        assert!(rendered.contains("recent_files_cleared: true"));
        assert!(rendered.contains("frequent_folders_cleared: false"));
        assert!(rendered.contains("failed to clear pinned folders"));
    }

    #[test]
    fn test_powershell_error_is_access_denied() {
        let err = ps_error_with_stderr("Access to the path is denied.");
        assert!(err.is_access_denied());

        let err2 = ps_error_with_stderr("UnauthorizedAccessException: Access denied");
        assert!(err2.is_access_denied());
    }

    #[test]
    fn test_powershell_error_is_execution_policy() {
        let err = ps_error_with_stderr("execution policy does not allow this script");
        assert!(err.is_execution_policy_error());

        let err2 = ps_error_with_stderr("Set-ExecutionPolicy required");
        assert!(err2.is_execution_policy_error());
    }

    #[test]
    fn test_powershell_error_is_cmdlet_not_found() {
        let err = ps_error_with_stderr("Get-Item is not recognized as a cmdlet");
        assert!(err.is_cmdlet_not_found());

        let err2 =
            ps_error_with_stderr("CommandNotFoundException: The term 'Get-Item' is not recognized");
        assert!(err2.is_cmdlet_not_found());
    }

    #[test]
    fn test_powershell_error_suggest_fix() {
        let err = ps_error_with_stderr("UnauthorizedAccessException");
        let suggestion = err.suggest_fix();
        assert!(suggestion.is_some());
        assert!(suggestion.unwrap().contains("administrator"));

        let err2 = ps_error_with_stderr("execution policy blocks this script");
        let suggestion2 = err2.suggest_fix();
        assert!(suggestion2.is_some());
        assert!(suggestion2.unwrap().contains("Set-ExecutionPolicy"));

        let err3 = ps_error_with_stderr("CommandNotFoundException");
        let suggestion3 = err3.suggest_fix();
        assert!(suggestion3.is_some());
        assert!(suggestion3.unwrap().contains("PowerShell version"));
    }

    #[test]
    fn test_powershell_error_is_transient() {
        let err = ps_error_with_stderr("The file is locked by another process");
        assert!(err.is_transient());

        let err2 = ps_error_with_stderr("The file is in use");
        assert!(err2.is_transient());

        let err2 = ps_error_with_stderr("Shell namespace is temporarily unavailable");
        assert!(err2.is_transient());

        let err3 = ps_error_with_stderr("Access denied");
        assert!(!err3.is_transient());
    }

    #[test]
    fn test_powershell_error_display() {
        let err = PowerShellError::builder(PowerShellOperation::PinFrequentFolder)
            .kind(PowerShellErrorKind::AccessDenied)
            .stdout("Some output")
            .stderr("Access denied")
            .script_path("C:\\scripts\\test.ps1")
            .parameters("C:\\test\\folder")
            .duration(Duration::from_millis(150))
            .os_error(5)
            .build();

        let display = format!("{}", err);
        assert!(display.contains("PinFrequentFolder"));
        assert!(display.contains("Exit code: 1"));
        assert!(display.contains("C:\\scripts\\test.ps1"));
        assert!(display.contains("C:\\test\\folder"));
        assert!(display.contains("Access denied"));
    }
}
