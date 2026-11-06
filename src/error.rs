//! Error types returned by Windows Quick Access operations.

use crate::QuickAccess;
use std::path::{Path, PathBuf};
use thiserror::Error;

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

    /// A numeric Quick Access category value is unknown.
    #[error("Unknown quick access type: {0}")]
    UnknownQuickAccessType(u32),

    /// A Windows API call returned a failing HRESULT or status code.
    #[error("Windows API error: {0}")]
    WindowsApi(i32),

    /// An operation exceeded its caller-provided wait timeout.
    ///
    /// For Shell and COM mutations this is an indeterminate outcome, not proof
    /// that the operation failed. The underlying operation may still complete
    /// after the caller receives this error.
    #[error("Operation timed out: {0}")]
    Timeout(String),

    /// A cooperating Quick Access lock denies mutation of this category.
    #[error("Quick Access category is locked: {qa_type:?}")]
    #[non_exhaustive]
    QuickAccessLocked {
        /// Category whose mutation was refused.
        qa_type: crate::quick_access_lock::QuickAccessLockTarget,
    },

    /// This manager already has an outstanding Quick Access lock.
    #[error("Quick Access manager already holds a lock")]
    ManagerAlreadyLocked,

    /// This manager has no outstanding Quick Access lock.
    #[error("Quick Access manager has no lock to release")]
    ManagerNotLocked,

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

    /// The backing file changed after a direct DestList mutation took its snapshot.
    #[error("DestList backing file changed during direct mutation: {0}")]
    DestListConcurrentModification(PathBuf),

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

        let invalid_path = WincentError::invalid_path("test/path", "test path");
        assert!(format!("{}", invalid_path).contains("test/path"));
    }

    #[test]
    fn test_result_type() {
        let success: WincentResult<()> = Ok(());
        assert!(success.is_ok());

        let failure: WincentResult<()> =
            Err(WincentError::InvalidArgument("missing value".to_string()));
        assert!(failure.is_err());

        let partial_empty = WincentError::PartialEmpty {
            recent_files_cleared: true,
            frequent_folders_cleared: false,
            source: Box::new(WincentError::SystemError(
                "failed to clear pinned folders".into(),
            )),
        };
        let rendered = format!("{}", partial_empty);
        assert!(rendered.contains("recent_files_cleared: true"));
        assert!(rendered.contains("frequent_folders_cleared: false"));
        assert!(rendered.contains("failed to clear pinned folders"));
    }
}
