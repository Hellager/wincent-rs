//! Windows path validation and comparison helpers.

use crate::{error::WincentError, WincentResult};
use std::path::Path;

/// Lightweight path normalization without I/O operations.
pub(crate) fn normalize_path_lightweight(path: &str) -> String {
    let mut result = path.to_lowercase().replace('/', "\\");
    if result.len() > 3 && result.ends_with('\\') {
        result.pop();
    }
    result
}

/// Normalizes a Windows path for comparison, using canonicalization when available.
pub(crate) fn normalize_path_for_comparison(path: &str) -> String {
    let normalized = Path::new(path).canonicalize().map_or_else(
        |_| path.to_string(),
        |value| value.to_string_lossy().into_owned(),
    );
    normalize_path_lightweight(&normalized)
}

/// Checks whether two Windows paths refer to the same location.
pub(crate) fn paths_equal(left: &str, right: &str) -> bool {
    let normalized_left = normalize_path_lightweight(left);
    let normalized_right = normalize_path_lightweight(right);
    normalized_left == normalized_right
        || normalize_path_for_comparison(left) == normalize_path_for_comparison(right)
}

#[derive(Debug, Copy, Clone, PartialEq, Eq)]
pub(crate) enum PathType {
    File,
    Directory,
}

/// Identifies an existing path as a regular file or directory.
pub(crate) fn detect_path_type(path: &str) -> WincentResult<PathType> {
    if path.is_empty() {
        return Err(WincentError::invalid_path_reason("Empty path provided"));
    }

    let metadata = std::fs::metadata(path).map_err(|error| {
        if error.kind() == std::io::ErrorKind::NotFound {
            WincentError::invalid_path(path, "Path does not exist")
        } else {
            WincentError::Io(error)
        }
    })?;
    if metadata.is_file() {
        Ok(PathType::File)
    } else if metadata.is_dir() {
        Ok(PathType::Directory)
    } else {
        Err(WincentError::invalid_path(
            path,
            "Path is neither a file nor a directory",
        ))
    }
}

/// Validates that a path exists and matches the expected type.
pub(crate) fn validate_path(path: &str, expected_type: PathType) -> WincentResult<()> {
    if path.is_empty() {
        return Err(WincentError::invalid_path_reason("Empty path provided"));
    }

    let path = Path::new(path);
    if !path.exists() {
        return Err(WincentError::invalid_path(path, "Path does not exist"));
    }

    match expected_type {
        PathType::File if !path.is_file() => {
            Err(WincentError::invalid_path(path, "Not a valid file"))
        }
        PathType::Directory if !path.is_dir() => {
            Err(WincentError::invalid_path(path, "Not a valid directory"))
        }
        _ => Ok(()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn detects_file_and_directory_paths() -> WincentResult<()> {
        let directory = tempfile::tempdir().map_err(WincentError::Io)?;
        let file = tempfile::NamedTempFile::new_in(directory.path()).map_err(WincentError::Io)?;

        assert_eq!(
            detect_path_type(&directory.path().to_string_lossy())?,
            PathType::Directory
        );
        assert_eq!(
            detect_path_type(&file.path().to_string_lossy())?,
            PathType::File
        );
        Ok(())
    }

    #[test]
    fn rejects_empty_and_missing_paths() {
        assert!(matches!(
            detect_path_type(""),
            Err(WincentError::InvalidPath(_))
        ));
        assert!(matches!(
            detect_path_type("Z:\\WincentMissing\\item.txt"),
            Err(WincentError::InvalidPath(_))
        ));
    }

    #[test]
    fn nonexistent_path_comparison_uses_lightweight_fallback() {
        assert_eq!(
            normalize_path_for_comparison("Z:/NonExistent/Path/"),
            normalize_path_lightweight("Z:/NonExistent/Path/")
        );
        assert!(paths_equal("Z:/NonExistent/Path/", "z:\\nonexistent\\path"));
    }

    #[test]
    fn validates_expected_path_type() -> WincentResult<()> {
        let directory = tempfile::tempdir().map_err(WincentError::Io)?;
        let file = tempfile::NamedTempFile::new_in(directory.path()).map_err(WincentError::Io)?;
        let directory_path = directory.path().to_string_lossy();
        let file_path = file.path().to_string_lossy();

        validate_path(&directory_path, PathType::Directory)?;
        validate_path(&file_path, PathType::File)?;
        assert!(validate_path(&directory_path, PathType::File).is_err());
        assert!(validate_path(&file_path, PathType::Directory).is_err());
        Ok(())
    }
}
