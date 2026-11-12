//! Windows path validation and comparison helpers.

use crate::{error::WincentError, WincentResult};
use std::io;
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
    normalize_path_for_comparison_with(path, |value| value.canonicalize())
}

fn normalize_path_for_comparison_with<F>(path: &str, canonicalize: F) -> String
where
    F: FnOnce(&Path) -> io::Result<std::path::PathBuf>,
{
    let normalized = canonicalize(Path::new(path)).map_or_else(
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
    detect_path_type_with(path, |value| {
        let metadata = std::fs::metadata(value)?;
        if metadata.is_file() {
            Ok(Some(PathType::File))
        } else if metadata.is_dir() {
            Ok(Some(PathType::Directory))
        } else {
            Ok(None)
        }
    })
}

fn detect_path_type_with<F>(path: &str, probe: F) -> WincentResult<PathType>
where
    F: FnOnce(&Path) -> io::Result<Option<PathType>>,
{
    if path.is_empty() {
        return Err(WincentError::invalid_path_reason("Empty path provided"));
    }

    let path_type = probe(Path::new(path)).map_err(|error| {
        if error.kind() == std::io::ErrorKind::NotFound {
            WincentError::invalid_path(path, "Path does not exist")
        } else {
            WincentError::Io(error)
        }
    })?;
    path_type
        .ok_or_else(|| WincentError::invalid_path(path, "Path is neither a file nor a directory"))
}

/// Validates that a path exists and matches the expected type.
pub(crate) fn validate_path(path: &str, expected_type: PathType) -> WincentResult<()> {
    validate_path_with(path, expected_type, detect_path_type)
}

fn validate_path_with<F>(path: &str, expected_type: PathType, probe: F) -> WincentResult<()>
where
    F: FnOnce(&str) -> WincentResult<PathType>,
{
    let actual_type = probe(path)?;
    if actual_type == expected_type {
        Ok(())
    } else {
        let reason = match expected_type {
            PathType::File => "Not a valid file",
            PathType::Directory => "Not a valid directory",
        };
        Err(WincentError::invalid_path(path, reason))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn detects_file_and_directory_paths_without_file_system_access() -> WincentResult<()> {
        assert_eq!(
            detect_path_type_with("C:\\fixture\\directory", |_| Ok(Some(PathType::Directory)))?,
            PathType::Directory
        );
        assert_eq!(
            detect_path_type_with("C:\\fixture\\file.txt", |_| Ok(Some(PathType::File)))?,
            PathType::File
        );
        Ok(())
    }

    #[test]
    fn rejects_empty_and_missing_paths() {
        assert!(matches!(
            detect_path_type_with("", |_| Ok(Some(PathType::File))),
            Err(WincentError::InvalidPath(_))
        ));
        assert!(matches!(
            detect_path_type_with("Z:\\WincentMissing\\item.txt", |_| {
                Err(io::Error::from(io::ErrorKind::NotFound))
            }),
            Err(WincentError::InvalidPath(_))
        ));
        assert!(matches!(
            detect_path_type_with("C:\\Denied", |_| {
                Err(io::Error::from(io::ErrorKind::PermissionDenied))
            }),
            Err(WincentError::Io(_))
        ));
        assert!(matches!(
            detect_path_type_with("C:\\Device", |_| Ok(None)),
            Err(WincentError::InvalidPath(_))
        ));
    }

    #[test]
    fn nonexistent_path_comparison_uses_lightweight_fallback() {
        assert_eq!(
            normalize_path_for_comparison_with("Z:/NonExistent/Path/", |_| {
                Err(io::Error::from(io::ErrorKind::NotFound))
            }),
            normalize_path_lightweight("Z:/NonExistent/Path/")
        );
        assert!(paths_equal("Z:/NonExistent/Path/", "z:\\nonexistent\\path"));
    }

    #[test]
    fn canonicalized_existing_paths_compare_equal_without_file_system_access() {
        let left = normalize_path_for_comparison_with("C:\\fixture\\nested\\..\\item.txt", |_| {
            Ok(std::path::PathBuf::from("C:\\fixture\\item.txt"))
        });
        let right = normalize_path_for_comparison_with("c:/fixture/item.txt", |_| {
            Ok(std::path::PathBuf::from("C:\\fixture\\item.txt"))
        });
        assert_eq!(left, right);
    }

    #[test]
    fn validates_expected_path_type_without_file_system_access() -> WincentResult<()> {
        validate_path_with("C:\\directory", PathType::Directory, |_| {
            Ok(PathType::Directory)
        })?;
        validate_path_with("C:\\file.txt", PathType::File, |_| Ok(PathType::File))?;
        assert!(validate_path_with("C:\\directory", PathType::File, |_| {
            Ok(PathType::Directory)
        })
        .is_err());
        assert!(
            validate_path_with("C:\\file.txt", PathType::Directory, |_| {
                Ok(PathType::File)
            })
            .is_err()
        );
        Ok(())
    }
}
