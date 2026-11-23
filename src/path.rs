//! Windows path validation and comparison helpers.

use crate::{error::WincentError, WincentResult};
use std::collections::HashMap;
use std::io;
use std::path::{Path, PathBuf};

/// Converts a path to the string representation used by Shell APIs.
pub(crate) fn path_to_shell_string(path: &Path) -> WincentResult<String> {
    if path.as_os_str().is_empty() {
        return Err(WincentError::invalid_path_reason("Empty path provided"));
    }
    Ok(path.to_string_lossy().into_owned())
}

/// Lightweight path normalization without I/O operations.
pub(crate) fn normalize_path_lightweight(path: &str) -> String {
    let mut result = path.to_lowercase().replace('/', "\\");
    if result.len() > 3 && result.ends_with('\\') {
        result.pop();
    }
    result
}

fn normalize_path_for_comparison_with<F>(path: &str, canonicalize: F) -> String
where
    F: Fn(&Path) -> io::Result<PathBuf>,
{
    let normalized = canonicalize(Path::new(path)).map_or_else(
        |_| path.to_string(),
        |value| value.to_string_lossy().into_owned(),
    );
    normalize_path_lightweight(&normalized)
}

#[derive(Default)]
pub(crate) struct PathComparisonCache {
    normalized_paths: HashMap<String, String>,
}

impl PathComparisonCache {
    /// Compares paths while reusing canonicalization results within a batch.
    pub(crate) fn paths_equal(&mut self, left: &str, right: &str) -> bool {
        self.paths_equal_with(left, right, |value| value.canonicalize())
    }

    fn paths_equal_with<F>(&mut self, left: &str, right: &str, canonicalize: F) -> bool
    where
        F: Fn(&Path) -> io::Result<PathBuf>,
    {
        let normalized_left = normalize_path_lightweight(left);
        let normalized_right = normalize_path_lightweight(right);
        if normalized_left == normalized_right {
            return true;
        }

        let canonicalized_left = self.normalized_path_with(left, &canonicalize);
        let canonicalized_right = self.normalized_path_with(right, &canonicalize);
        canonicalized_left == canonicalized_right
    }

    fn normalized_path_with<F>(&mut self, path: &str, canonicalize: &F) -> String
    where
        F: Fn(&Path) -> io::Result<PathBuf>,
    {
        self.normalized_paths
            .entry(path.to_owned())
            .or_insert_with(|| normalize_path_for_comparison_with(path, canonicalize))
            .clone()
    }
}

/// Checks whether two Windows paths refer to the same location.
pub(crate) fn paths_equal(left: &str, right: &str) -> bool {
    PathComparisonCache::default().paths_equal(left, right)
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
    use std::cell::Cell;

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
    fn comparison_cache_reuses_failed_canonicalization() {
        let mut cache = PathComparisonCache::default();
        let canonicalize_calls = Cell::new(0);
        let canonicalize = |_: &Path| {
            canonicalize_calls.set(canonicalize_calls.get() + 1);
            Err(io::Error::from(io::ErrorKind::NotFound))
        };

        assert!(!cache.paths_equal_with(
            "Z:/NonExistent/Path/one",
            "z:\\nonexistent\\path\\two",
            canonicalize,
        ));
        assert_eq!(canonicalize_calls.get(), 2);
        assert!(!cache.paths_equal_with(
            "Z:/NonExistent/Path/one",
            "z:\\nonexistent\\path\\two",
            canonicalize,
        ));
        assert_eq!(canonicalize_calls.get(), 2);
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
