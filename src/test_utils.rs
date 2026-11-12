#![allow(dead_code)]

use crate::WincentResult;
use std::fs::{self, File};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicUsize, Ordering};

static TEST_DIR_COUNTER: AtomicUsize = AtomicUsize::new(0);

/// Create test environment
pub(crate) fn setup_test_env() -> WincentResult<PathBuf> {
    let base = std::env::temp_dir().join("wincent-rs-tests");
    fs::create_dir_all(&base)?;

    loop {
        let counter = TEST_DIR_COUNTER.fetch_add(1, Ordering::Relaxed);
        let test_dir = base.join(format!("{}-{counter}", std::process::id()));
        match fs::create_dir(&test_dir) {
            Ok(()) => return Ok(test_dir),
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error.into()),
        }
    }
}

/// Clean up test environment
pub(crate) fn cleanup_test_env(path: &Path) -> WincentResult<()> {
    if path.exists() {
        fs::remove_dir_all(path)?;
    }
    Ok(())
}

/// Create test file and write content
pub(crate) fn create_test_file(dir: &Path, name: &str, content: &str) -> WincentResult<PathBuf> {
    let file_path = dir.join(name);
    let mut file = File::create(&file_path)?;
    file.write_all(content.as_bytes())?;
    Ok(file_path)
}
