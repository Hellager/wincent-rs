use crate::WincentResult;
use std::path::PathBuf;
use std::fs::{self, File};
use std::io::Write;

/// Create test environment
pub(crate) fn setup_test_env() -> WincentResult<PathBuf> {
    let test_dir = std::env::current_dir()?.join("tests").join("test_folder");
    fs::create_dir_all(&test_dir)?;
    Ok(test_dir)
}

/// Clean up test environment
pub(crate) fn cleanup_test_env(path: &PathBuf) -> WincentResult<()> {
    if path.exists() {
        fs::remove_dir_all(path)?;
    }
    Ok(())
}

/// Create test file and write content
pub(crate) fn create_test_file(dir: &PathBuf, name: &str, content: &str) -> WincentResult<PathBuf> {
    let file_path = dir.join(name);
    let mut file = File::create(&file_path)?;
    file.write_all(content.as_bytes())?;
    Ok(file_path)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use serial_test::serial;

    #[test]
    #[serial]
    fn test_setup_test_env() -> WincentResult<()> {
        // Create test environment
        let test_dir = setup_test_env()?;
        
        // Verify directory was created
        assert!(test_dir.exists(), "Test directory should exist");
        assert!(test_dir.is_dir(), "Path should be a directory");
        assert!(test_dir.ends_with("tests/test_folder"), "Incorrect directory path");
        
        // Cleanup
        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    #[serial]
    fn test_cleanup_test_env() -> WincentResult<()> {
        // Setup
        let test_dir = setup_test_env()?;
        
        // Create some test content
        create_test_file(&test_dir, "test.txt", "content")?;
        fs::create_dir(test_dir.join("subdir"))?;
        
        // Verify content exists
        assert!(test_dir.join("test.txt").exists());
        assert!(test_dir.join("subdir").exists());
        
        // Cleanup
        cleanup_test_env(&test_dir)?;
        
        // Verify directory and contents were removed
        assert!(!test_dir.exists(), "Test directory should be removed");
        Ok(())
    }

    #[test]
    #[serial]
    fn test_create_test_file() -> WincentResult<()> {
        // Setup
        let test_dir = setup_test_env()?;
        
        // Test creating file with normal content
        let file_path = create_test_file(&test_dir, "test1.txt", "Hello, World!")?;
        assert!(file_path.exists(), "File should exist");
        assert_eq!(
            fs::read_to_string(&file_path)?,
            "Hello, World!",
            "File content should match"
        );
        
        // Test creating file with empty content
        let empty_file = create_test_file(&test_dir, "empty.txt", "")?;
        assert!(empty_file.exists(), "Empty file should exist");
        assert_eq!(
            fs::read_to_string(&empty_file)?,
            "",
            "File should be empty"
        );
        
        // Test creating file with special characters
        let special_content = "Special chars: !@#$%^&*()_+";
        let special_file = create_test_file(&test_dir, "special.txt", special_content)?;
        assert_eq!(
            fs::read_to_string(&special_file)?,
            special_content,
            "Special characters should be preserved"
        );
        
        // Cleanup
        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    #[serial]
    fn test_create_test_file_in_nonexistent_directory() {
        let non_existent_dir = PathBuf::from("non_existent_directory");
        let result = create_test_file(&non_existent_dir, "test.txt", "content");
        assert!(result.is_err(), "Should fail when directory doesn't exist");
    }

    #[test]
    #[serial]
    fn test_cleanup_nonexistent_directory() -> WincentResult<()> {
        let non_existent_dir = PathBuf::from("non_existent_directory");
        let result = cleanup_test_env(&non_existent_dir)?;
        assert!(result == (), "Should succeed silently for non-existent directory");
        Ok(())
    }
}
