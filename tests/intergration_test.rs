#[cfg(test)]
mod tests {
    use std::fs::{self, File};
    use std::io::Write;
    use std::path::{Path, PathBuf};
    use std::{thread, time::Duration};
    use wincent::predule::*;

    /// Create test environment
    pub(crate) fn setup_test_env() -> WincentResult<PathBuf> {
        let test_dir = std::env::current_dir()?.join("tests").join("test_folder");
        fs::create_dir_all(&test_dir)?;
        Ok(test_dir)
    }

    /// Clean up test environment
    pub(crate) fn cleanup_test_env(path: &Path) -> WincentResult<()> {
        if path.exists() {
            fs::remove_dir_all(path)?;
        }
        Ok(())
    }

    /// Create test file and write content
    pub(crate) fn create_test_file(
        dir: &Path,
        name: &str,
        content: &str,
    ) -> WincentResult<PathBuf> {
        let file_path = dir.join(name);
        let mut file = File::create(&file_path)?;
        file.write_all(content.as_bytes())?;
        Ok(file_path)
    }

    #[test_log::test]
    #[ignore]
    fn test_feasibility_checks() -> WincentResult<()> {
        // Test script execution feasibility
        let script_feasible = check_script_feasible()?;
        println!("Script execution feasible: {}", script_feasible);

        // Test query feasibility
        let query_feasible = check_query_feasible()?;
        println!("Query operation feasible: {}", query_feasible);

        // Test pin/unpin feasibility
        let pinunpin_feasible = check_pinunpin_feasible()?;
        println!("Pin/Unpin operation feasible: {}", pinunpin_feasible);

        // If any check fails, try fixing
        if !script_feasible || !query_feasible || !pinunpin_feasible {
            println!("Attempting to fix feasibility...");
            fix_script_feasible()?;

            let fixed = check_feasible()?;
            if fixed {
                println!("Successfully fixed feasibility");
            } else {
                println!("Failed to fix feasibility");
            }
        }

        Ok(())
    }

    #[test]
    #[ignore]
    fn test_quick_access_operations() -> WincentResult<()> {
        let test_dir = setup_test_env()?;

        // Create test files
        let test_file = create_test_file(&test_dir, "test.txt", "test content")?;
        let test_path = test_file.to_str().unwrap();

        // Test adding to recent files
        add_to_recent_files(test_path)?;
        thread::sleep(Duration::from_millis(500));

        // Verify file was added
        assert!(
            is_in_recent_files(test_path)?,
            "File should be in recent files"
        );

        // Test adding folder to frequent folders
        let dir_path = test_dir.to_str().unwrap();
        add_to_frequent_folders(dir_path)?;
        thread::sleep(Duration::from_millis(500));

        // Verify folder was added
        assert!(
            is_in_frequent_folders(dir_path)?,
            "Folder should be in frequent folders"
        );

        // Test removal operations
        remove_from_recent_files(test_path)?;
        remove_from_frequent_folders(dir_path)?;
        thread::sleep(Duration::from_millis(500));

        // Verify removals
        assert!(
            !is_in_recent_files(test_path)?,
            "File should not be in recent files"
        );
        assert!(
            !is_in_frequent_folders(dir_path)?,
            "Folder should not be in frequent folders"
        );

        cleanup_test_env(&test_dir)?;
        Ok(())
    }

    #[test]
    #[ignore]
    fn test_visibility_operations() -> WincentResult<()> {
        // Save initial states
        let initial_recent = is_recent_files_visiable()?;
        let initial_frequent = is_frequent_folders_visible()?;

        // Test visibility toggling
        set_recent_files_visiable(!initial_recent)?;
        set_frequent_folders_visiable(!initial_frequent)?;

        // Verify changes
        assert_eq!(
            !initial_recent,
            is_recent_files_visiable()?,
            "Recent files visibility should be toggled"
        );
        assert_eq!(
            !initial_frequent,
            is_frequent_folders_visible()?,
            "Frequent folders visibility should be toggled"
        );

        // Restore initial states
        set_recent_files_visiable(initial_recent)?;
        set_frequent_folders_visiable(initial_frequent)?;

        Ok(())
    }

    #[test_log::test]
    #[ignore]
    fn test_empty_operations() -> WincentResult<()> {
        let test_dir = setup_test_env()?;

        // Test empty_recent_files
        let test_file = create_test_file(&test_dir, "test.txt", "test content")?;
        add_to_recent_files(test_file.to_str().unwrap())?;
        thread::sleep(Duration::from_millis(500));

        assert!(
            is_in_recent_files(test_file.to_str().unwrap())?,
            "File should be in recent files"
        );
        empty_recent_files()?;
        thread::sleep(Duration::from_millis(500));
        assert!(
            !is_in_recent_files(test_file.to_str().unwrap())?,
            "Recent files should be empty"
        );

        // Test empty_frequent_folders
        let dir_path = test_dir.to_str().unwrap();
        add_to_frequent_folders(dir_path)?;
        thread::sleep(Duration::from_millis(500));

        assert!(
            is_in_frequent_folders(dir_path)?,
            "Folder should be in frequent folders"
        );
        empty_frequent_folders()?;
        thread::sleep(Duration::from_millis(500));
        assert!(
            !is_in_frequent_folders(dir_path)?,
            "Frequent folders should be empty"
        );

        // Test empty_quick_access
        add_to_recent_files(test_file.to_str().unwrap())?;
        add_to_frequent_folders(dir_path)?;
        thread::sleep(Duration::from_millis(500));

        assert!(
            is_in_quick_access(test_file.to_str().unwrap())? || is_in_quick_access(dir_path)?,
            "Items should be in Quick Access"
        );

        empty_quick_access()?;
        thread::sleep(Duration::from_millis(500));

        assert!(
            !is_in_quick_access(test_file.to_str().unwrap())? && !is_in_quick_access(dir_path)?,
            "Quick Access should be empty"
        );

        cleanup_test_env(&test_dir)?;
        Ok(())
    }
}
