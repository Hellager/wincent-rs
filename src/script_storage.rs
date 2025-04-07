use crate::error::WincentError;
use crate::script_strategy::{PSScript, ScriptStrategyFactory};
use crate::WincentResult;
use std::fs::{self, File};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::time::{Duration, SystemTime};

/// Script storage manager
pub(crate) struct ScriptStorage;

impl ScriptStorage {
    const SCRIPT_VERSION: &'static str = env!("CARGO_PKG_VERSION");

    /// Retrieves Wincent temporary directory
    fn get_wincent_temp_dir() -> WincentResult<PathBuf> {
        let temp_dir = std::env::temp_dir().join("wincent");
        if !temp_dir.exists() {
            fs::create_dir_all(&temp_dir).map_err(WincentError::Io)?;
        }
        Ok(temp_dir)
    }

    /// Retrieves static scripts directory
    fn get_static_scripts_dir() -> WincentResult<PathBuf> {
        let static_dir = Self::get_wincent_temp_dir()?.join("static");
        if !static_dir.exists() {
            fs::create_dir_all(&static_dir).map_err(WincentError::Io)?;
        }
        Ok(static_dir)
    }

    /// Retrieves dynamic scripts directory
    fn get_dynamic_scripts_dir() -> WincentResult<PathBuf> {
        let dynamic_dir = Self::get_wincent_temp_dir()?.join("dynamic");
        if !dynamic_dir.exists() {
            fs::create_dir_all(&dynamic_dir).map_err(WincentError::Io)?;
        }

        // Cleanup expired scripts
        Self::cleanup_expired_scripts(&dynamic_dir)?;

        Ok(dynamic_dir)
    }

    fn parse_script_version(file_name: &str) -> Option<String> {
        let base_name = file_name.strip_suffix(".ps1")?;

        let parts: Vec<&str> = base_name.split('_').collect();
        if parts.len() >= 2 {
            Some(parts[1].to_string())
        } else {
            None
        }
    }

    /// Cleans up expired scripts (older than 24 hours)
    fn cleanup_expired_scripts(dir: &Path) -> WincentResult<()> {
        let expiry_duration = Duration::from_secs(24 * 60 * 60); // 24 hours
        let now = SystemTime::now();
        let current_version = Self::SCRIPT_VERSION;

        if let Ok(entries) = fs::read_dir(dir) {
            for entry in entries.flatten() {
                let path = entry.path();

                if path.is_file() && path.extension().is_some_and(|e| e == "ps1") {
                    let mut should_remove = false;
                    if let Some(file_version) = path
                        .file_name()
                        .and_then(|n| n.to_str())
                        .and_then(Self::parse_script_version)
                    {
                        should_remove = file_version != current_version;
                    }
                    if !should_remove {
                        if let Ok(metadata) = entry.metadata() {
                            if let Ok(created) = metadata.created() {
                                should_remove =
                                    now.duration_since(created).unwrap_or(Duration::ZERO)
                                        > expiry_duration;
                            }
                        }
                    }
                    if should_remove {
                        let _ = fs::remove_file(path);
                    }
                }
            }
        }

        Ok(())
    }

    /// Generates parameter hash for dynamic script filenames
    fn hash_parameter(parameter: &str) -> String {
        let digest = md5::compute(parameter.as_bytes());
        format!("{:x}", digest)[..8].to_string() // Take first 8 hexadecimal chars
    }

    /// Creates script file with proper encoding
    fn create_script_file(path: &Path, content: &str) -> WincentResult<()> {
        let mut file = File::create(path).map_err(WincentError::Io)?;

        // Write UTF-8 BOM
        let bom = [0xEF, 0xBB, 0xBF];
        file.write_all(&bom).map_err(WincentError::Io)?;

        // Write script content
        file.write_all(content.as_bytes())
            .map_err(WincentError::Io)?;

        file.flush().map_err(WincentError::Io)?;

        Ok(())
    }

    /// Retrieves static script path (parameter-less scripts)
    pub fn get_script_path(script_type: PSScript) -> WincentResult<PathBuf> {
        let static_dir = Self::get_static_scripts_dir()?;
        let script_name = format!("{:?}_{}.ps1", script_type, Self::SCRIPT_VERSION);
        let script_path = static_dir.join(script_name);

        // Create script if missing
        if !script_path.exists() {
            let content = ScriptStrategyFactory::generate_script(script_type, None)?;
            Self::create_script_file(&script_path, &content)?;
        }

        Ok(script_path)
    }

    /// Retrieves dynamic script path (scripts with parameters)
    pub fn get_dynamic_script_path(
        script_type: PSScript,
        parameter: &str,
    ) -> WincentResult<PathBuf> {
        let dynamic_dir = Self::get_dynamic_scripts_dir()?;
        let param_hash = Self::hash_parameter(parameter);
        let script_name = format!(
            "{:?}_{}_{}.ps1",
            script_type,
            Self::SCRIPT_VERSION,
            param_hash
        );
        let script_path = dynamic_dir.join(script_name);

        // Create script if missing
        if !script_path.exists() {
            let content = ScriptStrategyFactory::generate_script(script_type, Some(parameter))?;
            Self::create_script_file(&script_path, &content)?;
        }

        Ok(script_path)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use filetime::{set_file_mtime, FileTime};
    use std::fs;
    use std::time::SystemTime;

    fn current_version() -> String {
        env!("CARGO_PKG_VERSION").to_string()
    }

    #[test]
    fn test_version_parsing() {
        assert_eq!(
            ScriptStorage::parse_script_version("RefreshExplorer_0.5.2.ps1"),
            Some("0.5.2".into())
        );

        assert_eq!(
            ScriptStorage::parse_script_version("PinToFrequent_0.5.2_abcd1234.ps1"),
            Some("0.5.2".into())
        );

        assert_eq!(ScriptStorage::parse_script_version("InvalidName.ps1"), None);
    }

    #[test_log::test]
    fn test_temp_directory_creation() {
        let result = ScriptStorage::get_wincent_temp_dir();
        assert!(result.is_ok());
        let dir = result.unwrap();
        println!("Temporary directory: {}", dir.to_string_lossy());
        assert!(dir.exists());
        assert!(dir.ends_with("wincent"));
    }

    #[test]
    fn test_static_script_management() {
        let result = ScriptStorage::get_script_path(PSScript::RefreshExplorer);
        assert!(result.is_ok());
        let path = result.unwrap();
        let version_str = format!("_{}.ps1", current_version());
        assert!(path.exists());
        assert!(path.to_string_lossy().contains("RefreshExplorer"));
        assert!(
            path.to_string_lossy().contains(&version_str),
            "Path {} should contain version {}",
            path.display(),
            current_version()
        );

        // Clean up test files
        let _ = fs::remove_file(path);
    }

    #[test]
    fn test_dynamic_script_management() {
        let param = "C:\\Test\\Path";
        let result = ScriptStorage::get_dynamic_script_path(PSScript::PinToFrequentFolder, param);
        assert!(result.is_ok());
        let path = result.unwrap();
        assert!(path.exists());
        assert!(path.to_string_lossy().contains("PinToFrequentFolder"));

        let expected_pattern = format!("PinToFrequentFolder_{}_", current_version());
        assert!(
            path.to_string_lossy().contains(&expected_pattern),
            "Path {} should contain pattern {}",
            path.display(),
            expected_pattern
        );

        // Clean up test files
        let _ = fs::remove_file(path);
    }

    #[test]
    fn test_parameter_hashing() {
        let param1 = "C:\\Test\\Path1";
        let param2 = "C:\\Test\\Path2";

        let hash1 = ScriptStorage::hash_parameter(param1);
        let hash2 = ScriptStorage::hash_parameter(param2);

        assert_ne!(hash1, hash2);
        assert_eq!(hash1.len(), 8);
    }

    #[test]
    fn test_cleanup_logic() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir()?;

        // Create test files with different states
        let current_ver_file = temp_dir
            .path()
            .join(format!("Test_{}.ps1", env!("CARGO_PKG_VERSION")));
        let old_ver_file = temp_dir.path().join("Test_0.4.0.ps1");
        let expired_current_ver = temp_dir.path().join("Test_0.5.2_expired.ps1");
        // Create test files
        File::create(&current_ver_file)?;
        File::create(&old_ver_file)?;
        File::create(&expired_current_ver)?;
        // Set expiration time for the expired file (25 hours ago)
        let expired_time = SystemTime::now() - Duration::from_secs(25 * 3600);
        set_file_mtime(
            &expired_current_ver,
            FileTime::from_system_time(expired_time),
        )?;
        // Execute cleanup
        ScriptStorage::cleanup_expired_scripts(temp_dir.path())?;
        // Verify cleanup results
        assert!(
            current_ver_file.exists(),
            "Current version file should be preserved"
        );
        assert!(
            !old_ver_file.exists(),
            "Outdated version file should be removed"
        );
        assert!(
            !expired_current_ver.exists(),
            "Expired file should be removed regardless of version"
        );
        Ok(())
    }
}
