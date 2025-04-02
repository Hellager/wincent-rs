use crate::error::WincentError;
use crate::WincentResult;
use crate::script_strategy::{PSScript, ScriptStrategyFactory};
use std::fs::{self, File};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::time::{Duration, SystemTime};

/// Script storage manager
pub(crate) struct ScriptStorage;

impl ScriptStorage {
    /// Retrieves Wincent temporary directory
    fn get_wincent_temp_dir() -> WincentResult<PathBuf> {
        let temp_dir = std::env::temp_dir().join("wincent");
        if !temp_dir.exists() {
            fs::create_dir_all(&temp_dir)
                .map_err(WincentError::Io)?;
        }
        Ok(temp_dir)
    }
    
    /// Retrieves static scripts directory
    fn get_static_scripts_dir() -> WincentResult<PathBuf> {
        let static_dir = Self::get_wincent_temp_dir()?.join("static");
        if !static_dir.exists() {
            fs::create_dir_all(&static_dir)
                .map_err(WincentError::Io)?;
        }
        Ok(static_dir)
    }
    
    /// Retrieves dynamic scripts directory
    fn get_dynamic_scripts_dir() -> WincentResult<PathBuf> {
        let dynamic_dir = Self::get_wincent_temp_dir()?.join("dynamic");
        if !dynamic_dir.exists() {
            fs::create_dir_all(&dynamic_dir)
                .map_err(WincentError::Io)?;
        }
        
        // Cleanup expired scripts
        Self::cleanup_expired_scripts(&dynamic_dir)?;
        
        Ok(dynamic_dir)
    }
    
    /// Cleans up expired scripts (older than 24 hours)
    fn cleanup_expired_scripts(dir: &Path) -> WincentResult<()> {
        let expiry_duration = Duration::from_secs(24 * 60 * 60); // 24 hours
        let now = SystemTime::now();
        
        if let Ok(entries) = fs::read_dir(dir) {
            for entry in entries.flatten() {
                if let Ok(metadata) = entry.metadata() {
                    if metadata.is_file() {
                        if let Ok(created_time) = metadata.created() {
                            if now.duration_since(created_time).unwrap_or(Duration::from_secs(0)) > expiry_duration {
                                // Ignore removal errors and continue
                                let _ = fs::remove_file(entry.path());
                            }
                        }
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
        let mut file = File::create(path)
            .map_err(WincentError::Io)?;
            
        // Write UTF-8 BOM
        let bom = [0xEF, 0xBB, 0xBF];
        file.write_all(&bom)
            .map_err(WincentError::Io)?;
            
        // Write script content
        file.write_all(content.as_bytes())
            .map_err(WincentError::Io)?;
            
        file.flush()
            .map_err(WincentError::Io)?;
            
        Ok(())
    }
    
    /// Retrieves static script path (parameter-less scripts)
    pub fn get_script_path(script_type: PSScript) -> WincentResult<PathBuf> {
        let static_dir = Self::get_static_scripts_dir()?;
        let script_name = format!("{:?}.ps1", script_type);
        let script_path = static_dir.join(script_name);
        
        // Create script if missing
        if !script_path.exists() {
            let content = ScriptStrategyFactory::generate_script(script_type, None)?;
            Self::create_script_file(&script_path, &content)?;
        }
        
        Ok(script_path)
    }
    
    /// Retrieves dynamic script path (scripts with parameters)
    pub fn get_dynamic_script_path(script_type: PSScript, parameter: &str) -> WincentResult<PathBuf> {
        let dynamic_dir = Self::get_dynamic_scripts_dir()?;
        let param_hash = Self::hash_parameter(parameter);
        let script_name = format!("{:?}_{}.ps1", script_type, param_hash);
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
    use std::fs;
    
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
        assert!(path.exists());
        assert!(path.to_string_lossy().contains("RefreshExplorer"));
        
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
}
