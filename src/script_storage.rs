use crate::error::WincentError;
use crate::script_strategy::{PSScript, ScriptStrategyFactory};
use crate::WincentResult;
use std::fs::{self, File};
use std::io::{Read, Write};
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
        base_name
            .split('_')
            .rev()
            .find(|t| t.contains('.'))
            .map(|s| s.to_string())
    }

    fn current_process_token() -> String {
        std::process::id().to_string()
    }

    fn is_current_process_dynamic_script(file_name: &str) -> bool {
        let Some(base_name) = file_name.strip_suffix(".ps1") else {
            return false;
        };
        let mut parts = base_name.rsplitn(4, '_');
        let Some(hash) = parts.next() else {
            return false;
        };
        let Some(pid) = parts.next() else {
            return false;
        };
        let Some(version) = parts.next() else {
            return false;
        };
        let Some(script_name) = parts.next() else {
            return false;
        };

        !script_name.is_empty()
            && version.contains('.')
            && pid == Self::current_process_token()
            && pid.chars().all(|c| c.is_ascii_digit())
            && !hash.is_empty()
            && hash.chars().all(|c| c.is_ascii_hexdigit())
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
                    if path
                        .file_name()
                        .and_then(|n| n.to_str())
                        .is_some_and(Self::is_current_process_dynamic_script)
                    {
                        continue;
                    }

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
                            if let Ok(modified) = metadata.modified() {
                                should_remove =
                                    now.duration_since(modified).unwrap_or(Duration::ZERO)
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

    /// Generates parameter hash for dynamic script filenames.
    ///
    /// This is a cache key, not a security boundary. Use the full MD5 digest to
    /// avoid the easy cache collisions that come from a short prefix.
    fn hash_parameter(parameter: &str) -> String {
        let digest = md5::compute(parameter.as_bytes());
        format!("{:x}", digest)
    }

    /// Reads a script file and strips the leading UTF-8 BOM (if present), returning
    /// the raw script text. Returns `None` if the file cannot be read.
    fn read_script_content(path: &Path) -> Option<String> {
        let mut bytes = Vec::new();
        File::open(path).ok()?.read_to_end(&mut bytes).ok()?;
        // Strip 3-byte UTF-8 BOM written by create_script_file
        let content_bytes = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(&bytes);
        String::from_utf8(content_bytes.to_vec()).ok()
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

        let content = ScriptStrategyFactory::generate_script(script_type, None)?;

        // Write when missing or when on-disk content differs from freshly generated content.
        // This ensures that script format changes (e.g. escaping fixes) are propagated to
        // callers even when the same-version script file already exists on disk.
        // A concurrent process may make the same decision at the same time, but the write
        // is idempotent because the generated static content is deterministic.
        let needs_write = match Self::read_script_content(&script_path) {
            Some(existing) => existing != content,
            None => true,
        };
        if needs_write {
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
        let process_token = Self::current_process_token();
        let script_name = format!(
            "{:?}_{}_{}_{}.ps1",
            script_type,
            Self::SCRIPT_VERSION,
            process_token,
            param_hash
        );
        let script_path = dynamic_dir.join(script_name);

        let content = ScriptStrategyFactory::generate_script(script_type, Some(parameter))?;

        // Write when missing or when on-disk content differs from freshly generated content.
        // A concurrent process may race this read-compare-write sequence, but both writers
        // produce identical content for the same script type, version, and parameter.
        let needs_write = match Self::read_script_content(&script_path) {
            Some(existing) => existing != content,
            None => true,
        };
        if needs_write {
            Self::create_script_file(&script_path, &content)?;
        }

        Ok(script_path)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use filetime::{set_file_mtime, FileTime};
    use std::fs::{self, File};
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

        // Script-type names that themselves contain underscores
        assert_eq!(
            ScriptStorage::parse_script_version("Some_Script_0.5.2.ps1"),
            Some("0.5.2".into())
        );

        assert_eq!(
            ScriptStorage::parse_script_version("Some_Script_0.5.2_abcd1234.ps1"),
            Some("0.5.2".into())
        );

        assert_eq!(
            ScriptStorage::parse_script_version("PinToFrequent_0.5.2_12345_abcd1234.ps1"),
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

        // Static scripts are shared by all tests for this crate version. Leave
        // the deterministic cache file in place so parallel tests do not race
        // with this test's cleanup.
    }

    #[test]
    fn test_dynamic_script_management() {
        let param = "C:\\Test\\Path";
        let result = ScriptStorage::get_dynamic_script_path(PSScript::PinToFrequentFolder, param);
        assert!(result.is_ok());
        let path = result.unwrap();
        assert!(path.exists());
        assert!(path.to_string_lossy().contains("PinToFrequentFolder"));
        assert!(path
            .to_string_lossy()
            .contains(&format!("_{}_", ScriptStorage::current_process_token())));

        let expected_pattern = format!("PinToFrequentFolder_{}_", current_version());
        assert!(
            path.to_string_lossy().contains(&expected_pattern),
            "Path {} should contain pattern {}",
            path.display(),
            expected_pattern
        );

        let second =
            ScriptStorage::get_dynamic_script_path(PSScript::PinToFrequentFolder, param).unwrap();
        assert_eq!(path, second, "same process and parameter should reuse path");
    }

    #[test]
    fn test_parameter_hashing() {
        let param1 = "C:\\Test\\Path1";
        let param2 = "C:\\Test\\Path2";

        let hash1 = ScriptStorage::hash_parameter(param1);
        let hash2 = ScriptStorage::hash_parameter(param2);

        assert_ne!(hash1, hash2);
        assert_eq!(hash1.len(), 32);
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

    #[test]
    fn test_cleanup_mtime_branch_current_version() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir()?;
        let ver = env!("CARGO_PKG_VERSION");
        let file = temp_dir.path().join(format!("Script_{}.ps1", ver));
        File::create(&file)?;
        // Set mtime to 25 hours ago so it is expired
        let expired_time = SystemTime::now() - Duration::from_secs(25 * 3600);
        set_file_mtime(&file, FileTime::from_system_time(expired_time))?;
        ScriptStorage::cleanup_expired_scripts(temp_dir.path())?;
        assert!(
            !file.exists(),
            "File with current version but expired mtime should be removed"
        );
        Ok(())
    }

    #[test]
    fn cleanup_preserves_current_process_dynamic_scripts() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir()?;
        let current_pid = ScriptStorage::current_process_token();
        let file = temp_dir.path().join(format!(
            "PinToFrequentFolder_{}_{}_abcd1234.ps1",
            env!("CARGO_PKG_VERSION"),
            current_pid
        ));
        File::create(&file)?;
        let expired_time = SystemTime::now() - Duration::from_secs(25 * 3600);
        set_file_mtime(&file, FileTime::from_system_time(expired_time))?;

        ScriptStorage::cleanup_expired_scripts(temp_dir.path())?;

        assert!(
            file.exists(),
            "Current process dynamic script should be preserved even when expired"
        );
        Ok(())
    }

    #[test]
    fn cleanup_removes_expired_orphan_dynamic_scripts() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir()?;
        let other_pid = std::process::id().saturating_add(1);
        let file = temp_dir.path().join(format!(
            "PinToFrequentFolder_{}_{}_abcd1234.ps1",
            env!("CARGO_PKG_VERSION"),
            other_pid
        ));
        File::create(&file)?;
        let expired_time = SystemTime::now() - Duration::from_secs(25 * 3600);
        set_file_mtime(&file, FileTime::from_system_time(expired_time))?;

        ScriptStorage::cleanup_expired_scripts(temp_dir.path())?;

        assert!(
            !file.exists(),
            "Expired orphan dynamic scripts should be removed by normal cleanup"
        );
        Ok(())
    }

    #[test]
    fn test_static_script_overwrites_stale_content() -> WincentResult<()> {
        // get_script_path must overwrite an existing file whose content differs from
        // the freshly generated script (e.g. after a format/escaping fix in a new build
        // that kept the same semver).
        let result = ScriptStorage::get_script_path(PSScript::RefreshExplorer)?;
        // Overwrite the file with obviously wrong content (simulates stale cache).
        fs::write(&result, b"stale content")?;
        // Calling get_script_path again must detect the mismatch and regenerate.
        let result2 = ScriptStorage::get_script_path(PSScript::RefreshExplorer)?;
        assert_eq!(result, result2);
        let on_disk = ScriptStorage::read_script_content(&result2)
            .expect("should be able to read regenerated script");
        // The regenerated content must exactly match what the strategy produces,
        // not just contain a keyword. This catches the case where stale content
        // happens to include the substring but is otherwise wrong.
        let expected =
            ScriptStrategyFactory::generate_script(PSScript::RefreshExplorer, None).unwrap();
        assert_eq!(
            on_disk, expected,
            "regenerated script must exactly match the current strategy output"
        );
        Ok(())
    }

    #[test]
    fn test_dynamic_script_overwrites_stale_content() -> WincentResult<()> {
        let param = "C:\\Test\\StaleCheck";
        let result = ScriptStorage::get_dynamic_script_path(PSScript::PinToFrequentFolder, param)?;
        fs::write(&result, b"stale content")?;
        let result2 = ScriptStorage::get_dynamic_script_path(PSScript::PinToFrequentFolder, param)?;
        assert_eq!(result, result2);
        let on_disk = ScriptStorage::read_script_content(&result2)
            .expect("should be able to read regenerated script");
        let expected =
            ScriptStrategyFactory::generate_script(PSScript::PinToFrequentFolder, Some(param))
                .unwrap();
        assert_eq!(
            on_disk, expected,
            "regenerated dynamic script must exactly match the current strategy output"
        );
        let _ = fs::remove_file(result2);
        Ok(())
    }
}
