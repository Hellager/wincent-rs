use crate::error::WincentError;
use crate::script_storage::ScriptStorage;
use crate::script_strategy::PSScript;
use crate::utils::get_windows_recent_folder;
use crate::WincentResult;
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::{Command, Output};
use std::sync::{Arc, Mutex};
use std::time::{Duration, SystemTime};
use tokio::task;

/// PowerShell script executor
pub(crate) struct ScriptExecutor;

impl ScriptExecutor {
    /// Executes PowerShell script synchronously
    pub fn execute_ps_script(
        script_type: PSScript,
        parameter: Option<&str>,
    ) -> WincentResult<Output> {
        let script_path = match parameter {
            Some(param) => ScriptStorage::get_dynamic_script_path(script_type, param)?,
            None => ScriptStorage::get_script_path(script_type)?,
        };

        Command::new("powershell")
            .args([
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                script_path.to_str().ok_or_else(|| {
                    WincentError::InvalidPath("Failed to convert script path".to_string())
                })?,
            ])
            .output()
            .map_err(|e| WincentError::PowerShellExecution(e.to_string()))
    }

    /// Executes PowerShell script asynchronously
    pub async fn execute_ps_script_async(
        script_type: PSScript,
        parameter: Option<String>,
    ) -> WincentResult<Output> {
        // Clone parameters for async task
        let param_clone = parameter.clone();

        // Execute in separate thread
        let result = task::spawn_blocking(move || {
            Self::execute_ps_script(script_type, param_clone.as_deref())
        })
        .await
        .map_err(|e| WincentError::AsyncExecution(e.to_string()))??;

        Ok(result)
    }

    /// Parses script output into string collection
    pub fn parse_output_to_strings(output: Output) -> WincentResult<Vec<String>> {
        if !output.status.success() {
            let error = String::from_utf8_lossy(&output.stderr);
            return Err(WincentError::PowerShellExecution(error.to_string()));
        }

        let stdout = String::from_utf8_lossy(&output.stdout);
        let lines: Vec<String> = stdout
            .lines()
            .map(|s| s.trim().to_string())
            .filter(|s| !s.is_empty())
            .collect();

        Ok(lines)
    }

    /// Executes script with timeout protection
    #[allow(dead_code)]
    pub async fn execute_with_timeout(
        script_type: PSScript,
        parameter: Option<String>,
        timeout_secs: u64,
    ) -> WincentResult<Output> {
        let execution = Self::execute_ps_script_async(script_type, parameter);

        match tokio::time::timeout(Duration::from_secs(timeout_secs), execution).await {
            Ok(result) => result,
            Err(_) => Err(WincentError::Timeout(format!(
                "Script execution timed out after {} seconds",
                timeout_secs
            ))),
        }
    }
}

/// Cached script executor with automatic invalidation
pub(crate) struct CachedScriptExecutor {
    cache: Arc<Mutex<HashMap<CacheKey, CacheEntry>>>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct CacheKey {
    script_type: PSScript,
    parameter: Option<String>,
}

struct CacheEntry {
    result: Vec<String>,
    timestamp: SystemTime,
}

/// Windows Quick Access data file information
struct QuickAccessDataFiles {
    recent_files_path: PathBuf,
    frequent_folders_path: PathBuf,
}

impl QuickAccessDataFiles {
    /// Retrieves Quick Access data file paths
    fn new() -> WincentResult<Self> {
        let recent_folder = get_windows_recent_folder()?;
        let automatic_dest_dir = Path::new(&recent_folder).join("AutomaticDestinations");

        let recent_files_path =
            automatic_dest_dir.join("5f7b5f1e01b83767.automaticDestinations-ms");
        let frequent_folders_path =
            automatic_dest_dir.join("f01b4d95cf55d32a.automaticDestinations-ms");

        Ok(Self {
            recent_files_path,
            frequent_folders_path,
        })
    }

    /// Retrieves modification time for recent files data
    fn get_recent_files_modified_time(&self) -> WincentResult<SystemTime> {
        if self.recent_files_path.exists() {
            let metadata = fs::metadata(&self.recent_files_path).map_err(WincentError::Io)?;
            Ok(metadata.modified().unwrap_or(SystemTime::now()))
        } else {
            Ok(SystemTime::now())
        }
    }

    /// Retrieves modification time for frequent folders data
    fn get_frequent_folders_modified_time(&self) -> WincentResult<SystemTime> {
        if self.frequent_folders_path.exists() {
            let metadata = fs::metadata(&self.frequent_folders_path).map_err(WincentError::Io)?;
            Ok(metadata.modified().unwrap_or(SystemTime::now()))
        } else {
            Ok(SystemTime::now())
        }
    }

    /// Retrieves latest modification time for Quick Access data
    fn get_quick_access_modified_time(&self) -> WincentResult<SystemTime> {
        let recent_time = self.get_recent_files_modified_time()?;
        let frequent_time = self.get_frequent_folders_modified_time()?;

        // Return most recent timestamp
        Ok(recent_time.max(frequent_time))
    }

    /// Gets relevant modification time based on script type
    fn get_modified_time_for_script(&self, script_type: PSScript) -> WincentResult<SystemTime> {
        match script_type {
            PSScript::QueryRecentFile => self.get_recent_files_modified_time(),
            PSScript::QueryFrequentFolder => self.get_frequent_folders_modified_time(),
            PSScript::QueryQuickAccess => self.get_quick_access_modified_time(),
            _ => Ok(SystemTime::now()), // Non-query scripts use current time
        }
    }
}

impl CachedScriptExecutor {
    pub fn new() -> Self {
        Self {
            cache: Arc::new(Mutex::new(HashMap::new())),
        }
    }

    /// Determines if script type should be cached
    fn should_cache(script_type: PSScript) -> bool {
        matches!(
            script_type,
            PSScript::QueryQuickAccess | PSScript::QueryRecentFile | PSScript::QueryFrequentFolder
        )
    }

    /// Executes script with cache management
    pub async fn execute(
        &self,
        script_type: PSScript,
        parameter: Option<String>,
    ) -> WincentResult<Vec<String>> {
        // Bypass cache for non-query operations
        if !Self::should_cache(script_type) {
            let output = ScriptExecutor::execute_ps_script_async(script_type, parameter).await?;
            return ScriptExecutor::parse_output_to_strings(output);
        }

        let key = CacheKey {
            script_type,
            parameter: parameter.clone(),
        };

        // Initialize data file tracker
        let data_files = QuickAccessDataFiles::new()?;
        let current_modified_time = data_files.get_modified_time_for_script(script_type)?;

        // Cache check
        {
            let cache = self.cache.lock().unwrap();
            if let Some(entry) = cache.get(&key) {
                // Validate cache using modification timestamp
                if entry.timestamp >= current_modified_time {
                    return Ok(entry.result.clone());
                }
            }
        }

        // Cache miss: execute and store
        let output = ScriptExecutor::execute_ps_script_async(script_type, parameter).await?;
        let result = ScriptExecutor::parse_output_to_strings(output)?;

        // Update cache
        {
            let mut cache = self.cache.lock().unwrap();
            cache.insert(
                key,
                CacheEntry {
                    result: result.clone(),
                    timestamp: current_modified_time,
                },
            );
        }

        Ok(result)
    }

    /// Executes script with timeout protection
    #[allow(dead_code)]
    pub async fn execute_with_timeout(
        &self,
        script_type: PSScript,
        parameter: Option<String>,
        timeout_secs: u64,
    ) -> WincentResult<Vec<String>> {
        match tokio::time::timeout(
            Duration::from_secs(timeout_secs),
            self.execute(script_type, parameter),
        )
        .await
        {
            Ok(result) => result,
            Err(_) => Err(WincentError::Timeout(format!(
                "Script execution timed out after {} seconds",
                timeout_secs
            ))),
        }
    }

    /// Clears entire cache
    pub fn clear_cache(&self) {
        let mut cache = self.cache.lock().unwrap();
        cache.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::os::windows::process::ExitStatusExt;

    #[test]
    fn test_output_parsing() {
        let output = Output {
            status: std::process::ExitStatus::from_raw(0),
            stdout: "Line1\nLine2\n\nLine3".as_bytes().to_vec(),
            stderr: Vec::new(),
        };

        let result = ScriptExecutor::parse_output_to_strings(output).unwrap();
        assert_eq!(result.len(), 3);
        assert_eq!(result[0], "Line1");
        assert_eq!(result[1], "Line2");
        assert_eq!(result[2], "Line3");
    }

    #[test]
    fn test_error_output_handling() {
        let output = Output {
            status: std::process::ExitStatus::from_raw(1),
            stdout: Vec::new(),
            stderr: "Error message".as_bytes().to_vec(),
        };

        let result = ScriptExecutor::parse_output_to_strings(output);
        assert!(result.is_err());
        if let Err(WincentError::PowerShellExecution(msg)) = result {
            assert_eq!(msg, "Error message");
        } else {
            panic!("Expected PowerShellExecution error");
        }
    }

    #[test]
    fn test_cache_eligibility_check() {
        // Cache-eligible script types
        assert!(CachedScriptExecutor::should_cache(
            PSScript::QueryQuickAccess
        ));
        assert!(CachedScriptExecutor::should_cache(
            PSScript::QueryRecentFile
        ));
        assert!(CachedScriptExecutor::should_cache(
            PSScript::QueryFrequentFolder
        ));

        // Non-cacheable script types
        assert!(!CachedScriptExecutor::should_cache(
            PSScript::RefreshExplorer
        ));
        assert!(!CachedScriptExecutor::should_cache(
            PSScript::RemoveRecentFile
        ));
        assert!(!CachedScriptExecutor::should_cache(
            PSScript::PinToFrequentFolder
        ));
        assert!(!CachedScriptExecutor::should_cache(
            PSScript::UnpinFromFrequentFolder
        ));
        assert!(!CachedScriptExecutor::should_cache(
            PSScript::CheckQueryFeasible
        ));
        assert!(!CachedScriptExecutor::should_cache(
            PSScript::CheckPinUnpinFeasible
        ));
    }

    #[tokio::test]
    async fn test_quick_access_file_tracking() -> WincentResult<()> {
        let data_files = QuickAccessDataFiles::new()?;

        // Validate file path patterns
        assert!(data_files
            .recent_files_path
            .to_string_lossy()
            .contains("5f7b5f1e01b83767"));
        assert!(data_files
            .frequent_folders_path
            .to_string_lossy()
            .contains("f01b4d95cf55d32a"));

        // Basic timestamp checks
        let _ = data_files.get_recent_files_modified_time()?;
        let _ = data_files.get_frequent_folders_modified_time()?;
        let _ = data_files.get_quick_access_modified_time()?;

        Ok(())
    }

    #[tokio::test]
    async fn test_cache_management_workflow() -> WincentResult<()> {
        let executor = CachedScriptExecutor::new();

        // Prime cache
        {
            let mut cache_map = executor.cache.lock().unwrap();
            cache_map.insert(
                CacheKey {
                    script_type: PSScript::QueryQuickAccess,
                    parameter: None,
                },
                CacheEntry {
                    result: vec!["cached result".to_string()],
                    timestamp: SystemTime::now() + Duration::from_secs(3600),
                },
            );
        }

        // Cache hit scenario
        let result = executor.execute(PSScript::QueryQuickAccess, None).await?;
        assert_eq!(result, vec!["cached result".to_string()]);

        // Cache clearance test
        executor.clear_cache();
        let cache_map = executor.cache.lock().unwrap();
        assert!(cache_map.is_empty());

        Ok(())
    }
}
