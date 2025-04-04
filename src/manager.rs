//! Windows Quick Access management and operations
//!
//! Provides unified interface for querying and managing Windows Quick Access items,
//! including recent files and frequent folders. Implements caching for performance.

use crate::{
    error::WincentError,
    script_executor::CachedScriptExecutor,
    script_strategy::PSScript,
    QuickAccess,
    WincentResult,
    handle::{validate_path, PathType, add_file_to_recent_with_api},
    empty::{empty_recent_files_with_api, empty_normal_folders_with_jumplist_file},
    feasible::{check_script_feasible, fix_script_feasible},
};
use std::sync::Arc;
use tokio::sync::OnceCell;
use tokio::time::{timeout, Duration};

/// Windows Quick Access management system
pub struct QuickAccessManager {
    executor: Arc<CachedScriptExecutor>,
    script_feasibility: OnceCell<bool>,
    query_feasibility: OnceCell<bool>,
    handle_feasibility: OnceCell<bool>,
    lock_timeout: Duration,
}

impl QuickAccessManager {
    /// Creates new Quick Access manager
    pub async fn new() -> WincentResult<Self> {
        let manager = Self {
            executor: Arc::new(CachedScriptExecutor::new()),
            script_feasibility: OnceCell::new(),
            query_feasibility: OnceCell::new(),
            handle_feasibility: OnceCell::new(),
            lock_timeout: Duration::from_secs(10),
        };
        
        let check_result = tokio::task::spawn_blocking(check_script_feasible)
            .await??;
    
        if !check_result {
            let _ = fix_script_feasible();
            let recheck = tokio::task::spawn_blocking(check_script_feasible)
                .await??;
            manager.script_feasibility.set(recheck).unwrap_or(());
        } else {
            manager.script_feasibility.set(true).unwrap_or(());
        }
        
        Ok(manager)
    }
    
    pub async fn check_feasible(&self) -> WincentResult<bool> {
        let script_feasible = match self.script_feasibility.get() {
            Some(&feasible) => feasible,
            None => {
                let feasible = check_script_feasible()?;
                self.script_feasibility.set(feasible).unwrap_or(());
                feasible
            }
        };
        
        if !script_feasible {
            return Ok(false);
        }
        
        let query_feasible = match self.query_feasibility.get() {
            Some(&feasible) => feasible,
            None => {
                let feasible = self.check_query_feasible_async().await?;
                self.query_feasibility.set(feasible).unwrap_or(());
                feasible
            }
        };
        
        if !query_feasible {
            return Ok(false);
        }
        
        let handle_feasible = match self.handle_feasibility.get() {
            Some(&feasible) => feasible,
            None => {
                let feasible = self.check_handle_feasible_async().await?;
                self.handle_feasibility.set(feasible).unwrap_or(());
                feasible
            }
        };
        
        Ok(script_feasible && query_feasible && handle_feasible)
    }
    
    fn map_to_script_type(&self, qa_type: QuickAccess) -> WincentResult<PSScript> {
        match qa_type {
            QuickAccess::All => Ok(PSScript::QueryQuickAccess),
            QuickAccess::RecentFiles => Ok(PSScript::QueryRecentFile),
            QuickAccess::FrequentFolders => Ok(PSScript::QueryFrequentFolder)
        }
    }

    async fn check_query_feasible_async(&self) -> WincentResult<bool> {
        let _ = timeout(
            self.lock_timeout,
            self.executor.execute(PSScript::CheckQueryFeasible, None)
        ).await.map_err(|_| WincentError::Timeout("Query feasibility check timed out".to_string()))??;
        
        Ok(true)
    }
    
    async fn ensure_query_feasible(&self) -> WincentResult<()> {
        let _ = match self.query_feasibility.get() {
            Some(&false) => return Err(WincentError::UnsupportedOperation(
                "Query operation not feasible".to_string()
            )),
            Some(&true) => true,
            None => {
                let feasible = self.check_query_feasible_async().await?;
                self.query_feasibility.set(feasible).unwrap_or(());
                
                if !feasible {
                    return Err(WincentError::UnsupportedOperation(
                        "Query operation not feasible".to_string()
                    ));
                }
                true
            }
        };
        
        Ok(())
    }
    
    async fn check_handle_feasible_async(&self) -> WincentResult<bool> {
        let _ = timeout(
            self.lock_timeout,
            self.executor.execute(PSScript::CheckPinUnpinFeasible, None)
        ).await.map_err(|_| WincentError::Timeout("Handle feasibility check timed out".to_string()))??;
        
        Ok(true)
    }

    async fn ensure_handle_feasible(&self) -> WincentResult<()> {
        let _ = match self.handle_feasibility.get() {
            Some(&false) => return Err(WincentError::UnsupportedOperation(
                "Handle operation not feasible".to_string()
            )),
            Some(&true) => true,
            None => {
                let feasible = self.check_handle_feasible_async().await?;
                self.handle_feasibility.set(feasible).unwrap_or(());
                
                if !feasible {
                    return Err(WincentError::UnsupportedOperation(
                        "Handle operation not feasible".to_string()
                    ));
                }
                true
            }
        };
        
        Ok(())
    }
    
    /// Retrieves Quick Access items
    ///
    /// # Arguments
    ///
    /// * `qa_type` - Target Quick Access category
    ///
    /// # Returns
    ///
    /// List of path strings
    pub async fn get_items(&self, qa_type: QuickAccess) -> WincentResult<Vec<String>> {
        self.ensure_query_feasible().await?;
        
        let script_type = self.map_to_script_type(qa_type)?;
        self.executor.execute(script_type, None).await
    }
    
    /// Checks item presence in Quick Access
    ///
    /// # Arguments
    ///
    /// * `path` - Target path to check
    /// * `qa_type` - Quick Access category to search
    ///
    /// # Returns
    ///
    /// `true` if item exists, `false` otherwise
    pub async fn check_item(&self, path: &str, qa_type: QuickAccess) -> WincentResult<bool> {
        self.ensure_query_feasible().await?;
        
        let items = self.get_items(qa_type).await?;
        Ok(items.iter().any(|item| item == path))
    }
    
    /// Adds item to Quick Access
    ///
    /// # Arguments
    ///
    /// * `path` - Path to add
    /// * `qa_type` - Target Quick Access category
    pub async fn add_item(&self, path: &str, qa_type: QuickAccess) -> WincentResult<()> {
        self.ensure_handle_feasible().await?;
        
        match qa_type {
            QuickAccess::RecentFiles => {
                validate_path(path, PathType::File)?;
                add_file_to_recent_with_api(path)?;
            },
            QuickAccess::FrequentFolders => {
                validate_path(path, PathType::Directory)?;
                let result = self.executor.execute(PSScript::PinToFrequentFolder, Some(path.to_string())).await?;
                if !result.is_empty() {
                    return Err(WincentError::ScriptFailed(format!("Failed pinning folder: {}", path)));
                }
            },
            _ => return Err(WincentError::UnsupportedOperation(
                format!("Unsupported add operation for {:?}", qa_type)
            )),
        }
        
        self.executor.clear_cache();
        Ok(())
    }
    
    /// Removes item from Quick Access
    ///
    /// # Arguments
    ///
    /// * `path` - Path to remove
    /// * `qa_type` - Target Quick Access category
    pub async fn remove_item(&self, path: &str, qa_type: QuickAccess) -> WincentResult<()> {
        self.ensure_handle_feasible().await?;
        
        match qa_type {
            QuickAccess::RecentFiles => {
                validate_path(path, PathType::File)?;
                let result = self.executor.execute(PSScript::RemoveRecentFile, Some(path.to_string())).await?;
                if !result.is_empty() {
                    return Err(WincentError::ScriptFailed(format!("Failed removing file: {}", path)));
                }
            },
            QuickAccess::FrequentFolders => {
                validate_path(path, PathType::Directory)?;
                let result = self.executor.execute(PSScript::UnpinFromFrequentFolder, Some(path.to_string())).await?;
                if !result.is_empty() {
                    return Err(WincentError::ScriptFailed(format!("Failed unpinning folder: {}", path)));
                }
            },
            _ => return Err(WincentError::UnsupportedOperation(
                format!("Unsupported remove operation for {:?}", qa_type)
            )),
        }
        
        self.executor.clear_cache();
        Ok(())
    }
    
    /// Clears Quick Access items
    ///
    /// # Arguments
    ///
    /// * `qa_type` - Target Quick Access category to clear
    pub async fn empty_items(&self, qa_type: QuickAccess) -> WincentResult<()> {
        self.ensure_handle_feasible().await?;
        
        match qa_type {
            QuickAccess::RecentFiles => {
                empty_recent_files_with_api()?;
            },
            QuickAccess::FrequentFolders => {
                empty_normal_folders_with_jumplist_file()?;
                self.executor.execute(PSScript::EmptyPinnedFolders, None).await?;
            },
            QuickAccess::All => {
                Box::pin(self.empty_items(QuickAccess::RecentFiles)).await?;
                Box::pin(self.empty_items(QuickAccess::FrequentFolders)).await?;
            }
        }
        
        self.executor.clear_cache();
        Ok(())
    }
    
    /// Clears internal cache
    pub fn clear_cache(&self) {
        self.executor.clear_cache();
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    
    #[tokio::test]
    #[ignore = "Modifies system state"]
    async fn test_feasibility_check() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;

        let feasible = manager.check_feasible().await?;
        println!("Operation feasibility status: {}", feasible);
    
        Ok(())
    }
    
    #[tokio::test]
    async fn test_item_retrieval() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;
        
        // Test full retrieval
        let all_items = manager.get_items(QuickAccess::All).await?;
        println!("Total items found: {}", all_items.len());
        
        // Test recent files
        let recent_files = manager.get_items(QuickAccess::RecentFiles).await?;
        println!("Recent files count: {}", recent_files.len());
        
        // Test frequent folders
        let frequent_folders = manager.get_items(QuickAccess::FrequentFolders).await?;
        println!("Frequent folders count: {}", frequent_folders.len());
        
        Ok(())
    }
    
    #[tokio::test]
    async fn test_item_presence_check() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;
        let items = manager.get_items(QuickAccess::All).await?;
        
        if let Some(item) = items.first() {
            let exists = manager.check_item(item, QuickAccess::All).await?;
            assert!(exists, "Item should exist in collection");
        }
        
        let non_existent = "Z:\\Invalid\\Path\\Test.txt";
        let exists = manager.check_item(non_existent, QuickAccess::All).await?;
        assert!(!exists, "Non-existent item should not be present");
        
        Ok(())
    }
    
    #[tokio::test]
    #[ignore = "Modifies system state"]
    async fn test_item_management_cycle() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;
        
        // File management test
        let temp_file = tempfile::Builder::new()
            .prefix("wincent-test-")
            .suffix(".txt")
            .tempfile()?;
        let file_path = temp_file.path().to_str().unwrap();
        
        manager.add_item(file_path, QuickAccess::RecentFiles).await?;
        let exists = manager.check_item(file_path, QuickAccess::RecentFiles).await?;
        assert!(exists, "File should be added");
        
        manager.remove_item(file_path, QuickAccess::RecentFiles).await?;
        let exists = manager.check_item(file_path, QuickAccess::RecentFiles).await?;
        assert!(!exists, "File should be removed");
        
        // Directory management test
        let temp_dir = tempfile::Builder::new()
            .prefix("wincent-test-")
            .tempdir()?;
        let dir_path = temp_dir.path().to_str().unwrap();
        
        manager.add_item(dir_path, QuickAccess::FrequentFolders).await?;
        let exists = manager.check_item(dir_path, QuickAccess::FrequentFolders).await?;
        assert!(exists, "Directory should be added");
        
        manager.remove_item(dir_path, QuickAccess::FrequentFolders).await?;
        let exists = manager.check_item(dir_path, QuickAccess::FrequentFolders).await?;
        assert!(!exists, "Directory should be removed");
        
        Ok(())
    }
    
    #[tokio::test]
    #[ignore = "Modifies system state"]
    async fn test_collection_clearance() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;
        
        manager.empty_items(QuickAccess::RecentFiles).await?;
        let recent_files = manager.get_items(QuickAccess::RecentFiles).await?;
        assert!(recent_files.is_empty(), "Recent files should be cleared");
        
        manager.empty_items(QuickAccess::FrequentFolders).await?;
        let frequent_folders = manager.get_items(QuickAccess::FrequentFolders).await?;
        assert!(frequent_folders.is_empty(), "Frequent folders should be cleared");
        
        Ok(())
    }
}
