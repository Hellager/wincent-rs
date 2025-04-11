//! Windows Quick Access management and operations
//!
//! Provides unified interface for querying and managing Windows Quick Access items,
//! including recent files and frequent folders. Implements caching for performance.
//!
//! # Key Functionality
//! - Query Quick Access items (recent files, frequent folders)
//! - Add/remove items from Quick Access categories
//! - Clear entire Quick Access categories
//! - Feasibility checking for system operations
//! - Cached PowerShell script execution

use crate::{
    empty::{empty_normal_folders_with_jumplist_file, empty_recent_files_with_api},
    error::WincentError,
    handle::add_file_to_recent_with_api,
    script_executor::{CachedScriptExecutor, QuickAccessDataFiles},
    script_strategy::PSScript,
    utils::{validate_path, PathType},
    QuickAccess, WincentResult,
};
use std::sync::Arc;
use tokio::sync::OnceCell;
use tokio::time::Duration;

/// Represents system capability status for Quick Access operations
#[derive(Debug)]
struct FeasibilityStatus {
    handle: bool,
    query: bool,
}

impl FeasibilityStatus {
    async fn check(executor: &Arc<CachedScriptExecutor>, timeout_duration: Duration) -> Self {
        let query_feasible =
            Self::check_feasibility(executor, PSScript::CheckQueryFeasible, timeout_duration).await;

        let handle_feasible =
            Self::check_feasibility(executor, PSScript::CheckPinUnpinFeasible, timeout_duration)
                .await;

        Self {
            query: query_feasible,
            handle: handle_feasible,
        }
    }

    async fn check_feasibility(
        executor: &Arc<CachedScriptExecutor>,
        script: PSScript,
        timeout_duration: Duration,
    ) -> bool {
        let check_result =
            tokio::time::timeout(timeout_duration, executor.execute(script, None)).await;

        match check_result {
            Ok(Ok(_)) => true,
            Ok(Err(_)) => false, // Execution failed
            Err(_) => false,     // Timeout occurred
        }
    }
}

/// Windows Quick Access management system
///
/// # Example
/// ```rust,no_run
/// use wincent::predule::*;
///
/// #[tokio::main]
/// async fn main() {
///     let manager = QuickAccessManager::new().await.unwrap();
///     let items = manager.get_items(QuickAccess::RecentFiles).await.unwrap();
///     println!("Recent files: {:?}", items);
/// }
/// ```
pub struct QuickAccessManager {
    executor: Arc<CachedScriptExecutor>,
    feasibility: OnceCell<FeasibilityStatus>,
    lock_timeout: Duration,
}

#[derive(Debug)]
enum Operation {
    Add(PSScript),
    Remove(PSScript),
}

impl QuickAccessManager {
    /// Initializes new Quick Access manager with default configuration
    ///
    /// # Errors
    /// Returns `WincentError` if script executor initialization fails
    pub async fn new() -> WincentResult<Self> {
        Ok(Self {
            executor: Arc::new(CachedScriptExecutor::new()),
            feasibility: OnceCell::new(),
            lock_timeout: Duration::from_secs(10),
        })
    }

    /// Checks system capability for Quick Access operations
    ///
    /// In most case, no need to check feasibility
    ///
    /// # Returns
    /// Tuple (query_feasible, handle_feasible) indicating:
    /// - Ability to query Quick Access items
    /// - Ability to modify Quick Access items
    pub async fn check_feasible(&self) -> (bool, bool) {
        let status = self
            .feasibility
            .get_or_init(|| FeasibilityStatus::check(&self.executor, self.lock_timeout))
            .await;

        (status.query, status.handle)
    }

    fn map_to_script_type(&self, qa_type: QuickAccess) -> WincentResult<PSScript> {
        match qa_type {
            QuickAccess::All => Ok(PSScript::QueryQuickAccess),
            QuickAccess::RecentFiles => Ok(PSScript::QueryRecentFile),
            QuickAccess::FrequentFolders => Ok(PSScript::QueryFrequentFolder),
        }
    }

    async fn handle_operation(
        &self,
        operation: Operation,
        path: &str,
        qa_type: QuickAccess,
        path_type: PathType,
        force_update: bool,
    ) -> WincentResult<()> {
        validate_path(path, path_type)?;

        let script = match operation {
            Operation::Add(script) => script,
            Operation::Remove(script) => script,
        };

        let result = match qa_type {
            QuickAccess::RecentFiles => {
                if matches!(operation, Operation::Add(_)) {
                    add_file_to_recent_with_api(path)?;
                    if force_update {
                        let data_files = QuickAccessDataFiles::new()?;
                        data_files.remove_recent_file()?;
                    }
                    Vec::new()
                } else {
                    self.executor
                        .execute_with_timeout(script, Some(path.to_string()), 10)
                        .await?
                }
            }
            QuickAccess::FrequentFolders => {
                self.executor
                    .execute_with_timeout(script, Some(path.to_string()), 10)
                    .await?
            }
            _ => {
                return Err(WincentError::UnsupportedOperation(format!(
                    "Unsupported operation for {:?}",
                    qa_type
                )))
            }
        };

        if !result.is_empty() {
            return Err(WincentError::ScriptFailed(format!(
                "Operation failed for path: {}",
                path
            )));
        }

        self.executor.clear_cache();
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
        let script_type = self.map_to_script_type(qa_type)?;
        self.executor
            .execute_with_timeout(script_type, None, 10)
            .await
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
        let items = self.get_items(qa_type).await?;
        Ok(items.iter().any(|item| item == path))
    }

    /// Adds item to Quick Access
    ///
    /// # Arguments
    ///
    /// * `path` - Path to add
    /// * `qa_type` - Target Quick Access category
    pub async fn add_item(&self, path: &str, qa_type: QuickAccess, force_update: bool) -> WincentResult<()> {
        if self.check_item(path, qa_type.clone()).await? {
            return Err(WincentError::AlreadyExists(path.to_string()));
        }

        let script = match qa_type {
            QuickAccess::RecentFiles => PSScript::AddRecentFile,
            QuickAccess::FrequentFolders => PSScript::PinToFrequentFolder,
            _ => {
                return Err(WincentError::UnsupportedOperation(format!(
                    "Unsupported add operation for {:?}",
                    qa_type
                )))
            }
        };

        let path_type = match qa_type {
            QuickAccess::RecentFiles => PathType::File,
            QuickAccess::FrequentFolders => PathType::Directory,
            _ => unreachable!(),
        };

        self.handle_operation(Operation::Add(script), path, qa_type, path_type, force_update)
            .await
    }

    /// Removes item from Quick Access
    ///
    /// # Arguments
    ///
    /// * `path` - Path to remove
    /// * `qa_type` - Target Quick Access category
    pub async fn remove_item(&self, path: &str, qa_type: QuickAccess) -> WincentResult<()> {
        if !self.check_item(path, qa_type.clone()).await? {
            return Err(WincentError::NotInRecent(path.to_string()));
        }

        let script = match qa_type {
            QuickAccess::RecentFiles => PSScript::RemoveRecentFile,
            QuickAccess::FrequentFolders => PSScript::UnpinFromFrequentFolder,
            _ => {
                return Err(WincentError::UnsupportedOperation(format!(
                    "Unsupported remove operation for {:?}",
                    qa_type
                )))
            }
        };

        let path_type = match qa_type {
            QuickAccess::RecentFiles => PathType::File,
            QuickAccess::FrequentFolders => PathType::Directory,
            _ => unreachable!(),
        };

        self.handle_operation(Operation::Remove(script), path, qa_type, path_type, false)
            .await
    }

    /// Clears Quick Access items
    ///
    /// # Arguments
    ///
    /// * `qa_type` - Target Quick Access category to clear
    pub async fn empty_items(&self, qa_type: QuickAccess) -> WincentResult<()> {
        match qa_type {
            QuickAccess::RecentFiles => {
                empty_recent_files_with_api()?;
            }
            QuickAccess::FrequentFolders => {
                empty_normal_folders_with_jumplist_file()?;
                self.executor
                    .execute_with_timeout(PSScript::EmptyPinnedFolders, None, 10)
                    .await?;
            }
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
        let (query, handle) = manager.check_feasible().await;
        println!("Query feasibility: {}", query);
        println!("Handle feasibility: {}", handle);
        Ok(())
    }

    #[tokio::test]
    async fn test_item_retrieval() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;

        {
            let _ = manager.feasibility.set(FeasibilityStatus {
                handle: true,
                query: true,
            });
            let items = manager.get_items(QuickAccess::All).await?;
            println!("Items with feasibility=true: {}", items.len());
        }

        {
            let _ = manager.feasibility.get();
            let items = manager.get_items(QuickAccess::All).await?;
            println!("Items with feasibility=None: {}", items.len());
        }

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
    async fn test_item_operations() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;
        let temp_file = tempfile::Builder::new()
            .prefix("wincent-test-")
            .suffix(".txt")
            .tempfile()?;
        let file_path = temp_file.path().to_str().unwrap();

        {
            let _ = manager.feasibility.set(FeasibilityStatus {
                handle: true,
                query: true,
            });
            manager
                .add_item(file_path, QuickAccess::RecentFiles, false)
                .await?;
            manager
                .remove_item(file_path, QuickAccess::RecentFiles)
                .await?;
        }

        {
            let _ = manager.feasibility.get();
            manager
                .add_item(file_path, QuickAccess::RecentFiles, false)
                .await?;
            manager
                .remove_item(file_path, QuickAccess::RecentFiles)
                .await?;
        }

        Ok(())
    }

    #[tokio::test]
    #[ignore = "Modifies system state"]
    async fn test_empty_operations() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;

        {
            let _ = manager.feasibility.set(FeasibilityStatus {
                handle: true,
                query: true,
            });
            manager.empty_items(QuickAccess::RecentFiles).await?;
        }

        {
            let _ = manager.feasibility.get();
            manager.empty_items(QuickAccess::FrequentFolders).await?;
        }

        Ok(())
    }
}
