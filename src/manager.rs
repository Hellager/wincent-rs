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
    empty::{empty_frequent_folders, empty_recent_files_with_api},
    error::WincentError,
    handle::add_file_to_recent_with_api,
    retry::RetryPolicy,
    script_executor::{CachedScriptExecutor, QuickAccessDataFiles},
    script_strategy::PSScript,
    utils::{validate_path, PathType},
    QuickAccess, WincentResult,
};
use std::sync::Arc;
use tokio::sync::OnceCell;
use tokio::time::Duration;

/// Result of a batch operation containing succeeded and failed items
///
/// # Example
/// ```rust,no_run
/// use wincent::prelude::*;
///
/// #[tokio::main]
/// async fn main() -> WincentResult<()> {
///     let manager = QuickAccessManager::new().await?;
///
///     let items = vec![
///         ("C:\\file1.txt".to_string(), QuickAccess::RecentFiles),
///         ("C:\\file2.txt".to_string(), QuickAccess::RecentFiles),
///     ];
///
///     let result = manager.add_items_batch(&items, true).await?;
///
///     println!("Success rate: {:.1}%", result.success_rate() * 100.0);
///
///     if !result.is_complete_success() {
///         for (path, error) in &result.failed {
///             eprintln!("Failed: {} - {}", path, error);
///         }
///     }
///
///     Ok(())
/// }
/// ```
#[derive(Debug)]
pub struct BatchResult {
    /// Successfully processed items
    pub succeeded: Vec<String>,
    /// Failed items with error details
    pub failed: Vec<(String, WincentError)>,
}

impl BatchResult {
    /// Returns true if all operations succeeded
    pub fn is_complete_success(&self) -> bool {
        self.failed.is_empty()
    }

    /// Returns true if at least one operation succeeded
    pub fn has_partial_success(&self) -> bool {
        !self.succeeded.is_empty()
    }

    /// Returns the success rate (0.0 to 1.0)
    pub fn success_rate(&self) -> f64 {
        let total = self.succeeded.len() + self.failed.len();
        if total == 0 {
            return 1.0;
        }
        self.succeeded.len() as f64 / total as f64
    }

    /// Returns the total number of operations
    pub fn total(&self) -> usize {
        self.succeeded.len() + self.failed.len()
    }
}

/// Represents system capability status for Quick Access operations
#[derive(Debug)]
struct FeasibilityStatus {
    handle: bool,
    query: bool,
}

impl FeasibilityStatus {
    async fn check(
        executor: &Arc<CachedScriptExecutor>,
        timeout_duration: Duration,
        retry_policy: &RetryPolicy,
    ) -> Self {
        let query_feasible = Self::check_feasibility(
            executor,
            PSScript::CheckQueryFeasible,
            timeout_duration,
            retry_policy,
        )
        .await;

        let handle_feasible = Self::check_feasibility(
            executor,
            PSScript::CheckPinUnpinFeasible,
            timeout_duration,
            retry_policy,
        )
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
        retry_policy: &RetryPolicy,
    ) -> bool {
        // Use retry mechanism for feasibility checks
        let check_result = executor
            .execute_with_retry_and_timeout(script, None, timeout_duration, retry_policy)
            .await;

        match check_result {
            Ok(_) => true,
            Err(_) => false, // Execution failed (after retries)
        }
    }
}

/// Builder for configuring QuickAccessManager
///
/// # Example
/// ```rust,no_run
/// use wincent::prelude::*;
/// use std::time::Duration;
///
/// #[tokio::main]
/// async fn main() -> WincentResult<()> {
///     let manager = QuickAccessManager::builder()
///         .timeout(Duration::from_secs(30))
///         .check_feasibility_on_init()
///         .build()
///         .await?;
///     Ok(())
/// }
/// ```
pub struct QuickAccessManagerBuilder {
    timeout: Duration,
    cache_enabled: bool,
    executor: Option<Arc<CachedScriptExecutor>>,
    feasibility_check_on_init: bool,
    retry_policy: RetryPolicy,
}

impl Default for QuickAccessManagerBuilder {
    fn default() -> Self {
        Self {
            timeout: Duration::from_secs(10),
            cache_enabled: true,
            executor: None,
            feasibility_check_on_init: false,
            retry_policy: RetryPolicy::default(),
        }
    }
}

impl QuickAccessManagerBuilder {
    /// Sets the timeout for script execution
    ///
    /// # Arguments
    /// * `duration` - Timeout duration (must be > 0)
    ///
    /// # Panics
    /// Panics if duration is zero
    ///
    /// # Example
    /// ```rust,no_run
    /// use wincent::prelude::*;
    /// use std::time::Duration;
    ///
    /// # async fn example() -> WincentResult<()> {
    /// let manager = QuickAccessManager::builder()
    ///     .timeout(Duration::from_secs(30))
    ///     .build()
    ///     .await?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn timeout(mut self, duration: Duration) -> Self {
        assert!(!duration.is_zero(), "Timeout must be greater than zero");
        self.timeout = duration;
        self
    }

    /// Disables query result caching
    ///
    /// Useful for scenarios requiring real-time data without caching.
    ///
    /// # Example
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # async fn example() -> WincentResult<()> {
    /// let manager = QuickAccessManager::builder()
    ///     .disable_cache()
    ///     .build()
    ///     .await?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn disable_cache(mut self) -> Self {
        self.cache_enabled = false;
        self
    }

    /// Provides a custom script executor (mainly for testing)
    ///
    /// # Arguments
    /// * `executor` - Custom CachedScriptExecutor instance
    ///
    /// # Example
    /// ```rust,no_run
    /// use wincent::prelude::*;
    /// use wincent::script_executor::CachedScriptExecutor;
    /// use std::sync::Arc;
    ///
    /// # async fn example() -> WincentResult<()> {
    /// let custom_executor = Arc::new(CachedScriptExecutor::new());
    /// let manager = QuickAccessManager::builder()
    ///     .executor(custom_executor)
    ///     .build()
    ///     .await?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn executor(mut self, executor: Arc<CachedScriptExecutor>) -> Self {
        self.executor = Some(executor);
        self
    }

    /// Enables feasibility check during initialization
    ///
    /// This will verify system capabilities before returning the manager.
    /// If the system does not support query operations, an error will be returned.
    ///
    /// # Example
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # async fn example() -> WincentResult<()> {
    /// let manager = QuickAccessManager::builder()
    ///     .check_feasibility_on_init()
    ///     .build()
    ///     .await?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn check_feasibility_on_init(mut self) -> Self {
        self.feasibility_check_on_init = true;
        self
    }

    /// Sets the retry policy for transient error handling
    ///
    /// # Arguments
    /// * `policy` - Retry policy configuration
    ///
    /// # Example
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # async fn example() -> WincentResult<()> {
    /// let manager = QuickAccessManager::builder()
    ///     .retry_policy(RetryPolicy::aggressive())
    ///     .build()
    ///     .await?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn retry_policy(mut self, policy: RetryPolicy) -> Self {
        self.retry_policy = policy;
        self
    }

    /// Disables retry mechanism
    ///
    /// # Example
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # async fn example() -> WincentResult<()> {
    /// let manager = QuickAccessManager::builder()
    ///     .no_retry()
    ///     .build()
    ///     .await?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn no_retry(mut self) -> Self {
        self.retry_policy = RetryPolicy::no_retry();
        self
    }

    /// Uses fast retry policy (2 retries, short delays)
    ///
    /// # Example
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # async fn example() -> WincentResult<()> {
    /// let manager = QuickAccessManager::builder()
    ///     .fast_retry()
    ///     .build()
    ///     .await?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn fast_retry(mut self) -> Self {
        self.retry_policy = RetryPolicy::fast();
        self
    }

    /// Uses aggressive retry policy (5 retries, longer delays)
    ///
    /// # Example
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # async fn example() -> WincentResult<()> {
    /// let manager = QuickAccessManager::builder()
    ///     .aggressive_retry()
    ///     .build()
    ///     .await?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn aggressive_retry(mut self) -> Self {
        self.retry_policy = RetryPolicy::aggressive();
        self
    }

    /// Builds the QuickAccessManager instance
    ///
    /// # Returns
    ///
    /// Returns a configured QuickAccessManager instance.
    ///
    /// # Errors
    ///
    /// Returns an error if:
    /// - Feasibility check is enabled and the system does not support query operations
    ///
    /// # Example
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # async fn example() -> WincentResult<()> {
    /// let manager = QuickAccessManager::builder()
    ///     .build()
    ///     .await?;
    /// # Ok(())
    /// # }
    /// ```
    pub async fn build(self) -> WincentResult<QuickAccessManager> {
        let executor = self
            .executor
            .unwrap_or_else(|| Arc::new(CachedScriptExecutor::new()));

        if !self.cache_enabled {
            executor.clear_cache();
        }

        let manager = QuickAccessManager {
            executor,
            feasibility: Arc::new(OnceCell::new()),
            lock_timeout: self.timeout,
            retry_policy: self.retry_policy,
        };

        if self.feasibility_check_on_init {
            let (can_query, _can_modify) = manager.check_feasible().await;
            if !can_query {
                return Err(WincentError::SystemError(
                    "System does not support Quick Access query operations".to_string(),
                ));
            }
        }

        Ok(manager)
    }
}

/// Windows Quick Access management system
///
/// # Example
/// ```rust,no_run
/// use wincent::prelude::*;
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
    feasibility: Arc<OnceCell<FeasibilityStatus>>,
    lock_timeout: Duration,
    retry_policy: RetryPolicy,
}

#[derive(Debug)]
enum Operation {
    Add(PSScript),
    Remove(PSScript),
}

impl QuickAccessManager {
    /// Creates a new builder for QuickAccessManager
    ///
    /// This is the recommended way to create a QuickAccessManager with custom configuration.
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    /// use std::time::Duration;
    ///
    /// #[tokio::main]
    /// async fn main() -> WincentResult<()> {
    ///     let manager = QuickAccessManager::builder()
    ///         .timeout(Duration::from_secs(30))
    ///         .check_feasibility_on_init()
    ///         .build()
    ///         .await?;
    ///     Ok(())
    /// }
    /// ```
    pub fn builder() -> QuickAccessManagerBuilder {
        QuickAccessManagerBuilder::default()
    }

    /// Initializes new Quick Access manager with default configuration
    ///
    /// This is equivalent to `QuickAccessManager::builder().build().await`.
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// #[tokio::main]
    /// async fn main() -> WincentResult<()> {
    ///     let manager = QuickAccessManager::new().await?;
    ///     Ok(())
    /// }
    /// ```
    pub async fn new() -> WincentResult<Self> {
        Self::builder().build().await
    }

    /// Checks system capability for Quick Access operations
    ///
    /// In most case, this is not needed
    ///
    /// # Arguments
    ///
    /// None
    ///
    /// # Returns
    ///
    /// Tuple (query_feasible, handle_feasible) indicating:
    /// - query_feasible: Ability to query Quick Access items
    /// - handle_feasible: Ability to modify Quick Access items
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// #[tokio::main]
    /// async fn main() -> WincentResult<()> {
    ///     let manager = QuickAccessManager::new().await?;
    ///     let (can_query, can_modify) = manager.check_feasible().await;
    ///     println!("Can query: {}, Can modify: {}", can_query, can_modify);
    ///     Ok(())
    /// }
    /// ```
    pub async fn check_feasible(&self) -> (bool, bool) {
        // Use get_or_init for concurrent-safe single initialization
        let status = self
            .feasibility
            .get_or_init(|| async {
                FeasibilityStatus::check(&self.executor, self.lock_timeout, &self.retry_policy).await
            })
            .await;

        (status.query, status.handle)
    }

    /// Gets the current retry policy
    ///
    /// # Example
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # async fn example() -> WincentResult<()> {
    /// let manager = QuickAccessManager::new().await?;
    /// let policy = manager.retry_policy();
    /// println!("Max attempts: {}", policy.max_attempts);
    /// # Ok(())
    /// # }
    /// ```
    pub fn retry_policy(&self) -> &RetryPolicy {
        &self.retry_policy
    }

    /// Sets a new retry policy
    ///
    /// This will invalidate any cached feasibility check results, so the next
    /// call to `check_feasible()` will re-run the checks with the new policy.
    ///
    /// # Example
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # async fn example() -> WincentResult<()> {
    /// let mut manager = QuickAccessManager::new().await?;
    /// manager.set_retry_policy(RetryPolicy::aggressive());
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_retry_policy(&mut self, policy: RetryPolicy) {
        self.retry_policy = policy;
        // Invalidate cached feasibility check results by replacing the entire Arc
        self.feasibility = Arc::new(OnceCell::new());
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
                    // Add recent file may not show in the explorer recent files list
                    // But it did will show in windows recent folder
                    // So if we did need it show, we need force update the list
                    if force_update {
                        let data_files = QuickAccessDataFiles::new()?;
                        data_files.remove_recent_file()?;
                    }
                    Vec::new()
                } else {
                    self.executor
                        .execute_with_retry_and_timeout(
                            script,
                            Some(path.to_string()),
                            self.lock_timeout,
                            &self.retry_policy,
                        )
                        .await?
                }
            }
            QuickAccess::FrequentFolders => {
                self.executor
                    .execute_with_retry_and_timeout(
                        script,
                        Some(path.to_string()),
                        self.lock_timeout,
                        &self.retry_policy,
                    )
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
    /// * `qa_type` - Target Quick Access category (Recent Files, Frequent Folders, or All)
    ///
    /// # Returns
    ///
    /// List of path strings representing items in the specified category
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// #[tokio::main]
    /// async fn main() -> WincentResult<()> {
    ///     let manager = QuickAccessManager::new().await?;
    ///     
    ///     // Get all Quick Access items
    ///     let all_items = manager.get_items(QuickAccess::All).await?;
    ///     
    ///     // Get only recent files
    ///     let recent_files = manager.get_items(QuickAccess::RecentFiles).await?;
    ///     
    ///     Ok(())
    /// }
    /// ```
    pub async fn get_items(&self, qa_type: QuickAccess) -> WincentResult<Vec<String>> {
        let script_type = self.map_to_script_type(qa_type)?;
        self.executor
            .execute_with_retry_and_timeout(
                script_type,
                None,
                self.lock_timeout,
                &self.retry_policy,
            )
            .await
    }

    /// Checks if an item exists in Quick Access with exact path matching
    ///
    /// This method performs exact path comparison. For partial/fuzzy matching,
    /// use `contains_item()` instead.
    ///
    /// # Arguments
    ///
    /// * `path` - Exact path to check
    /// * `qa_type` - Quick Access category to search (Recent Files, Frequent Folders, or All)
    ///
    /// # Returns
    ///
    /// `true` if the exact path exists in the specified category, `false` otherwise
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// #[tokio::main]
    /// async fn main() -> WincentResult<()> {
    ///     let manager = QuickAccessManager::new().await?;
    ///
    ///     // Exact match - must match full path
    ///     let exists = manager.check_item_exact(
    ///         "C:\\Users\\Documents\\file.txt",
    ///         QuickAccess::RecentFiles
    ///     ).await?;
    ///
    ///     println!("File exists: {}", exists);
    ///     Ok(())
    /// }
    /// ```
    pub async fn check_item_exact(&self, path: &str, qa_type: QuickAccess) -> WincentResult<bool> {
        let items = self.get_items(qa_type).await?;
        Ok(items.iter().any(|item| item == path))
    }

    /// Checks if any item in Quick Access contains the given keyword
    ///
    /// This method performs substring matching. For exact path matching,
    /// use `check_item_exact()` instead.
    ///
    /// # Arguments
    ///
    /// * `keyword` - Keyword or partial path to search for
    /// * `qa_type` - Quick Access category to search (Recent Files, Frequent Folders, or All)
    ///
    /// # Returns
    ///
    /// `true` if any item contains the keyword, `false` otherwise
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// #[tokio::main]
    /// async fn main() -> WincentResult<()> {
    ///     let manager = QuickAccessManager::new().await?;
    ///
    ///     // Fuzzy match - matches any item containing "Documents"
    ///     let exists = manager.contains_item(
    ///         "Documents",
    ///         QuickAccess::RecentFiles
    ///     ).await?;
    ///
    ///     // This will match paths like:
    ///     // - "C:\\Users\\Documents\\file.txt"
    ///     // - "D:\\My Documents\\report.pdf"
    ///
    ///     println!("Found items containing 'Documents': {}", exists);
    ///     Ok(())
    /// }
    /// ```
    pub async fn contains_item(&self, keyword: &str, qa_type: QuickAccess) -> WincentResult<bool> {
        let items = self.get_items(qa_type).await?;
        Ok(items.iter().any(|item| item.contains(keyword)))
    }

    /// Checks item presence in Quick Access (exact match)
    ///
    /// **Deprecated**: Use `check_item_exact()` for clarity. This method will be
    /// removed in v0.2.0.
    ///
    /// # Arguments
    ///
    /// * `path` - Target path to check
    /// * `qa_type` - Quick Access category to search (Recent Files, Frequent Folders, or All)
    ///
    /// # Returns
    ///
    /// `true` if item exists in the specified category, `false` otherwise
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// #[tokio::main]
    /// async fn main() -> WincentResult<()> {
    ///     let manager = QuickAccessManager::new().await?;
    ///     
    ///     let exists = manager.check_item(
    ///         "C:\\path\\to\\file.txt",
    ///         QuickAccess::RecentFiles
    ///     ).await?;
    ///     
    ///     println!("File exists in Recent Files: {}", exists);
    ///     Ok(())
    /// }
    /// ```
    #[deprecated(
        since = "0.1.3",
        note = "Use `check_item_exact()` for exact matching or `contains_item()` for fuzzy matching. This method will be removed or redefined in v0.2.0."
    )]
    pub async fn check_item(&self, path: &str, qa_type: QuickAccess) -> WincentResult<bool> {
        self.check_item_exact(path, qa_type).await
    }

    /// Adds an item to Quick Access
    ///
    /// # Arguments
    ///
    /// * `path` - Path to add to Quick Access
    /// * `qa_type` - Target Quick Access category (Recent Files or Frequent Folders)
    /// * `force_update` - Whether to force update Explorer display
    ///     - For Recent Files, setting to true will force update the Explorer display list
    ///     - For Frequent Folders, this parameter has no effect
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// #[tokio::main]
    /// async fn main() -> WincentResult<()> {
    ///     let manager = QuickAccessManager::new().await?;
    ///     
    ///     // Add file to Recent Files
    ///     manager.add_item("C:\\path\\to\\file.txt", QuickAccess::RecentFiles, true).await?;
    ///     
    ///     // Add folder to Frequent Folders
    ///     manager.add_item("C:\\path\\to\\folder", QuickAccess::FrequentFolders, false).await?;
    ///     
    ///     Ok(())
    /// }
    /// ```
    pub async fn add_item(
        &self,
        path: &str,
        qa_type: QuickAccess,
        force_update: bool,
    ) -> WincentResult<()> {
        if self.check_item_exact(path, qa_type.clone()).await? {
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

        self.handle_operation(
            Operation::Add(script),
            path,
            qa_type,
            path_type,
            force_update,
        )
        .await
    }

    /// Removes item from Quick Access
    ///
    /// # Arguments
    ///
    /// * `path` - Path to remove
    /// * `qa_type` - Target Quick Access category
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// #[tokio::main]
    /// async fn main() -> WincentResult<()> {
    ///     let manager = QuickAccessManager::new().await?;
    ///     
    ///     // Remove file from Recent Files
    ///     manager.remove_item("C:\\path\\to\\file.txt", QuickAccess::RecentFiles).await?;
    ///     
    ///     // Remove folder from Frequent Folders
    ///     manager.remove_item("C:\\path\\to\\folder", QuickAccess::FrequentFolders).await?;
    ///     
    ///     Ok(())
    /// }
    /// ```
    pub async fn remove_item(&self, path: &str, qa_type: QuickAccess) -> WincentResult<()> {
        if !self.check_item_exact(path, qa_type.clone()).await? {
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

    /// Adds multiple items to Quick Access in batch
    ///
    /// This method is more efficient than calling `add_item()` multiple times:
    /// - Reduces PowerShell process overhead
    /// - Refreshes Explorer only once at the end
    /// - Continues on errors and reports all results
    ///
    /// # Arguments
    ///
    /// * `items` - List of (path, QuickAccess type) tuples to add
    /// * `force_update` - Whether to refresh Explorer after all operations
    ///
    /// # Returns
    ///
    /// Returns `BatchResult` containing succeeded and failed items
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// #[tokio::main]
    /// async fn main() -> WincentResult<()> {
    ///     let manager = QuickAccessManager::new().await?;
    ///
    ///     let items = vec![
    ///         ("C:\\file1.txt".to_string(), QuickAccess::RecentFiles),
    ///         ("C:\\file2.txt".to_string(), QuickAccess::RecentFiles),
    ///         ("C:\\folder1".to_string(), QuickAccess::FrequentFolders),
    ///     ];
    ///
    ///     let result = manager.add_items_batch(&items, true).await?;
    ///
    ///     println!("Succeeded: {}", result.succeeded.len());
    ///     println!("Failed: {}", result.failed.len());
    ///     println!("Success rate: {:.1}%", result.success_rate() * 100.0);
    ///
    ///     // Handle failures
    ///     for (path, error) in &result.failed {
    ///         eprintln!("Failed to add {}: {}", path, error);
    ///     }
    ///
    ///     Ok(())
    /// }
    /// ```
    pub async fn add_items_batch(
        &self,
        items: &[(String, QuickAccess)],
        force_update: bool,
    ) -> WincentResult<BatchResult> {
        let mut succeeded = Vec::new();
        let mut failed = Vec::new();

        // Process each item, continuing on errors
        for (path, qa_type) in items {
            match self.add_item_internal(path, qa_type.clone()).await {
                Ok(_) => succeeded.push(path.clone()),
                Err(e) => failed.push((path.clone(), e)),
            }
        }

        // Refresh Explorer once at the end if requested and at least one succeeded
        if force_update && !succeeded.is_empty() {
            if let Err(e) = crate::utils::refresh_explorer_window() {
                // Refresh failure is not critical, just log it
                eprintln!("Warning: Failed to refresh Explorer: {}", e);
            }
        }

        // Clear cache after batch operation
        if !succeeded.is_empty() {
            self.clear_cache();
        }

        Ok(BatchResult { succeeded, failed })
    }

    /// Removes multiple items from Quick Access in batch
    ///
    /// This method is more efficient than calling `remove_item()` multiple times
    /// by processing all items and reporting results together.
    ///
    /// # Arguments
    ///
    /// * `items` - List of (path, QuickAccess type) tuples to remove
    ///
    /// # Returns
    ///
    /// Returns `BatchResult` containing succeeded and failed items
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// #[tokio::main]
    /// async fn main() -> WincentResult<()> {
    ///     let manager = QuickAccessManager::new().await?;
    ///
    ///     let items = vec![
    ///         ("C:\\file1.txt".to_string(), QuickAccess::RecentFiles),
    ///         ("C:\\file2.txt".to_string(), QuickAccess::RecentFiles),
    ///     ];
    ///
    ///     let result = manager.remove_items_batch(&items).await?;
    ///
    ///     if result.is_complete_success() {
    ///         println!("All items removed successfully");
    ///     } else {
    ///         println!("Partial success: {}/{} removed",
    ///             result.succeeded.len(),
    ///             items.len()
    ///         );
    ///     }
    ///
    ///     Ok(())
    /// }
    /// ```
    pub async fn remove_items_batch(
        &self,
        items: &[(String, QuickAccess)],
    ) -> WincentResult<BatchResult> {
        let mut succeeded = Vec::new();
        let mut failed = Vec::new();

        // Process each item, continuing on errors
        for (path, qa_type) in items {
            match self.remove_item_internal(path, qa_type.clone()).await {
                Ok(_) => succeeded.push(path.clone()),
                Err(e) => failed.push((path.clone(), e)),
            }
        }

        // Clear cache after batch operation
        if !succeeded.is_empty() {
            self.clear_cache();
        }

        Ok(BatchResult { succeeded, failed })
    }

    /// Internal method to add item without refreshing Explorer
    async fn add_item_internal(&self, path: &str, qa_type: QuickAccess) -> WincentResult<()> {
        // Validate qa_type first before any operations
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

        if self.check_item_exact(path, qa_type.clone()).await? {
            return Err(WincentError::AlreadyExists(path.to_string()));
        }

        let path_type = match qa_type {
            QuickAccess::RecentFiles => PathType::File,
            QuickAccess::FrequentFolders => PathType::Directory,
            _ => unreachable!(),
        };

        // Always pass false for force_update in internal method
        self.handle_operation(Operation::Add(script), path, qa_type, path_type, false)
            .await
    }

    /// Internal method to remove item
    async fn remove_item_internal(&self, path: &str, qa_type: QuickAccess) -> WincentResult<()> {
        // Validate qa_type first before any operations
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

        if !self.check_item_exact(path, qa_type.clone()).await? {
            return Err(WincentError::NotInRecent(path.to_string()));
        }

        let path_type = match qa_type {
            QuickAccess::RecentFiles => PathType::File,
            QuickAccess::FrequentFolders => PathType::Directory,
            _ => unreachable!(),
        };

        self.handle_operation(Operation::Remove(script), path, qa_type, path_type, false)
            .await
    }

    async fn empty_all_items_internal(&self, also_system_default: bool) -> WincentResult<()> {
        if let Err(source) = Box::pin(self.empty_items_internal(
            QuickAccess::RecentFiles,
            also_system_default,
        ))
        .await
        {
            return Err(WincentError::PartialEmpty {
                recent_files_cleared: false,
                frequent_folders_cleared: false,
                source: Box::new(source),
            });
        }

        if let Err(source) = Box::pin(self.empty_items_internal(
            QuickAccess::FrequentFolders,
            also_system_default,
        ))
        .await
        {
            return Err(WincentError::PartialEmpty {
                recent_files_cleared: true,
                frequent_folders_cleared: false,
                source: Box::new(source),
            });
        }

        Ok(())
    }

    async fn refresh_empty_items(&self, force_refresh: bool) -> WincentResult<()> {
        if force_refresh {
            self.executor
                .execute(PSScript::RefreshExplorer, None)
                .await?;
        }
        Ok(())
    }

    fn should_clear_cache_after_empty(result: &WincentResult<()>) -> bool {
        match result {
            Ok(()) => true,
            Err(WincentError::PartialEmpty {
                recent_files_cleared,
                frequent_folders_cleared,
                ..
            }) => *recent_files_cleared || *frequent_folders_cleared,
            Err(_) => false,
        }
    }

    async fn empty_items_internal(
        &self,
        qa_type: QuickAccess,
        also_system_default: bool,
    ) -> WincentResult<()> {
        match qa_type {
            QuickAccess::RecentFiles => {
                empty_recent_files_with_api()?;
            }
            QuickAccess::FrequentFolders => {
                empty_frequent_folders(also_system_default)?;
                // Regular frequent folders and pinned folders are backed by different behaviors in
                // Windows Quick Access. The helper above clears regular entries, while pinned
                // folders still require a dedicated PowerShell cleanup step.
                self.executor
                    .execute_with_retry_and_timeout(
                        PSScript::EmptyPinnedFolders,
                        None,
                        self.lock_timeout,
                        &self.retry_policy,
                    )
                    .await?;
            }
            QuickAccess::All => {
                self.empty_all_items_internal(also_system_default).await?;
            }
        }

        Ok(())
    }

    /// Clears Quick Access items
    ///
    /// Frequent folder cleanup is a two-step process:
    /// 1. `empty_frequent_folders()` clears regular frequent folders.
    /// 2. `PSScript::EmptyPinnedFolders` clears folders that were pinned in Quick Access.
    ///
    /// When clearing `QuickAccess::All`, the operation runs recent-files cleanup first and then
    /// frequent-folders cleanup. If the second step fails after the first one has already
    /// succeeded, this method returns `WincentError::PartialEmpty` so callers can detect the
    /// partial-success state.
    ///
    /// # Arguments
    ///
    /// * `qa_type` - Target Quick Access category to clear
    /// * `force_refresh` - Whether to force refresh Explorer after clearing
    /// * `also_system_default` - Whether to also clear system default items
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// #[tokio::main]
    /// async fn main() -> WincentResult<()> {
    ///     let manager = QuickAccessManager::new().await?;
    ///     
    ///     // Clear recent files with Explorer refresh
    ///     manager.empty_items(QuickAccess::RecentFiles, true, false).await?;
    ///     
    ///     // Clear all items including system defaults
    ///     manager.empty_items(QuickAccess::All, true, true).await?;
    ///     
    ///     Ok(())
    /// }
    /// ```
    pub async fn empty_items(
        &self,
        qa_type: QuickAccess,
        force_refresh: bool,
        also_system_default: bool,
    ) -> WincentResult<()> {
        let result = self.empty_items_internal(qa_type, also_system_default).await;

        if Self::should_clear_cache_after_empty(&result) {
            self.executor.clear_cache();
        }

        if result.is_ok() {
            self.refresh_empty_items(force_refresh).await?;
        }

        result
    }

    /// Clears internal cache
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// #[tokio::main]
    /// async fn main() -> WincentResult<()> {
    ///     let manager = QuickAccessManager::new().await?;
    ///     manager.clear_cache();
    ///     Ok(())
    /// }
    /// ```
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

        // Pre-populate feasibility cache
        manager.feasibility.set(FeasibilityStatus {
            handle: true,
            query: true,
        }).ok();

        let items = manager.get_items(QuickAccess::All).await?;
        println!("Items with feasibility=true: {}", items.len());

        // Verify cache is populated
        assert!(manager.feasibility.get().is_some());

        let items = manager.get_items(QuickAccess::All).await?;
        println!("Items with feasibility cached: {}", items.len());

        Ok(())
    }

    #[tokio::test]
    async fn test_item_presence_check() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;
        let items = manager.get_items(QuickAccess::All).await?;

        if let Some(item) = items.first() {
            let exists = manager.check_item_exact(item, QuickAccess::All).await?;
            assert!(exists, "Item should exist in collection");
        }

        let non_existent = "Z:\\Invalid\\Path\\Test.txt";
        let exists = manager.check_item_exact(non_existent, QuickAccess::All).await?;
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

        // Pre-populate feasibility cache
        manager.feasibility.set(FeasibilityStatus {
            handle: true,
            query: true,
        }).ok();

        manager
            .add_item(file_path, QuickAccess::RecentFiles, false)
            .await?;
        manager
            .remove_item(file_path, QuickAccess::RecentFiles)
            .await?;

        // Verify cache is still populated
        assert!(manager.feasibility.get().is_some());

        manager
            .add_item(file_path, QuickAccess::RecentFiles, false)
            .await?;
        manager
            .remove_item(file_path, QuickAccess::RecentFiles)
            .await?;

        Ok(())
    }

    #[tokio::test]
    #[ignore = "Modifies system state"]
    async fn test_empty_operations() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;

        // Pre-populate feasibility cache
        manager.feasibility.set(FeasibilityStatus {
            handle: true,
            query: true,
        }).ok();

        manager
            .empty_items(QuickAccess::RecentFiles, false, false)
            .await?;

        // Verify cache is still populated
        assert!(manager.feasibility.get().is_some());

        manager
            .empty_items(QuickAccess::FrequentFolders, false, false)
            .await?;

        Ok(())
    }

    #[tokio::test]
    async fn test_check_item_exact_vs_contains() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;
        let items = manager.get_items(QuickAccess::All).await?;

        if let Some(full_path) = items.first() {
            // exact match with full path should succeed
            let exact = manager
                .check_item_exact(full_path, QuickAccess::All)
                .await?;
            assert!(exact, "check_item_exact should match full path");

            // contains match with full path should also succeed
            let fuzzy = manager.contains_item(full_path, QuickAccess::All).await?;
            assert!(fuzzy, "contains_item should match full path");

            // exact match with partial path should fail
            if full_path.len() > 3 {
                let partial = &full_path[..full_path.len() - 1];
                let exact_partial = manager
                    .check_item_exact(partial, QuickAccess::All)
                    .await?;
                assert!(
                    !exact_partial,
                    "check_item_exact should not match partial path"
                );
            }
        }

        // non-existent path should return false for both
        let non_existent = "Z:\\Invalid\\Path\\Test.txt";
        assert!(
            !manager
                .check_item_exact(non_existent, QuickAccess::All)
                .await?
        );
        assert!(
            !manager
                .contains_item(non_existent, QuickAccess::All)
                .await?
        );

        Ok(())
    }

    // Builder pattern tests
    #[tokio::test]
    async fn test_builder_default() -> WincentResult<()> {
        let manager = QuickAccessManager::builder().build().await?;
        assert_eq!(manager.executor.cache_size(), 0);
        assert_eq!(manager.lock_timeout, Duration::from_secs(10));
        Ok(())
    }

    #[tokio::test]
    async fn test_builder_custom_timeout() -> WincentResult<()> {
        let custom_timeout = Duration::from_secs(30);
        let manager = QuickAccessManager::builder()
            .timeout(custom_timeout)
            .build()
            .await?;
        assert_eq!(manager.lock_timeout, custom_timeout);
        Ok(())
    }

    #[tokio::test]
    async fn test_builder_disable_cache() -> WincentResult<()> {
        let manager = QuickAccessManager::builder()
            .disable_cache()
            .build()
            .await?;

        // Cache should be cleared
        assert_eq!(manager.executor.cache_size(), 0);
        Ok(())
    }

    #[tokio::test]
    async fn test_builder_custom_executor() -> WincentResult<()> {
        let custom_executor = Arc::new(CachedScriptExecutor::new());
        let executor_ptr = Arc::as_ptr(&custom_executor);

        let manager = QuickAccessManager::builder()
            .executor(custom_executor)
            .build()
            .await?;

        // Verify the same executor instance is used
        assert_eq!(Arc::as_ptr(&manager.executor), executor_ptr);
        Ok(())
    }

    #[tokio::test]
    #[ignore = "Modifies system state"]
    async fn test_builder_check_feasibility_on_init() -> WincentResult<()> {
        // This should succeed on a compatible system
        let manager = QuickAccessManager::builder()
            .check_feasibility_on_init()
            .build()
            .await?;

        // Feasibility should already be checked
        let (can_query, _can_modify) = manager.check_feasible().await;
        assert!(can_query, "System should support query operations");
        Ok(())
    }

    #[tokio::test]
    async fn test_builder_combined_options() -> WincentResult<()> {
        let custom_timeout = Duration::from_secs(20);
        let manager = QuickAccessManager::builder()
            .timeout(custom_timeout)
            .disable_cache()
            .build()
            .await?;

        assert_eq!(manager.lock_timeout, custom_timeout);
        assert_eq!(manager.executor.cache_size(), 0);
        Ok(())
    }

    #[tokio::test]
    async fn test_builder_vs_new_equivalence() -> WincentResult<()> {
        let manager_new = QuickAccessManager::new().await?;
        let manager_builder = QuickAccessManager::builder().build().await?;

        // Both should have the same default timeout
        assert_eq!(manager_new.lock_timeout, manager_builder.lock_timeout);
        assert_eq!(manager_new.lock_timeout, Duration::from_secs(10));
        Ok(())
    }

    #[tokio::test]
    #[should_panic(expected = "Timeout must be greater than zero")]
    async fn test_builder_zero_timeout_panics() {
        let _ = QuickAccessManager::builder()
            .timeout(Duration::from_secs(0))
            .build()
            .await;
    }

    #[tokio::test]
    async fn test_builder_method_chaining() -> WincentResult<()> {
        // Test that builder methods can be chained fluently
        let manager = QuickAccessManager::builder()
            .timeout(Duration::from_secs(15))
            .disable_cache()
            .timeout(Duration::from_secs(25)) // Override previous timeout
            .build()
            .await?;

        assert_eq!(manager.lock_timeout, Duration::from_secs(25));
        Ok(())
    }

    #[tokio::test]
    async fn test_builder_with_actual_query() -> WincentResult<()> {
        let manager = QuickAccessManager::builder()
            .timeout(Duration::from_secs(15))
            .build()
            .await?;

        // Verify the manager actually works
        let _items = manager.get_items(QuickAccess::All).await?;
        // If we get here without error, the query succeeded
        Ok(())
    }

    // Batch operations tests
    #[tokio::test]
    async fn test_batch_result_methods() {
        let result = BatchResult {
            succeeded: vec!["file1.txt".to_string(), "file2.txt".to_string()],
            failed: vec![(
                "file3.txt".to_string(),
                WincentError::NotInRecent("file3.txt".to_string()),
            )],
        };

        assert_eq!(result.total(), 3);
        assert_eq!(result.succeeded.len(), 2);
        assert_eq!(result.failed.len(), 1);
        assert!(!result.is_complete_success());
        assert!(result.has_partial_success());
        assert!((result.success_rate() - 0.666).abs() < 0.01);
    }

    #[tokio::test]
    async fn test_batch_result_complete_success() {
        let result = BatchResult {
            succeeded: vec!["file1.txt".to_string(), "file2.txt".to_string()],
            failed: vec![],
        };

        assert!(result.is_complete_success());
        assert!(result.has_partial_success());
        assert_eq!(result.success_rate(), 1.0);
    }

    #[tokio::test]
    async fn test_batch_result_complete_failure() {
        let result = BatchResult {
            succeeded: vec![],
            failed: vec![
                (
                    "file1.txt".to_string(),
                    WincentError::NotInRecent("file1.txt".to_string()),
                ),
                (
                    "file2.txt".to_string(),
                    WincentError::NotInRecent("file2.txt".to_string()),
                ),
            ],
        };

        assert!(!result.is_complete_success());
        assert!(!result.has_partial_success());
        assert_eq!(result.success_rate(), 0.0);
    }

    #[tokio::test]
    async fn test_batch_result_empty() {
        let result = BatchResult {
            succeeded: vec![],
            failed: vec![],
        };

        assert!(result.is_complete_success());
        assert!(!result.has_partial_success());
        assert_eq!(result.success_rate(), 1.0);
        assert_eq!(result.total(), 0);
    }

    #[tokio::test]
    async fn test_add_items_batch_empty() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;
        let items: Vec<(String, QuickAccess)> = vec![];

        let result = manager.add_items_batch(&items, false).await?;

        assert!(result.is_complete_success());
        assert_eq!(result.total(), 0);
        Ok(())
    }

    #[tokio::test]
    async fn test_remove_items_batch_empty() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;
        let items: Vec<(String, QuickAccess)> = vec![];

        let result = manager.remove_items_batch(&items).await?;

        assert!(result.is_complete_success());
        assert_eq!(result.total(), 0);
        Ok(())
    }

    #[tokio::test]
    async fn test_add_items_batch_with_invalid_type() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;
        let items = vec![
            ("C:\\test.txt".to_string(), QuickAccess::All), // Invalid type
        ];

        let result = manager.add_items_batch(&items, false).await?;

        assert!(!result.is_complete_success());
        assert_eq!(result.succeeded.len(), 0);
        assert_eq!(result.failed.len(), 1);
        assert!(matches!(
            result.failed[0].1,
            WincentError::UnsupportedOperation(_)
        ));
        Ok(())
    }

    #[tokio::test]
    async fn test_remove_items_batch_with_invalid_type() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;
        let items = vec![
            ("C:\\test.txt".to_string(), QuickAccess::All), // Invalid type
        ];

        let result = manager.remove_items_batch(&items).await?;

        assert!(!result.is_complete_success());
        assert_eq!(result.succeeded.len(), 0);
        assert_eq!(result.failed.len(), 1);
        assert!(matches!(
            result.failed[0].1,
            WincentError::UnsupportedOperation(_)
        ));
        Ok(())
    }

    #[tokio::test]
    #[ignore = "Modifies system state"]
    async fn test_add_items_batch_actual() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;

        // Create temporary files
        let temp_file1 = tempfile::Builder::new()
            .prefix("wincent-batch-test-1-")
            .suffix(".txt")
            .tempfile()?;
        let temp_file2 = tempfile::Builder::new()
            .prefix("wincent-batch-test-2-")
            .suffix(".txt")
            .tempfile()?;

        let file1_path = temp_file1.path().to_str().unwrap().to_string();
        let file2_path = temp_file2.path().to_str().unwrap().to_string();

        let items = vec![
            (file1_path.clone(), QuickAccess::RecentFiles),
            (file2_path.clone(), QuickAccess::RecentFiles),
        ];

        // Add items in batch
        let result = manager.add_items_batch(&items, false).await?;

        assert_eq!(result.succeeded.len(), 2);
        assert_eq!(result.failed.len(), 0);
        assert!(result.is_complete_success());

        // Verify items were added
        assert!(
            manager
                .check_item_exact(&file1_path, QuickAccess::RecentFiles)
                .await?
        );
        assert!(
            manager
                .check_item_exact(&file2_path, QuickAccess::RecentFiles)
                .await?
        );

        // Clean up
        let _ = manager.remove_items_batch(&items).await?;

        Ok(())
    }

    #[tokio::test]
    #[ignore = "Modifies system state"]
    async fn test_remove_items_batch_actual() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;

        // Create temporary files
        let temp_file1 = tempfile::Builder::new()
            .prefix("wincent-batch-test-1-")
            .suffix(".txt")
            .tempfile()?;
        let temp_file2 = tempfile::Builder::new()
            .prefix("wincent-batch-test-2-")
            .suffix(".txt")
            .tempfile()?;

        let file1_path = temp_file1.path().to_str().unwrap().to_string();
        let file2_path = temp_file2.path().to_str().unwrap().to_string();

        // Add items first
        manager
            .add_item(&file1_path, QuickAccess::RecentFiles, false)
            .await?;
        manager
            .add_item(&file2_path, QuickAccess::RecentFiles, false)
            .await?;

        let items = vec![
            (file1_path.clone(), QuickAccess::RecentFiles),
            (file2_path.clone(), QuickAccess::RecentFiles),
        ];

        // Remove items in batch
        let result = manager.remove_items_batch(&items).await?;

        assert_eq!(result.succeeded.len(), 2);
        assert_eq!(result.failed.len(), 0);
        assert!(result.is_complete_success());

        // Verify items were removed
        assert!(
            !manager
                .check_item_exact(&file1_path, QuickAccess::RecentFiles)
                .await?
        );
        assert!(
            !manager
                .check_item_exact(&file2_path, QuickAccess::RecentFiles)
                .await?
        );

        Ok(())
    }

    #[tokio::test]
    #[ignore = "Modifies system state"]
    async fn test_batch_operations_partial_failure() -> WincentResult<()> {
        let manager = QuickAccessManager::new().await?;

        // Create one valid temporary file
        let temp_file = tempfile::Builder::new()
            .prefix("wincent-batch-test-")
            .suffix(".txt")
            .tempfile()?;
        let valid_path = temp_file.path().to_str().unwrap().to_string();

        let items = vec![
            (valid_path.clone(), QuickAccess::RecentFiles),
            (
                "Z:\\NonExistent\\File.txt".to_string(),
                QuickAccess::RecentFiles,
            ),
        ];

        // Add items - one should succeed, one should fail
        let result = manager.add_items_batch(&items, false).await?;

        assert!(result.has_partial_success());
        assert!(!result.is_complete_success());
        assert_eq!(result.succeeded.len(), 1);
        assert_eq!(result.failed.len(), 1);
        assert!((result.success_rate() - 0.5).abs() < 0.01);

        // Clean up the successful one
        let _ = manager
            .remove_item(&valid_path, QuickAccess::RecentFiles)
            .await;

        Ok(())
    }
}
