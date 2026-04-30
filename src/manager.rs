//! Synchronous facade for Windows Quick Access operations.

use crate::{
    batch::{self, BatchOptions, BatchResult},
    empty::{self, EmptyOptions},
    error::WincentError,
    handle::{
        add_to_frequent_folders_with_timeout, add_to_recent_files_with_options,
        remove_from_frequent_folders_with_timeout, remove_from_recent_files_with_timeout,
        AddRecentFileOptions,
    },
    query,
    utils::paths_equal,
    QuickAccess, WincentResult,
};
use std::time::Duration;

/// Builder for configuring [`QuickAccessManager`].
pub struct QuickAccessManagerBuilder {
    timeout: Duration,
}

impl Default for QuickAccessManagerBuilder {
    fn default() -> Self {
        Self {
            timeout: Duration::from_secs(10),
        }
    }
}

impl QuickAccessManagerBuilder {
    /// Sets the timeout for shell operations that support timeout control.
    ///
    /// # Panics
    ///
    /// Panics if `duration` is zero.
    pub fn timeout(mut self, duration: Duration) -> Self {
        assert!(!duration.is_zero(), "Timeout must be greater than zero");
        self.timeout = duration;
        self
    }

    /// Builds a configured [`QuickAccessManager`].
    pub fn build(self) -> QuickAccessManager {
        QuickAccessManager {
            timeout: self.timeout,
        }
    }
}

/// Windows Quick Access manager.
///
/// This type is a thin synchronous facade over the `query`, `handle`, `empty`,
/// and `batch` modules.
pub struct QuickAccessManager {
    timeout: Duration,
}

impl Default for QuickAccessManager {
    fn default() -> Self {
        Self::new()
    }
}

impl QuickAccessManager {
    /// Creates a new builder for [`QuickAccessManager`].
    pub fn builder() -> QuickAccessManagerBuilder {
        QuickAccessManagerBuilder::default()
    }

    /// Creates a manager with default configuration.
    pub fn new() -> Self {
        Self::builder().build()
    }

    /// Returns the configured shell operation timeout.
    pub fn timeout(&self) -> Duration {
        self.timeout
    }

    /// Retrieves Quick Access items.
    pub fn get_items(&self, qa_type: QuickAccess) -> WincentResult<Vec<String>> {
        match qa_type {
            QuickAccess::RecentFiles => query::get_recent_files(),
            QuickAccess::FrequentFolders => query::get_frequent_folders(),
            QuickAccess::All => query::get_quick_access_items(),
        }
    }

    /// Checks if an item exists in Quick Access using Windows path semantics.
    pub fn check_item_exact(&self, path: &str, qa_type: QuickAccess) -> WincentResult<bool> {
        let items = self.get_items(qa_type)?;
        Ok(items.iter().any(|item| paths_equal(item, path)))
    }

    /// Checks if any item in Quick Access contains the given keyword.
    pub fn contains_item(&self, keyword: &str, qa_type: QuickAccess) -> WincentResult<bool> {
        let items = self.get_items(qa_type)?;
        Ok(items.iter().any(|item| item.contains(keyword)))
    }

    /// Adds an item to Recent Files or Frequent Folders.
    pub fn add_item(
        &self,
        path: &str,
        qa_type: QuickAccess,
        force_update: bool,
    ) -> WincentResult<()> {
        if matches!(qa_type, QuickAccess::All) {
            return Err(WincentError::UnsupportedOperation(format!(
                "Unsupported add operation for {:?}",
                qa_type
            )));
        }

        // This preflight check is best-effort; Explorer state may still change
        // before the shell operation runs.
        if self.check_item_exact(path, qa_type.clone())? {
            return Err(WincentError::AlreadyExists(path.to_string()));
        }

        match qa_type {
            QuickAccess::RecentFiles => {
                add_to_recent_files_with_options(path, AddRecentFileOptions { force_update })
            }
            QuickAccess::FrequentFolders => {
                add_to_frequent_folders_with_timeout(path, self.timeout)
            }
            QuickAccess::All => unreachable!(),
        }
    }

    /// Removes an item from Recent Files or Frequent Folders.
    pub fn remove_item(&self, path: &str, qa_type: QuickAccess) -> WincentResult<()> {
        if matches!(qa_type, QuickAccess::All) {
            return Err(WincentError::UnsupportedOperation(format!(
                "Unsupported remove operation for {:?}",
                qa_type
            )));
        }

        // This preflight check is best-effort; Explorer state may still change
        // before the shell operation runs.
        if !self.check_item_exact(path, qa_type.clone())? {
            return Err(WincentError::NotInRecent(path.to_string()));
        }

        match qa_type {
            QuickAccess::RecentFiles => remove_from_recent_files_with_timeout(path, self.timeout),
            QuickAccess::FrequentFolders => {
                remove_from_frequent_folders_with_timeout(path, self.timeout)
            }
            QuickAccess::All => unreachable!(),
        }
    }

    /// Adds multiple items to Quick Access, collecting per-item failures.
    pub fn add_items_batch(
        &self,
        items: &[(String, QuickAccess)],
        force_update: bool,
    ) -> BatchResult {
        batch::add_items_batch(
            items,
            BatchOptions {
                timeout: self.timeout,
                force_update,
            },
        )
    }

    /// Removes multiple items from Quick Access, collecting per-item failures.
    pub fn remove_items_batch(&self, items: &[(String, QuickAccess)]) -> BatchResult {
        batch::remove_items_batch(
            items,
            BatchOptions {
                timeout: self.timeout,
                force_update: false,
            },
        )
    }

    /// Clears Quick Access items.
    pub fn empty_items(
        &self,
        qa_type: QuickAccess,
        force_refresh: bool,
        also_pinned_folders: bool,
    ) -> WincentResult<()> {
        empty::empty_items(
            qa_type,
            EmptyOptions {
                also_pinned_folders,
                force_refresh,
            },
        )
    }

    /// Clears internal caches.
    ///
    /// The v0.2 manager no longer owns a script-result cache, so this is a no-op.
    pub fn clear_cache(&self) {}
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builder_default_timeout() {
        let manager = QuickAccessManager::builder().build();
        assert_eq!(manager.timeout(), Duration::from_secs(10));
    }

    #[test]
    fn builder_custom_timeout() {
        let manager = QuickAccessManager::builder()
            .timeout(Duration::from_secs(30))
            .build();
        assert_eq!(manager.timeout(), Duration::from_secs(30));
    }

    #[test]
    #[should_panic(expected = "Timeout must be greater than zero")]
    fn builder_zero_timeout_panics() {
        let _ = QuickAccessManager::builder().timeout(Duration::ZERO);
    }

    #[test]
    fn manager_new_uses_default_timeout() {
        let manager = QuickAccessManager::new();
        assert_eq!(manager.timeout(), Duration::from_secs(10));
    }

    #[test]
    fn unsupported_add_all_returns_error() {
        let manager = QuickAccessManager::new();
        let error = manager
            .add_item("C:\\test.txt", QuickAccess::All, false)
            .unwrap_err();

        assert!(matches!(error, WincentError::UnsupportedOperation(_)));
    }

    #[test]
    fn unsupported_remove_all_returns_error() {
        let manager = QuickAccessManager::new();
        let error = manager
            .remove_item("C:\\test.txt", QuickAccess::All)
            .unwrap_err();

        assert!(matches!(error, WincentError::UnsupportedOperation(_)));
    }
}
