//! Synchronous facade for Windows Quick Access operations.

#[cfg(feature = "visible")]
use crate::visible;
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

/// Options for adding a single item to Quick Access.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct AddOptions {
    /// Refresh Recent Files display data after adding a recent file.
    ///
    /// This option only applies to [`QuickAccess::RecentFiles`]. Frequent
    /// folders are pinned through Explorer's shell verb and ignore this value.
    force_update: bool,
}

impl AddOptions {
    /// Creates default add options.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether Recent Files display data is refreshed after adding a recent file.
    #[must_use]
    pub fn force_update(&self) -> bool {
        self.force_update
    }

    /// Sets whether Recent Files display data is refreshed after adding a recent file.
    #[must_use]
    pub fn with_force_update(mut self, force_update: bool) -> Self {
        self.force_update = force_update;
        self
    }

    /// Refreshes Recent Files display data after adding a recent file.
    #[must_use]
    pub fn refresh_recent_files(self) -> Self {
        self.with_force_update(true)
    }
}

/// Builder for configuring [`QuickAccessManager`].
#[derive(Debug, Clone)]
#[must_use]
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
    #[must_use]
    pub fn timeout(mut self, duration: Duration) -> Self {
        assert!(!duration.is_zero(), "Timeout must be greater than zero");
        self.timeout = duration;
        self
    }

    /// Tries to set the timeout for shell operations that support timeout control.
    ///
    /// Unlike [`QuickAccessManagerBuilder::timeout`], this returns
    /// [`WincentError::InvalidArgument`] instead of panicking when `duration` is
    /// zero.
    pub fn try_timeout(mut self, duration: Duration) -> WincentResult<Self> {
        if duration.is_zero() {
            return Err(WincentError::InvalidArgument(
                "timeout must be greater than zero".to_string(),
            ));
        }
        self.timeout = duration;
        Ok(self)
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
#[derive(Debug, Clone)]
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
    #[must_use]
    pub fn builder() -> QuickAccessManagerBuilder {
        QuickAccessManagerBuilder::default()
    }

    /// Creates a manager with default configuration.
    #[must_use]
    pub fn new() -> Self {
        Self::builder().build()
    }

    /// Returns the configured shell operation timeout.
    #[must_use]
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
        options: AddOptions,
    ) -> WincentResult<()> {
        if matches!(qa_type, QuickAccess::All) {
            return Err(WincentError::UnsupportedOperation(format!(
                "Unsupported add operation for {:?}",
                qa_type
            )));
        }

        // This preflight check is best-effort; Explorer state may still change
        // before the shell operation runs.
        if self.check_item_exact(path, qa_type)? {
            return Err(WincentError::AlreadyExists(path.to_string()));
        }

        match qa_type {
            QuickAccess::RecentFiles => add_to_recent_files_with_options(
                path,
                AddRecentFileOptions {
                    force_update: options.force_update(),
                },
            ),
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
        if !self.check_item_exact(path, qa_type)? {
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
        options: BatchOptions,
    ) -> BatchResult {
        batch::add_items_batch(
            items,
            BatchOptions::from_parts(self.timeout, options.force_update()),
        )
    }

    /// Removes multiple items from Quick Access, collecting per-item failures.
    pub fn remove_items_batch(&self, items: &[(String, QuickAccess)]) -> BatchResult {
        batch::remove_items_batch(items, BatchOptions::from_parts(self.timeout, false))
    }

    /// Clears Quick Access items.
    pub fn empty_items(&self, qa_type: QuickAccess, options: EmptyOptions) -> WincentResult<()> {
        empty::empty_items(qa_type, options)
    }

    /// Clears internal caches.
    ///
    /// The v0.2 manager no longer owns a script-result cache, so this is a no-op.
    pub fn clear_cache(&self) {}

    /// Checks whether a Quick Access section is visible in Explorer.
    #[cfg(feature = "visible")]
    pub fn is_visible(&self, qa_type: QuickAccess) -> WincentResult<bool> {
        visible::is_visible(qa_type)
    }

    /// Sets whether a Quick Access section is visible in Explorer.
    #[cfg(feature = "visible")]
    pub fn set_visible(&self, qa_type: QuickAccess, visible: bool) -> WincentResult<()> {
        visible::set_visible(qa_type, visible)
    }

    /// Shows a Quick Access section in Explorer.
    #[cfg(feature = "visible")]
    pub fn show_section(&self, qa_type: QuickAccess) -> WincentResult<()> {
        self.set_visible(qa_type, true)
    }

    /// Hides a Quick Access section in Explorer.
    #[cfg(feature = "visible")]
    pub fn hide_section(&self, qa_type: QuickAccess) -> WincentResult<()> {
        self.set_visible(qa_type, false)
    }
}

#[cfg(feature = "destlist")]
impl QuickAccessManager {
    /// Parses the recent-files `.automaticDestinations-ms` file and returns all entries.
    pub fn get_recent_files_metadata(&self) -> WincentResult<Vec<crate::destlist::DestListEntry>> {
        let parsed = crate::destlist::parse_file(crate::destlist::recent_files_dest_path()?)?;
        Ok(crate::destlist::entries(parsed.dest_list()))
    }

    /// Parses the frequent-folders `.automaticDestinations-ms` file and returns all entries.
    pub fn get_frequent_folders_metadata(
        &self,
    ) -> WincentResult<Vec<crate::destlist::DestListEntry>> {
        let parsed = crate::destlist::parse_file(crate::destlist::frequent_folders_dest_path()?)?;
        Ok(crate::destlist::entries(parsed.dest_list()))
    }
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
    fn builder_try_timeout_returns_error_for_zero() {
        let error = QuickAccessManager::builder()
            .try_timeout(Duration::ZERO)
            .unwrap_err();

        assert!(matches!(error, WincentError::InvalidArgument(_)));
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
            .add_item("C:\\test.txt", QuickAccess::All, AddOptions::new())
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
