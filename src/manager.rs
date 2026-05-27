//! Synchronous facade for Windows Quick Access operations.

#[cfg(feature = "visible")]
use crate::visible;
use crate::{
    batch::{self, BatchFailure, BatchOptions, BatchResult},
    empty::{self, EmptyOptions},
    error::WincentError,
    handle::{
        add_to_frequent_folders_with_timeout, add_to_recent_files_with_options,
        remove_from_frequent_folders_with_timeout, remove_from_recent_files_with_timeout,
        AddRecentFileOptions,
    },
    query,
    quick_access_lock::{QuickAccessLock, QuickAccessLockTarget},
    recent_links::delete_recent_links_for_target,
    utils::paths_equal,
    QuickAccess, WincentResult,
};
use std::path::{Path, PathBuf};
use std::time::Duration;

fn path_to_shell_string(path: &Path) -> WincentResult<String> {
    if path.as_os_str().is_empty() {
        return Err(WincentError::invalid_path_reason("Empty path provided"));
    }

    Ok(path.to_string_lossy().into_owned())
}

/// Options for adding a single item to Quick Access.
///
/// Use [`AddOptions::refresh_recent_files`] when a newly added Recent Files
/// entry should become visible immediately in Explorer. Frequent Folders are
/// pinned through Explorer shell verbs and ignore that option.
///
/// # Examples
///
/// ```rust
/// use wincent::AddOptions;
///
/// let options = AddOptions::new().refresh_recent_files();
/// assert!(options.force_update());
/// ```
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

/// Options for removing a single item from Quick Access.
///
/// By default removal only asks Explorer to remove the visible Quick Access
/// item. Use [`RemoveOptions::deep_clean_recent_links`] when the matching
/// shortcut in the Windows Recent folder should be deleted too.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct RemoveOptions {
    deep_clean_recent_links: bool,
}

impl RemoveOptions {
    /// Creates default remove options.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether matching `.lnk` files in the Windows Recent folder are deleted.
    #[must_use]
    pub fn deep_clean_recent_links_enabled(&self) -> bool {
        self.deep_clean_recent_links
    }

    /// Sets whether matching `.lnk` files in the Windows Recent folder are deleted.
    #[must_use]
    pub fn with_deep_clean_recent_links(mut self, deep_clean_recent_links: bool) -> Self {
        self.deep_clean_recent_links = deep_clean_recent_links;
        self
    }

    /// Deletes matching `.lnk` files from the Windows Recent folder after removing an item.
    #[must_use]
    pub fn deep_clean_recent_links(self) -> Self {
        self.with_deep_clean_recent_links(true)
    }

    pub(crate) fn should_deep_clean_links_for(&self, qa_type: QuickAccess) -> bool {
        matches!(
            qa_type,
            QuickAccess::RecentFiles | QuickAccess::FrequentFolders
        ) && self.deep_clean_recent_links
    }
}

/// Item used by batch Quick Access operations.
///
/// Batch operations accept typed items instead of a bare `(path, category)`
/// tuple so each path carries its intended Quick Access category.
///
/// # Examples
///
/// ```rust
/// use std::path::Path;
/// use wincent::{QuickAccess, QuickAccessItem};
///
/// let file = QuickAccessItem::recent_file("C:\\Work\\report.docx");
/// assert_eq!(file.qa_type(), QuickAccess::RecentFiles);
/// assert_eq!(file.path(), Path::new("C:\\Work\\report.docx"));
/// ```
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QuickAccessItem {
    path: PathBuf,
    qa_type: QuickAccess,
}

impl QuickAccessItem {
    /// Creates a batch item for the given Quick Access category.
    #[must_use]
    pub fn new<P: Into<PathBuf>>(path: P, qa_type: QuickAccess) -> Self {
        Self {
            path: path.into(),
            qa_type,
        }
    }

    /// Creates a Recent Files batch item.
    #[must_use]
    pub fn recent_file<P: Into<PathBuf>>(path: P) -> Self {
        Self::new(path, QuickAccess::RecentFiles)
    }

    /// Creates a Frequent Folders batch item.
    #[must_use]
    pub fn frequent_folder<P: Into<PathBuf>>(path: P) -> Self {
        Self::new(path, QuickAccess::FrequentFolders)
    }

    /// Path to process.
    #[must_use]
    pub fn path(&self) -> &Path {
        &self.path
    }

    /// Quick Access category for this item.
    #[must_use]
    pub fn qa_type(&self) -> QuickAccess {
        self.qa_type
    }
}

/// Builder for configuring [`QuickAccessManager`].
///
/// # Examples
///
/// ```rust
/// use std::time::Duration;
/// use wincent::prelude::QuickAccessManager;
///
/// let manager = QuickAccessManager::builder()
///     .timeout(Duration::from_secs(30))
///     .build();
///
/// assert_eq!(manager.timeout(), Duration::from_secs(30));
/// ```
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
    ///
    /// # Errors
    ///
    /// Returns [`WincentError::InvalidArgument`] when `duration` is zero.
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
///
/// Prefer this facade over lower-level modules. It keeps path conversion,
/// timeout handling, batch conversion, and feature-gated helpers behind one
/// stable API surface.
///
/// # Examples
///
/// ```rust,no_run
/// use wincent::prelude::*;
///
/// # fn main() -> WincentResult<()> {
/// let manager = QuickAccessManager::new();
/// let recent = manager.get_item_paths(QuickAccess::RecentFiles)?;
///
/// for path in recent {
///     println!("{}", path.display());
/// }
/// # Ok(())
/// # }
/// ```
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

    /// Retrieves Quick Access items as strings.
    ///
    /// The returned values reflect Explorer's current Quick Access state and
    /// may include paths that no longer exist on disk.
    ///
    /// # Errors
    ///
    /// Returns an error if both the native COM query path and the PowerShell
    /// fallback fail. Common causes include an unavailable desktop session,
    /// inaccessible Shell namespaces, PowerShell execution failure, or registry
    /// and I/O failures while resolving system locations.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # fn main() -> WincentResult<()> {
    /// let manager = QuickAccessManager::new();
    /// let items = manager.get_items(QuickAccess::All)?;
    /// println!("{} Quick Access items", items.len());
    /// # Ok(())
    /// # }
    /// ```
    pub fn get_items(&self, qa_type: QuickAccess) -> WincentResult<Vec<String>> {
        match qa_type {
            QuickAccess::RecentFiles => query::get_recent_files(),
            QuickAccess::FrequentFolders => query::get_frequent_folders(),
            QuickAccess::All => query::get_quick_access_items(),
        }
    }

    /// Retrieves Quick Access items as path buffers.
    ///
    /// The returned paths are Explorer's current path strings converted into
    /// [`PathBuf`]. They may point to files or folders that no longer exist.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`QuickAccessManager::get_items`].
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # fn main() -> WincentResult<()> {
    /// let manager = QuickAccessManager::new();
    /// for path in manager.get_item_paths(QuickAccess::RecentFiles)? {
    ///     println!("{}", path.display());
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn get_item_paths(&self, qa_type: QuickAccess) -> WincentResult<Vec<PathBuf>> {
        Ok(self
            .get_items(qa_type)?
            .into_iter()
            .map(PathBuf::from)
            .collect())
    }

    /// Checks if an item exists in Quick Access using Windows path semantics.
    ///
    /// This performs exact path comparison with Windows-style normalization
    /// rather than substring matching. Use [`QuickAccessManager::contains_item`]
    /// for simple string containment.
    ///
    /// # Errors
    ///
    /// Returns [`WincentError::InvalidPath`] when `path` is empty. Also returns
    /// query errors from [`QuickAccessManager::get_items`].
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # fn main() -> WincentResult<()> {
    /// let manager = QuickAccessManager::new();
    /// let exists = manager.check_item_exact(
    ///     "C:\\Work\\report.docx",
    ///     QuickAccess::RecentFiles,
    /// )?;
    /// println!("exact match: {exists}");
    /// # Ok(())
    /// # }
    /// ```
    pub fn check_item_exact<P: AsRef<Path>>(
        &self,
        path: P,
        qa_type: QuickAccess,
    ) -> WincentResult<bool> {
        let path = path_to_shell_string(path.as_ref())?;
        let items = self.get_items(qa_type)?;
        Ok(items.iter().any(|item| paths_equal(item, &path)))
    }

    /// Checks if any item in Quick Access contains the given keyword.
    ///
    /// This is a plain, case-sensitive substring check against Explorer's path
    /// strings. It is useful for loose search, but it can produce false
    /// positives for path membership checks.
    ///
    /// # Errors
    ///
    /// Returns query errors from [`QuickAccessManager::get_items`].
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # fn main() -> WincentResult<()> {
    /// let manager = QuickAccessManager::new();
    /// if manager.contains_item("Projects", QuickAccess::All)? {
    ///     println!("Found a matching Quick Access item");
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn contains_item(&self, keyword: &str, qa_type: QuickAccess) -> WincentResult<bool> {
        let items = self.get_items(qa_type)?;
        Ok(items.iter().any(|item| item.contains(keyword)))
    }

    /// Adds an item to Recent Files or Frequent Folders.
    ///
    /// `QuickAccess::RecentFiles` records a file in the Recent Files list.
    /// `QuickAccess::FrequentFolders` pins a folder through Explorer's shell
    /// verbs. `QuickAccess::All` is not a valid target for mutation.
    ///
    /// The duplicate check is best-effort. Explorer state can change between
    /// the preflight query and the shell operation.
    ///
    /// # Errors
    ///
    /// Returns [`WincentError::InvalidPath`] when `path` is empty, missing, or
    /// has the wrong file/folder type for `qa_type`; [`WincentError::AlreadyExists`]
    /// if the preflight query finds the item; [`WincentError::UnsupportedOperation`]
    /// for `QuickAccess::All`; or a Shell/PowerShell error if the underlying
    /// operation fails.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # fn main() -> WincentResult<()> {
    /// let manager = QuickAccessManager::new();
    /// manager.add_item(
    ///     "C:\\Work\\report.docx",
    ///     QuickAccess::RecentFiles,
    ///     AddOptions::new().refresh_recent_files(),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_item<P: AsRef<Path>>(
        &self,
        path: P,
        qa_type: QuickAccess,
        options: AddOptions,
    ) -> WincentResult<()> {
        let path = path_to_shell_string(path.as_ref())?;

        match qa_type {
            QuickAccess::RecentFiles => {
                self.ensure_not_present(&path, qa_type)?;
                add_to_recent_files_with_options(
                    &path,
                    AddRecentFileOptions {
                        force_update: options.force_update(),
                    },
                )
            }
            QuickAccess::FrequentFolders => {
                self.ensure_not_present(&path, qa_type)?;
                add_to_frequent_folders_with_timeout(&path, self.timeout)
            }
            unsupported => Err(unsupported_add(unsupported)),
        }
    }

    /// Removes an item from Recent Files or Frequent Folders.
    ///
    /// `QuickAccess::All` is not a valid target for mutation. Use
    /// [`QuickAccessManager::empty_items`] when the goal is to clear a whole
    /// category.
    ///
    /// # Errors
    ///
    /// Returns [`WincentError::InvalidPath`] when `path` is empty;
    /// [`WincentError::NotInQuickAccess`] if the preflight query does not find
    /// the item; [`WincentError::UnsupportedOperation`] for `QuickAccess::All`;
    /// or a Shell/PowerShell error if the underlying operation fails.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # fn main() -> WincentResult<()> {
    /// let manager = QuickAccessManager::new();
    /// manager.remove_item("C:\\Work\\report.docx", QuickAccess::RecentFiles)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove_item<P: AsRef<Path>>(&self, path: P, qa_type: QuickAccess) -> WincentResult<()> {
        self.remove_item_with_options(path, qa_type, RemoveOptions::new())
    }

    /// Removes an item from Quick Access with additional cleanup options.
    ///
    /// [`RemoveOptions::deep_clean_recent_links`] applies to both
    /// [`QuickAccess::RecentFiles`] and [`QuickAccess::FrequentFolders`].
    pub fn remove_item_with_options<P: AsRef<Path>>(
        &self,
        path: P,
        qa_type: QuickAccess,
        options: RemoveOptions,
    ) -> WincentResult<()> {
        let path = path_to_shell_string(path.as_ref())?;

        match qa_type {
            QuickAccess::RecentFiles => {
                self.ensure_present(&path, qa_type)?;
                remove_from_recent_files_with_timeout(&path, self.timeout)?;
                if options.should_deep_clean_links_for(qa_type) {
                    delete_recent_links_for_target(&path, self.timeout)?;
                }
                Ok(())
            }
            QuickAccess::FrequentFolders => {
                self.ensure_present(&path, qa_type)?;
                remove_from_frequent_folders_with_timeout(&path, self.timeout)?;
                if options.should_deep_clean_links_for(qa_type) {
                    delete_recent_links_for_target(&path, self.timeout)?;
                }
                Ok(())
            }
            unsupported => Err(unsupported_remove(unsupported)),
        }
    }

    /// Adds multiple items to Quick Access, collecting per-item failures.
    ///
    /// This method does not short-circuit and does not return `Result`.
    /// Per-item failures are collected in [`BatchResult::failed`].
    ///
    /// Failure ordering is phase-based. Items whose paths cannot be converted
    /// into shell strings are reported first, followed by failures from the
    /// underlying batch add operations. The failure list is not guaranteed to
    /// preserve input order.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// let manager = QuickAccessManager::new();
    /// let items = [
    ///     QuickAccessItem::recent_file("C:\\Work\\report.docx"),
    ///     QuickAccessItem::frequent_folder("C:\\Work"),
    /// ];
    ///
    /// let result = manager.add_items_batch(&items, BatchOptions::new());
    /// for failure in result.failed() {
    ///     eprintln!("{}: {}", failure.path(), failure.error());
    /// }
    /// ```
    pub fn add_items_batch(&self, items: &[QuickAccessItem], options: BatchOptions) -> BatchResult {
        let (items, failures) = self.convert_batch_items(items);
        let result = batch::add_items_batch(&items, options, self.timeout);
        merge_batch_failures(result, failures)
    }

    /// Removes multiple items from Quick Access, collecting per-item failures.
    ///
    /// This method does not short-circuit and does not return `Result`.
    /// Per-item failures are collected in [`BatchResult::failed`].
    ///
    /// Failure ordering is phase-based. Items whose paths cannot be converted
    /// into shell strings are reported first, followed by failures from the
    /// underlying batch remove operations. The failure list is not guaranteed
    /// to preserve input order.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// let manager = QuickAccessManager::new();
    /// let items = [QuickAccessItem::recent_file("C:\\Work\\report.docx")];
    ///
    /// let result = manager.remove_items_batch(&items);
    /// if !result.is_complete_success() {
    ///     eprintln!("{} item(s) failed", result.failed().len());
    /// }
    /// ```
    pub fn remove_items_batch(&self, items: &[QuickAccessItem]) -> BatchResult {
        let (items, failures) = self.convert_batch_items(items);
        let result = batch::remove_items_batch(&items, self.timeout);
        merge_batch_failures(result, failures)
    }

    /// Removes multiple items from Quick Access with additional cleanup options.
    ///
    /// If deep link cleanup is enabled and a matching `.lnk` deletion fails
    /// after the shell remove succeeds, that item is reported in
    /// [`BatchResult::failed`].
    pub fn remove_items_batch_with_options(
        &self,
        items: &[QuickAccessItem],
        options: RemoveOptions,
    ) -> BatchResult {
        let (items, failures) = self.convert_batch_items(items);
        let result = batch::remove_items_batch_with_options(&items, options, self.timeout);
        merge_batch_failures(result, failures)
    }

    /// Locks Explorer's Recent Files and Frequent Folders automatic destination files.
    ///
    /// The returned guard holds OS file handles until it is dropped or explicitly
    /// unlocked. Use [`QuickAccessLock::unlock`] with
    /// [`crate::QuickAccessUnlockOptions::cleanup_new_recent_links`] to delete
    /// `.lnk` files that appeared in the Windows Recent folder while locked.
    pub fn lock_quick_access(&self) -> WincentResult<QuickAccessLock> {
        QuickAccessLock::lock()
    }

    /// Locks Explorer's Recent Files automatic destination file.
    pub fn lock_recent_files(&self) -> WincentResult<QuickAccessLock> {
        QuickAccessLock::lock_target(QuickAccessLockTarget::RecentFiles)
    }

    /// Locks Explorer's Frequent Folders automatic destination file.
    pub fn lock_frequent_folders(&self) -> WincentResult<QuickAccessLock> {
        QuickAccessLock::lock_target(QuickAccessLockTarget::FrequentFolders)
    }

    /// Clears Quick Access items.
    ///
    /// `QuickAccess::RecentFiles` clears Recent Files. `QuickAccess::FrequentFolders`
    /// clears user-visited frequent folders and optionally pinned folders.
    /// `QuickAccess::All` clears both categories.
    ///
    /// # Errors
    ///
    /// Returns Shell, PowerShell, registry, or I/O errors from the requested
    /// cleanup operations. If one category is cleared and a later cleanup step
    /// fails, returns [`WincentError::PartialEmpty`] so callers can distinguish
    /// partial progress from a complete failure.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # fn main() -> WincentResult<()> {
    /// let manager = QuickAccessManager::new();
    /// manager.empty_items(
    ///     QuickAccess::All,
    ///     EmptyOptions::new().remove_pinned_folders().refresh_explorer(),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn empty_items(&self, qa_type: QuickAccess, options: EmptyOptions) -> WincentResult<()> {
        empty::empty_items(qa_type, options)
    }

    /// Clears internal caches.
    ///
    /// Currently this is a no-op retained for API compatibility.
    pub fn clear_cache(&self) {}

    /// Checks whether a Quick Access section is visible in Explorer.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when the current user's Explorer settings
    /// cannot be read.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// # #[cfg(feature = "visible")]
    /// # fn main() -> wincent::WincentResult<()> {
    /// use wincent::prelude::*;
    ///
    /// let manager = QuickAccessManager::new();
    /// let visible = manager.is_visible(QuickAccess::RecentFiles)?;
    /// println!("Recent Files visible: {visible}");
    /// # Ok(())
    /// # }
    /// # #[cfg(not(feature = "visible"))]
    /// # fn main() {}
    /// ```
    #[cfg(feature = "visible")]
    pub fn is_visible(&self, qa_type: QuickAccess) -> WincentResult<bool> {
        visible::is_visible(qa_type)
    }

    /// Sets whether a Quick Access section is visible in Explorer.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when the current user's Explorer settings
    /// cannot be created or updated.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// # #[cfg(feature = "visible")]
    /// # fn main() -> wincent::WincentResult<()> {
    /// use wincent::prelude::*;
    ///
    /// let manager = QuickAccessManager::new();
    /// manager.set_visible(QuickAccess::RecentFiles, true)?;
    /// # Ok(())
    /// # }
    /// # #[cfg(not(feature = "visible"))]
    /// # fn main() {}
    /// ```
    #[cfg(feature = "visible")]
    pub fn set_visible(&self, qa_type: QuickAccess, visible: bool) -> WincentResult<()> {
        visible::set_visible(qa_type, visible)
    }

    /// Shows a Quick Access section in Explorer.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when Explorer visibility settings cannot be
    /// updated.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// # #[cfg(feature = "visible")]
    /// # fn main() -> wincent::WincentResult<()> {
    /// use wincent::prelude::*;
    ///
    /// QuickAccessManager::new().show_section(QuickAccess::FrequentFolders)?;
    /// # Ok(())
    /// # }
    /// # #[cfg(not(feature = "visible"))]
    /// # fn main() {}
    /// ```
    #[cfg(feature = "visible")]
    pub fn show_section(&self, qa_type: QuickAccess) -> WincentResult<()> {
        self.set_visible(qa_type, true)
    }

    /// Hides a Quick Access section in Explorer.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when Explorer visibility settings cannot be
    /// updated.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// # #[cfg(feature = "visible")]
    /// # fn main() -> wincent::WincentResult<()> {
    /// use wincent::prelude::*;
    ///
    /// QuickAccessManager::new().hide_section(QuickAccess::RecentFiles)?;
    /// # Ok(())
    /// # }
    /// # #[cfg(not(feature = "visible"))]
    /// # fn main() {}
    /// ```
    #[cfg(feature = "visible")]
    pub fn hide_section(&self, qa_type: QuickAccess) -> WincentResult<()> {
        self.set_visible(qa_type, false)
    }

    fn convert_batch_items(
        &self,
        items: &[QuickAccessItem],
    ) -> (Vec<(String, QuickAccess)>, Vec<BatchFailure>) {
        let mut converted = Vec::new();
        let mut failures = Vec::new();

        for item in items {
            match path_to_shell_string(item.path()) {
                Ok(path) => converted.push((path, item.qa_type())),
                Err(error) => {
                    failures.push(BatchFailure::new(item.path().display().to_string(), error))
                }
            }
        }

        (converted, failures)
    }

    fn ensure_not_present(&self, path: &str, qa_type: QuickAccess) -> WincentResult<()> {
        // This preflight check is best-effort; Explorer state may still change
        // before the shell operation runs.
        if self.check_item_exact(path, qa_type)? {
            return Err(WincentError::already_exists(path, qa_type));
        }

        Ok(())
    }

    fn ensure_present(&self, path: &str, qa_type: QuickAccess) -> WincentResult<()> {
        // This preflight check is best-effort; Explorer state may still change
        // before the shell operation runs.
        if !self.check_item_exact(path, qa_type)? {
            return Err(WincentError::not_in_quick_access(path, qa_type));
        }

        Ok(())
    }
}

fn unsupported_add(qa_type: QuickAccess) -> WincentError {
    WincentError::UnsupportedOperation(format!("Unsupported add operation for {:?}", qa_type))
}

fn unsupported_remove(qa_type: QuickAccess) -> WincentError {
    WincentError::UnsupportedOperation(format!("Unsupported remove operation for {:?}", qa_type))
}

fn merge_batch_failures(
    result: BatchResult,
    mut initial_failures: Vec<BatchFailure>,
) -> BatchResult {
    if initial_failures.is_empty() {
        return result;
    }

    let (succeeded, mut failed) = result.into_parts();
    initial_failures.append(&mut failed);
    BatchResult::new(succeeded, initial_failures)
}

#[cfg(feature = "destlist")]
impl QuickAccessManager {
    /// Parses the recent-files `.automaticDestinations-ms` file and returns all entries.
    ///
    /// # Errors
    ///
    /// Returns I/O errors when the file cannot be read, or DestList parse errors
    /// when Explorer's backing file is missing, truncated, corrupt, or uses an
    /// unsupported format version.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// # #[cfg(feature = "destlist")]
    /// # fn main() -> wincent::WincentResult<()> {
    /// use wincent::prelude::*;
    ///
    /// let manager = QuickAccessManager::new();
    /// for entry in manager.get_recent_files_metadata()? {
    ///     println!("{} ({})", entry.path(), entry.access_count());
    /// }
    /// # Ok(())
    /// # }
    /// # #[cfg(not(feature = "destlist"))]
    /// # fn main() {}
    /// ```
    pub fn get_recent_files_metadata(&self) -> WincentResult<Vec<crate::destlist::DestListEntry>> {
        let parsed = crate::destlist::parse_file(crate::destlist::recent_files_dest_path()?)?;
        Ok(crate::destlist::entries(parsed.dest_list()))
    }

    /// Parses the frequent-folders `.automaticDestinations-ms` file and returns all entries.
    ///
    /// # Errors
    ///
    /// Returns I/O errors when the file cannot be read, or DestList parse errors
    /// when Explorer's backing file is missing, truncated, corrupt, or uses an
    /// unsupported format version.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// # #[cfg(feature = "destlist")]
    /// # fn main() -> wincent::WincentResult<()> {
    /// use wincent::prelude::*;
    ///
    /// let manager = QuickAccessManager::new();
    /// let entries = manager.get_frequent_folders_metadata()?;
    /// println!("{} frequent-folder entries", entries.len());
    /// # Ok(())
    /// # }
    /// # #[cfg(not(feature = "destlist"))]
    /// # fn main() {}
    /// ```
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
    fn remove_options_default_disables_deep_clean() {
        assert!(!RemoveOptions::new().deep_clean_recent_links_enabled());
    }

    #[test]
    fn remove_options_builder_enables_deep_clean() {
        let options = RemoveOptions::new().deep_clean_recent_links();
        assert!(options.deep_clean_recent_links_enabled());
        assert!(!options
            .with_deep_clean_recent_links(false)
            .deep_clean_recent_links_enabled());
    }

    #[test]
    fn remove_options_deep_clean_targets_recent_files_and_frequent_folders() {
        let options = RemoveOptions::new().deep_clean_recent_links();

        assert!(options.should_deep_clean_links_for(QuickAccess::RecentFiles));
        assert!(options.should_deep_clean_links_for(QuickAccess::FrequentFolders));
        assert!(!options.should_deep_clean_links_for(QuickAccess::All));
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

    #[test]
    fn quick_access_item_constructors_preserve_path_and_type() {
        let item = QuickAccessItem::recent_file(PathBuf::from("C:\\test.txt"));
        assert_eq!(item.path(), Path::new("C:\\test.txt"));
        assert_eq!(item.qa_type(), QuickAccess::RecentFiles);

        let item = QuickAccessItem::frequent_folder("C:\\test");
        assert_eq!(item.path(), Path::new("C:\\test"));
        assert_eq!(item.qa_type(), QuickAccess::FrequentFolders);
    }

    #[test]
    fn batch_items_report_invalid_empty_paths_per_item() {
        let manager = QuickAccessManager::new();
        let result =
            manager.add_items_batch(&[QuickAccessItem::recent_file("")], BatchOptions::new());

        assert_eq!(result.failed().len(), 1);
        assert!(matches!(
            result.failed()[0].error(),
            WincentError::InvalidPath(_)
        ));
    }
}
