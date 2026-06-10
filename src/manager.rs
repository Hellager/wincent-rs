//! Synchronous facade for Windows Quick Access operations.

use crate::visible;
use crate::{
    backend::{QuickAccessBackend, SystemQuickAccessBackend},
    batch::{self, BatchFailure, BatchOptions, BatchResult},
    empty::{self, EmptyOptions},
    error::{QuickAccessPostMutationStep, WincentError},
    quick_access_lock::{QuickAccessLock, QuickAccessLockTarget},
    utils::paths_equal,
    QuickAccess, RetryPolicy, WincentResult,
};
use std::fmt;
use std::path::{Path, PathBuf};
use std::sync::Arc;
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
/// entry should be pushed toward immediate visibility in Explorer. The refresh
/// deletes Explorer's Recent Files backing data and refreshes open Explorer
/// windows, but it does not synchronously verify that Windows has rebuilt the
/// item. Frequent Folders are pinned through Explorer shell verbs and ignore
/// that option.
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
    ///
    /// If the add succeeds but the display refresh fails, the operation returns
    /// [`WincentError::PostMutationFailure`] so callers can distinguish a
    /// completed add from a failed post-mutation refresh.
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
#[derive(Clone)]
#[must_use]
pub struct QuickAccessManagerBuilder {
    timeout: Duration,
    retry_policy: RetryPolicy,
    backend: Arc<dyn QuickAccessBackend>,
}

impl fmt::Debug for QuickAccessManagerBuilder {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("QuickAccessManagerBuilder")
            .field("timeout", &self.timeout)
            .field("retry_policy", &self.retry_policy)
            .finish_non_exhaustive()
    }
}

impl Default for QuickAccessManagerBuilder {
    fn default() -> Self {
        Self {
            timeout: Duration::from_secs(10),
            retry_policy: RetryPolicy::standard(),
            backend: Arc::new(SystemQuickAccessBackend),
        }
    }
}

impl QuickAccessManagerBuilder {
    /// Sets the caller-side wait timeout for shell operations that support timeout control.
    ///
    /// This bounds how long wincent waits for supported Shell, COM, and
    /// PowerShell operations to report a result. It does not cancel a Shell COM
    /// call that is already running inside Windows; after a timeout, that
    /// underlying operation may still complete later and affect Explorer state.
    ///
    /// # Panics
    ///
    /// Panics if `duration` is zero.
    pub fn timeout(mut self, duration: Duration) -> Self {
        assert!(!duration.is_zero(), "Timeout must be greater than zero");
        self.timeout = duration;
        self
    }

    /// Tries to set the caller-side wait timeout for shell operations that support timeout control.
    ///
    /// This bounds how long wincent waits for supported Shell, COM, and
    /// PowerShell operations to report a result. It does not cancel a Shell COM
    /// call that is already running inside Windows; after a timeout, that
    /// underlying operation may still complete later and affect Explorer state.
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

    /// Sets the retry policy for transient PowerShell fallback failures.
    pub fn retry_policy(mut self, retry_policy: RetryPolicy) -> Self {
        self.retry_policy = retry_policy;
        self
    }

    #[cfg(test)]
    pub(crate) fn with_backend_for_tests(mut self, backend: Arc<dyn QuickAccessBackend>) -> Self {
        self.backend = backend;
        self
    }

    /// Builds a configured [`QuickAccessManager`].
    pub fn build(self) -> QuickAccessManager {
        self.try_build()
            .expect("invalid retry policy in QuickAccessManagerBuilder")
    }

    /// Tries to build a configured [`QuickAccessManager`].
    ///
    /// Unlike [`QuickAccessManagerBuilder::build`], this returns
    /// [`WincentError::InvalidArgument`] instead of panicking when the retry
    /// policy is invalid.
    ///
    /// # Errors
    ///
    /// Returns [`WincentError::InvalidArgument`] when the configured retry
    /// policy fails validation.
    pub fn try_build(self) -> WincentResult<QuickAccessManager> {
        self.retry_policy.validate()?;
        Ok(QuickAccessManager {
            timeout: self.timeout,
            retry_policy: self.retry_policy,
            backend: self.backend,
        })
    }
}

/// Windows Quick Access manager.
///
/// This type is a thin synchronous facade over the `query`, `handle`, `empty`,
/// and `batch` modules.
///
/// Prefer this facade over lower-level modules. It keeps path conversion,
/// timeout handling, batch conversion, visibility helpers, and DestList helpers
/// behind one stable API surface.
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
#[derive(Clone)]
pub struct QuickAccessManager {
    timeout: Duration,
    retry_policy: RetryPolicy,
    backend: Arc<dyn QuickAccessBackend>,
}

impl fmt::Debug for QuickAccessManager {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("QuickAccessManager")
            .field("timeout", &self.timeout)
            .finish_non_exhaustive()
    }
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
    #[must_use]
    pub fn new() -> Self {
        Self::builder().build()
    }

    /// Returns the configured shell operation timeout.
    #[must_use]
    pub fn timeout(&self) -> Duration {
        self.timeout
    }

    /// Returns the configured retry policy for transient PowerShell fallback failures.
    #[must_use]
    pub fn retry_policy(&self) -> &RetryPolicy {
        &self.retry_policy
    }

    #[cfg(test)]
    pub(crate) fn with_backend_for_tests(
        timeout: Duration,
        backend: Arc<dyn QuickAccessBackend>,
    ) -> Self {
        Self {
            timeout,
            retry_policy: RetryPolicy::standard(),
            backend,
        }
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
        self.execute_with_retry(|| self.backend.get_items(qa_type, self.timeout))
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
    /// Recent Files duplicate checks are best-effort preflight queries.
    /// Frequent Folders duplicate checks happen inside the Shell pin operation
    /// to reduce the race window before invoking Explorer's pin verb.
    ///
    /// # Errors
    ///
    /// Returns [`WincentError::InvalidPath`] when `path` is empty, missing, or
    /// has the wrong file/folder type for `qa_type`; [`WincentError::AlreadyExists`]
    /// if the item is already present; [`WincentError::UnsupportedOperation`] for
    /// `QuickAccess::All`; [`WincentError::PostMutationFailure`] when a Recent
    /// Files add succeeded but display refresh failed; or a Shell/PowerShell
    /// error if the underlying operation fails.
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
                self.backend
                    .validate_path(&path, crate::utils::PathType::File)?;
                self.ensure_not_present(&path, qa_type)?;
                self.backend.add_recent_file(&path, self.timeout)?;
                if options.force_update() {
                    self.refresh_recent_files_display_after_add(&path)?;
                }
                Ok(())
            }
            QuickAccess::FrequentFolders => {
                self.backend
                    .validate_path(&path, crate::utils::PathType::Directory)?;
                self.ensure_not_present(&path, qa_type)?;
                self.execute_with_retry(|| self.backend.add_frequent_folder(&path, self.timeout))
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
    /// [`WincentError::PostMutationFailure`] when a Recent Files remove
    /// succeeded but Explorer refresh failed; or a Shell/PowerShell error if
    /// the underlying operation fails.
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
                self.backend
                    .validate_path(&path, crate::utils::PathType::File)?;
                self.ensure_present(&path, qa_type)?;
                self.execute_with_retry(|| self.backend.remove_recent_file(&path, self.timeout))?;
                self.refresh_explorer_after_recent_remove(&path)?;
                if options.should_deep_clean_links_for(qa_type) {
                    self.backend
                        .delete_recent_links_for_target(&path, self.timeout)?;
                }
                Ok(())
            }
            QuickAccess::FrequentFolders => {
                self.backend
                    .validate_path(&path, crate::utils::PathType::Directory)?;
                self.ensure_present(&path, qa_type)?;
                self.execute_with_retry(|| {
                    self.backend.remove_frequent_folder(&path, self.timeout)
                })?;
                if options.should_deep_clean_links_for(qa_type) {
                    self.backend
                        .delete_recent_links_for_target(&path, self.timeout)?;
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
        let result = batch::add_items_batch(&items, options, self.timeout, self.backend.as_ref());
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
        let result = batch::remove_items_batch(&items, self.timeout, self.backend.as_ref());
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
        let result = batch::remove_items_batch_with_options(
            &items,
            options,
            self.timeout,
            self.backend.as_ref(),
        );
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
    /// deletes Explorer's Frequent Folders backing file. `QuickAccess::All`
    /// clears both categories.
    ///
    /// Deleting the Frequent Folders backing file can cause Explorer to rebuild
    /// default folder pins and remove user-pinned folders even when
    /// [`EmptyOptions::also_pinned_folders`] is `false`. Set
    /// [`EmptyOptions::remove_pinned_folders`] to additionally invoke Explorer's
    /// unpin verb for visible Frequent Folders items.
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
        empty::empty_items_with_backend(qa_type, options, self.timeout, self.backend.as_ref())
    }

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
    /// # fn main() -> wincent::WincentResult<()> {
    /// use wincent::prelude::*;
    ///
    /// let manager = QuickAccessManager::new();
    /// let visible = manager.is_visible(QuickAccess::RecentFiles)?;
    /// println!("Recent Files visible: {visible}");
    /// # Ok(())
    /// # }
    /// ```
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
    /// # fn main() -> wincent::WincentResult<()> {
    /// use wincent::prelude::*;
    ///
    /// let manager = QuickAccessManager::new();
    /// manager.set_visible(QuickAccess::RecentFiles, true)?;
    /// # Ok(())
    /// # }
    /// ```
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
    /// # fn main() -> wincent::WincentResult<()> {
    /// use wincent::prelude::*;
    ///
    /// QuickAccessManager::new().show_section(QuickAccess::FrequentFolders)?;
    /// # Ok(())
    /// # }
    /// ```
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
    /// # fn main() -> wincent::WincentResult<()> {
    /// use wincent::prelude::*;
    ///
    /// QuickAccessManager::new().hide_section(QuickAccess::RecentFiles)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn hide_section(&self, qa_type: QuickAccess) -> WincentResult<()> {
        self.set_visible(qa_type, false)
    }

    pub fn set_visible_with_options(
        &self,
        qa_type: QuickAccess,
        visible: bool,
        options: visible::VisibilityOptions,
    ) -> WincentResult<()> {
        visible::set_visible_with_options(qa_type, visible, options)
    }

    pub fn show_section_with_options(
        &self,
        qa_type: QuickAccess,
        options: visible::VisibilityOptions,
    ) -> WincentResult<()> {
        self.set_visible_with_options(qa_type, true, options)
    }

    pub fn hide_section_with_options(
        &self,
        qa_type: QuickAccess,
        options: visible::VisibilityOptions,
    ) -> WincentResult<()> {
        self.set_visible_with_options(qa_type, false, options)
    }

    pub fn set_recent_files_visible_with_options(
        &self,
        visible: bool,
        options: visible::VisibilityOptions,
    ) -> WincentResult<()> {
        visible::set_recent_files_visible_with_options(visible, options)
    }

    pub fn set_frequent_folders_visible_with_options(
        &self,
        visible: bool,
        options: visible::VisibilityOptions,
    ) -> WincentResult<()> {
        visible::set_frequent_folders_visible_with_options(visible, options)
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

    fn refresh_recent_files_display_after_add(&self, path: &str) -> WincentResult<()> {
        self.backend
            .delete_recent_files_backing_data()
            .map_err(|source| {
                WincentError::post_mutation_failure(
                    path,
                    QuickAccess::RecentFiles,
                    QuickAccessPostMutationStep::DeleteRecentFilesBackingData,
                    source,
                )
            })?;
        self.backend.refresh_explorer().map_err(|source| {
            WincentError::post_mutation_failure(
                path,
                QuickAccess::RecentFiles,
                QuickAccessPostMutationStep::RefreshExplorer,
                source,
            )
        })
    }

    fn refresh_explorer_after_recent_remove(&self, path: &str) -> WincentResult<()> {
        self.backend.refresh_explorer().map_err(|source| {
            WincentError::post_mutation_failure(
                path,
                QuickAccess::RecentFiles,
                QuickAccessPostMutationStep::RefreshExplorer,
                source,
            )
        })
    }

    fn execute_with_retry<T, F>(&self, mut action: F) -> WincentResult<T>
    where
        F: FnMut() -> WincentResult<T>,
    {
        for retry_attempt in 0.. {
            match action() {
                Ok(value) => return Ok(value),
                Err(error) if self.should_retry(&error, retry_attempt) => {
                    std::thread::sleep(self.retry_policy.calculate_delay(retry_attempt));
                }
                Err(error) => return Err(error),
            }
        }

        unreachable!("retry loop always returns")
    }

    fn should_retry(&self, error: &WincentError, retry_attempt: u32) -> bool {
        if retry_attempt >= self.retry_policy.max_attempts() {
            return false;
        }

        matches!(error, WincentError::PowerShellExecution(err) if err.is_transient())
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
    /// # fn main() -> wincent::WincentResult<()> {
    /// use wincent::prelude::*;
    ///
    /// let manager = QuickAccessManager::new();
    /// for entry in manager.get_recent_files_metadata()? {
    ///     println!("{} ({})", entry.path(), entry.access_count());
    /// }
    /// # Ok(())
    /// # }
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
    /// # fn main() -> wincent::WincentResult<()> {
    /// use wincent::prelude::*;
    ///
    /// let manager = QuickAccessManager::new();
    /// let entries = manager.get_frequent_folders_metadata()?;
    /// println!("{} frequent-folder entries", entries.len());
    /// # Ok(())
    /// # }
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
    use std::fs;
    use std::sync::Mutex;

    #[derive(Default)]
    struct FakeBackend {
        items: Mutex<Vec<String>>,
        calls: Mutex<Vec<String>>,
        get_item_timeouts: Mutex<Vec<Duration>>,
        validate_path_error: Mutex<Option<WincentError>>,
        add_frequent_folder_error: Mutex<Option<WincentError>>,
        delete_recent_files_backing_data_error: Mutex<Option<WincentError>>,
        refresh_explorer_error: Mutex<Option<WincentError>>,
    }

    impl FakeBackend {
        fn with_items(items: Vec<String>) -> Self {
            Self {
                items: Mutex::new(items),
                calls: Mutex::new(Vec::new()),
                get_item_timeouts: Mutex::new(Vec::new()),
                validate_path_error: Mutex::new(None),
                add_frequent_folder_error: Mutex::new(None),
                delete_recent_files_backing_data_error: Mutex::new(None),
                refresh_explorer_error: Mutex::new(None),
            }
        }

        fn calls(&self) -> Vec<String> {
            self.calls.lock().unwrap().clone()
        }

        fn record(&self, call: impl Into<String>) {
            self.calls.lock().unwrap().push(call.into());
        }

        fn get_item_timeouts(&self) -> Vec<Duration> {
            self.get_item_timeouts.lock().unwrap().clone()
        }

        fn fail_validate_path_with(&self, error: WincentError) {
            *self.validate_path_error.lock().unwrap() = Some(error);
        }

        fn fail_add_frequent_folder_with(&self, error: WincentError) {
            *self.add_frequent_folder_error.lock().unwrap() = Some(error);
        }

        fn fail_delete_recent_files_backing_data_with(&self, error: WincentError) {
            *self.delete_recent_files_backing_data_error.lock().unwrap() = Some(error);
        }

        fn fail_refresh_explorer_with(&self, error: WincentError) {
            *self.refresh_explorer_error.lock().unwrap() = Some(error);
        }
    }

    impl crate::backend::QuickAccessBackend for FakeBackend {
        fn validate_path(
            &self,
            _path: &str,
            _expected: crate::utils::PathType,
        ) -> WincentResult<()> {
            if let Some(error) = self.validate_path_error.lock().unwrap().take() {
                return Err(error);
            }
            Ok(())
        }

        fn get_items(
            &self,
            _qa_type: QuickAccess,
            timeout: Duration,
        ) -> WincentResult<Vec<String>> {
            self.get_item_timeouts.lock().unwrap().push(timeout);
            Ok(self.items.lock().unwrap().clone())
        }

        fn add_recent_file(&self, path: &str, _timeout: Duration) -> WincentResult<()> {
            self.record(format!("add_recent_file:{path}"));
            Ok(())
        }

        fn add_frequent_folder(&self, path: &str, _timeout: Duration) -> WincentResult<()> {
            self.record(format!("add_frequent_folder:{path}"));
            if let Some(error) = self.add_frequent_folder_error.lock().unwrap().take() {
                return Err(error);
            }
            Ok(())
        }

        fn remove_recent_file(&self, path: &str, _timeout: Duration) -> WincentResult<()> {
            self.record(format!("remove_recent_file:{path}"));
            Ok(())
        }

        fn remove_frequent_folder(&self, path: &str, _timeout: Duration) -> WincentResult<()> {
            self.record(format!("remove_frequent_folder:{path}"));
            Ok(())
        }

        fn delete_recent_links_for_target(
            &self,
            path: &str,
            _timeout: Duration,
        ) -> WincentResult<()> {
            self.record(format!("delete_recent_links:{path}"));
            Ok(())
        }

        fn delete_recent_files_backing_data(&self) -> WincentResult<()> {
            self.record("delete_recent_files_backing_data");
            if let Some(error) = self
                .delete_recent_files_backing_data_error
                .lock()
                .unwrap()
                .take()
            {
                return Err(error);
            }
            Ok(())
        }

        fn clear_recent_files(&self, _timeout: Duration) -> WincentResult<()> {
            self.record("clear_recent_files");
            Ok(())
        }

        fn clear_frequent_folders_jumplist(&self) -> WincentResult<()> {
            self.record("clear_frequent_folders_jumplist");
            Ok(())
        }

        fn refresh_explorer(&self) -> WincentResult<()> {
            self.record("refresh_explorer");
            if let Some(error) = self.refresh_explorer_error.lock().unwrap().take() {
                return Err(error);
            }
            Ok(())
        }
    }

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
    fn builder_default_retry_policy_is_standard() {
        let manager = QuickAccessManager::builder().build();
        assert_eq!(
            manager.retry_policy().max_attempts(),
            RetryPolicy::standard().max_attempts()
        );
    }

    #[test]
    fn builder_custom_retry_policy_is_exposed() {
        let policy = RetryPolicy::no_retry();
        let manager = QuickAccessManager::builder()
            .retry_policy(policy.clone())
            .build();

        assert_eq!(manager.retry_policy().max_attempts(), policy.max_attempts());
    }

    #[test]
    #[should_panic(expected = "invalid retry policy in QuickAccessManagerBuilder")]
    fn builder_build_panics_for_invalid_retry_policy() {
        let _ = QuickAccessManager::builder()
            .retry_policy(
                RetryPolicy::new()
                    .with_max_attempts(1)
                    .with_initial_delay(Duration::ZERO),
            )
            .build();
    }

    #[test]
    fn builder_try_build_returns_error_for_invalid_retry_policy() {
        let result = QuickAccessManager::builder()
            .retry_policy(
                RetryPolicy::new()
                    .with_max_attempts(1)
                    .with_initial_delay(Duration::ZERO),
            )
            .try_build();

        assert!(matches!(result, Err(WincentError::InvalidArgument(_))));
    }

    #[test]
    fn get_items_passes_manager_timeout_to_backend() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::default());
        let manager = QuickAccessManager::builder()
            .timeout(Duration::from_secs(7))
            .with_backend_for_tests(backend.clone())
            .build();

        let _ = manager.get_items(QuickAccess::RecentFiles)?;

        assert_eq!(backend.get_item_timeouts(), vec![Duration::from_secs(7)]);
        Ok(())
    }

    fn transient_powershell_error() -> WincentError {
        use crate::error::{PowerShellError, PowerShellErrorKind, PowerShellOperation};

        WincentError::PowerShellExecution(Box::new(
            PowerShellError::builder(PowerShellOperation::QueryQuickAccess)
                .kind(PowerShellErrorKind::Timeout)
                .exit_code(None)
                .stderr("timeout")
                .script_path("test.ps1")
                .build(),
        ))
    }

    fn non_transient_powershell_error() -> WincentError {
        use crate::error::{PowerShellError, PowerShellErrorKind, PowerShellOperation};

        WincentError::PowerShellExecution(Box::new(
            PowerShellError::builder(PowerShellOperation::QueryQuickAccess)
                .kind(PowerShellErrorKind::ProcessFailed)
                .exit_code(Some(1))
                .stderr("permanent failure")
                .script_path("test.ps1")
                .build(),
        ))
    }

    #[test]
    fn retry_policy_retries_transient_powershell_errors() -> WincentResult<()> {
        let manager = QuickAccessManager::builder()
            .retry_policy(
                RetryPolicy::new()
                    .with_max_attempts(2)
                    .with_initial_delay(Duration::from_millis(1))
                    .with_jitter(false),
            )
            .build();
        let attempts = Mutex::new(0u32);

        let result = manager.execute_with_retry(|| {
            let mut attempts = attempts.lock().unwrap();
            *attempts += 1;
            if *attempts < 3 {
                Err(transient_powershell_error())
            } else {
                Ok("ok")
            }
        })?;

        assert_eq!(result, "ok");
        assert_eq!(*attempts.lock().unwrap(), 3);
        Ok(())
    }

    #[test]
    fn retry_policy_no_retry_returns_transient_error_immediately() {
        let manager = QuickAccessManager::builder()
            .retry_policy(RetryPolicy::no_retry())
            .build();
        let attempts = Mutex::new(0u32);

        let result: WincentResult<()> = manager.execute_with_retry(|| {
            *attempts.lock().unwrap() += 1;
            Err(transient_powershell_error())
        });

        assert!(matches!(result, Err(WincentError::PowerShellExecution(_))));
        assert_eq!(*attempts.lock().unwrap(), 1);
    }

    #[test]
    fn retry_policy_does_not_retry_non_transient_powershell_errors() {
        let manager = QuickAccessManager::builder()
            .retry_policy(
                RetryPolicy::new()
                    .with_max_attempts(2)
                    .with_initial_delay(Duration::from_millis(1))
                    .with_jitter(false),
            )
            .build();
        let attempts = Mutex::new(0u32);

        let result: WincentResult<()> = manager.execute_with_retry(|| {
            *attempts.lock().unwrap() += 1;
            Err(non_transient_powershell_error())
        });

        assert!(matches!(result, Err(WincentError::PowerShellExecution(_))));
        assert_eq!(*attempts.lock().unwrap(), 1);
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
    fn add_recent_file_refresh_runs_add_then_display_refresh() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::default());
        let manager = QuickAccessManager::builder()
            .with_backend_for_tests(backend.clone())
            .build();

        manager.add_item(
            "C:\\report.docx",
            QuickAccess::RecentFiles,
            AddOptions::new().refresh_recent_files(),
        )?;

        assert_eq!(
            backend.calls(),
            vec![
                "add_recent_file:C:\\report.docx".to_string(),
                "delete_recent_files_backing_data".to_string(),
                "refresh_explorer".to_string(),
            ]
        );
        Ok(())
    }

    #[test]
    fn add_recent_file_without_refresh_skips_display_refresh() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::default());
        let manager = QuickAccessManager::builder()
            .with_backend_for_tests(backend.clone())
            .build();

        manager.add_item(
            "C:\\report.docx",
            QuickAccess::RecentFiles,
            AddOptions::new(),
        )?;

        assert_eq!(
            backend.calls(),
            vec!["add_recent_file:C:\\report.docx".to_string()]
        );
        Ok(())
    }

    #[test]
    fn add_recent_file_refresh_failure_is_post_mutation_failure() {
        let backend = Arc::new(FakeBackend::default());
        backend.fail_delete_recent_files_backing_data_with(WincentError::SystemError(
            "recent folder unavailable".to_string(),
        ));
        let manager = QuickAccessManager::builder()
            .with_backend_for_tests(backend.clone())
            .build();

        let error = manager
            .add_item(
                "C:\\report.docx",
                QuickAccess::RecentFiles,
                AddOptions::new().refresh_recent_files(),
            )
            .unwrap_err();

        assert!(matches!(
            error,
            WincentError::PostMutationFailure {
                ref path,
                qa_type: QuickAccess::RecentFiles,
                step: crate::error::QuickAccessPostMutationStep::DeleteRecentFilesBackingData,
                ref source,
            } if path == "C:\\report.docx"
                && matches!(source.as_ref(), WincentError::SystemError(message) if message == "recent folder unavailable")
        ));
        assert_eq!(
            backend.calls(),
            vec![
                "add_recent_file:C:\\report.docx".to_string(),
                "delete_recent_files_backing_data".to_string(),
            ]
        );
    }

    #[test]
    fn add_duplicate_skips_backend_mutation() {
        let backend = Arc::new(FakeBackend::with_items(vec!["C:\\report.docx".to_string()]));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        let error = manager
            .add_item(
                "c:/REPORT.docx",
                QuickAccess::RecentFiles,
                AddOptions::new(),
            )
            .unwrap_err();

        assert!(matches!(error, WincentError::AlreadyExists { .. }));
        assert!(backend.calls().is_empty());
    }

    #[test]
    fn add_frequent_folder_rejects_existing_before_mutation() {
        let backend = Arc::new(FakeBackend::with_items(vec!["C:\\Projects".to_string()]));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        let result = manager.add_item(
            "C:\\Projects",
            QuickAccess::FrequentFolders,
            AddOptions::new(),
        );

        assert!(matches!(result, Err(WincentError::AlreadyExists { .. })));
        assert_eq!(backend.get_item_timeouts(), vec![Duration::from_secs(10)]);
        assert!(backend.calls().is_empty());
    }

    #[test]
    fn add_frequent_folder_propagates_backend_already_exists() {
        let backend = Arc::new(FakeBackend::default());
        backend.fail_add_frequent_folder_with(WincentError::already_exists(
            "C:\\Projects",
            QuickAccess::FrequentFolders,
        ));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        let result = manager.add_item(
            "C:\\Projects",
            QuickAccess::FrequentFolders,
            AddOptions::new(),
        );

        assert!(matches!(result, Err(WincentError::AlreadyExists { .. })));
        assert_eq!(backend.get_item_timeouts(), vec![Duration::from_secs(10)]);
        assert_eq!(
            backend.calls(),
            vec!["add_frequent_folder:C:\\Projects".to_string()]
        );
    }

    #[test]
    fn add_frequent_folder_validate_error_skips_membership_preflight() {
        let backend = Arc::new(FakeBackend::default());
        backend
            .fail_validate_path_with(WincentError::invalid_path_reason("Path is not a directory"));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        let result = manager.add_item(
            "C:\\Projects",
            QuickAccess::FrequentFolders,
            AddOptions::new(),
        );

        assert!(matches!(result, Err(WincentError::InvalidPath(_))));
        assert!(backend.get_item_timeouts().is_empty());
        assert!(backend.calls().is_empty());
    }

    #[test]
    fn remove_deep_clean_runs_after_successful_remove() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::with_items(vec!["C:\\report.docx".to_string()]));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        manager.remove_item_with_options(
            "C:\\report.docx",
            QuickAccess::RecentFiles,
            RemoveOptions::new().deep_clean_recent_links(),
        )?;

        assert_eq!(
            backend.calls(),
            vec![
                "remove_recent_file:C:\\report.docx".to_string(),
                "refresh_explorer".to_string(),
                "delete_recent_links:C:\\report.docx".to_string(),
            ]
        );
        Ok(())
    }

    #[test]
    fn remove_recent_file_refreshes_explorer() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::with_items(vec!["C:\\report.docx".to_string()]));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        manager.remove_item("C:\\report.docx", QuickAccess::RecentFiles)?;

        assert_eq!(
            backend.calls(),
            vec![
                "remove_recent_file:C:\\report.docx".to_string(),
                "refresh_explorer".to_string(),
            ]
        );
        Ok(())
    }

    #[test]
    fn remove_recent_file_refresh_failure_is_post_mutation_failure() {
        let backend = Arc::new(FakeBackend::with_items(vec!["C:\\report.docx".to_string()]));
        backend.fail_refresh_explorer_with(WincentError::SystemError("refresh failed".to_string()));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        let error = manager
            .remove_item("C:\\report.docx", QuickAccess::RecentFiles)
            .unwrap_err();

        assert!(matches!(
            error,
            WincentError::PostMutationFailure {
                ref path,
                qa_type: QuickAccess::RecentFiles,
                step: crate::error::QuickAccessPostMutationStep::RefreshExplorer,
                ref source,
            } if path == "C:\\report.docx"
                && matches!(source.as_ref(), WincentError::SystemError(message) if message == "refresh failed")
        ));
        assert_eq!(
            backend.calls(),
            vec![
                "remove_recent_file:C:\\report.docx".to_string(),
                "refresh_explorer".to_string(),
            ]
        );
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
    fn add_recent_file_validates_path_before_membership_check() {
        let manager = QuickAccessManager::new();
        let missing = unique_temp_path("missing-add-recent.txt");
        let error = manager
            .add_item(&missing, QuickAccess::RecentFiles, AddOptions::new())
            .unwrap_err();

        assert!(matches!(error, WincentError::InvalidPath(_)));
    }

    #[test]
    fn remove_recent_file_validates_path_before_membership_check() {
        let manager = QuickAccessManager::new();
        let missing = unique_temp_path("missing-remove-recent.txt");
        let error = manager
            .remove_item(&missing, QuickAccess::RecentFiles)
            .unwrap_err();

        assert!(matches!(error, WincentError::InvalidPath(_)));
    }

    #[test]
    fn add_recent_file_rejects_existing_directory_before_membership_check() -> WincentResult<()> {
        let manager = QuickAccessManager::new();
        let directory = unique_temp_path("directory-as-recent-file");
        fs::create_dir_all(&directory)?;

        let error = manager
            .add_item(&directory, QuickAccess::RecentFiles, AddOptions::new())
            .unwrap_err();

        fs::remove_dir_all(&directory)?;
        assert!(matches!(error, WincentError::InvalidPath(_)));
        Ok(())
    }

    #[test]
    fn add_frequent_folder_rejects_existing_file_before_membership_check() -> WincentResult<()> {
        let manager = QuickAccessManager::new();
        let file = unique_temp_path("file-as-frequent-folder.txt");
        fs::write(&file, b"test")?;

        let error = manager
            .add_item(&file, QuickAccess::FrequentFolders, AddOptions::new())
            .unwrap_err();

        fs::remove_file(&file)?;
        assert!(matches!(error, WincentError::InvalidPath(_)));
        Ok(())
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

    fn unique_temp_path(name: &str) -> PathBuf {
        std::env::temp_dir().join(format!("wincent-rs-{}-{}", std::process::id(), name))
    }
}
