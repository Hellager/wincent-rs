//! Synchronous facade for Windows Quick Access operations.

use crate::visible;
use crate::{
    backend::{QuickAccessBackend, SystemQuickAccessBackend},
    empty,
    error::WincentError,
    monitor::{QuickAccessChangeEvent, QuickAccessMonitor, QuickAccessMonitorOptions},
    quick_access_lock::{QuickAccessLock, QuickAccessLockTarget},
    status::QuickAccessStatus,
    utils::paths_equal,
    QuickAccess, WincentResult,
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
    backend: Arc<dyn QuickAccessBackend>,
}

impl fmt::Debug for QuickAccessManagerBuilder {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("QuickAccessManagerBuilder")
            .field("timeout", &self.timeout)
            .finish_non_exhaustive()
    }
}

impl Default for QuickAccessManagerBuilder {
    fn default() -> Self {
        Self {
            timeout: Duration::from_secs(10),
            backend: Arc::new(SystemQuickAccessBackend),
        }
    }
}

impl QuickAccessManagerBuilder {
    /// Sets the caller-side wait timeout for shell operations that support timeout control.
    ///
    /// This bounds how long wincent waits for supported Shell and COM
    /// operations to report a result. It does not cancel a Shell COM call that
    /// is already running inside Windows. A timeout therefore means
    /// the final mutation state is unknown: the underlying operation may still
    /// complete later and affect Explorer state. Mutation operations are not
    /// retried automatically.
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
    /// This bounds how long wincent waits for supported Shell and COM
    /// operations to report a result. It does not cancel a Shell COM call that
    /// is already running inside Windows. A timeout therefore means
    /// the final mutation state is unknown: the underlying operation may still
    /// complete later and affect Explorer state. Mutation operations are not
    /// retried automatically.
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

    #[cfg(test)]
    pub(crate) fn with_backend_for_tests(mut self, backend: Arc<dyn QuickAccessBackend>) -> Self {
        self.backend = backend;
        self
    }

    /// Builds a configured [`QuickAccessManager`].
    pub fn build(self) -> QuickAccessManager {
        QuickAccessManager {
            timeout: self.timeout,
            backend: self.backend,
        }
    }
}

/// Windows Quick Access manager.
///
/// This type is a thin synchronous facade over Quick Access query, mutation,
/// clearing, monitoring, locking, refresh, and visibility operations.
///
/// Prefer this facade over lower-level modules. It keeps path conversion,
/// timeout handling, visibility helpers, and DestList helpers behind one stable
/// API surface.
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

    /// Returns the configured caller-side wait timeout for supported Shell and
    /// COM operations.
    #[must_use]
    pub fn timeout(&self) -> Duration {
        self.timeout
    }

    #[cfg(test)]
    pub(crate) fn with_backend_for_tests(
        timeout: Duration,
        backend: Arc<dyn QuickAccessBackend>,
    ) -> Self {
        Self { timeout, backend }
    }

    /// Retrieves Quick Access items as strings.
    ///
    /// This reads Explorer's on-disk DestList directly and applies the
    /// `CAutoList` rules plus Explorer's Folder Options hidden-item setting.
    /// Queries do not invoke Shell COM.
    ///
    /// The returned values match the UI-visible result from persisted backing
    /// data. Hidden filtering uses the attributes serialized when each item
    /// was added, not the target's current filesystem attributes. An already-
    /// open Explorer view may temporarily differ while its UI cache refreshes.
    /// Returned paths may also point to items that no longer exist on disk.
    ///
    /// # Errors
    ///
    /// Returns I/O, path-resolution, unsupported-version, or DestList parse
    /// errors when the persisted backing data cannot be read and evaluated.
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
        self.backend.get_items(qa_type, self.timeout)
    }

    /// Retrieves Quick Access items as path buffers.
    ///
    /// The returned paths are the UI-visible DestList results converted into
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

    /// Checks whether Quick Access contains the given path.
    ///
    /// This is exact **path** membership using Windows path semantics:
    /// case-insensitive, slash-normalized, and canonicalized when the path
    /// exists. It is not a substring search; use
    /// [`QuickAccessManager::contains_keyword`] for keyword matching.
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
    /// let exists = manager.contains_path(
    ///     "C:\\Work\\report.docx",
    ///     QuickAccess::RecentFiles,
    /// )?;
    /// println!("exact match: {exists}");
    /// # Ok(())
    /// # }
    /// ```
    pub fn contains_path<P: AsRef<Path>>(
        &self,
        path: P,
        qa_type: QuickAccess,
    ) -> WincentResult<bool> {
        let path = path_to_shell_string(path.as_ref())?;
        let items = self.get_items(qa_type)?;
        Ok(items.iter().any(|item| paths_equal(item, &path)))
    }

    /// Returns where a path appears in a default Quick Access Shell query.
    ///
    /// Both lists are read from Explorer's `.automaticDestinations-ms` backing
    /// files and evaluated with the same rules as the DestList API. The result
    /// mirrors the observed default Shell query without running Shell/COM code
    /// or modifying state. It does not inspect the currently rendered Explorer
    /// window, which may remain stale until refreshed.
    ///
    /// A path omitted by both rule-based results returns the default
    /// [`QuickAccessStatus`], for which
    /// [`QuickAccessStatus::is_visible`] is `false`. This does not distinguish
    /// "absent from the backing file" from "present but filtered out"; use
    /// [`crate::destlist::entries`] on a parsed backing file to inspect the raw
    /// entries.
    ///
    /// Explorer persists its own changes asynchronously, so a path that was
    /// just added can be absent, or a path that was just removed can remain, in
    /// the rule-based result for a short time. A missing backing file is treated
    /// as an empty list.
    ///
    /// # Errors
    ///
    /// Returns [`WincentError::InvalidPath`] when `path` is empty. Returns I/O or
    /// DestList parse errors when an existing backing file cannot be read or
    /// parsed, and [`WincentError::DestListUnsupportedVersion`] when Explorer
    /// uses an unsupported DestList version.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # fn main() -> WincentResult<()> {
    /// let manager = QuickAccessManager::new();
    /// let status = manager.get_item_status(r"C:\Work\report.docx")?;
    /// if let Some(position) = status.recent_file() {
    ///     println!("recent file at position {position}");
    /// }
    /// if let Some(slot) = status.frequent_folder() {
    ///     println!("frequent folder pinned={}", slot.is_pinned());
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn get_item_status<P: AsRef<Path>>(&self, path: P) -> WincentResult<QuickAccessStatus> {
        let path = path_to_shell_string(path.as_ref())?;
        crate::status::status_for_path(&path)
    }

    /// Finds Quick Access items containing the given keyword.
    ///
    /// This is a case-insensitive substring check against Explorer's path
    /// strings. It is useful for loose search, but it can produce false
    /// positives. Use [`QuickAccessManager::contains_path`] for exact path
    /// membership. An empty keyword returns all queried items, matching
    /// [`str::contains`] behavior.
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
    /// for item in manager.find_items_containing("projects", QuickAccess::All)? {
    ///     println!("{item}");
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn find_items_containing(
        &self,
        keyword: &str,
        qa_type: QuickAccess,
    ) -> WincentResult<Vec<String>> {
        let keyword = keyword.to_lowercase();
        Ok(self
            .get_items(qa_type)?
            .into_iter()
            .filter(|item| item.to_lowercase().contains(&keyword))
            .collect())
    }

    /// Checks whether any Quick Access item contains the given keyword.
    ///
    /// This is a case-insensitive substring search, not exact path matching;
    /// use [`QuickAccessManager::contains_path`] for path membership. An empty
    /// keyword matches any non-empty list, matching [`str::contains`] behavior.
    ///
    /// Items are evaluated lazily without collecting a temporary match list.
    /// Use [`QuickAccessManager::find_items_containing`] when the matching
    /// items themselves are needed.
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
    /// if manager.contains_keyword("projects", QuickAccess::All)? {
    ///     println!("Found a matching Quick Access item");
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn contains_keyword(&self, keyword: &str, qa_type: QuickAccess) -> WincentResult<bool> {
        let keyword = keyword.to_lowercase();
        Ok(self
            .get_items(qa_type)?
            .iter()
            .any(|item| item.to_lowercase().contains(&keyword)))
    }

    /// Monitors Quick Access for snapshot changes.
    ///
    /// This starts a background polling thread. The callback is invoked when
    /// the monitored snapshot changes, including changes caused by this crate's
    /// own add, remove, empty, and restore operations. Keep the returned
    /// [`QuickAccessMonitor`] alive for as long as monitoring should continue.
    ///
    /// # Errors
    ///
    /// Returns query errors from [`QuickAccessManager::get_items`] when the
    /// initial snapshot cannot be read.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # fn main() -> WincentResult<()> {
    /// let manager = QuickAccessManager::new();
    /// let _monitor = manager.watch_quick_access(
    ///     QuickAccessMonitorOptions::new().with_qa_type(QuickAccess::All),
    ///     |event| {
    ///         if let Ok(event) = event {
    ///             println!("{} Quick Access items", event.current_items().len());
    ///         }
    ///     },
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn watch_quick_access<F>(
        &self,
        options: QuickAccessMonitorOptions,
        callback: F,
    ) -> WincentResult<QuickAccessMonitor>
    where
        F: FnMut(WincentResult<QuickAccessChangeEvent>) + Send + 'static,
    {
        QuickAccessMonitor::start(self.clone(), options, callback)
    }

    /// Adds an existing file or directory to Quick Access.
    ///
    /// Files are added to Recent Files and directories are pinned to Frequent
    /// Folders. Other path types and missing paths are rejected. Existing items
    /// return [`WincentError::AlreadyExists`]. Explorer updates its UI
    /// asynchronously, so the visible result may lag behind a successful call.
    ///
    /// `Ok(())` means that the selected Shell mutation returned successfully. It
    /// does not guarantee that Explorer has already persisted the change, that a
    /// new rule-based query observes it, or that an open Explorer window has
    /// redrawn. Use [`QuickAccessManager::get_item_status`] when persisted-state
    /// confirmation is required and [`QuickAccessManager::refresh`] when an open
    /// Explorer view should be notified and refreshed.
    ///
    /// A timeout is an indeterminate outcome rather than a confirmed failure.
    /// The underlying Shell operation can still complete after this method
    /// returns [`WincentError::Timeout`], so callers should inspect state before
    /// deciding whether another mutation is appropriate.
    ///
    /// # Errors
    ///
    /// Returns [`WincentError::InvalidPath`] when `path` is empty, missing, or
    /// is neither a file nor directory; [`WincentError::AlreadyExists`] when the
    /// item is already present; or the underlying operation error.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # fn main() -> WincentResult<()> {
    /// let manager = QuickAccessManager::new();
    /// manager.add_item("C:\\Work\\report.docx")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_item<P: AsRef<Path>>(&self, path: P) -> WincentResult<()> {
        let path = path_to_shell_string(path.as_ref())?;
        let path_type = self.backend.path_type(&path)?;
        let qa_type = match path_type {
            crate::utils::PathType::File => QuickAccess::RecentFiles,
            crate::utils::PathType::Directory => QuickAccess::FrequentFolders,
        };
        self.ensure_not_present(&path, qa_type)?;

        match path_type {
            crate::utils::PathType::File => self.backend.add_recent_file(&path, self.timeout),
            crate::utils::PathType::Directory => {
                self.backend.add_frequent_folder(&path, self.timeout)
            }
        }
    }

    /// Removes an existing file or directory from Quick Access.
    ///
    /// The operation is selected automatically:
    ///
    /// - files use the Recent Files `remove` Shell verb, which blocks the entry
    ///   from the UI but can retain its DestList entry and numeric stream;
    /// - visible unpinned directories use the native `removefromhome` Shell
    ///   verb, which blocks the entry from the UI while retaining its DestList
    ///   entry and numeric stream;
    /// - pinned directories use the native `unpinfromhome` Shell verb. Windows
    ///   can retain an entry with cleared usage or physically remove a zero-use
    ///   entry.
    ///
    /// `Ok(())` means that the selected Shell mutation returned successfully.
    /// Frequent-folder removal also verifies the live namespace while handling
    /// Shell verbs that can behave like toggles. Neither result guarantees that
    /// Explorer has already persisted the final DestList state or redrawn an
    /// open window.
    ///
    /// A timeout is an indeterminate outcome rather than a confirmed failure.
    /// The underlying Shell operation can still complete after this method
    /// returns [`WincentError::Timeout`]. Mutations are never retried
    /// automatically; callers should inspect state before issuing another add or
    /// remove operation.
    ///
    /// # Errors
    ///
    /// Returns [`WincentError::InvalidPath`] when `path` is empty, missing, or
    /// is neither a file nor directory;
    /// [`WincentError::NotInQuickAccess`] if the preflight query does not find
    /// the item, or an underlying Shell/COM error.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use wincent::prelude::*;
    ///
    /// # fn main() -> WincentResult<()> {
    /// let manager = QuickAccessManager::new();
    /// manager.remove_item("C:\\Work\\report.docx")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove_item<P: AsRef<Path>>(&self, path: P) -> WincentResult<()> {
        let path = path_to_shell_string(path.as_ref())?;
        let path_type = self.backend.path_type(&path)?;
        let qa_type = match path_type {
            crate::utils::PathType::File => QuickAccess::RecentFiles,
            crate::utils::PathType::Directory => QuickAccess::FrequentFolders,
        };

        match path_type {
            crate::utils::PathType::File => {
                self.ensure_present(&path, qa_type)?;
                self.backend.remove_recent_file(&path, self.timeout)
            }
            crate::utils::PathType::Directory => {
                self.ensure_present(&path, qa_type)?;
                let pin_status = self.backend.frequent_folder_pin_status(&path)?;
                match pin_status {
                    crate::destlist::FrequentFolderPinStatus::Pinned => {
                        self.backend
                            .remove_pinned_frequent_folder(&path, self.timeout)?;
                    }
                    crate::destlist::FrequentFolderPinStatus::Unpinned => {
                        self.backend
                            .remove_unpinned_frequent_folder(&path, self.timeout)?;
                    }
                    crate::destlist::FrequentFolderPinStatus::NotFound => {
                        return Err(WincentError::not_in_quick_access(&path, qa_type));
                    }
                }
                Ok(())
            }
        }
    }

    /// Locks Explorer's Recent Files and Frequent Folders automatic destination files.
    ///
    /// The returned guard holds OS file handles until it is dropped or explicitly
    /// unlocked. While held, Explorer cannot add or remove items in the locked
    /// Quick Access categories because the backing files are not writable.
    ///
    /// Use [`QuickAccessLock::unlock`] with
    /// [`crate::QuickAccessUnlockOptions::cleanup_new_recent_links`] to delete
    /// `.lnk` files that appeared in the Windows Recent folder while locked.
    pub fn lock_quick_access(&self) -> WincentResult<QuickAccessLock> {
        QuickAccessLock::lock()
    }

    /// Locks Explorer's Recent Files automatic destination file.
    ///
    /// While held, Explorer cannot add or remove Recent Files entries because
    /// the Recent Files backing file is not writable.
    pub fn lock_recent_files(&self) -> WincentResult<QuickAccessLock> {
        QuickAccessLock::lock_target(QuickAccessLockTarget::RecentFiles)
    }

    /// Locks Explorer's Frequent Folders automatic destination file.
    ///
    /// While held, Explorer cannot add or remove Frequent Folders entries
    /// because the Frequent Folders backing file is not writable.
    pub fn lock_frequent_folders(&self) -> WincentResult<QuickAccessLock> {
        QuickAccessLock::lock_target(QuickAccessLockTarget::FrequentFolders)
    }

    /// Notifies the Shell about Quick Access changes and refreshes matching
    /// Quick Access or Home Explorer windows.
    ///
    /// This refreshes presentation state only. It does not wait for an earlier
    /// mutation to be persisted and does not verify the result of that mutation.
    ///
    /// # Errors
    ///
    /// Returns an error when matching Explorer windows are found but none of
    /// their active Shell views can be refreshed, or when the refresh worker
    /// exceeds the manager timeout.
    pub fn refresh(&self) -> WincentResult<()> {
        self.backend.refresh(self.timeout)
    }

    /// Clears Quick Access items.
    ///
    /// `QuickAccess::RecentFiles` removes all Recent Files history and its
    /// corresponding shortcuts. `QuickAccess::FrequentFolders` removes all
    /// recorded Frequent Folders and their corresponding shortcuts.
    /// `QuickAccess::All` clears both categories.
    ///
    /// User-created Frequent Folders entries, including user-pinned folders,
    /// are not preserved. Windows may recreate system-default pinned folders.
    /// A successful clear refreshes matching Quick Access or Home windows.
    ///
    /// # Errors
    ///
    /// Returns Shell, registry, or I/O errors from the requested cleanup
    /// operations. If one category is cleared and a later cleanup step
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
    /// manager.empty_items(QuickAccess::All)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn empty_items(&self, qa_type: QuickAccess) -> WincentResult<()> {
        empty::empty_items_with_backend(qa_type, self.timeout, self.backend.as_ref())
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
    /// When hiding `QuickAccess::FrequentFolders` through this registry-backed
    /// API, Explorer hides unpinned frequent folders, while pinned folders
    /// remain visible. Passing `QuickAccess::All` with `visible` set to `false`
    /// has the same Frequent Folders behavior.
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
    /// Hiding `QuickAccess::FrequentFolders` through this registry-backed API
    /// hides unpinned frequent folders, while pinned folders remain visible.
    /// Passing `QuickAccess::All` has the same Frequent Folders behavior.
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

    /// Sets whether a Quick Access section is visible in Explorer, with optional
    /// Explorer refresh.
    ///
    /// When hiding `QuickAccess::FrequentFolders` through this registry-backed
    /// API, Explorer hides unpinned frequent folders, while pinned folders
    /// remain visible. Passing `QuickAccess::All` with `visible` set to `false`
    /// has the same Frequent Folders behavior.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when the current user's Explorer settings
    /// cannot be created or updated.
    /// [`WincentError::VisibilityPostMutationFailure`] means the registry write
    /// completed but the requested Explorer refresh failed; do not blindly
    /// retry the visibility mutation.
    pub fn set_visible_with_options(
        &self,
        qa_type: QuickAccess,
        visible: bool,
        options: visible::VisibilityOptions,
    ) -> WincentResult<()> {
        visible::set_visible_with_options(qa_type, visible, options)
    }

    /// Shows a Quick Access section in Explorer, with optional Explorer refresh.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when Explorer visibility settings cannot be
    /// updated. [`WincentError::VisibilityPostMutationFailure`] means the
    /// registry write completed but the requested Explorer refresh failed.
    pub fn show_section_with_options(
        &self,
        qa_type: QuickAccess,
        options: visible::VisibilityOptions,
    ) -> WincentResult<()> {
        self.set_visible_with_options(qa_type, true, options)
    }

    /// Hides a Quick Access section in Explorer, with optional Explorer refresh.
    ///
    /// Hiding `QuickAccess::FrequentFolders` through this registry-backed API
    /// hides unpinned frequent folders, while pinned folders remain visible.
    /// Passing `QuickAccess::All` has the same Frequent Folders behavior.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when Explorer visibility settings cannot be
    /// updated. [`WincentError::VisibilityPostMutationFailure`] means the
    /// registry write completed but the requested Explorer refresh failed.
    pub fn hide_section_with_options(
        &self,
        qa_type: QuickAccess,
        options: visible::VisibilityOptions,
    ) -> WincentResult<()> {
        self.set_visible_with_options(qa_type, false, options)
    }

    /// Sets whether Recent Files are visible, with optional Explorer refresh.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when Explorer visibility settings cannot be
    /// updated. [`WincentError::VisibilityPostMutationFailure`] means the
    /// registry write completed but the requested Explorer refresh failed.
    pub fn set_recent_files_visible_with_options(
        &self,
        visible: bool,
        options: visible::VisibilityOptions,
    ) -> WincentResult<()> {
        visible::set_recent_files_visible_with_options(visible, options)
    }

    /// Sets whether Frequent Folders are visible, with optional Explorer refresh.
    ///
    /// When hiding Frequent Folders through this registry-backed API, Explorer
    /// hides unpinned frequent folders, while pinned folders remain visible.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when Explorer visibility settings cannot be
    /// updated. [`WincentError::VisibilityPostMutationFailure`] means the
    /// registry write completed but the requested Explorer refresh failed.
    pub fn set_frequent_folders_visible_with_options(
        &self,
        visible: bool,
        options: visible::VisibilityOptions,
    ) -> WincentResult<()> {
        visible::set_frequent_folders_visible_with_options(visible, options)
    }

    /// Checks whether recently used files appear in Windows Recommended items.
    ///
    /// This reads the current-user `Start_TrackDocs` Explorer Advanced value
    /// used by the Windows Recommended recent-files setting. Policy values,
    /// MDM settings, Windows edition, or Explorer version can override the
    /// effective UI state.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when the current-user Explorer Advanced key
    /// cannot be read.
    pub fn is_start_recommended_section_visible(&self) -> WincentResult<bool> {
        visible::is_start_recommended_section_visible()
    }

    /// Sets whether recently used files appear in Windows Recommended items.
    ///
    /// Passing `true` writes `Start_TrackDocs = 1`; passing `false` writes
    /// `Start_TrackDocs = 0` under the current-user Explorer Advanced key.
    /// This controls whether recently used files appear in Windows Recommended
    /// items.
    ///
    /// On some Windows 11 builds, setting this to `false` can also make
    /// Explorer Recent Files invisible. In that state, make Start Recommended
    /// visible again before relying on `ShowRecent` to control Recent Files
    /// visibility.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when the current-user Explorer Advanced key
    /// cannot be created or updated.
    pub fn set_start_recommended_section_visible(&self, visible: bool) -> WincentResult<()> {
        visible::set_start_recommended_section_visible(visible)
    }

    /// Shows recently used files in Windows Recommended items.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when the current-user Explorer Advanced key
    /// cannot be created or updated.
    pub fn show_start_recommended_section(&self) -> WincentResult<()> {
        self.set_start_recommended_section_visible(true)
    }

    /// Hides recently used files in Windows Recommended items.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when the current-user Explorer Advanced key
    /// cannot be created or updated.
    pub fn hide_start_recommended_section(&self) -> WincentResult<()> {
        self.set_start_recommended_section_visible(false)
    }

    /// Sets whether recently used files appear in Windows Recommended items,
    /// with optional Explorer refresh.
    ///
    /// On some Windows 11 builds, hiding Start Recommended can also make
    /// Explorer Recent Files invisible. If that happens, set Start Recommended
    /// visible again before relying on `ShowRecent` to control Recent Files
    /// visibility.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when the current-user Explorer Advanced key
    /// cannot be created or updated.
    /// [`WincentError::VisibilityPostMutationFailure`] means the registry write
    /// completed but the requested Explorer refresh failed; do not blindly
    /// retry the visibility mutation.
    pub fn set_start_recommended_section_visible_with_options(
        &self,
        visible: bool,
        options: visible::VisibilityOptions,
    ) -> WincentResult<()> {
        visible::set_start_recommended_section_visible_with_options(visible, options)
    }

    /// Shows recently used files in Windows Recommended items, with optional
    /// Explorer refresh.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when the current-user Explorer Advanced key
    /// cannot be created or updated.
    /// [`WincentError::VisibilityPostMutationFailure`] means the registry write
    /// completed but the requested Explorer refresh failed.
    pub fn show_start_recommended_section_with_options(
        &self,
        options: visible::VisibilityOptions,
    ) -> WincentResult<()> {
        self.set_start_recommended_section_visible_with_options(true, options)
    }

    /// Hides recently used files in Windows Recommended items, with optional
    /// Explorer refresh.
    ///
    /// # Errors
    ///
    /// Returns registry I/O errors when the current-user Explorer Advanced key
    /// cannot be created or updated.
    /// [`WincentError::VisibilityPostMutationFailure`] means the registry write
    /// completed but the requested Explorer refresh failed.
    pub fn hide_start_recommended_section_with_options(
        &self,
        options: visible::VisibilityOptions,
    ) -> WincentResult<()> {
        self.set_start_recommended_section_visible_with_options(false, options)
    }

    fn ensure_not_present(&self, path: &str, qa_type: QuickAccess) -> WincentResult<()> {
        // This preflight check is best-effort; Explorer state may still change
        // before the shell operation runs.
        match self.contains_path(path, qa_type) {
            Ok(true) => Err(WincentError::already_exists(path, qa_type)),
            Ok(false) => Ok(()),
            Err(WincentError::Io(error)) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
            Err(error) => Err(error),
        }
    }

    fn ensure_present(&self, path: &str, qa_type: QuickAccess) -> WincentResult<()> {
        // This preflight check is best-effort; Explorer state may still change
        // before the shell operation runs.
        if !self.contains_path(path, qa_type)? {
            return Err(WincentError::not_in_quick_access(path, qa_type));
        }

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Mutex;

    #[derive(Default)]
    struct FakeBackend {
        items: Mutex<Vec<String>>,
        calls: Mutex<Vec<String>>,
        get_item_timeouts: Mutex<Vec<Duration>>,
        get_items_error: Mutex<Option<WincentError>>,
        refresh_timeouts: Mutex<Vec<Duration>>,
        validate_path_error: Mutex<Option<WincentError>>,
        add_frequent_folder_error: Mutex<Option<WincentError>>,
        remove_recent_file_error: Mutex<Option<WincentError>>,
        frequent_pin_status: Mutex<Option<crate::destlist::FrequentFolderPinStatus>>,
        refresh_error: Mutex<Option<WincentError>>,
    }

    impl FakeBackend {
        fn with_items(items: Vec<String>) -> Self {
            Self {
                items: Mutex::new(items),
                calls: Mutex::new(Vec::new()),
                get_item_timeouts: Mutex::new(Vec::new()),
                get_items_error: Mutex::new(None),
                refresh_timeouts: Mutex::new(Vec::new()),
                validate_path_error: Mutex::new(None),
                add_frequent_folder_error: Mutex::new(None),
                remove_recent_file_error: Mutex::new(None),
                frequent_pin_status: Mutex::new(None),
                refresh_error: Mutex::new(None),
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

        fn fail_get_items_with(&self, error: WincentError) {
            *self.get_items_error.lock().unwrap() = Some(error);
        }

        fn fail_add_frequent_folder_with(&self, error: WincentError) {
            *self.add_frequent_folder_error.lock().unwrap() = Some(error);
        }

        fn fail_remove_recent_file_with(&self, error: WincentError) {
            *self.remove_recent_file_error.lock().unwrap() = Some(error);
        }

        fn set_frequent_pin_status(&self, status: crate::destlist::FrequentFolderPinStatus) {
            *self.frequent_pin_status.lock().unwrap() = Some(status);
        }
    }

    impl crate::backend::QuickAccessBackend for FakeBackend {
        fn path_type(&self, path: &str) -> WincentResult<crate::utils::PathType> {
            if let Some(error) = self.validate_path_error.lock().unwrap().take() {
                return Err(error);
            }
            if Path::new(path).extension().is_some() {
                Ok(crate::utils::PathType::File)
            } else {
                Ok(crate::utils::PathType::Directory)
            }
        }

        fn get_items(
            &self,
            _qa_type: QuickAccess,
            timeout: Duration,
        ) -> WincentResult<Vec<String>> {
            self.get_item_timeouts.lock().unwrap().push(timeout);
            if let Some(error) = self.get_items_error.lock().unwrap().take() {
                return Err(error);
            }
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
            if let Some(error) = self.remove_recent_file_error.lock().unwrap().take() {
                return Err(error);
            }
            Ok(())
        }

        fn frequent_folder_pin_status(
            &self,
            _path: &str,
        ) -> WincentResult<crate::destlist::FrequentFolderPinStatus> {
            Ok(self
                .frequent_pin_status
                .lock()
                .unwrap()
                .unwrap_or(crate::destlist::FrequentFolderPinStatus::Pinned))
        }

        fn remove_pinned_frequent_folder(
            &self,
            path: &str,
            timeout: Duration,
        ) -> WincentResult<()> {
            self.record(format!(
                "remove_pinned_frequent_folder:{path}:{}",
                timeout.as_secs()
            ));
            Ok(())
        }

        fn remove_unpinned_frequent_folder(
            &self,
            path: &str,
            timeout: Duration,
        ) -> WincentResult<()> {
            self.record(format!(
                "remove_unpinned_frequent_folder:{path}:{}",
                timeout.as_secs()
            ));
            Ok(())
        }

        fn clear_recent_files(&self, _timeout: Duration) -> WincentResult<()> {
            self.record("clear_recent_files");
            Ok(())
        }

        fn clear_frequent_folders_jumplist(&self, _timeout: Duration) -> WincentResult<()> {
            self.record("clear_frequent_folders_jumplist");
            Ok(())
        }

        fn notify_shell_change(&self) {
            self.record("notify_shell_change");
        }

        fn refresh_shell_views(&self, timeout: Duration) -> WincentResult<()> {
            self.record("refresh_shell_views");
            self.refresh_timeouts.lock().unwrap().push(timeout);
            if let Some(error) = self.refresh_error.lock().unwrap().take() {
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

    #[test]
    fn find_items_containing_returns_matching_items() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::with_items(vec![
            "C:\\Projects\\report.docx".to_string(),
            "C:\\Temp\\notes.txt".to_string(),
            "D:\\Projects\\README.md".to_string(),
        ]));
        let manager = QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend);

        let matches = manager.find_items_containing("projects", QuickAccess::All)?;

        assert_eq!(
            matches,
            vec![
                "C:\\Projects\\report.docx".to_string(),
                "D:\\Projects\\README.md".to_string(),
            ]
        );
        Ok(())
    }

    #[test]
    fn find_items_containing_returns_empty_vec_without_matches() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::with_items(vec![
            "C:\\Temp\\notes.txt".to_string()
        ]));
        let manager = QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend);

        let matches = manager.find_items_containing("Projects", QuickAccess::All)?;

        assert!(matches.is_empty());
        Ok(())
    }

    #[test]
    fn find_items_containing_empty_keyword_returns_all_items() -> WincentResult<()> {
        let items = vec![
            "C:\\Projects\\report.docx".to_string(),
            "C:\\Temp\\notes.txt".to_string(),
        ];
        let backend = Arc::new(FakeBackend::with_items(items.clone()));
        let manager = QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend);

        let matches = manager.find_items_containing("", QuickAccess::All)?;

        assert_eq!(matches, items);
        Ok(())
    }

    #[test]
    fn contains_keyword_matches_substring_case_insensitively() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::with_items(vec!["C:\\Projects".to_string()]));
        let manager = QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend);

        assert!(manager.contains_keyword("projects", QuickAccess::All)?);
        assert!(!manager.contains_keyword("Missing", QuickAccess::All)?);
        Ok(())
    }

    #[test]
    fn contains_keyword_empty_keyword_matches_non_empty_list() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::with_items(vec!["C:\\Projects".to_string()]));
        let manager = QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend);

        assert!(
            manager.contains_keyword("", QuickAccess::All)?,
            "an empty keyword matches any item, matching str::contains behavior"
        );

        let empty_backend = Arc::new(FakeBackend::with_items(vec![]));
        let empty_manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), empty_backend);
        assert!(!empty_manager.contains_keyword("", QuickAccess::All)?);
        Ok(())
    }

    #[test]
    fn empty_frequent_folders_refreshes_after_clear() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::with_items(vec!["C:\\Projects".to_string()]));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(9), backend.clone());

        manager.empty_items(QuickAccess::FrequentFolders)?;

        assert!(backend.get_item_timeouts().is_empty());
        assert_eq!(
            backend.calls(),
            vec![
                "clear_frequent_folders_jumplist".to_string(),
                "notify_shell_change".to_string(),
                "refresh_shell_views".to_string(),
            ]
        );
        Ok(())
    }

    #[test]
    fn refresh_notifies_before_refreshing_views() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::with_items(Vec::new()));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(9), backend.clone());

        manager.refresh()?;

        assert_eq!(
            backend.calls(),
            vec![
                "notify_shell_change".to_string(),
                "refresh_shell_views".to_string(),
            ]
        );
        assert_eq!(
            *backend.refresh_timeouts.lock().unwrap(),
            vec![Duration::from_secs(9)]
        );
        Ok(())
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
    fn add_item_dispatches_file_without_post_processing() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::default());
        let manager = QuickAccessManager::builder()
            .with_backend_for_tests(backend.clone())
            .build();

        manager.add_item("C:\\report.docx")?;

        assert_eq!(
            backend.calls(),
            vec!["add_recent_file:C:\\report.docx".to_string()]
        );
        Ok(())
    }

    #[test]
    fn add_item_dispatches_directory_without_post_processing() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::default());
        let manager = QuickAccessManager::builder()
            .with_backend_for_tests(backend.clone())
            .build();

        manager.add_item("C:\\Projects")?;

        assert_eq!(
            backend.calls(),
            vec!["add_frequent_folder:C:\\Projects".to_string()]
        );
        Ok(())
    }

    #[test]
    fn add_duplicate_skips_backend_mutation() {
        let backend = Arc::new(FakeBackend::with_items(vec!["C:\\report.docx".to_string()]));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        let error = manager.add_item("c:/REPORT.docx").unwrap_err();

        assert!(matches!(error, WincentError::AlreadyExists { .. }));
        assert!(backend.calls().is_empty());
    }

    #[test]
    fn add_frequent_folder_rejects_existing_before_mutation() {
        let backend = Arc::new(FakeBackend::with_items(vec!["C:\\Projects".to_string()]));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        let result = manager.add_item("C:\\Projects");

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

        let result = manager.add_item("C:\\Projects");

        assert!(matches!(result, Err(WincentError::AlreadyExists { .. })));
        assert_eq!(backend.get_item_timeouts(), vec![Duration::from_secs(10)]);
        assert_eq!(
            backend.calls(),
            vec!["add_frequent_folder:C:\\Projects".to_string()]
        );
    }

    #[test]
    fn add_item_path_type_error_skips_membership_preflight() {
        let backend = Arc::new(FakeBackend::default());
        backend
            .fail_validate_path_with(WincentError::invalid_path_reason("Path is not a directory"));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        let result = manager.add_item("C:\\Projects");

        assert!(matches!(result, Err(WincentError::InvalidPath(_))));
        assert!(backend.get_item_timeouts().is_empty());
        assert!(backend.calls().is_empty());
    }

    #[test]
    fn add_item_allows_missing_destlist_during_duplicate_preflight() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::default());
        backend.fail_get_items_with(WincentError::Io(std::io::Error::from(
            std::io::ErrorKind::NotFound,
        )));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        manager.add_item("C:\\report.docx")?;

        assert_eq!(
            backend.calls(),
            vec!["add_recent_file:C:\\report.docx".to_string()]
        );
        Ok(())
    }

    #[test]
    fn status_queries_reject_empty_path_before_reading_destlist() {
        let manager = QuickAccessManager::new();

        assert!(matches!(
            manager.get_item_status(""),
            Err(WincentError::InvalidPath(_))
        ));
    }

    #[test]
    fn remove_recent_file_default_does_not_refresh_explorer() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::with_items(vec!["C:\\report.docx".to_string()]));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        manager.remove_item("C:\\report.docx")?;

        assert_eq!(
            backend.calls(),
            vec!["remove_recent_file:C:\\report.docx".to_string()]
        );
        Ok(())
    }

    #[test]
    fn remove_recent_file_propagates_error_without_retrying() {
        let backend = Arc::new(FakeBackend::with_items(vec!["C:\\report.docx".to_string()]));
        backend.fail_remove_recent_file_with(WincentError::Timeout(
            "mutation outcome is unknown".to_string(),
        ));
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        let result = manager.remove_item("C:\\report.docx");

        assert!(matches!(result, Err(WincentError::Timeout(_))));
        assert_eq!(
            backend.calls(),
            vec!["remove_recent_file:C:\\report.docx".to_string()]
        );
    }

    #[test]
    fn remove_unpinned_frequent_folder_uses_native_verb_path() -> WincentResult<()> {
        let backend = Arc::new(FakeBackend::with_items(vec!["C:\\Projects".to_string()]));
        backend.set_frequent_pin_status(crate::destlist::FrequentFolderPinStatus::Unpinned);
        let manager =
            QuickAccessManager::with_backend_for_tests(Duration::from_secs(10), backend.clone());

        manager.remove_item("C:\\Projects")?;

        assert_eq!(
            backend.calls(),
            vec!["remove_unpinned_frequent_folder:C:\\Projects:10".to_string()]
        );
        Ok(())
    }

    #[test]
    fn add_item_rejects_missing_path_before_membership_check() {
        let manager = QuickAccessManager::new();
        let missing = unique_temp_path("missing-add-recent.txt");
        let error = manager.add_item(&missing).unwrap_err();

        assert!(matches!(error, WincentError::InvalidPath(_)));
    }

    #[test]
    fn remove_recent_file_validates_path_before_membership_check() {
        let manager = QuickAccessManager::new();
        let missing = unique_temp_path("missing-remove-recent.txt");
        let error = manager.remove_item(&missing).unwrap_err();

        assert!(matches!(error, WincentError::InvalidPath(_)));
    }

    fn unique_temp_path(name: &str) -> PathBuf {
        std::env::temp_dir().join(format!("wincent-rs-{}-{}", std::process::id(), name))
    }
}
