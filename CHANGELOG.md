## Unreleased

## [0.3.1-alpha.2] - 2026-11-09

### Added
- Public `standalone` module for querying Recent Files and Frequent Folders,
  checking folder pin state, adding, pinning, removing, clearing, locking,
  unlocking, and refreshing without creating a `QuickAccessManager`.
  Functions and lock types are also re-exported from the crate root and prelude.
- Explicit, nonzero caller-supplied timeouts for standalone Shell/COM mutations
  and refreshes. Timeouts bound supported waits, do not cancel native work,
  and do not trigger automatic mutation retries.
- Owned `RecentFilesLock` and `FrequentFoldersLock` guards that share mutation
  gates with managers and other wincent processes. Explicit unlock removes and
  returns new Recent shortcut paths; dropping a guard releases it without cleanup.
- Comprehensive VM-only acceptance tests for manager and standalone APIs on
  target Windows 10 and Windows 11 builds, covering independent Shell query
  comparisons, status, native mutations, visibility, scoped clearing, lock
  lifetime, and cross-process mutation denial.
- Standalone unit tests for timeout forwarding, input and error contracts,
  pinned versus unpinned removal, and lock release after cleanup failure.
- VM runners with guest identity and binary hash checks, source and evidence
  archives, file/registry backup verification, and whole-VM snapshot rollback.
- A sequential `examples/standalone.rs` application demonstrating queries,
  additions, removals, locks, hiding, and restoration of visibility settings.

### Changed
- Standalone operations reuse the manager's backend filtering and native
  mutation paths without changing existing manager behavior. Pinning accepts
  existing unpinned folders and rejects already-pinned folders.
- Renamed the internal `quick_access_lock` module to `lock` and updated module
  references and cross-process test entry points.
- The VM acceptance runner executes standalone unit tests before integration
  checks and requires positive passing counts for both phases.

### Removed
- Public access to the `utils` module and `get_windows_recent_folder` helper.
  `WindowsVersion` and `get_windows_version` remain available through `prelude`.
- The crate-root `find_windows_recent_links_for_target` export; the helper is
  now a crate-private, test-only implementation detail.

### Documentation
- Linked both root READMEs to the standalone application example.
- Added English and Chinese VM acceptance guides describing coverage, recovery,
  evidence requirements, and build-specific validation limits.

## [0.3.0-alpha.1] - 2026-11-06

### Added
- Kind-aware DestList parsing for Recent Files and Frequent Folders, including
  stream-backed visibility candidates and bounded Shell-query results
- `QuickAccessManager::get_item_status()` for checking a path's Quick Access
  status, including Frequent Folder pin state, with `QuickAccessStatus` and
  `FrequentFolderSlot`

### Changed
- `QuickAccessManager::remove_item` accepts only a path and automatically
  selects Recent Files, unpinned Frequent Folders, or pinned Frequent Folders
  removal.
- Visible unpinned Frequent Folders now use the native `removefromhome` Shell
  verb; pinned folders use `unpinfromhome`, while Recent files continue to use
  `remove`. All three paths now match Explorer's native UI removal behavior.
- `QuickAccessManager::lock(qa_type)` and `unlock()` now own one lock per manager
  and its clones. Unlock returns paths of new Recent `.lnk` files and attempts
  to remove them; independent managers still rely on backing-file lock conflicts.
- Visibility queries and setters use `VisibilityTarget`, including
  `StartRecommended`. `VisibilityTarget::All` affects only Explorer Recent Files
  and Frequent Folders. Setters write current-user registry values without
  guaranteeing an immediate UI change.
- DestList v4 Frequent Folders now require `access_count >= 3` for unpinned
  visibility candidates, matching Windows 10 1809 and 22H2 observations
- Zero-length DestList streams are represented as an uninitialized empty state;
  non-empty streams shorter than the 32-byte header are rejected as malformed
- DestList parsing now supports only persisted versions 4 and 6; persisted
  versions below 4 return `DestListUnsupportedVersion`
- Add/remove mutations execute once without automatic retry. Their documentation
  now distinguishes successful Shell invocation, persisted DestList state, and
  the currently rendered Explorer UI. Mutation timeouts are explicitly
  documented as indeterminate outcomes.
- Recent File removal now uses only the native Shell COM `remove` verb, without
  a fallback implementation.

### Removed
- `QuickAccessManager::watch_quick_access`, the `monitor` module and its
  `QuickAccessMonitorOptions`, `QuickAccessMonitor`, and
  `QuickAccessChangeEvent` types, plus the example CLI `watch` command.
  Applications can poll `get_items` using their own scheduling and shutdown policy.
- Visibility `show`/`hide` and category-specific aliases, `VisibilityOptions`,
  `*_with_options` methods, and the visibility-only post-mutation error.
  Use `set_visible(VisibilityTarget, bool)` instead.
- The separate Start Recommended visibility methods; use
  `is_visible(VisibilityTarget::StartRecommended)` and
  `set_visible(VisibilityTarget::StartRecommended, visible)` instead.
- Category-specific manager lock methods and public `QuickAccessLock` and
  unlock report types; use the manager's `lock(qa_type)` and `unlock()` methods.
- Public `RetryPolicy` configuration, `QuickAccessManagerBuilder::retry_policy`,
  `QuickAccessManagerBuilder::try_build`, and the example CLI `retry` command
- PowerShell script generation, caching, execution, and public error types
- Legacy live Shell query helpers and the Explorer Desktop-navigation refresh
  experiment
- Production-unreachable generic Frequent Folder removal compatibility helpers

## [0.2.6] - 2026-07-09

### Added
- `QuickAccessManager::frequent_folder_pin_status()` for reading Frequent Folders pinned, unpinned, or absent state from DestList metadata
- `QuickAccessManager::find_items_containing()` for returning fuzzy Quick Access matches instead of only a boolean
- `QuickAccessManager::watch_quick_access()` with `QuickAccessMonitorOptions`, `QuickAccessMonitor`, and `QuickAccessChangeEvent`
- `QuickAccessChangeEvent::is_reorder()` for detecting order-only monitor events
- Example CLI commands for Frequent Folders pin status, fuzzy finding, and watching Quick Access changes

### Changed
- Frequent Folders removal is now state-first and verification-driven, so pinned and unpinned frequent entries are handled without relying on Windows version alone
- Fuzzy matching through `contains_item()` and `find_items_containing()` is now case-insensitive
- Quick Access monitor poll interval configuration now uses `try_poll_interval()` to return `InvalidArgument` for zero intervals instead of exposing a panic setter
- README and rustdoc now describe current monitor, fuzzy search, DestList, visibility, and Frequent Folders removal behavior more accurately

### Fixed
- Removed unpinned Frequent Folders entries through Shell verbs instead of requiring destructive DestList removal
- Preserved Frequent Folders remove correctness when `pintohome` behaves as a toggle on both Windows 10 and Windows 11
- Made reorder-only monitor events visible to callers instead of requiring them to infer from empty added/removed lists
- Avoided case-sensitive misses in fuzzy Quick Access matching

### Removed
- Removed public and example CLI access to destructive experimental DestList removal helpers; the rebuild-based helpers are now test-only internals
- Removed temporary manual Shell verb diagnostics and kept their Win10/Win11 observations in implementation rustdoc

## [0.2.5] - 2026-06-22

### Added
- Start Recommended visibility APIs for controlling whether recently used files appear in Windows Recommended items
- Example CLI commands for reading, setting, showing, and hiding the Recommended recent-files visibility setting

### Changed
- Release workflow no longer passes feature flags now that former optional features are built in
- GitHub Actions workflows now use Node 24-compatible action versions
- Example CLI DestList `remove-entries` matching now uses Windows-style lightweight path comparison

### Fixed
- Example CLI now rejects unknown `--limit` options instead of silently ignoring misspelled arguments
- Example CLI now reports a clear error when `dest remove-entries` finds no matching DestList entries

### Documentation
- Clarified that Start Recommended APIs control recently used files in Windows Recommended items
- Clarified that Quick Access locks block add/remove mutations for the locked category while held
- Fixed stale rustdoc examples and cross-references that pointed at private or removed APIs
- Corrected PowerShell execution safety notes to describe path escaping and argument passing accurately

## [0.2.4] - 2026-06-18

### Added
- `RestoreDefaultsOptions::deep_lnk_cleanup()` for opt-in deletion of unresolved or unknown-type `.lnk` files during restore cleanup
- `EmptyOptions::with_pinned_folders_timeout()` and `pinned_folders_timeout()` for explicit pinned-folder snapshot/unpin timeout control
- Example CLI support for `empty --pinned-timeout-ms N`

### Changed
- DestList visible-entry deduplication now uses Windows-style path keys, covering slash variants, trailing separators, and Unicode case folding
- Dynamic PowerShell script cache files are process-isolated with `{script}_{version}_{pid}_{hash}.ps1` names to avoid cross-process cleanup races
- `.lnk` target type resolution only timeout-protects network-target metadata probes when shell link attributes are unavailable
- PowerShell unpin scripts keep the Windows 11 `pintohome` toggle fallback while separating normalize, find, wait, and invoke helpers
- Test utilities now create process-unique temporary directories instead of relying on the cargo working directory

### Fixed
- Prevented frequent-folder pin timeout fallback from toggling a folder back off after a late native COM success
- Mapped PowerShell pin "already exists" sentinel output back to `AlreadyExists`
- Preserved existing frequent-folder pin errors for timeouts and already-existing folders without invoking unsafe fallback paths
- Avoided deleting current-process dynamic PowerShell scripts during cache cleanup while still allowing expired orphan scripts to age out
- Rejected zero pinned-folder cleanup timeouts at execution time
- Added an internal debug assertion and documentation for validating custom `RetryPolicy` values before execution

### Documentation
- Normalized repository documentation and comments to UTF-8-friendly wording
- Clarified experimental DestList remove stability expectations, including backing-file rebuilds and possible `.lnk` deletion
- Updated README files and the example CLI help to match the current built-in visibility/DestList APIs and cleanup timeout behavior

## [0.2.3] - 2026-05-27

### Added
- `RemoveOptions` for optional deep cleanup after Quick Access removals
- `QuickAccessManager::remove_item_with_options` and `remove_items_batch_with_options`
- Deep cleanup of matching Windows Recent `.lnk` shortcuts for Recent Files and Frequent Folders removals
- `QuickAccessLock` APIs for locking Recent Files, Frequent Folders, or both Explorer automatic destination backing files
- Unlock reporting for initial, current, new, deleted, and failed Recent `.lnk` cleanup paths
- Example CLI support for `--deep-clean` removal and `lock [recent|frequent|all] [--cleanup-new-links]`

### Changed
- Batch remove can now run optional deep cleanup and reports cleanup failures as per-item failures
- Shared Windows wide-string conversion helper across shortcut parsing and backing-file locking code

### Fixed
- Compared lock/unlock Recent `.lnk` snapshots with Windows path semantics instead of case-sensitive `PathBuf` equality
- Made unlock cleanup best-effort so successful and failed `.lnk` deletions are both visible in the report

## [0.2.2] - 2026-05-27

### Added
- Optional `visible` feature for reading and updating Explorer Quick Access section visibility
- Experimental `destlist` rebuild-based removal helpers for deleting matching Recent Files or Frequent Folders entries
- DestList parser support for Explorer DestList versions 1, 3, 4, and 6
- Explorer navigation refresh cycle for the Recent Files page
- Interactive CLI example replacing the previous collection of small sample programs

### Changed
- Aligned public types with Rust API Guidelines by hiding implementation fields behind accessors and marking public enums as non-exhaustive
- Made `QuickAccessManager` the primary facade for lower-level Quick Access operations and hidden implementation modules
- Updated manager APIs to accept path-like inputs and use typed options instead of boolean flags
- Moved shell timeout configuration out of `BatchOptions`; batch options now focus on batch behavior only
- Improved README and rustdoc coverage, including clearer examples, failure ordering, batch preflight behavior, and force-update semantics
- Refactored Explorer refresh logic and internal PowerShell error construction

### Fixed
- Hardened PowerShell script generation and cached script handling
- Avoided localized Explorer window title matching when refreshing windows
- Reduced partial-state risks in experimental DestList removal
- Preserved partial frequent-folder clear state in failure cases
- Hardened Recent Files handling tests and fallback behavior
- Returned `UnsupportedOperation` for unknown mutation targets such as `QuickAccess::All` removal
- Clarified COM STA timeout, cleanup, and `S_FALSE` guard behavior

### Removed
- Removed misspelled public visibility API aliases
- Removed unused internal Windows utility helpers

## [0.2.1] - 2026-05-22

### Added
- Read-only `destlist` feature for parsing `.automaticDestinations-ms` Jump List files
- Metadata access for DestList entries, including pin status, rank, access count, score, and FILETIME
- Manager helpers for reading recent-files and frequent-folders metadata from Jump List backing files

### Changed
- Updated README and examples to match the synchronous API
- Updated ignore rules for local analysis artifacts

## [0.2.0] - 2026-04-30

### Added
- Native Windows API query operations
- Batch operations API for `QuickAccessManager`
- Builder pattern for configuring `QuickAccessManager`
- Retry mechanism with exponential backoff for PowerShell operations
- Localized PowerShell error handling with state consistency protection
- Exact and fuzzy existence-check semantics

### Changed
- Migrated `QuickAccessManager` back to a synchronous API
- Migrated core Quick Access operations to native COM paths with PowerShell fallbacks
- Restructured the public API and unified function naming
- Consolidated COM integration into a unified module
- Optimized Explorer refresh behavior and expanded test coverage
- Updated documentation and examples for the redesigned API

### Fixed
- Hardened script execution against PowerShell injection and stale cached scripts
- Improved non-zero PowerShell exit retry handling
- Fixed path comparison correctness and pin idempotency
- Improved `refresh_explorer_native` robustness
- Optimized `empty_items` to avoid duplicate refreshes
- Added partial success reporting for clear operations
- Overhauled test infrastructure and corrected doc comments
- Fixed spelling errors

### Removed
- Removed the deprecated feasible module

## [0.1.2] - 2025-04-12

### Added
- Async support with tokio runtime
- Operation timeout protection
- Lazy initialization using OnceCell
- Force update option for file operations
- System default items control
- Comprehensive feasibility checking system

### Changed
- Complete API redesign with async/await
- Enhanced error handling with timeout errors
- Improved script execution caching
- Better performance with batch operations
- Updated documentation with best practices
- Restructured system requirement checks

### Fixed
- Script execution deadlock issues
- Explorer refresh reliability
- Registry operation race conditions
- PowerShell script timeout handling

### Removed
- Direct registry visibility controls
- Deprecated feasibility check methods

## [0.1.1] - 2025-01-20

### Added

- Intergration test
- Module document

### Changed

- Better module management


## [0.1.0] - 2025-01-19

### Added

- Detailed check feasible function
- Add file to recent file function
- Clear quick access items function

### Changed

- Make functions synchronous
- Better error handling
- Better test coverage
- Support chinese


### Fixed

- Fix get visibility value issue
- Fix access registry issue
- Fix query script issue

### Removed

- Replace old examples
