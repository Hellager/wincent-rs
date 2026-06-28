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
