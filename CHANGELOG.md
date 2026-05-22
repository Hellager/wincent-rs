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
