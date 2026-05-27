# wincent-rs

[![](https://img.shields.io/crates/v/wincent.svg)](https://crates.io/crates/wincent)
[![][img_doc]][doc]
![Crates.io Total Downloads](https://img.shields.io/crates/d/wincent)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Hellager/wincent-rs/publish.yml)
![Crates.io License](https://img.shields.io/crates/l/wincent)

Read this in other languages: [English](README.md) | [中文](README.cn.md)

## Overview

Wincent is a Rust library for managing Windows Quick Access functionality. It provides a safe interface for querying, adding, removing, and clearing recent files and frequent folders, with native Windows API fast paths and PowerShell fallbacks.

## Features

- Query recent files and frequent folders
- Add and remove items with duplicate detection
- Clear categories (optionally including pinned folders)
- Check item existence by exact path or keyword
- Batch add/remove with per-item error collection
- Timeout protection for all shell operations
- Optional visibility control for Quick Access sections (`visible` feature)
- Optional DestList metadata access (`destlist` feature)

## Installation

Add the following to your `Cargo.toml`:

```toml
[dependencies]
wincent = "0.1.2"
```

Optional features:

```toml
[dependencies]
wincent = { version = "0.1.2", features = ["visible", "destlist"] }
```

## Quick Start

```rust,no_run
use wincent::prelude::*;

fn main() -> WincentResult<()> {
    let manager = QuickAccessManager::new();

    // --- Add ---
    // Add a file to Recent Files. Returns Err(AlreadyExists) if already present.
    manager.add_item("C:\\Projects\\report.docx", QuickAccess::RecentFiles, true)?;

    // Pin a folder to Frequent Folders.
    manager.add_item("C:\\Projects", QuickAccess::FrequentFolders, false)?;

    // --- Query ---
    // List all recent files.
    let recent = manager.get_items(QuickAccess::RecentFiles)?;
    println!("Recent files ({}):", recent.len());
    for path in &recent {
        println!("  {path}");
    }

    // Check exact membership (Windows path semantics: case-insensitive).
    let exists = manager.check_item_exact("C:\\Projects\\report.docx", QuickAccess::RecentFiles)?;
    println!("report.docx in Recent Files: {exists}");

    // Fuzzy check: any item whose path string contains the keyword.
    let any_match = manager.contains_item("Projects", QuickAccess::All)?;
    println!("Any Quick Access item contains 'Projects': {any_match}");

    // --- Remove ---
    // Remove a file. Returns Err(NotInRecent) if not present.
    manager.remove_item("C:\\Projects\\report.docx", QuickAccess::RecentFiles)?;

    Ok(())
}
```

## Manager API

### Construction

| Method | Description |
|--------|-------------|
| `QuickAccessManager::new()` | Create a manager with the default 10-second timeout. |
| `QuickAccessManager::builder()` | Return a `QuickAccessManagerBuilder` for custom configuration. |
| `builder.timeout(duration)` | Set the shell operation timeout (panics on zero). |
| `builder.try_timeout(duration)` | Set the timeout, returning `Err` instead of panicking on zero. |
| `builder.build()` | Build the configured `QuickAccessManager`. |

### Query

| Method | Description |
|--------|-------------|
| `get_items(qa_type)` | Return all paths in the given category (`RecentFiles`, `FrequentFolders`, or `All`). |
| `check_item_exact(path, qa_type)` | Return `true` if the path is present (Windows case-insensitive comparison). |
| `contains_item(keyword, qa_type)` | Return `true` if any item's path string contains `keyword` (case-sensitive substring). |

### Mutation

| Method | Description |
|--------|-------------|
| `add_item(path, qa_type, force_update)` | Add an item. `QuickAccess::All` is rejected. Returns `Err(AlreadyExists)` if already present. |
| `remove_item(path, qa_type)` | Remove an item. `QuickAccess::All` is rejected. Returns `Err(NotInRecent)` if absent. |
| `add_items_batch(items, force_update)` | Add multiple items, collecting failures without short-circuiting. |
| `remove_items_batch(items)` | Remove multiple items, collecting failures without short-circuiting. |
| `empty_items(qa_type, force_refresh, also_pinned_folders)` | Clear a category. When `also_pinned_folders` is `true`, pinned folder entries are also removed. |

### Optional features

| Method | Feature | Description |
|--------|---------|-------------|
| `is_visible(qa_type)` | `visible` | Read the Explorer registry flag that controls section visibility. |
| `set_visible(qa_type, visible)` | `visible` | Write the Explorer registry flag. |
| `get_recent_files_metadata()` | `destlist` | Parse the recent-files `.automaticDestinations-ms` file and return `DestListEntry` records. |
| `get_frequent_folders_metadata()` | `destlist` | Parse the frequent-folders `.automaticDestinations-ms` file and return `DestListEntry` records. |

## Error Handling

All fallible methods return `WincentResult<T>` (`Result<T, WincentError>`). Common variants:

| Variant | When raised |
|---------|-------------|
| `AlreadyExists(path)` | `add_item` called for a path already in Quick Access. |
| `NotInRecent(path)` | `remove_item` called for a path not in Quick Access. |
| `InvalidPath(path)` | A supplied path is empty, malformed, or has the wrong type. |
| `UnsupportedOperation(msg)` | `add_item`/`remove_item` called with `QuickAccess::All`. |
| `Timeout(msg)` | A shell operation exceeded the configured timeout. |
| `PartialEmpty { … }` | `empty_items` cleared some categories before a later step failed. |
| `PowerShellExecution(err)` | A PowerShell fallback script failed or could not start. |
| `SystemError(msg)` | A non-I/O Windows API call failed. |

### Handling PowerShell errors

```rust,no_run
use wincent::prelude::*;

fn main() -> WincentResult<()> {
    let manager = QuickAccessManager::new();

    match manager.add_item("C:\\folder", QuickAccess::FrequentFolders, false) {
        Ok(()) => println!("Added."),
        Err(WincentError::AlreadyExists(p)) => println!("{p} is already pinned."),
        Err(WincentError::PowerShellExecution(err)) => {
            if err.is_access_denied() {
                println!("Access denied — try running as administrator.");
            } else if err.is_execution_policy_error() {
                println!("PowerShell execution policy blocks scripts.");
            } else if let Some(fix) = err.suggest_fix() {
                println!("Suggestion: {fix}");
            } else {
                println!("Script error: {err}");
            }
        }
        Err(e) => println!("Error: {e}"),
    }

    Ok(())
}
```

## System Requirements and Limitations

- **OS**: Windows 10 or Windows 11.
- **Rust**: 1.60.0 or later.
- **PowerShell**: Required for fallback operations when the native Windows API path is unavailable. Scripts are launched with process-scoped `-ExecutionPolicy Bypass`; this does not change user or machine policy, but hardened environments may still block execution.
- **Permissions**: Most operations run under the current user's account. Pinning/unpinning folders may require the user's desktop session. Elevated permissions are not normally required.
- **Consistency**: Quick Access state is maintained by Windows Explorer. Results may lag behind mutations by a short interval, and Explorer may rebuild state asynchronously across versions.

## Best Practices

- Pass `force_update = true` to `add_item` when adding recent files for immediate visibility in Explorer.
- When clearing frequent folders, set `also_pinned_folders = true` only when you intend to remove user-pinned entries.
- Handle `PartialEmpty` explicitly if your code needs to distinguish "nothing cleared" from "partially cleared".
- Use `check_item_exact` (Windows path semantics) rather than `contains_item` (case-sensitive substring) for membership tests.

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b wincent/amazing-feature`)
3. Commit your changes (`git commit -m 'feat: Add some amazing feature'`)
4. Push to the branch (`git push origin wincent/amazing-feature`)
5. Open a Pull Request

### Development Setup

```bash
# Clone the repository
git clone https://github.com/Hellager/wincent-rs.git
cd wincent-rs

# Build and run unit tests (no Explorer session required)
cargo build
cargo test

# Run integration tests that require an interactive desktop session
cargo test -- --ignored
```

## Disclaimer

This library interacts with system-level Quick Access functionality. Always ensure you have appropriate permissions and create backups before making significant changes.

## Support

If you encounter any issues or have questions, please file an issue on our GitHub repository.

## Thanks

- [Castorix31](https://learn.microsoft.com/en-us/answers/questions/1087928/how-to-get-recent-docs-list-and-delete-some-of-the)
- [Yohan Ney](https://stackoverflow.com/questions/30051634/is-it-possible-programmatically-add-folders-to-the-windows-10-quick-access-panel)
- [libyal](https://github.com/libyal/dtformats/blob/main/documentation/Jump%20lists%20format.asciidoc)
- [Eric Zimmerman](https://github.com/EricZimmerman/JumpList)
- [kacos2000](https://github.com/kacos2000/Jumplist-Browser)

## License

Distributed under the MIT License. See `LICENSE` for more information.

## Author

Developed with 🦀 by [@Hellager](https://github.com/Hellager)

[img_doc]: https://img.shields.io/badge/doc-latest-orange
[doc]: https://docs.rs/wincent/latest/wincent/
