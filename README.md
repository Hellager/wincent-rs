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
- Add and remove items with duplicate detection, including pinned and unpinned frequent folders
- Clear categories with optional explicit pinned-folder cleanup
- Restore default Quick Access state with conservative `.lnk` cleanup and opt-in deep cleanup
- Check item existence by exact path or keyword
- Query whether a frequent folder is pinned, unpinned, or absent
- Batch add/remove with per-item error collection
- Caller-side timeout protection for Shell and PowerShell operations
- Visibility control for Quick Access sections
- Recent files visibility control in Windows Recommended items
- DestList metadata access

## Installation

Add the following to your `Cargo.toml`:

```toml
[dependencies]
wincent = "0.2.5"
```

## Quick Start

```rust,no_run
use wincent::prelude::*;

fn main() -> WincentResult<()> {
    let manager = QuickAccessManager::new();

    // --- Add ---
    // Add a file to Recent Files. Returns Err(AlreadyExists) if already present.
    manager.add_item(
        "C:\\Projects\\report.docx",
        QuickAccess::RecentFiles,
        AddOptions::new()
            .force_recent_files_rebuild()
            .refresh_explorer(),
    )?;

    // Pin a folder to Frequent Folders.
    manager.add_item("C:\\Projects", QuickAccess::FrequentFolders, AddOptions::new())?;

    // --- Query ---
    // List all recent files as PathBuf values.
    let recent = manager.get_item_paths(QuickAccess::RecentFiles)?;
    println!("Recent files ({}):", recent.len());
    for path in &recent {
        println!("  {}", path.display());
    }

    // Check exact membership (Windows path semantics: case-insensitive).
    let exists = manager.check_item_exact("C:\\Projects\\report.docx", QuickAccess::RecentFiles)?;
    println!("report.docx in Recent Files: {exists}");

    // Fuzzy check: any item whose path string contains the keyword, case-insensitively.
    let any_match = manager.contains_item("projects", QuickAccess::All)?;
    println!("Any Quick Access item contains 'projects': {any_match}");
    for item in manager.find_items_containing("projects", QuickAccess::All)? {
        println!("Matched Quick Access item: {item}");
    }

    // Check whether a Frequent Folders path is pinned, unpinned, or absent.
    let pin_status = manager.frequent_folder_pin_status("C:\\Projects")?;
    println!("Projects Frequent Folders pin status: {pin_status:?}");

    // Watch snapshot changes. Keep the guard alive while monitoring.
    // This reports changes from Explorer and from this crate's own operations.
    let _monitor = manager.watch_quick_access(
        QuickAccessMonitorOptions::new().with_qa_type(QuickAccess::All),
        |event| {
            if let Ok(event) = event {
                println!(
                    "Quick Access changed: +{} -{} reorder={}",
                    event.added_items().len(),
                    event.removed_items().len(),
                    event.is_reorder()
                );
            }
        },
    )?;

    // --- Remove ---
    // Remove a file. Returns Err(NotInQuickAccess) if not present.
    manager.remove_item("C:\\Projects\\report.docx", QuickAccess::RecentFiles)?;

    Ok(())
}
```

## System Requirements and Limitations

- **OS**: Windows 10 or Windows 11.
- **Rust**: 1.85.0 or later.
- **Consistency**: Quick Access state is maintained by Windows Explorer. Results may lag behind mutations by a short interval, and Explorer may rebuild state asynchronously across versions.
- **Timeouts**: Timeout limits how long the caller waits, not how long the underlying Shell or COM call runs. A timed-out COM operation may still complete and affect Explorer state.
- **Frequent Folders removal**: `remove_item(..., QuickAccess::FrequentFolders)` removes pinned folders by unpinning them. For unpinned frequent entries, it pins first and then unpins through Explorer Shell verbs; if the second step fails, the folder may remain pinned.
- **Pinned-folder cleanup timeout**: when explicitly removing visible pinned folders during an `empty` operation, `EmptyOptions::with_pinned_folders_timeout()` can override the snapshot/unpin timeout. If unset, the operation uses the manager timeout.
- **Restore cleanup**: default restore cleanup deletes only `.lnk` files whose target type is resolved as the requested file or folder category. Use `RestoreDefaultsOptions::deep_lnk_cleanup()` or CLI `restore --deep` to also delete unresolved or unknown-type `.lnk` files.
- **Start menu Recommended visibility**: the Start Recommended APIs target the Windows 11 Start menu. They write the current user's `Explorer\Advanced\Start_TrackDocs` value and are not blocked on Windows 10, where the setting may have no visible Start menu effect. On some Windows 11 builds, hiding Start Recommended can also make Explorer Recent Files invisible; if `ShowRecent` alone does not restore Recent Files visibility, show Start Recommended first and then control Recent Files with `ShowRecent`.

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
- [Grant Funtila](https://www.ninjaone.com/blog/clear-the-recommended-section-in-the-start-menu-in-windows-11/)

## License

Distributed under the MIT License. See `LICENSE` for more information.

## Author

Developed with 🦀 by [@Hellager](https://github.com/Hellager)

[img_doc]: https://img.shields.io/badge/doc-latest-orange
[doc]: https://docs.rs/wincent/latest/wincent/
