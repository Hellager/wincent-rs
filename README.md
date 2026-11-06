# wincent-rs

[![](https://img.shields.io/crates/v/wincent.svg)](https://crates.io/crates/wincent)
[![][img_doc]][doc]
![Crates.io Total Downloads](https://img.shields.io/crates/d/wincent)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Hellager/wincent-rs/publish.yml)
![Crates.io License](https://img.shields.io/crates/l/wincent)

Read this in other languages: [English](README.md) | [中文](README.cn.md)

## Overview

Wincent is a Rust library for managing Windows Quick Access functionality. Queries parse Explorer's DestList files directly and apply rules recovered from observed Explorer behavior. Mutations use `SHAddToRecentDocs` or native Shell COM.

## Features

- Query recent files and frequent folders directly from DestList backing data
- Add and remove items with duplicate detection, including pinned and unpinned frequent folders
- Clear Recent Files, Frequent Folders, or both categories
- Check item existence by exact path or keyword
- Query persisted Recent Files positions and Frequent Folders positions/pin state
- Caller-side timeout protection for native Shell mutations and refreshes
- Visibility control for Quick Access sections
- Recent files visibility control in Windows Recommended items
- DestList parsing and diagnostics

## Installation

Add the following to your `Cargo.toml`:

```toml
[dependencies]
wincent = "0.2.6"
```

## Quick Start

```rust,no_run
use wincent::prelude::*;

fn main() -> WincentResult<()> {
    let manager = QuickAccessManager::new();

    // --- Add ---
    // Add a file to Recent Files. Ok means the Shell mutation returned successfully;
    // Explorer may persist and display the change later.
    manager.add_item("C:\\Projects\\report.docx")?;

    // Pin a folder to Frequent Folders.
    manager.add_item("C:\\Projects")?;

    // --- Query ---
    // List all recent files as PathBuf values.
    let recent = manager.get_item_paths(QuickAccess::RecentFiles)?;
    println!("Recent files ({}):", recent.len());
    for path in &recent {
        println!("  {}", path.display());
    }

    // Check exact membership (Windows path semantics: case-insensitive).
    let exists = manager.contains_path("C:\\Projects\\report.docx", QuickAccess::RecentFiles)?;
    println!("report.docx in Recent Files: {exists}");

    // Fuzzy check: any item whose path string contains the keyword, case-insensitively.
    let any_match = manager.contains_keyword("projects", QuickAccess::All)?;
    println!("Any Quick Access item contains 'projects': {any_match}");
    for item in manager.find_items_containing("projects", QuickAccess::All)? {
        println!("Matched Quick Access item: {item}");
    }

    // Predict the default Shell query result from persisted DestList data:
    // zero-based Recent Files position, zero-based Frequent Folders position
    // and pin state, or neither.
    let status = manager.get_item_status("C:\\Projects\\report.docx")?;
    if let Some(position) = status.recent_file() {
        println!("Recent Files position: {position}");
    }
    if let Some(slot) = status.frequent_folder() {
        println!("Frequent Folders position {} pinned {}", slot.position(), slot.is_pinned());
    }

    // Notify the Shell first, then refresh matching Quick Access/Home views.
    manager.refresh()?;

    // --- Remove ---
    // Remove a file. The same eventual-consistency rules apply as for add_item.
    manager.remove_item("C:\\Projects\\report.docx")?;

    Ok(())
}
```

## System Requirements and Limitations

- **OS**: Windows 10 or Windows 11.
- **Rust**: 1.85.0 or later.
- **Query semantics**: Query APIs return the Explorer UI-visible result calculated from persisted DestList data. The recovered rules cover ordering, metadata filtering, deduplication, hidden-item Folder Options, and result limits without invoking a live Shell query.
- **UI consistency**: An already-open Explorer view may temporarily differ from the persisted UI-visible result while its UI cache refreshes. Explorer may also rebuild state asynchronously across Windows versions.
- **Mutation success**: `Ok(())` from `add_item` or `remove_item` means the selected Shell mutation returned successfully. It does not guarantee that DestList persistence, a later query, and an already-open Explorer view are immediately synchronized.
- **Timeouts**: Timeout limits how long the caller waits, not how long the underlying Shell or COM call runs. A timed-out operation has an unknown final state and may still complete later. Mutations are not retried automatically; inspect state before issuing another mutation.
- **Automatic removal routing**: `remove_item(path)` detects files and directories. Recent files use the native `remove` Shell verb, unpinned Frequent Folders use `removefromhome`, and pinned Frequent Folders use `unpinfromhome`.
- **Empty behavior**: `empty_items(...)` always refreshes Explorer after a successful clear. Clearing Frequent Folders removes user-created entries, including user pins; Windows may rebuild its system-default pins.
- **Start menu Recommended visibility**: `VisibilityTarget::StartRecommended` targets recently used files in the Windows 11 Start menu via the current user's `Explorer\Advanced\Start_TrackDocs` value. The setting is not blocked on Windows 10, where it may have no visible Start menu effect. On some Windows 11 builds, hiding Start Recommended can also make Explorer Recent Files invisible; if `ShowRecent` alone does not restore Recent Files visibility, show Start Recommended first.
- **Visibility writes**: `set_visible(target, bool)` only writes current-user registry settings; success does not guarantee an immediate UI change. `is_visible(target)` reads configured values, not the rendered UI. `VisibilityTarget::All` covers only Explorer Recent Files and Frequent Folders and writes two values in sequence; the second write may fail after the first succeeds.

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
