# wincent-rs

![Crates.io Version](https://img.shields.io/crates/v/wincent)
[![][img_doc]][doc]
![Crates.io Total Downloads](https://img.shields.io/crates/d/wincent)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Hellager/wincent-rs/publish.yml)
![Crates.io License](https://img.shields.io/crates/l/wincent)

## Overview

Wincent is a rust library for managing Windows quick access functionality, providing comprehensive control over your file system's quick access content.

## Features

- 🔍 Query Quick Access Contents
- ➕ Add Items to Quick Access
- 🗑️ Remove Specific Quick Access Entries
- 🧹 Clear Quick Access Items
- 👁️ Toggle Visibility of Quick Access Items


## Installation

Add the following to your `Cargo.toml`:

```toml
[dependencies]
wincent = "0.1.1"
```

## Quick Start

### Querying  Quick  Access  Contents

```rust
use wincent::{
    feasible::{check_feasible, fix_feasible}, 
    query::get_quick_access_items, 
    error::WincentError
};

fn main() -> Result<(), WincentError> {
    // Check if quick access is feasible
    if !check_feasible()? {
        fix_feasible()?;
    }

    // List all current quick access items
    let quick_access_items = get_quick_access_items()?;
    for item in quick_access_items {
        println!("Quick Access item: {}", item);
    }

    Ok(())
}
```

### Removing  a  Quick  Access  Entry

```rust
use wincent::{
    query::get_recent_files, 
    handle::remove_from_recent_files, 
    error::WincentError
};

fn main() -> Result<(), WincentError> {
    // Remove sensitive files from recent items
    let recent_files = get_recent_files()?;
    for item in recent_files {
        if item.contains("password") {
            remove_from_recent_files(&item)?;
        }
    }

    Ok(())
}
```

### Toggling  Visibility

```rust
use wincent::{
    visible::{is_recent_files_visiable, set_recent_files_visiable}, 
    error::WincentError
};

fn main() -> Result<(), WincentError> {
    let is_visible = is_recent_files_visiable()?;
    println!("Recent files visibility: {}", is_visible);

    set_recent_files_visiable(!is_visible)?;
    println!("Visibility toggled");

    Ok(())
}
```

## Error Handling

The library uses Rust's `Result` type for comprehensive error management, allowing precise handling of potential issues during quick access manipulation.

## Compatibility

- Supports Windows 10 and Windows 11
- Requires Rust 1.60.0 or later

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

# Install development dependencies
cargo build
cargo test
```


## Disclaimer

This library interacts with system-level Quick Access functionality. Always ensure you have appropriate permissions and create backups before making significant changes.

## Support

If you encounter any issues or have questions, please file an issue on our GitHub repository.

## Thanks

- [Castorix31](https://learn.microsoft.com/en-us/answers/questions/1087928/how-to-get-recent-docs-list-and-delete-some-of-the)
- [Yohan Ney](https://stackoverflow.com/questions/30051634/is-it-possible-programmatically-add-folders-to-the-windows-10-quick-access-panel)

## License

Distributed under the MIT License. See `LICENSE` for more information.

## Author

Developed with 🦀 by [@Hellager](https://github.com/Hellager)

[img_doc]: https://img.shields.io/badge/doc-latest-orange
[doc]: https://docs.rs/wincent/latest/wincent/