# wincent-rs

## Overview

Wincent is a rust library for managing Windows quick access functionality, providing comprehensive control over your file system's quick access content.

## Features

- 🔍 Query Quick Access Contents
- 🗑️ Remove Specific Quick Access Entries
- 👁️ Toggle Visibility of Quick Access Items

## Installation

Add the following to your `Cargo.toml`:

```toml
[dependencies]
wincent = "*"
```

## Quick Start

### Querying Quick Access Contents

```rust
use wincent::get_quick_access_items;

fn main() -> Result<(), Box<dyn std::error::Error>> {    
    // List all current quick access items
    let quick_access_items = get_quick_access_items()?;
    for item in quick_access_items {
        println!("Quick Access items: {:?}", item);
    }

    Ok(())
}
```

### Removing a Quick Access Entry

```rust
use wincent::remove_from_recent_files;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Remove a specific path from quick access
    match remove_from_recent_files("C:\\path\\to\\file.txt").await {
        Ok(()) => println!("File removed from recent files."),
        Err(e) => eprintln!("Error removing file from recent files: {:?}", e),
    }

    Ok(())
}
```

### Toggling Visibility

```rust
use wincent::{QuickAccess, set_visiable};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // hide Frequent Folders
    match set_visible(QuickAccess::FrequentFolders, true) {
        Ok(_) => println!("Successfully set visibility for Frequent Folders."),
        Err(e) => eprintln!("Failed to set visibility: {:?}", e),
    }

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

## Roadmap

- [ ] Test on more windows version
- [ ] Better way to interact with quick access

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