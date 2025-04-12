# wincent-rs

[![](https://img.shields.io/crates/v/wincent.svg)](https://crates.io/crates/wincent)
[![][img_doc]][doc]
![Crates.io Total Downloads](https://img.shields.io/crates/d/wincent)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Hellager/wincent-rs/publish.yml)
![Crates.io License](https://img.shields.io/crates/l/wincent)

Read this in other languages: [English](README.md) | [中文](README.cn.md)

## Overview

Wincent is a Rust library for managing Windows Quick Access functionality, providing comprehensive control over recent files and frequent folders with async support and robust error handling.

## Features

- 🔍 Comprehensive Quick Access Management
  - Query recent files and frequent folders
  - Add/Remove items with force update options
  - Clear categories with system default control
  - Check item existence

- 🛡️ Robust Operation Handling
  - Operation timeout protection
  - Cached script execution
  - Comprehensive error handling

- ⚡ Performance Optimizations
  - Lazy initialization with OnceCell
  - Script execution caching
  - Batch operation support
  - Force refresh control

## Installation

Add the following to your `Cargo.toml`:

```toml
[dependencies]
wincent = "0.1.2"
```

## Quick Start

### Basic Operations

```rust
use wincent::predule::*;

#[tokio::main]
async fn main() -> WincentResult<()> {
    // Initialize manager
    let manager = QuickAccessManager::new().await?;
    
    // Add file to Recent Files with force update
    manager.add_item(
        "C:\\path\\to\\file.txt",
        QuickAccess::RecentFiles,
        true
    ).await?;
    
    // Query all items
    let items = manager.get_items(QuickAccess::All).await?;
    println!("Quick Access items: {:?}", items);
    
    Ok(())
}
```

### Advanced Usage

```rust
use wincent::predule::*;

#[tokio::main]
async fn main() -> WincentResult<()> {
    let manager = QuickAccessManager::new().await?;
    
    // Clear recent files with force refresh
    manager.empty_items(
        QuickAccess::RecentFiles,
        true,  // force refresh
        false  // preserve system defaults
    ).await?;
    
    Ok(())
}
```

## Best Practices

- Use `force_update` when adding recent files for immediate visibility
- Check operation feasibility only when necessary
- Clear cache after bulk operations
- Consider `also_system_default` carefully when clearing items
- Handle timeouts appropriately in production environments

## System Requirements and Limitations

- **API Dependencies**: The library relies on Windows system APIs for its core functionality. Due to Windows security policies and updates, certain operations may require elevated permissions or could be restricted.

- **Environment Compatibility**: Different Windows environments (including versions, configurations, and installed software) may affect the library's functionality. In particular:
  - Third-party software may modify relevant registry entries
  - System security policies may restrict PowerShell script execution
  - Windows Explorer integration might vary across different Windows versions

- **Pre-flight Checks**: Before performing operations, especially for folder management, it's recommended to use the built-in feasibility checks:
  ```rust
  let (can_query, can_modify) = manager.check_feasible().await;
  if !can_modify {
      println!("Warning: System environment may restrict modification operations");
  }
  ```

These limitations are inherent to Windows Quick Access functionality and not specific to this library. We provide comprehensive error handling and status checking mechanisms to help you handle these scenarios gracefully.

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
