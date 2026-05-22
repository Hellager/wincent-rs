//! Example demonstrating the Builder pattern for QuickAccessManager.
//!
//! Run with: cargo run --example builder_pattern

use std::time::Duration;
use wincent::prelude::*;

fn main() -> WincentResult<()> {
    println!("=== QuickAccessManager Builder Pattern Examples ===\n");

    println!("1. Creating manager with default configuration:");
    let manager_default = QuickAccessManager::new();
    println!("   Manager created with default settings\n");

    println!("2. Creating manager with custom timeout (30 seconds):");
    let _manager_custom_timeout = QuickAccessManager::builder()
        .timeout(Duration::from_secs(30))
        .build();
    println!("   Manager created with 30s timeout\n");

    println!("3. Creating manager with custom options:");
    let _manager_combined = QuickAccessManager::builder()
        .timeout(Duration::from_secs(20))
        .build();
    println!("   Manager created with 20s timeout\n");

    println!("4. Querying Quick Access items:");
    let items = manager_default.get_items(QuickAccess::All)?;
    println!("   Found {} items in Quick Access", items.len());
    if let Some(first) = items.first() {
        println!("   First item: {}", first);
    }
    println!();

    println!("=== All examples completed successfully! ===");

    Ok(())
}
