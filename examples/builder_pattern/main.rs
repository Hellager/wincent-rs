//! Example demonstrating the Builder pattern for QuickAccessManager
//!
//! Run with: cargo run --example builder_pattern

use std::time::Duration;
use wincent::prelude::*;

#[tokio::main]
async fn main() -> WincentResult<()> {
    println!("=== QuickAccessManager Builder Pattern Examples ===\n");

    // Example 1: Default configuration (equivalent to QuickAccessManager::new())
    println!("1. Creating manager with default configuration:");
    let manager_default = QuickAccessManager::new().await?;
    println!("   ✓ Manager created with default settings\n");

    // Example 2: Custom timeout
    println!("2. Creating manager with custom timeout (30 seconds):");
    let _manager_custom_timeout = QuickAccessManager::builder()
        .timeout(Duration::from_secs(30))
        .build()
        .await?;
    println!("   ✓ Manager created with 30s timeout\n");

    // Example 3: Disable cache for real-time data
    println!("3. Creating manager with cache disabled:");
    let _manager_no_cache = QuickAccessManager::builder()
        .disable_cache()
        .build()
        .await?;
    println!("   ✓ Manager created without caching\n");

    // Example 4: Check feasibility on initialization
    println!("4. Creating manager with feasibility check:");
    match QuickAccessManager::builder()
        .check_feasibility_on_init()
        .build()
        .await
    {
        Ok(manager) => {
            println!("   ✓ Manager created and system is compatible");
            let (can_query, can_modify) = manager.check_feasible().await;
            println!("   - Can query: {}", can_query);
            println!("   - Can modify: {}", can_modify);
        }
        Err(e) => {
            println!("   ✗ System is not compatible: {}", e);
        }
    }
    println!();

    // Example 5: Combining multiple options
    println!("5. Creating manager with multiple custom options:");
    let _manager_combined = QuickAccessManager::builder()
        .timeout(Duration::from_secs(20))
        .disable_cache()
        .build()
        .await?;
    println!("   ✓ Manager created with 20s timeout and no cache\n");

    // Example 6: Using the manager to query items
    println!("6. Querying Quick Access items:");
    let items = manager_default.get_items(QuickAccess::All).await?;
    println!("   Found {} items in Quick Access", items.len());
    if !items.is_empty() {
        println!("   First item: {}", items[0]);
    }
    println!();

    println!("=== All examples completed successfully! ===");

    Ok(())
}
