//! Example demonstrating batch operations for QuickAccessManager
//!
//! Run with: cargo run --example batch_operations

use std::time::Instant;
use wincent::prelude::*;

#[tokio::main]
async fn main() -> WincentResult<()> {
    println!("=== QuickAccessManager Batch Operations Examples ===\n");

    let manager = QuickAccessManager::new().await?;

    // Example 1: Batch add multiple files
    println!("1. Batch adding multiple files:");
    let files_to_add = vec![
        (
            "G:\\Github\\wincent-rs\\README.md".to_string(),
            QuickAccess::RecentFiles,
        ),
        (
            "G:\\Github\\wincent-rs\\Cargo.toml".to_string(),
            QuickAccess::RecentFiles,
        ),
        (
            "G:\\Github\\wincent-rs\\LICENSE".to_string(),
            QuickAccess::RecentFiles,
        ),
    ];

    let start = Instant::now();
    let result = manager.add_items_batch(&files_to_add, true).await?;
    let duration = start.elapsed();

    println!("   ✓ Batch operation completed in {:?}", duration);
    println!("   - Succeeded: {}", result.succeeded.len());
    println!("   - Failed: {}", result.failed.len());
    println!(
        "   - Success rate: {:.1}%",
        result.success_rate() * 100.0
    );

    if !result.failed.is_empty() {
        println!("   Failed items:");
        for (path, error) in &result.failed {
            println!("     - {}: {}", path, error);
        }
    }
    println!();

    // Example 2: Batch add with mixed types
    println!("2. Batch adding mixed types (files and folders):");
    let mixed_items = vec![
        (
            "G:\\Github\\wincent-rs\\src".to_string(),
            QuickAccess::FrequentFolders,
        ),
        (
            "G:\\Github\\wincent-rs\\examples".to_string(),
            QuickAccess::FrequentFolders,
        ),
    ];

    let result = manager.add_items_batch(&mixed_items, false).await?;
    println!("   ✓ Added {} folders", result.succeeded.len());
    println!();

    // Example 3: Demonstrate error handling with invalid paths
    println!("3. Batch operation with some invalid paths:");
    let items_with_errors = vec![
        (
            "G:\\Github\\wincent-rs\\README.md".to_string(),
            QuickAccess::RecentFiles,
        ),
        (
            "Z:\\NonExistent\\File.txt".to_string(),
            QuickAccess::RecentFiles,
        ),
        (
            "C:\\Invalid\\Path\\File.txt".to_string(),
            QuickAccess::RecentFiles,
        ),
    ];

    let result = manager.add_items_batch(&items_with_errors, false).await?;

    if result.has_partial_success() {
        println!("   ✓ Partial success:");
        println!("     - Succeeded: {}", result.succeeded.len());
        println!("     - Failed: {}", result.failed.len());
        println!(
            "     - Success rate: {:.1}%",
            result.success_rate() * 100.0
        );
    }

    if !result.is_complete_success() {
        println!("   Failed items:");
        for (path, error) in &result.failed {
            println!("     - {}: {}", path, error);
        }
    }
    println!();

    // Example 4: Batch remove operations
    println!("4. Batch removing items:");
    let items_to_remove = vec![
        (
            "G:\\Github\\wincent-rs\\README.md".to_string(),
            QuickAccess::RecentFiles,
        ),
        (
            "G:\\Github\\wincent-rs\\Cargo.toml".to_string(),
            QuickAccess::RecentFiles,
        ),
    ];

    let result = manager.remove_items_batch(&items_to_remove).await?;
    println!("   ✓ Removed {} items", result.succeeded.len());
    if !result.failed.is_empty() {
        println!("   Failed to remove {} items", result.failed.len());
    }
    println!();

    // Example 5: Performance comparison
    println!("5. Performance comparison (single vs batch):");

    // Single operations
    let test_files = vec![
        "G:\\Github\\wincent-rs\\README.md",
        "G:\\Github\\wincent-rs\\Cargo.toml",
        "G:\\Github\\wincent-rs\\LICENSE",
    ];

    println!("   Testing single operations...");
    let start = Instant::now();
    for file in &test_files {
        let _ = manager
            .add_item(file, QuickAccess::RecentFiles, false)
            .await;
    }
    let single_duration = start.elapsed();
    println!("   - Single operations: {:?}", single_duration);

    // Clean up
    for file in &test_files {
        let _ = manager.remove_item(file, QuickAccess::RecentFiles).await;
    }

    // Batch operations
    println!("   Testing batch operations...");
    let batch_items: Vec<_> = test_files
        .iter()
        .map(|f| (f.to_string(), QuickAccess::RecentFiles))
        .collect();

    let start = Instant::now();
    let _ = manager.add_items_batch(&batch_items, false).await?;
    let batch_duration = start.elapsed();
    println!("   - Batch operations: {:?}", batch_duration);

    if single_duration > batch_duration {
        let improvement = (1.0 - batch_duration.as_secs_f64() / single_duration.as_secs_f64())
            * 100.0;
        println!(
            "   ✓ Batch is {:.1}% faster than single operations",
            improvement
        );
    }
    println!();

    // Clean up batch items
    let _ = manager.remove_items_batch(&batch_items).await?;

    println!("=== All examples completed successfully! ===");

    Ok(())
}
