//! Windows Quick Access item retrieval and inspection
//!
//! Provides read-only access to system Quick Access metadata including:
//! - Recent file tracking
//! - Frequent folder usage
//! - Combined access patterns
//!
//! # Key Functionality
//! - Full Quick Access inventory retrieval
//! - Category-specific queries
//! - Path existence verification
//! - PowerShell-based data collection
//!
//! # Data Characteristics
//! - Updated automatically by Windows Explorer
//! - Contains user-specific activity data
//! - Maximum 20 items per category (Windows default)

use crate::{error::WincentError, QuickAccess, WincentResult};
use windows::core::VARIANT;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER,
    COINIT_APARTMENTTHREADED,
};
use windows::Win32::UI::Shell::{Folder, FolderItem, FolderItems, IShellDispatch, Shell};

/// Shell namespace for frequent folders
pub(crate) const FREQUENT_FOLDERS_NAMESPACE: &str = "shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}";

pub(crate) struct ComGuard;

impl ComGuard {
    pub(crate) fn initialize() -> WincentResult<Self> {
        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED)
                .ok()
                .map_err(|e| {
                    WincentError::SystemError(format!("Failed to initialize COM: {}", e))
                })?;
        }

        Ok(Self)
    }
}

impl Drop for ComGuard {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

enum QueryFilter {
    All,
    FilesOnly,
    FoldersOnly,
}

fn query_namespace_for(qa_type: QuickAccess) -> (&'static str, QueryFilter) {
    match qa_type {
        QuickAccess::All => (
            "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}",
            QueryFilter::All,
        ),
        QuickAccess::RecentFiles => (
            "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}",
            QueryFilter::FilesOnly,
        ),
        QuickAccess::FrequentFolders => (
            "shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}",
            QueryFilter::FoldersOnly,
        ),
    }
}

fn should_keep_item(item: &FolderItem, filter: &QueryFilter) -> bool {
    unsafe {
        match filter {
            QueryFilter::All => true,
            QueryFilter::FilesOnly => item
                .IsFolder()
                .map(bool::from)
                .map(|is_folder| !is_folder)
                .unwrap_or(false),
            QueryFilter::FoldersOnly => item.IsFolder().map(bool::from).unwrap_or(false),
        }
    }
}

pub(crate) fn item_path(item: &FolderItem) -> Option<String> {
    unsafe {
        item.Path().ok().and_then(|path| {
            let value = path.to_string();
            if value.trim().is_empty() {
                None
            } else {
                Some(value)
            }
        })
    }
}

pub(crate) fn folder_items(folder: &Folder) -> WincentResult<FolderItems> {
    unsafe {
        folder.Items().map_err(|e| {
            WincentError::SystemError(format!("Failed to enumerate Quick Access items: {}", e))
        })
    }
}

pub(crate) fn shell_dispatch() -> WincentResult<IShellDispatch> {
    unsafe {
        CoCreateInstance(&Shell, None, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER).map_err(|e| {
            WincentError::SystemError(format!("Failed to create Shell COM object: {}", e))
        })
    }
}

pub(crate) fn shell_folder(namespace: &str) -> WincentResult<Folder> {
    let shell = shell_dispatch()?;
    let variant = VARIANT::from(namespace);

    unsafe {
        shell.NameSpace(&variant).map_err(|e| {
            WincentError::SystemError(format!(
                "Failed to open shell namespace {}: {}",
                namespace, e
            ))
        })
    }
}

pub(crate) fn query_recent_native(qa_type: QuickAccess) -> WincentResult<Vec<String>> {
    let _com = ComGuard::initialize()?;
    let (namespace, filter) = query_namespace_for(qa_type);
    let folder = shell_folder(namespace)?;
    let items = folder_items(&folder)?;
    let count = unsafe {
        items.Count().map_err(|e| {
            WincentError::SystemError(format!("Failed to read Quick Access item count: {}", e))
        })?
    };

    let mut paths = Vec::with_capacity(count.max(0) as usize);

    for index in 0..count {
        let index_variant = VARIANT::from(index);
        let item = unsafe {
            match items.Item(&index_variant) {
                Ok(item) => item,
                Err(_) => continue,
            }
        };

        if should_keep_item(&item, &filter) {
            if let Some(path) = item_path(&item) {
                paths.push(path);
            }
        }
    }

    Ok(paths)
}

/// Queries recent items from Quick Access using PowerShell scripts (for comparison/fallback).
fn query_recent_powershell(qa_type: QuickAccess) -> WincentResult<Vec<String>> {
    use crate::script_executor::ScriptExecutor;
    use crate::script_storage::ScriptStorage;
    use crate::script_strategy::PSScript;

    let script_type = match qa_type {
        QuickAccess::RecentFiles => PSScript::QueryRecentFile,
        QuickAccess::FrequentFolders => PSScript::QueryFrequentFolder,
        QuickAccess::All => PSScript::QueryQuickAccess,
    };

    let start = std::time::Instant::now();
    let script_path = ScriptStorage::get_script_path(script_type)?;
    let output = ScriptExecutor::execute_ps_script(script_type, None)?;
    let duration = start.elapsed();

    ScriptExecutor::parse_output_to_strings(output, script_type, script_path, None, duration)
}

/// Queries recent items from Quick Access using the native Shell COM API or PowerShell fallback.
///
/// This function attempts to query using native COM API first, falling back to PowerShell if COM fails.
/// The native approach is significantly faster (~10-50ms vs ~200-500ms).
pub(crate) fn query_recent(qa_type: QuickAccess) -> WincentResult<Vec<String>> {
    // Try native COM first (fast path)
    let qa_type_clone = qa_type.clone();
    query_recent_native(qa_type).or_else(|_| {
        // Fallback to PowerShell if COM fails (compatibility)
        query_recent_powershell(qa_type_clone)
    })
}

/****************************************************** Query Quick Access ******************************************************/

/// Gets a list of recent files from Windows Quick Access.
///
/// # Returns
///
/// Returns a vector of file paths as strings.
///
/// # Example
///
/// ```rust
/// use wincent::{query::get_recent_files, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     let recent_files = get_recent_files()?;
///     for file in recent_files {
///         println!("Recent file: {}", file);
///     }
///     Ok(())
/// }
/// ```
pub fn get_recent_files() -> WincentResult<Vec<String>> {
    query_recent(QuickAccess::RecentFiles)
}

/// Gets a list of frequent folders from Windows Quick Access.
///
/// # Returns
///
/// Returns a vector of folder paths as strings.
///
/// # Example
///
/// ```rust
/// use wincent::{query::get_frequent_folders, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     let folders = get_frequent_folders()?;
///     for folder in folders {
///         println!("Frequent folder: {}", folder);
///     }
///     Ok(())
/// }
/// ```
pub fn get_frequent_folders() -> WincentResult<Vec<String>> {
    query_recent(QuickAccess::FrequentFolders)
}

/// Gets a list of all items from Windows Quick Access, including both recent files and frequent folders.
///
/// # Returns
///
/// Returns a vector of strings containing the paths of all Quick Access items.
///
/// # Example
///
/// ```rust
/// use wincent::{query::get_quick_access_items, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     match get_quick_access_items() {
///         Ok(items) => {
///             println!("Found {} Quick Access items:", items.len());
///             for item in items {
///                 println!("  - {}", item);
///             }
///         },
///         Err(e) => println!("Failed to get Quick Access items: {}", e)
///     }
///     Ok(())
/// }
/// ```
pub fn get_quick_access_items() -> WincentResult<Vec<String>> {
    query_recent(QuickAccess::All)
}

/****************************************************** Check Quick Access ******************************************************/

/// Checks if an exact file path exists in the Windows Recent Files list.
///
/// This function performs exact path comparison. For partial/fuzzy matching,
/// use `is_in_recent_files()` instead.
///
/// # Arguments
///
/// * `path` - The exact file path to search for
///
/// # Returns
///
/// Returns `true` if the exact path is found in the recent files list.
///
/// # Example
///
/// ```rust
/// use wincent::{query::is_recent_file_exact, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Exact match only
///     let exists = is_recent_file_exact("C:\\Users\\Documents\\file.txt")?;
///     if exists {
///         println!("Exact file path found in recent files");
///     }
///     Ok(())
/// }
/// ```
pub fn is_recent_file_exact(path: &str) -> WincentResult<bool> {
    let items = get_recent_files()?;
    Ok(items.iter().any(|item| item == path))
}

/// Checks if a file path or keyword exists in the Windows Recent Files list.
///
/// **Note**: This function performs substring matching (fuzzy match). If you need
/// exact path matching, use `is_recent_file_exact()` instead.
///
/// # Arguments
///
/// * `keyword` - The file path or partial path to search for (substring match)
///
/// # Returns
///
/// Returns `true` if any recent file path contains the keyword.
///
/// # Example
///
/// ```rust
/// use wincent::{query::is_in_recent_files, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Fuzzy match - matches any path containing "Documents"
///     let file_exists = is_in_recent_files("Documents")?;
///
///     // This will match paths like:
///     // - "C:\\Users\\Documents\\report.docx"
///     // - "D:\\My Documents\\file.txt"
///
///     if file_exists {
///         println!("File found in recent files");
///     }
///     Ok(())
/// }
/// ```
pub fn is_in_recent_files(keyword: &str) -> WincentResult<bool> {
    let items = get_recent_files()?;

    Ok(items.iter().any(|item| item.contains(keyword)))
}

/// Checks if an exact folder path exists in the Windows Frequent Folders list.
///
/// This function performs exact path comparison. For partial/fuzzy matching,
/// use `is_in_frequent_folders()` instead.
///
/// # Arguments
///
/// * `path` - The exact folder path to search for
///
/// # Returns
///
/// Returns `true` if the exact path is found in the frequent folders list.
///
/// # Example
///
/// ```rust
/// use wincent::{query::is_frequent_folder_exact, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     let folder_exists = is_frequent_folder_exact("C:\\Users\\Documents")?;
///     if folder_exists {
///         println!("Exact folder path found in frequent folders list");
///     }
///     Ok(())
/// }
/// ```
pub fn is_frequent_folder_exact(path: &str) -> WincentResult<bool> {
    let items = get_frequent_folders()?;
    Ok(items.iter().any(|item| item == path))
}

/// Checks if a folder path or keyword exists in the Windows Frequent Folders list.
///
/// **Note**: This function performs substring matching (fuzzy match). If you need
/// exact path matching, use `is_frequent_folder_exact()` instead.
///
/// # Arguments
///
/// * `keyword` - The folder path or partial path to search for (substring match)
///
/// # Returns
///
/// Returns `true` if any frequent folder path contains the keyword.
///
/// # Example
///
/// ```rust
/// use wincent::{query::is_in_frequent_folders, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Fuzzy match - matches any path containing "Projects"
///     let folder_exists = is_in_frequent_folders("Projects")?;
///     if folder_exists {
///         println!("Found folder in frequent folders list");
///     } else {
///         println!("Folder not found in frequent folders list");
///     }
///     Ok(())
/// }
/// ```
pub fn is_in_frequent_folders(keyword: &str) -> WincentResult<bool> {
    let items = get_frequent_folders()?;

    Ok(items.iter().any(|item| item.contains(keyword)))
}

/// Checks if an exact path exists in the Windows Quick Access list.
///
/// This function performs exact path comparison. For partial/fuzzy matching,
/// use `is_in_quick_access()` instead.
///
/// # Arguments
///
/// * `path` - The exact path to search for
///
/// # Returns
///
/// Returns `true` if the exact path is found in either recent files or frequent folders.
///
/// # Example
///
/// ```rust
/// use wincent::{query::is_in_quick_access_exact, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     let exists = is_in_quick_access_exact("C:\\Users\\Documents\\file.txt")?;
///     if exists {
///         println!("Exact path found in Quick Access");
///     }
///     Ok(())
/// }
/// ```
pub fn is_in_quick_access_exact(path: &str) -> WincentResult<bool> {
    let items = get_quick_access_items()?;
    Ok(items.iter().any(|item| item == path))
}

/// Checks if a path or keyword exists in the Windows Quick Access list.
///
/// **Note**: This function performs substring matching (fuzzy match). If you need
/// exact path matching, use `is_in_quick_access_exact()` instead.
///
/// # Arguments
///
/// * `keyword` - The path or partial path to search for (substring match)
///
/// # Returns
///
/// Returns `true` if any path in recent files or frequent folders contains the keyword.
///
/// # Example
///
/// ```rust
/// use wincent::{query::is_in_quick_access, error::WincentError};
///
/// fn main() -> Result<(), WincentError> {
///     // Fuzzy match - check for items containing "Documents"
///     if is_in_quick_access("Documents")? {
///         println!("Found item in Quick Access");
///     }
///
///     // Check for items in a specific location
///     if is_in_quick_access("C:\\Projects\\")? {
///         println!("Found items from Projects folder");
///     }
///     Ok(())
/// }
/// ```
pub fn is_in_quick_access(keyword: &str) -> WincentResult<bool> {
    let items = get_quick_access_items()?;

    Ok(items.iter().any(|item| item.contains(keyword)))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_query_recent_files() -> WincentResult<()> {
        let files = query_recent(QuickAccess::RecentFiles)?;

        if !files.is_empty() {
            assert!(
                files.iter().all(|path| !path.is_empty()),
                "Paths should not be empty"
            );

            for path in &files {
                assert!(
                    path.contains(":\\") || path.starts_with("\\\\"),
                    "Path should be a valid Windows path format: {}",
                    path
                );
            }
        }

        Ok(())
    }

    #[test]
    fn test_query_frequent_folders() -> WincentResult<()> {
        let folders = query_recent(QuickAccess::FrequentFolders)?;

        if !folders.is_empty() {
            assert!(
                folders.iter().all(|path| !path.is_empty()),
                "Paths should not be empty"
            );

            for path in &folders {
                assert!(
                    path.contains(":\\") || path.starts_with("\\\\"),
                    "Path should be a valid Windows path format: {}",
                    path
                );
            }
        }

        Ok(())
    }

    #[test_log::test]
    fn test_query_quick_access() -> WincentResult<()> {
        let items = query_recent(QuickAccess::All)?;

        if !items.is_empty() {
            assert!(
                items.iter().all(|path| !path.is_empty()),
                "Paths should not be empty"
            );

            for path in &items {
                assert!(
                    path.contains(":\\") || path.starts_with("\\\\"),
                    "Path should be a valid Windows path format: {}",
                    path
                );
            }
        }

        Ok(())
    }

    #[test]
    fn test_exact_vs_fuzzy_matching() -> WincentResult<()> {
        let items = query_recent(QuickAccess::All)?;

        if let Some(full_path) = items.first() {
            // exact match with full path should succeed
            assert!(
                is_in_quick_access_exact(full_path)?,
                "exact match should find full path"
            );
            assert!(
                is_in_quick_access(full_path)?,
                "fuzzy match should also find full path"
            );

            // exact match with partial path should fail
            if full_path.len() > 3 {
                let partial = &full_path[..full_path.len() - 1];
                assert!(
                    !is_in_quick_access_exact(partial)?,
                    "exact match should not find partial path"
                );
                // fuzzy match with partial path should succeed
                assert!(
                    is_in_quick_access(partial)?,
                    "fuzzy match should find partial path"
                );
            }
        }

        // non-existent path should return false for both
        let non_existent = "Z:\\Invalid\\Path\\Test.txt";
        assert!(!is_in_quick_access_exact(non_existent)?);
        assert!(!is_in_quick_access(non_existent)?);

        Ok(())
    }

    #[test]
    fn test_native_vs_powershell_results_recent_files() -> WincentResult<()> {
        let native_results = query_recent_native(QuickAccess::RecentFiles)?;
        let powershell_results = query_recent_powershell(QuickAccess::RecentFiles)?;

        println!("\n=== Native API vs PowerShell Results Comparison (Recent Files) ===");
        println!("Native API returned {} items", native_results.len());
        println!("PowerShell returned {} items", powershell_results.len());

        // Both should return the same items (order may differ)
        assert_eq!(
            native_results.len(),
            powershell_results.len(),
            "Native API and PowerShell should return the same number of items"
        );

        // Check that all items from native API exist in PowerShell results
        for item in &native_results {
            assert!(
                powershell_results.contains(item),
                "Native API item '{}' not found in PowerShell results",
                item
            );
        }

        // Check that all items from PowerShell exist in native API results
        for item in &powershell_results {
            assert!(
                native_results.contains(item),
                "PowerShell item '{}' not found in Native API results",
                item
            );
        }

        println!("✓ Results match perfectly");

        Ok(())
    }

    #[test]
    fn test_native_vs_powershell_results_frequent_folders() -> WincentResult<()> {
        let native_results = query_recent_native(QuickAccess::FrequentFolders)?;
        let powershell_results = query_recent_powershell(QuickAccess::FrequentFolders)?;

        println!("\n=== Native API vs PowerShell Results Comparison (Frequent Folders) ===");
        println!("Native API returned {} items", native_results.len());
        println!("PowerShell returned {} items", powershell_results.len());

        assert_eq!(
            native_results.len(),
            powershell_results.len(),
            "Native API and PowerShell should return the same number of items"
        );

        for item in &native_results {
            assert!(
                powershell_results.contains(item),
                "Native API item '{}' not found in PowerShell results",
                item
            );
        }

        for item in &powershell_results {
            assert!(
                native_results.contains(item),
                "PowerShell item '{}' not found in Native API results",
                item
            );
        }

        println!("✓ Results match perfectly");

        Ok(())
    }

    #[test]
    fn test_native_vs_powershell_results_all() -> WincentResult<()> {
        let native_results = query_recent_native(QuickAccess::All)?;
        let powershell_results = query_recent_powershell(QuickAccess::All)?;

        println!("\n=== Native API vs PowerShell Results Comparison (All Items) ===");
        println!("Native API returned {} items", native_results.len());
        println!("PowerShell returned {} items", powershell_results.len());

        assert_eq!(
            native_results.len(),
            powershell_results.len(),
            "Native API and PowerShell should return the same number of items"
        );

        for item in &native_results {
            assert!(
                powershell_results.contains(item),
                "Native API item '{}' not found in PowerShell results",
                item
            );
        }

        for item in &powershell_results {
            assert!(
                native_results.contains(item),
                "PowerShell item '{}' not found in Native API results",
                item
            );
        }

        println!("✓ Results match perfectly");

        Ok(())
    }

    #[test]
    #[ignore = "Performance benchmark - run manually with: cargo test test_native_vs_powershell_performance -- --ignored --nocapture"]
    fn test_native_vs_powershell_performance() -> WincentResult<()> {
        use std::time::Instant;

        println!("\n=== Native API vs PowerShell Performance Comparison ===\n");

        // Test Recent Files
        println!("Testing Recent Files Query:");
        let start = Instant::now();
        let native_files = query_recent_native(QuickAccess::RecentFiles)?;
        let native_duration = start.elapsed();
        println!("  Native API: {:?} ({} items)", native_duration, native_files.len());

        let start = Instant::now();
        let ps_files = query_recent_powershell(QuickAccess::RecentFiles)?;
        let ps_duration = start.elapsed();
        println!("  PowerShell: {:?} ({} items)", ps_duration, ps_files.len());

        let speedup = ps_duration.as_secs_f64() / native_duration.as_secs_f64();
        println!("  Speedup: {:.2}x", speedup);
        println!("  Time saved: {:?}\n", ps_duration.saturating_sub(native_duration));

        // Test Frequent Folders
        println!("Testing Frequent Folders Query:");
        let start = Instant::now();
        let native_folders = query_recent_native(QuickAccess::FrequentFolders)?;
        let native_duration = start.elapsed();
        println!("  Native API: {:?} ({} items)", native_duration, native_folders.len());

        let start = Instant::now();
        let ps_folders = query_recent_powershell(QuickAccess::FrequentFolders)?;
        let ps_duration = start.elapsed();
        println!("  PowerShell: {:?} ({} items)", ps_duration, ps_folders.len());

        let speedup = ps_duration.as_secs_f64() / native_duration.as_secs_f64();
        println!("  Speedup: {:.2}x", speedup);
        println!("  Time saved: {:?}\n", ps_duration.saturating_sub(native_duration));

        // Test All Items
        println!("Testing All Quick Access Items Query:");
        let start = Instant::now();
        let native_all = query_recent_native(QuickAccess::All)?;
        let native_duration = start.elapsed();
        println!("  Native API: {:?} ({} items)", native_duration, native_all.len());

        let start = Instant::now();
        let ps_all = query_recent_powershell(QuickAccess::All)?;
        let ps_duration = start.elapsed();
        println!("  PowerShell: {:?} ({} items)", ps_duration, ps_all.len());

        let speedup = ps_duration.as_secs_f64() / native_duration.as_secs_f64();
        println!("  Speedup: {:.2}x", speedup);
        println!("  Time saved: {:?}\n", ps_duration.saturating_sub(native_duration));

        // Consistency test (10 runs)
        println!("=== Consistency Test (10 runs of All Items) ===\n");

        let mut native_times = Vec::new();
        let mut ps_times = Vec::new();

        for i in 1..=10 {
            let start = Instant::now();
            let _ = query_recent_native(QuickAccess::All)?;
            let native_time = start.elapsed();
            native_times.push(native_time);

            let start = Instant::now();
            let _ = query_recent_with_powershell(QuickAccess::All)?;
            let ps_time = start.elapsed();
            ps_times.push(ps_time);

            println!("Run {}: Native {:?}, PowerShell {:?}", i, native_time, ps_time);
        }

        // Calculate statistics
        let native_avg: f64 = native_times.iter().map(|d| d.as_secs_f64()).sum::<f64>() / 10.0;
        let ps_avg: f64 = ps_times.iter().map(|d| d.as_secs_f64()).sum::<f64>() / 10.0;

        let native_min = native_times.iter().min().unwrap();
        let native_max = native_times.iter().max().unwrap();
        let ps_min = ps_times.iter().min().unwrap();
        let ps_max = ps_times.iter().max().unwrap();

        println!("\nNative API Statistics:");
        println!("  Average: {:.2}ms", native_avg * 1000.0);
        println!("  Min: {:?}", native_min);
        println!("  Max: {:?}", native_max);

        println!("\nPowerShell Statistics:");
        println!("  Average: {:.2}ms", ps_avg * 1000.0);
        println!("  Min: {:?}", ps_min);
        println!("  Max: {:?}", ps_max);

        println!("\nOverall Speedup: {:.2}x", ps_avg / native_avg);

        Ok(())
    }
}
