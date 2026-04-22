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
use crate::com::{ComGuard, ComInitStatus};
use windows::core::VARIANT;
use windows::Win32::System::Com::{
    CoCreateInstance, CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER,
};
use windows::Win32::UI::Shell::{Folder, FolderItem, FolderItems, IShellDispatch, Shell};

/// Shell namespace GUID for frequent folders in Windows Quick Access
///
/// This GUID corresponds to the "Frequent Folders" virtual folder in Windows Explorer.
/// It provides access to folders that Windows tracks as frequently accessed by the user.
pub(crate) const FREQUENT_FOLDERS_NAMESPACE: &str = "shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}";

/// Filter type for Quick Access item queries
///
/// Determines which types of items should be included in query results.
enum QueryFilter {
    /// Include all items (both files and folders)
    All,
    /// Include only files (exclude folders)
    FilesOnly,
    /// Include only folders (exclude files)
    FoldersOnly,
}

/// Maps a QuickAccess query type to its corresponding shell namespace and filter
///
/// # Arguments
///
/// * `qa_type` - The type of Quick Access items to query
///
/// # Returns
///
/// A tuple containing:
/// - Shell namespace GUID string
/// - Filter to apply to the results
///
/// # Shell Namespace GUIDs
///
/// - `shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}` - Recent Items (contains both files and folders)
/// - `shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}` - Frequent Folders
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

/// Determines if a FolderItem should be included based on the query filter
///
/// # Arguments
///
/// * `item` - The FolderItem to check
/// * `filter` - The filter criteria to apply
///
/// # Returns
///
/// `true` if the item matches the filter criteria, `false` otherwise
///
/// # Safety
///
/// This function calls unsafe COM methods (`IsFolder()`). The caller must ensure
/// the FolderItem is valid and COM is properly initialized.
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

/// Extracts the file system path from a FolderItem
///
/// # Arguments
///
/// * `item` - The FolderItem to extract the path from
///
/// # Returns
///
/// `Some(String)` containing the path if available and non-empty, `None` otherwise
///
/// # Safety
///
/// This function calls unsafe COM methods (`Path()`). The caller must ensure
/// the FolderItem is valid and COM is properly initialized.
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

/// Retrieves the FolderItems collection from a Folder
///
/// # Arguments
///
/// * `folder` - The Folder to enumerate items from
///
/// # Returns
///
/// A `FolderItems` collection on success, or an error if enumeration fails
///
/// # Errors
///
/// Returns `WincentError::SystemError` if the COM call to enumerate items fails
///
/// # Safety
///
/// This function calls unsafe COM methods. The caller must ensure COM is properly initialized.
pub(crate) fn folder_items(folder: &Folder) -> WincentResult<FolderItems> {
    unsafe {
        folder.Items().map_err(|e| {
            WincentError::SystemError(format!("Failed to enumerate Quick Access items: {}", e))
        })
    }
}

/// Creates a Shell COM object (IShellDispatch)
///
/// This is the entry point for accessing Windows Shell functionality through COM.
///
/// # Returns
///
/// An `IShellDispatch` interface on success, or an error if creation fails
///
/// # Errors
///
/// Returns `WincentError::SystemError` if the COM object creation fails
///
/// # Safety
///
/// This function calls unsafe COM methods. The caller must ensure COM is properly initialized.
pub(crate) fn shell_dispatch() -> WincentResult<IShellDispatch> {
    unsafe {
        CoCreateInstance(&Shell, None, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER).map_err(|e| {
            WincentError::SystemError(format!("Failed to create Shell COM object: {}", e))
        })
    }
}

/// Opens a shell namespace by its GUID
///
/// Shell namespaces are virtual folders in Windows Explorer identified by GUIDs.
/// Common namespaces include Recent Items, Frequent Folders, etc.
///
/// # Arguments
///
/// * `namespace` - Shell namespace GUID string (e.g., "shell:::{GUID}")
///
/// # Returns
///
/// A `Folder` interface representing the namespace on success, or an error if opening fails
///
/// # Errors
///
/// Returns `WincentError::SystemError` if:
/// - Shell COM object creation fails
/// - The namespace GUID is invalid or inaccessible
///
/// # Safety
///
/// This function calls unsafe COM methods. The caller must ensure COM is properly initialized.
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

/// Queries Quick Access items using the native Windows Shell COM API
///
/// This is the fast path for querying Quick Access items, typically completing in 10-50ms.
/// It directly accesses Windows Shell namespaces through COM interfaces.
///
/// # Arguments
///
/// * `qa_type` - The type of Quick Access items to query (RecentFiles, FrequentFolders, or All)
///
/// # Returns
///
/// A vector of file/folder paths on success, or an error if the query fails
///
/// # Errors
///
/// Returns errors if:
/// - COM initialization fails (apartment mismatch or other COM errors)
/// - Shell namespace access fails
/// - Item enumeration fails
///
/// # Performance
///
/// Native COM API is significantly faster than PowerShell:
/// - Native: ~10-50ms
/// - PowerShell: ~200-500ms
///
/// # Safety
///
/// This function initializes COM in STA mode and performs unsafe COM operations.
/// COM is automatically cleaned up via RAII when the function returns.
pub(crate) fn query_recent_native(qa_type: QuickAccess) -> WincentResult<Vec<String>> {
    let _com = ComGuard::try_initialize().map_err(|status| match status {
        ComInitStatus::ApartmentMismatch => {
            WincentError::ComApartmentMismatch(
                "Thread already initialized with incompatible COM apartment model".to_string()
            )
        }
        ComInitStatus::OtherError(hr) => {
            WincentError::SystemError(format!("Failed to initialize COM: 0x{:08X}", hr))
        }
        _ => unreachable!(),
    })?;
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
///
/// This is the compatibility fallback when native COM API fails or is unavailable.
/// PowerShell execution is slower but more compatible across different system configurations.
///
/// # Arguments
///
/// * `qa_type` - The type of Quick Access items to query
///
/// # Returns
///
/// A vector of file/folder paths on success, or an error if PowerShell execution fails
///
/// # Errors
///
/// Returns errors if:
/// - PowerShell script execution fails
/// - Script output parsing fails
/// - PowerShell is not available on the system
///
/// # Performance
///
/// PowerShell is significantly slower than native COM API:
/// - PowerShell: ~200-500ms
/// - Native: ~10-50ms
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
/// This function implements a two-tier strategy:
/// 1. **Fast path**: Attempts native COM API first (~10-50ms)
/// 2. **Fallback**: Falls back to PowerShell if COM fails (~200-500ms)
///
/// The fallback ensures compatibility even when COM initialization fails due to
/// apartment model mismatches or other COM-related issues.
///
/// # Arguments
///
/// * `qa_type` - The type of Quick Access items to query
///
/// # Returns
///
/// A vector of file/folder paths on success, or an error if both methods fail
///
/// # Errors
///
/// Returns an error only if both native COM and PowerShell fallback fail.
/// Common failure scenarios:
/// - PowerShell is not available
/// - Quick Access is disabled or corrupted
/// - Insufficient permissions
///
/// # Example
///
/// ```rust,no_run
/// use wincent::{QuickAccess, query::query_recent};
///
/// // Query recent files with automatic fallback
/// let files = query_recent(QuickAccess::RecentFiles)?;
/// # Ok::<(), wincent::error::WincentError>(())
/// ```
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
/// This function retrieves files that Windows has tracked as recently accessed.
/// The list is automatically maintained by Windows Explorer and typically contains
/// up to 20 items (Windows default limit).
///
/// # Returns
///
/// Returns a vector of file paths as strings. Paths are in Windows format (e.g., `C:\Users\...`).
///
/// # Errors
///
/// Returns an error if:
/// - Quick Access is disabled or inaccessible
/// - Both native COM and PowerShell methods fail
/// - Insufficient permissions to access Quick Access data
///
/// # Example
///
/// ```rust,no_run
/// use wincent::query::get_recent_files;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     let recent_files = get_recent_files()?;
///
///     println!("Found {} recent files:", recent_files.len());
///     for file in recent_files {
///         println!("  - {}", file);
///     }
///     Ok(())
/// }
/// ```
pub fn get_recent_files() -> WincentResult<Vec<String>> {
    query_recent(QuickAccess::RecentFiles)
}

/// Gets a list of frequent folders from Windows Quick Access.
///
/// This function retrieves folders that Windows has tracked as frequently accessed.
/// Windows automatically maintains this list based on user activity, typically containing
/// up to 20 items (Windows default limit).
///
/// # Returns
///
/// Returns a vector of folder paths as strings. Paths are in Windows format (e.g., `C:\Users\Documents`).
///
/// # Errors
///
/// Returns an error if:
/// - Quick Access is disabled or inaccessible
/// - Both native COM and PowerShell methods fail
/// - Insufficient permissions to access Quick Access data
///
/// # Example
///
/// ```rust,no_run
/// use wincent::query::get_frequent_folders;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     let folders = get_frequent_folders()?;
///
///     println!("Found {} frequent folders:", folders.len());
///     for folder in folders {
///         println!("  - {}", folder);
///     }
///     Ok(())
/// }
/// ```
pub fn get_frequent_folders() -> WincentResult<Vec<String>> {
    query_recent(QuickAccess::FrequentFolders)
}

/// Gets a list of all items from Windows Quick Access (Recent Items namespace).
///
/// This function queries the Recent Items namespace, which contains both recently accessed
/// files and folders. Windows maintains this as a single unified list, not as separate
/// Recent Files and Frequent Folders lists.
///
/// # Returns
///
/// Returns a vector of strings containing the paths of all items in the Recent Items namespace.
/// The list includes both files and folders, typically up to 20 items (Windows default limit
/// for the Recent Items namespace).
///
/// # Errors
///
/// Returns an error if:
/// - Quick Access is disabled or inaccessible
/// - Both native COM and PowerShell methods fail
/// - Insufficient permissions to access Quick Access data
///
/// # Example
///
/// ```rust,no_run
/// use wincent::query::get_quick_access_items;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     match get_quick_access_items() {
///         Ok(items) => {
///             println!("Found {} Quick Access items:", items.len());
///             for item in items {
///                 println!("  - {}", item);
///             }
///         },
///         Err(e) => eprintln!("Failed to get Quick Access items: {}", e)
///     }
///     Ok(())
/// }
/// ```
///
/// # Note
///
/// This function queries the "Recent Items" namespace (`shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}`),
/// which is different from the "Frequent Folders" namespace. To query frequent folders specifically,
/// use [`get_frequent_folders()`] instead.
pub fn get_quick_access_items() -> WincentResult<Vec<String>> {
    query_recent(QuickAccess::All)
}

/****************************************************** Check Quick Access ******************************************************/

/// Checks if an exact file path exists in the Windows Recent Files list.
///
/// This function performs **exact path comparison** (case-sensitive on the full path).
/// For partial/fuzzy matching, use [`is_in_recent_files()`] instead.
///
/// # Arguments
///
/// * `path` - The exact file path to search for (e.g., `"C:\\Users\\Documents\\file.txt"`)
///
/// # Returns
///
/// Returns `true` if the exact path is found in the recent files list, `false` otherwise.
///
/// # Errors
///
/// Returns an error if Quick Access query fails.
///
/// # Example
///
/// ```rust,no_run
/// use wincent::query::is_recent_file_exact;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     // Exact match only - must match the full path exactly
///     let exists = is_recent_file_exact("C:\\Users\\Documents\\report.docx")?;
///     if exists {
///         println!("Exact file path found in recent files");
///     } else {
///         println!("File not found (exact match required)");
///     }
///     Ok(())
/// }
/// ```
///
/// # See Also
///
/// - [`is_in_recent_files()`] - For fuzzy/substring matching
pub fn is_recent_file_exact(path: &str) -> WincentResult<bool> {
    let items = get_recent_files()?;
    Ok(items.iter().any(|item| item == path))
}

/// Checks if a file path or keyword exists in the Windows Recent Files list.
///
/// **Note**: This function performs **substring matching** (fuzzy match). If you need
/// exact path matching, use [`is_recent_file_exact()`] instead.
///
/// # Arguments
///
/// * `keyword` - The file path or partial path to search for (substring match)
///
/// # Returns
///
/// Returns `true` if any recent file path contains the keyword, `false` otherwise.
///
/// # Errors
///
/// Returns an error if Quick Access query fails.
///
/// # Example
///
/// ```rust,no_run
/// use wincent::query::is_in_recent_files;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     // Fuzzy match - matches any path containing "Documents"
///     let file_exists = is_in_recent_files("Documents")?;
///
///     // This will match paths like:
///     // - "C:\\Users\\Alice\\Documents\\report.docx"
///     // - "D:\\My Documents\\file.txt"
///     // - "C:\\Documents and Settings\\..."
///
///     if file_exists {
///         println!("Found file(s) matching 'Documents'");
///     }
///
///     // You can also search by filename
///     if is_in_recent_files("report.docx")? {
///         println!("Found report.docx in recent files");
///     }
///
///     Ok(())
/// }
/// ```
///
/// # See Also
///
/// - [`is_recent_file_exact()`] - For exact path matching
pub fn is_in_recent_files(keyword: &str) -> WincentResult<bool> {
    let items = get_recent_files()?;

    Ok(items.iter().any(|item| item.contains(keyword)))
}

/// Checks if an exact folder path exists in the Windows Frequent Folders list.
///
/// This function performs **exact path comparison** (case-sensitive on the full path).
/// For partial/fuzzy matching, use [`is_in_frequent_folders()`] instead.
///
/// # Arguments
///
/// * `path` - The exact folder path to search for (e.g., `"C:\\Users\\Documents"`)
///
/// # Returns
///
/// Returns `true` if the exact path is found in the frequent folders list, `false` otherwise.
///
/// # Errors
///
/// Returns an error if Quick Access query fails.
///
/// # Example
///
/// ```rust,no_run
/// use wincent::query::is_frequent_folder_exact;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     // Exact match - must match the full path exactly
///     let folder_exists = is_frequent_folder_exact("C:\\Users\\Documents")?;
///     if folder_exists {
///         println!("Exact folder path found in frequent folders list");
///     } else {
///         println!("Folder not found (exact match required)");
///     }
///     Ok(())
/// }
/// ```
///
/// # See Also
///
/// - [`is_in_frequent_folders()`] - For fuzzy/substring matching
pub fn is_frequent_folder_exact(path: &str) -> WincentResult<bool> {
    let items = get_frequent_folders()?;
    Ok(items.iter().any(|item| item == path))
}

/// Checks if a folder path or keyword exists in the Windows Frequent Folders list.
///
/// **Note**: This function performs **substring matching** (fuzzy match). If you need
/// exact path matching, use [`is_frequent_folder_exact()`] instead.
///
/// # Arguments
///
/// * `keyword` - The folder path or partial path to search for (substring match)
///
/// # Returns
///
/// Returns `true` if any frequent folder path contains the keyword, `false` otherwise.
///
/// # Errors
///
/// Returns an error if Quick Access query fails.
///
/// # Example
///
/// ```rust,no_run
/// use wincent::query::is_in_frequent_folders;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     // Fuzzy match - matches any path containing "Projects"
///     let folder_exists = is_in_frequent_folders("Projects")?;
///     if folder_exists {
///         println!("Found folder(s) matching 'Projects'");
///     } else {
///         println!("No folders matching 'Projects' found");
///     }
///
///     // Search by drive letter
///     if is_in_frequent_folders("D:\\")? {
///         println!("Found folders on D: drive");
///     }
///
///     Ok(())
/// }
/// ```
///
/// # See Also
///
/// - [`is_frequent_folder_exact()`] - For exact path matching
pub fn is_in_frequent_folders(keyword: &str) -> WincentResult<bool> {
    let items = get_frequent_folders()?;

    Ok(items.iter().any(|item| item.contains(keyword)))
}

/// Checks if an exact path exists in the Windows Quick Access list (Recent Items namespace).
///
/// This function performs **exact path comparison** (case-sensitive on the full path).
/// For partial/fuzzy matching, use [`is_in_quick_access()`] instead.
///
/// # Arguments
///
/// * `path` - The exact path to search for (file or folder)
///
/// # Returns
///
/// Returns `true` if the exact path is found in the Recent Items namespace, `false` otherwise.
///
/// # Errors
///
/// Returns an error if Quick Access query fails.
///
/// # Example
///
/// ```rust,no_run
/// use wincent::query::is_in_quick_access_exact;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     // Check for exact file path
///     let exists = is_in_quick_access_exact("C:\\Users\\Documents\\file.txt")?;
///     if exists {
///         println!("Exact path found in Quick Access");
///     }
///
///     // Check for exact folder path
///     let exists = is_in_quick_access_exact("C:\\Users\\Documents")?;
///     if exists {
///         println!("Exact folder path found in Quick Access");
///     }
///
///     Ok(())
/// }
/// ```
///
/// # Note
///
/// This function queries the "Recent Items" namespace, not the "Frequent Folders" namespace.
///
/// # See Also
///
/// - [`is_in_quick_access()`] - For fuzzy/substring matching
pub fn is_in_quick_access_exact(path: &str) -> WincentResult<bool> {
    let items = get_quick_access_items()?;
    Ok(items.iter().any(|item| item == path))
}

/// Checks if a path or keyword exists in the Windows Quick Access list (Recent Items namespace).
///
/// **Note**: This function performs **substring matching** (fuzzy match). If you need
/// exact path matching, use [`is_in_quick_access_exact()`] instead.
///
/// # Arguments
///
/// * `keyword` - The path or partial path to search for (substring match)
///
/// # Returns
///
/// Returns `true` if any path in the Recent Items namespace contains the keyword,
/// `false` otherwise.
///
/// # Errors
///
/// Returns an error if Quick Access query fails.
///
/// # Example
///
/// ```rust,no_run
/// use wincent::query::is_in_quick_access;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     // Fuzzy match - check for items containing "Documents"
///     if is_in_quick_access("Documents")? {
///         println!("Found item(s) matching 'Documents'");
///     }
///
///     // Check for items in a specific location
///     if is_in_quick_access("C:\\Projects\\")? {
///         println!("Found items from Projects folder");
///     }
///
///     // Search by filename
///     if is_in_quick_access("report.docx")? {
///         println!("Found report.docx in Quick Access");
///     }
///
///     Ok(())
/// }
/// ```
///
/// # Note
///
/// This function queries the "Recent Items" namespace, not the "Frequent Folders" namespace.
///
/// # See Also
///
/// - [`is_in_quick_access_exact()`] - For exact path matching
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
    fn test_com_s_false_reference_counting() -> WincentResult<()> {
        // Tests that ComGuard::initialize() correctly handles S_FALSE
        // and properly calls CoUninitialize to balance reference counts.
        //
        // Strategy: Use multiple init/uninit cycles to amplify reference leaks.
        // If the library leaks references, COM will remain initialized after
        // the final CoUninitialize(), which we can detect.
        use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED};

        unsafe {
            // Cycle 1: User init -> library call -> user uninit
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            assert_eq!(hr.0, 0, "First CoInitializeEx should return S_OK");

            {
                let _guard = ComGuard::try_initialize();
                assert!(_guard.is_ok(), "Should handle S_FALSE correctly: {:?}", _guard);
            } // Guard drops here, should call CoUninitialize

            CoUninitialize();

            // Cycle 2: Repeat to amplify potential leaks
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            assert_eq!(hr.0, 0, "Second cycle should return S_OK (COM was uninitialized)");

            {
                let _guard = ComGuard::try_initialize();
                assert!(_guard.is_ok(), "Should handle S_FALSE on second cycle: {:?}", _guard);
            }

            CoUninitialize();

            // Cycle 3: One more time
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            assert_eq!(hr.0, 0, "Third cycle should return S_OK");

            {
                let _guard = ComGuard::try_initialize();
                assert!(_guard.is_ok(), "Should handle S_FALSE on third cycle: {:?}", _guard);
            }

            CoUninitialize();

            // Verification: Try to initialize with incompatible mode
            // If there are leaked references, this will fail with RPC_E_CHANGED_MODE
            let hr = CoInitializeEx(None, COINIT_MULTITHREADED);
            assert!(
                hr.0 == 0 || hr.0 == 1,
                "Should be able to initialize with MTA (no leaked STA references). Got: 0x{:08X}",
                hr.0
            );
            CoUninitialize();
        }

        Ok(())
    }

    #[test]
    fn test_com_apartment_mismatch() -> WincentResult<()> {
        // Tests that ComGuard::initialize() correctly detects and reports
        // RPC_E_CHANGED_MODE when called from an MTA thread.
        use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED};

        unsafe {
            // Initialize as MTA (multi-threaded apartment)
            let hr = CoInitializeEx(None, COINIT_MULTITHREADED);
            assert!(hr.is_ok() || hr.0 == 1, "MTA init should succeed");

            // Call STA-only function (should return ComApartmentMismatch)
            let result = ComGuard::try_initialize().map_err(|status| match status {
                ComInitStatus::ApartmentMismatch => {
                    WincentError::ComApartmentMismatch(
                        "Thread already initialized with incompatible COM apartment model".to_string()
                    )
                }
                ComInitStatus::OtherError(hr) => {
                    WincentError::SystemError(format!("Failed to initialize COM: 0x{:08X}", hr))
                }
                _ => unreachable!(),
            });

            match result {
                Err(WincentError::ComApartmentMismatch(msg)) => {
                    assert!(
                        msg.contains("incompatible"),
                        "Error message should mention incompatibility: {}",
                        msg
                    );
                }
                _ => panic!("Expected ComApartmentMismatch error, got: {:?}", result),
            }

            // Clean up MTA
            CoUninitialize();
        }

        Ok(())
    }

    #[test]
    fn test_apartment_mismatch_triggers_fallback() -> WincentResult<()> {
        // Tests that query_recent() attempts fallback when native COM fails
        // due to apartment mismatch (RPC_E_CHANGED_MODE).
        //
        // This test verifies the fallback MECHANISM, not whether PowerShell succeeds.
        // We verify that:
        // 1. Native COM fails with apartment mismatch
        // 2. The error is NOT propagated directly (fallback was attempted)
        use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED};

        unsafe {
            // Initialize as MTA
            let hr = CoInitializeEx(None, COINIT_MULTITHREADED);
            assert!(hr.is_ok() || hr.0 == 1, "MTA init should succeed");

            // First verify that native COM fails with apartment mismatch
            let native_result = query_recent_native(QuickAccess::RecentFiles);
            assert!(
                matches!(native_result, Err(WincentError::ComApartmentMismatch(_))),
                "Native should fail with ComApartmentMismatch, got: {:?}",
                native_result
            );

            // Now test query_recent() which has fallback
            let result = query_recent(QuickAccess::RecentFiles);

            // The result should NOT be ComApartmentMismatch (fallback was attempted)
            // It may succeed (PowerShell works) or fail with other errors (PowerShell unavailable)
            match result {
                Err(WincentError::ComApartmentMismatch(_)) => {
                    panic!("ComApartmentMismatch should not propagate - fallback should be attempted");
                }
                Ok(_) => {
                    // Fallback succeeded - good!
                }
                Err(_) => {
                    // Fallback was attempted but failed (e.g., PowerShell unavailable) - acceptable
                    // The important thing is that ComApartmentMismatch was not propagated
                }
            }

            CoUninitialize();
        }

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
            let _ = query_recent_powershell(QuickAccess::All)?;
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
