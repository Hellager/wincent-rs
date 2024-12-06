use powershell_script::PsError;

pub enum QuickAccess {
    FrequentFolders,
    RecentFiles,
    All
}

#[derive(Debug)]
pub enum WincentError {
    ScriptError(PsError),
    IoError(std::io::Error),
    ConvertError(std::array::TryFromSliceError),
    ExecuteError(tokio::task::JoinError),
    TimeoutError(tokio::time::error::Elapsed)
}

const SCRIPT_TIMEOUT: u64 = 3;

/******************************** Utils ********************************/

fn is_admin() -> bool {
    use windows::Win32::Foundation::BOOL;
    use windows::Win32::UI::Shell::IsUserAnAdmin;

    unsafe {
        IsUserAnAdmin() == BOOL(1)
    } 
}

/// Checks if the PowerShell execution policy for the current user is set to a 
/// feasible value that allows script execution.
///
/// This function runs a PowerShell script to retrieve the current execution policy 
/// for the user. It considers the following [policies](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-7.4) as feasible for script execution:
/// - `AllSigned`
/// - `Bypass`
/// - `RemoteSigned`
/// - `Unrestricted`
///
/// The function runs asynchronously and returns a `Result<bool, WincentError>`. 
/// The boolean indicates whether the execution policy is feasible (`true`) or not 
/// (`false`). In case of an error during script execution or if the operation 
/// times out, it returns a `WincentError`.
///
/// # Errors
///
/// This function can return the following errors:
/// - `WincentError::ScriptError` if there is an error executing the PowerShell script.
/// - `WincentError::TimeoutError` if the operation exceeds the specified timeout.
/// - `WincentError::ExecuteError` if there is an error in executing the script handle.
///
/// # Example
///
/// ```
/// let is_feasible = check_feasible().await?;
/// if is_feasible {
///     println!("The execution policy is feasible for script execution.");
/// } else {
///     println!("The execution policy is not feasible.");
/// }
/// ```
pub async fn check_feasible() -> Result<bool, WincentError> {
    use powershell_script::PsScriptBuilder;

    const SCRIPT: &str = r#"
        $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
        Get-ExecutionPolicy -Scope CurrentUser
    "#;

    let ps = PsScriptBuilder::new()
        .no_profile(true)
        .non_interactive(true)
        .hidden(false)
        .print_commands(false)
        .build();

    let handle = tokio::task::spawn_blocking(move || {
            ps.run(SCRIPT).map(|output| {
                output.stdout().map_or(false, |policy| {
                    matches!(policy.as_str(), "AllSigned" | "Bypass" | "RemoteSigned" | "Unrestricted")
                })
            }).map_err(WincentError::ScriptError)
        });

    tokio::time::timeout(tokio::time::Duration::from_secs(SCRIPT_TIMEOUT), handle)
        .await
        .map_err(WincentError::TimeoutError)?
        .map_err(WincentError::ExecuteError)?
}

/// Sets the PowerShell execution policy for the current user to `RemoteSigned`.
///
/// This function runs a PowerShell script that modifies the execution policy 
/// to `RemoteSigned`, which allows the execution of scripts that are signed 
/// by a trusted publisher and scripts that are created locally without a signature.
///
/// The function runs asynchronously and returns a `Result<(), WincentError>`. 
/// On success, it returns `Ok(())`. In case of an error during script execution 
/// or if the operation times out, it returns a `WincentError`.
///
/// # Errors
///
/// This function can return the following errors:
/// - `WincentError::ScriptError` if there is an error executing the PowerShell script.
/// - `WincentError::TimeoutError` if the operation exceeds the specified timeout.
/// - `WincentError::ExecuteError` if there is an error in executing the script handle.
///
/// # Example
///
/// ```
/// fix_feasible().await?;
/// println!("Execution policy has been set to RemoteSigned.");
/// ```
pub async fn fix_feasible() -> Result<(), WincentError> {
    use powershell_script::PsScriptBuilder;

    const SCRIPT: &str = r#"
        $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
    "#;

    let ps = PsScriptBuilder::new()
        .no_profile(true)
        .non_interactive(true)
        .hidden(false)
        .print_commands(false)
        .build();

    let handle = tokio::task::spawn_blocking(move || {
        ps.run(SCRIPT)
            .map(|_| ())
            .map_err(WincentError::ScriptError)
    });

    tokio::time::timeout(tokio::time::Duration::from_secs(SCRIPT_TIMEOUT), handle)
        .await
        .map_err(WincentError::TimeoutError)?
        .map_err(WincentError::ExecuteError)?
}


/// Refreshes all open Windows Explorer windows asynchronously.
///
/// This function constructs and executes a PowerShell script that refreshes all
/// currently open Windows Explorer windows. It uses the `powershell_script` crate
/// to build and run the PowerShell script in a non-interactive manner.
///
/// # Errors
///
/// This function returns a `Result<(), WincentError>`, which can be:
/// - `Ok(())` if the operation was successful.
/// - `Err(WincentError::ScriptError)` if there was an error executing the PowerShell script.
/// - `Err(WincentError::TimeoutError)` if the operation timed out.
/// - `Err(WincentError::ExecuteError)` if there was an error during execution.
///
/// # Example
///
/// ```rust
/// match refresh_explorer_window().await {
///     Ok(()) => println!("Explorer windows refreshed successfully."),
///     Err(e) => eprintln!("Failed to refresh explorer windows: {:?}", e),
/// }
/// ```
///
/// # Notes
///
/// The PowerShell script executed by this function sets the output encoding to UTF-8,
/// creates a Shell.Application COM object, retrieves all open windows, and refreshes each one.
/// The script is run in a blocking manner using `tokio::task::spawn_blocking`, and a timeout
/// is applied to ensure that the operation does not hang indefinitely.
pub async fn refresh_explorer_window() -> Result<(), WincentError> {
    use powershell_script::PsScriptBuilder;
    use std::io::{Error, ErrorKind};

    if !check_feasible().await? {
        return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
    }

    const SCRIPT: &str = r#"
        $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
        $shellApplication = New-Object -ComObject Shell.Application;
        $windows = $shellApplication.Windows();
        $windows | ForEach-Object { $_.Refresh() }
    "#;

    let ps = PsScriptBuilder::new()
        .no_profile(true)
        .non_interactive(true)
        .hidden(false)
        .print_commands(false)
        .build();

    let handle = tokio::task::spawn_blocking(move || {
        ps.run(SCRIPT)
            .map(|_| ())
            .map_err(WincentError::ScriptError)
    });

    tokio::time::timeout(tokio::time::Duration::from_secs(SCRIPT_TIMEOUT), handle)
        .await
        .map_err(WincentError::TimeoutError)?
        .map_err(WincentError::ExecuteError)?
}

/************************* Query Quick Access *************************/

/// Queries recent items from the Quick Access section of Windows Explorer.
///
/// This asynchronous function retrieves a list of recent items based on the specified
/// `recent_type`. It uses a PowerShell script to access the Quick Access feature of
/// Windows Explorer and returns the paths of the recent items as a vector of strings.
///
/// # Parameters
///
/// - `recent_type`: An enum of type `QuickAccess` that specifies the type of recent items
///   to query. It can be one of the following:
///   - `QuickAccess::FrequentFolders`: Retrieves frequently accessed folders.
///   - `QuickAccess::RecentFiles`: Retrieves recently accessed files (excluding folders).
///   - `QuickAccess::All`: Retrieves all recent items (both files and folders).
///
/// # Returns
///
/// This function returns a `Result<Vec<String>, WincentError>`, which can be:
/// - `Ok(Vec<String>)`: A vector containing the paths of the recent items if the operation
///   is successful.
/// - `Err(WincentError)`: An error of type `WincentError` if the operation fails. Possible
///   errors include:
///   - `WincentError::ScriptError`: If there is an error executing the PowerShell script.
///   - `WincentError::ExecuteError`: If there is an error during the execution of the task.
///   - `WincentError::TimeoutError`: If the operation times out.
///
/// # Example
///
/// ```rust
/// match query_recent(QuickAccess::RecentFiles).await {
///     Ok(paths) => println!("Recent files: {:?}", paths),
///     Err(e) => eprintln!("Error querying recent files: {:?}", e),
/// }
/// ```
///
/// # Notes
///
/// The function constructs a PowerShell script that sets the output encoding to UTF-8,
/// creates a Shell.Application COM object, and retrieves the items from the specified
/// Quick Access namespace. The results are processed to filter out empty lines and
/// return only valid paths.
async fn query_recent(recent_type: QuickAccess) -> Result<Vec<String>, WincentError> {
    use powershell_script::PsScriptBuilder;
    use std::io::{Error, ErrorKind};

    if !check_feasible().await? {
        return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
    }

    let (shell_namespace, condition) = match recent_type {
        QuickAccess::FrequentFolders => ("3936E9E4-D92C-4EEE-A85A-BC16D5EA0819", ""),
        QuickAccess::RecentFiles => ("679f85cb-0220-4080-b29b-5540cc05aab6", "| where {$_.IsFolder -eq $false}"),
        QuickAccess::All => ("679f85cb-0220-4080-b29b-5540cc05aab6", ""),
    };

    let script = format!(r#"
        $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
        $shell = New-Object -ComObject Shell.Application;
        $shell.Namespace('shell:::{{{}}}').Items() {} | ForEach-Object {{ $_.Path }};
    "#, shell_namespace, condition);

    let ps = PsScriptBuilder::new()
        .no_profile(true)
        .non_interactive(true)
        .hidden(false)
        .print_commands(false)
        .build();

    let handle = tokio::task::spawn_blocking(move || {
        ps.run(&script).map(|output| {
            output.stdout()
                .map_or_else(Vec::new, |data| {
                    data.lines()
                        .filter_map(|item| {
                            let trimmed = item.trim();
                            if !trimmed.is_empty() {
                                Some(trimmed.to_string())
                            } else {
                                None
                            }
                        })
                        .collect()
                })
        }).map_err(WincentError::ScriptError)
    });

    match tokio::time::timeout(tokio::time::Duration::from_secs(SCRIPT_TIMEOUT), handle).await {
        Ok(res) => {
            res.map_err(WincentError::ExecuteError)?
        },
        Err(e) => Err(WincentError::TimeoutError(e)),
    }
}

/// Retrieves a list of recently accessed files from the Quick Access section of Windows Explorer.
///
/// This asynchronous function calls `query_recent` with the `QuickAccess::RecentFiles`
/// variant to obtain a list of recently accessed files.
///
/// # Returns
///
/// This function returns a `Result<Vec<String>, WincentError>`, which can be:
/// - `Ok(Vec<String>)`: A vector containing the paths of the recently accessed files if the
///   operation is successful.
/// - `Err(WincentError)`: An error of type `WincentError` if the operation fails.
///
/// # Example
///
/// ```rust
/// match get_recent_files().await {
///     Ok(files) => println!("Recent files: {:?}", files),
///     Err(e) => eprintln!("Error retrieving recent files: {:?}", e),
/// }
/// ```
pub async fn get_recent_files() -> Result<Vec<String>, WincentError> {
    query_recent(QuickAccess::RecentFiles).await
}

/// Retrieves a list of frequently accessed folders from the Quick Access section of Windows Explorer.
///
/// This asynchronous function calls `query_recent` with the `QuickAccess::FrequentFolders`
/// variant to obtain a list of frequently accessed folders.
///
/// # Returns
///
/// This function returns a `Result<Vec<String>, WincentError>`, which can be:
/// - `Ok(Vec<String>)`: A vector containing the paths of the frequently accessed folders if the
///   operation is successful.
/// - `Err(WincentError)`: An error of type `WincentError` if the operation fails.
///
/// # Example
///
/// ```rust
/// match get_frequent_folders().await {
///     Ok(folders) => println!("Frequent folders: {:?}", folders),
///     Err(e) => eprintln!("Error retrieving frequent folders: {:?}", e),
/// }
/// ```
pub async fn get_frequent_folders() -> Result<Vec<String>, WincentError> {
    query_recent(QuickAccess::FrequentFolders).await
}

/// Retrieves a list of all items in the Quick Access section of Windows Explorer.
///
/// This asynchronous function calls `query_recent` with the `QuickAccess::All`
/// variant to obtain a list of all items (both files and folders) in Quick Access.
///
/// # Returns
///
/// This function returns a `Result<Vec<String>, WincentError>`, which can be:
/// - `Ok(Vec<String>)`: A vector containing the paths of all items in Quick Access if the
///   operation is successful.
/// - `Err(WincentError)`: An error of type `WincentError` if the operation fails.
///
/// # Example
///
/// ```rust
/// match get_quick_access_items().await {
///     Ok(items) => println!("Quick Access items: {:?}", items),
///     Err(e) => eprintln!("Error retrieving Quick Access items: {:?}", e),
/// }
/// ```
pub async fn get_quick_access_items() -> Result<Vec<String>, WincentError> {
    query_recent(QuickAccess::All).await
}

/************************* Check Existence *************************/

/// Checks if any of the specified keywords are present in the Quick Access items.
///
/// This asynchronous function queries the Quick Access section of Windows Explorer
/// for items based on the specified `specific_type` and checks if any of the items
/// contain any of the provided keywords.
///
/// # Parameters
///
/// - `keywords`: A vector of string slices (`Vec<&str>`) representing the keywords to search for
///   in the Quick Access items.
/// - `specific_type`: An optional parameter of type `Option<QuickAccess>`, which specifies the
///   type of Quick Access items to search through. It can be one of the following:
///   - `Some(QuickAccess::FrequentFolders)`: Search in frequently accessed folders.
///   - `Some(QuickAccess::RecentFiles)`: Search in recently accessed files.
///   - `Some(QuickAccess::All)`: Search in all Quick Access items (both files and folders).
///   - `None`: Search in all Quick Access items by default.
///
/// # Returns
///
/// This function returns a `Result<bool, WincentError>`, which can be:
/// - `Ok(true)`: If at least one of the keywords is found in the Quick Access items.
/// - `Ok(false)`: If none of the keywords are found.
/// - `Err(WincentError)`: An error of type `WincentError` if the operation fails, such as issues
///   retrieving the Quick Access items.
///
/// # Example
///
/// ```rust
/// match is_in_quick_access(vec!["example", "test"], Some(QuickAccess::RecentFiles)).await {
///     Ok(found) => {
///         if found {
///             println!("At least one keyword was found in recent files.");
///         } else {
///             println!("No keywords found in recent files.");
///         }
///     },
///     Err(e) => eprintln!("Error checking Quick Access items: {:?}", e),
/// }
/// ```
///
/// # Notes
///
/// The function retrieves the relevant Quick Access items based on the specified type,
/// then checks each item to see if it contains any of the provided keywords. The search
/// is case-sensitive and checks for substring matches.
pub async fn is_in_quick_access(keywords: Vec<&str>, specific_type: Option<QuickAccess>) -> Result<bool, WincentError> {
    let target_items = match specific_type {
        Some(QuickAccess::FrequentFolders) => get_frequent_folders().await?,
        Some(QuickAccess::RecentFiles) => get_recent_files().await?,
        Some(QuickAccess::All) => get_quick_access_items().await?,
        None => get_quick_access_items().await?,
    };

    let is_found = target_items.iter().any(|item| {
        keywords.iter().any(|&keyword| item.contains(keyword))
    });

    Ok(is_found)
}

/************************* Remove Recent File *************************/

/// Handles recent files by removing a specified file from the recent files list.
///
/// This asynchronous function removes a file from the recent files in Windows Explorer
/// if `is_remove` is set to true. If `is_remove` is false, it returns an error indicating
/// that the operation is unsupported.
///
/// # Parameters
///
/// - `path`: A string slice representing the path of the file to be removed from recent files.
/// - `is_remove`: A boolean indicating whether to remove the specified file. If true, the file
///   will be removed; if false, an error will be returned.
///
/// # Returns
///
/// This function returns a `Result<(), WincentError>`, which can be:
/// - `Ok(())`: If the operation was successful.
/// - `Err(WincentError)`: An error of type `WincentError` if the operation fails, such as issues
///   executing the PowerShell script or if the operation is unsupported.
///
/// # Example
///
/// ```rust
/// match handle_recent_files("C:\\path\\to\\file.txt", true).await {
///     Ok(()) => println!("File removed from recent files."),
///     Err(e) => eprintln!("Error handling recent files: {:?}", e),
/// }
/// ```
async fn handle_recent_files(path: &str, is_remove: bool) -> Result<(), WincentError> {
    use powershell_script::PsScriptBuilder;
    use std::io::{Error, ErrorKind};

    if !check_feasible().await? {
        return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
    }

    if !is_remove {
        return Err(WincentError::IoError(std::io::ErrorKind::Unsupported.into()));
    }

    let script = format!(r#"
        $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
        $shell = New-Object -ComObject Shell.Application;
        $files = $shell.Namespace("shell:::{{679f85cb-0220-4080-b29b-5540cc05aab6}}").Items() | where {{$_.IsFolder -eq $false}};
        $target = $files | where {{$_.Path -match ${}}};
        $target.InvokeVerb("remove");
    "#, path);

    let ps = PsScriptBuilder::new()
        .no_profile(true)
        .non_interactive(true)
        .hidden(false)
        .print_commands(false)
        .build();

    let handle = tokio::task::spawn_blocking(move || {
        ps.run(&script).map_err(WincentError::ScriptError)
    });

    match tokio::time::timeout(tokio::time::Duration::from_secs(SCRIPT_TIMEOUT), handle).await {
        Ok(res) => {
            res.map(|_| ())
            .map_err(WincentError::ExecuteError)
        },
        Err(e) => Err(WincentError::TimeoutError(e)),
    }
}

/// Removes a specified file from the recent files list in Windows Explorer.
///
/// This asynchronous function checks if the specified file exists and is a valid file.
/// If it is, it calls `handle_recent_files` to remove the file from the recent files list.
/// If the file does not exist or is not a valid file, it returns an appropriate error.
///
/// # Parameters
///
/// - `path`: A string slice representing the path of the file to be removed from recent files.
///
/// # Returns
///
/// This function returns a `Result<(), WincentError>`, which can be:
/// - `Ok(())`: If the operation was successful.
/// - `Err(WincentError)`: An error of type `WincentError` if the operation fails, such as if the
///   file does not exist, is not a valid file, or if there are issues handling recent files.
///
/// # Example
///
/// ```rust
/// match remove_from_recent_files("C:\\path\\to\\file.txt").await {
///     Ok(()) => println!("File removed from recent files."),
///     Err(e) => eprintln!("Error removing file from recent files: {:?}", e),
/// }
/// ```
pub async fn remove_from_recent_files(path: &str) -> Result<(), WincentError> {
    use std::fs;
    use std::path::Path;

    if fs::metadata(path).is_err() {
        return Err(WincentError::IoError(std::io::ErrorKind::NotFound.into()));
    }

    if !Path::new(path).is_file() {
        return Err(WincentError::IoError(std::io::ErrorKind::InvalidData.into()));
    }

    let in_quick_access = match is_in_quick_access(vec![path], Some(QuickAccess::RecentFiles)).await {
        Ok(result) => result,
        Err(e) => return Err(e),
    };

    if !in_quick_access {
        return Ok(());
    }

    handle_recent_files(path, true).await?;
    Ok(())
}

/************************* Remove/Add Frequent Folders *************************/

/// Handles the pinning of a specified folder to the Windows home directory.
///
/// This asynchronous function takes a path to a folder as input and uses a PowerShell script
/// to pin that folder to the user's home directory in Windows. The function constructs a PowerShell
/// script that utilizes the `Shell.Application` COM object to perform the pinning operation.
///
/// # Arguments
///
/// * `path` - A string slice that holds the path to the folder that needs to be pinned.
///
/// # Returns
///
/// This function returns a `Result<(), WincentError>`. On success, it returns `Ok(())`. 
/// If an error occurs, it returns a `WincentError` which can indicate various issues:
/// - `ScriptError` if there was an error running the PowerShell script.
/// - `ExecuteError` if the script execution failed.
/// - `TimeoutError` if the operation exceeds the predefined timeout duration.
///
/// # Errors
///
/// The function may fail due to:
/// - Issues with the PowerShell script execution.
/// - The specified path being invalid or inaccessible.
/// - The operation timing out if it takes longer than the defined `SCRIPT_TIMEOUT`.
///
/// # Example
///
/// ```
/// let result = handle_frequent_folders("C:\\path\\to\\folder").await;
/// match result {
///     Ok(_) => println!("Folder pinned successfully!"),
///     Err(e) => eprintln!("Error pinning folder: {:?}", e),
/// }
/// ```
async fn handle_frequent_folders(path: &str) -> Result<(), WincentError> {
    use powershell_script::PsScriptBuilder;
    use std::io::{Error, ErrorKind};

    if !check_feasible().await? {
        return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
    }

    let script: String = format!(r#"
            $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
            $shell = New-Object -ComObject Shell.Application;
            $shell.Namespace("{}").Self.InvokeVerb("pintohome");
        "#, path);

    let ps = PsScriptBuilder::new()
        .no_profile(true)
        .non_interactive(true)
        .hidden(true)
        .print_commands(true)
        .build();

    let handle = tokio::task::spawn_blocking(move || {
        ps.run(&script).map_err(WincentError::ScriptError)
    });

    match tokio::time::timeout(tokio::time::Duration::from_secs(SCRIPT_TIMEOUT), handle).await {
        Ok(res) => {
            res.map(|_| ())
            .map_err(WincentError::ExecuteError)
        },
        Err(e) => Err(WincentError::TimeoutError(e)),
    }
}

/// Adds a specified folder to the list of frequent folders.
///
/// This asynchronous function checks if the provided path is valid and represents a directory.
/// If the path is valid, it calls the `handle_frequent_folders` function to pin the folder
/// to the user's home directory in Windows.
///
/// # Arguments
///
/// * `path` - A string slice that holds the path to the folder that needs to be added to the
///   frequent folders list.
///
/// # Returns
///
/// This function returns a `Result<(), WincentError>`. On success, it returns `Ok(())`. 
/// If an error occurs, it returns a `WincentError` which can indicate various issues:
/// - `IoError` if there is an issue accessing the file system, such as the path not existing
///   or being invalid.
/// - `InvalidData` if the specified path does not point to a directory.
///
/// # Errors
///
/// The function may fail due to:
/// - The specified path not being accessible or not existing, resulting in an I/O error.
/// - The specified path not being a directory, leading to an `InvalidData` error.
///
/// # Example
///
/// ```
/// let result = add_to_frequent_folders("C:\\path\\to\\folder").await;
/// match result {
///     Ok(_) => println!("Folder added to frequent folders successfully!"),
///     Err(e) => eprintln!("Error adding folder: {:?}", e),
/// }
/// ```
pub async fn add_to_frequent_folders(path: &str) -> Result<(), WincentError> {
    if let Err(e) = std::fs::metadata(path) {
        return Err(WincentError::IoError(e));
    }

    if !std::path::Path::new(path).is_dir() {
        return Err(WincentError::IoError(std::io::ErrorKind::InvalidData.into()));
    }

    handle_frequent_folders(path).await?;

    Ok(())
}

/// Removes a specified folder from the list of frequent folders.
///
/// This asynchronous function checks if the provided path is valid and represents a directory.
/// If the path is valid and the folder is currently in the frequent folders list, it calls
/// the `handle_frequent_folders` function to remove the folder from the user's home directory
/// in Windows.
///
/// # Arguments
///
/// * `path` - A string slice that holds the path to the folder that needs to be removed from
///   the frequent folders list.
///
/// # Returns
///
/// This function returns a `Result<(), WincentError>`. On success, it returns `Ok(())`. 
/// If an error occurs, it returns a `WincentError` which can indicate various issues:
/// - `IoError` if there is an issue accessing the file system, such as the path not existing
///   or being invalid.
/// - `InvalidData` if the specified path does not point to a directory.
/// - Any error returned from the `is_in_quick_access` function if it fails to check the folder's
///   presence in the frequent folders.
///
/// # Errors
///
/// The function may fail due to:
/// - The specified path not being accessible or not existing, resulting in an I/O error.
/// - The specified path not being a directory, leading to an `InvalidData` error.
/// - Issues with checking if the folder is in the quick access list.
///
/// # Example
///
/// ```
/// let result = remove_from_frequent_folders("C:\\path\\to\\folder").await;
/// match result {
///     Ok(_) => println!("Folder removed from frequent folders successfully!"),
///     Err(e) => eprintln!("Error removing folder: {:?}", e),
/// }
/// ```
pub async fn remove_from_frequent_folders(path: &str) -> Result<(), WincentError> {
    if let Err(e) = std::fs::metadata(path) {
        return Err(WincentError::IoError(e));
    }

    if !std::path::Path::new(path).is_dir() {
        return Err(WincentError::IoError(std::io::ErrorKind::InvalidData.into()));
    }

    let is_in_quick_access = match is_in_quick_access(vec![path], Some(QuickAccess::FrequentFolders)).await {
        Ok(result) => result,
        Err(e) => return Err(e),
    };

    if is_in_quick_access {
        handle_frequent_folders(path).await?;
    }

    Ok(())
}

/************************* Check/Set Visibility  *************************/

/// Retrieves the registry key for Quick Access settings in Windows.
///
/// This function attempts to open the registry key located at
/// `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer`.
/// If successful, it returns the corresponding `RegKey`. If it fails,
/// it returns a `WincentError` indicating the type of error encountered.
///
/// # Errors
///
/// This function can return the following errors:
/// - `WincentError::IoError`: If the registry key cannot be found or if there is an
///   I/O error while attempting to open the key.
///
/// # Examples
///
/// ```
/// match get_quick_access_reg() {
///     Ok(key) => println!("Successfully retrieved Quick Access registry key."),
///     Err(e) => eprintln!("Failed to retrieve Quick Access registry key: {:?}", e),
/// }
/// ```
fn get_quick_access_reg() -> Result<winreg::RegKey, WincentError> {
    use winreg::enums::*;
    use winreg::RegKey;

    let hklm = RegKey::predef(HKEY_CURRENT_USER);
    hklm.open_subkey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer")
        .map_err(WincentError::IoError)
}

/// Checks if the specified Quick Access target is visible in Windows.
///
/// This function retrieves the visibility setting for a given Quick Access target
/// (Frequent Folders, Recent Files, or All) from the Windows registry. It returns
/// `Ok(true)` if the target is visible, `Ok(false)` if it is not, or an error
/// if the operation fails.
///
/// # Parameters
///
/// - `target`: A `QuickAccess` enum value representing the target to check visibility for.
///
/// # Errors
///
/// This function can return the following errors:
/// - `WincentError::IoError`: If there is an I/O error while accessing the registry.
/// - `WincentError::ConvertError`: If there is an error converting the registry value.
///
/// # Examples
///
/// ```
/// match is_visible(QuickAccess::FrequentFolders) {
///     Ok(visible) => println!("Frequent Folders visibility: {}", visible),
///     Err(e) => eprintln!("Failed to check visibility: {:?}", e),
/// }
/// ```
pub fn is_visialbe(target: QuickAccess) -> Result<bool, WincentError> {
    let reg_key = get_quick_access_reg()?;
    let reg_value = match target {
        QuickAccess::FrequentFolders => "ShowFrequent",
        QuickAccess::RecentFiles => "ShowRecent",
        QuickAccess::All => "ShowRecent",
    };

    let val = reg_key.get_raw_value(reg_value).map_err(WincentError::IoError)?;

    let visibility = u32::from_ne_bytes(val.bytes[0..4].try_into().map_err(|e| WincentError::ConvertError(e))?);
    
    Ok(visibility != 0)
}

/// Sets the visibility of the specified Quick Access target in Windows.
///
/// This function updates the visibility setting for a given Quick Access target
/// (Frequent Folders, Recent Files, or All) in the Windows registry. It takes a
/// boolean value indicating whether the target should be visible (`true`) or not
/// (`false`). If the operation is successful, it returns `Ok(())`. If it fails,
/// it returns a `WincentError` indicating the type of error encountered.
///
/// # Parameters
///
/// - `target`: A `QuickAccess` enum value representing the target to set visibility for.
/// - `visible`: A boolean indicating the desired visibility state.
///
/// # Errors
///
/// This function can return the following errors:
/// - `WincentError::IoError`: If there is an I/O error while accessing the registry.
///
/// # Examples
///
/// ```
/// match set_visible(QuickAccess::FrequentFolders, true) {
///     Ok(_) => println!("Successfully set visibility for Frequent Folders."),
///     Err(e) => eprintln!("Failed to set visibility: {:?}", e),
/// }
/// ```
pub fn set_visiable(target: QuickAccess, visiable: bool) -> Result<(), WincentError> {
    let reg_key = get_quick_access_reg()?;
    let reg_value = match target {
        QuickAccess::FrequentFolders => "ShowFrequent",
        QuickAccess::RecentFiles => "ShowRecent",
        QuickAccess::All => "ShowRecent",
    };

    reg_key.set_value(reg_value, &u32::from(visiable)).map_err(WincentError::IoError)?;

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn init_logger() {
        let _ = env_logger::builder()
            .target(env_logger::Target::Stdout)
            .filter_level(log::LevelFilter::Trace)
            .is_test(true)
            .try_init();
    }

    #[tokio::test]
    async fn test_refresh() -> Result<(), WincentError> {
        refresh_explorer_window().await
    }

    #[tokio::test]
    async fn test_feasible() -> Result<(), WincentError> {
        match check_feasible().await {
            Ok(is_feasible) => {
                if !is_feasible {
                    if let Err(e) = fix_feasible().await {
                        panic!("Failed to fix feasibility: {:?}", e);
                    }
                }
    
                match check_feasible().await {
                    Ok(double_check) if double_check => {
                        assert!(true);
                    },
                    Err(e) => {
                        panic!("Error during double check: {:?}", e);
                    },
                    _ => {}
                }
    
                Ok(())
            },
            Err(e) => Err(e),
        }
    }

    #[tokio::test]
    async fn test_query_recent() -> Result<(), WincentError> {
        let recent_files = get_recent_files().await?;
        let frequent_folders = get_frequent_folders().await?;
        let quick_access = get_quick_access_items().await?;

        let seperate = recent_files.len() + frequent_folders.len();
        let total = quick_access.len();

        assert_eq!(seperate, total);

        Ok(())
    }

    #[tokio::test]
    async fn test_check_exists() -> Result<(), WincentError> {
        use std::path::Path;

        let quick_access = get_recent_files().await?;

        let full_path = quick_access[0].clone();
        let filename = Path::new(&full_path).file_name().unwrap().to_str().unwrap();
        let check_once = is_in_quick_access(vec![filename], None).await?;

        assert_eq!(check_once, true);

        let reversed: String = filename.chars().rev().collect();
        let check_twice = is_in_quick_access(vec![&reversed], None).await?;
        
        assert_eq!(check_twice, false);

        Ok(())
    }

    #[test]
    fn test_visiable() {
        init_logger();

        let iff = is_visialbe(QuickAccess::FrequentFolders).unwrap();
        let irf = is_visialbe(QuickAccess::RecentFiles).unwrap();

        assert_eq!(iff, true);
        assert_eq!(irf, true);
    }
}
