use crate::WincentError;

// Handles recent files by removing a specified file from the recent files list.
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
pub(crate) async fn handle_recent_files_with_ps_script(path: &str, is_remove: bool) -> Result<(), WincentError> {
    use powershell_script::PsScriptBuilder;
    use std::io::{Error, ErrorKind};

    if !crate::check_feasible()? {
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

    match tokio::time::timeout(tokio::time::Duration::from_secs(crate::SCRIPT_TIMEOUT), handle).await {
        Ok(res) => {
            res.map(|_| ())
            .map_err(WincentError::ExecuteError)
        },
        Err(e) => Err(WincentError::TimeoutError(e)),
    }
}

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
pub(crate) async fn handle_frequent_folders_with_ps_script(path: &str) -> Result<(), WincentError> {
    use powershell_script::PsScriptBuilder;
    use std::io::{Error, ErrorKind};

    if !crate::check_feasible()? {
        return Err(WincentError::IoError(Error::from(ErrorKind::PermissionDenied)));
    }

    // `pintohome` is a toggle function, which means if you invoke it once, the folder pinned, invoke it twice, the folder will unpinned
    let script: String = format!(r#"
            $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
            $shell = New-Object -ComObject Shell.Application;
            $shell.Namespace("{}").Self.InvokeVerb("pintohome");
        "#, path);

    let mut vec_with_bom = Vec::new();
    vec_with_bom.extend_from_slice(&[0xEF, 0xBB, 0xBF]);
    vec_with_bom.extend_from_slice(script.as_bytes());  
    
    let bom_script = String::from_utf8(vec_with_bom).unwrap();

    let ps = PsScriptBuilder::new()
        .no_profile(true)
        .non_interactive(true)
        .hidden(true)
        .print_commands(false)
        .build();

    println!("print out bom script");
    println!("{:?}", bom_script);

    let handle = tokio::task::spawn_blocking(move || {
        ps.run(&bom_script).map_err(WincentError::ScriptError)
    });

    match tokio::time::timeout(tokio::time::Duration::from_secs(crate::SCRIPT_TIMEOUT), handle).await {
        Ok(res) => {
            res.map(|_| ())
            .map_err(WincentError::ExecuteError)
        },
        Err(e) => Err(WincentError::TimeoutError(e)),
    }
}

#[cfg(test)]
mod handle_test {
    use crate::{query::query_recent_with_ps_script, utils::init_test_logger};
    use log::debug;
    use super::*;

    #[ignore]
    #[tokio::test]
    async fn test_remove_recent_files() -> Result<(), WincentError> {
        init_test_logger();
        let recent_files = query_recent_with_ps_script(crate::QuickAccess::RecentFiles).await?;
        let target_file = recent_files[recent_files.len() - 1].clone();
        debug!("try remove file {} from recent files.", &target_file);
        let _ = handle_recent_files_with_ps_script(&target_file, true).await;

        let new_recent = query_recent_with_ps_script(crate::QuickAccess::RecentFiles).await?;

        let is_in_quick_access = new_recent.contains(&target_file);
        assert_eq!(is_in_quick_access, false);

        Ok(())
    }
    
    #[ignore]
    #[tokio::test]
    async fn test_handle_frequent_folders() -> Result<(), WincentError> {
        init_test_logger();
        let current_dir = std::env::current_dir().unwrap();

        debug!("try pin folder {:?} to frequent folders", current_dir.display());
        let current_dir_path = current_dir.to_str().unwrap();
        let _ = handle_frequent_folders_with_ps_script(current_dir_path).await?;

        let frequent_folders = query_recent_with_ps_script(crate::QuickAccess::FrequentFolders).await?;
        let is_in_quick_access = frequent_folders.contains(&current_dir_path.to_string());
        assert_eq!(is_in_quick_access, true);

        debug!("try unpin folder {:?} from frequent folders", current_dir.display());
        let frequent_folders = query_recent_with_ps_script(crate::QuickAccess::FrequentFolders).await?;
        let is_in_quick_access = frequent_folders.contains(&current_dir_path.to_string());
        assert_eq!(is_in_quick_access, false);

        Ok(())
    }
}
