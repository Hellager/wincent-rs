use crate::{QuickAccess, WincentError};

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
/// match query_recent_with_ps_script(QuickAccess::RecentFiles).await {
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
pub(crate) async fn query_recent_with_ps_script(recent_type: QuickAccess) -> Result<Vec<String>, WincentError> {
    use powershell_script::PsScriptBuilder;

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
        .hidden(true)
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

    match tokio::time::timeout(tokio::time::Duration::from_secs(crate::SCRIPT_TIMEOUT), handle).await {
        Ok(res) => {
            res.map_err(WincentError::ExecuteError)?
        },
        Err(e) => Err(WincentError::TimeoutError(e)),
    }
}

#[cfg(test)]
mod query_test {
    use crate::utils::init_test_logger;
    use log::debug;
    use super::*;

    #[tokio::test]
    async fn test_print_out_quick_access() -> Result<(), WincentError> {
        init_test_logger();
        let quick_access = query_recent_with_ps_script(QuickAccess::All).await?;

        debug!("{} items in current quick access", quick_access.len());
        for (index, item) in quick_access.iter().enumerate() {
            debug!("{}. {}", index + 1, item);
        }

        Ok(())
    }
}
