use crate::error::{WincentError, WincentResult};
use std::io::Write;
use std::process::Command;
use tempfile::Builder;

pub(crate) enum Script {
    RefreshExplorer,
    QueryQuickAccess,
    QuertRecentFile,
    QueryFrequentFolder,
    RemoveRecentFile,
    PinToFrequentFolder,
    UnpinFromFrequentFolder,
}

static REFRESH_EXPLORER: &str = r#"
    $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
    $shellApplication = New-Object -ComObject Shell.Application;
    $windows = $shellApplication.Windows();
    $windows | ForEach-Object { $_.Refresh() }
"#;

static QUERY_RECENT_FILE: &str = r#"
    $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
    $shell = New-Object -ComObject Shell.Application;
    $shell.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}').Items() | where { $_.IsFolder -eq $false } | ForEach-Object { $_.Path };
"#;

static QUERY_FREQUENT_FOLDER: &str = r#"
    $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
    $shell = New-Object -ComObject Shell.Application;
    $shell.Namespace('shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}').Items() | ForEach-Object { $_.Path };
"#;

static QUERY_QUICK_ACCESS: &str = r#"
    $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
    $shell = New-Object -ComObject Shell.Application;
    $shell.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}').Items() | ForEach-Object { $_.Path };
"#;

/// Generates PowerShell script content based on the specified method and optional parameters.
///
/// This function returns the content of a PowerShell script as a string, depending on the provided `method`.
/// If the method requires parameters, it checks if they are provided and formats the script accordingly.
///
/// # Parameters
///
/// - `method`: An enum value of type `Script` that specifies which script to generate.
/// - `para`: An optional string slice that provides additional parameters for certain scripts.
///
/// # Returns
///
/// Returns a `WincentResult<String>`, which contains the generated PowerShell script content as a string.
/// If the operation is successful, it returns `Ok(content)`. If a required parameter is missing,
/// it returns `WincentError::MissingParameter`.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let script_content = get_script_content(Script::RemoveRecentFile, Some("C:\\Path\\To\\File.txt"))?;
///     println!("{}", script_content);
///     Ok(())
/// }
/// ```
pub(crate) fn get_script_content(method: Script, para: Option<&str>) -> WincentResult<String> {
    match method {
        Script::RefreshExplorer => Ok(REFRESH_EXPLORER.to_string()),
        Script::QuertRecentFile => Ok(QUERY_RECENT_FILE.to_string()),
        Script::QueryFrequentFolder => Ok(QUERY_FREQUENT_FOLDER.to_string()),
        Script::QueryQuickAccess => Ok(QUERY_QUICK_ACCESS.to_string()),
        Script::RemoveRecentFile => {
            if let Some(data) = para {
                let content = format!(r#"
                    $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
                    $shell = New-Object -ComObject Shell.Application;
                    $files = $shell.Namespace("shell:::{{679f85cb-0220-4080-b29b-5540cc05aab6}}").Items() | where {{$_.IsFolder -eq $false}};
                    $target = $files | where {{$_.Path -match ${}}};
                    $target.InvokeVerb("remove");
                "#, data);
                return Ok(content);                
            } else {
                return Err(WincentError::MissingParemeter);
            }
        },
        Script::PinToFrequentFolder => {
            if let Some(data) = para {
                let content = format!(r#"
                    $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
                    $shell = New-Object -ComObject Shell.Application;
                    $shell.Namespace("{}").Self.InvokeVerb("pintohome");
                "#, data);
                return Ok(content);                
            } else {
                return Err(WincentError::MissingParemeter);
            }
        },
        Script::UnpinFromFrequentFolder => {
            if let Some(data) = para {
                let content = format!(r#"
                    $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
                    $shell = New-Object -ComObject Shell.Application;
                    $folders = $shell.Namespace("shell:::{{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}}").Items();
                    $target = $folders | where {{$_.Path -match ${}}};
                    $target.InvokeVerb("unpinfromhome");
                "#, data);
                return Ok(content);                
            } else {
                return Err(WincentError::MissingParemeter);
            }
        },
    }
}

/// Executes a PowerShell script generated based on the specified method and optional parameters.
///
/// This function retrieves the content of a PowerShell script, writes it to a temporary file,
/// and then executes the script using PowerShell. It returns the output of the script execution.
///
/// # Parameters
///
/// - `method`: An enum value of type `Script` that specifies which script to execute.
/// - `para`: An optional string slice that provides additional parameters for certain scripts.
///
/// # Returns
///
/// Returns a `WincentResult<std::process::Output>`, which contains the output of the executed script.
/// If the operation is successful, it returns `Ok(output)`. If there is an error during script generation,
/// file creation, or execution, it returns an appropriate `WincentError`.
///
/// # Example
///
/// ```rust
/// fn main() -> Result<(), WincentError> {
///     let output = execute_ps_script(Script::QueryRecentFile, None)?;
///     println!("Script output: {:?}", output);
///     Ok(())
/// }
/// ```
pub(crate) fn execute_ps_script(method: Script, para: Option<&str>) -> WincentResult<std::process::Output> {
    let content = get_script_content(method, para)?;
    let temp_script_file = Builder::new()
        .prefix("wincent_")
        .suffix(".ps1")
        .rand_bytes(5)
        .tempfile()
        .map_err(|e| WincentError::Io(e))?;

    let bom = [0xEF, 0xBB, 0xBF];
    let mut file = temp_script_file.as_file();
    file.write_all(&bom)?;
    file.write_all(content.as_bytes())?;
    file.flush()?;

    Command::new("powershell")
        .args([
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            temp_script_file.into_temp_path().to_str().ok_or_else(|| 
                WincentError::InvalidPath("Failed to convert temp file path".to_string())
            )?,
        ])
        .output()
        .map_err(|e| WincentError::PowerShellExecution(e.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_get_pin_frequent_folder_script() {
        let path = "C:\\Users\\User\\Documents";
        let script = get_script_content(Script::PinToFrequentFolder, Some(path)).unwrap();
        assert!(script.contains("pintohome"));
    }

    #[test]
    fn test_get_unpin_frequent_folder_script() {
        let path = "C:\\Users\\User\\Documents";
        let script = get_script_content(Script::UnpinFromFrequentFolder, Some(path)).unwrap();
        assert!(script.contains("unpinfromhome"));
    }

    #[test]
    fn test_get_remove_recent_files_script() {
        let path = "C:\\Users\\User\\Documents";
        let script = get_script_content(Script::RemoveRecentFile, Some(path)).unwrap();
        assert!(script.contains("remove"));
    }

    #[test]
    fn test_script_content_validity() {
        let path = "C:\\Users\\User\\Documents";
        assert!(!get_script_content(Script::RefreshExplorer, None).unwrap().is_empty());
        assert!(!get_script_content(Script::QueryQuickAccess, None).unwrap().is_empty());
        assert!(!get_script_content(Script::QuertRecentFile, None).unwrap().is_empty());
        assert!(!get_script_content(Script::QueryFrequentFolder, None).unwrap().is_empty());
        assert!(!get_script_content(Script::RemoveRecentFile, Some(path)).unwrap().is_empty());
        assert!(!get_script_content(Script::PinToFrequentFolder, Some(path)).unwrap().is_empty());
        assert!(!get_script_content(Script::UnpinFromFrequentFolder, Some(path)).unwrap().is_empty());
    }
}
