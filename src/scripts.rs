use crate::{WincentResult, error::WincentError};
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
    CheckQueryFeasible,
    CheckPinUnpinFeasible,
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

static CHECK_QUERY_FEASIBLE: &str = r#"
    $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;

    $timeout = 5

    $scriptBlock = {
        $shell = New-Object -ComObject Shell.Application
        $shell.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}').Items() | ForEach-Object { $_.Path };
    }.ToString()

    $arguments = "-Command & {$scriptBlock}"
    $process = Start-Process powershell -ArgumentList $arguments -NoNewWindow -PassThru

    if (-not $process.WaitForExit($timeout * 1000)) {
        try {
            $process.Kill()
            Write-Error "Process execution timed out (${timeout}s), forcefully terminated"
            exit 1
        }
        catch {
            Write-Error "Error occurred while terminating process: $_"
            exit 1
        }
    }
"#;

static CHECK_PIN_UNPIN_FEASIBLE: &str = r#"
    $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;

    $currentPath = $PSScriptRoot

    $scriptBlock = {
        param($scriptPath)
        $shell = New-Object -ComObject Shell.Application
        $shell.Namespace($scriptPath).Self.InvokeVerb('pintohome')

        $folders = $shell.Namespace('shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}').Items();
        $target = $folders | Where-Object {$_.Path -match ${$scriptPath}};
        $target.InvokeVerb('unpinfromhome');
    }.ToString()

    $arguments = "-Command & {$scriptBlock} -scriptPath '$currentPath'"
    $process = Start-Process powershell -ArgumentList $arguments -NoNewWindow -PassThru

    $timeout = 5
    if (-not $process.WaitForExit($timeout * 1000)) {
        try {
            $process.Kill()
            Write-Error "Process execution timed out (${timeout}s), forcefully terminated"
            exit 1
        }
        catch {
            Write-Error "Error occurred while terminating process: $_"
            exit 1
        }
    }
"#;

/// Generates PowerShell script content based on the specified method and optional parameters.
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
                    $target = $files | where {{$_.Path -eq "{}"}};
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
                    $target = $folders | Where-Object {{$_.Path -eq "{}"}};
                    $target.InvokeVerb("unpinfromhome");
                "#, data);
                return Ok(content);                
            } else {
                return Err(WincentError::MissingParemeter);
            }
        },
        Script::CheckQueryFeasible => Ok(CHECK_QUERY_FEASIBLE.to_string()),
        Script::CheckPinUnpinFeasible => Ok(CHECK_PIN_UNPIN_FEASIBLE.to_string()),
    }
}

/// Executes a PowerShell script generated based on the specified method and optional parameters.
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
