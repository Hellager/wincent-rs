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

/************************* Utils *************************/
pub async fn refresh_explorer_window() -> Result<(), WincentError> {
    use powershell_script::PsScriptBuilder;

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
async fn query_recent(recent_type: QuickAccess) -> Result<Vec<String>, WincentError> {
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

pub async fn get_recent_files() -> Result<Vec<String>, WincentError> {
    query_recent(QuickAccess::RecentFiles).await
}

pub async fn get_frequent_folders() -> Result<Vec<String>, WincentError> {
    query_recent(QuickAccess::FrequentFolders).await
}

pub async fn get_quick_access_items() -> Result<Vec<String>, WincentError> {
    query_recent(QuickAccess::All).await
}

/************************* Check Existence *************************/

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

async fn handle_recent_files(path: &str, is_remove: bool) -> Result<(), WincentError> {
    use powershell_script::PsScriptBuilder;

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

pub async fn remove_from_recent_files(path: &str) -> Result<(), WincentError> {
    use std::fs;
    use std::path::Path;

    // Check if the file exists
    if fs::metadata(path).is_err() {
        return Err(WincentError::IoError(std::io::ErrorKind::NotFound.into()));
    }

    // Check if the path is a file
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

async fn handle_frequent_folders(path: &str) -> Result<(), WincentError> {
    use powershell_script::PsScriptBuilder;

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

fn get_quick_access_reg() -> Result<winreg::RegKey, WincentError> {
    use winreg::enums::*;
    use winreg::RegKey;

    let hklm = RegKey::predef(HKEY_CURRENT_USER);
    hklm.open_subkey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer")
        .map_err(WincentError::IoError)
}

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
