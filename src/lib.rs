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

pub enum SupportedOsVersion {
    Win10,
    Win11,
}

const SCRIPT_TIMEOUT: u64 = 3;

/************************* Utils *************************/
async fn refresh_explorer_window() -> Result<(), WincentError> {
    use powershell_script::PsScriptBuilder;

    const SCRIPT: &str = r#"
        $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
        $shellApplication = New-Object -ComObject Shell.Application;
        $windows = $shellApplication.Windows();
        $count = $windows.Count();

        foreach( $i in 0..($count-1) ) {
            $item = $windows.Item( $i )
            $item.Refresh() 
        }
    "#;

    let ps = PsScriptBuilder::new()
        .no_profile(true)
        .non_interactive(true)
        .hidden(false)
        .print_commands(false)
        .build();

    let handle = tokio::task::spawn_blocking(move || {
        match ps.run(SCRIPT) {
            Ok(_) => {
                return Ok(());
            },
            Err(e) => {
                return Err(WincentError::ScriptError(e))
            }
        }
    });

    match tokio::time::timeout(tokio::time::Duration::from_secs(SCRIPT_TIMEOUT), handle).await {
        Ok(res) => {
            match res {
                Ok(h_res) => {
                    return h_res;
                },
                Err(e) => {
                    return Err(WincentError::ExecuteError(e));
                }
            }
        },
        Err(e) => {
            return Err(WincentError::TimeoutError(e))
        } 
    }
}

fn check_os_version() -> Result<SupportedOsVersion, WincentError> {
    use sysinfo::System;

    if let Some(os) = System::name() {
        if os == "Windows" {
            match System::os_version() {
                Some(version) => {
                    if version.starts_with("10") {
                        return Ok(SupportedOsVersion::Win10);
                    } else if version.starts_with("11") {
                        return Ok(SupportedOsVersion::Win11);
                    }
                },
                None => {
                    return Err(WincentError::IoError(std::io::ErrorKind::Unsupported.into()));
                },
            }
        }
    }

    Err(WincentError::IoError(std::io::ErrorKind::Unsupported.into()))
}

/************************* Query Quick Access *************************/
async fn query_recent(recent_type: QuickAccess) -> Result<Vec<String>, WincentError> {
    use powershell_script::PsScriptBuilder;
    let shell_namespace: &str;
    let mut condition: &str = "";
    
    match recent_type {
        QuickAccess::FrequentFolders => shell_namespace = "3936E9E4-D92C-4EEE-A85A-BC16D5EA0819",
        QuickAccess::RecentFiles => {
            shell_namespace = "679f85cb-0220-4080-b29b-5540cc05aab6";
            condition = "| where {$_.IsFolder -eq $false}";
        },
        QuickAccess::All => shell_namespace = "679f85cb-0220-4080-b29b-5540cc05aab6",
    }

    let script: String = format!(r#"
            $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
            $shell = New-Object -ComObject Shell.Application;
            $shell.Namespace('shell:::{{{}}}').Items() {} | ForEach-Object {{ $_.Path }};
        "#, shell_namespace.to_string(), condition);

    let ps = PsScriptBuilder::new()
        .no_profile(true)
        .non_interactive(true)
        .hidden(false)
        .print_commands(false)
        .build();

    let _ = refresh_explorer_window();
    let handle = tokio::task::spawn_blocking(move || {
        match ps.run(&script) {
            Ok(output) => {
                let mut res: Vec<String> = vec![];
                if let Some(data) = output.stdout() {
                    let recents = data.split("\r\n");
                    for (_idx, item) in recents.enumerate() {
                        // debug!("{:?}. {:?}", idx+1, item);
                        if !item.is_empty() {
                            res.push(item.to_string());
                        } 
                    }
                }
            
                return Ok(res);
            },
            Err(e) => {
                return Err(WincentError::ScriptError(e))
            }
        }
    });

    match tokio::time::timeout(tokio::time::Duration::from_secs(SCRIPT_TIMEOUT), handle).await {
        Ok(res) => {
            match res {
                Ok(h_res) => {
                    return h_res;
                },
                Err(e) => {
                    return Err(WincentError::ExecuteError(e));
                }
            }
        },
        Err(e) => {
            return Err(WincentError::TimeoutError(e))
        } 
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
    let target_items: Vec<String>;

    if let Some(target) = specific_type {
        match target {
            QuickAccess::FrequentFolders => {
                target_items = get_frequent_folders().await?;
            },
            QuickAccess::RecentFiles => {
                target_items = get_recent_files().await?;
            },
            QuickAccess::All => {
                target_items = get_quick_access_items().await?;
            }
        }
    } else {
        target_items = get_quick_access_items().await?;
    }

    for item in target_items {
        for keyword in &keywords {
            if item.contains(*keyword) {
                return Ok(true);
            }
        }
    }

    Ok(false)
}


/************************* Check/Set Visibility  *************************/

fn get_quick_access_reg() -> Result<winreg::RegKey, WincentError> {
    use winreg::enums::*;
    use winreg::RegKey;

    let hklm = RegKey::predef(HKEY_CURRENT_USER);
    match hklm.open_subkey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer") {
        Ok(key) => {
            return Ok(key);
        },
        Err(e) => {
            return Err(WincentError::IoError(e));
        }
    }
}

pub fn is_visialbe(target: QuickAccess) -> Result<bool, WincentError> {
    let reg_key = get_quick_access_reg()?;
    let reg_value: &str;
    match target {
        QuickAccess::FrequentFolders => reg_value = "ShowFrequent",
        QuickAccess::RecentFiles => reg_value = "ShowRecent",
        QuickAccess::All => reg_value = "ShowRecent",
    }

    match reg_key.get_raw_value(reg_value) {
        Ok(val) => { // REG_DWORD 
            match &val.bytes[0..4].try_into() {
                Ok(arr) => Ok(u32::from_ne_bytes(*arr) != 0),
                Err(e) => Err(WincentError::ConvertError(*e)),
            }
        },
        Err(e) => {
            return Err(WincentError::IoError(e));
        }
    }
}

pub fn set_visiable(target: QuickAccess, visiable: bool) -> Result<(), WincentError> {
    let reg_key = get_quick_access_reg()?;
    let reg_value: &str;
    match target {
        QuickAccess::FrequentFolders => reg_value = "ShowFrequent",
        QuickAccess::RecentFiles => reg_value = "ShowRecent",
        QuickAccess::All => reg_value = "ShowRecent",
    }

    match reg_key.set_value(reg_value, &u32::from(visiable)) {
        Ok(_) => Ok(()),
        Err(e) => Err(WincentError::IoError(e)),
    }
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

    #[test]
    fn test_check_os_version() {
        if let Err(e) = check_os_version() {
            panic!("{:?}", e);
            
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
