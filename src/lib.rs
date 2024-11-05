use powershell_script::PsError;

pub enum QuickAccess {
    FrequentFolders,
    RecentFiles,
    All
}

#[derive(Debug)]
pub enum WincentError {
    ScriptError(PsError),
}

fn refresh_explorer_window() -> Result<(), WincentError> {
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

    match ps.run(SCRIPT) {
        Ok(_) => {
            return Ok(());
        },
        Err(e) => {
            return Err(WincentError::ScriptError(e))
        }
    }
}

pub fn get_frequent_folders() -> Result<Vec<String>, WincentError> {
    use powershell_script::{PsScriptBuilder, PsError};

    const SCRIPT: &str = r#"
        $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
        $shell = New-Object -ComObject Shell.Application;
        $paths = $shell.Namespace("shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}").Items() | ForEach-Object { $_.Path };
        $paths
    "#;

    let ps = PsScriptBuilder::new()
        .no_profile(true)
        .non_interactive(true)
        .hidden(true)
        .print_commands(false)
        .build();

    match ps.run(SCRIPT) {
        Ok(output) => {
            if let Some(data) = output.stdout() {
                let recents = data.split("\r\n");
                let mut folders: Vec<String> = vec![];
                for item in recents {
                    if !item.is_empty() {
                       folders.push(item.to_owned()); 
                    }
                }

                return Ok(folders);
            }
        },
        Err(e) => {
            return Err(WincentError::ScriptError(e))
        }
    }

    Err(WincentError::ScriptError(PsError::PowershellNotFound))
}

pub fn get_quick_access_items() -> Result<Vec<String>, WincentError> {
    use powershell_script::{PsScriptBuilder, PsError};

    const SCRIPT: &str = r#"
        $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
        $shell = New-Object -ComObject Shell.Application;
        $paths = $shell.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}").Items() | ForEach-Object { $_.Path };
        $paths
    "#;

    let ps = PsScriptBuilder::new()
        .no_profile(true)
        .non_interactive(true)
        .hidden(true)
        .print_commands(false)
        .build();

    match ps.run(SCRIPT) {
        Ok(output) => {
            if let Some(data) = output.stdout() {
                let recents = data.split("\r\n");
                let mut quick_access: Vec<String> = vec![];
                for item in recents {
                    if !item.is_empty() {
                        quick_access.push(item.to_owned()); 
                    }
                }

                return Ok(quick_access);
            }
        },
        Err(e) => {
            return Err(WincentError::ScriptError(e))
        }
    }

    Err(WincentError::ScriptError(PsError::PowershellNotFound))
}

// /// If given path is not exists, when call this script, the condition `where {$_.IsFolder -eq $false}` will return false even it used to be a file
// fn get_recent_files_old() -> Result<Vec<String>, WincentError> {
//     use powershell_script::PsScriptBuilder;

//     const SCRIPT: &str = r#"
//         $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
//         $shell = New-Object -ComObject Shell.Application;
//         $paths = $shell.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}").Items() | where {$_.IsFolder -eq $false} | ForEach-Object { $_.Path };
//         $paths
//     "#;

//     let ps = PsScriptBuilder::new()
//         .no_profile(true)
//         .non_interactive(true)
//         .hidden(true)
//         .print_commands(false)
//         .build();

//     match ps.run(SCRIPT) {
//         Ok(output) => {
//             if let Some(data) = output.stdout() {
//                 let recents = data.split("\r\n");
//                 let mut files: Vec<String> = vec![];
//                 for item in recents {
//                     if !item.is_empty() {
//                         files.push(item.to_owned()); 
//                     }
//                 }

//                 return Ok(files);
//             }
//         },
//         Err(e) => {
//             return Err(WincentError::ScriptError(e))
//         }
//     }

//     Err(WincentError::ScriptError(PsError::PowershellNotFound))
// }

pub fn get_recent_files() -> Result<Vec<String>, WincentError> {
    let mut files: Vec<String> = vec![];
    let quick_access = get_quick_access_items()?;
    let frequent_folders = get_frequent_folders()?;
    for item in quick_access {
        if !frequent_folders.contains(&item) {
            files.push(item);
        }
    }

    Ok(files)
}

pub fn is_in_quick_access(path: &str) -> Result<bool, WincentError> {
    let cur_quick_access = get_quick_access_items()?;

    for item in cur_quick_access {
        if item.contains(path) {
            return Ok(true);
        }
    }

    Ok(false)
}

fn get_quick_access_reg() -> Result<winreg::RegKey, std::io::Error> {
    use winreg::enums::*;
    use winreg::RegKey;

    let hklm = RegKey::predef(HKEY_CURRENT_USER);
    let cur_ver = hklm.open_subkey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer")?;

    Ok(cur_ver)
}

pub fn is_visialbe(target: QuickAccess) -> Result<bool, std::io::Error> {
    let reg_key = get_quick_access_reg()?;
    let reg_value: &str;
    match target {
        QuickAccess::FrequentFolders => reg_value = "ShowFrequent",
        QuickAccess::RecentFiles => reg_value = "ShowRecent",
        QuickAccess::All => reg_value = "ShowRecent",
    }
    let is_visiable: u32 = reg_key.get_value(reg_value)?;
    Ok(is_visiable != 0)
}

pub fn set_visiable(target: QuickAccess, visiable: bool) -> Result<(), std::io::Error> {
    let reg_key = get_quick_access_reg()?;
    let reg_value: &str;
    match target {
        QuickAccess::FrequentFolders => reg_value = "ShowFrequent",
        QuickAccess::RecentFiles => reg_value = "ShowRecent",
        QuickAccess::All => reg_value = "ShowRecent",
    }
    reg_key.set_value(reg_value, &u32::from(visiable))?;
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

    #[test]
    fn test_refresh() -> Result<(), WincentError> {
        refresh_explorer_window()
    }

    #[test]
    fn test_query_quick_access() -> Result<(), WincentError> {
        init_logger();

        let recent_files: Vec<String> = get_recent_files()?;
        let frequent_folders: Vec<String> = get_frequent_folders()?;
        let quick_access: Vec<String> = get_quick_access_items()?;

        // debug!("recent files");
        // for (idx, item) in recent_files.iter().enumerate() {
        //     debug!("{}. {}", idx, item);
        // }
        // debug!("frequent folders");
        // for (idx, item) in frequent_folders.iter().enumerate() {
        //     debug!("{}. {}", idx, item);
        // }
        // debug!("quick access items");
        // for (idx, item) in quick_access.iter().enumerate() {
        //     debug!("{}. {}", idx, item);
        // }

        assert_eq!(quick_access.len(), (recent_files.len() + frequent_folders.len()));

        for folder in frequent_folders {
            assert_eq!(quick_access.contains(&folder), true);
        }

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
