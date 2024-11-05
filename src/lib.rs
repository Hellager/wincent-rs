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
    ConvertError(std::array::TryFromSliceError)
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

/************************* Query Quick Access *************************/
fn query_recent(recent_type: QuickAccess) -> Result<Vec<String>, WincentError> {
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

    refresh_explorer_window();
    let output = ps.run(&script).unwrap();

    let mut res: Vec<String> = vec![];
    if let Some(data) = output.stdout() {
        let recents = data.split("\r\n");
        for (_idx, item) in recents.enumerate() {
            if !item.is_empty() {
                res.push(item.to_string());
            } 
        }
    }

    Ok(res)
}

pub fn get_recent_files() -> Result<Vec<String>, WincentError> {
    query_recent(QuickAccess::RecentFiles)
}

pub fn get_frequent_folders() -> Result<Vec<String>, WincentError> {
    query_recent(QuickAccess::FrequentFolders)
}

pub fn get_quick_access_items() -> Result<Vec<String>, WincentError> {
    query_recent(QuickAccess::All)
}

/************************* Check Existence *************************/

pub fn is_in_quick_access(keywords: Vec<&str>, specific_type: Option<QuickAccess>) -> Result<bool, WincentError> {
    let target_items: Vec<String>;

    if let Some(target) = specific_type {
        match target {
            QuickAccess::FrequentFolders => {
                target_items = get_frequent_folders()?;
            },
            QuickAccess::RecentFiles => {
                target_items = get_recent_files()?;
            },
            QuickAccess::All => {
                target_items = get_quick_access_items()?;
            }
        }
    } else {
        target_items = get_quick_access_items()?;
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

    #[test]
    fn test_refresh() -> Result<(), WincentError> {
        refresh_explorer_window()
    }

    #[test]
    fn test_query_recent() -> Result<(), WincentError> {
        let recent_files = get_recent_files()?;
        let frequent_folders = get_frequent_folders()?;
        let quick_access = get_quick_access_items()?;

        let seperate = recent_files.len() + frequent_folders.len();
        let total = quick_access.len();

        assert_eq!(seperate, total);

        Ok(())
    }

    #[test]
    fn test_check_exists() -> Result<(), WincentError> {
        use std::path::Path;

        let quick_access = get_recent_files()?;

        let full_path = quick_access[0].clone();
        let filename = Path::new(&full_path).file_name().unwrap().to_str().unwrap();
        let check_once = is_in_quick_access(vec![filename], None)?;

        assert_eq!(check_once, true);

        let reversed: String = filename.chars().rev().collect();
        let check_twice = is_in_quick_access(vec![&reversed], None)?;
        
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
