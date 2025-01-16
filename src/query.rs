use crate::{
    QuickAccess, 
    WincentResult,
    error::WincentError,
    scripts::{Script, execute_ps_script}
};


/// Queries recent items from Quick Access using a PowerShell script.
pub(crate) fn query_recent_with_ps_script(qa_type: QuickAccess) -> WincentResult<Vec<String>> {
    let output = match qa_type {
        QuickAccess::All => execute_ps_script(Script::QueryQuickAccess, None)?,
        QuickAccess::RecentFiles => execute_ps_script(Script::QuertRecentFile, None)?,
        QuickAccess::FrequentFolders => execute_ps_script(Script::QueryFrequentFolder, None)?,
    };

    if output.status.success() {
        let stdout_str = String::from_utf8(output.stdout)
            .map_err(|e| WincentError::Utf8(e))?;
        
        let data: Vec<String> = stdout_str
            .lines()
            .map(str::trim)
            .filter(|line| !line.is_empty())
            .map(String::from)
            .collect();

        Ok(data)
    } else {
        let error = String::from_utf8(output.stderr)?;
        Err(WincentError::ScriptFailed(error))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_query_recent_files() -> WincentResult<()> {
        let files = query_recent_with_ps_script(QuickAccess::RecentFiles)?;
        
        if !files.is_empty() {
            assert!(files.iter().all(|path| !path.is_empty()), "Paths should not be empty");
            
            for path in &files {
                assert!(path.contains(":\\"), "Path should be a valid Windows path format: {}", path);
            }
        }

        Ok(())
    }

    #[test]
    fn test_query_frequent_folders() -> WincentResult<()> {
        let folders = query_recent_with_ps_script(QuickAccess::FrequentFolders)?;
        
        if !folders.is_empty() {
            assert!(folders.iter().all(|path| !path.is_empty()), "Paths should not be empty");
            
            for path in &folders {
                assert!(path.contains(":\\"), "Path should be a valid Windows path format: {}", path);
            }
        }

        Ok(())
    }

    #[test_log::test]
    fn test_query_quick_access() -> WincentResult<()> {
        let items = query_recent_with_ps_script(QuickAccess::All)?;

        if !items.is_empty() {
            assert!(items.iter().all(|path| !path.is_empty()), "Paths should not be empty");
            
            for path in &items {
                assert!(path.contains(":\\"), "Path should be a valid Windows path format: {}", path);
            }
        }

        Ok(())
    }
}
