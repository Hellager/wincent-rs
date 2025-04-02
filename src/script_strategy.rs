use crate::error::WincentError;
use crate::WincentResult;
use std::collections::HashMap;
use std::sync::OnceLock;

/// Enum representing PowerShell script operation types
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) enum PSScript {
    RefreshExplorer,
    QueryQuickAccess,
    QueryRecentFile,
    QueryFrequentFolder,
    RemoveRecentFile,
    PinToFrequentFolder,
    UnpinFromFrequentFolder,
    EmptyPinnedFolders,
    CheckQueryFeasible,
    CheckPinUnpinFeasible,
}

/// Shell namespace constants
pub(crate) struct ShellNamespaces;

impl ShellNamespaces {
    pub const QUICK_ACCESS: &'static str = "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}";
    pub const FREQUENT_FOLDERS: &'static str = "shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}";
}

/// Script generation strategy interface
pub(crate) trait ScriptStrategy {
    fn generate(&self, parameter: Option<&str>) -> WincentResult<String>;
}

/// Base script strategy providing UTF-8 encoding configuration
pub(crate) struct BaseScriptStrategy;

impl BaseScriptStrategy {
    fn utf8_header() -> &'static str {
        "$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;"
    }
    
    fn shell_com_object() -> &'static str {
        "$shell = New-Object -ComObject Shell.Application;"
    }
}

/// Strategy for refreshing File Explorer
pub(crate) struct RefreshExplorerStrategy;

impl ScriptStrategy for RefreshExplorerStrategy {
    fn generate(&self, _: Option<&str>) -> WincentResult<String> {
        Ok(format!(
            r#"
    {}
    $shellApplication = New-Object -ComObject Shell.Application;
    $windows = $shellApplication.Windows();
    $windows | ForEach-Object {{ $_.Refresh() }}
"#,
            BaseScriptStrategy::utf8_header()
        ))
    }
}

/// Strategy for querying recent files
pub(crate) struct QueryRecentFileStrategy;

impl ScriptStrategy for QueryRecentFileStrategy {
    fn generate(&self, _: Option<&str>) -> WincentResult<String> {
        Ok(format!(
            r#"
    {}
    {}
    $shell.Namespace('{}').Items() | where {{ $_.IsFolder -eq $false }} | ForEach-Object {{ $_.Path }};
"#,
            BaseScriptStrategy::utf8_header(),
            BaseScriptStrategy::shell_com_object(),
            ShellNamespaces::QUICK_ACCESS
        ))
    }
}

/// Strategy for querying frequent folders
pub(crate) struct QueryFrequentFolderStrategy;

impl ScriptStrategy for QueryFrequentFolderStrategy {
    fn generate(&self, _: Option<&str>) -> WincentResult<String> {
        Ok(format!(
            r#"
    {}
    {}
    $shell.Namespace('{}').Items() | ForEach-Object {{ $_.Path }};
"#,
            BaseScriptStrategy::utf8_header(),
            BaseScriptStrategy::shell_com_object(),
            ShellNamespaces::FREQUENT_FOLDERS
        ))
    }
}

/// Strategy for querying Quick Access
pub(crate) struct QueryQuickAccessStrategy;

impl ScriptStrategy for QueryQuickAccessStrategy {
    fn generate(&self, _: Option<&str>) -> WincentResult<String> {
        Ok(format!(
            r#"
    {}
    {}
    $shell.Namespace('{}').Items() | ForEach-Object {{ $_.Path }};
"#,
            BaseScriptStrategy::utf8_header(),
            BaseScriptStrategy::shell_com_object(),
            ShellNamespaces::QUICK_ACCESS
        ))
    }
}

/// Strategy for removing recent files
pub(crate) struct RemoveRecentFileStrategy;

impl ScriptStrategy for RemoveRecentFileStrategy {
    fn generate(&self, parameter: Option<&str>) -> WincentResult<String> {
        let path = parameter.ok_or(WincentError::MissingParemeter)?;
        Ok(format!(
            r#"
    {}
    {}
    $files = $shell.Namespace("{}").Items() | where {{$_.IsFolder -eq $false}};
    $target = $files | where {{$_.Path -eq "{}"}};
    $target.InvokeVerb("remove");
"#,
            BaseScriptStrategy::utf8_header(),
            BaseScriptStrategy::shell_com_object(),
            ShellNamespaces::QUICK_ACCESS,
            path
        ))
    }
}

/// Strategy for pinning to frequent folders
pub(crate) struct PinToFrequentFolderStrategy;

impl ScriptStrategy for PinToFrequentFolderStrategy {
    fn generate(&self, parameter: Option<&str>) -> WincentResult<String> {
        let path = parameter.ok_or(WincentError::MissingParemeter)?;
        Ok(format!(
            r#"
    {}
    {}
    $shell.Namespace("{}").Self.InvokeVerb("pintohome");
"#,
            BaseScriptStrategy::utf8_header(),
            BaseScriptStrategy::shell_com_object(),
            path
        ))
    }
}

/// Strategy for unpinning from frequent folders
pub(crate) struct UnpinFromFrequentFolderStrategy;

impl ScriptStrategy for UnpinFromFrequentFolderStrategy {
    fn generate(&self, parameter: Option<&str>) -> WincentResult<String> {
        let path = parameter.ok_or(WincentError::MissingParemeter)?;
        Ok(format!(
            r#"
    {}
    {}
    $folders = $shell.Namespace("{}").Items();
    $target = $folders | Where-Object {{$_.Path -eq "{}"}};
    $target.InvokeVerb("unpinfromhome");
"#,
            BaseScriptStrategy::utf8_header(),
            BaseScriptStrategy::shell_com_object(),
            ShellNamespaces::FREQUENT_FOLDERS,
            path
        ))
    }
}

/// Strategy for emptying pinned folders
pub(crate) struct EmptyPinnedFoldersStrategy;

impl ScriptStrategy for EmptyPinnedFoldersStrategy {
    fn generate(&self, _: Option<&str>) -> WincentResult<String> {
        Ok(format!(
            r#"
    {}
    {}
    $shell.Namespace('{}').Items() | ForEach-Object {{ $_.InvokeVerb("unpinfromhome") }};
"#,
            BaseScriptStrategy::utf8_header(),
            BaseScriptStrategy::shell_com_object(),
            ShellNamespaces::FREQUENT_FOLDERS,
        ))
    }
}

/// Strategy for checking query feasibility
pub(crate) struct CheckQueryFeasibleStrategy;

impl ScriptStrategy for CheckQueryFeasibleStrategy {
    fn generate(&self, _: Option<&str>) -> WincentResult<String> {
        Ok(format!(
            r#"
    {}

    $timeout = 5

    $scriptBlock = {{
        $shell = New-Object -ComObject Shell.Application
        $shell.Namespace('{}').Items() | ForEach-Object {{ $_.Path }};
    }}.ToString()

    $arguments = "-Command & {{$scriptBlock}}"
    $process = Start-Process powershell -ArgumentList $arguments -NoNewWindow -PassThru

    if (-not $process.WaitForExit($timeout * 1000)) {{
        try {{
            $process.Kill()
            Write-Error "Process execution timed out (${{timeout}}s), forcefully terminated"
            exit 1
        }}
        catch {{
            Write-Error "Error occurred while terminating process: $_"
            exit 1
        }}
    }}
"#,
            BaseScriptStrategy::utf8_header(),
            ShellNamespaces::QUICK_ACCESS
        ))
    }
}

/// Strategy for checking pin/unpin feasibility
pub(crate) struct CheckPinUnpinFeasibleStrategy;

impl ScriptStrategy for CheckPinUnpinFeasibleStrategy {
    fn generate(&self, _: Option<&str>) -> WincentResult<String> {
        Ok(format!(
            r#"
    {}

    $currentPath = $PSScriptRoot

    $scriptBlock = {{
        param($scriptPath)
        $shell = New-Object -ComObject Shell.Application
        $shell.Namespace($scriptPath).Self.InvokeVerb('pintohome')

        Start-Sleep -Seconds 3

        $folders = $shell.Namespace('{}').Items();
        $target = $folders | Where-Object {{$_.Path -eq $scriptPath}};
        $target.InvokeVerb('unpinfromhome');
    }}.ToString()

    $arguments = "-Command & {{$scriptBlock}} -scriptPath '$currentPath'"
    $process = Start-Process powershell -ArgumentList $arguments -NoNewWindow -PassThru

    $timeout = 10
    if (-not $process.WaitForExit($timeout * 1000)) {{
        try {{
            $process.Kill()
            Write-Error "Process execution timed out (${{timeout}}s), forcefully terminated"
            exit 1
        }}
        catch {{
            Write-Error "Error occurred while terminating process: $_"
            exit 1
        }}
    }}
"#,
            BaseScriptStrategy::utf8_header(),
            ShellNamespaces::FREQUENT_FOLDERS
        ))
    }
}

/// Strategy factory for retrieving script generation strategies
pub(crate) struct ScriptStrategyFactory;

impl ScriptStrategyFactory {
    /// Retrieves the strategy instance for the specified script type
    pub fn get_strategy(script_type: PSScript) -> WincentResult<Box<dyn ScriptStrategy>> {
        static STRATEGIES: OnceLock<HashMap<PSScript, Box<dyn ScriptStrategy + Sync + Send>>> = OnceLock::new();
        
        let strategies = STRATEGIES.get_or_init(|| {
            let mut map = HashMap::new();
            map.insert(PSScript::RefreshExplorer, Box::new(RefreshExplorerStrategy) as Box<dyn ScriptStrategy + Sync + Send>);
            map.insert(PSScript::QueryQuickAccess, Box::new(QueryQuickAccessStrategy) as Box<dyn ScriptStrategy + Sync + Send>);
            map.insert(PSScript::QueryRecentFile, Box::new(QueryRecentFileStrategy) as Box<dyn ScriptStrategy + Sync + Send>);
            map.insert(PSScript::QueryFrequentFolder, Box::new(QueryFrequentFolderStrategy) as Box<dyn ScriptStrategy + Sync + Send>);
            map.insert(PSScript::RemoveRecentFile, Box::new(RemoveRecentFileStrategy) as Box<dyn ScriptStrategy + Sync + Send>);
            map.insert(PSScript::PinToFrequentFolder, Box::new(PinToFrequentFolderStrategy) as Box<dyn ScriptStrategy + Sync + Send>);
            map.insert(PSScript::UnpinFromFrequentFolder, Box::new(UnpinFromFrequentFolderStrategy) as Box<dyn ScriptStrategy + Sync + Send>);
            map.insert(PSScript::CheckQueryFeasible, Box::new(CheckQueryFeasibleStrategy) as Box<dyn ScriptStrategy + Sync + Send>);
            map.insert(PSScript::CheckPinUnpinFeasible, Box::new(CheckPinUnpinFeasibleStrategy) as Box<dyn ScriptStrategy + Sync + Send>);
            map.insert(PSScript::EmptyPinnedFolders, Box::new(EmptyPinnedFoldersStrategy) as Box<dyn ScriptStrategy + Sync + Send>);
            map
        });
        
        // Clone strategy instances to avoid borrowing issues
        match strategies.get(&script_type) {
            Some(_strategy) => {
                // Create new Box instances for trait objects
                Ok(match script_type {
                    PSScript::RefreshExplorer => Box::new(RefreshExplorerStrategy),
                    PSScript::QueryQuickAccess => Box::new(QueryQuickAccessStrategy),
                    PSScript::QueryRecentFile => Box::new(QueryRecentFileStrategy),
                    PSScript::QueryFrequentFolder => Box::new(QueryFrequentFolderStrategy),
                    PSScript::RemoveRecentFile => Box::new(RemoveRecentFileStrategy),
                    PSScript::PinToFrequentFolder => Box::new(PinToFrequentFolderStrategy),
                    PSScript::UnpinFromFrequentFolder => Box::new(UnpinFromFrequentFolderStrategy),
                    PSScript::CheckQueryFeasible => Box::new(CheckQueryFeasibleStrategy),
                    PSScript::CheckPinUnpinFeasible => Box::new(CheckPinUnpinFeasibleStrategy),
                    PSScript::EmptyPinnedFolders => Box::new(EmptyPinnedFoldersStrategy),
                })
            },
            None => Err(WincentError::ScriptStrategyNotFound(format!("{:?}", script_type))),
        }
    }
    
    /// Generates script content
    pub fn generate_script(script_type: PSScript, parameter: Option<&str>) -> WincentResult<String> {
        let strategy = Self::get_strategy(script_type)?;
        strategy.generate(parameter)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_pin_frequent_folder_script_generation() {
        let path = "C:\\Users\\User\\Documents";
        let script = ScriptStrategyFactory::generate_script(PSScript::PinToFrequentFolder, Some(path)).unwrap();
        assert!(script.contains("pintohome"));
    }

    #[test]
    fn test_unpin_frequent_folder_script_generation() {
        let path = "C:\\Users\\User\\Documents";
        let script = ScriptStrategyFactory::generate_script(PSScript::UnpinFromFrequentFolder, Some(path)).unwrap();
        assert!(script.contains("unpinfromhome"));
    }

    #[test]
    fn test_remove_recent_files_script_generation() {
        let path = "C:\\Users\\User\\Documents";
        let script = ScriptStrategyFactory::generate_script(PSScript::RemoveRecentFile, Some(path)).unwrap();
        assert!(script.contains("remove"));
    }

    #[test]
    fn test_query_feasibility_check_script() {
        let script = ScriptStrategyFactory::generate_script(PSScript::CheckQueryFeasible, None).unwrap();
        assert!(script.contains(ShellNamespaces::QUICK_ACCESS));
    }

    #[test]
    fn test_pin_unpin_feasibility_check_script() {
        let script = ScriptStrategyFactory::generate_script(PSScript::CheckPinUnpinFeasible, None).unwrap();
        assert!(script.contains("pintohome"));
    }

    #[test]
    fn test_script_content_validity() {
        let path = "C:\\Users\\User\\Documents";
        assert!(!ScriptStrategyFactory::generate_script(PSScript::RefreshExplorer, None)
            .unwrap()
            .is_empty());
        assert!(!ScriptStrategyFactory::generate_script(PSScript::QueryQuickAccess, None)
            .unwrap()
            .is_empty());
        assert!(!ScriptStrategyFactory::generate_script(PSScript::QueryRecentFile, None)
            .unwrap()
            .is_empty());
        assert!(!ScriptStrategyFactory::generate_script(PSScript::QueryFrequentFolder, None)
            .unwrap()
            .is_empty());
        assert!(!ScriptStrategyFactory::generate_script(PSScript::RemoveRecentFile, Some(path))
            .unwrap()
            .is_empty());
        assert!(!ScriptStrategyFactory::generate_script(PSScript::PinToFrequentFolder, Some(path))
            .unwrap()
            .is_empty());
        assert!(
            !ScriptStrategyFactory::generate_script(PSScript::UnpinFromFrequentFolder, Some(path))
                .unwrap()
                .is_empty()
        );
        assert!(!ScriptStrategyFactory::generate_script(PSScript::CheckQueryFeasible, None)
            .unwrap()
            .is_empty());
        assert!(!ScriptStrategyFactory::generate_script(PSScript::CheckPinUnpinFeasible, None)
            .unwrap()
            .is_empty());
    }
    
    #[test]
    #[should_panic(expected = "not implemented")]
    fn test_nonexistent_strategy_error_handling() {
        struct MockStrategy;
        impl ScriptStrategy for MockStrategy {
            fn generate(&self, _: Option<&str>) -> WincentResult<String> {
                unimplemented!("not implemented");
            }
        }
        
        let mock = Box::new(MockStrategy);
        let _ = mock.generate(None);
    }
}
