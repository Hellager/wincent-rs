use std::env;
use std::io::{self, Write};
use std::path::Path;
use std::time::Duration;

use wincent::error::WincentError;
use wincent::prelude::*;

fn main() {
    let args: Vec<String> = env::args().skip(1).collect();
    let result = if args.is_empty() {
        interactive_loop()
    } else {
        run(args)
    };

    if let Err(error) = result {
        print_error(&error);
        std::process::exit(1);
    }
}

fn interactive_loop() -> WincentResult<()> {
    println!("wincent interactive example CLI");
    println!("Type `help` to list commands, `exit` or `quit` to leave.");

    loop {
        print!("wincent> ");
        io::stdout().flush().map_err(WincentError::Io)?;

        let mut line = String::new();
        let bytes = io::stdin().read_line(&mut line).map_err(WincentError::Io)?;
        if bytes == 0 {
            println!();
            return Ok(());
        }

        let line = line.trim();
        if line.is_empty() {
            continue;
        }
        if matches!(line, "exit" | "quit") {
            return Ok(());
        }

        let args = match split_command_line(line) {
            Ok(args) => args,
            Err(error) => {
                print_error(&error);
                continue;
            }
        };

        if args.is_empty() {
            continue;
        }

        if let Err(error) = run(args) {
            print_error(&error);
        }
    }
}

fn run(args: Vec<String>) -> WincentResult<()> {
    if args.is_empty() || args[0] == "help" || args[0] == "--help" || args[0] == "-h" {
        print_help();
        return Ok(());
    }

    let (timeout, args) = parse_global_options(args)?;
    if args.is_empty() {
        print_help();
        return Ok(());
    }
    let manager = QuickAccessManager::builder().try_timeout(timeout)?.build();

    match args[0].as_str() {
        "list" => cmd_list(&manager, &args[1..]),
        "list-paths" => cmd_list_paths(&manager, &args[1..]),
        "contains-path" => cmd_contains_path(&manager, &args[1..]),
        "status" => cmd_status(&manager, &args[1..]),
        "contains-keyword" => cmd_contains_keyword(&manager, &args[1..]),
        "find" => cmd_find(&manager, &args[1..]),
        "add" => cmd_add(&manager, &args[1..]),
        "add-timed" => cmd_add_timed(&manager, &args[1..]),
        "remove" => cmd_remove(&manager, &args[1..]),
        "refresh" => cmd_refresh(&manager, &args[1..]),
        "lock" => cmd_lock(&manager, &args[1..]),
        "empty" => cmd_empty(&manager, &args[1..]),
        "visible" => cmd_visible(&manager, &args[1..]),
        other => Err(WincentError::InvalidArgument(format!(
            "unknown command: {other}"
        ))),
    }
}

fn parse_global_options(args: Vec<String>) -> WincentResult<(Duration, Vec<String>)> {
    let mut timeout = Duration::from_secs(10);
    let mut rest = Vec::new();
    let mut iter = args.into_iter();

    while let Some(arg) = iter.next() {
        match arg.as_str() {
            "--timeout-ms" => {
                let value = iter.next().ok_or_else(|| {
                    WincentError::InvalidArgument("--timeout-ms requires a value".to_string())
                })?;
                timeout = Duration::from_millis(parse_u64(&value, "timeout-ms")?);
            }
            "--timeout-secs" => {
                let value = iter.next().ok_or_else(|| {
                    WincentError::InvalidArgument("--timeout-secs requires a value".to_string())
                })?;
                timeout = Duration::from_secs(parse_u64(&value, "timeout-secs")?);
            }
            _ => rest.push(arg),
        }
    }

    Ok((timeout, rest))
}

fn print_help() {
    println!(
        r#"wincent example CLI

Usage:
  cargo run --example cli
  cargo run --example cli -- [--timeout-ms N] <command> [args]

Interactive:
  help
  exit
  quit

QuickAccessManager:
  list <recent|frequent|all> [--paths]
  list-paths <recent|frequent|all>
  contains-path <recent|frequent|all> <path>
  status <path>
    prints recent position, frequent position and pin state, or none.
  contains-keyword <recent|frequent|all> <keyword>
  find <recent|frequent|all> <keyword>
  add <path>
  add-timed <path>
    prints successful addition time as one JSON object.
  remove <path>
    detects Recent files and pinned or unpinned Frequent Folders automatically.
  refresh
    notifies the Shell, then refreshes open Explorer views.
  lock [recent|frequent|all] [--cleanup-new-links]
    Locked targets block Explorer from adding or removing matching Quick Access items.
  empty <recent|frequent|all>
  visible get <recent|frequent|start|all>
  visible set <recent|frequent|start|all> <true|false>
  The start target controls recent files in Windows Recommended items.
  Successful writes do not guarantee the UI changes immediately.
"#
    );
}

fn split_command_line(line: &str) -> WincentResult<Vec<String>> {
    let mut args = Vec::new();
    let mut current = String::new();
    let chars = line.chars();
    let mut quote = None;

    for ch in chars {
        match (quote, ch) {
            (Some(q), c) if c == q => quote = None,
            (Some(_), c) => current.push(c),
            (None, '"' | '\'') => quote = Some(ch),
            (None, c) if c.is_whitespace() => {
                if !current.is_empty() {
                    args.push(std::mem::take(&mut current));
                }
            }
            (None, c) => current.push(c),
        }
    }

    if let Some(q) = quote {
        return Err(WincentError::InvalidArgument(format!(
            "unterminated quote: {q}"
        )));
    }

    if !current.is_empty() {
        args.push(current);
    }

    Ok(args)
}

fn cmd_list(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "list <recent|frequent|all> [--paths]")?;
    let qa_type = parse_category(&args[0], true)?;

    if args.iter().skip(1).any(|arg| arg == "--paths") {
        return cmd_list_paths(manager, &args[..1]);
    }

    let items = manager.get_items(qa_type)?;
    print_strings("items", &items);
    Ok(())
}

fn cmd_list_paths(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "list-paths <recent|frequent|all>")?;
    let qa_type = parse_category(&args[0], true)?;
    let items = manager.get_item_paths(qa_type)?;

    println!("paths: {}", items.len());
    for path in items {
        println!("{}", path.display());
    }
    Ok(())
}

fn cmd_contains_path(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 2, "contains-path <recent|frequent|all> <path>")?;
    let qa_type = parse_category(&args[0], true)?;
    let exists = manager.contains_path(&args[1], qa_type)?;
    println!("{exists}");
    Ok(())
}

fn cmd_status(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "status <path>")?;
    let path = &args[0];
    let status = manager.get_item_status(path)?;
    let recent_file = match status.recent_file() {
        Some(position) => position.to_string(),
        None => "null".to_string(),
    };
    let frequent_folder = match status.frequent_folder() {
        Some(slot) => format!(
            "{{\"position\":{},\"pinned\":{}}}",
            slot.position(),
            slot.is_pinned()
        ),
        None => "null".to_string(),
    };
    let lists = match status.lists() {
        Some(QuickAccess::RecentFiles) => "recent",
        Some(QuickAccess::FrequentFolders) => "frequent",
        Some(QuickAccess::All) => "all",
        _ => "none",
    };
    println!(
        "{{\"path\":{path:?},\"recent_file\":{recent_file},\"frequent_folder\":{frequent_folder},\"lists\":\"{lists}\"}}"
    );
    Ok(())
}

fn cmd_contains_keyword(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 2, "contains-keyword <recent|frequent|all> <keyword>")?;
    let qa_type = parse_category(&args[0], true)?;
    let exists = manager.contains_keyword(&args[1], qa_type)?;
    println!("{exists}");
    Ok(())
}

fn cmd_find(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 2, "find <recent|frequent|all> <keyword>")?;
    let qa_type = parse_category(&args[0], true)?;
    let matches = manager.find_items_containing(&args[1], qa_type)?;
    print_strings("matches", &matches);
    Ok(())
}

fn cmd_add(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "add <path>")?;
    manager.add_item(&args[0])?;
    println!("added {}", args[0]);
    Ok(())
}

fn cmd_add_timed(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "add-timed <path>")?;
    let started = std::time::Instant::now();
    manager.add_item(&args[0])?;
    println!(
        "{{\"mode\":\"shell\",\"category\":\"auto\",\"total_ns\":{}}}",
        started.elapsed().as_nanos()
    );
    Ok(())
}

fn cmd_remove(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "remove <path>")?;
    manager.remove_item(&args[0])?;
    println!("removed {}", args[0]);
    Ok(())
}

fn cmd_refresh(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    if !args.is_empty() {
        return Err(WincentError::InvalidArgument("usage: refresh".to_string()));
    }
    manager.refresh()?;
    println!("refreshed Explorer");
    Ok(())
}

fn cmd_lock(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    let usage = || {
        WincentError::InvalidArgument(
            "usage: lock [recent|frequent|all] [--cleanup-new-links]".to_string(),
        )
    };
    let mut cleanup = false;
    let mut target = None;
    for arg in args {
        if arg == "--cleanup-new-links" {
            if cleanup {
                return Err(usage());
            }
            cleanup = true;
        } else {
            if target.is_some() {
                return Err(usage());
            }
            target = Some(parse_category(arg, true)?);
        }
    }
    let target = target.unwrap_or(QuickAccess::All);
    manager.lock(target)?;
    println!("locked Quick Access backing files");
    println!("target: {target:?}");
    println!("press Enter to unlock");

    let mut line = String::new();
    io::stdin().read_line(&mut line).map_err(WincentError::Io)?;

    if cleanup {
        let report = manager.unlock_and_cleanup()?;
        println!("cleaned_shortcuts: {}", report.cleaned_shortcuts().len());
        for path in report.cleaned_shortcuts() {
            println!("  cleaned {}", path.display());
        }
        for failure in report.cleanup_failures() {
            println!(
                "  failed to clean {}: {}",
                failure.path().display(),
                failure.error()
            );
        }
    } else {
        manager.unlock()?;
        println!("released lock without shortcut cleanup");
    }
    Ok(())
}

fn cmd_empty(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "empty <recent|frequent|all>")?;
    if args.len() != 1 {
        return Err(WincentError::InvalidArgument(
            "usage: empty <recent|frequent|all>".to_string(),
        ));
    }
    let qa_type = parse_category(&args[0], true)?;

    manager.empty_items(qa_type)?;
    println!("cleared {}", category_name(qa_type));
    Ok(())
}

fn cmd_visible(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "visible <command>")?;
    match args[0].as_str() {
        "get" => {
            require_exact_len(&args[1..], 1, "visible get <recent|frequent|start|all>")?;
            let target = parse_visibility_target(&args[1])?;
            println!("{}", manager.is_visible(target)?);
        }
        "set" => {
            require_exact_len(
                &args[1..],
                2,
                "visible set <recent|frequent|start|all> <true|false>",
            )?;
            let target = parse_visibility_target(&args[1])?;
            manager.set_visible(target, parse_bool(&args[2])?)?;
            println!("updated visibility");
        }
        other => {
            return Err(WincentError::InvalidArgument(format!(
                "unknown visible command: {other}"
            )));
        }
    }
    Ok(())
}

fn parse_category(value: &str, allow_all: bool) -> WincentResult<QuickAccess> {
    match value {
        "recent" | "recent-files" | "files" => Ok(QuickAccess::RecentFiles),
        "frequent" | "frequent-folders" | "folders" => Ok(QuickAccess::FrequentFolders),
        "all" if allow_all => Ok(QuickAccess::All),
        "all" => Err(WincentError::UnsupportedOperation(
            "QuickAccess::All is not valid for this operation".to_string(),
        )),
        other => Err(WincentError::InvalidArgument(format!(
            "unknown Quick Access category: {other}"
        ))),
    }
}

fn parse_visibility_target(value: &str) -> WincentResult<VisibilityTarget> {
    match value {
        "recent" | "recent-files" | "files" => Ok(VisibilityTarget::RecentFiles),
        "frequent" | "frequent-folders" | "folders" => Ok(VisibilityTarget::FrequentFolders),
        "start" | "start-recommended" => Ok(VisibilityTarget::StartRecommended),
        "all" => Ok(VisibilityTarget::All),
        other => Err(WincentError::InvalidArgument(format!(
            "unknown visibility target: {other}"
        ))),
    }
}

fn category_name(qa_type: QuickAccess) -> &'static str {
    match qa_type {
        QuickAccess::RecentFiles => "Recent Files",
        QuickAccess::FrequentFolders => "Frequent Folders",
        QuickAccess::All => "All",
        _ => "Unknown",
    }
}

fn print_strings(label: &str, values: &[String]) {
    println!("{label}: {}", values.len());
    for value in values {
        println!("{value}");
    }
}

fn print_error(error: &WincentError) {
    eprintln!("error: {error}");
    match error {
        WincentError::InvalidPath(error) => {
            eprintln!("invalid_path.reason: {}", error.reason_text());
            eprintln!("invalid_path.path: {:?}", error.path().map(Path::display));
        }
        WincentError::AlreadyExists { path, qa_type, .. } => {
            eprintln!("already_exists.path: {path}");
            eprintln!("already_exists.qa_type: {qa_type:?}");
        }
        WincentError::NotInQuickAccess { path, qa_type, .. } => {
            eprintln!("not_in_quick_access.path: {path}");
            eprintln!("not_in_quick_access.qa_type: {qa_type:?}");
        }
        WincentError::PartialEmpty {
            recent_files_cleared,
            frequent_folders_cleared,
            source,
            ..
        } => {
            eprintln!("partial_empty.recent_files_cleared: {recent_files_cleared}");
            eprintln!("partial_empty.frequent_folders_cleared: {frequent_folders_cleared}");
            eprintln!("partial_empty.source: {source}");
        }
        _ => {}
    }
}

fn parse_bool(value: &str) -> WincentResult<bool> {
    match value {
        "true" | "1" | "yes" | "on" | "show" => Ok(true),
        "false" | "0" | "no" | "off" | "hide" => Ok(false),
        other => Err(WincentError::InvalidArgument(format!(
            "invalid bool value: {other}"
        ))),
    }
}

fn parse_u64(value: &str, name: &str) -> WincentResult<u64> {
    value
        .parse::<u64>()
        .map_err(|_| WincentError::InvalidArgument(format!("invalid {name}")))
}

fn require_len(args: &[String], min: usize, usage: &str) -> WincentResult<()> {
    if args.len() < min {
        Err(WincentError::InvalidArgument(format!("usage: {usage}")))
    } else {
        Ok(())
    }
}

fn require_exact_len(args: &[String], expected: usize, usage: &str) -> WincentResult<()> {
    if args.len() != expected {
        Err(WincentError::InvalidArgument(format!("usage: {usage}")))
    } else {
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn split_command_line_preserves_quoted_windows_path_backslashes() {
        let args = split_command_line(r#"remove "D:\Temp\wincent-rs-test\scripts\Copy-Admin.ps1""#)
            .unwrap();

        assert_eq!(
            args,
            vec![
                "remove",
                r#"D:\Temp\wincent-rs-test\scripts\Copy-Admin.ps1"#
            ]
        );
    }

    #[test]
    fn split_command_line_preserves_quoted_path_with_spaces() {
        let args = split_command_line(r#"remove "D:\Temp\wincent rs test\file.ps1""#).unwrap();

        assert_eq!(args, vec!["remove", r#"D:\Temp\wincent rs test\file.ps1"#]);
    }

    #[test]
    fn split_command_line_preserves_trailing_backslash_before_closing_quote() {
        let args = split_command_line(r#"list-paths "D:\Temp\wincent-rs-test\""#).unwrap();

        assert_eq!(args, vec!["list-paths", r#"D:\Temp\wincent-rs-test\"#]);
    }

    #[test]
    fn split_command_line_reports_unterminated_quote() {
        let error = split_command_line(r#"remove "D:\Temp"#).unwrap_err();

        assert!(
            matches!(error, WincentError::InvalidArgument(message) if message == "unterminated quote: \"")
        );
    }

    #[test]
    fn visibility_target_parses_start_without_changing_quick_access_categories() {
        assert_eq!(
            parse_visibility_target("start").unwrap(),
            VisibilityTarget::StartRecommended
        );
        assert_eq!(
            parse_visibility_target("all").unwrap(),
            VisibilityTarget::All
        );
        assert!(parse_category("start", true).is_err());
    }

    #[test]
    fn visible_set_start_requires_bool_argument() {
        let error = require_exact_len(
            &["start".to_string()],
            2,
            "visible set <recent|frequent|start|all> <true|false>",
        )
        .unwrap_err();

        assert!(
            matches!(error, WincentError::InvalidArgument(message) if message == "usage: visible set <recent|frequent|start|all> <true|false>")
        );
    }

    #[test]
    fn visible_set_rejects_removed_refresh_option_before_writing() {
        let error = require_exact_len(
            &[
                "recent".to_string(),
                "true".to_string(),
                "--refresh-explorer".to_string(),
            ],
            2,
            "visible set <recent|frequent|start|all> <true|false>",
        )
        .unwrap_err();

        assert!(matches!(error, WincentError::InvalidArgument(_)));
    }
}
