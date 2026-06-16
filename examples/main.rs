use std::env;
use std::io::{self, Write};
use std::path::Path;
use std::path::PathBuf;
use std::time::Duration;

use wincent::error::{InvalidPathError, PowerShellError, PowerShellErrorKind, WincentError};
use wincent::prelude::*;

use wincent::destlist::{
    entries as dest_entries, experimental_remove_entries_by_rebuild,
    experimental_remove_entry_paths_by_rebuild, filetime_to_system_time,
    frequent_folders_dest_path, parse_bytes as parse_dest_bytes, parse_file as parse_dest_file,
    recent_files_dest_path, AutomaticDestinations, AutomaticDestinationsKind,
    ExperimentalRemoveOptions,
};

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
        "features" => cmd_features(),
        "list" => cmd_list(&manager, &args[1..]),
        "list-paths" => cmd_list_paths(&manager, &args[1..]),
        "check" => cmd_check(&manager, &args[1..]),
        "contains" => cmd_contains(&manager, &args[1..]),
        "add" => cmd_add(&manager, &args[1..]),
        "remove" => cmd_remove(&manager, &args[1..]),
        "batch-add" => cmd_batch_add(&manager, &args[1..]),
        "batch-remove" => cmd_batch_remove(&manager, &args[1..]),
        "lock" => cmd_lock(&manager, &args[1..]),
        "empty" => cmd_empty(&manager, &args[1..]),
        "retry" => cmd_retry(&args[1..]),
        "classify" => cmd_classify(&args[1..]),
        "invalid-path" => cmd_invalid_path(&args[1..]),
        "visible" => cmd_visible(&manager, &args[1..]),
        "dest" => cmd_dest(&manager, &args[1..]),
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
  cargo run --example main
  cargo run --example main -- [--timeout-ms N] <command> [args]

Interactive:
  help
  exit
  quit

Core:
  features
  list <recent|frequent|all> [--paths]
  list-paths <recent|frequent|all>
  check <recent|frequent|all> <path>
  contains <recent|frequent|all> <keyword>
  add <recent|frequent> <path> [--refresh]
  remove <recent|frequent> <path> [--deep-clean]
  batch-add [--refresh] <recent:path|frequent:path>...
  batch-remove [--deep-clean] <recent:path|frequent:path>...
  lock [recent|frequent|all] [--cleanup-new-links]
  empty <recent|frequent|all> [--pinned] [--refresh]
Utility APIs:
  retry <default|none|fast|standard|aggressive|custom> [--attempt N] [custom options]
  classify <stderr text>
  invalid-path <reason> [path]

Visibility APIs:
  visible get <recent|frequent|all>
  visible set <recent|frequent|all> <true|false>
  visible show <recent|frequent|all>
  visible hide <recent|frequent|all>
  visible get-recent | get-frequent
  visible set-recent <true|false> | set-frequent <true|false>

DestList APIs:
  dest path <recent|frequent>
  dest parse <recent|frequent|file> [path] [--limit N]
  dest parse-bytes <path> [--limit N]
  dest manager <recent|frequent> [--limit N]
  dest filetime <value>
  dest remove <recent|frequent> [--delay-ms N] <path>...
  dest remove-entries <recent|frequent> [--delay-ms N] <path>...
"#
    );
}

fn split_command_line(line: &str) -> WincentResult<Vec<String>> {
    let mut args = Vec::new();
    let mut current = String::new();
    let mut chars = line.chars().peekable();
    let mut quote = None;

    while let Some(ch) = chars.next() {
        match (quote, ch) {
            (Some(q), c) if c == q => quote = None,
            (Some(_), '\\') => {
                if let Some(next) = chars.next() {
                    current.push(next);
                } else {
                    current.push('\\');
                }
            }
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

fn cmd_features() -> WincentResult<()> {
    println!("visible: built-in");
    println!("destlist: built-in");
    Ok(())
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

fn cmd_check(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 2, "check <recent|frequent|all> <path>")?;
    let qa_type = parse_category(&args[0], true)?;
    let exists = manager.check_item_exact(&args[1], qa_type)?;
    println!("{exists}");
    Ok(())
}

fn cmd_contains(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 2, "contains <recent|frequent|all> <keyword>")?;
    let qa_type = parse_category(&args[0], true)?;
    let exists = manager.contains_item(&args[1], qa_type)?;
    println!("{exists}");
    Ok(())
}

fn cmd_add(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 2, "add <recent|frequent> <path> [--refresh]")?;
    let qa_type = parse_category(&args[0], false)?;
    let mut options = AddOptions::new();
    if args.iter().skip(2).any(|arg| arg == "--refresh") {
        options = options.force_recent_files_rebuild().refresh_explorer();
    }

    manager.add_item(&args[1], qa_type, options)?;
    println!("added {}", args[1]);
    Ok(())
}

fn cmd_remove(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 2, "remove <recent|frequent> <path> [--deep-clean]")?;
    let qa_type = parse_category(&args[0], false)?;
    let mut options = RemoveOptions::new();
    if args.iter().skip(2).any(|arg| arg == "--deep-clean") {
        options = options.deep_clean_recent_links();
    }

    manager.remove_item_with_options(&args[1], qa_type, options)?;
    println!("removed {}", args[1]);
    Ok(())
}

fn cmd_batch_add(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    let (refresh, item_args) = split_refresh(args);
    if item_args.is_empty() {
        return Err(WincentError::InvalidArgument(
            "batch-add requires at least one item".to_string(),
        ));
    }

    let items = parse_batch_items(&item_args)?;
    let options = if refresh {
        BatchOptions::new()
            .force_recent_files_rebuild()
            .refresh_explorer()
    } else {
        BatchOptions::new()
    };
    let result = manager.add_items_batch(&items, options);
    print_batch_result(result);
    Ok(())
}

fn cmd_batch_remove(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    let (refresh, refresh_args) = split_refresh(args);
    let (deep_clean, item_args) = split_deep_clean(&refresh_args);
    if item_args.is_empty() {
        return Err(WincentError::InvalidArgument(
            "batch-remove requires at least one item".to_string(),
        ));
    }

    let items = parse_batch_items(&item_args)?;
    let options = if deep_clean {
        RemoveOptions::new().deep_clean_recent_links()
    } else {
        RemoveOptions::new()
    };
    let batch_options = if refresh {
        BatchOptions::new().refresh_explorer()
    } else {
        BatchOptions::new()
    };
    let result = manager.remove_items_batch_with_batch_options(&items, batch_options, options);
    print_batch_result(result);
    Ok(())
}

fn cmd_lock(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    let cleanup = args.iter().any(|arg| arg == "--cleanup-new-links");
    let target = args
        .iter()
        .find(|arg| arg.as_str() != "--cleanup-new-links")
        .map(|arg| parse_lock_target(arg))
        .transpose()?
        .unwrap_or(QuickAccessLockTarget::All);
    let lock = match target {
        QuickAccessLockTarget::RecentFiles => manager.lock_recent_files()?,
        QuickAccessLockTarget::FrequentFolders => manager.lock_frequent_folders()?,
        QuickAccessLockTarget::All => manager.lock_quick_access()?,
        _ => {
            return Err(WincentError::InvalidArgument(
                "unsupported lock target".to_string(),
            ))
        }
    };
    println!("locked Quick Access backing files");
    println!("target: {:?}", lock.target());
    println!("recent_folder: {}", lock.recent_folder().display());
    println!("initial_lnk_paths: {}", lock.initial_lnk_paths().len());
    println!("press Enter to unlock");

    let mut line = String::new();
    io::stdin().read_line(&mut line).map_err(WincentError::Io)?;

    let options = if cleanup {
        QuickAccessUnlockOptions::new().cleanup_new_recent_links()
    } else {
        QuickAccessUnlockOptions::new()
    };
    let report = lock.unlock(options)?;
    println!("current_lnk_paths: {}", report.current_lnk_paths().len());
    println!("new_lnk_paths: {}", report.new_lnk_paths().len());
    println!("deleted_lnk_paths: {}", report.deleted_lnk_paths().len());
    println!(
        "failed_lnk_deletions: {}",
        report.failed_lnk_deletions().len()
    );
    for path in report.deleted_lnk_paths() {
        println!("  deleted {}", path.display());
    }
    for failure in report.failed_lnk_deletions() {
        println!("  failed {}: {}", failure.path().display(), failure.error());
    }
    Ok(())
}

fn parse_lock_target(value: &str) -> WincentResult<QuickAccessLockTarget> {
    match value {
        "recent" | "recent-files" | "files" => Ok(QuickAccessLockTarget::RecentFiles),
        "frequent" | "frequent-folders" | "folders" => Ok(QuickAccessLockTarget::FrequentFolders),
        "all" => Ok(QuickAccessLockTarget::All),
        other => Err(WincentError::InvalidArgument(format!(
            "unknown lock target: {other}"
        ))),
    }
}

fn cmd_empty(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(
        args,
        1,
        "empty <recent|frequent|all> [--pinned] [--refresh]",
    )?;
    let qa_type = parse_category(&args[0], true)?;
    let mut options = EmptyOptions::new();

    for arg in args.iter().skip(1) {
        match arg.as_str() {
            "--pinned" => options = options.remove_pinned_folders(),
            "--refresh" => options = options.refresh_explorer(),
            other => {
                return Err(WincentError::InvalidArgument(format!(
                    "unknown empty option: {other}"
                )))
            }
        }
    }

    manager.empty_items(qa_type, options)?;
    println!("cleared {}", category_name(qa_type));
    Ok(())
}

fn cmd_retry(args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "retry <policy> [--attempt N]")?;
    let mut attempt = 0;
    let mut policy = match args[0].as_str() {
        "default" | "standard" => RetryPolicy::standard(),
        "none" | "no-retry" => RetryPolicy::no_retry(),
        "fast" => RetryPolicy::fast(),
        "aggressive" => RetryPolicy::aggressive(),
        "custom" => RetryPolicy::new(),
        other => {
            return Err(WincentError::InvalidArgument(format!(
                "unknown retry policy: {other}"
            )))
        }
    };

    let mut index = 1;
    while index < args.len() {
        match args[index].as_str() {
            "--attempt" => {
                index += 1;
                attempt = parse_u32(args.get(index), "attempt")?;
            }
            "--max-attempts" => {
                index += 1;
                policy = policy.with_max_attempts(parse_u32(args.get(index), "max-attempts")?);
            }
            "--initial-ms" => {
                index += 1;
                policy = policy.with_initial_delay(Duration::from_millis(parse_u64(
                    required_arg(args.get(index), "initial-ms")?,
                    "initial-ms",
                )?));
            }
            "--max-ms" => {
                index += 1;
                policy = policy.with_max_delay(Duration::from_millis(parse_u64(
                    required_arg(args.get(index), "max-ms")?,
                    "max-ms",
                )?));
            }
            "--factor" => {
                index += 1;
                let factor = required_arg(args.get(index), "factor")?
                    .parse::<f64>()
                    .map_err(|_| WincentError::InvalidArgument("invalid factor".to_string()))?;
                policy = policy.with_backoff_factor(factor);
            }
            "--jitter" => {
                index += 1;
                policy = policy.with_jitter(parse_bool(required_arg(args.get(index), "jitter")?)?);
            }
            other => {
                return Err(WincentError::InvalidArgument(format!(
                    "unknown retry option: {other}"
                )))
            }
        }
        index += 1;
    }

    println!("max_attempts: {}", policy.max_attempts());
    println!("initial_delay_ms: {}", policy.initial_delay().as_millis());
    println!("max_delay_ms: {}", policy.max_delay().as_millis());
    println!("backoff_factor: {}", policy.backoff_factor());
    println!("jitter: {}", policy.jitter());
    println!(
        "delay_at_{attempt}_ms: {}",
        policy.calculate_delay(attempt).as_millis()
    );
    Ok(())
}

fn cmd_classify(args: &[String]) -> WincentResult<()> {
    if args.is_empty() {
        return Err(WincentError::InvalidArgument(
            "classify requires stderr text".to_string(),
        ));
    }

    let stderr = args.join(" ");
    let inferred = PowerShellError::infer_kind_from_stderr(&stderr);
    let classified = PowerShellError::classify_with(&stderr, Some(&custom_classifier));
    println!("infer_kind_from_stderr: {inferred:?}");
    println!("classify_with: {classified:?}");
    Ok(())
}

fn custom_classifier(stderr: &str) -> Option<PowerShellErrorKind> {
    let lower = stderr.to_lowercase();
    if lower.contains("access denied") || lower.contains("拒绝访问") {
        Some(PowerShellErrorKind::AccessDenied)
    } else if lower.contains("timed out") || lower.contains("timeout") {
        Some(PowerShellErrorKind::Timeout)
    } else {
        None
    }
}

fn cmd_invalid_path(args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "invalid-path <reason> [path]")?;
    let error = if let Some(path) = args.get(1) {
        InvalidPathError::new(path, &args[0])
    } else {
        InvalidPathError::reason(&args[0])
    };

    println!("display: {error}");
    println!("reason: {}", error.reason_text());
    match error.path() {
        Some(path) => println!("path: {}", path.display()),
        None => println!("path: <none>"),
    }
    Ok(())
}

fn cmd_visible(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "visible <command>")?;
    match args[0].as_str() {
        "get" => {
            require_len(&args[1..], 1, "visible get <recent|frequent|all>")?;
            let qa_type = parse_category(&args[1], true)?;
            println!("{}", manager.is_visible(qa_type)?);
        }
        "set" => {
            require_len(
                &args[1..],
                2,
                "visible set <recent|frequent|all> <true|false>",
            )?;
            let qa_type = parse_category(&args[1], true)?;
            manager.set_visible(qa_type, parse_bool(&args[2])?)?;
            println!("updated visibility");
        }
        "show" => {
            require_len(&args[1..], 1, "visible show <recent|frequent|all>")?;
            manager.show_section(parse_category(&args[1], true)?)?;
            println!("shown");
        }
        "hide" => {
            require_len(&args[1..], 1, "visible hide <recent|frequent|all>")?;
            manager.hide_section(parse_category(&args[1], true)?)?;
            println!("hidden");
        }
        "get-recent" => println!("{}", wincent::visible::is_recent_files_visible()?),
        "get-frequent" => println!("{}", wincent::visible::is_frequent_folders_visible()?),
        "set-recent" => {
            require_len(&args[1..], 1, "visible set-recent <true|false>")?;
            wincent::visible::set_recent_files_visible(parse_bool(&args[1])?)?;
            println!("updated recent visibility");
        }
        "set-frequent" => {
            require_len(&args[1..], 1, "visible set-frequent <true|false>")?;
            wincent::visible::set_frequent_folders_visible(parse_bool(&args[1])?)?;
            println!("updated frequent visibility");
        }
        other => {
            return Err(WincentError::InvalidArgument(format!(
                "unknown visible command: {other}"
            )))
        }
    }
    Ok(())
}

fn cmd_dest(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "dest <command>")?;
    match args[0].as_str() {
        "path" => {
            require_len(&args[1..], 1, "dest path <recent|frequent>")?;
            println!("{}", dest_path(parse_dest_kind(&args[1])?)?.display());
            Ok(())
        }
        "parse" => cmd_dest_parse(&args[1..]),
        "parse-bytes" => cmd_dest_parse_bytes(&args[1..]),
        "manager" => cmd_dest_manager(manager, &args[1..]),
        "filetime" => {
            require_len(&args[1..], 1, "dest filetime <value>")?;
            let value = parse_u64(&args[1], "filetime")?;
            match filetime_to_system_time(value) {
                Some(time) => println!("{time:?}"),
                None => println!("<before unix epoch or out of range>"),
            }
            Ok(())
        }
        "remove" => cmd_dest_remove(false, &args[1..]),
        "remove-entries" => cmd_dest_remove(true, &args[1..]),
        other => Err(WincentError::InvalidArgument(format!(
            "unknown dest command: {other}"
        ))),
    }
}

fn cmd_dest_parse(args: &[String]) -> WincentResult<()> {
    require_len(
        args,
        1,
        "dest parse <recent|frequent|file> [path] [--limit N]",
    )?;
    let limit = parse_limit(args, 20)?;
    let parsed = match args[0].as_str() {
        "recent" => parse_dest_file(recent_files_dest_path()?)?,
        "frequent" => parse_dest_file(frequent_folders_dest_path()?)?,
        "file" => {
            let path = args.get(1).ok_or_else(|| {
                WincentError::InvalidArgument("dest parse file requires a path".to_string())
            })?;
            parse_dest_file(path)?
        }
        other => {
            return Err(WincentError::InvalidArgument(format!(
                "unknown parse target: {other}"
            )))
        }
    };
    print_dest(&parsed, limit);
    Ok(())
}

fn cmd_dest_parse_bytes(args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "dest parse-bytes <path> [--limit N]")?;
    let limit = parse_limit(args, 20)?;
    let data = std::fs::read(&args[0]).map_err(WincentError::Io)?;
    let parsed = parse_dest_bytes(data)?;
    print_dest(&parsed, limit);
    Ok(())
}

fn cmd_dest_manager(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "dest manager <recent|frequent> [--limit N]")?;
    let limit = parse_limit(args, 20)?;
    let entries = match args[0].as_str() {
        "recent" => manager.get_recent_files_metadata()?,
        "frequent" => manager.get_frequent_folders_metadata()?,
        other => {
            return Err(WincentError::InvalidArgument(format!(
                "unknown metadata target: {other}"
            )))
        }
    };
    print_dest_entries(&entries, limit);
    Ok(())
}

fn cmd_dest_remove(use_entries: bool, args: &[String]) -> WincentResult<()> {
    require_len(
        args,
        2,
        "dest remove <recent|frequent> [--delay-ms N] <path>...",
    )?;
    let kind = parse_dest_kind(&args[0])?;
    let mut delay = Duration::from_millis(500);
    let mut paths = Vec::new();
    let mut index = 1;

    while index < args.len() {
        if args[index] == "--delay-ms" {
            index += 1;
            delay = Duration::from_millis(parse_u64(
                required_arg(args.get(index), "delay-ms")?,
                "delay-ms",
            )?);
        } else {
            paths.push(PathBuf::from(&args[index]));
        }
        index += 1;
    }

    if paths.is_empty() {
        return Err(WincentError::InvalidArgument(
            "dest remove requires at least one target path".to_string(),
        ));
    }

    let options = ExperimentalRemoveOptions::new().with_rebuild_delay(delay);
    let report = if use_entries {
        let parsed = parse_dest_file(dest_path(kind)?)?;
        let requested: Vec<String> = paths
            .iter()
            .map(|path| path.to_string_lossy().to_string())
            .collect();
        let matching: Vec<_> = parsed
            .dest_list()
            .entries()
            .iter()
            .filter(|entry| requested.iter().any(|path| entry.path() == path))
            .cloned()
            .collect();
        experimental_remove_entries_by_rebuild(kind, &matching, options)?
    } else {
        experimental_remove_entry_paths_by_rebuild(kind, &paths, options)?
    };

    print_remove_report(&report);
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

fn parse_dest_kind(value: &str) -> WincentResult<AutomaticDestinationsKind> {
    match value {
        "recent" | "recent-files" | "files" => Ok(AutomaticDestinationsKind::RecentFiles),
        "frequent" | "frequent-folders" | "folders" => {
            Ok(AutomaticDestinationsKind::FrequentFolders)
        }
        other => Err(WincentError::InvalidArgument(format!(
            "unknown AutomaticDestinations kind: {other}"
        ))),
    }
}

fn dest_path(kind: AutomaticDestinationsKind) -> WincentResult<PathBuf> {
    match kind {
        AutomaticDestinationsKind::RecentFiles => recent_files_dest_path(),
        AutomaticDestinationsKind::FrequentFolders => frequent_folders_dest_path(),
        _ => Err(WincentError::UnsupportedOperation(
            "unknown AutomaticDestinations kind".to_string(),
        )),
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

fn parse_batch_items(args: &[String]) -> WincentResult<Vec<QuickAccessItem>> {
    args.iter()
        .map(|item| {
            let (kind, path) = item.split_once(':').ok_or_else(|| {
                WincentError::InvalidArgument(format!(
                    "batch item must be recent:path or frequent:path: {item}"
                ))
            })?;
            match kind {
                "recent" | "file" | "files" => Ok(QuickAccessItem::recent_file(path)),
                "frequent" | "folder" | "folders" => Ok(QuickAccessItem::frequent_folder(path)),
                other => Err(WincentError::InvalidArgument(format!(
                    "unknown batch item type: {other}"
                ))),
            }
        })
        .collect()
}

fn split_refresh(args: &[String]) -> (bool, Vec<String>) {
    let mut refresh = false;
    let mut rest = Vec::new();
    for arg in args {
        if arg == "--refresh" {
            refresh = true;
        } else {
            rest.push(arg.clone());
        }
    }
    (refresh, rest)
}

fn split_deep_clean(args: &[String]) -> (bool, Vec<String>) {
    let mut deep_clean = false;
    let mut rest = Vec::new();
    for arg in args {
        if arg == "--deep-clean" {
            deep_clean = true;
        } else {
            rest.push(arg.clone());
        }
    }
    (deep_clean, rest)
}

fn print_batch_result(result: BatchResult) {
    println!("total: {}", result.total());
    println!("succeeded: {}", result.succeeded().len());
    println!("failed: {}", result.failed().len());
    println!("success_rate: {:.1}%", result.success_rate() * 100.0);
    println!("complete_success: {}", result.is_complete_success());
    println!("partial_success: {}", result.has_partial_success());

    for path in result.succeeded() {
        println!("ok: {path}");
    }
    for failure in result.failed() {
        println!("failed: {}: {}", failure.path(), failure.error());
    }
}

fn print_dest(parsed: &AutomaticDestinations, limit: usize) {
    let cfb = parsed.cfb_info();
    let dest = parsed.dest_list();

    println!("cfb.sector_size: {}", cfb.sector_size());
    println!("cfb.mini_sector_size: {}", cfb.mini_sector_size());
    println!("cfb.mini_cutoff_size: {}", cfb.mini_cutoff_size());
    println!("cfb.directory_entries: {}", cfb.directory_entries().len());
    for entry in cfb.directory_entries().iter().take(limit) {
        println!(
            "cfb.entry name={} type={} start_sector={} stream_size={}",
            entry.name(),
            entry.object_type(),
            entry.start_sector(),
            entry.stream_size()
        );
    }

    println!("dest.version: {}", dest.version());
    println!("dest.declared_entry_count: {}", dest.declared_entry_count());
    println!("dest.pinned_entry_count: {}", dest.pinned_entry_count());
    println!("dest.header_counter_raw: {}", dest.header_counter_raw());
    println!("dest.header_counter_f32: {}", dest.header_counter_f32());
    println!("dest.last_entry_id: {}", dest.last_entry_id());
    println!("dest.last_entry_number: {}", dest.last_entry_number());
    println!(
        "dest.add_delete_action_count: {}",
        dest.add_delete_action_count()
    );
    println!("dest.diagnostics: {}", dest.diagnostics().len());
    print_dest_entries(&dest_entries(dest), limit);
}

fn print_dest_entries(entries: &[wincent::destlist::DestListEntry], limit: usize) {
    println!("entries: {}", entries.len());
    for entry in entries.iter().take(limit) {
        println!(
            concat!(
                "entry offset={} len={} mru={} checksum={} id={} number={} number_unknown={} ",
                "host={} stream={} pinned={} pin_status={} pin_order={:?} rank={} recent_rank={} ",
                "count={} access_count={} score={} last_access={:?} ",
                "last_interaction={:?} sps_size={:?} reserved_78={:?} reserved_7c={:?} ",
                "droid={} mac={} warnings={} raw_path={} path={}"
            ),
            entry.entry_offset(),
            entry.entry_len(),
            entry.mru_position(),
            entry.checksum(),
            entry.entry_id(),
            entry.entry_number(),
            entry.entry_number_unknown(),
            entry.hostname(),
            entry.stream_name(),
            entry.is_pinned(),
            entry.pin_status(),
            entry.pin_order(),
            entry.rank(),
            entry.recent_rank(),
            entry.count(),
            entry.access_count(),
            entry.score(),
            entry.last_access_filetime(),
            entry.last_interaction_filetime(),
            entry.sps_size(),
            entry.reserved_78(),
            entry.reserved_7c(),
            entry.file_droid(),
            entry.file_droid_mac(),
            entry.warnings().len(),
            entry.raw_path(),
            entry.path()
        );
    }
}

fn print_remove_report(report: &wincent::destlist::ExperimentalRemoveReport) {
    println!("kind: {:?}", report.kind());
    println!("recent_folder: {}", report.recent_folder().display());
    println!("dest_path: {}", report.dest_path().display());
    print_strings("requested_paths", report.requested_paths());
    print_strings("matching_paths_before", report.matching_paths_before());
    println!("deleted_lnk_paths: {}", report.deleted_lnk_paths().len());
    for path in report.deleted_lnk_paths() {
        println!("{}", path.display());
    }
    print_strings(
        "missing_lnk_target_paths",
        report.missing_lnk_target_paths(),
    );
    println!("dest_deleted: {}", report.dest_deleted());
    println!("rebuilt: {}", report.rebuilt());
    println!(
        "rebuild_parse_elapsed: {:?}",
        report.rebuild_parse_elapsed()
    );
    println!("rebuild_parse_error: {:?}", report.rebuild_parse_error());
    print_strings(
        "remaining_paths_after_rebuild",
        report.remaining_paths_after_rebuild(),
    );
    println!("success: {}", report.success());
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
        WincentError::PowerShellExecution(error) => {
            eprintln!("powershell.kind: {:?}", error.kind());
            eprintln!("powershell.operation: {:?}", error.operation());
            eprintln!("powershell.exit_code: {:?}", error.exit_code());
            eprintln!("powershell.script_path: {}", error.script_path().display());
            eprintln!("powershell.parameters: {:?}", error.parameters());
            eprintln!("powershell.duration: {:?}", error.duration());
            eprintln!("powershell.io_error: {:?}", error.io_error());
            eprintln!("powershell.os_error: {:?}", error.os_error());
            eprintln!("powershell.is_access_denied: {}", error.is_access_denied());
            eprintln!(
                "powershell.is_execution_policy_error: {}",
                error.is_execution_policy_error()
            );
            eprintln!("powershell.is_timeout: {}", error.is_timeout());
            eprintln!(
                "powershell.is_cmdlet_not_found: {}",
                error.is_cmdlet_not_found()
            );
            eprintln!("powershell.is_transient: {}", error.is_transient());
            eprintln!("powershell.suggest_fix: {:?}", error.suggest_fix());
            eprintln!(
                "powershell.stderr_has_error: {}",
                error.stderr_contains("error")
            );
            eprintln!("powershell.stdout: {}", error.raw_stdout());
            eprintln!("powershell.stderr: {}", error.raw_stderr());
            eprintln!(
                "powershell.classification_with: {:?}",
                error.classification_with(custom_classifier)
            );
        }
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

fn parse_limit(args: &[String], default: usize) -> WincentResult<usize> {
    let mut limit = default;
    let mut index = 0;
    while index < args.len() {
        if args[index] == "--limit" {
            index += 1;
            let value = required_arg(args.get(index), "limit")?;
            limit = value
                .parse::<usize>()
                .map_err(|_| WincentError::InvalidArgument("invalid limit".to_string()))?;
        }
        index += 1;
    }
    Ok(limit)
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

fn parse_u32(value: Option<&String>, name: &str) -> WincentResult<u32> {
    required_arg(value, name)?
        .parse::<u32>()
        .map_err(|_| WincentError::InvalidArgument(format!("invalid {name}")))
}

fn parse_u64(value: &str, name: &str) -> WincentResult<u64> {
    value
        .parse::<u64>()
        .map_err(|_| WincentError::InvalidArgument(format!("invalid {name}")))
}

fn required_arg<'a>(value: Option<&'a String>, name: &str) -> WincentResult<&'a str> {
    value
        .map(String::as_str)
        .ok_or_else(|| WincentError::InvalidArgument(format!("{name} requires a value")))
}

fn require_len(args: &[String], min: usize, usage: &str) -> WincentResult<()> {
    if args.len() < min {
        Err(WincentError::InvalidArgument(format!("usage: {usage}")))
    } else {
        Ok(())
    }
}
