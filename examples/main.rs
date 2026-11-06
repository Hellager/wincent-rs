use std::env;
use std::io::{self, Write};
use std::path::Path;
use std::path::PathBuf;
use std::time::{Duration, Instant};

use wincent::error::{InvalidPathError, WincentError};
use wincent::prelude::*;

use wincent::destlist::{
    entries as dest_entries, filetime_to_system_time, frequent_folders_dest_path,
    parse_bytes as parse_dest_bytes, parse_file as parse_dest_file,
    parse_file_with_kind as parse_dest_file_with_kind, recent_files_dest_path,
    AutomaticDestinations, DestListKind,
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
    if args[0] == "clear" {
        return cmd_clear();
    }

    let manager = QuickAccessManager::builder().try_timeout(timeout)?.build();

    match args[0].as_str() {
        "features" => cmd_features(),
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
        "invalid-path" => cmd_invalid_path(&args[1..]),
        "visible" => cmd_visible(&manager, &args[1..]),
        "dest" => cmd_dest(&manager, &args[1..]),
        "bench" => cmd_bench(&manager, &args[1..]),
        other => Err(WincentError::InvalidArgument(format!(
            "unknown command: {other}"
        ))),
    }
}

fn cmd_clear() -> WincentResult<()> {
    print!("\x1B[2J\x1B[H");
    io::stdout().flush().map_err(WincentError::Io)
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
  clear
  exit
  quit

Core:
  features
  list <recent|frequent|all> [--paths]
  list-paths <recent|frequent|all>
  contains-path <recent|frequent|all> <path>
  status <path>
    prints where the DestList rules place the path in a default Shell query:
    recent position, frequent position and pin state, or none.
  contains-keyword <recent|frequent|all> <keyword>
  find <recent|frequent|all> <keyword>
  add <path>
  add-timed <path>
    prints successful addition time as one JSON object.
  remove <path>
    detects Recent files and pinned or unpinned Frequent Folders automatically.
  refresh
    notifies the Shell, then refreshes open Explorer views.
  lock [recent|frequent|all]
    Locked targets block Explorer from adding or removing matching Quick Access items.
  empty <recent|frequent|all>
Utility APIs:
  invalid-path <reason> [path]

Visibility APIs:
  visible get <recent|frequent|all>
  visible set <recent|frequent|all> <true|false>
  visible get-start-recommended
  visible set-start-recommended <true|false>
  Start Recommended commands control recent files in Windows Recommended items.
  Successful writes do not guarantee the UI changes immediately.

DestList APIs:
  dest path <recent|frequent>
  dest reveal <recent|frequent>
  dest parse <recent|frequent|file> [path] [--limit N]
  dest cfb <recent|frequent|file> [path] [recent|frequent]
  dest rules <recent|frequent|file> [path] [recent|frequent]
  dest parse-bytes <path> [--limit N]
  dest manager <recent|frequent> [--limit N]
  dest filetime <value>
  bench <recent|frequent> [--limit N]
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
    if args.len() > 1 {
        return Err(WincentError::InvalidArgument(
            "usage: lock [recent|frequent|all]".to_string(),
        ));
    }
    let target = args
        .first()
        .map(|arg| parse_category(arg, true))
        .transpose()?
        .unwrap_or(QuickAccess::All);
    manager.lock(target)?;
    println!("locked Quick Access backing files");
    println!("target: {target:?}");
    println!("press Enter to unlock");

    let mut line = String::new();
    io::stdin().read_line(&mut line).map_err(WincentError::Io)?;

    let new_links = manager.unlock()?;
    println!("new_lnk_paths: {}", new_links.len());
    for path in new_links {
        println!("  {path}");
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
            require_exact_len(&args[1..], 1, "visible get <recent|frequent|all>")?;
            let qa_type = parse_category(&args[1], true)?;
            println!("{}", manager.is_visible(qa_type)?);
        }
        "set" => {
            require_exact_len(
                &args[1..],
                2,
                "visible set <recent|frequent|all> <true|false>",
            )?;
            let qa_type = parse_category(&args[1], true)?;
            manager.set_visible(qa_type, parse_bool(&args[2])?)?;
            println!("updated visibility");
        }
        "get-start-recommended" => {
            require_exact_len(&args[1..], 0, "visible get-start-recommended")?;
            println!("{}", manager.is_start_recommended_section_visible()?);
        }
        "set-start-recommended" => {
            require_exact_len(&args[1..], 1, "visible set-start-recommended <true|false>")?;
            manager.set_start_recommended_section_visible(parse_bool(&args[1])?)?;
            println!("updated recommended recent files visibility");
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
        "reveal" => {
            require_len(&args[1..], 1, "dest reveal <recent|frequent>")?;
            cmd_dest_reveal(parse_dest_kind(&args[1])?)
        }
        "parse" => cmd_dest_parse(&args[1..]),
        "cfb" => cmd_dest_cfb(&args[1..]),
        "rules" => cmd_dest_rules(&args[1..]),
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
        other => Err(WincentError::InvalidArgument(format!(
            "unknown dest command: {other}"
        ))),
    }
}

fn cmd_dest_reveal(kind: DestKind) -> WincentResult<()> {
    let path = dest_path(kind)?;
    println!("revealing {}", path.display());
    reveal_in_explorer(&path)
}

fn reveal_in_explorer(path: &Path) -> WincentResult<()> {
    if !path.exists() {
        return Err(WincentError::InvalidPath(InvalidPathError::new(
            path,
            "Path does not exist",
        )));
    }

    let select = format!("/select,{}", path.display());
    let mut child = std::process::Command::new("explorer.exe")
        .arg(select)
        .spawn()
        .map_err(|error| {
            WincentError::SystemError(format!("failed to launch explorer: {error}"))
        })?;

    // explorer.exe exits with status 1 even when /select succeeds, so exit
    // status is unreliable; only a spawn failure counts as an error.
    let _ = child.wait();
    println!("opened Explorer at the dest path's folder with the file selected");
    Ok(())
}

fn cmd_dest_parse(args: &[String]) -> WincentResult<()> {
    require_len(
        args,
        1,
        "dest parse <recent|frequent|file> [path] [--limit N]",
    )?;
    let (parsed, limit) = match args[0].as_str() {
        "recent" => (
            parse_dest_file(recent_files_dest_path()?)?,
            parse_limit(&args[1..], 20)?,
        ),
        "frequent" => (
            parse_dest_file(frequent_folders_dest_path()?)?,
            parse_limit(&args[1..], 20)?,
        ),
        "file" => {
            let path = args.get(1).ok_or_else(|| {
                WincentError::InvalidArgument("dest parse file requires a path".to_string())
            })?;
            let trailing = if args.len() > 2 { &args[2..] } else { &[] };
            (parse_dest_file(path)?, parse_limit(trailing, 20)?)
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

fn cmd_dest_cfb(args: &[String]) -> WincentResult<()> {
    let parsed = parse_dest_target(args, "cfb")?;
    let entries = parsed.cfb_entries();
    println!("cfb_entries: {}", entries.len());
    for entry in entries {
        println!("{}", entry.path());
    }
    Ok(())
}

fn cmd_dest_rules(args: &[String]) -> WincentResult<()> {
    let parsed = parse_dest_target(args, "rules")?;
    let entries = parsed.rule_entries();
    println!("rule_entries: {}", entries.len());
    for entry in entries {
        println!("{}", entry.path());
    }
    Ok(())
}

fn parse_dest_target(args: &[String], command: &str) -> WincentResult<AutomaticDestinations> {
    require_len(
        args,
        1,
        &format!("dest {command} <recent|frequent|file> [path] [recent|frequent]"),
    )?;
    let parsed = match args[0].as_str() {
        "recent" => parse_dest_file(recent_files_dest_path()?)?,
        "frequent" => parse_dest_file(frequent_folders_dest_path()?)?,
        "file" => {
            let path = args.get(1).ok_or_else(|| {
                WincentError::InvalidArgument(format!("dest {command} file requires a path"))
            })?;
            match args.get(2).map(String::as_str) {
                Some("recent") => parse_dest_file_with_kind(path, DestListKind::RecentFiles)?,
                Some("frequent") => parse_dest_file_with_kind(path, DestListKind::FrequentFolders)?,
                Some(other) => {
                    return Err(WincentError::InvalidArgument(format!(
                        "unknown DestList kind for {command} file: {other}"
                    )))
                }
                None => parse_dest_file(path)?,
            }
        }
        other => {
            return Err(WincentError::InvalidArgument(format!(
                "unknown {command} target: {other}"
            )))
        }
    };
    Ok(parsed)
}

fn cmd_dest_parse_bytes(args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "dest parse-bytes <path> [--limit N]")?;
    let limit = parse_limit(&args[1..], 20)?;
    let data = std::fs::read(&args[0]).map_err(WincentError::Io)?;
    let parsed = parse_dest_bytes(data)?;
    print_dest(&parsed, limit);
    Ok(())
}

fn cmd_dest_manager(_manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "dest manager <recent|frequent> [--limit N]")?;
    let limit = parse_limit(&args[1..], 20)?;
    let entries = match args[0].as_str() {
        "recent" => read_metadata(QuickAccess::RecentFiles)?,
        "frequent" => read_metadata(QuickAccess::FrequentFolders)?,
        other => {
            return Err(WincentError::InvalidArgument(format!(
                "unknown metadata target: {other}"
            )))
        }
    };
    print_dest_entries(&entries, limit);
    Ok(())
}

fn cmd_bench(manager: &QuickAccessManager, args: &[String]) -> WincentResult<()> {
    require_len(args, 1, "bench <recent|frequent> [--limit N]")?;
    let category = parse_category(&args[0], false)?;
    let iterations = parse_limit(&args[1..], 100)?;
    if iterations == 0 {
        return Err(WincentError::InvalidArgument(
            "bench limit must be greater than zero".to_string(),
        ));
    }

    // Warm-up pass keeps one-time costs (COM init, file cache) out of the loop.
    let _ = manager.get_items(category)?;
    let _ = read_metadata(category)?;

    let mut shell_ms = Vec::with_capacity(iterations);
    let mut rules_ms = Vec::with_capacity(iterations);
    for _ in 0..iterations {
        let start = Instant::now();
        let _ = manager.get_items(category)?;
        shell_ms.push(start.elapsed().as_secs_f64() * 1000.0);

        let start = Instant::now();
        let _ = read_metadata(category)?;
        rules_ms.push(start.elapsed().as_secs_f64() * 1000.0);
    }

    print_bench_result("native shell query", &shell_ms);
    print_bench_result("cfb rule parsing", &rules_ms);
    println!("iterations: {iterations}");
    Ok(())
}

fn print_bench_result(label: &str, samples_ms: &[f64]) {
    let mut sorted = samples_ms.to_vec();
    sorted.sort_by(f64::total_cmp);
    let median = sorted[sorted.len() / 2];
    let mean = sorted.iter().sum::<f64>() / sorted.len() as f64;
    println!(
        "{label}: median {median:.3} ms, mean {mean:.3} ms, min {:.3} ms, max {:.3} ms",
        sorted[0],
        sorted[sorted.len() - 1]
    );
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

#[derive(Debug, Clone, Copy)]
enum DestKind {
    RecentFiles,
    FrequentFolders,
}

fn parse_dest_kind(value: &str) -> WincentResult<DestKind> {
    match value {
        "recent" | "recent-files" | "files" => Ok(DestKind::RecentFiles),
        "frequent" | "frequent-folders" | "folders" => Ok(DestKind::FrequentFolders),
        other => Err(WincentError::InvalidArgument(format!(
            "unknown AutomaticDestinations kind: {other}"
        ))),
    }
}

fn dest_path(kind: DestKind) -> WincentResult<PathBuf> {
    match kind {
        DestKind::RecentFiles => recent_files_dest_path(),
        DestKind::FrequentFolders => frequent_folders_dest_path(),
    }
}

fn read_metadata(qa_type: QuickAccess) -> WincentResult<Vec<wincent::destlist::DestListEntry>> {
    let path = match qa_type {
        QuickAccess::RecentFiles => recent_files_dest_path()?,
        QuickAccess::FrequentFolders => frequent_folders_dest_path()?,
        _ => {
            return Err(WincentError::UnsupportedOperation(
                "metadata requires Recent Files or Frequent Folders".to_string(),
            ));
        }
    };
    let parsed = parse_dest_file(path)?;
    Ok(dest_entries(parsed.dest_list()))
}

fn category_name(qa_type: QuickAccess) -> &'static str {
    match qa_type {
        QuickAccess::RecentFiles => "Recent Files",
        QuickAccess::FrequentFolders => "Frequent Folders",
        QuickAccess::All => "All",
        _ => "Unknown",
    }
}

fn print_dest(parsed: &AutomaticDestinations, limit: usize) {
    let cfb = parsed.cfb_info();
    let dest = parsed.dest_list();

    println!("cfb.major_version: {}", cfb.major_version());
    println!("cfb.sector_size: {}", cfb.sector_size());
    println!("cfb.mini_sector_size: {}", cfb.mini_sector_size());
    println!("cfb.mini_cutoff_size: {}", cfb.mini_cutoff_size());
    println!("cfb.directory_entries: {}", cfb.directory_entries().len());
    for entry in cfb.directory_entries().iter().take(limit) {
        println!(
            "cfb.entry name={} path={} type={} start_sector={:?} stream_size={}",
            entry.name(),
            entry.path().display(),
            entry.object_type(),
            entry.start_sector(),
            entry.stream_size()
        );
    }
    println!("cfb.streams: {}", cfb.streams().len());
    for stream in cfb.streams().iter().take(limit) {
        println!(
            "cfb.stream name={} path={} size={}",
            stream.name(),
            stream.path().display(),
            stream.len()
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
    println!("dest.header_reserved_1c: {}", dest.header_reserved_1c());
    match parsed.dest_list_property_store() {
        Some(stream) => println!(
            "dest.property_store: size={} value={:?}",
            stream.len(),
            parsed.dest_list_property_store_value()
        ),
        None => println!("dest.property_store: <absent>"),
    }
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

fn parse_limit(args: &[String], default: usize) -> WincentResult<usize> {
    let mut limit = default;
    let mut index = 0;
    while index < args.len() {
        match args[index].as_str() {
            "--limit" => {
                index += 1;
                let value = required_arg(args.get(index), "limit")?;
                limit = value
                    .parse::<usize>()
                    .map_err(|_| WincentError::InvalidArgument("invalid limit".to_string()))?;
            }
            other => {
                return Err(WincentError::InvalidArgument(format!(
                    "unknown limit option: {other}"
                )))
            }
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
    fn parse_limit_rejects_unknown_options() {
        let error = parse_limit(&["--limti".to_string(), "3".to_string()], 20).unwrap_err();

        assert!(
            matches!(error, WincentError::InvalidArgument(message) if message == "unknown limit option: --limti")
        );
    }

    #[test]
    fn parse_limit_accepts_limit_after_file_path_is_removed_by_caller() {
        let limit = parse_limit(&["--limit".to_string(), "3".to_string()], 20).unwrap();

        assert_eq!(limit, 3);
    }

    #[test]
    fn visible_set_start_recommended_requires_bool_argument() {
        let error = require_exact_len(
            &[] as &[String],
            1,
            "visible set-start-recommended <true|false>",
        )
        .unwrap_err();

        assert!(
            matches!(error, WincentError::InvalidArgument(message) if message == "usage: visible set-start-recommended <true|false>")
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
            "visible set <recent|frequent|all> <true|false>",
        )
        .unwrap_err();

        assert!(matches!(error, WincentError::InvalidArgument(_)));
    }
}
