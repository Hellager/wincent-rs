# wincent-rs

[![](https://img.shields.io/crates/v/wincent.svg)](https://crates.io/crates/wincent)
[![][img_doc]][doc]
![Crates.io Total Downloads](https://img.shields.io/crates/d/wincent)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Hellager/wincent-rs/publish.yml)
![Crates.io License](https://img.shields.io/crates/l/wincent)

其他语言版本：[English](README.md) | [中文](README.cn.md)

## 概述

Wincent 是一个用于管理 Windows 快速访问功能的 Rust 库，提供对最近使用文件和常用文件夹的安全操作接口，底层采用原生 Windows API 优先、PowerShell 兜底的双路策略。

## 功能特性

- 查询最近文件和常用文件夹
- 添加和删除项目，自动检测重复项
- 清空分类（可选同时移除已固定文件夹）
- 精确路径或关键词检查项目是否存在
- 批量添加/删除，逐项收集错误
- 所有 Shell 操作均有超时保护
- 可选的快速访问分区可见性控制（`visible` feature）
- 可选的 DestList 元数据访问（`destlist` feature）

## 安装

在 `Cargo.toml` 中添加以下依赖：

```toml
[dependencies]
wincent = "0.2.2"
```

启用可选 feature：

```toml
[dependencies]
wincent = { version = "0.2.2", features = ["visible", "destlist"] }
```

## 快速开始

```rust,no_run
use wincent::prelude::*;

fn main() -> WincentResult<()> {
    let manager = QuickAccessManager::new();

    // --- 添加 ---
    // 添加文件到最近文件。若已存在则返回 Err(AlreadyExists)。
    manager.add_item(
        "C:\\Projects\\report.docx",
        QuickAccess::RecentFiles,
        AddOptions::new().refresh_recent_files(),
    )?;

    // 固定文件夹到常用文件夹。
    manager.add_item("C:\\Projects", QuickAccess::FrequentFolders, AddOptions::new())?;

    // --- 查询 ---
    // 以 PathBuf 形式列出所有最近文件。
    let recent = manager.get_item_paths(QuickAccess::RecentFiles)?;
    println!("最近文件（{}个）：", recent.len());
    for path in &recent {
        println!("  {}", path.display());
    }

    // 精确成员检查（Windows 路径语义：不区分大小写）。
    let exists = manager.check_item_exact("C:\\Projects\\report.docx", QuickAccess::RecentFiles)?;
    println!("report.docx 在最近文件中：{exists}");

    // 模糊检查：任意项目的路径字符串包含指定关键词。
    let any_match = manager.contains_item("Projects", QuickAccess::All)?;
    println!("快速访问中存在包含 'Projects' 的项目：{any_match}");

    // --- 删除 ---
    // 移除文件。若不存在则返回 Err(NotInQuickAccess)。
    manager.remove_item("C:\\Projects\\report.docx", QuickAccess::RecentFiles)?;

    Ok(())
}
```

## Manager API

### 构造

| 方法 | 说明 |
|------|------|
| `QuickAccessManager::new()` | 使用默认 10 秒超时创建管理器。 |
| `QuickAccessManager::builder()` | 返回 `QuickAccessManagerBuilder` 以进行自定义配置。 |
| `builder.timeout(duration)` | 设置 Shell 操作超时时间（为零时 panic）。 |
| `builder.try_timeout(duration)` | 设置超时时间，为零时返回 `Err` 而非 panic。 |
| `builder.build()` | 构建配置好的 `QuickAccessManager`。 |

### 查询

| 方法 | 说明 |
|------|------|
| `get_items(qa_type)` | 返回指定分类（`RecentFiles`、`FrequentFolders` 或 `All`）的所有路径。 |
| `get_item_paths(qa_type)` | 以 `PathBuf` 形式返回所有路径。路径反映 Explorer 状态，可能已不再存在于磁盘上。 |
| `check_item_exact(path, qa_type)` | 若路径存在则返回 `true`（Windows 不区分大小写比较）。 |
| `contains_item(keyword, qa_type)` | 若任意项目的路径字符串包含 `keyword` 则返回 `true`（区分大小写的子串匹配）。 |

### 修改

| 方法 | 说明 |
|------|------|
| `add_item(path, qa_type, options)` | 添加项目（接受 `impl AsRef<Path>`）。不接受 `QuickAccess::All`。若已存在则返回 `Err(AlreadyExists)`。 |
| `remove_item(path, qa_type)` | 删除项目（接受 `impl AsRef<Path>`）。不接受 `QuickAccess::All`。若不存在则返回 `Err(NotInQuickAccess)`。 |
| `add_items_batch(items, options)` | 批量添加多个 `QuickAccessItem`，遇错不中断，逐项收集失败结果。 |
| `remove_items_batch(items)` | 批量删除多个项目，遇错不中断，逐项收集失败结果。 |
| `empty_items(qa_type, options)` | 清空指定分类。使用 `EmptyOptions::remove_pinned_folders()` 可同时移除已固定文件夹。 |

### 可选 Feature 方法

| 方法 | Feature | 说明 |
|------|---------|------|
| `is_visible(qa_type)` | `visible` | 读取控制分区显示状态的 Explorer 注册表标志。 |
| `set_visible(qa_type, visible)` | `visible` | 写入 Explorer 注册表标志。 |
| `show_section(qa_type)` / `hide_section(qa_type)` | `visible` | 显示或隐藏分区，避免直接传入裸布尔值。 |
| `get_recent_files_metadata()` | `destlist` | 解析最近文件 `.automaticDestinations-ms` 文件并返回 `DestListEntry` 记录。 |
| `get_frequent_folders_metadata()` | `destlist` | 解析常用文件夹 `.automaticDestinations-ms` 文件并返回 `DestListEntry` 记录。 |

## 错误处理

所有可能失败的方法均返回 `WincentResult<T>`（即 `Result<T, WincentError>`）。常见错误变体：

| 变体 | 触发时机 |
|------|----------|
| `AlreadyExists { path, qa_type }` | `add_item` 目标路径已在快速访问分类中存在。 |
| `NotInQuickAccess { path, qa_type }` | `remove_item` 目标路径不在快速访问分类中。 |
| `InvalidPath(error)` | 提供的路径为空、格式错误、缺失或类型不符。 |
| `UnsupportedOperation(msg)` | `add_item`/`remove_item` 使用了 `QuickAccess::All`。 |
| `Timeout(msg)` | Shell 操作超过配置的超时时间。 |
| `PartialEmpty { … }` | `empty_items` 在清空部分分类后，后续步骤失败。 |
| `PowerShellExecution(err)` | PowerShell 兜底脚本执行失败或无法启动。 |
| `SystemError(msg)` | 非 I/O 的 Windows API 调用失败。 |

### PowerShell 错误处理示例

```rust,no_run
use wincent::prelude::*;

fn main() -> WincentResult<()> {
    let manager = QuickAccessManager::new();

    match manager.add_item("C:\\folder", QuickAccess::FrequentFolders, AddOptions::new()) {
        Ok(()) => println!("添加成功。"),
        Err(WincentError::AlreadyExists { path, .. }) => println!("{path} 已固定。"),
        Err(WincentError::PowerShellExecution(err)) => {
            if err.is_access_denied() {
                println!("权限不足，请尝试以管理员身份运行。");
            } else if err.is_execution_policy_error() {
                println!("PowerShell 执行策略阻止了脚本运行。");
            } else if let Some(fix) = err.suggest_fix() {
                println!("建议：{fix}");
            } else {
                println!("脚本错误：{err}");
            }
        }
        Err(e) => println!("错误：{e}"),
    }

    Ok(())
}
```

## 系统要求与限制

- **操作系统**：Windows 10 或 Windows 11。
- **Rust**：1.60.0 或更高版本。
- **PowerShell**：原生 Windows API 路径不可用时，作为兜底方案使用。脚本以进程级 `-ExecutionPolicy Bypass` 启动，不修改用户或机器策略，但在加固环境下仍可能被拦截。
- **权限**：大多数操作在当前用户账户下运行，无需提升权限。固定/取消固定文件夹操作需要活跃的桌面会话。
- **状态一致性**：快速访问状态由 Windows Explorer 维护，修改操作后结果可能有短暂延迟，Explorer 也可能在不同版本间以异步方式重建状态。

## 注意事项

- 核心功能高度依赖系统 API，Windows 安全策略收紧时相关调用可能受限或失效。
- 个人系统环境可能与测试环境存在差异（例如第三方软件修改了相关注册表项），建议在使用前通过查询函数验证功能是否可用，尤其是文件夹相关操作。
- `visible` feature 会修改注册表，可能影响 Explorer 窗口布局，请谨慎使用。

## 最佳实践

- 添加最近文件时使用 `AddOptions::refresh_recent_files()` 以立即在 Explorer 中显示。
- 仅在确实需要移除用户固定条目时，才使用 `EmptyOptions::remove_pinned_folders()`。
- 若代码需要区分"未清空任何内容"与"部分清空"，应显式处理 `PartialEmpty`。
- 成员检查优先使用 `check_item_exact`（Windows 路径语义），而非 `contains_item`（区分大小写的子串匹配）。

## 贡献指南

1. Fork 本仓库
2. 创建功能分支 (`git checkout -b wincent/amazing-feature`)
3. 提交更改 (`git commit -m 'feat: 添加某个很棒的功能'`)
4. 推送到分支 (`git push origin wincent/amazing-feature`)
5. 开启一个 Pull Request

### 开发环境设置

```bash
# 克隆仓库
git clone https://github.com/Hellager/wincent-rs.git
cd wincent-rs

# 构建并运行单元测试（无需 Explorer 桌面会话）
cargo build
cargo test

# 运行需要交互式桌面会话的集成测试
cargo test -- --ignored
```

## 免责声明

本库与系统级快速访问功能进行交互。在进行重要更改之前，请确保您具有适当的权限并创建备份。

## 支持

如果您遇到任何问题或有疑问，请在 GitHub 仓库上提出 issue。

## 致谢

- [Castorix31](https://learn.microsoft.com/en-us/answers/questions/1087928/how-to-get-recent-docs-list-and-delete-some-of-the)
- [Yohan Ney](https://stackoverflow.com/questions/30051634/is-it-possible-programmatically-add-folders-to-the-windows-10-quick-access-panel)
- [libyal](https://github.com/libyal/dtformats/blob/main/documentation/Jump%20lists%20format.asciidoc)
- [Eric Zimmerman](https://github.com/EricZimmerman/JumpList)
- [kacos2000](https://github.com/kacos2000/Jumplist-Browser)

## 许可证

基于 MIT 许可证分发。更多信息请参见 `LICENSE` 文件。

## 作者

由 [@Hellager](https://github.com/Hellager) 用 🦀 开发

[img_doc]: https://img.shields.io/badge/doc-latest-orange
[doc]: https://docs.rs/wincent/latest/wincent/
