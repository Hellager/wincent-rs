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
- 以保守 `.lnk` 清理恢复默认快速访问状态，并提供可选深度清理
- 精确路径或关键词检查项目是否存在
- 批量添加/删除，逐项收集错误
- Shell 和 PowerShell 操作具备调用方等待超时保护
- 快速访问分区可见性控制
- Windows 推荐的项目中最近使用文件可见性控制
- DestList 元数据访问

## 安装

在 `Cargo.toml` 中添加以下依赖：

```toml
[dependencies]
wincent = "0.2.5"
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
        AddOptions::new()
            .force_recent_files_rebuild()
            .refresh_explorer(),
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

## 系统要求与限制

- **操作系统**：Windows 10 或 Windows 11。
- **Rust**：1.85.0 或更高版本。
- **状态一致性**：快速访问状态由 Windows Explorer 维护，修改操作后结果可能有短暂延迟，Explorer 也可能在不同版本间以异步方式重建状态。
- **超时语义**：超时限制的是调用方等待结果的时间，而不是底层 Shell 或 COM 调用的实际运行时间。已经超时的 COM 操作仍可能稍后完成并影响 Explorer 状态。
- **固定文件夹清理超时**：在 `empty` 操作中显式移除可见固定文件夹时，可用 `EmptyOptions::with_pinned_folders_timeout()` 覆盖 snapshot/unpin 超时。未设置时使用 manager timeout。
- **恢复清理**：默认恢复清理只删除目标类型可解析为对应文件或文件夹分类的 `.lnk` 文件。使用 `RestoreDefaultsOptions::deep_lnk_cleanup()` 或 CLI `restore --deep` 时，也会删除无法解析或目标类型未知的 `.lnk` 文件。
- **推荐的项目中最近使用文件可见性**：Start Recommended 命名的 API 写入当前用户的 `Explorer\Advanced\Start_TrackDocs` 值，控制最近使用的文件是否显示在 Windows 推荐的项目中。策略值、MDM 设置、Windows 版本或 Explorer 版本可能覆盖或延迟可见效果。
- **实验性 DestList 删除**：experimental remove API 会直接重建 Explorer backing file，并可能删除 Recent 文件夹中匹配的 `.lnk` 文件；其稳定性弱于 parser/query API。

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
- [Grant Funtila](https://www.ninjaone.com/blog/clear-the-recommended-section-in-the-start-menu-in-windows-11/)

## 许可证

基于 MIT 许可证分发。更多信息请参见 `LICENSE` 文件。

## 作者

由 [@Hellager](https://github.com/Hellager) 用 🦀 开发

[img_doc]: https://img.shields.io/badge/doc-latest-orange
[doc]: https://docs.rs/wincent/latest/wincent/
