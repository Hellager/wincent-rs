# wincent-rs

[![](https://img.shields.io/crates/v/wincent.svg)](https://crates.io/crates/wincent)
[![][img_doc]][doc]
![Crates.io Total Downloads](https://img.shields.io/crates/d/wincent)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Hellager/wincent-rs/publish.yml)
![Crates.io License](https://img.shields.io/crates/l/wincent)

其他语言版本：[English](README.md) | [中文](README.cn.md)

## 概述

Wincent 是一个用于管理 Windows 快速访问功能的 Rust 库，提供异步支持和强大的错误处理，全面控制最近使用的文件和常用文件夹。

## 功能特性

- 🔍 全面的快速访问管理
  - 查询最近文件和常用文件夹
  - 支持强制更新的添加/删除操作
  - 可控制系统默认项的清理
  - 项目存在性检查

- 🛡️ 强大的操作处理
  - 操作超时保护
  - 脚本执行缓存
  - 全面的错误处理

- ⚡ 性能优化
  - 使用 OnceCell 实现延迟初始化
  - 脚本执行缓存
  - 批量操作支持
  - 强制刷新控制

## 安装

在 `Cargo.toml` 中添加以下依赖：

```toml
[dependencies]
wincent = "0.1.2"
```

## 快速开始

### 基础操作

```rust
use wincent::predule::*;

#[tokio::main]
async fn main() -> WincentResult<()> {
    // 初始化管理器
    let manager = QuickAccessManager::new().await?;
    
    // 添加文件到最近文件并强制更新
    manager.add_item(
        "C:\\path\\to\\file.txt",
        QuickAccess::RecentFiles,
        true
    ).await?;
    
    // 查询所有项目
    let items = manager.get_items(QuickAccess::All).await?;
    println!("快速访问项目: {:?}", items);
    
    Ok(())
}
```

### 高级用法

```rust
use wincent::predule::*;

#[tokio::main]
async fn main() -> WincentResult<()> {
    let manager = QuickAccessManager::new().await?;
    
    // 清空最近文件并强制刷新
    manager.empty_items(
        QuickAccess::RecentFiles,
        true,  // 强制刷新
        false  // 保留系统默认项
    ).await?;
    
    Ok(())
}
```

## 最佳实践

- 添加最近文件时使用 `force_update` 以立即显示
- 仅在必要时检查操作可行性
- 批量操作后清理缓存
- 谨慎使用 `also_system_default` 清理项目
- 在生产环境中妥善处理超时情况

## 系统要求和限制

- **API 依赖**: 本库核心功能依赖于 Windows 系统 API。由于 Windows 安全策略和更新，某些操作可能需要提升权限或受到限制。

- **环境兼容性**: 不同的 Windows 环境（包括版本、配置和已安装软件）可能影响库的功能。特别是：
  - 第三方软件可能修改相关注册表项
  - 系统安全策略可能限制 PowerShell 脚本执行
  - Windows 资源管理器集成在不同 Windows 版本中可能有所不同

- **预检查建议**: 在执行操作之前，特别是文件夹管理操作，建议使用内置的可行性检查：
  ```rust
  let (can_query, can_modify) = manager.check_feasible().await;
  if !can_modify {
      println!("警告：系统环境可能限制修改操作");
  }
  ```

这些限制是 Windows 快速访问功能本身的固有特性，并非本库特有。我们提供了全面的错误处理和状态检查机制，以帮助您优雅地处理这些场景。

## 注意事项

- 功能相关实现高度依赖系统 api，windows 可能会出于安全性考虑收紧相关权限或调用，届时可能失效
- 个人系统环境可能与测试环境不同，导致功能无法正常使用，作者已观察到类似问题，大概率是由于部分软件修改了相关注册表导致，未定位到具体注册表项，使用前可调用相关函数检查是否可行，主要是对文件夹的操作
- 可见性部分会修改注册表，可能会导致意外结果，大概率窗口布局会受到影响，请谨慎使用

## 错误处理

该库使用 Rust 的 `Result` 类型进行全面的错误管理，允许在操作快速访问过程中精确处理潜在问题。

## 兼容性

- 支持 Windows 10 和 Windows 11
- 需要 Rust 1.60.0 或更高版本

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

# 安装开发依赖
cargo build
cargo test
```

## 免责声明

本库与系统级快速访问功能进行交互。在进行重要更改之前，请确保您具有适当的权限并创建备份。

## 支持

如果您遇到任何问题或有疑问，请在我们的 GitHub 仓库上提出 issue。

## 致谢

- [Castorix31](https://learn.microsoft.com/en-us/answers/questions/1087928/how-to-get-recent-docs-list-and-delete-some-of-the)
- [Yohan Ney](https://stackoverflow.com/questions/30051634/is-it-possible-programmatically-add-folders-to-the-windows-10-quick-access-panel)

## 许可证

基于 MIT 许可证分发。更多信息请参见 `LICENSE` 文件。

## 作者

由 [@Hellager](https://github.com/Hellager) 用 🦀 开发

[img_doc]: https://img.shields.io/badge/doc-latest-orange
[doc]: https://docs.rs/wincent/latest/wincent/