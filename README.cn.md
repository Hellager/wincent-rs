# wincent-rs

![Crates.io Version](https://img.shields.io/crates/v/wincent)
[![][img_doc]][doc]
![Crates.io Total Downloads](https://img.shields.io/crates/d/wincent)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Hellager/wincent-rs/publish.yml)
![Crates.io License](https://img.shields.io/crates/l/wincent)

其他语言版本：[English](README.md) | [中文](README.cn.md)

## 概述

Wincent 是一个用于管理 Windows 快速访问及最近使用的文件( Win11 )/常用文件夹和最近访问的文件( Win10 )功能的 Rust 库，提供对其中内容的全面控制。

## 功能特性

- 🔍 查询快速访问内容
- ➕ 添加项目到快速访问
- 🗑️ 移除特定快速访问条目
- 🧹 清空快速访问项目
- 👁️ 切换快速访问项目的可见性

## 安装

在 `Cargo.toml` 中添加以下依赖：

```toml
[dependencies]
wincent = "0.1.1"
```

## 注意事项

- 功能相关实现高度依赖系统 api，windows 可能会出于安全性考虑收紧相关权限或调用，届时可能失效
- 个人系统环境可能与测试环境不同，导致功能无法正常使用，作者已观察到类似问题，大概率是由于部分软件修改了相关注册表导致，未定位到具体注册表项，使用前可调用相关函数检查是否可行，主要是对文件夹的操作
- 可见性部分会修改注册表，可能会导致意外结果，大概率窗口布局会受到影响，请谨慎使用

## 快速开始

### 查询快速访问内容

```rust
use wincent::{
    feasible::{check_script_feasible, fix_script_feasible}, 
    query::get_quick_access_items, 
    error::WincentError
};

fn main() -> Result<(), WincentError> {
    // 检查快速访问是否可用
    if !check_script_feasible()? {
        println!("Fixing script execution policy...");
        fix_script_feasible()?;
    }

    // 列出所有当前快速访问项目
    let quick_access_items = get_quick_access_items()?;
    for item in quick_access_items {
        println!("快速访问项目: {}", item);
    }

    Ok(())
}
```

### 移除快速访问条目

```rust
use wincent::{
    query::get_recent_files, 
    handle::remove_from_recent_files, 
    error::WincentError
};

fn main() -> Result<(), WincentError> {
    // 从最近项目中移除敏感文件
    let recent_files = get_recent_files()?;
    for item in recent_files {
        if item.contains("password") {
            remove_from_recent_files(&item)?;
        }
    }

    Ok(())
}
```

### 切换可见性

```rust
use wincent::{
    visible::{is_recent_files_visiable, set_recent_files_visiable}, 
    error::WincentError
};

fn main() -> Result<(), WincentError> {
    let is_visible = is_recent_files_visiable()?;
    println!("最近文件可见性: {}", is_visible);

    set_recent_files_visiable(!is_visible)?;
    println!("可见性已切换");

    Ok(())
}
```

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
