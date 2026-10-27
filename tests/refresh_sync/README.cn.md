# 刷新同步实验（仅 VM）

## 问题

当 Quick Access 窗口已经打开时，直接替换 `.automaticDestinations-ms` 文件内容，哪种刷新方式可以让
**已打开 UI**、**终端/Shell 查询**和**磁盘文件**三者一致？

本实验比较：

1. **当前程序刷新**：`QuickAccessManager::refresh_explorer_view()`
   （对匹配的 Quick Access/Home 窗口调用 `IShellView::Refresh`）；
2. **SHChangeNotify + refresh**：先调用
   `SHChangeNotify(SHCNE_ASSOCCHANGED)`，再执行同一个 refresh。

另外单独测试 `SHChangeNotify`，确认它本身是否足以更新已打开视图。

## 稳定复现方法

`main.rs` 构造两个 Recent Files 状态：

| 状态 | Raw 条目数 | 可见项 |
| --- | ---: | --- |
| A（基线） | 5 | 5 |
| B | 6 | 6（新增 `refresh-sync-marker.txt`） |

实验步骤：

1. 打开一个显示 A 的 Quick Access 窗口；
2. 只用 `ReplaceFileW` 替换 Recent CFB 字节为 B，不发送通知、不刷新；
3. 验证文件解析、public rule 查询和新的 live Shell 查询都返回 B，而已打开窗口仍显示 A；
4. 执行当前 refresh，截图并记录；重启 Explorer 重置到 A；
5. 再次 A→B，执行 `SHChangeNotify + refresh`；
6. 重置后再执行 `SHChangeNotify`，观察单独通知是否生效，随后再 refresh。

测试在 `targets/phases` 下写阶段 marker；host 轮询 marker 并截图。`result.json` 记录
file/rule/Shell 观察和耗时。

## 文件

| 文件 | 作用 |
| --- | --- |
| `main.rs` | VM 集成测试；`#[ignore]` + `WINCENT_REFRESH_SYNC=1` 双重门禁 |
| `run_refresh_sync_vm.ps1` | host 编排、阶段轮询和截图 |
| `vm_refresh_sync_start.ps1` | guest session 1 计划任务启动器 |
| `vm_refresh_sync_guest.ps1` | guest baseline/恢复和测试执行 |
| `.env.example` | VM/凭据/输出目录配置模板 |
| `.env` | 本机实际配置，含 guest 密码，已被 git 忽略 |

## 配置与运行

```powershell
Copy-Item tests/refresh_sync/.env.example tests/refresh_sync/.env
# 编辑 .env，然后：

pwsh -File tests/refresh_sync/run_refresh_sync_vm.ps1 `
    -PhaseSleepMs 6000
```

测试二进制由 `cargo test --release --test refresh_sync --no-run` 构建，在 guest 中通过
`refresh_sync --ignored --exact --nocapture` 运行。

`WINCENT_REFRESH_SYNC_PHASE_SLEEP_MS`（默认 3000）控制每个阶段在屏幕上停留的时间；
VM 较慢时应增大该值，避免 host 截到错误阶段。

## 结果（Windows 10 Home 22H2 `10.0.19045`）

三次有效运行（`v2`、`v3`、`v4`），每次使用全新的 Explorer 窗口和同一 A/B 夹具。

### 观察矩阵

| 阶段 | 文件含 marker | Rule 查询含 marker | Live Shell 含 marker | 已打开 UI 显示 marker |
| --- | :---: | :---: | :---: | :---: |
| B 替换后、未刷新 | 是 | 是 | 是 | **否（5 项）** |
| 仅当前 refresh | 是 | 是 | 是 | v2：是；v3/v4：**窗口空白/不可见** |
| SHChangeNotify + refresh | 是 | 是 | 是 | **是（6 项），3/3** |
| 仅 SHChangeNotify | 是 | 是 | 是 | **否（5 项），2/2** |

每次 A→B 替换后，UI 都显示旧列表，而文件、rule 查询和 live Shell 查询已经是 B。这是一个稳定的
不同步复现：只有 UI 滞后；终端查询已经与文件一致。

### 耗时

| 操作 | v2 | v3 | v4 | 均值 |
| --- | ---: | ---: | ---: | ---: |
| 仅当前 refresh（ms） | 53.253 | 79.384 | 53.128 | 61.922 |
| 仅 `SHChangeNotify`（ms） | 0.100 | 0.105 | 0.070 | 0.092 |
| `SHChangeNotify` + refresh（ms） | 57.767 | 56.961 | 58.575 | 57.768 |

UI 稳定后的重复 refresh：`n = 15`，mean **56.390 ms**，median **50.210 ms**，
min **23.224 ms**，max **123.704 ms**。

通知本身约 **0.09 ms**，耗时几乎全部来自 refresh；加上通知不会明显增加总耗时。

## 结论

- **不同步确实存在且可稳定复现**：raw CFB 替换后，已打开 Quick Access 窗口保留旧列表，
  而文件、rule 查询和新的 Shell 查询已经返回新项。
- **单独 `SHChangeNotify` 不足以更新 UI**：两次 notify-only 试验中窗口都保持旧内容。
- **单独当前 refresh 不可靠**：一次更新成功，两次在阶段截图中 Quick Access 窗口不可见
  （只截到桌面背景，尽管 refresh 调用返回成功）。不能把它作为 raw 替换后的唯一同步步骤。
- **`SHChangeNotify(SHCNE_ASSOCCHANGED)` + 当前 refresh 是有效且稳定的组合**：
  3/3 次更新 UI，terminal/file 本来已同步。耗时约 **57.8 ms**，主要来自 refresh。
- 对已经一致的视图做普通刷新时，单独当前 refresh 仍是最便宜的路径。

实践建议：在 Explorer 外部修改 backing file 后，先发 `SHCNE_ASSOCCHANGED`，再执行当前
refresh。

## 边界

- UI 同步结论来自阶段截图，不是 UI Automation 文本断言；“空白”只说明截图时窗口不可见，
  不等于已经做进程级崩溃证明。
- 只有三次有效运行、一种夹具大小和一个 Windows 10 22H2 镜像。
- 测试会强制重启 Explorer 并替换 Quick Access backing file，只能用于专用 VM 或快照。
- 为实验新增了 `QuickAccessManager::refresh_explorer_view()`；它执行与 mutation 后相同的
  native refresh，不发送 change notification。
