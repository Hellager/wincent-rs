# Quick Access `get_item` 延迟测试（仅 VM）

本目录测量 Quick Access 列表的两条独立读取路径：

- **Shell 查询**：通过 STA COM 枚举当前 Quick Access Shell 命名空间。
- **CFB 解析**：读取并解析 `.automaticDestinations-ms` backing 文件本身。

测试在专用 Windows VM 内完成；host 只负责编排、文件传输、截图和证据回收。

当前数据是初步结果：一轮正式 20 样本运行加两轮重复运行，不用于容量规划。

## 文件

| 文件 | 作用 |
| --- | --- |
| `main.rs` | Rust 集成测试；默认 `#[ignore]`，并额外要求 `WINCENT_GET_ITEM_LATENCY=1`，避免在开发机误运行 |
| `run_get_item_latency_vm.ps1` | host 编排：构建、上传、启动、轮询 marker、截图、回收证据、校验恢复状态 |
| `vm_get_item_latency_start.ps1` | guest 启动器（session 0）：清理旧状态，在 Explorer 所在交互 session 创建临时 `/IT` 计划任务 |
| `vm_get_item_latency_guest.ps1` | guest 编排（session 1）：baseline 备份、运行测试、恢复、marker、环境/最终状态 JSON、证据 zip |
| `.env.example` | 配置模板；复制为 `.env` 后填写 VM 路径和 guest 账号 |
| `.env` | 本机实际配置，含 guest 密码，已被 git 忽略 |

## 测试夹具

测试在计时前构造以下状态，然后重启 Explorer，让 Shell 重新加载 backing 文件。
夹具构造不计入耗时。

| 列表 | raw 条目数 | 组成 |
| --- | ---: | --- |
| Recent Files | 120 | 20 visible + 100 hidden |
| Frequent Folders | 100 | 46 pinned + 4 visible unpinned + 50 hidden unpinned |

Frequent Folders 的 access count 覆盖 `0`、`1`、`3`、`5`：pinned 使用 `0`，
visible unpinned 使用 `3`，hidden unpinned 包含 `1`、`3`、`5`。Explorer 重启后，测试会
重新断言 raw 条目数、pinned 数、规则可见数和 access count 分布。

夹具从一个空的 v4 backing 文件开始，通过 crate 的公开 mutation API 创建；测试不会复用
当前用户的 Quick Access 状态。

## 场景

| 场景 | Shell 可见数 | Raw 数 |
| --- | ---: | ---: |
| `recent_full` | 20（结果窗口满） | 120 |
| `recent_under` | 8（不足窗口） | 120 |
| `frequent_full` | 50 | 100 |
| `frequent_under` | 40 | 100 |

Recent Files 在满窗口时返回 20 项；再隐藏 12 个 visible 条目后返回 8 项。Frequent
Folders 满状态为 46 pinned + 4 unpinned；将 10 个 pinned 条目 unpin 并隐藏后，剩 36
pinned + 4 unpinned。

## 测量方法

- **Shell 查询**：每次在一个新 worker 线程初始化 STA COM，打开 Recent Items 或
  Frequent Folders 命名空间，枚举 `FolderItems`，按文件/文件夹过滤并收集路径。COM
  初始化和命名空间访问都计入计时区间。
- **CFB 解析**：`wincent::destlist::parse_file` 打开 backing 文件，解析 CFB 容器和 raw
  DestList 条目，并统计 `cfb_entries()`。
- **计时区间**：一次读取调用，从调用前到返回。夹具构造、Explorer 重启、raw 校验和恢复
  都不计入。
- **样本**：每个场景/方法 2 次预热，再执行 20 次计时；每轮交替两种方法的执行顺序，避免
  固定由某一方法先运行。
- **统计**：对成功调用计算 mean、median、nearest-rank p95、min、max、sample stdev。

host 只有在预期 shell 数量、raw 数量和场景统计都写入结果后才判定成功。

## 环境要求

- Windows host，PowerShell 5.1 或更高，安装 Rust 工具链和 `cargo`。
- VMware Workstation，提供 `vmrun.exe`。
- Windows guest：目标用户已登录，`explorer.exe` 运行在 **session 1**；VMware Tools
  guest control 可用，且配置账号有权限。
- 测试期间不要用该 guest 做无关工作。测试会停止/重启 Explorer 并替换 Quick Access
  backing 文件；应使用专用 VM 或快照。

## 配置

```powershell
Copy-Item tests/get_item_latency/.env.example tests/get_item_latency/.env
# 编辑 .env：填写 vmrun.exe、.vmx、guest 用户名/密码、证据根目录和默认参数。
```

host 脚本的解析顺序：

```text
命令行参数 > .env > 非敏感内置默认值
```

`.env` 键：

| 键 | 含义 |
| --- | --- |
| `WINCENT_VM_VMRUN` | `vmrun.exe` 完整路径 |
| `WINCENT_VM_VMX` | 目标 `.vmx` 完整路径 |
| `WINCENT_VM_GUEST_USER` | vmrun guest control 使用的 guest 账号 |
| `WINCENT_VM_GUEST_PASSWORD` | guest 密码 |
| `WINCENT_EVIDENCE_ROOT` | host 证据根目录；相对路径从仓库根目录解析 |
| `WINCENT_GET_ITEM_LATENCY_ITERS` | 每个场景/方法的默认计时迭代数 |
| `WINCENT_GET_ITEM_LATENCY_WARMUP` | 每个场景/方法的默认预热次数 |

`WINCENT_VM_VMX`、`WINCENT_VM_VMRUN`、`WINCENT_VM_GUEST_USER`、
`WINCENT_VM_GUEST_PASSWORD` 没有硬编码兜底值。缺少时脚本会在接触 VM 之前停止，并提示创建
`.env`。

## 运行

```powershell
# 每个场景/方法 20 次计时 + 2 次预热
pwsh -File tests/get_item_latency/run_get_item_latency_vm.ps1 `
    -EvidenceDir temp/track_wincent/evidence/my-get-item-latency-run `
    -Iters 20 -Warmup 2
```

脚本内部构建命令：

```text
cargo test --release --test get_item_latency --no-run
```

然后在 guest 中运行：

```text
get_item_latency --ignored --exact --nocapture
```

## 当前初步结果

环境：专用 Windows 10 Home 22H2 `10.0.19045` VM，Explorer 运行在 session `1`。
一轮正式运行，每个场景/方法 20 个样本。单位毫秒。

| 场景 | Shell mean | Shell median | Shell p95 | CFB mean | CFB median | CFB p95 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `recent_full` | 87.629 | 68.063 | 241.312 | 1.549 | 1.366 | 2.257 |
| `recent_under` | 52.270 | 44.982 | 78.077 | 1.882 | 1.631 | 2.518 |
| `frequent_full` | 56.933 | 54.477 | 78.596 | 1.400 | 1.298 | 1.990 |
| `frequent_under` | 43.368 | 37.758 | 63.126 | 1.569 | 1.401 | 2.339 |

正式运行 + 两轮重复运行合计（每个场景/方法 `n = 60`）的中位数：

| 场景 | Shell median | CFB median |
| --- | ---: | ---: |
| `recent_full` | 68.630 | 1.468 |
| `recent_under` | 44.510 | 1.530 |
| `frequent_full` | 52.269 | 1.353 |
| `frequent_under` | 37.048 | 1.340 |

时间统计观察：

- 在每个场景中，CFB 解析都比 Shell 查询快约 25–50 倍（中位数约 1.3–1.5 ms，
  Shell 约 37–69 ms）。
- 结果窗口满时更慢：Recent 为 20 vs 8，Frequent 为 50 vs 40。
- Shell 查询偶发长尾（观察到 0.2–1.2 s 的样本），部分运行的 mean 会被拉高；跨场景比较
  应优先使用 median 和 p95。

## 验证与恢复

每轮 guest 编排都会：

1. 备份两个 Quick Access CFB 和完整 Recent `.lnk` 集合；
2. 写入含逐文件 SHA-256 的 baseline manifest；
3. 运行夹具构造与测量测试；
4. 停止 Explorer，写回两个 CFB，删除测试新增的 `.lnk`，恢复 baseline `.lnk`，重启
   Explorer；
5. 比较 CFB SHA-256 和 `.lnk` manifest；
6. 输出 `final-state.json`；失败时写 `FAILURE.txt`/`RESTORE_FAILURE.txt`，并打包证据 zip。

host 只有在 `restored = true`、无 failure marker、且 guest 报告的 exe SHA-256 与本机构建
一致时才会判定成功。证据写入 `temp/track_wincent/evidence/`，该目录已被 git 忽略。

## 边界

- 结论只适用于该 VM 镜像、夹具组成、Quick Access 条目数量、磁盘和 Defender 状态。
- 没有 ETW/Procmon 分解，不能把耗时归因到某个 Shell 或 CFB parser 内部阶段。
- 每轮 20 个样本足以做初步 mean/median/p95；p95 接近最大值（n=20 时为第二高样本）。
  尾延迟分析应增加样本量。
- 测试会强制重启 Explorer，必须在专用 VM 上运行。中断可能留下修改后的 Quick Access
  状态；可从 baseline 副本或 VM 快照恢复。
- 预热不是可选优化：Explorer 重启后的前几次查询可能明显慢于稳态。正式运行每个
  场景/方法使用 2 次预热。
