# Quick Access `add_item` 平均耗时测试（仅 VM）

本目录用于测量 `QuickAccessManager::add_item` 向 **Recent Files** 和
**Frequent Folders** 添加新项的端到端耗时。所有 Quick Access 变更都在专用 Windows 10
guest VM 内完成；host 只负责编排、文件传输、截图和证据回收。

当前数据是初步结果：每个目标两轮独立运行、每轮 20 个样本，不用于容量规划。本文是
`README.md` 的中文版本。

## 文件

| 文件 | 作用 |
| --- | --- |
| `main.rs` | Rust 集成测试；默认 `#[ignore]`，并额外要求 `WINCENT_ADD_LATENCY=1`，避免在开发机误运行 |
| `run_add_item_latency_vm.ps1` | host 编排：构建、上传、启动、轮询 marker、截图、回收证据、校验恢复状态 |
| `vm_add_item_latency_start.ps1` | guest 启动器（session 0）：清理旧状态，在 Explorer 所在的交互 session 创建临时 `/IT` 计划任务 |
| `vm_add_item_latency_guest.ps1` | guest 编排（session 1）：baseline 备份、运行测试、恢复、marker、环境/最终状态 JSON、证据 zip |
| `.env.example` | 配置模板；复制为 `.env` 后填写 VM 路径和 guest 账号 |
| `.env` | 本机实际配置，含 guest 密码，已被 git 忽略 |

## 测量方法

只测 shell/COM 的 `add_item` 路径。

- **Recent Files**：路径校验 → Shell 重复预检 → 在 STA 线程调用
  `SHAddToRecentDocs(SHARD_PATHW, path)`。
- **Frequent Folders**：路径校验 → Shell 可见性预检 → Shell pin verb（带 manager
  的重试策略）。
- **计时区间**：一次 `manager.add_item(path, kind, options)` 调用，从调用前到返回
  `Ok`/`Err`。不包含进程启动、测试框架、目标创建、最终 raw 校验和 CFB 恢复。
- **样本**：每个系列 2 次预热（同代码路径，不计入统计）+ 20 次计时调用。每次调用都使用
  全新目标（Recent 为 `.txt` 文件，Frequent 为新目录），避免重复检测提前短路。
- **顺序**：每轮轮换系列顺序，避免某个系列固定处于相同的缓存/Explorer 位置。
- **统计**：对成功调用计算 mean、median、nearest-rank p95、min、max、sample stdev。

全部样本结束后，测试直接解析两个 `.automaticDestinations-ms`，统计本轮目标中有多少以 raw
DestList entry 的形式存在。

## 环境要求

- Windows host，PowerShell 5.1 或更高，安装 Rust 工具链和 `cargo`。
- VMware Workstation，提供 `vmrun.exe`。
- Windows guest：目标用户已登录，`explorer.exe` 运行在 **session 1**；VMware Tools
  guest control 可用，且配置的账号有权限。
- 测试期间不要用该 guest 做无关工作，因为测试会真实修改 Quick Access 状态。

## 配置

```powershell
Copy-Item tests/add_item_latency/.env.example tests/add_item_latency/.env
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
| `WINCENT_ADD_LATENCY_SERIES` | 默认系列，如 `recent,frequent` |
| `WINCENT_ADD_LATENCY_ITERS` | 默认每系列计时迭代数 |
| `WINCENT_ADD_LATENCY_WARMUP` | 默认每系列预热次数 |

`WINCENT_VM_VMX`、`WINCENT_VM_VMRUN`、`WINCENT_VM_GUEST_USER`、
`WINCENT_VM_GUEST_PASSWORD` 没有硬编码兜底值。缺少时脚本会在接触 VM 之前停止，并提示创建
`.env`。

## 运行

```powershell
# Recent + Frequent，每系列 20 次计时 + 2 次预热
pwsh -File tests/add_item_latency/run_add_item_latency_vm.ps1 `
    -EvidenceDir temp/track_wincent/evidence/my-add-latency-run `
    -Series 'recent,frequent' -Iters 20 -Warmup 2

# 可选：AddOptions::new().refresh_explorer()
pwsh -File tests/add_item_latency/run_add_item_latency_vm.ps1 `
    -EvidenceDir temp/track_wincent/evidence/my-add-latency-refresh `
    -Series 'recent-refresh,frequent-refresh' -Iters 20 -Warmup 2
```

脚本内部构建命令：

```text
cargo test --release --test add_item_latency --no-run
```

然后在 guest 中运行：

```text
add_item_shell_latency --ignored --exact --nocapture
```

## 当前初步结果

环境：专用 Windows 10 Home 22H2 `10.0.19045` VM，Explorer 运行在 session `1`。
两轮独立默认运行（v2、v4-repeat），每轮 20 个样本；每个目标合并 `n = 40`。单位毫秒。

| 目标 | n | mean | median | p95 | min | max | stdev |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Recent Files | 40 | **589.440** | 555.351 | 846.560 | 159.100 | 979.939 | 167.269 |
| Frequent Folders | 40 | **442.816** | 437.952 | 816.634 | 190.709 | 838.845 | 179.890 |

每轮均值：

| 运行 | Recent 均值 | Frequent 均值 |
| --- | ---: | ---: |
| v2 | 601.716 | 442.261 |
| v4-repeat | 577.163 | 443.371 |

补充的 `refresh_explorer()` 运行（每系列 20 样本，独立批次）：Recent **584.025 ms**、
Frequent **643.158 ms**。Frequent 的约 200 ms 增量只作为方向性参考，因为 Explorer
窗口/负载状态没有与默认轮严格配对。

时间统计观察：

- 该 VM 上 `add_item` 平均约 **0.59 s**（Recent）、**0.44 s**（Frequent）。
- 单次调用波动较大（stdev 约为均值的 26%–46%），应同时看 mean、median 和 p95。
- Frequent 两轮均值差异约 0.2%；Recent 差异约 4%。

## 验证与恢复

每轮 guest 编排都会：

1. 备份两个 Quick Access CFB 和完整 Recent `.lnk` 集合；
2. 写入含逐文件 SHA-256 的 baseline manifest；
3. 运行 benchmark；
4. 停止 Explorer，写回两个 CFB，删除测试新增的 `.lnk`，恢复 baseline `.lnk`，重启
   Explorer；
5. 比较 CFB SHA-256 和 `.lnk` manifest；
6. 输出 `final-state.json`；失败时写 `FAILURE.txt`/`RESTORE_FAILURE.txt`，并打包证据 zip。

host 只有在 `restored = true`、无 failure marker、且 guest 报告的 exe SHA-256 与本机构建
一致时才会判定成功。证据写入 `temp/track_wincent/evidence/`，该目录已被 git 忽略。


## 边界

- 结论只适用于该 VM 镜像、当前 Quick Access 条目数量、磁盘和 Defender 状态。
- 没有 ETW/Procmon 分解，不能把耗时归因到某个 Shell 内部阶段。
- 每轮 20 个样本足以做初步 mean/median/p95；p95 接近最大值（n=20 时为第二高样本）。
  尾延迟分析应增加样本量。
- refresh 变体是独立批次，与默认配置的对比只具有方向性，不是严格控制变量实验。
- 预热不是可选优化。布局验证时曾用 0 次预热，第一次冷启动 Frequent pin 触发了 30 s 的 COM
  超时。至少保留 1–2 次预热；正式运行使用 2 次。
