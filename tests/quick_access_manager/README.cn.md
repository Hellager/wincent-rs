# QuickAccessManager 综合 VM 测试

面向新 Windows 系统版本的公开 API 验收套件。只在指定 VMware guest 中执行；
宿主只编译、传文件、取证和回滚，不运行任何 Quick Access 测试。

## 运行

前置条件：VMware Workstation、可用快照、VMware Tools、已登录的测试账户和
同会话 Explorer。宿主需要 Rust/Cargo 和 PowerShell 7，guest 使用系统
Windows PowerShell 5.1。测试会清空 Quick Access、取消用户固定、切换可见性并重启
Explorer，请只用专用 VM，不要在真实工作桌面手动执行 ignored 测试。

复制 `.env.example` 为 `.env`，填写 VMX、vmrun、用户名、密码和平台；`.env`
已由 `tests/.gitignore` 忽略。也可以使用进程环境变量或显式参数覆盖配置。

```powershell
pwsh -File tests/quick_access_manager/run_quick_access_manager_vm.ps1 `
    -Platform win10
```

Win11 使用同一套测试，仅改为 `-Platform win11` 并指定对应 VM。
可用 `-EvidenceDir <全新目录>` 指定证据目录。加密 VM 可设置
`WINCENT_VM_PASSWORD`。不要把真实密码写入已跟踪的文件。

测试期间不要操作 guest 桌面，也不要运行其他会写 Recent 的程序。脚本不会升级系统、
安装工具或修改库实现。

## 覆盖矩阵

| 功能 | 检查内容 |
| --- | --- |
| 构造与超时 | `new/default/builder/build/timeout/try_timeout`、builder/manager clone、零超时错误和 panic 契约 |
| 查询 | `get_items/get_item_paths` 三类同序、All 的 Recent-first 去重、精确路径与关键词区别、空关键词、大小写和无匹配 |
| 独立 Shell 对照 | 独立 STA 中真实 COM 枚举，与 manager 返回结果逐项同序比较，不把 manager 查询冒充 Shell oracle |
| 状态 | 原始 entry 全量检查：可见序号、窗口外不可见、pin 状态、缺失路径默认值、`lists/is_visible` |
| 路径与错误 | 文件和目录、空格和中文、大小写、正斜线、canonical、目录末尾斜线和点；空路径、缺失路径、重复项、不在列表 |
| 原生增删 | 25 个 Recent 文件、两个 pinned 目录、原生访问生成的 unpinned 父目录；新增可见、删除消失 |
| 锁 | Recent/Frequent/All；克隆共享、重复 lock/unlock 错误、独立 manager 与跨进程阻断、backing 文件写拒绝、锁内查询 |
| 锁生命周期 | 不相交类别同时持锁、单 clone drop 后仍锁定、最后一个 clone drop 释放、unlock 返回并删除新增 `.lnk` |
| 可见性 | 四种 target 的 hide/read/show/read；All 不含 Start；混合状态 All=false；fixture 前显式启用 Start 跟踪 |
| 刷新 | 打开的 Quick Access/Home 窗口；有内容和清空后 API 返回成功，不将其描述为像素级 UI 验证 |
| 清空 | 每种类别都先种入文件与目录；检查目标消失及另一类别保留；重复清空；允许 Windows 重建默认 pin |

原生写入后使用有界轮询等待持久化，不将 `Ok(())` 等同于立即可见，不自动重试超时 mutation。
前置 fixture 失败立即中止依赖阶段并记录 `suite.completed=false`，不会把能力缺失静默
跳过后声称完整通过。对新系统使用原生 API 构造数据，不注入 Win10 v4 fixture。

前 3 个 Recent 文件按 **66 秒间隔**加入，以原生方式生成可见的未固定父目录。
短时重复访问会合并计分，仅快速加入 25 个文件不足以生成该 fixture。
这两个等待是测试数据的必要准备，不是性能计时；含快照和回滚，一轮需要数分钟。
每个清空用例使用全新文件/目录路径。Recent 的 remove 可能留下 Shell 缓存中的屏蔽状态，
不能通过重新 add 上一阶段已移除的文件来建立下一阶段基线。

## 隔离与恢复

1. 宿主使用 `cargo test --no-run` 生成精确构建产物，记录 SHA-256 和源代码状态。
2. 在任何 guest 改动前创建唯一整机快照；快照失败则停止，不降级为无保护运行。
3. guest 校验 VMware 身份、实际 build 对应平台、账户、二进制哈希和 Explorer 会话；
   不把交互会话号写死为 1。
4. 停 Explorer，备份整个 Recent 树及三项可见性注册表值的存在性、类型和值。
   Rust 测试还要求运行标记和备份完成标记，默认 `#[ignore]`。
5. 成功或失败均在 guest `finally` 中停止测试进程、恢复文件和注册表、删除 fixture，
   在重启 Explorer 前比较完整文件树路径/长度/SHA-256 和注册表。
6. 另存 Explorer 重启后清单，不把系统对未初始化 CFB 的规范化伪装成精确恢复。
7. 宿主等待清理真正结束，再回收证据并回滚整机快照。恢复原电源状态，成功时只删除本轮
   临时快照；失败时保留本轮快照。宿主强制退出或崩溃无法保证 `finally` 执行，届时必须
   使用保留快照手动恢复，不能继续在未知状态上试验。

每轮独立目录和 run ID，拒绝复用已有证据目录，不会清除其他测试的旧备份。失败 marker
不代表恢复完成，只有 `finished.txt` 表示本轮清理尝试已结束。

## 证据与验收

默认目录：`temp/track_wincent/evidence/qam-<platform>-<run-id>/`。

| 文件 | 用途 |
| --- | --- |
| `result.json` | 检查明细、总数、失败数、完成标记、二进制退出码 |
| `checks.tsv`、`test.log` | 流式检查和完整测试日志，失败时仍保留已完成检查 |
| `environment.json` | build/UBR、用户、会话、exe SHA-256、提交、Shell DLL/Explorer 版本与哈希 |
| `final-state.json` | guest 文件/注册表恢复、fixture 删除、测试进程残留 |
| `host-result.json` | 整机快照回滚与最终整体通过标记 |
| `archive-verification.json` | 归档内每个原始文件（包括隐藏文件）的长度和 SHA-256 校验 |
| `evidence.zip` | 原始备份、恢复清单、注册表、阶段 CFB、raw entry、独立 Shell 输出、子进程日志 |
| `suite-source/`、`source.diff`、`source-status.txt` | 对应构建的套件源码及工作区改动，不收集 `.env` |
| `guest-running.png` | 尽力采集的截图，仅作辅助，不作为 UI 通过判据 |

完整通过要求：检查数大于零、`completed=true`、失败数为零、哈希匹配、guest
`restored=true`、无残留测试进程，以及宿主 `passed=true` / `snapshot_reverted=true`。
任何报告缺失、阶段未执行、恢复失败都不能视为通过。

## 测试范围与边界

本套件验证 `QuickAccessManager` 公开 API 的正常行为与错误契约。验收结论仅适用于
实际受测的库版本和 Windows build；每个目标系统版本都应重新运行本套件并保留证据，
不能将一次通过外推为其他版本或未来版本的兼容保证。

测试会将库的可见数量限制、排序和过滤规则与受测系统的真实 Shell 对账，
使用原生方式准备 fixture，不注入其他 Windows 版本的持久化数据，也不修改私有 API。
遇到不支持的格式或 Shell 行为变化时，应明确失败。

本套件不穷举底层 DestList API、损坏/CRC 恢复矩阵、故意制造的 `PartialEmpty`、
COM 中途崩溃、所有第三方并发写入、组策略覆盖、网络路径、私有 ABI 稳定性或逐像素 UI。
不使用亚秒级超时断言：原生操作耗时不确定，超时也不能证明原生修改已取消。
