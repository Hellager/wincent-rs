# 状态查询 VM 验证

验证 `QuickAccessManager::quick_access_status()` / `recent_file_position()` /
`frequent_folder_slot()` 与 Explorer Shell 结果的一致性，并测量查询成本。

## 运行

```powershell
pwsh -File tests/status_query/run_status_query_vm.ps1 `
    -Platform win10 `
    -EvidenceDir temp/track_wincent/evidence/status-query-<date>-v1
```

配置见同目录 `.env`（从 `.env.example` 复制，`tests/.gitignore` 会忽略该文件）。
`-NoMutate` 可跳过 add/remove 闭环。
加密 VM 可设置可选的 `WINCENT_VM_PASSWORD`；host 优先通过 VMware `-interactive` 进入
Explorer session，并保留计划任务回退。

## 行为

1. 备份 Recent/Frequent CFB 与 Recent `.lnk`；
2. 构造平台 fixture：Win10 使用 25 个 Recent + 2 个 direct unpinned Frequent；Win11
   使用 25 个 Recent + 2 个系统 pin 的 Frequent，不调用 v4 direct writer；
3. 逐项校验序号、pin 状态、窗口/槽位外条目、路径变体；
4. 执行 add→可见、remove→不可见、pin→可见、unpin→不可见的闭环；
5. 测量 parse / stat / 查询耗时；
6. `finally` 按 absent/zero-length/populated 原始状态恢复两个 CFB，并比较链接清单。

`-Platform win10|win11` 也可通过 `.env` 的 `WINCENT_VM_PLATFORM` 设置；guest 会校验实际
Windows build，平台不匹配时在任何 Quick Access 修改前停止。

Win11 Recent 跟踪按 capability 处理：如果系统 `add_item` 未持久化 Recent 条目，测试会
明确跳过 Recent mutation 并写入 `fixture_capabilities`；原生 Frequent pin/remove 仍会执行。

## 证据

`result.json`（检查明细）、`timings.csv`、`environment.json`、`final-state.json`、
`benchmark.log`、截图与 `status-query-evidence.zip`。
