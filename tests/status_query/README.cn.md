# 状态查询 VM 验证

验证 `QuickAccessManager::quick_access_status()` / `recent_file_position()` /
`frequent_folder_slot()` 与 Explorer Shell 结果的一致性，并测量查询成本。

## 运行

```powershell
pwsh -File tests/status_query/run_status_query_vm.ps1 `
    -EvidenceDir temp/track_wincent/evidence/status-query-<date>-v1
```

配置见同目录 `.env`（从 `.env.example` 复制，`tests/.gitignore` 会忽略该文件）。
`-NoMutate` 可跳过 add/remove 闭环。

## 行为

1. 备份 Recent/Frequent CFB 与 Recent `.lnk`；
2. 构造 fixture：25 个 Recent 文件（`add_item`）+ 2 个 Frequent 未固定目录（`add_item_direct`）；
3. 逐项校验序号、pin 状态、窗口/槽位外条目、路径变体；
4. 执行 add→可见、remove→不可见、pin→可见、unpin→不可见的闭环；
5. 测量 parse / stat / 查询耗时；
6. `finally` 恢复两个 CFB 与 `.lnk`，比较 SHA-256 与链接清单。

## 证据

`result.json`（检查明细）、`timings.csv`、`environment.json`、`final-state.json`、
`benchmark.log`、截图与 `status-query-evidence.zip`。
