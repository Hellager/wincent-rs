# Refresh synchronization experiment (VM only)

## Question

When `.automaticDestinations-ms` bytes are replaced while a Quick Access window
is open, which refresh path makes the **open UI**, the **terminal/Shell query**
and the **on-disk file** agree?

The experiment compares:

1. **current program refresh**: `QuickAccessManager::refresh_explorer()`
   (`SHChangeNotify(SHCNE_ASSOCCHANGED)` followed by `IShellView::Refresh` on
   matching Quick Access/Home windows);
2. **explicit SHChangeNotify + refresh**: an additional explicit notification
   followed by the same unified refresh.

It also isolates `SHChangeNotify` alone to see whether the notification by
itself updates the open view.

## Stable reproduction

`main.rs` builds two Recent Files states:

| State | Raw entries | Visible items |
| --- | ---: | --- |
| Windows 10 A (baseline) | 5 | 5 |
| Windows 10 B | 6 | 6 (adds `refresh-sync-marker.txt`) |
| Windows 11 A/B | Native v6 snapshots created through system APIs | Records actual Shell counts; B adds the marker |

The test:

1. opens one Quick Access window showing A;
2. replaces only the Recent CFB bytes with B (`ReplaceFileW`), without
   notification or refresh;
3. verifies that the file parser, the public rule-based query and a fresh live
   Shell query all see B while the open window still shows A;
4. applies the current refresh, captures the result, restarts Explorer and
   resets to A;
5. replaces A with B again and applies `SHChangeNotify + refresh`;
6. resets and tests `SHChangeNotify` alone, followed by a refresh.

Phase markers are written to `targets/phases`; the host polls them and captures
screenshots. `result.json` records file/rule/Shell observations and timings.

## Files

| File | Purpose |
| --- | --- |
| `main.rs` | VM integration test. Ignored and gated behind `WINCENT_REFRESH_SYNC=1`. |
| `run_refresh_sync_vm.ps1` | Host orchestration, phase polling and screenshots. |
| `vm_refresh_sync_start.ps1` | Guest session-1 launcher with an interactive path and task fallback. |
| `vm_refresh_sync_guest.ps1` | Guest baseline/restore and test execution. |
| `.env.example` | VM/credential/output configuration template. |
| `.env` | Local values, including the guest password. Git-ignored. |

## Configuration and run

```powershell
Copy-Item tests/refresh_sync/.env.example tests/refresh_sync/.env
# Edit .env, then:

pwsh -File tests/refresh_sync/run_refresh_sync_vm.ps1 `
    -Platform win10 `
    -PhaseSleepMs 6000
```

The test binary is built with
`cargo test --release --test refresh_sync --no-run` and run with
`refresh_sync --ignored --exact --nocapture` in the guest.

`WINCENT_REFRESH_SYNC_PHASE_SLEEP_MS` (default 3000) is the time each phase
stays on screen; use a larger value on a slow VM so the host can capture the
correct screenshot.

Encrypted VMs can set the optional `WINCENT_VM_PASSWORD`; the host passes it
through `vmrun -vp`.

Windows 10 uses deterministic synthetic-v4 A/B states. Windows 11 creates
native-v6 A/B snapshots through system `add_item`, then performs whole-file
replacement. This does not implement or validate a v6 entry writer.
If native Recent tracking cannot produce both snapshots, the Windows 11 run
returns `status: "skipped"` with reason `native_recent_tracking_unavailable`.
That result is an expected capability outcome, not a successful refresh sample.

Set `WINCENT_VM_PLATFORM=win10|win11` or pass `-Platform`. The guest validates
the actual Windows build before modifying Quick Access.

## Results (Windows 10 Home 22H2 `10.0.19045`)

The historical baseline below contains three valid runs (`v2`, `v3`, `v4`)
collected before the public API was unified. In this section, "view-only
refresh" means the former `IShellView::Refresh`-only implementation.

### Observation matrix

| Stage | File has marker | Rule query has marker | Live Shell has marker | Open UI shows marker |
| --- | :---: | :---: | :---: | :---: |
| After B replacement, no refresh | yes | yes | yes | **no (5 items)** |
| View-only refresh | yes | yes | yes | v2: yes; v3/v4: **window blank/unavailable** |
| SHChangeNotify + view-only refresh | yes | yes | yes | **yes (6 items), 3/3** |
| SHChangeNotify only | yes | yes | yes | **no (5 items), 2/2** |

The open UI was stale in every A→B replacement before refresh, while file,
rule-based terminal query and live Shell query were already on B. This is a
stable reproduction of the desync: only the UI lags; the terminal query is
already synchronized with the file.

### Timing

| Operation | v2 | v3 | v4 | Mean |
| --- | ---: | ---: | ---: | ---: |
| View-only refresh (ms) | 53.253 | 79.384 | 53.128 | 61.922 |
| `SHChangeNotify` only (ms) | 0.100 | 0.105 | 0.070 | 0.092 |
| `SHChangeNotify` + refresh (ms) | 57.767 | 56.961 | 58.575 | 57.768 |

Repeated current-refresh cost after the UI was settled: `n = 15`, mean
**56.390 ms**, median **50.210 ms**, min **23.224 ms**, max **123.704 ms**.

The notification itself costs about **0.09 ms**; the refresh dominates. Adding
the notification does not materially change the measured refresh cost.

## Conclusion

- **The desync exists and is reproducible**: after a raw CFB replacement, the
  open Quick Access window keeps the old list while the file, the rule-based
  query and a fresh Shell query already return B.
- **`SHChangeNotify` alone is not sufficient** for this case: the open UI stayed
  on A in both notify-only trials.
- **View-only refresh is not reliable** for a replaced backing file: it
  updated the UI in one run, but in two runs the Quick Access window was not
  visible when the phase screenshot was taken (only the desktop background was
  captured, although the refresh call returned success). Do not use it as the
  sole synchronization step after raw replacement.
- **`SHChangeNotify(SHCNE_ASSOCCHANGED)` + view-only refresh is the effective and
  stable combination**: it updated the open UI in 3/3 runs, with terminal and
  file already in sync. Its measured cost is about **57.8 ms**, dominated by
  the refresh; the notification adds about **0.09 ms**.

The public `QuickAccessManager::refresh_explorer()` now implements that
notification-plus-targeted-view-refresh sequence as one operation.

## Limitations

- UI synchronization is inferred from phase screenshots, not from a UI
  Automation text dump. The “blank” current-refresh result shows that the
  window was not visible in the capture; it is not a process-level crash proof.
- Only three valid runs, one fixture size and one Windows 10 22H2 image.
- The test forcibly restarts Explorer and replaces Quick Access backing files;
  run it only on a dedicated VM or snapshot.
- The current test code calls the unified API. The historical timing table is
  retained to document why the notification was folded into that API.
