# Quick Access `add_item` latency test (VM only)

Measures the end-to-end latency of `QuickAccessManager::add_item` when adding
new items to **Recent Files** and **Frequent Folders**. The benchmark runs
inside a dedicated Windows 10 or Windows 11 VM; the host only orchestrates the VM, transfers
artifacts, captures screenshots and collects evidence.

This directory contains everything needed to reproduce the run. The current
numbers are preliminary: two independent 20-sample runs per target, not a
capacity-planning benchmark.

## Files

| File | Purpose |
| --- | --- |
| `main.rs` | Rust integration test. Ignored by default and gated behind `WINCENT_ADD_LATENCY=1` so it cannot mutate a developer host by accident. |
| `run_add_item_latency_vm.ps1` | Host orchestrator: builds the test, uploads it, starts the guest run, polls markers, captures screenshots, pulls evidence and validates restore state. |
| `vm_add_item_latency_start.ps1` | Guest launcher: cleans stale state, prefers the interactive session, and retains a scheduled-task fallback. |
| `vm_add_item_latency_guest.ps1` | Guest orchestrator (session 1): baseline backup, test execution, restore, markers, environment/final-state JSON, evidence zip. |
| `.env.example` | Configuration template. Copy to `.env` and fill in the VM paths and guest account. |
| `.env` | Local values, including the guest password. Git-ignored. |

## What is measured

Only the shell/COM `add_item` path is measured.

- **Recent Files**: path validation, Shell duplicate preflight, then
  `SHAddToRecentDocs(SHARD_PATHW, path)` on an STA thread.
- **Frequent Folders**: path validation, Shell visibility preflight, then the
  native Shell pin operation.
- **Timed interval**: one `manager.add_item(path)` call, from
  immediately before the call until it returns `Ok`/`Err`. Process startup,
  test framework, target creation, final raw verification and CFB restore are
  excluded.
- **Samples**: 2 warm-up calls per series (same code path, not recorded) and
  20 timed calls. Each call uses a brand-new target (`.txt` file for Recent, new
  directory for Frequent), so duplicate detection cannot short-circuit a sample.
- **Order**: series order rotates each round to avoid one series always running
  in the same cache/Explorer position.
- **Statistics**: mean, median, nearest-rank p95, min, max and sample standard
  deviation over successful calls.

After all samples the test parses both `.automaticDestinations-ms` files and
reports how many of its targets are present as raw DestList entries.

## Requirements

- Windows host with PowerShell 5.1 or later and a Rust toolchain with
  `cargo`.
- VMware Workstation with `vmrun.exe`.
- A Windows guest where the target user is logged in and `explorer.exe` runs in
  **session 1**. VMware Tools guest control must work with the configured
  account.
- The guest must not be used for unrelated work while the test runs, because
  the test mutates real Quick Access state.

## Configuration

```powershell
Copy-Item tests/add_item_latency/.env.example tests/add_item_latency/.env
# Edit .env: vmrun.exe, .vmx, guest user/password, evidence root, defaults.
```

The host script resolves settings as:

```text
command-line parameter > .env value > non-secret built-in default
```

`.env` keys:

| Key | Meaning |
| --- | --- |
| `WINCENT_VM_VMRUN` | Full path to `vmrun.exe` |
| `WINCENT_VM_VMX` | Full path to the target `.vmx` |
| `WINCENT_VM_PLATFORM` | `win10` or `win11`; selected by the host and verified in the guest |
| `WINCENT_VM_PASSWORD` | Optional VM encryption password passed through `vmrun -vp` |
| `WINCENT_VM_GUEST_USER` | Guest account for vmrun guest control |
| `WINCENT_VM_GUEST_PASSWORD` | Guest password |
| `WINCENT_EVIDENCE_ROOT` | Host evidence root; relative paths resolve from the repo root |
| `WINCENT_ADD_LATENCY_SERIES` | Default series, e.g. `recent,frequent` |
| `WINCENT_ADD_LATENCY_ITERS` | Default timed iterations per series |
| `WINCENT_ADD_LATENCY_WARMUP` | Default warm-up iterations per series |

`WINCENT_VM_VMX`, `WINCENT_VM_VMRUN`, `WINCENT_VM_GUEST_USER` and
`WINCENT_VM_GUEST_PASSWORD` have no hard-coded fallback. If they are missing,
the script stops before touching the VM and asks you to create `.env`.

The host prefers VMware `-interactive` to enter the Explorer session. The
guest launcher retains the interactive scheduled-task path as a fallback.

## Run

```powershell
# Recent + Frequent, 20 timed + 2 warm-up calls per series
pwsh -File tests/add_item_latency/run_add_item_latency_vm.ps1 `
    -Platform win10 `
    -EvidenceDir temp/track_wincent/evidence/my-add-latency-run `
    -Series 'recent,frequent' -Iters 20 -Warmup 2
```

The script builds the test with:

```text
cargo test --release --test add_item_latency --no-run
```

then runs the resulting binary in the guest with
`add_item_shell_latency --ignored --exact --nocapture`.

## Current preliminary results

Environment: dedicated Windows 10 Home 22H2 `10.0.19045` VM, Explorer in
session `1`. Two independent default runs (v2, v4-repeat), 20 samples each;
combined `n = 40` per target. Units are milliseconds.

| Target | n | mean | median | p95 | min | max | stdev |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Recent Files | 40 | **589.440** | 555.351 | 846.560 | 159.100 | 979.939 | 167.269 |
| Frequent Folders | 40 | **442.816** | 437.952 | 816.634 | 190.709 | 838.845 | 179.890 |

Per-run means:

| Run | Recent mean | Frequent mean |
| --- | ---: | ---: |
| v2 | 601.716 | 442.261 |
| v4-repeat | 577.163 | 443.371 |

Supplementary legacy explicit-refresh run (20 samples each, different batch):
Recent **584.025 ms**, Frequent **643.158 ms**. The Frequent refresh delta
(+~200 ms) is directional only because the Explorer window/load state was not
paired with the default runs.

Timing observations:

- This VM needs roughly **0.59 s** per Recent add and **0.44 s** per Frequent
  add, on average.
- Individual calls vary widely (stdev ~26%–46% of the mean); report median and
  p95 together with the mean.
- Frequent means repeated within ~0.2% across the two default runs; Recent
  means differed by ~4%.

## Verification and restore

For every run the guest orchestrator:

1. backs up both Quick Access CFB files and the complete Recent `.lnk` set;
2. writes a baseline manifest with per-file SHA-256;
3. runs the benchmark;
4. stops Explorer, writes back both CFB files, deletes test-created `.lnk`
   files, restores the baseline `.lnk` set, and restarts Explorer;
5. compares CFB SHA-256 values and the `.lnk` manifest with the baseline;
6. writes `final-state.json`, `FAILURE.txt`/`RESTORE_FAILURE.txt` on errors,
   and an evidence zip.

The host refuses to report success unless `restored = true`, no failure marker
exists, and the guest-reported executable SHA-256 matches the locally built
binary. Evidence is written under `temp/track_wincent/evidence/` and is
git-ignored.

## Limitations

- Results apply only to this VM image, Quick Access entry count, disk and
  Defender state.
- No ETW/Procmon decomposition, so the latency cannot be attributed to a
  specific internal Shell stage.
- 20 samples per run is enough for a preliminary mean/median/p95; p95 is close
  to the maximum sample (second-highest for n=20). Increase the sample count for
  tail-latency work.
- Warm-up is not optional. In a layout smoke test with zero warm-ups, the first
  cold Frequent pin hit the 30 s COM timeout. Keep at least 1–2 warm-ups; the
  official runs used 2.
