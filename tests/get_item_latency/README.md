# Quick Access `get_item` latency test (VM only)

Measures Quick Access list retrieval latency through two independent paths:

- **Shell query**: enumerate the live Quick Access Shell namespace through STA COM.
- **CFB parse**: read and parse the `.automaticDestinations-ms` backing file itself.

The benchmark runs inside a dedicated Windows VM. The host only orchestrates the
VM, transfers artifacts, captures screenshots and collects evidence.

The current numbers are preliminary: one official 20-sample run plus two repeat
runs, not a capacity-planning benchmark.

## Files

| File | Purpose |
| --- | --- |
| `main.rs` | Rust integration test. Ignored by default and gated behind `WINCENT_GET_ITEM_LATENCY=1` so it cannot mutate a developer host by accident. |
| `run_get_item_latency_vm.ps1` | Host orchestrator: builds the test, uploads it, starts the guest run, polls markers, captures screenshots, pulls evidence and validates the restore state. |
| `vm_get_item_latency_start.ps1` | Guest launcher (session 0): cleans stale state and creates a temporary `/IT` scheduled task in the interactive Explorer session. |
| `vm_get_item_latency_guest.ps1` | Guest orchestrator (session 1): baseline backup, test execution, restore, markers, environment/final-state JSON, evidence zip. |
| `.env.example` | Configuration template. Copy to `.env` and fill in the VM paths and guest account. |
| `.env` | Local values, including the guest password. Git-ignored. |

## Fixture

The test builds the following state before measuring, then restarts Explorer so
the Shell reloads the backing files. Fixture construction is not timed.

| List | Raw entries | Composition |
| --- | ---: | --- |
| Recent Files | 120 | 20 visible + 100 hidden |
| Frequent Folders | 100 | 46 pinned + 4 visible unpinned + 50 hidden unpinned |

Frequent Folders access counts cover `0`, `1`, `3` and `5`. Pinned entries use
`0`; visible unpinned entries use `3`; hidden unpinned entries include `1`, `3`
and `5`. The test asserts the raw counts, pinned count, rule-visible count and
access-count distribution after Explorer restarts.

The fixture starts from an empty v4 backing file and is created through the
crate's public mutation API; the test does not reuse the current user's Quick
Access state.

## Scenarios

| Scenario | Shell-visible count | Raw count |
| --- | ---: | ---: |
| `recent_full` | 20 (result window full) | 120 |
| `recent_under` | 8 (under the window) | 120 |
| `frequent_full` | 50 | 100 |
| `frequent_under` | 40 | 100 |

The Shell query is expected to return 20 Recent Files in the full window and 8
after 12 visible entries are hidden. Frequent Folders returns 46 pinned plus 4
unpinned entries in the full state, then 36 pinned plus 4 unpinned after 10
pinned entries are unpinned and hidden.

## What is measured

- **Shell query**: a fresh worker thread initializes STA COM, opens the Recent
  Items or Frequent Folders namespace, enumerates `FolderItems`, filters
  files/folders and collects paths. COM initialization and namespace access are
  inside the timed interval.
- **CFB parse**: `wincent::destlist::parse_file` opens the backing file, parses
  the CFB container and raw DestList entries, and counts `cfb_entries()`.
- **Timed interval**: one retrieval call, from immediately before the call until
  it returns. Fixture build, Explorer restarts, raw verification and restore are
  excluded.
- **Samples**: 2 warm-up calls per scenario/method, then 20 timed calls. Methods
  alternate their order each round so neither always runs first.
- **Statistics**: mean, median, nearest-rank p95, min, max and sample standard
  deviation over successful calls.

The host refuses to report success unless the expected shell counts, raw counts
and scenario statistics were written.

## Requirements

- Windows host with PowerShell 5.1 or later and a Rust toolchain with `cargo`.
- VMware Workstation with `vmrun.exe`.
- A Windows guest where the target user is logged in and `explorer.exe` runs in
  **session 1**. VMware Tools guest control must work with the configured
  account.
- The guest must not be used for unrelated work while the test runs. The test
  stops and restarts Explorer and replaces Quick Access backing files; use a
  dedicated VM or a snapshot.

## Configuration

```powershell
Copy-Item tests/get_item_latency/.env.example tests/get_item_latency/.env
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
| `WINCENT_VM_GUEST_USER` | Guest account for vmrun guest control |
| `WINCENT_VM_GUEST_PASSWORD` | Guest password |
| `WINCENT_EVIDENCE_ROOT` | Host evidence root; relative paths resolve from the repo root |
| `WINCENT_GET_ITEM_LATENCY_ITERS` | Default timed iterations per scenario/method |
| `WINCENT_GET_ITEM_LATENCY_WARMUP` | Default warm-up iterations per scenario/method |

`WINCENT_VM_VMX`, `WINCENT_VM_VMRUN`, `WINCENT_VM_GUEST_USER` and
`WINCENT_VM_GUEST_PASSWORD` have no hard-coded fallback. If they are missing,
the script stops before touching the VM and asks you to create `.env`.

## Run

```powershell
# 20 timed + 2 warm-up calls per scenario and method
pwsh -File tests/get_item_latency/run_get_item_latency_vm.ps1 `
    -EvidenceDir temp/track_wincent/evidence/my-get-item-latency-run `
    -Iters 20 -Warmup 2
```

The script builds the test with:

```text
cargo test --release --test get_item_latency --no-run
```

then runs the resulting binary in the guest with:

```text
get_item_latency --ignored --exact --nocapture
```

## Current preliminary results

Environment: dedicated Windows 10 Home 22H2 `10.0.19045` VM, Explorer in
session `1`. One official run with 20 samples per scenario and method. Units
are milliseconds.

| Scenario | Shell mean | Shell median | Shell p95 | CFB mean | CFB median | CFB p95 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `recent_full` | 87.629 | 68.063 | 241.312 | 1.549 | 1.366 | 2.257 |
| `recent_under` | 52.270 | 44.982 | 78.077 | 1.882 | 1.631 | 2.518 |
| `frequent_full` | 56.933 | 54.477 | 78.596 | 1.400 | 1.298 | 1.990 |
| `frequent_under` | 43.368 | 37.758 | 63.126 | 1.569 | 1.401 | 2.339 |

Aggregate median over the official and two repeat runs (`n = 60` per
scenario/method):

| Scenario | Shell median | CFB median |
| --- | ---: | ---: |
| `recent_full` | 68.630 | 1.468 |
| `recent_under` | 44.510 | 1.530 |
| `frequent_full` | 52.269 | 1.353 |
| `frequent_under` | 37.048 | 1.340 |

Timing observations:

- CFB parsing is roughly one to two orders of magnitude faster than the live
  Shell query in every scenario (median around 1.3–1.5 ms vs 37–69 ms).
- Both lists are slower when the result window is full: 20 vs 8 Recent Files and
  50 vs 40 Frequent Folders.
- Shell queries occasionally have long tails (observed 0.2–1.2 s samples), so
  the mean is inflated in some runs. Use median and p95 for comparisons.

## Verification and restore

For every run the guest orchestrator:

1. backs up both Quick Access CFB files and the complete Recent `.lnk` set;
2. writes a baseline manifest with per-file SHA-256;
3. runs the fixture-building and measuring test;
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

- Results apply only to this VM image, fixture composition, Quick Access entry
  count, disk and Defender state.
- No ETW/Procmon decomposition, so the latency cannot be attributed to a
  specific internal Shell or CFB parser stage.
- 20 samples per run is enough for a preliminary mean/median/p95; p95 is close
  to the maximum sample (second-highest for n=20). Increase the sample count for
  tail-latency work.
- The test forcibly restarts Explorer, so it must run on a dedicated VM. An
  interrupted run may leave Quick Access modified; restore from the baseline
  copies or a VM snapshot.
- Warm-up is not optional: the first queries after an Explorer restart can be
  much slower than steady state. The official run used 2 warm-ups per
  scenario/method.
