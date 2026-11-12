# QuickAccessManager VM acceptance suite

End-to-end validation of the **current manager and standalone APIs** on a dedicated
Windows 10 or Windows 11 VMware guest. This is a correctness suite, not a
latency benchmark. See [README.cn.md](README.cn.md) for Chinese instructions.

## Safety and execution

The suite **clears Quick Access history and user pins**, changes three visibility
settings, stops/restarts Explorer, and creates/removes native Shell entries.
Never run its ignored tests directly on your workstation.

Requirements: VMware Workstation with working snapshots and guest tools, an
interactive logged-in test user with Explorer, host Rust/Cargo, and host
PowerShell 7. Guest scripts support Windows PowerShell 5.1. Guest credentials
must permit VMware guest operations and restarting that user's Explorer.

1. Copy `.env.example` to `.env` and set the VM path, `vmrun` path, guest
   credentials and `WINCENT_VM_PLATFORM=win10` or `win11`. `.env` is git-ignored.
   Process environment variables and explicit script parameters also work.
2. Log the test user into the guest desktop. Close unrelated guest applications
   that may write Recent history; do not interact with the guest during the run.
3. Run from the repository root:

   ```powershell
   pwsh -File tests/quick_access_manager/run_quick_access_manager_vm.ps1 `
       -Platform win10
   ```

   `-Platform win11` runs the same native fixtures on a Windows 11 guest.
   `-EvidenceDir <new-directory>` selects the local evidence directory.
   Encrypted VMs can provide `WINCENT_VM_PASSWORD` or `-VmPassword`.

Host execution is limited to compilation (`cargo test --no-run`), VMware
orchestration, screenshots and evidence transfer. There is **no host test
execution** and no skip-build option that could select a stale executable.
No Cargo manifest changes or additional dependencies are required.

The runner creates a unique whole-VM snapshot **before guest changes**. Snapshot
creation failure stops execution. A unique guest directory and run ID prevent
stale markers from another run being accepted. The guest verifies its VMware
identity, actual OS build, executable SHA-256, username and nonzero Explorer
session before any Quick Access mutation. The Rust tests are ignored by default
and additionally require the guest acknowledgement and completed-backup marker.

The guest backs up the **entire Recent tree**, including AutomaticDestinations,
CustomDestinations, shortcuts, absent/empty files and wincent gate files. The
three registry values are backed up with existence, type and value. It restores
these in `finally` and compares recursive path/length/SHA-256 manifests and
registry values while Explorer is stopped. Post-restart state is also recorded:
Explorer may normalize an uninitialized CFB, so that state is not falsely
reported as byte-identical.

The host waits for cleanup completion, even if a failure marker appears first,
collects evidence, and **reverts the whole-VM snapshot on success or failure**.
The original power state is preserved. Successful runs delete only their own
temporary snapshot; failed runs retain theirs for recovery. An unavailable
rollback makes the host run fail. Do not interrupt the host process: `finally`
cannot run after a host crash or forced termination. Keep the snapshot and
manually recover the VM in that case.

## Coverage

| Area | Checks |
| --- | --- |
| Construction | `new`, `default`, `builder`, builder default/clone, `build`, `timeout`, `try_timeout`; zero rejected/panics as documented |
| Visibility | `set_visible` / `is_visible` hide/show for all four targets, mixed `All`, `All` excludes Start; Start tracking explicitly enabled before fixtures |
| Query | All three categories, `get_items` / `get_item_paths` order equivalence, Recent-first deduplicated `All`, keyword case/empty/no-match behavior |
| Independent oracle | Real `Shell.Application` COM enumeration in its own STA, compared in order with manager results; not a second call to `get_items` |
| Status | Every raw persisted entry, visible position, filtered/out-of-window entries, Frequent pin state, default missing status and `lists` |
| Paths | Native files/folders with spaces and Unicode, uppercase/slashes/canonical paths, directory trailing slash/dot, exact membership versus substrings |
| Errors | Empty/missing paths, existing-but-absent file/folder, duplicate file/folder and category details |
| Native mutation | 25 Recent files, two pinned folders, native unpinned parent, add/remove convergence, pinned and unpinned removal |
| Locks | Recent/Frequent/All, clone-shared lock ownership, repeated lock/unlock errors, independent manager and child-process mutation denial, backing-file write denial, queries while locked |
| Lock lifetime | Disjoint categories, clone lifetime, plain unlock preserving `.lnk`, cleanup success/failure reports, and lock release after failed deletion |
| Refresh | Open Quick Access/Home window, populated and cleared state; checks API result, not pixels |
| Clear | Recent/Frequent/All seeded with both file and folder, scoped preservation, repeated clear; does not incorrectly require Windows default pins to disappear |
| Standalone functions | Query parity with manager and independent Shell results, explicit timeouts, zero-timeout rejection, path types, pin status, pinning existing unpinned folders, both removal routes, scoped clear and refresh |
| Standalone locks | Typed guards, disjoint locks, overlapping acquisition, manager interoperability, cross-process mutation denial, plain unlock, explicit cleanup, delete-sharing failure and release on drop |
| Non-DestList unit tests | Dynamically discovers and runs every library test outside `destlist::`, covering pure logic, timeout handling, concurrency resource lifetime and error categories |

Every phase uses bounded convergence polling after native mutation. A failed
prerequisite aborts its dependent phases and records `suite.completed=false`;
unavailable Recent tracking is **not silently skipped or counted as a pass**.
Mutation timeouts are not retried. A guest watchdog terminates the run before
restoration; the host snapshot is the final protection from delayed Shell work.
The first three Recent files are added 66 seconds apart to create a native,
visible unpinned parent folder. Rapid repeated visits coalesce into one usage
point, so adding all 25 files immediately is insufficient for this fixture.
These two waits prepare test data; they are not performance measurements.
Allow several minutes for a complete run, including snapshot creation and rollback.
The standalone phase prepares two additional unpinned parents using the same
66-second visit intervals. It runs after manager coverage, with fresh paths.
Before the acceptance suite, the guest lists the library test binary, excludes
`destlist::`, and runs every remaining test by exact name. The discovered names
are saved in `unit-tests-list.txt`, so newly added modules cannot be missed by a
stale hard-coded filter. These tests are never executed on the host.
Each clear case uses fresh paths: removing a Recent item can leave a blocked
entry in Shell's cached state, so re-adding the previous mutation target is not
a reliable way to prepare the next case.

## Evidence and acceptance

Default output: `temp/track_wincent/evidence/qam-<platform>-<run-id>/`.

- `result.json`: per-check IDs, results/details, count, completion and process exit.
- `checks.tsv`, `test.log`: streaming checks and Rust test output, including failures.
- `unit-tests.log`, `unit-tests-list.txt`: non-DestList unit output and the exact
  discovered test set; `result.json` records its positive passing count as
  `unit_checks` and success as `unit_passed`.
- `environment.json`: actual build/UBR, display version, user/session, binary hash,
  source commit, Explorer / Shell DLL versions and hashes.
- `final-state.json`: guest restoration, target removal and remaining test processes.
- `host-result.json`: whole-VM snapshot rollback and **overall pass/fail**.
- `archive-verification.json`: every archived original file, including hidden
  files, verified against the baseline length and SHA-256 manifest.
- `evidence.zip`: guest baseline, restored manifests and registry values, fixture /
  post-clear CFB files, raw records, independent Shell results and lock-child logs.
- `suite-source/`, `source.diff`, `source-status.txt`: exact suite sources and
  local source changes, without `.env` credentials.
- `library-source/`: library sources and Cargo manifests, including new modules
  not yet tracked by Git. Both uploaded executables have verified SHA-256 hashes.
- `guest-running.png`: best-effort screenshot; not a UI correctness assertion.

Accept a build only when `result.json` is completed with zero failures and a
positive check count, `unit_passed=true` with `unit_checks>0`, guest restoration
is successful, both executable hashes match, and `host-result.json` reports `passed=true` and
`snapshot_reverted=true`. Missing reports are failures, not skipped tests.

## Scope and limitations

This suite validates normal behavior and error contracts of the public
`QuickAccessManager` and `quick_access` APIs. Results apply only to the library revision and Windows
build actually tested. Run the suite again on each target system version and
retain its evidence; a passing run does not establish compatibility with other
or future versions.

The test compares the library's visibility limits, ordering and filtering rules
with the actual Shell on the tested build. It uses native fixtures rather than
injecting another Windows version's persisted data or patching private APIs.
An unsupported format or changed Shell behavior must fail visibly.

Out of scope: exhaustive low-level DestList API coverage, corruption/CRC recovery
matrices, induced `PartialEmpty` failures, forced mid-COM crashes, every
third-party concurrent writer, policy-managed visibility, network paths, private
ABI stability, and pixel-level UI rendering. No subsecond timeout assertion is
used: timing is nondeterministic and a timeout cannot establish that a native
mutation was cancelled.
