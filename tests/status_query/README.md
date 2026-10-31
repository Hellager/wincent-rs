# Status Query VM Validation

Validates `QuickAccessManager::quick_access_status()`, `recent_file_position()`,
and `frequent_folder_slot()` against Explorer Shell query results and measures
their query cost.

## Run

```powershell
pwsh -File tests/status_query/run_status_query_vm.ps1 `
    -Platform win10 `
    -EvidenceDir temp/track_wincent/evidence/status-query-<date>-v1
```

Copy `.env.example` to `.env` in this directory and fill in the VM settings.
The test-local `.env` file is ignored by Git. Pass `-NoMutate` to skip the
add/remove round trip.
Encrypted VMs can set the optional `WINCENT_VM_PASSWORD`. The host prefers
VMware `-interactive` for the Explorer session and keeps the scheduled-task
path as a fallback.

## Behavior

1. Backs up the Recent/Frequent CFB files and the Recent `.lnk` set.
2. Creates platform-specific fixtures: Windows 10 uses 25 Recent files and two
   direct unpinned Frequent folders; Windows 11 uses 25 Recent files and two
   system-pinned Frequent folders without invoking the v4 direct writer.
3. Checks positions, pin state, entries outside the result windows, and path
   variants.
4. Runs add-to-present, remove-to-absent, pin-to-present, and unpin-to-absent
   round trips.
5. Measures parse, metadata, and query timings.
6. Restores each CFB to its original absent/zero-length/populated state in
   `finally`, then compares the link manifest.

Set `-Platform win10|win11` or `WINCENT_VM_PLATFORM` in `.env`. The guest checks
the actual Windows build and stops before modifying Quick Access if it does not
match the requested platform.

Windows 11 Recent tracking is capability-based: if system `add_item` calls do
not persist Recent entries, Recent mutation checks are explicitly skipped and
reported in `fixture_capabilities`; native Frequent pin/remove checks still run.

## Evidence

The run produces `result.json`, `timings.csv`, `environment.json`,
`final-state.json`, `benchmark.log`, screenshots, and
`status-query-evidence.zip`.
