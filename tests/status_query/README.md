# Status Query VM Validation

Validates `QuickAccessManager::quick_access_status()`, `recent_file_position()`,
and `frequent_folder_slot()` against Explorer Shell query results and measures
their query cost.

## Run

```powershell
pwsh -File tests/status_query/run_status_query_vm.ps1 `
    -EvidenceDir temp/track_wincent/evidence/status-query-<date>-v1
```

Copy `.env.example` to `.env` in this directory and fill in the VM settings.
The test-local `.env` file is ignored by Git. Pass `-NoMutate` to skip the
add/remove round trip.

## Behavior

1. Backs up the Recent/Frequent CFB files and the Recent `.lnk` set.
2. Creates fixtures with 25 Recent files through `add_item` and two unpinned
   Frequent folders through `add_item_direct`.
3. Checks positions, pin state, entries outside the result windows, and path
   variants.
4. Runs add-to-present, remove-to-absent, pin-to-present, and unpin-to-absent
   round trips.
5. Measures parse, metadata, and query timings.
6. Restores both CFB files and the `.lnk` set in `finally`, then compares the
   SHA-256 hashes and link manifests.

## Evidence

The run produces `result.json`, `timings.csv`, `environment.json`,
`final-state.json`, `benchmark.log`, screenshots, and
`status-query-evidence.zip`.
