# Findings

Track what was wrong, what was done to fix it, and how to test. Update this file when implementing fixes so future changes don't reintroduce the same issues.

The point of this file is to capture **patterns**, not history — git already has history. Add a row when you learn something the codebase can't tell you on its own (a non-obvious constraint, a subtle invariant, a workaround for a specific behavior).

## Summary of changes

| # | Area | What was wrong | What was fixed | Quick check |
|---|------|----------------|----------------|-------------|
| 1 | Get-AllDeviceData (compliance) | Per-device sequential GETs for `deviceCompliancePolicyStates` and `settingStates`. At 200 noncompliant devices × ~3 rules = ~800 round trips. | Two batched prefetch passes via `Invoke-RKGraphBatch`; main loop becomes hashtable lookup. | `Get-IntuneAnomaliesReport -Verbose` on a tenant with noncompliant devices: `[Get-AllDeviceData] Compliance policy-states batch` line shows total time |
| 2 | Get-BitLockerLapsAssignmentContext (Settings Catalog) | `configurationPolicies?$expand=settings` is hard-capped at ~1-2 policies per cosmos page. 75 policies → ~50 sequential page requests, ~10s. | Drop `$expand`; list policies bare, then batch-fetch `/settings` per policy (~4 parallel sub-requests of 20). | `-Verbose` shows `[BitLockerLaps] configurationPolicies list` and `Settings batch-fetch` separately |
| 3 | Resolve-IntuneLapsAnomalies (deviceName fallback) | When primary `azureAdDeviceId` lookup failed, code did `LapsCredentialByDeviceId.Values \| Where-Object { $_.deviceName -eq $d.DeviceName }`. O(N×M). | Pre-built `LapsCredentialByDeviceName` hashtable at index time; fallback is now O(1) lookup. | Eyeball the LAPS tab still shows credential rows for devices where the primary join misses |
| 4 | Get-ApplicationFailures | (a) `$apps \| Where-Object { $_.Id -eq ... }` ran twice per failed row. (b) One sequential `?$expand=assignments` GET per failed app. | (a) Pre-built id→app hashtable. (b) Collect unique app ids and issue them as a `$batch`. | `[Report] Get-ApplicationFailures` total in verbose output; on tenants with many failing apps, the saving is ~250ms per app |
| 5 | Get-BitLockerLapsAssignmentContext (group members) | Per referenced group: 2 sequential GETs (devices + users) on `transitiveMembers`. 10 groups = 20 sequential calls. | Batch all (group, devices\|users) requests in one `$batch`; only follow `@odata.nextLink` sequentially for groups >999 members (rare). | One `POST .../v1.0/$batch` for the whole set in verbose output |
| 6 | Intune Anomalies stat tiles | 10 tiles, 7 unique color tokens — `t-rust`, `t-amber`, `t-violet` each duplicated. | Reassigned BITLOCKER → `t-steel` (already defined but unused), added `t-indigo` for LAPS BACKUP, `t-pink` for DEPRECATED SETTINGS. | Open the HTML report; no two tiles share a color |
| 7 | BitLocker / LAPS tabs (columns) | Manufacturer + Model + Policy Assigned columns surfaced but redundant — Manufacturer/Model are device-inventory data, Policy Assigned is implied by Applied Policies. | Removed all three columns from headers, rows, dropdowns, JS filter wiring, and `Resolve-*` output objects. Column indices updated. | All 10 stat tiles render; no orphan filter dropdowns; columns: Customer, Device Name, Primary User, Serial Number, Encrypted, Applied Policies, Key Escrowed, Severity, Status |
| 8 | Dead code | `Get-GroupActivationDetails` in EntraAdminRoles.ps1 and `Invoke-RKSolutionsWithConnection.ps1` both defined but zero callers. | Removed (~115 lines). Load order in `RKSolutions.psm1` updated. | `Import-Module ./module/RKSolutions.psd1 -Force` still imports clean |

## Detailed findings

### Microsoft Graph `$expand=settings` is hard-capped on cosmos endpoints

When you query `deviceManagement/configurationPolicies?$expand=settings`, Graph silently reduces the page size to fit a payload limit. Adding `$top=999` does **not** override this — observed ~1-2 policies per page on a real tenant.

**Pattern**: never combine `$expand` with a large list endpoint when the expanded payload is heavy. Fetch the bare list, then `$batch` the per-item expand calls. Settings Catalog is the canonical example.

### Graph `$batch` is up to 20 sub-requests per HTTP round trip

`Invoke-RKGraphBatch` chunks larger sets automatically, retries 429/5xx (honoring `Retry-After`), and correlates responses back via caller-supplied `Id`. Use it any time you have N independent GETs to the same Graph instance — sub-requests run in parallel server-side.

**Don't use it when:**
- The work is paginated (the helper only batches first-page calls; if a sub-response has `@odata.nextLink` you must follow it sequentially for that one).
- The same call already returns the data in a single payload (e.g., `informationProtection/bitlocker/recoveryKeys` returns everything at once).

### Per-row `| Where-Object` against a list is the biggest accidental O(N²) trap

Two patterns to grep for and refuse:

```powershell
foreach ($row in $rows) {
    $match = $list | Where-Object { $_.Id -eq $row.Id }       # bad
}
```

```powershell
$entry = $hashtable.Values | Where-Object { $_.foo -eq $bar }  # bad
```

Both should be pre-indexed hashtables. The first pattern bit `Get-ApplicationFailures` (50 apps × 500-app catalog = 25,000 iterations) and `Get-AllDeviceData` (the original compliance section). The second bit `Resolve-IntuneLapsAnomalies` (devices × LAPS entries).

### Module load order matters

Private scripts are dot-sourced in two phases by `RKSolutions.psm1`:
1. `$sharedOrder` — utilities (`Connect-ToMgGraph`, `Invoke-GraphRequestWithPaging`, `Invoke-RKGraphBatch`, `Send-EmailWithAttachment`, etc.). These must load before anything that calls them.
2. `$domainOrder` — report-specific scripts. **`IntuneAnomalies.ps1` depends on helpers defined in `IntuneEnrollmentFlows.ps1`** (`Get-DetailedPolicyAssignments`, `Test-IntuneFilter`, `$script:AllFilters`), so flows must load first.

If you add a new private helper that's used cross-domain, register it in `$sharedOrder`, not `$domainOrder`.

### `$script:AllGroups` is populated by `Get-AllIntunePoliciesWithAssignments` only

`Get-GroupInfo` uses `$script:AllGroups` as a cache. When `Get-IntuneEnrollmentFlowsReport` runs first (which calls `Get-AllIntunePoliciesWithAssignments`), subsequent `Get-IntuneAnomaliesReport` runs in the same session get per-group lookup dedup for free. When the anomalies report runs cold, every referenced group hits Graph individually. If you ever batch this, do it inside `Get-BitLockerLapsAssignmentContext` (the consumer), not by refactoring the shared `Get-DetailedPolicyAssignments` — that helper is used by many other code paths.
