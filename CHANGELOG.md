# Changelog

All notable changes to this project will be documented in this file.

The format follows [Conventional Commits](https://www.conventionalcommits.org/) and this project adheres to [Semantic Versioning](https://semver.org/). Release notes for each version are also generated from git history by the automation pipeline using the same conventional types (feat, fix, docs, refactor, test, etc.).

## [Unreleased]

### Features

- Intune Anomalies Report gains two new checks: **BitLocker key escrow** and **Windows LAPS backup**. Each check resolves Intune policy assignments end-to-end (Settings Catalog by `settingDefinitionId`, legacy `windows10EndpointProtectionConfiguration`, Endpoint Security intents) and honours include / exclude groups, All Users / All Devices targets, and assignment filters - so the report flags both "no policy applied" and "policy applied but no key / no backup in Entra".
- A single bulk pass (`Get-BitLockerLapsAssignmentContext`) collects policies + assignments + filters + transitive group members + `informationProtection/bitlocker/recoveryKeys` + `directory/deviceLocalCredentials`; per-device evaluators (`Resolve-IntuneBitLockerAnomalies`, `Resolve-IntuneLapsAnomalies`) then emit severity-tagged anomaly rows.
- Two new dashboard tiles, tabs, and DataTables panels in the Intune Anomalies HTML report.
- **New `Get-IntuneAnomaliesReport -ShowExcludedDevices` switch.** Surfaces devices that were *deliberately* excluded from BitLocker / LAPS policies (via exclude group or assignment filter) as Info-severity rows. Off by default so the report stays focused on actual anomalies.
- **Deprecated Settings Catalog detection** in the Intune Anomalies report, backed by a cached download from [royklo/IntuneSettingsCatalogData](https://github.com/royklo/IntuneSettingsCatalogData) so the catalog never gets fetched twice in a session.
- **Auto-repair flow in `Connect-RKGraph`** for tokens missing newly-added scopes. Clears the MSAL token cache and re-grants consent automatically instead of failing the report with a permission error.
- **`Get-RKSolutionsConsent` helper script** (in `~/Downloads/Grant-RKSolutionsConsent.ps1`) for first-run admin consent without running the module.
- **Search bar in every DataTables panel** of the shared template, injected by `initRKTable` into the existing `.rk-filter-bar`. Works alongside existing per-column filters.

### Performance

New private helper `Invoke-RKGraphBatch` (Microsoft Graph `$batch` wrapper with chunking to 20 sub-requests, 429/5xx retry honoring `Retry-After`, and id-based response correlation) plus five hot paths converted to use it. Measured on a 75-policy / 2-device tenant: **19.0s → 8.2s end-to-end (–57%)**. Scaling wins at larger tenant sizes are proportionally bigger.

- **Compliance fetch** in `Get-AllDeviceData`: two batched prefetch passes (policy states, then per-rule setting states) replace 1+N×M sequential per-device GETs. At 1000 devices with 5% noncompliant × 3 rules each: ~200 round trips → ~10 batches.
- **Settings Catalog policy discovery** in `Get-BitLockerLapsAssignmentContext`: drop `?$expand=settings` (Graph hard-caps page size to ~1-2 policies per cosmos page with `$expand`) and instead list bare policies, then `$batch` `/settings` per policy. On a 75-policy tenant: 9.9s → 1.3s.
- **App assignment expand** in `Get-ApplicationFailures`: one batched `$expand=assignments` round for all failed apps instead of N sequential GETs. At 50 failing apps: ~12s → ~2s projected.
- **Per-group `transitiveMembers`** in `Get-BitLockerLapsAssignmentContext`: device-members + user-members for all referenced groups in one batch, with sequential `@odata.nextLink` follow-up only for groups exceeding 999 members.
- **Verbose stopwatch instrumentation** across `Get-IntuneAnomaliesReport` and `Get-AllDeviceData` so subsequent perf work can be measured instead of guessed.

### Fixes

- **O(N×M) lookup in `Resolve-IntuneLapsAnomalies`.** The deviceName fallback enumerated all LAPS credentials for every device whose primary `azureAdDeviceId` join failed. At 1000 devices × 1000 LAPS entries: ~1M iterations per report. Now an O(1) hashtable lookup via a parallel `LapsCredentialByDeviceName` index built at fetch time.
- **O(N×M) lookup in `Get-ApplicationFailures`.** Per-row `$apps | Where-Object { $_.Id -eq ... }` ran twice per failed row (once for the if-check, once inline for `displayName`). Replaced with an id-keyed hashtable.
- **JavaScript template rendering bug** in the report HTML. A missing backtick escape in `select.append($('<option>')...)` inside a PowerShell here-string broke the DataTables initialisation. Fixed.

### UI

- **BitLocker and LAPS tabs trim three columns** (Manufacturer, Model, Policy Assigned). Manufacturer and Model are device-inventory data that distract from the security finding each row describes; Policy Assigned is redundant with Applied Policies. Other inventory-focused tabs (Devices Without Autopilot Hash, Inactive Devices, Disabled Primary Users) keep these columns.
- **"Not Encrypted" tab merged into BitLocker tab.** Unencrypted devices now surface there as `Device not encrypted and no BitLocker policy assigned` (Critical) or `BitLocker policy assigned but device not encrypted` (Critical) rows.
- **Every stat tile in the Intune Anomalies report now uses a unique color.** Three tokens were duplicated previously (rust, amber, violet each used twice across 10 tiles). Reassigned BITLOCKER → `t-steel` (already defined but unused), added new `t-indigo` for LAPS BACKUP and `t-pink` for DEPRECATED SETTINGS. Light + dark theme parity preserved.

### Maintenance

- Expose `AzureAdDeviceId` on the device records produced by `Get-AllDeviceData` (required to join managed devices with Entra recovery keys and LAPS local credentials).
- Add `BitlockerKey.ReadBasic.All` and `DeviceLocalCredential.ReadBasic.All` to the default `Connect-RKGraph` scopes and to `docs/PERMISSIONS.md`. ReadBasic variants only - the report never reads recovery passwords or local admin passwords.
- **Remove ~115 lines of dead code.** `Get-GroupActivationDetails` in `EntraAdminRoles.ps1` and the whole `Invoke-RKSolutionsWithConnection.ps1` helper file (load-order entry too) - defined but no callers.
- **Six new Pester smoke tests** (all CI-safe, no Graph connection required): report-cmdlet shape contracts, BitLocker / LAPS anomaly resolver column contracts (would have caught this PR's column-trim regression at the data layer), and a scope-drift check between `Connect-RKGraph` defaults and `docs/PERMISSIONS.md`.

---

## [1.1.0]

### Features

- New `Get-CustomSecurityAttributesReport` cmdlet with auto-discovery of attribute sets across users, devices, and enterprise applications.
- Shared HTML report template (`Get-RKSolutionsReportTemplate`) with rksolutions.nl branding, Geist/Geist Mono typography, pill-style tabs, and neutral dark theme.
- All 5 reports migrated to shared template, eliminating duplicated HTML/CSS/JS.
- Light/dark theme support for table backgrounds.
- DataTable column widths recalculate on tab switch.

### Security

- HTML-encode all Graph API data before HTML interpolation across all report generators (stored XSS prevention).
- Fix JavaScript filter dropdown injection — use jQuery DOM API instead of string concatenation.
- Validate `-From` parameter as email address or GUID in `Send-EmailWithAttachment`.

### Fixes

- Replace quadratic `$array +=` patterns with `List[PSObject].Add()` in M365License and CustomSecurityAttributes.
- Add `Write-Verbose` to empty catch blocks in IntuneEnrollmentFlows for diagnostics.
- Add `Write-Warning` when Graph paging silently caps results at 10,000 items.
- Fix report file deletion before confirming email was sent in `Get-IntuneAnomaliesReport`.
- Guard undefined `$mermaidDiagram` variable with null-check.
- Fix typo `OperationSystemEdtionOverview` to `OperatingSystemEditionOverview`.
- Fix double-encoded `&rarr;` HTML entity in PIM Audit Logs.
- Sanitize tenant name in export file paths.
- Use unique temp file instead of hardcoded `C:\temp` path.
- Initialize `$emailSent` before conditional block.
- Remove dead OS detection variables in `Export-Results`.
- PIM Audit Logs: fix column widths with `table-layout:fixed`.
- M365 License: replace assignment type color badges with plain text.

### Maintenance

- Remove unused `$Filters` and `$AttributeNames` parameters from `Get-CustomSecurityAttributeData`.
- Add `CustomSecAttributeAssignment.Read.All` and `CustomSecAttributeDefinition.Read.All` scopes.

---

## [1.0.0] - (initial)

### Features

- Initial release of the RKSolutions PowerShell module.
- Cmdlets: Connect-RKGraph, Disconnect-RKGraph, Get-IntuneEnrollmentFlowsReport, Get-IntuneAnomaliesReport, Get-EntraAdminRolesReport, Get-M365LicenseAssignmentReport, Get-DeviceEvaluationContext, Get-CloudPCProvisioningPolicyGroupInfo.
- Connects to Microsoft Graph and generates HTML/CSV reports for Intune enrollment flows, anomalies, Entra admin roles, and M365 license assignment.
