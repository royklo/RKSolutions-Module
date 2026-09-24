# Changelog

All notable changes to this project will be documented in this file.

The format follows [Conventional Commits](https://www.conventionalcommits.org/) and this project adheres to [Semantic Versioning](https://semver.org/). Release notes for each version are also generated from git history by the automation pipeline using the same conventional types (feat, fix, docs, refactor, test, etc.).



## [1.2.2] - 2026-09-24

Patch release for the Intune Anomalies report.

### Changes

- **Autopilot device preparation (v2) is recognised.** The *Not in Autopilot* tab (was *No Autopilot Hash*) only lists devices with neither a hardware hash nor a corporate identifier.
- **Cloud PCs are left out of the BitLocker and Autopilot tabs**, where they were always false positives.
- **Every tab explains what it checks.**
- **Noncompliant reasons use the names from the Intune admin center**, e.g. *Require BitLocker* or *Is active (no compliance check-in in the last 31 days)*.

### Fixes

- **Noncompliant reason no longer shows "Unknown"** for devices that fail the built-in Default Device Compliance Policy.
- **Application failure percentages are correct** (100% was shown as 1%).

---

## [1.2.1] - 2026-06-18

Patch release - fixes a perf regression and a missing-data symptom in the Intune Anomalies report's compliance fetch, plus a layout fix that affects every report.

### Fixes

- **Noncompliant Reason column now shows the actual rule name.** On tenants where Microsoft Graph returned a `deviceCompliancePolicyStates` record more than once for the same device (which happens when a single policy resolves via more than one assignment path), the rule-detail fetch silently failed and the Reason column fell back to "Unknown". `Get-IntuneAnomaliesReport` now dedupes those records before fetching, so the real rule name shows up (e.g. `Windows10CompliancePolicy.SignatureOutOfDate`).
- **Report no longer hangs ~60s on tenants with noncompliant devices.** When the rule-detail fetch failed, the Microsoft Graph `$batch` retry loop sat on a permanent 400 with exponential backoff. The retry path now fast-fails non-transient HTTP errors and logs Graph's actual error body so future regressions are easier to spot. End-to-end report runtime on the reproduction tenant dropped from ~70s to ~10s.
- **Rule-detail fetch has a per-request fallback.** If the batch path ever fails again (for any reason), the report retries each rejected sub-request individually instead of returning empty data. Slower but correct.

### Polish

- **Long values no longer overflow table cells.** UPNs, long device names, and other unbreakable strings now wrap inside their column instead of bleeding into the neighbouring column. Shared report template fix - applies to every report, not just the Intune Anomalies one.

---

## [1.2.0] - 2026-06-17

Adds BitLocker, Windows LAPS, and deprecated-settings detection to `Get-IntuneAnomaliesReport`, and significantly speeds the report up via Graph `$batch`.

### New detections

- **BitLocker key escrow** — flags devices where Intune assigned a BitLocker policy but no OS-volume recovery key is in Entra, plus devices that aren't encrypted at all. Resolves assignments across Settings Catalog, legacy device configurations, and Endpoint Security intents, honouring include / exclude groups and assignment filters.
- **Windows LAPS backup** — flags devices covered by an Entra-backed LAPS policy with no local admin credential backed up, or whose most recent backup is older than 60 days.
- **Deprecated Settings Catalog settings** — walks every Settings Catalog policy in the tenant and flags settings Microsoft has marked deprecated. Catalog data is cached.
- New `-ShowExcludedDevices` switch surfaces devices that were deliberately excluded from BitLocker / LAPS policies (via exclude group or assignment filter) as Info-severity rows. Off by default.

### Performance

End-to-end report runtime on a 75-policy / 2-device tenant dropped from ~19s to ~8s. Larger tenants see proportionally bigger wins.

The main lever is a new private Graph `$batch` helper used across five hot paths (compliance fetch, Settings Catalog policy `/settings`, app `?$expand=assignments`, group `transitiveMembers`, and the prior `$expand=settings` paginated fetch). Two quadratic per-row lookups in the BitLocker/LAPS and app-failure code paths were also replaced with hashtable indexes.

### Fixes

- DataTables didn't always initialise correctly in the rendered report — fixed.
- The two quadratic lookups above no longer slow down the report at higher device counts.

### Permissions

The Intune Anomalies report now requires two additional Graph scopes to check BitLocker and LAPS state:

- `BitlockerKey.ReadBasic.All`
- `DeviceLocalCredential.ReadBasic.All`

Both are the `ReadBasic` variants — the report can confirm that a recovery key or password backup exists, but never reads the actual recovery password or local admin password. `Connect-RKGraph` auto-detects missing scopes on existing tokens and re-grants consent for you.

### Polish

- BitLocker and LAPS tabs drop the Manufacturer, Model, and Policy Assigned columns (irrelevant to the security finding each row describes). Inventory-focused tabs keep them.
- Every stat tile in the report now uses a distinct color.

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
