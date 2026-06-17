# Changelog

All notable changes to this project will be documented in this file.

The format follows [Conventional Commits](https://www.conventionalcommits.org/) and this project adheres to [Semantic Versioning](https://semver.org/). Release notes for each version are also generated from git history by the automation pipeline using the same conventional types (feat, fix, docs, refactor, test, etc.).

## [Unreleased]

### Features

- Intune Anomalies Report gains two new checks: **BitLocker key escrow** and **Windows LAPS backup**. Each check resolves Intune policy assignments end-to-end (Settings Catalog by `settingDefinitionId`, legacy `windows10EndpointProtectionConfiguration`, Endpoint Security intents) and honours include / exclude groups, All Users / All Devices targets, and assignment filters - so the report flags both "no policy applied" and "policy applied but no key / no backup in Entra".
- A single bulk pass (`Get-BitLockerLapsAssignmentContext`) collects policies + assignments + filters + transitive group members + `informationProtection/bitlocker/recoveryKeys` + `directory/deviceLocalCredentials`; per-device evaluators (`Resolve-IntuneBitLockerAnomalies`, `Resolve-IntuneLapsAnomalies`) then emit severity-tagged anomaly rows.
- Two new dashboard tiles, tabs, and DataTables panels in the Intune Anomalies HTML report.

### Maintenance

- Expose `AzureAdDeviceId` on the device records produced by `Get-AllDeviceData` (required to join managed devices with Entra recovery keys and LAPS local credentials).
- Add `BitlockerKey.ReadBasic.All` and `DeviceLocalCredential.ReadBasic.All` to the default Connect-RKGraph scopes and to `docs/PERMISSIONS.md`. ReadBasic variants only - the report never reads recovery passwords or local admin passwords.

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
