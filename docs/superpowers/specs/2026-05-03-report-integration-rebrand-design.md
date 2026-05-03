# Design: Custom Security Attributes Report Integration + Brand Rebrand

**Date:** 2026-05-03
**Author:** Roy Klooster + Claude

## Overview

Two tasks in one spec:

1. **Integrate** `Get-CustomSecurityAttributesReport.ps1` (source: `/Users/roy/github/Get-CustomSecurityAttributesReport.ps1`) into the RKSolutions module as `Get-CustomSecurityAttributesReport`
2. **Rebrand** all 5 HTML reports (4 existing + 1 new) with the rksolutions.nl editorial parchment brand, using a shared template function

### Delivery Order

Task 1 (integration) and Task 2 (rebrand) are delivered together since the new report will be built on the shared template from the start. However, commits are structured incrementally:

1. Create shared template (`New-RKSolutionsReportTemplate.ps1`)
2. Integrate the new Custom Security Attributes report (built on template)
3. Migrate each existing report to the shared template (one commit per report)
4. Update manifest, PSM1, tests, permissions docs

## Task 1: Custom Security Attributes Report Integration

### New Files

| File | Purpose |
|------|---------|
| `module/Public/Get-CustomSecurityAttributesReport.ps1` | Public cmdlet (thin wrapper) |
| `module/Private/CustomSecurityAttributes.ps1` | Data fetching + HTML report generation |

### Public Cmdlet Parameters

Follow the existing module pattern (same as Get-IntuneAnomaliesReport, Get-EntraAdminRolesReport). The cmdlet must include a `.SYNOPSIS` comment-based help block matching the pattern in existing public cmdlets.

```powershell
function Get-CustomSecurityAttributesReport {
    [CmdletBinding()]
    param(
        # Domain-specific parameters (unique to this report)
        [string]   $AttributeSet = "CustomerData",
        [string[]] $AttributeNames,
        [hashtable]$Filters,
        # Common report parameters (shared across all report cmdlets)
        [switch]   $SendEmail,
        [string[]] $Recipient,
        [string]   $From,
        [string]   $ExportPath,
        [switch]   $DebugMode
    )
}
```

**Note:** `$AttributeSet`, `$AttributeNames`, and `$Filters` are domain-specific additions required by the Custom Security Attributes API. They follow the same principle as `Get-IntuneEnrollmentFlowsReport`'s domain-specific `$Device`, `$AssignmentOverviewOnly`, and `$ApplyPlatformFilter` parameters. The `$ExportCSV` switch from the standalone script is removed to maintain the common parameter contract.

**Differences from standalone script:**
- No auth parameters on the cmdlet. User calls `Connect-RKGraph` first (module pattern).
- No `Install-Requirements` or `Connect-ToMgGraph` calls inline. Uses existing Graph context.
- Script's `Connect-ToMgGraph` helper is already in `module/Private/Connect-ToMgGraph.ps1`.
- Uses `Send-EmailWithAttachment` from `module/Private/Send-EmailWithAttachment.ps1`.

### Required Scopes

The standalone script requires: `CustomSecAttributeAssignment.Read.All`, `User.Read.All`, `Organization.Read.All`, `Mail.Send`.

`User.Read.All`, `Organization.Read.All`, and `Mail.Send` are already in `Connect-RKGraph` default scopes.

**Action:** Add `CustomSecAttributeAssignment.Read.All` to `Connect-RKGraph`'s default `$RequiredScopes` array.

### Private Helper Functions

`CustomSecurityAttributes.ps1` will contain:

- `Get-CustomSecurityAttributeData` -- fetches attribute definitions and user data from Graph API
- `New-CustomSecurityAttributesHTMLReport` -- generates HTML using the shared template

### Data Flow

```
User calls Get-CustomSecurityAttributesReport
  -> Verify Graph context (Get-MgContext)
  -> Get tenant info (Organization.Read.All)
  -> Get attribute definitions (beta/directory/customSecurityAttributeDefinitions)
  -> Query users with attributes (v1.0/users with $filter + ConsistencyLevel=eventual)
  -> Build report data array
  -> Call New-CustomSecurityAttributesHTMLReport (shared template)
  -> Optional: Send-EmailWithAttachment
```

### Module Loader Updates (RKSolutions.psm1)

The PSM1 has two hardcoded script loading lists that must be updated:

1. **`$sharedOrder` array (line 27):** Add `'New-RKSolutionsReportTemplate.ps1'` to load the shared template before any report-specific script.
2. **`$domainOrder` array (line 41):** Add `'CustomSecurityAttributes.ps1'` to the report-specific loading list.
3. **`Export-ModuleMember` call (line 55):** Add `'Get-CustomSecurityAttributesReport'` to the exported functions list.

### Manifest Updates

- Add `Get-CustomSecurityAttributesReport` to `FunctionsToExport` in `RKSolutions.psd1`
- Update description to mention Custom Security Attributes
- Bump version to `1.1.0`
- Correct `PowerShellVersion` from `'5.1'` to `'7.0'` (the PSM1 already enforces 7.0 at runtime; the manifest should match)
- Add `1.1.0` entry to `ReleaseNotes`

### Test Updates

- Add `Get-CustomSecurityAttributesReport` to consistency tests
- Add expected parameters: `AttributeSet`, `ExportPath`
- The new public cmdlet must have a `.SYNOPSIS` help block (validated by existing "Get-Help is filled" test)

## Task 2: Brand Rebrand with Shared Template

### Architecture: Shared Template Function

Create `module/Private/New-RKSolutionsReportTemplate.ps1` containing a single function:

```powershell
function New-RKSolutionsReportTemplate {
    param(
        [string]$TenantName,
        [string]$ReportTitle,        # e.g. "Anomalies"
        [string]$ReportSlug,         # e.g. "intune-anomalies" (for breadcrumb)
        [string]$Eyebrow,            # e.g. "INTUNE ANOMALIES"
        [string]$Lede,               # italic subtitle
        [string]$StatsCardsHtml,     # pre-built stat cards HTML
        [string]$BodyContentHtml,    # tabs, tables, filters -- everything unique
        [string]$CustomCss,          # report-specific CSS (e.g. shimmer animations)
        [string]$ReportDate,
        [string[]]$Tags              # optional tag pills
    )
}
```

This function returns the complete HTML document string using PowerShell expandable here-strings (`@"..."@`) with direct variable interpolation. No `{{placeholder}}` string replacement -- parameters are embedded directly in the here-string via `$StatsCardsHtml`, `$BodyContentHtml`, etc. Each report only needs to provide its unique content.

### Report Template Parameter Values

| Report | ReportSlug | Eyebrow | ReportTitle accent word | Lede |
|--------|-----------|---------|------------------------|------|
| Intune Anomalies | `intune-anomalies` | `INTUNE ANOMALIES` | `Anomalies` | Device compliance overview with flagged anomalies across encryption, activity, and application health. |
| Entra Admin Roles | `entra-admin-roles` | `ENTRA ADMIN ROLES` | `Admin Roles` | Privileged role assignments including permanent, eligible, group-based, and service principal assignments. |
| M365 License Assignment | `m365-licenses` | `M365 LICENSE ASSIGNMENT` | `License` | License assignment overview including direct, inherited, and disabled user assignments. |
| Intune Enrollment Flows (Overview) | `intune-enrollment-overview` | `INTUNE ENROLLMENT FLOWS` | `Enrollment` | Assignment overview across all Intune policy types with group and filter targeting. |
| Intune Enrollment Flows (Device) | `intune-enrollment-device` | `DEVICE ENROLLMENT FLOW` | `Enrollment` | Device-specific enrollment flow with policy assignments, group membership, and filter evaluation. |
| Custom Security Attributes | `custom-security-attributes` | `CUSTOM SECURITY ATTRIBUTES` | `Security Attributes` | Users with custom security attribute assignments across the specified attribute set. |

### What the Template Provides (single source of truth)

1. `<!DOCTYPE html>` + `<head>` with meta, Google Fonts, CDN links
2. CSS custom properties (`:root` light tokens, `[data-theme="dark"]` dark tokens)
3. All shared CSS: typography, breadcrumb pill, eyebrow, title block, stat tiles, filter bar, table styles, badges, tag pills, footer, theme toggle, DataTables overrides
4. `$CustomCss` injection point in `<style>` block for report-specific styles
5. Theme toggle switch (fixed top-right)
6. Breadcrumb pill: `<- cd ./reports/{slug}`
7. Eyebrow + Playfair title + lede
8. Stats cards area (from `$StatsCardsHtml`)
9. Body content area (from `$BodyContentHtml`)
10. PS prompt footer
11. JavaScript: theme toggle, DataTable init helper, filter infrastructure

### Validated Brand Tokens

**Light mode (chalk-white):**

| Token | Hex | Use |
|-------|-----|-----|
| --bg-base | #fcfbf8 | Page background |
| --bg-elevated | #f3f0e8 | Cards, filter bar, code blocks |
| --bg-warm | #e8e4dc | Table header strip, tag pills |
| --border | #e2ddd2 | Borders, table grid |
| --border-dashed | #ddd8ce | Section dividers |
| --text | #181410 | Headlines |
| --text-body | #46423a | Running text, table cells |
| --text-muted | #787468 | Captions, secondary info |
| --text-dim | #a8a298 | Decorative chrome |
| --accent | #84441c | Rust accent, links, eyebrows |
| --accent-hover | #a85b1d | Hover state |
| --accent-soft | #fceadb | Callout backgrounds, hover rows |

**Dark mode (warm parchment, matching rksolutions.nl):**

| Token | Hex | Use |
|-------|-----|-----|
| --bg-base | #1a1710 | Page background |
| --bg-elevated | #252118 | Cards, elevated surfaces |
| --bg-deep | #151210 | Alternating table rows |
| --border | #3a3228 | Borders |
| --text | #e0d8c8 | Headlines |
| --text-body | #e0d8c8 | Running text |
| --text-muted | #8a7a60 | Captions, mono identifiers |
| --text-dim | #5a5040 | Decorative chrome |
| --accent | #c8a060 | Amber accent (replaces rust) |

**Status badges:**

| Status | Light BG | Dark Style |
|--------|----------|------------|
| OK/Compliant | #2d7a3a (white text) | rgba(90,160,90,0.2) border+text #6abf6a |
| Warning | #c46a1a (white text) | rgba(200,150,50,0.18) border+text #d4a840 |
| Error/NonCompliant | #c0392b (white text) | rgba(200,80,70,0.18) border+text #e06050 |
| N/A | #e8e4dc (muted text) | rgba(140,120,90,0.15) border+text #8a7a60 |

**Stat tile colors (solid background, white text in light; tinted bg + colored text in dark):**

| Tile | Light BG | Dark BG | Dark Text |
|------|----------|---------|-----------|
| Rust (hero) | #84441c | #3a2810 | #e8c080 |
| Olive | #4a6830 | #1a2c10 | #a0d870 |
| Steel | #3e5c78 | #10202e | #78c0e8 |
| Rose | #8a3e38 | #2e1410 | #e88878 |

### Migration Per Report

Each report's private helper gets refactored:

1. **Remove** all inline CSS, `<head>`, `<style>`, theme toggle HTML, footer HTML, DataTables init boilerplate
2. **Keep** only report-specific logic: stat cards HTML, filter dropdowns, table generation, tab structure
3. **Call** `New-RKSolutionsReportTemplate` with the report-specific content
4. Report-specific CSS (e.g. shimmer animation for Entra eligible badges) is passed via the `-CustomCss` parameter

**Files affected:**
- `module/Private/IntuneAnomalies.ps1` -- `New-IntuneAnomaliesHTMLReport`
- `module/Private/EntraAdminRoles.ps1` -- `New-AdminRoleHTMLReport`
- `module/Private/M365License.ps1` -- `New-HTMLReport`
- `module/Private/IntuneEnrollmentFlows.ps1` -- `New-AssignmentOverviewHtmlReport` + `New-DeviceVisualizationHtmlReport`
- `module/Private/CustomSecurityAttributes.ps1` -- new, built on template from start

### CDN Dependencies (unchanged)

- Bootstrap 5.3.0 CSS
- DataTables 1.13.6 + Buttons
- jQuery 3.7.0
- Font Awesome 6.4.0
- jszip, pdfmake (for export buttons)
- **New:** Google Fonts (Playfair Display, Source Serif 4, JetBrains Mono)

### Signature Elements (every report)

1. **Breadcrumb pill** -- `<- cd ./reports/{slug}`
2. **Eyebrow** -- `REPORT NAME . GENERATED YYYY-MM-DD`
3. **Playfair title** -- one word in accent color
4. **Source Serif italic lede**
5. **Dashed divider** between sections
6. **PS prompt footer** -- `PS C:\Blog\rksolutions>` + blinking cursor + `rksolutions.nl`
7. **Theme toggle** -- sun/moon, fixed top-right

## Testing Strategy

- Existing Pester consistency tests extended for new cmdlet
- Manual visual verification of all 5 reports in both themes
- No Graph connection required for template unit testing

## File Summary

| Action | File |
|--------|------|
| Create | `module/Private/New-RKSolutionsReportTemplate.ps1` |
| Create | `module/Public/Get-CustomSecurityAttributesReport.ps1` |
| Create | `module/Private/CustomSecurityAttributes.ps1` |
| Modify | `module/Private/IntuneAnomalies.ps1` |
| Modify | `module/Private/EntraAdminRoles.ps1` |
| Modify | `module/Private/M365License.ps1` |
| Modify | `module/Private/IntuneEnrollmentFlows.ps1` |
| Modify | `module/Public/Connect-RKGraph.ps1` |
| Modify | `module/RKSolutions.psd1` |
| Modify | `module/RKSolutions.psm1` |
| Modify | `Tests/Consistency.Tests.ps1` |
| Modify | `docs/PERMISSIONS.md` |
