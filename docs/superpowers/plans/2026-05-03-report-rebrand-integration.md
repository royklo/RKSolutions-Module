# Report Rebrand + Custom Security Attributes Integration

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Create a shared HTML report template with rksolutions.nl branding, integrate Get-CustomSecurityAttributesReport, and migrate all existing reports to the shared template.

**Architecture:** A single `New-RKSolutionsReportTemplate` PowerShell function provides the full HTML shell (head, CSS, JS, brand elements). Each report only builds its unique content (stats cards, tables, filters) and passes it to the template. The new Custom Security Attributes report is built on the template from the start; existing reports are migrated one at a time.

**Tech Stack:** PowerShell 7.0+, Microsoft Graph API, HTML5/CSS3/JS, Bootstrap 5, DataTables, Google Fonts (Playfair Display, Source Serif 4, JetBrains Mono)

**Spec:** `docs/superpowers/specs/2026-05-03-report-integration-rebrand-design.md`

---

### Task 1: Create Shared HTML Report Template

**Files:**
- Create: `module/Private/New-RKSolutionsReportTemplate.ps1`

This is the foundation. Every subsequent task depends on it.

- [ ] **Step 1: Create `New-RKSolutionsReportTemplate.ps1`**

The function accepts: `$TenantName`, `$ReportTitle`, `$ReportSlug`, `$Eyebrow`, `$Lede`, `$StatsCardsHtml`, `$BodyContentHtml`, `$CustomCss`, `$ReportDate`, `$Tags`.

Returns a complete HTML document string via expandable here-string with:
- Google Fonts + CDN links (Bootstrap 5, DataTables, Font Awesome, jQuery, jszip, pdfmake)
- CSS custom properties for light (`:root`) and dark (`[data-theme="dark"]`) themes using the validated brand tokens from the spec
- Shared CSS: typography (Playfair Display for titles, Source Serif 4 for body, JetBrains Mono for monospace), breadcrumb pill, eyebrow line, title block, stat tile colors (rust/olive/steel/rose), filter bar, table styles, status badges, tag pills, theme toggle, DataTables overrides, PS prompt footer
- `$CustomCss` injected inside `<style>` block
- HTML structure: theme toggle -> breadcrumb pill -> eyebrow -> title -> lede -> tags -> dashed divider -> stats cards -> body content -> PS prompt footer
- JavaScript: theme toggle with localStorage, DataTable init helper function `initRKTable(selector, options)`, pagination color management

- [ ] **Step 2: Register in PSM1**

Edit `module/RKSolutions.psm1` line 27 `$sharedOrder` array: add `'New-RKSolutionsReportTemplate.ps1'` after `'ConvertTo-DateString.ps1'`.

- [ ] **Step 3: Commit**

```
feat: add shared HTML report template with rksolutions.nl brand
```

---

### Task 2: Integrate Custom Security Attributes Report

**Files:**
- Create: `module/Private/CustomSecurityAttributes.ps1`
- Create: `module/Public/Get-CustomSecurityAttributesReport.ps1`
- Modify: `module/Public/Connect-RKGraph.ps1`
- Modify: `module/RKSolutions.psm1`

- [ ] **Step 1: Create `CustomSecurityAttributes.ps1`**

Two functions:
- `Get-CustomSecurityAttributeData`: fetches attribute definitions from `beta/directory/customSecurityAttributeDefinitions`, queries users via `v1.0/users` with `$filter` and `ConsistencyLevel=eventual` header, builds PSCustomObject array
- `New-CustomSecurityAttributesHTMLReport`: builds stats cards HTML (total users + up to 3 unique attribute counts), filter dropdowns, table rows, then calls `New-RKSolutionsReportTemplate` with slug `custom-security-attributes`, eyebrow `CUSTOM SECURITY ATTRIBUTES`, etc.

Port the data-fetching logic from `/Users/roy/github/Get-CustomSecurityAttributesReport.ps1` (lines 1526-1644), adapting to use existing Graph context instead of inline auth.

- [ ] **Step 2: Create `Get-CustomSecurityAttributesReport.ps1`**

Thin wrapper with `.SYNOPSIS` help block. Parameters: `$AttributeSet` (default "CustomerData"), `$AttributeNames`, `$Filters`, `$SendEmail`, `$Recipient`, `$From`, `$ExportPath`, `$DebugMode`. Follows the same pattern as `Get-IntuneAnomaliesReport.ps1`.

- [ ] **Step 3: Add scope to Connect-RKGraph**

Add `'CustomSecAttributeAssignment.Read.All'` to the `$RequiredScopes` default array in `module/Public/Connect-RKGraph.ps1`.

- [ ] **Step 4: Register in PSM1**

- Add `'CustomSecurityAttributes.ps1'` to `$domainOrder` array
- Add `'Get-CustomSecurityAttributesReport'` to `Export-ModuleMember`

- [ ] **Step 5: Commit**

```
feat: add Get-CustomSecurityAttributesReport cmdlet
```

---

### Task 3: Migrate Intune Anomalies Report to Shared Template

**Files:**
- Modify: `module/Private/IntuneAnomalies.ps1`

- [ ] **Step 1: Refactor `New-IntuneAnomaliesHTMLReport`**

Keep: parameter list, count calculations, stat cards HTML generation, tab structure, table row generation, filter logic.

Remove: entire `<head>` section, all CSS (`:root` vars, theme toggle, stats cards, tables, badges, DataTables overrides, footer), theme toggle HTML, footer HTML, DataTable init JS boilerplate.

Replace: build `$statsCardsHtml` (8 stat tiles using the template's `.stat-tile.t-{color}` classes), `$bodyContentHtml` (tabs + tables + filters), and any report-specific `$customCss`. Call `New-RKSolutionsReportTemplate` with slug `intune-anomalies`, eyebrow `INTUNE ANOMALIES`, title accent word `Anomalies`, lede from spec.

- [ ] **Step 2: Verify structure**

Confirm the function still outputs HTML to `$ExportPath` and sets `$script:ExportPath` for email flow.

- [ ] **Step 3: Commit**

```
refactor: migrate Intune Anomalies report to shared brand template
```

---

### Task 4: Migrate Entra Admin Roles Report to Shared Template

**Files:**
- Modify: `module/Private/EntraAdminRoles.ps1`

- [ ] **Step 1: Refactor `New-AdminRoleHTMLReport`**

Same pattern as Task 3. Keep report-specific: shimmer animation CSS (pass via `$CustomCss`), stat cards for Permanent/Eligible/Group/ServicePrincipal counts, tab structure with role tables, badge styling for assignment types, group-jump highlight animation.

Call `New-RKSolutionsReportTemplate` with slug `entra-admin-roles`, eyebrow `ENTRA ADMIN ROLES`, title accent word `Admin Roles`.

- [ ] **Step 2: Commit**

```
refactor: migrate Entra Admin Roles report to shared brand template
```

---

### Task 5: Migrate M365 License Assignment Report to Shared Template

**Files:**
- Modify: `module/Private/M365License.ps1`

- [ ] **Step 1: Refactor `New-HTMLReport`**

Keep report-specific: stat cards for Direct/Inherited/Both/Disabled, license filter badges, subscription overview table, disabled users table.

Call `New-RKSolutionsReportTemplate` with slug `m365-licenses`, eyebrow `M365 LICENSE ASSIGNMENT`, title accent word `License`.

- [ ] **Step 2: Commit**

```
refactor: migrate M365 License report to shared brand template
```

---

### Task 6: Migrate Intune Enrollment Flows Reports to Shared Template

**Files:**
- Modify: `module/Private/IntuneEnrollmentFlows.ps1`

- [ ] **Step 1: Refactor `New-AssignmentOverviewHtmlReport`**

Call `New-RKSolutionsReportTemplate` with slug `intune-enrollment-overview`, eyebrow `INTUNE ENROLLMENT FLOWS`, title accent word `Enrollment`.

- [ ] **Step 2: Refactor `New-DeviceVisualizationHtmlReport`**

Call `New-RKSolutionsReportTemplate` with slug `intune-enrollment-device`, eyebrow `DEVICE ENROLLMENT FLOW`, title accent word `Enrollment`. Keep Mermaid diagram integration CSS via `$CustomCss`.

- [ ] **Step 3: Commit**

```
refactor: migrate Intune Enrollment Flows reports to shared brand template
```

---

### Task 7: Update Manifest, Tests, and Docs

**Files:**
- Modify: `module/RKSolutions.psd1`
- Modify: `Tests/Consistency.Tests.ps1`
- Modify: `docs/PERMISSIONS.md`

- [ ] **Step 1: Update manifest**

- Add `'Get-CustomSecurityAttributesReport'` to `FunctionsToExport`
- Bump `ModuleVersion` to `'1.1.0'`
- Change `PowerShellVersion` from `'5.1'` to `'7.0'`
- Update `Description` to include Custom Security Attributes
- Add 1.1.0 to `ReleaseNotes`

- [ ] **Step 2: Update consistency tests**

Add to `$script:expectedParameters`:
```powershell
'Get-CustomSecurityAttributesReport' = @('AttributeSet', 'ExportPath')
```

- [ ] **Step 3: Update PERMISSIONS.md**

Add `CustomSecAttributeAssignment.Read.All` to summary table and add per-cmdlet section for `Get-CustomSecurityAttributesReport`.

- [ ] **Step 4: Run tests**

```bash
pwsh -Command "Invoke-Pester ./Tests/Consistency.Tests.ps1 -Verbose"
```

Expected: All tests pass, including the new cmdlet.

- [ ] **Step 5: Commit**

```
feat: update manifest, tests, and docs for v1.1.0
```
