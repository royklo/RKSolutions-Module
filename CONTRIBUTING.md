# Contributing to the RKSolutions PowerShell Module

Thank you for your interest in contributing. This project is maintained by Roy Klooster / RK Solutions. Below is how to contribute via fork and pull request, how to report bugs, and what we expect from code changes.

## How to contribute (fork and pull request)

1. **Fork the repository**
   Click **Fork** on [the GitHub repo](https://github.com/royklo/RKSolutions-Module).
2. **Clone your fork**
   ```powershell
   git clone https://github.com/YOUR_USERNAME/RKSolutions-Module.git
   cd RKSolutions-Module
   ```
3. **Create a branch** for your change
   ```powershell
   git checkout -b feature/your-feature-name
   # or: git checkout -b fix/bug-description
   ```
4. **Make your changes** in `module/` (see [Development setup](#development-setup)).
5. **Run tests** (see [Testing](#testing)).
6. **Commit and push** to your fork
   ```powershell
   git add .
   git commit -m "Short description of your change"
   git push origin feature/your-feature-name
   ```
7. **Open a pull request**
   Go to the [original repository](https://github.com/royklo/RKSolutions-Module) and open a **New pull request** from your branch. Fill in the PR template (summary of changes, related issue if any, checklist).

## Development setup

- **PowerShell 5.1+** or **PowerShell 7+** (see `module/RKSolutions.psd1`).
- No build step: the module is a script module. After editing files under `module/`, re-import:
  ```powershell
  Import-Module ./module/RKSolutions.psd1 -Force
  ```

## How to raise a bug

Use the **Bug report** issue template so we get the information we need:

1. Go to [New issue](https://github.com/royklo/RKSolutions-Module/issues/new).
2. Choose **Bug report**.
3. Fill in:
   - **Description** — What went wrong?
   - **Steps to reproduce** — Exact commands or steps.
   - **Expected behavior** — What you expected.
   - **Actual behavior** — What happened instead.
   - **Environment** — PowerShell version, OS, module version (e.g. `Get-Module RKSolutions | Select-Object Version`).
   - **Additional context** — Logs, screenshots, or other details.

## How to request a feature

Use the **Feature request** template:

1. Go to [New issue](https://github.com/royklo/RKSolutions-Module/issues/new).
2. Choose **Feature request**.
3. Describe the feature, the use case, and your proposed solution.

## Testing

From the repository root:

```powershell
Invoke-Pester ./Tests/Consistency.Tests.ps1
```

This checks that exported functions and key parameters are consistent and that comment-based help is present.

## Microsoft Graph permissions

When you add or change features that call Microsoft Graph, you must keep permissions documented and in sync.

### Where permissions are defined

- **Per cmdlet:** Each Public cmdlet that needs Graph has a `-RequiredScopes` parameter (default array of permission strings). That default is the single source of truth for what that cmdlet needs.
- **Connect-RKGraph:** Its default `-RequiredScopes` is the **union** of all report cmdlets so one connection works for every report. When you add a new report cmdlet or new Graph calls, consider whether Connect-RKGraph’s default list should include the new scopes.
- **Docs:** [docs/PERMISSIONS.md](docs/PERMISSIONS.md) lists every permission and which cmdlet(s) use it, plus a per-cmdlet breakdown. It is generated from the scripts—so when you change `RequiredScopes` or add a cmdlet, you must update PERMISSIONS.md.

### When you add or change Graph-dependent code

1. **Identify required scopes**  
   Check [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference) for the APIs you call (e.g. `GET /deviceManagement/managedDevices` → `DeviceManagementManagedDevices.Read.All`).

2. **Update the cmdlet**  
   - New cmdlet: add a `-RequiredScopes` parameter with a default array of the minimum scopes needed.  
   - Existing cmdlet: add any new scope to the existing `$RequiredScopes` default.

3. **Update Connect-RKGraph (if applicable)**  
   If the new/updated cmdlet should work with a single `Connect-RKGraph` call (no extra consent), add the new scope(s) to the default `$RequiredScopes` in `module/Public/Connect-RKGraph.ps1`.

4. **Update docs/PERMISSIONS.md**  
   - In the **Summary** table: add or update the permission row and the “Used by” column.  
   - In **Per cmdlet**: add a new section for a new cmdlet, or update the existing cmdlet’s list so it matches the script’s `RequiredScopes` default.

5. **In your PR**  
   Mention that new/updated Graph permissions were added and that PERMISSIONS.md was updated.

### Checklist for permission changes

- [ ] Required scopes identified from Graph API docs.
- [ ] Cmdlet’s default `-RequiredScopes` (or equivalent) updated.
- [ ] Connect-RKGraph default scopes updated if the cmdlet should work with one connection.
- [ ] [docs/PERMISSIONS.md](docs/PERMISSIONS.md) updated (summary table and per-cmdlet section).
- [ ] PR description notes the permission changes.

## Adding a new cmdlet

1. Implement in `module/Public/` and add the function name to `FunctionsToExport` in `module/RKSolutions.psd1` (and to `Export-ModuleMember` in `module/RKSolutions.psm1` if applicable).
2. If the cmdlet calls Microsoft Graph, define its required permissions and keep [Microsoft Graph permissions](#microsoft-graph-permissions) in sync: set the cmdlet’s default `-RequiredScopes`, update Connect-RKGraph if desired, and update **docs/PERMISSIONS.md**.
3. Update `docs/CMDLET-REFERENCE.md` with synopsis, parameters, and examples.
4. Run the consistency tests; update `Tests/Consistency.Tests.ps1` if the expected cmdlet or parameter list changes.

Thank you for contributing.
