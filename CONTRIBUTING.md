# Contributing to the RKSolutions PowerShell Module

Thank you for your interest in contributing. This project is maintained by Roy Klooster / RK Solutions. Below is how to contribute via fork and pull request, how to report bugs, and what we expect from code changes.

## How to contribute (fork and pull request)

1. **Fork the repository**
   Click **Fork** on [the GitHub repo](https://github.com/royklo/RK-Solutions-PSModule).
2. **Clone your fork**
   ```powershell
   git clone https://github.com/YOUR_USERNAME/RK-Solutions-PSModule.git
   cd RK-Solutions-PSModule
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
   Go to the [original repository](https://github.com/royklo/RK-Solutions-PSModule) and open a **New pull request** from your branch. Fill in the PR template (summary of changes, related issue if any, checklist).

## Development setup

- **PowerShell 5.1+** or **PowerShell 7+** (see `module/RKSolutions.psd1`).
- No build step: the module is a script module. After editing files under `module/`, re-import:
  ```powershell
  Import-Module ./module/RKSolutions.psd1 -Force
  ```

## How to raise a bug

Use the **Bug report** issue template so we get the information we need:

1. Go to [New issue](https://github.com/royklo/RK-Solutions-PSModule/issues/new).
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

1. Go to [New issue](https://github.com/royklo/RK-Solutions-PSModule/issues/new).
2. Choose **Feature request**.
3. Describe the feature, the use case, and your proposed solution.

## Testing

From the repository root:

```powershell
Invoke-Pester ./Tests/Consistency.Tests.ps1
```

This checks that exported functions and key parameters are consistent and that comment-based help is present.

## Adding a new cmdlet

1. Implement in `module/Public/` and add the function name to `FunctionsToExport` in `module/RKSolutions.psd1` (and to `Export-ModuleMember` in `module/RKSolutions.psm1` if applicable).
2. Update `docs/CMDLET-REFERENCE.md` with synopsis, parameters, and examples.
3. Run the consistency tests; update `Tests/Consistency.Tests.ps1` if the expected cmdlet or parameter list changes.

Thank you for contributing.
