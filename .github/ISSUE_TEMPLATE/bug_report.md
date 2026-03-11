---
name: Bug report
about: Report a bug or unexpected behavior
title: '[Bug] '
labels: bug
assignees: ''
---

## Description

A clear description of what the bug is.

## Steps to reproduce

1. Run command / do step '...'
2. Run command / do step '...'
3. See error or wrong output

## Expected behavior

What you expected to happen.

## Actual behavior

What actually happened (error message, wrong output, or no output).

## Environment

- **PowerShell version:** (output of `$PSVersionTable.PSVersion` — must be 7.0 or higher)
- **PowerShell host:** Are you running `pwsh` (PowerShell 7) or `powershell.exe` (Windows PowerShell 5.1)?
- **OS:** (e.g. Windows 11, macOS 14, Ubuntu 22.04)
- **Module version:** (output of `Get-Module RKSolutions | Select-Object Version`)

## Installation & Import

- **How did you install the module?**
  - [ ] PowerShell Gallery (`Install-Module -Name RKSolutions`)
  - [ ] GitHub (cloned/downloaded and `Import-Module ./module/RKSolutions.psd1`)

- **Does the module load correctly?**
  Run `Get-Command -Module RKSolutions` and paste the output:
  ```powershell
  # Paste output here
  ```

## Additional context

Logs, screenshots, or any other details that might help.
