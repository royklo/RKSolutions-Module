# RKSolutions PowerShell Module

[PowerShell Gallery](https://www.powershellgallery.com/packages/RKSolutions)  
[CI](https://github.com/royklo/RK-Solutions-PSModule/actions)  
[License: MIT](LICENSE)

PowerShell module for Microsoft Graph–backed reports: Intune Enrollment Flows, Intune Anomalies, Entra Admin Roles, and M365 License Assignment. Connects to Microsoft Graph and generates HTML/CSV reports.

## About

This module is maintained by **Roy Klooster** (RK Solutions).

- **Repository:** [https://github.com/royklo/RK-Solutions-PSModule](https://github.com/royklo/RK-Solutions-PSModule)
- **PowerShell Gallery:** [https://www.powershellgallery.com/packages/RKSolutions](https://www.powershellgallery.com/packages/RKSolutions)

## Repository structure

```
RK-Solutions-PSModule/
├── README.md                 # This file
├── LICENSE                   # MIT
├── CONTRIBUTING.md           # How to contribute, fork & PR
├── .gitignore
├── .github/
│   ├── ISSUE_TEMPLATE/       # Bug report, feature request
│   ├── PULL_REQUEST_TEMPLATE.md
│   └── workflows/           # build-and-test.yml, trigger-publish.yml
├── docs/
│   └── CMDLET-REFERENCE.md   # Parameters and example output
├── CHANGELOG.md              # Release history
├── module/                   # Script module (see module/README.md)
│   ├── RKSolutions.psd1
│   ├── RKSolutions.psm1
│   ├── README.md
│   ├── Public/               # Exported cmdlets
│   └── Private/              # Helpers
└── Tests/
    └── Consistency.Tests.ps1 # Pester tests
```

## Prerequisites

- **PowerShell 5.1+** or **PowerShell 7+** (cross-platform)
- **Microsoft.Graph.Authentication** (and other Graph modules as required by the cmdlets)
- Microsoft Graph permissions / app registration for the reports you run

## Installation

### From PowerShell Gallery (recommended)

```powershell
Install-Module -Name RKSolutions -Scope CurrentUser
```

### From source (GitHub)

```powershell
git clone https://github.com/royklo/RK-Solutions-PSModule.git
cd RK-Solutions-PSModule
Import-Module ./module/RKSolutions.psd1 -Force
```

Always run `Import-Module` from the **repository root** and use `./module/RKSolutions.psd1`.

## Quick start

```powershell
# Connect to Microsoft Graph
Connect-RKGraph -Scopes 'DeviceManagementManagedDevices.Read.All', 'User.Read.All'

# Generate reports (examples)
Get-IntuneEnrollmentFlowsReport -AssignmentOverviewOnly
Get-IntuneAnomaliesReport
Get-EntraAdminRolesReport
Get-M365LicenseAssignmentReport

# Disconnect when done
Disconnect-RKGraph
```

## Cmdlets

| Cmdlet | Description |
|--------|-------------|
| **Connect-RKGraph** | Establishes a Microsoft Graph session for report cmdlets. |
| **Disconnect-RKGraph** | Disconnects and clears the Graph session. |
| **Get-IntuneEnrollmentFlowsReport** | Generates Intune assignment overview and/or device visualization report. |
| **Get-IntuneAnomaliesReport** | Generates Intune anomalies report. |
| **Get-EntraAdminRolesReport** | Generates Entra admin roles report. |
| **Get-M365LicenseAssignmentReport** | Generates M365 license assignment report. |
| **Get-DeviceEvaluationContext** | Returns device evaluation context (used by enrollment flows report). |
| **Get-CloudPCProvisioningPolicyGroupInfo** | Returns Cloud PC provisioning policy group info. |

For full parameter details and examples, see **[Cmdlet Reference](docs/CMDLET-REFERENCE.md)**.

## Contributing

We welcome contributions: fork the repo, make your changes, and open a pull request. See **[CONTRIBUTING.md](CONTRIBUTING.md)** for the workflow and how to report bugs.

## Issues

- [Bug report](https://github.com/royklo/RK-Solutions-PSModule/issues/new?template=bug_report.md)
- [Feature request](https://github.com/royklo/RK-Solutions-PSModule/issues/new?template=feature_request.md)

## License

[MIT](LICENSE) — Copyright (c) Roy Klooster - RK Solutions.
