# RKSolutions PowerShell Script Module

This is the script implementation of the RKSolutions module. It provides cmdlets to connect to Microsoft Graph and generate reports for Intune Enrollment Flows, Intune Anomalies, Entra Admin Roles, and M365 License Assignment.

## Module structure

```
module/
├── RKSolutions.psd1           # Module manifest
├── RKSolutions.psm1           # Root script (dot-sources Public + Private)
├── README.md                  # This file
├── Public/                    # Exported cmdlets
│   ├── Connect-RKGraph.ps1
│   ├── Disconnect-RKGraph.ps1
│   ├── Get-IntuneEnrollmentFlowsReport.ps1
│   ├── Get-IntuneAnomaliesReport.ps1
│   ├── Get-EntraAdminRolesReport.ps1
│   └── Get-M365LicenseAssignmentReport.ps1
└── Private/                   # Helpers (not exported, or exported via manifest)
    ├── Connect-ToMgGraph.ps1
    ├── Invoke-RKSolutionsWithConnection.ps1
    ├── Invoke-GraphRequestWithPaging.ps1
    ├── Send-EmailWithAttachment.ps1
    ├── ConvertTo-DateString.ps1
    ├── Install-Requirements.ps1
    ├── Export-Results.ps1
    ├── IntuneEnrollmentFlows.ps1
    ├── IntuneAnomalies.ps1
    ├── EntraAdminRoles.ps1
    └── M365License.ps1
```

## Loading the script module

Run from the repository root so the path resolves to this repo's module folder.

From the repository root:

```powershell
Import-Module ./module/RKSolutions.psd1 -Force
```

Or from the `module` folder:

```powershell
Import-Module ./RKSolutions.psd1 -Force
```

## Quick start

```powershell
Connect-RKGraph -Scopes 'DeviceManagementManagedDevices.Read.All', 'User.Read.All'
Get-IntuneEnrollmentFlowsReport
Disconnect-RKGraph
```

For full documentation, prerequisites, and contributing, see the repository root **README.md** and **CONTRIBUTING.md**.
