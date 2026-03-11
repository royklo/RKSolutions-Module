@{
    RootModule        = 'RKSolutions.psm1'
    ModuleVersion     = '1.0.1'
    GUID              = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'
    Author            = 'Roy Klooster'
    CompanyName       = 'RK Solutions'
    Copyright         = '(c) 2026 Roy Klooster - RK Solutions. All rights reserved.'
    Description       = 'PowerShell module consolidating Intune Enrollment Flows, Intune Anomalies, Entra Admin Roles, and M365 License Assignment reports. Connects to Microsoft Graph and generates HTML/CSV reports.'
    PowerShellVersion = '5.1'
    RequiredModules   = @('Microsoft.Graph.Authentication')
    FunctionsToExport = @(
        'Connect-RKGraph',
        'Disconnect-RKGraph',
        'Get-IntuneEnrollmentFlowsReport',
        'Get-IntuneAnomaliesReport',
        'Get-EntraAdminRolesReport',
        'Get-M365LicenseAssignmentReport'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags         = @('RKSolutions', 'Microsoft365', 'MicrosoftIntune', 'MicrosoftEntraID', 'MicrosoftGraph', 'DeviceManagement', 'Reporting')
            LicenseUri   = 'https://opensource.org/licenses/MIT'
            ProjectUri   = 'https://github.com/royklo/RKSolutions-Module'
            ReleaseNotes = @'
1.0.1 - Requires PowerShell 7.0 or higher. Fixed encoding issues for Windows compatibility. Improved error messaging when running on unsupported PowerShell versions.
1.0.0 - Initial module release. Consolidates Intune Enrollment Flows, Intune Anomalies, Entra Admin Roles, and M365 License Assignment reports into a single module.
'@
        }
    }
}
