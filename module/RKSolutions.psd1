@{
    RootModule        = 'RKSolutions.psm1'
    ModuleVersion     = '1.1.0'
    GUID              = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'
    Author            = 'Roy Klooster'
    CompanyName       = 'RK Solutions'
    Copyright         = '(c) 2026 Roy Klooster - RK Solutions. All rights reserved.'
    Description       = 'PowerShell module consolidating Intune Enrollment Flows, Intune Anomalies, Entra Admin Roles, M365 License Assignment, and Custom Security Attributes reports. Connects to Microsoft Graph and generates interactive HTML reports with the Carbon Ember design system.'
    PowerShellVersion = '7.0'
    RequiredModules   = @('Microsoft.Graph.Authentication')
    FunctionsToExport = @(
        'Connect-RKGraph',
        'Disconnect-RKGraph',
        'Get-IntuneEnrollmentFlowsReport',
        'Get-IntuneAnomaliesReport',
        'Get-EntraAdminRolesReport',
        'Get-M365LicenseAssignmentReport',
        'Get-CustomSecurityAttributesReport'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags         = @('RKSolutions', 'Microsoft365', 'MicrosoftIntune', 'MicrosoftEntraID', 'MicrosoftGraph', 'DeviceManagement', 'Reporting', 'CustomSecurityAttributes')
            LicenseUri   = 'https://opensource.org/licenses/MIT'
            ProjectUri   = 'https://github.com/royklo/RKSolutions-Module'
            ReleaseNotes = @'
1.1.0 - Rebranded all HTML reports with Carbon Ember design (Geist/Geist Mono typography, pill-style tabs, neutral dark theme). Added Get-CustomSecurityAttributesReport with multi-entity support (users, devices, enterprise apps), dynamic attribute set discovery, and coverage metrics. Fixed XSS vulnerabilities in HTML output. Added equal-width tabs, dark mode pagination, and locked DataTable column widths. Added skuFamily for OS edition detection, Windows 11 25H2 support. Added CustomSecAttributeAssignment.Read.All and CustomSecAttributeDefinition.Read.All scopes.
1.0.1 - Requires PowerShell 7.0 or higher. Fixed encoding issues for Windows compatibility. Improved error messaging when running on unsupported PowerShell versions.
1.0.0 - Initial module release. Consolidates Intune Enrollment Flows, Intune Anomalies, Entra Admin Roles, and M365 License Assignment reports into a single module.
'@
        }
    }
}
