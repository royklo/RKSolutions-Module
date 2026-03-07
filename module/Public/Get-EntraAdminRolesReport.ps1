<#
.SYNOPSIS
    Generates an HTML report of Microsoft Entra ID administrative role assignments (including PIM and group-based).
    Connect first with Connect-RKGraph; this cmdlet uses the existing connection.
#>
function Get-EntraAdminRolesReport {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)] [switch] $SendEmail,
    [Parameter(Mandatory = $false)] [string[]] $Recipient,
    [Parameter(Mandatory = $false)] [string] $From,
    [Parameter(Mandatory = $false)] [string] $ExportPath,
    [Parameter(Mandatory = $false)] [switch] $DebugMode
)

$ErrorActionPreference = 'Stop'
# Scopes required by this report (authorization is handled by Connect-RKGraph)
$requiredScopes = @('Directory.Read.All', 'PrivilegedEligibilitySchedule.Read.AzureADGroup', 'Organization.Read.All', 'AuditLog.Read.All', 'RoleManagement.Read.Directory', 'Mail.Send', 'RoleAssignmentSchedule.Read.Directory')
try {
    $connected = Invoke-RKSolutionsWithConnection -RequiredScopes $requiredScopes -ParameterSetName 'Interactive' -DebugMode:$DebugMode
    if (-not $connected) { throw 'Failed to connect to Microsoft Graph API.' }

    Invoke-EntraAdminRolesReportCore -SendEmail:$SendEmail -Recipient $Recipient -From $From -ExportPath $ExportPath -DebugMode:$DebugMode
}
catch { Write-Error "Error: $_"; throw $_ }
finally {
    # Session left connected; use Disconnect-RKGraph when done.
}
}
