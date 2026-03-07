<#
.SYNOPSIS
    Generates an HTML report of Microsoft 365 license assignments across the tenant.
    Connect first with Connect-RKGraph; this cmdlet uses the existing connection.
#>
function Get-M365LicenseAssignmentReport {
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
$requiredScopes = @('User.Read.All', 'AuditLog.Read.All', 'GroupMember.Read.All', 'Group.Read.All', 'Directory.Read.All', 'Organization.Read.All', 'RoleManagement.Read.Directory', 'Mail.Send', 'CloudLicensing.Read')
try {
    $connected = Invoke-RKSolutionsWithConnection -RequiredScopes $requiredScopes -ParameterSetName 'Interactive' -DebugMode:$DebugMode
    if (-not $connected) { throw 'Failed to connect to Microsoft Graph API.' }

    Invoke-M365LicenseReportCore -SendEmail:$SendEmail -Recipient $Recipient -From $From -ExportPath $ExportPath
}
catch { Write-Error "Error: $_"; throw $_ }
finally {
    # Session left connected; use Disconnect-RKGraph when done.
}
}
