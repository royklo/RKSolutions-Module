<#
.SYNOPSIS
    Generates an HTML report of Microsoft Entra ID administrative role assignments (including PIM and group-based).
    Connect first with Connect-RKGraph, or pass auth parameters to this cmdlet.
#>
function Get-EntraAdminRolesReport {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)] [string[]] $RequiredScopes = @('Directory.Read.All', 'PrivilegedEligibilitySchedule.Read.AzureADGroup', 'Organization.Read.All', 'AuditLog.Read.All', 'RoleManagement.Read.Directory', 'Mail.Send', 'RoleAssignmentSchedule.Read.Directory'),
    [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')] [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')] [Parameter(Mandatory = $false, ParameterSetName = 'Identity')] [Parameter(Mandatory = $true, ParameterSetName = 'AccessToken')] [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')] [string] $TenantId,
    [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')] [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')] [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')] [string] $ClientId,
    [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')] [object] $ClientSecret,
    [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')] [string] $CertificateThumbprint,
    [Parameter(Mandatory = $true, ParameterSetName = 'Identity')] [switch] $Identity,
    [Parameter(Mandatory = $true, ParameterSetName = 'AccessToken')] [object] $AccessToken,
    [Parameter(Mandatory = $false)] [switch] $SendEmail,
    [Parameter(Mandatory = $false)] [string[]] $Recipient,
    [Parameter(Mandatory = $false)] [string] $From,
    [Parameter(Mandatory = $false)] [string] $ExportPath,
    [Parameter(Mandatory = $false)] [switch] $DebugMode
)

$ErrorActionPreference = 'Stop'
if ($ClientSecret -is [string]) { $ClientSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force }
if ($AccessToken -is [string]) { $AccessToken = ConvertTo-SecureString $AccessToken -AsPlainText -Force }
try {
    $connected = Invoke-RKSolutionsWithConnection -RequiredScopes $RequiredScopes -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -Identity:$Identity -AccessToken $AccessToken -DebugMode:$DebugMode -ParameterSetName $PSCmdlet.ParameterSetName
    if (-not $connected) { throw 'Failed to connect to Microsoft Graph API.' }

    Invoke-EntraAdminRolesReportCore -SendEmail:$SendEmail -Recipient $Recipient -From $From -ExportPath $ExportPath -DebugMode:$DebugMode
}
catch { Write-Error "Error: $_"; throw $_ }
finally {
    # Session left connected; use Disconnect-RKGraph when done.
}
}
