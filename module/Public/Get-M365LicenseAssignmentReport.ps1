<#
.SYNOPSIS
    Generates an HTML report of Microsoft 365 license assignments across the tenant.
    Connect first with Connect-RKGraph, or pass auth parameters to this cmdlet.
#>
function Get-M365LicenseAssignmentReport {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)] [string[]] $RequiredScopes = @('User.Read.All', 'AuditLog.Read.All', 'GroupMember.Read.All', 'Group.Read.All', 'Directory.Read.All', 'Organization.Read.All', 'RoleManagement.Read.Directory', 'Mail.Send', 'CloudLicensing.Read'),
    [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')] [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')] [Parameter(Mandatory = $false, ParameterSetName = 'Identity')] [Parameter(Mandatory = $true, ParameterSetName = 'AccessToken')] [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')] [string] $TenantId,
    [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')] [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')] [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')] [string] $ClientId,
    [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')] [SecureString] $ClientSecret,
    [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')] [string] $CertificateThumbprint,
    [Parameter(Mandatory = $true, ParameterSetName = 'Identity')] [switch] $Identity,
    [Parameter(Mandatory = $true, ParameterSetName = 'AccessToken')] [SecureString] $AccessToken,
    [Parameter(Mandatory = $false)] [switch] $SendEmail,
    [Parameter(Mandatory = $false)] [string[]] $Recipient,
    [Parameter(Mandatory = $false)] [string] $From,
    [Parameter(Mandatory = $false)] [string] $ExportPath,
    [Parameter(Mandatory = $false)] [switch] $DebugMode
)

$ErrorActionPreference = 'Stop'
try {
    $connected = Invoke-RKSolutionsWithConnection -RequiredScopes $RequiredScopes -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -Identity:$Identity -AccessToken $AccessToken -DebugMode:$DebugMode -ParameterSetName $PSCmdlet.ParameterSetName
    if (-not $connected) { throw 'Failed to connect to Microsoft Graph API.' }

    Invoke-M365LicenseReportCore -SendEmail:$SendEmail -Recipient $Recipient -From $From -ExportPath $ExportPath
}
catch { Write-Error "Error: $_"; throw $_ }
finally {
    # Session left connected; use Disconnect-RKGraph when done.
}
}
