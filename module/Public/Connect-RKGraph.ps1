<#
.SYNOPSIS
    Connects to Microsoft Graph for use with RKSolutions report cmdlets.
.DESCRIPTION
    Establishes a Microsoft Graph session. Run this once, then run report cmdlets (e.g. Get-IntuneEnrollmentFlowsReport,
    Get-EntraAdminRolesReport) without passing auth parameters; they will use this connection.
    Default -RequiredScopes includes permissions needed for all report cmdlets. You can pass -RequiredScopes to limit scopes.
    When done, run Disconnect-RKGraph to clear the session.
.PARAMETER RequiredScopes
    API permission scopes. Default is the union of all scopes required by the four report cmdlets so one connection works for every report without re-prompting. You can pass -RequiredScopes to limit scopes.
#>
function Connect-RKGraph {
    [CmdletBinding(DefaultParameterSetName = 'Interactive')]
    param(
        [Parameter(Mandatory = $false)]
        [string[]] $RequiredScopes = @(
            'User.Read',
            'User.Read.All',
            'Group.Read.All',
            'GroupMember.Read.All',
            'Device.Read.All',
            'DeviceManagementConfiguration.Read.All',
            'DeviceManagementApps.Read.All',
            'DeviceManagementManagedDevices.Read.All',
            'DeviceManagementServiceConfig.Read.All',
            'Directory.Read.All',
            'Organization.Read.All',
            'AuditLog.Read.All',
            'RoleManagement.Read.Directory',
            'RoleAssignmentSchedule.Read.Directory',
            'PrivilegedEligibilitySchedule.Read.AzureADGroup',
            'Mail.Send',
            'CloudLicensing.Read',
            'CloudPC.Read.All',
            'CustomSecAttributeAssignment.Read.All',
            'CustomSecAttributeDefinition.Read.All'
        ),
        [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Identity')]
        [Parameter(Mandatory = $true, ParameterSetName = 'AccessToken')]
        [string] $TenantId,
        [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
        [string] $ClientId,
        [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')]
        [SecureString] $ClientSecret,
        [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')]
        [string] $CertificateThumbprint,
        [Parameter(Mandatory = $true, ParameterSetName = 'Identity')]
        [switch] $Identity,
        [Parameter(Mandatory = $true, ParameterSetName = 'AccessToken')]
        [SecureString] $AccessToken,
        [Parameter(Mandatory = $false)]
        [switch] $DebugMode
    )
    $params = @{}
    foreach ($key in $PSBoundParameters.Keys) {
        $params[$key] = $PSBoundParameters[$key]
    }
    $connected = Connect-ToMgGraph @params
    if ($connected) {
        $ctx = Get-MgContext
        Write-Host "Connected to Microsoft Graph (tenant: $($ctx.TenantId))." -ForegroundColor Green
    }
    return $connected
}
