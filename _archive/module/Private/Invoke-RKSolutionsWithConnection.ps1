# Private: Build connection params from parameter set and call Connect-ToMgGraph (used by Public report cmdlets)
function Invoke-RKSolutionsWithConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]] $RequiredScopes,
        [Parameter(Mandatory = $false)]
        [string] $TenantId,
        [Parameter(Mandatory = $false)]
        [string] $ClientId,
        [Parameter(Mandatory = $false)]
        [SecureString] $ClientSecret,
        [Parameter(Mandatory = $false)]
        [string] $CertificateThumbprint,
        [Parameter(Mandatory = $false)]
        [switch] $Identity,
        [Parameter(Mandatory = $false)]
        [SecureString] $AccessToken,
        [Parameter(Mandatory = $false)]
        [switch] $DebugMode,
        [Parameter(Mandatory = $true)]
        [string] $ParameterSetName
    )
    $connectionParams = @{ RequiredScopes = $RequiredScopes }
    if ($ParameterSetName -eq 'ClientSecret') {
        $connectionParams.TenantId = $TenantId
        $connectionParams.ClientId = $ClientId
        $connectionParams.ClientSecret = $ClientSecret
    } elseif ($ParameterSetName -eq 'Certificate') {
        $connectionParams.TenantId = $TenantId
        $connectionParams.ClientId = $ClientId
        $connectionParams.CertificateThumbprint = $CertificateThumbprint
    } elseif ($ParameterSetName -eq 'Identity') {
        $connectionParams.Identity = $true
        if ($TenantId) { $connectionParams.TenantId = $TenantId }
    } elseif ($ParameterSetName -eq 'AccessToken') {
        $connectionParams.AccessToken = $AccessToken
        $connectionParams.TenantId = $TenantId
    } else {
        if ($TenantId) { $connectionParams.TenantId = $TenantId }
        if ($ClientId) { $connectionParams.ClientId = $ClientId }
    }
    if ($DebugMode) { $connectionParams.DebugMode = $true }
    return Connect-ToMgGraph @connectionParams
}
