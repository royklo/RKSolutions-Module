# Private: Connect to Microsoft Graph (Interactive, ClientSecret, Certificate, Identity, AccessToken)
function Connect-ToMgGraph {
    [CmdletBinding(DefaultParameterSetName = 'Interactive')]
    param(
        [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
        [Parameter(Mandatory = $false, ParameterSetName = 'ClientSecret')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Certificate')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Identity')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AccessToken')]
        [string[]] $RequiredScopes = @('User.Read'),

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

    Install-Requirements -DebugMode:$DebugMode | Out-Null
    $AuthMethod = $PSCmdlet.ParameterSetName
    Write-Verbose "Using authentication method: $AuthMethod"

    $contextInfo = Get-MgContext -ErrorAction SilentlyContinue
    $reconnect = $false

    if ($contextInfo) {
        if ($AuthMethod -eq 'Interactive') {
            $currentScopes = $contextInfo.Scopes
            $missingScopes = $RequiredScopes | Where-Object { $_ -notin $currentScopes }
            if ($missingScopes) {
                Write-Verbose "Missing required scopes; reconnecting."
                $reconnect = $true
            } else {
                Write-Verbose "Already connected with required scopes."
                return $contextInfo
            }
        } else {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            $reconnect = $true
        }
    } else {
        $reconnect = $true
    }

    if ($reconnect) {
        try {
            switch ($AuthMethod) {
                'Interactive' {
                    $p = @{ Scopes = $RequiredScopes; NoWelcome = $true }
                    if ($TenantId) { $p.TenantId = $TenantId }
                    if ($ClientId) { $p.ClientId = $ClientId }
                    Connect-MgGraph @p
                }
                'ClientSecret' {
                    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential (New-Object System.Management.Automation.PSCredential($ClientId, $ClientSecret)) -NoWelcome
                }
                'Certificate' {
                    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
                }
                'Identity' {
                    $p = @{ Identity = $true; NoWelcome = $true }
                    if ($TenantId) { $p.TenantId = $TenantId }
                    Connect-MgGraph @p
                }
                'AccessToken' {
                    Connect-MgGraph -AccessToken $AccessToken -NoWelcome
                }
            }
            $newContext = Get-MgContext
            if ($newContext) { return $newContext }
            throw 'Connection attempt completed but unable to confirm connection'
        } catch {
            Write-Error "Error connecting to Microsoft Graph: $_"
            return $null
        }
    }
    return $contextInfo
}
