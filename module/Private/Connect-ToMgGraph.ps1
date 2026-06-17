# Private: Connect to Microsoft Graph (Interactive, ClientSecret, Certificate, Identity, AccessToken)

function Clear-RKMsalTokenCache {
    <#
        Wipes the persistent MSAL token cache used by Microsoft.Graph.Authentication so
        the next Connect-MgGraph is forced into a fresh interactive sign-in. Disconnect-MgGraph
        only clears in-process state; without this, MSAL silently reuses a refresh token
        whose scope set predates any newly added module scopes.
    #>
    [CmdletBinding()]
    param()
    $cleared = 0
    $filePaths = @(
        "$HOME/.IdentityService/msal.cache"
        "$HOME/.local/share/.IdentityService/msal.cache"
        "$HOME/.config/.IdentityService/msal.cache"
    )
    if ($env:LOCALAPPDATA) { $filePaths += "$env:LOCALAPPDATA/.IdentityService/msal.cache" }
    foreach ($p in $filePaths) {
        if ($p -and (Test-Path $p)) {
            try { Remove-Item $p -Force -ErrorAction Stop; $cleared++; Write-Verbose "Removed $p" } catch { Write-Verbose "Could not remove $p : $($_.Exception.Message)" }
        }
    }
    if ($IsMacOS) {
        # Mg SDK on macOS persists the MSAL cache to the login Keychain. Multiple item
        # names are possible across SDK versions; loop each known service name until
        # `security` reports nothing left to delete.
        $services = @('Microsoft.Developer.IdentityService', 'com.microsoft.adalcache', 'com.microsoft.identitymodel.adalcache', 'MSALCache')
        foreach ($svc in $services) {
            do {
                & security delete-generic-password -s $svc 2>$null | Out-Null
                $deleted = ($LASTEXITCODE -eq 0)
                if ($deleted) { $cleared++; Write-Verbose "Removed Keychain item service=$svc" }
            } while ($deleted)
        }
    }
    return $cleared
}

function Invoke-RKConsentGrantRepair {
    <#
        When Connect-MgGraph completes but the resulting token is missing scopes the
        caller asked for, the most common cause is the Mg SDK reusing a cached refresh
        token whose oauth2PermissionGrant predates the scope addition. If the current
        session holds DelegatedPermissionGrant.ReadWrite.All this function PATCHes the
        user's existing grant to add the missing scopes, wipes the MSAL token cache,
        and triggers a fresh interactive sign-in. Returns $true on success, $false
        when the gap could not be closed automatically.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string[]] $MissingScopes,
        [Parameter(Mandatory)] [string[]] $RequiredScopes,
        [Parameter(Mandatory)] [PSCustomObject] $CurrentContext
    )

    if ('DelegatedPermissionGrant.ReadWrite.All' -notin $CurrentContext.Scopes) {
        Write-Warning "Connected, but missing scopes after sign-in: $($MissingScopes -join ', ')"
        Write-Warning "Auto-repair needs DelegatedPermissionGrant.ReadWrite.All (not present in your session). Options to fix:"
        Write-Warning "  1) Have a Global Admin grant admin consent for those scopes on the 'Microsoft Graph Command Line Tools' enterprise app, OR"
        Write-Warning "  2) Reconnect with Connect-MgGraph -Scopes DelegatedPermissionGrant.ReadWrite.All,$($MissingScopes -join ',') -UseDeviceCode"
        return $false
    }

    Write-Host "Detected scope gap after sign-in; attempting auto-repair of consent grant..." -ForegroundColor Yellow

    try {
        $clientId = $CurrentContext.ClientId
        $sp = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$clientId'&`$select=id,displayName").value | Select-Object -First 1
        if (-not $sp) { Write-Warning "Could not locate service principal for clientId $clientId"; return $false }

        $meId = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me?$select=id').id
        $grant = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=clientId eq '$($sp.id)' and principalId eq '$meId'").value | Select-Object -First 1
        if (-not $grant) { Write-Warning "No existing consent grant found to patch."; return $false }

        $current = ($grant.scope -split '\s+') | Where-Object { $_ }
        $toAdd = $MissingScopes | Where-Object { $_ -notin $current }
        if ($toAdd) {
            $newScope = ($current + $toAdd) -join ' '
            Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$($grant.id)" -Body @{ scope = $newScope } | Out-Null
            Write-Host "  Added to consent grant: $($toAdd -join ', ')" -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "Consent grant patch failed: $($_.Exception.Message)"
        return $false
    }

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    [void](Clear-RKMsalTokenCache)
    Write-Host "Reconnecting (browser will open for fresh interactive sign-in)..." -ForegroundColor Yellow

    try {
        $p = @{ Scopes = $RequiredScopes; NoWelcome = $true }
        if ($CurrentContext.TenantId) { $p.TenantId = $CurrentContext.TenantId }
        if ($CurrentContext.ClientId) { $p.ClientId = $CurrentContext.ClientId }
        Connect-MgGraph @p
        $verify = Get-MgContext
        $stillMissing = $RequiredScopes | Where-Object { $_ -notin $verify.Scopes }
        if ($stillMissing) {
            Write-Warning "Auto-repair completed, but scopes still missing: $($stillMissing -join ', '). Close this PowerShell session and try again from a fresh terminal."
            return $false
        }
        return $true
    } catch {
        Write-Warning "Reconnect after consent patch failed: $($_.Exception.Message)"
        return $false
    }
}

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
        [switch] $SkipAutoRepair,

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
            if ($newContext) {
                # Detect the MSAL silent-reuse trap: Connect-MgGraph "succeeded" but the
                # resulting token is missing scopes we explicitly asked for. Auto-repair
                # by PATCHing the consent grant + wiping the token cache, unless opted out.
                if ($AuthMethod -eq 'Interactive' -and -not $SkipAutoRepair) {
                    $stillMissing = $RequiredScopes | Where-Object { $_ -notin $newContext.Scopes }
                    if ($stillMissing) {
                        if (Invoke-RKConsentGrantRepair -MissingScopes $stillMissing -RequiredScopes $RequiredScopes -CurrentContext $newContext) {
                            $newContext = Get-MgContext
                        }
                    }
                }
                return $newContext
            }
            throw 'Connection attempt completed but unable to confirm connection'
        } catch {
            Write-Error "Error connecting to Microsoft Graph: $_"
            return $null
        }
    }
    return $contextInfo
}
