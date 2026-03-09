# Private: Install Microsoft.Graph.Authentication (PS 5.1 vs 7+ version logic)
function Install-Requirements {
    [CmdletBinding()]
    param([switch]$DebugMode)

    $PsVersion = $PSVersionTable.PSVersion.Major
    $requiredModules = @('Microsoft.Graph.Authentication')

    foreach ($module in $requiredModules) {
        if ($PsVersion -ge 7) {
            $moduleInstalled = Get-Module -ListAvailable -Name $module
            if (-not $moduleInstalled) {
                if ($DebugMode) { Write-Host "Installing latest version of module: $module" -ForegroundColor Cyan }
                Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
            }
        } else {
            $moduleVersion = '2.25.0'
            $moduleInstalled = Get-Module -ListAvailable -Name $module | Where-Object { $_.Version -eq $moduleVersion }
            if (-not $moduleInstalled) {
                if ($DebugMode) { Write-Host "Installing module: $module version $moduleVersion" -ForegroundColor Cyan }
                Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -RequiredVersion $moduleVersion -SkipPublisherCheck
            }
        }
    }

    foreach ($module in $requiredModules) {
        if (-not (Get-Module -Name $module)) {
            if ($PsVersion -ge 7) {
                Import-Module -Name $module -Force -ErrorAction Stop
            } else {
                Import-Module -Name $module -RequiredVersion '2.25.0' -Force -ErrorAction Stop
            }
        }
    }
}
