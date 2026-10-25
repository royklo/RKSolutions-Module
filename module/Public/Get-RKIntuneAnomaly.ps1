<#
.SYNOPSIS
    Returns the Intune anomalies of the connected tenant as objects, one property per category.

.DESCRIPTION
    Collects the Windows devices, BitLocker and LAPS state and Settings Catalog policies once, reads the failed app installs from the Intune app install status report (top 50 by failed devices), and returns:
    BitLocker, Laps, Noncompliant, Inactive, DisabledPrimaryUser, MissingAutopilot, MultiUser, AppInstallFailures,
    DeprecatedSettings and OperatingSystemEditions (rows), plus Warnings and the BitLockerKeysReadable,
    LapsCredentialsReadable and AppsReadable flags. Get-IntuneAnomaliesReport renders the same object as HTML.
    Connect first with Connect-RKGraph.

.PARAMETER InactiveDays
    A device whose last Intune sync is older than this many days is inactive. Default 90.

.PARAMETER MaxLapsBackupAgeDays
    A LAPS backup older than this many days is stale. Default 60.

.PARAMETER ShowExcludedDevices
    Also return devices that a BitLocker or LAPS policy leaves out on purpose (exclude group or filter), as Info rows.

.EXAMPLE
    Get-RKIntuneAnomaly | Select-Object -ExpandProperty BitLocker | Export-Csv bitlocker.csv
#>
function Get-RKIntuneAnomaly {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [ValidateRange(1, 3650)] [int] $InactiveDays = 90,
        [ValidateRange(1, 3650)] [int] $MaxLapsBackupAgeDays = 60,
        [switch] $ShowExcludedDevices
    )

    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) { throw 'Not connected to Microsoft Graph. Run Connect-RKGraph first.' }

    $data = Get-RKIntuneAnomalyData
    # Apps come from the Intune install status report (Get-ApplicationFailures), not from Core.
    $appsReadable = $true
    $appFailures = @()
    try { $appFailures = @(Get-ApplicationFailures -TenantName $data.TenantName) }
    catch {
        if ((Get-RKCoreGraphErrorStatus -ErrorRecord $_) -notin 401, 403) { throw }
        $appsReadable = $false
    }
    $results = [ordered]@{
        BitLocker           = Resolve-IntuneBitLockerAnomalies -Data $data
        Laps                = Resolve-IntuneLapsAnomalies -Data $data -MaxBackupAgeDays $MaxLapsBackupAgeDays
        Noncompliant        = Resolve-IntuneNoncompliantDevices -Data $data
        Inactive            = Resolve-IntuneInactiveDevices -Data $data -Days $InactiveDays
        DisabledPrimaryUser = Resolve-IntuneDisabledPrimaryUsers -Data $data
        MissingAutopilot    = Resolve-IntuneMissingAutopilotDevices -Data $data
        MultiUser           = Resolve-IntuneMultiUserDevices -Data $data
        AppInstallFailures  = [pscustomobject]@{ Rows = $appFailures }
        DeprecatedSettings  = Resolve-IntuneDeprecatedAnomalies -Data $data
    }
    $out = [ordered]@{ TenantName = $data.TenantName; CollectedAt = $data.CollectedAt }
    foreach ($name in $results.Keys) { $out[$name] = @($results[$name].Rows) }
    # @(...) keeps an empty result an empty array; without it, zero rows would become $null.
    $out.BitLocker = @($out.BitLocker | Where-Object { $ShowExcludedDevices -or $_.Severity -ne 'Info' })
    $out.Laps = @($out.Laps | Where-Object { $ShowExcludedDevices -or $_.Severity -ne 'Info' })
    $out.OperatingSystemEditions = @($data.Devices | Select-Object Customer, DeviceName, PrimaryUser, OperatingSystemEdition, OSFriendlyname)
    $out.Warnings = @($results.Values | ForEach-Object { $_.Warnings } | Sort-Object Area, Message -Unique)
    $out.BitLockerKeysReadable = $data.BitLockerKeysReadable
    $out.LapsCredentialsReadable = $data.LapsCredentialsReadable
    $out.AppsReadable = $appsReadable
    [pscustomobject]$out
}
