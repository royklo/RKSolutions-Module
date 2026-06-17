<#
.SYNOPSIS
    Generates an interactive HTML report of Intune anomalies (app failures, multi-user devices, encryption, Autopilot, compliance, etc.).
    Connect first with Connect-RKGraph; this cmdlet uses the existing connection.
#>
function Get-IntuneAnomaliesReport {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)] [switch] $SendEmail,
    [Parameter(Mandatory = $false)] [string[]] $Recipient,
    [Parameter(Mandatory = $false)] [string] $From,
    [Parameter(Mandatory = $false)] [string] $ExportPath,
    # Surface devices that were *deliberately* excluded from BitLocker / LAPS policies
    # (via exclude group or assignment filter) as Info-severity rows. Off by default
    # so the report stays focused on actual anomalies; intentional exclusions are
    # admin choices, not gaps.
    [Parameter(Mandatory = $false)] [switch] $ShowExcludedDevices,
    [Parameter(Mandatory = $false)] [switch] $DebugMode
)

$ErrorActionPreference = 'Stop'
try {
    $swReport = [System.Diagnostics.Stopwatch]::StartNew()
    $swPhase  = [System.Diagnostics.Stopwatch]::new()

    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) { throw 'Not connected to Microsoft Graph. Run Connect-RKGraph first.' }

    $tenantInfo = Invoke-MgGraphRequest -Uri 'beta/organization' -Method Get -OutputType PSObject
    $tenantname = $tenantInfo.value[0].displayName

    $swPhase.Restart()
    $AllEntraIDUsers = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/beta/users/?`$select=id,userPrincipalName,userType,accountEnabled" | Where-Object { $_.UserType -eq 'Member' }
    $DisabledEntraUsers = $AllEntraIDUsers | Where-Object { $_.accountEnabled -eq $false } | Select-Object id, userPrincipalName, userType, accountEnabled
    $swPhase.Stop(); Write-Verbose ("[Report] Entra users fetch: {0:N2}s ({1} members)" -f $swPhase.Elapsed.TotalSeconds, $AllEntraIDUsers.Count)

    Write-Host 'Starting device data collection...' -ForegroundColor Yellow
    $swPhase.Restart()
    $DeviceData = Get-AllDeviceData
    $swPhase.Stop(); Write-Verbose ("[Report] Get-AllDeviceData: {0:N2}s" -f $swPhase.Elapsed.TotalSeconds)

    $swPhase.Restart()
    $AutopilotProfilesInformation = Get-AutopilotProfilesInformation
    $UserDrivenAutopilotProfiles = $AutopilotProfilesInformation | Where-Object { $_.outOfBoxExperienceSettings.deviceUsageType -eq 'SingleUser' }
    $swPhase.Stop(); Write-Verbose ("[Report] Autopilot profiles: {0:N2}s" -f $swPhase.Elapsed.TotalSeconds)

    Write-Host 'Resolving BitLocker / Windows LAPS policy assignments and key escrow state...' -ForegroundColor Yellow
    $swPhase.Restart()
    $SecurityKeyContext = Get-BitLockerLapsAssignmentContext -DebugMode:$DebugMode
    $swPhase.Stop(); Write-Verbose ("[Report] BitLocker/LAPS context: {0:N2}s" -f $swPhase.Elapsed.TotalSeconds)

    $swPhase.Restart()
    $Report_BitLockerStatus    = Resolve-IntuneBitLockerAnomalies   -Devices $DeviceData -Context $SecurityKeyContext -TenantName $tenantname -ShowExcludedDevices:$ShowExcludedDevices
    $Report_LapsStatus         = Resolve-IntuneLapsAnomalies        -Devices $DeviceData -Context $SecurityKeyContext -TenantName $tenantname -ShowExcludedDevices:$ShowExcludedDevices
    $Report_DeprecatedSettings = Resolve-IntuneDeprecatedAnomalies  -Context $SecurityKeyContext -TenantName $tenantname
    $swPhase.Stop(); Write-Verbose ("[Report] BitLocker + LAPS + Deprecated resolution: {0:N2}s" -f $swPhase.Elapsed.TotalSeconds)

    $swPhase.Restart()
    $Report_ApplicationFailureReport = Get-ApplicationFailures
    $swPhase.Stop(); Write-Verbose ("[Report] Get-ApplicationFailures: {0:N2}s" -f $swPhase.Elapsed.TotalSeconds)
    $Report_DevicesWithMultipleUsers = $DeviceData | Where-Object { $_.usersLoggedOnCount -gt 1 -and $_.EnrollmentProfile -in $UserDrivenAutopilotProfiles.displayName } | Select-Object Customer, DeviceName, PrimaryUser, EnrollmentProfile, usersLoggedOnCount, usersLoggedOnIds
    $Report_OperatingSystemEditionOverview = $DeviceData | Select-Object Customer, DeviceName, PrimaryUser, OperatingSystemEdition, OSFriendlyname
    # Not Encrypted is now folded into the BitLocker tab - unencrypted devices
    # show up there as 'Device not encrypted and no BitLocker policy assigned' or
    # 'BitLocker policy assigned but device not encrypted' so we don't need a
    # standalone tab anymore.
    $Report_DevicesWithoutAutopilotHash = $DeviceData | Where-Object { $_.DeviceHashUploaded -eq $false } | Select-Object Customer, DeviceName, PrimaryUser, Serialnumber, DeviceManufacturer, DeviceModel
    $Report_InactiveDevices = $DeviceData | Where-Object { $_.LastContact -lt (Get-Date).AddDays(-90) } | Select-Object Customer, DeviceName, PrimaryUser, Serialnumber, DeviceManufacturer, DeviceModel, LastContact
    $Report_DisabledPrimaryUsers = $DeviceData | Where-Object { $_.PrimaryUser -in $DisabledEntraUsers.userPrincipalName } | Select-Object Customer, DeviceName, PrimaryUser, Serialnumber, DeviceManufacturer, DeviceModel

    $Report_NoncompliantDevices = [System.Collections.Generic.List[PSObject]]::new()
    $NoncompliantDevicesRaw = $DeviceData | Where-Object { $_.ComplianceStatus -eq 'noncompliant' }
    foreach ($device in $NoncompliantDevicesRaw) {
        if ($device.NoncompliantBasedOn) {
            $reasons = $device.NoncompliantBasedOn -split ', '
            foreach ($reason in $reasons) {
                if ($reason.Trim()) {
                    $Report_NoncompliantDevices.Add([PSCustomObject]@{ Customer = $device.Customer; DeviceName = $device.DeviceName; PrimaryUser = $device.PrimaryUser; Serialnumber = $device.Serialnumber; DeviceManufacturer = $device.DeviceManufacturer; DeviceModel = $device.DeviceModel; ComplianceStatus = $device.ComplianceStatus; NoncompliantBasedOn = $reason.Trim(); NoncompliantAlert = $device.NoncompliantAlert })
                }
            }
        } else {
            $Report_NoncompliantDevices.Add([PSCustomObject]@{ Customer = $device.Customer; DeviceName = $device.DeviceName; PrimaryUser = $device.PrimaryUser; Serialnumber = $device.Serialnumber; DeviceManufacturer = $device.DeviceManufacturer; DeviceModel = $device.DeviceModel; ComplianceStatus = $device.ComplianceStatus; NoncompliantBasedOn = 'Unknown'; NoncompliantAlert = $device.NoncompliantAlert })
        }
    }

    $swPhase.Restart()
    New-IntuneAnomaliesHTMLReport -TenantName $tenantname -Report_ApplicationFailureReport $Report_ApplicationFailureReport -Report_DevicesWithMultipleUsers $Report_DevicesWithMultipleUsers -Report_DevicesWithoutAutopilotHash $Report_DevicesWithoutAutopilotHash -Report_InactiveDevices $Report_InactiveDevices -Report_NoncompliantDevices $Report_NoncompliantDevices -Report_OperatingSystemEditionOverview $Report_OperatingSystemEditionOverview -Report_DisabledPrimaryUsers $Report_DisabledPrimaryUsers -Report_BitLockerStatus $Report_BitLockerStatus -Report_LapsStatus $Report_LapsStatus -Report_DeprecatedSettings $Report_DeprecatedSettings -ExportPath $ExportPath
    $swPhase.Stop(); Write-Verbose ("[Report] HTML build: {0:N2}s" -f $swPhase.Elapsed.TotalSeconds)
    $swReport.Stop(); Write-Verbose ("[Report] TOTAL Get-IntuneAnomaliesReport: {0:N2}s" -f $swReport.Elapsed.TotalSeconds)

    $emailSent = $false
    if ($SendEmail -and $Recipient) {
        $subject = "$tenantname - Intune Anomalies Report"
        $bodyHtml = "<html><body style=`"font-family: Segoe UI, Arial, sans-serif;`"><h2>Intune Anomalies Report</h2><p>Attached is the latest Intune anomalies report for <strong>$tenantname</strong>.</p><p>Open the attached HTML in a browser for the interactive dashboard.</p><p style='color:#666;'>Generated by RKSolutions - please do not reply.</p></body></html>"
        $emailSent = Send-EmailWithAttachment -Recipient $Recipient -AttachmentPath $script:ExportPath -From $From -Subject $subject -BodyHtml $bodyHtml
        if ($emailSent) { Write-Host 'INFO: Email sent successfully.' -ForegroundColor Green } else { Write-Host 'ERROR: Failed to send email.' -ForegroundColor Red }
    }
    if ($SendEmail -and $emailSent -and (Test-Path -Path $script:ExportPath)) { Remove-Item -Path $script:ExportPath -Force; Write-Host 'INFO: Temporary report file deleted.' -ForegroundColor Green }
}
catch { Write-Error "Error: $_"; throw $_ }
finally {
    # Session left connected; use Disconnect-RKGraph when done.
}
}
