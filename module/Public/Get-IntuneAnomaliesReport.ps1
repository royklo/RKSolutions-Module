<#
.SYNOPSIS
    Generates an interactive HTML report of Intune anomalies (app failures, multi-user devices, encryption, Autopilot, compliance, etc.).
    Connect first with Connect-RKGraph, or pass auth parameters to this cmdlet.
#>
function Get-IntuneAnomaliesReport {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)] [string[]] $RequiredScopes = @('User.Read', 'DeviceManagementManagedDevices.Read.All', 'DeviceManagementConfiguration.Read.All', 'DeviceManagementServiceConfig.Read.All', 'DeviceManagementApps.Read.All', 'User.Read.All', 'Directory.Read.All', 'Mail.Send', 'CloudPC.Read.All'),
    [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')] [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')] [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')] [Parameter(Mandatory = $false, ParameterSetName = 'Identity')] [Parameter(Mandatory = $true, ParameterSetName = 'AccessToken')] [string] $TenantId,
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

    $tenantInfo = Invoke-MgGraphRequest -Uri 'beta/organization' -Method Get -OutputType PSObject
    $tenantname = $tenantInfo.value[0].displayName

    $AllEntraIDUsers = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/beta/users/?`$select=id,userPrincipalName,userType,accountEnabled" | Where-Object { $_.UserType -eq 'Member' }
    $DisabledEntraUsers = $AllEntraIDUsers | Where-Object { $_.accountEnabled -eq $false } | Select-Object id, userPrincipalName, userType, accountEnabled

    Write-Host 'Starting device data collection...' -ForegroundColor Yellow
    $DeviceData = Get-AllDeviceData
    $AutopilotProfilesInformation = Get-AutopilotProfilesInformation
    $UserDrivenAutopilotProfiles = $AutopilotProfilesInformation | Where-Object { $_.outOfBoxExperienceSettings.deviceUsageType -eq 'SingleUser' }

    $Report_ApplicationFailureReport = Get-ApplicationFailures
    $Report_DevicesWithMultipleUsers = $DeviceData | Where-Object { $_.usersLoggedOnCount -gt 1 -and $_.EnrollmentProfile -in $UserDrivenAutopilotProfiles.displayName } | Select-Object Customer, DeviceName, PrimaryUser, EnrollmentProfile, usersLoggedOnCount, usersLoggedOnIds
    $Report_OperationSystemEdtionOverview = $DeviceData | Select-Object Customer, DeviceName, PrimaryUser, OperatingSystemEdition, OSFriendlyname
    $Report_NotEncryptedDevices = $DeviceData | Where-Object { $_.Encrypted -eq $false } | Select-Object Customer, DeviceName, PrimaryUser, Serialnumber, DeviceManufacturer, DeviceModel
    $Report_DevicesWithoutAutopilotHash = $DeviceData | Where-Object { $_.DeviceHashUploaded -eq $false } | Select-Object Customer, DeviceName, PrimaryUser, Serialnumber, DeviceManufacturer, DeviceModel
    $Report_InactiveDevices = $DeviceData | Where-Object { $_.LastContact -lt (Get-Date).AddDays(-90) } | Select-Object Customer, DeviceName, PrimaryUser, Serialnumber, DeviceManufacturer, DeviceModel, LastContact
    $Report_DisabledPrimaryUsers = $DeviceData | Where-Object { $_.PrimaryUser -in $DisabledEntraUsers.userPrincipalName } | Select-Object Customer, DeviceName, PrimaryUser, Serialnumber, DeviceManufacturer, DeviceModel

    $Report_NoncompliantDevices = @()
    $NoncompliantDevicesRaw = $DeviceData | Where-Object { $_.ComplianceStatus -eq 'noncompliant' }
    foreach ($device in $NoncompliantDevicesRaw) {
        if ($device.NoncompliantBasedOn) {
            $reasons = $device.NoncompliantBasedOn -split ', '
            foreach ($reason in $reasons) {
                if ($reason.Trim()) {
                    $Report_NoncompliantDevices += [PSCustomObject]@{ Customer = $device.Customer; DeviceName = $device.DeviceName; PrimaryUser = $device.PrimaryUser; Serialnumber = $device.Serialnumber; DeviceManufacturer = $device.DeviceManufacturer; DeviceModel = $device.DeviceModel; ComplianceStatus = $device.ComplianceStatus; NoncompliantBasedOn = $reason.Trim(); NoncompliantAlert = $device.NoncompliantAlert }
                }
            }
        } else {
            $Report_NoncompliantDevices += [PSCustomObject]@{ Customer = $device.Customer; DeviceName = $device.DeviceName; PrimaryUser = $device.PrimaryUser; Serialnumber = $device.Serialnumber; DeviceManufacturer = $device.DeviceManufacturer; DeviceModel = $device.DeviceModel; ComplianceStatus = $device.ComplianceStatus; NoncompliantBasedOn = 'Unknown'; NoncompliantAlert = $device.NoncompliantAlert }
        }
    }

    New-IntuneAnomaliesHTMLReport -TenantName $tenantname -Report_ApplicationFailureReport $Report_ApplicationFailureReport -Report_DevicesWithMultipleUsers $Report_DevicesWithMultipleUsers -Report_NotEncryptedDevices $Report_NotEncryptedDevices -Report_DevicesWithoutAutopilotHash $Report_DevicesWithoutAutopilotHash -Report_InactiveDevices $Report_InactiveDevices -Report_NoncompliantDevices $Report_NoncompliantDevices -Report_OperationSystemEdtionOverview $Report_OperationSystemEdtionOverview -Report_DisabledPrimaryUsers $Report_DisabledPrimaryUsers -ExportPath $ExportPath

    if ($SendEmail -and $Recipient) {
        $subject = "$tenantname - Intune Anomalies Report"
        $bodyHtml = "<html><body style=`"font-family: Segoe UI, Arial, sans-serif;`"><h2>Intune Anomalies Report</h2><p>Attached is the latest Intune anomalies report for <strong>$tenantname</strong>.</p><p>Open the attached HTML in a browser for the interactive dashboard.</p><p style='color:#666;'>Generated by RKSolutions - please do not reply.</p></body></html>"
        $emailSent = Send-EmailWithAttachment -Recipient $Recipient -AttachmentPath $script:ExportPath -From $From -Subject $subject -BodyHtml $bodyHtml
        if ($emailSent) { Write-Host 'INFO: Email sent successfully.' -ForegroundColor Green } else { Write-Host 'ERROR: Failed to send email.' -ForegroundColor Red }
    }
    if ($SendEmail -and (Test-Path -Path $script:ExportPath)) { Remove-Item -Path $script:ExportPath -Force; Write-Host 'INFO: Temporary report file deleted.' -ForegroundColor Green }
}
catch { Write-Error "Error: $_"; throw $_ }
finally {
    # Session left connected; use Disconnect-RKGraph when done.
}
}
