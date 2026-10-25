function Get-AllDeviceData {
    [CmdletBinding()]
    param(
        # Entra member users (id, userPrincipalName), for the logged-on users lookup.
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Users,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $TenantName,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.Collections.Generic.List[object]] $Warnings
    )

    function Get-OperatingSystemProductType {
        param (
            $Customer
        )

        @{
            "0"   = "unknown"
            "4"   = "Windows 10/11 Enterprise"
            "27"  = "Windows 10/11 Enterprise N"
            "48"  = "Windows 10/11 Professional"
            "49"  = "Windows 10/11 Professional for workstation N"
            "72"  = "Windows 10/11 Enterprise Evaluation"
            "119" = "Windows 10 TeamOS"
            "121" = "Windows 10/11 Education"
            "122" = "Windows 10/11 Education N"
            "125" = "Windows 10 Enterprise LTSC"
            "136" = "Hololens"
            "175" = "Windows 10 / 11 Enterprise Multi-session"
        }.$Customer
    }

    function Get-OSFriendlyName {
        param (
            [string]$OperatingSystemVersion
        )

        switch -Regex ($OperatingSystemVersion) {
            "^10\.0\.19043" { return "Windows 10 21H1" }
            "^10\.0\.19044" { return "Windows 10 21H2" }
            "^10\.0\.19045" { return "Windows 10 22H2" }
            "^10\.0\.22000" { return "Windows 11 21H2" }
            "^10\.0\.22621" { return "Windows 11 22H2" }
            "^10\.0\.22631" { return "Windows 11 23H2" }
            "^10\.0\.22635" { return "Windows 11 23H2 Insider Preview" }
            "^10\.0\.261" { return "Windows 11 24H2" }
            "^10\.0\.262" { return "Windows 11 25H2" }
            default { return "Other" }
        }
    }

    function Convert-Size {
        [cmdletbinding()]
        param(
            [validateset("Bytes", "KB", "MB", "GB", "TB")]
            [string]$From,
            [validateset("Bytes", "KB", "MB", "GB", "TB")]
            [string]$To,
            [Parameter(Mandatory = $true)]
            [double]$Value,
            [int]$Precision = 4
        )
        switch ($From) {
            "Bytes" { $value = $Value }
            "KB" { $value = $Value * 1024 }
            "MB" { $value = $Value * 1024 * 1024 }
            "GB" { $value = $Value * 1024 * 1024 * 1024 }
            "TB" { $value = $Value * 1024 * 1024 * 1024 * 1024 }
        }

        switch ($To) {
            "Bytes" { return $value }
            "KB" { $Value = $Value / 1KB }
            "MB" { $Value = $Value / 1MB }
            "GB" { $Value = $Value / 1GB }
            "TB" { $Value = $Value / 1TB }

        }

        $Calc = [Math]::Round($value, $Precision, [MidPointRounding]::AwayFromZero)
        return "$calc $to"

    }

    # Optimized Properties List - Only essential properties for better performance
    $Properties = @(
        'Id',                 # Required for compliance data fetching and unique identification
        'DeviceName',
        'azureADDeviceId',           # Canonical casing - the join key for BitLocker keys + LAPS credentials
        'azureActiveDirectoryDeviceId', # Legacy fallback property; same GUID on modern tenants
        'ManagedDeviceOwnerType',
        'UserPrincipalName',  # Primary user
        'SerialNumber',
        'ManagedDeviceName',
        'Manufacturer',
        'Model',
        'ProcessorArchitecture',
        'WiFiMacAddress',
        'EthernetMacAddress',
        'TotalStorageSpaceInBytes',
        'FreeStorageSpaceInBytes',
        'EnrolledDateTime',
        'LastSyncDateTime',
        'EnrollmentProfileName',
        'IsEncrypted',
        'DeviceEnrollmentType',
        'OperatingSystem',
        'OSVersion',
        'ComplianceState',
        'usersLoggedOn',      # Contains userId for logged-on users
        'hardwareInformation', # Contains nested properties like tpmVersion, OS details, BiosVersion
        'managementAgent', # Indicates the management agent used (e.g., Intune)
        'skuFamily' # OS edition (Pro, Enterprise, Home, etc.) - more reliable than hardwareInformation.operatingSystemEdition
    )

    $swTotal = [System.Diagnostics.Stopwatch]::StartNew()

    # Get all Windows Devices from Microsoft Intune
    $swStep = [System.Diagnostics.Stopwatch]::StartNew()
    $AllDeviceData = @(Invoke-RKCoreGraphRequest -Uri "beta/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'&`$select=$($Properties -join ',')" -All)
    #filter out managed by MDE
    $AllDeviceData = @($AllDeviceData | Where-Object { $_.managementAgent -ne "msSense" })
    $swStep.Stop()
    Write-Verbose ("[Get-AllDeviceData] Managed devices fetch: {0:N2}s ({1} devices)" -f $swStep.Elapsed.TotalSeconds, $AllDeviceData.Count)

    # Get all AutoPilot registered devices under "Enrollment"
    Write-Verbose "Fetching Autopilot devices..."
    $swStep.Restart()
    $AutopilotDevices = @(Invoke-RKCoreGraphRequest -Uri 'beta/deviceManagement/windowsAutopilotDeviceIdentities' -All)
    $swStep.Stop()
    Write-Verbose ("[Get-AllDeviceData] Autopilot devices fetch: {0:N2}s ({1} devices)" -f $swStep.Elapsed.TotalSeconds, $AutopilotDevices.Count)

    # Pre-build Autopilot lookup hashtable (serialNumber -> device object) for O(1) lookups
    $AutopilotLookup = @{}
    foreach ($ap in $AutopilotDevices) {
        if ($ap.serialNumber) { $AutopilotLookup[[string]$ap.serialNumber] = $ap }
    }

    # Autopilot device preparation (v2) associates devices via corporate identifiers
    # ("Manufacturer,Model,Serial" or a bare serial) instead of a hardware hash.
    $CorporateIdentifierSerials = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $CorporateIdentifiers = @(Invoke-RKCoreGraphRequest -Uri 'beta/deviceManagement/importedDeviceIdentities' -All)
    foreach ($ci in $CorporateIdentifiers) {
        if ($ci.importedDeviceIdentifier) { [void]$CorporateIdentifierSerials.Add(($ci.importedDeviceIdentifier -split ',')[-1].Trim()) }
    }

    # Only devices that BitLocker or Autopilot would flag need their deviceType: the list endpoint always reports
    # 'windowsRT', and only a per-device GET returns 'cloudPC' (Windows 365 and Dev Box).
    $swStep.Restart()
    $candidates = @($AllDeviceData | Where-Object {
            -not $_.isEncrypted -or (-not $AutopilotLookup.ContainsKey([string]$_.serialNumber) -and -not ($_.serialNumber -and $CorporateIdentifierSerials.Contains([string]$_.serialNumber)))
        })
    $CloudPcIds = [System.Collections.Generic.HashSet[string]]::new()
    $deviceTypeRequests = @(foreach ($d in $candidates) { [pscustomobject]@{ Id = "dt:$($d.id)"; Url = "beta/deviceManagement/managedDevices/$($d.id)?`$select=id,deviceType" } })
    $failedLookups = 0
    foreach ($resp in @(Invoke-RKCoreGraphBatch -Requests $deviceTypeRequests)) {
        if ($resp.Status -ne 200) { $failedLookups++; continue }
        if ($resp.Body.deviceType -eq 'cloudPC') { [void]$CloudPcIds.Add(($resp.Id -replace '^dt:', '')) }
    }
    if ($failedLookups -gt 0) {
        foreach ($area in 'BitLocker', 'Autopilot') { $Warnings.Add([pscustomobject]@{ Area = $area; Message = "The device type of $failedLookups device(s) could not be read; they are treated as physical devices, so a Cloud PC among them can be listed." }) }
    }
    $swStep.Stop()
    Write-Verbose ("[Get-AllDeviceData] Device type lookup: {0:N2}s ({1} candidates, {2} Cloud PCs)" -f $swStep.Elapsed.TotalSeconds, $candidates.Count, $CloudPcIds.Count)

    # Pre-build user lookup hashtable (id -> UPN) for O(1) lookups
    $UserLookup = @{}
    foreach ($u in $Users) {
        if ($u.id -and $u.userPrincipalName) { $UserLookup[$u.id] = $u.userPrincipalName }
    }

    # Pre-fetch compliance rules for all noncompliant devices via Graph $batch.
    # Two batched passes replace 1+N per-device round trips: first the policy
    # state list per device, then the settingStates per nonCompliant policy.
    $ComplianceRulesByDevice = @{}
    $NonCompliantDevices = $AllDeviceData | Where-Object { $_.complianceState -eq 'noncompliant' }
    if ($NonCompliantDevices -and $NonCompliantDevices.Count -gt 0) {
        # Module-scoped map: reset so a previous tenant's day count never leaks into this report.
        $script:ComplianceSettingFriendlyNames['RequireRemainContact'] = 'Is active (no compliance check-in within the compliance status validity period)'
        try {
            $checkinDays = (Invoke-RKCoreGraphRequest -Uri 'beta/deviceManagement/settings').deviceComplianceCheckinThresholdDays
            if ($checkinDays) { $script:ComplianceSettingFriendlyNames['RequireRemainContact'] = "Is active (no compliance check-in in the last $checkinDays days)" }
        } catch {
            $Warnings.Add([pscustomobject]@{ Area = 'Compliance'; Message = "The compliance status validity period could not be read, so its reason does not show the number of days: $($_.Exception.Message)" })
        }
        Write-Verbose "Fetching compliance details for $($NonCompliantDevices.Count) noncompliant devices..."

        $policyStateRequests = foreach ($d in $NonCompliantDevices) {
            [PSCustomObject]@{
                Id  = "ps:$($d.id)"
                Url = "beta/deviceManagement/managedDevices/$($d.id)/deviceCompliancePolicyStates"
            }
        }

        $swStep.Restart()
        $policyStateResponses = Invoke-RKCoreGraphBatch -Requests @($policyStateRequests)
        $swStep.Stop()
        Write-Verbose ("[Get-AllDeviceData] Compliance policy-states batch: {0:N2}s ({1} devices, {2} responses)" -f $swStep.Elapsed.TotalSeconds, $NonCompliantDevices.Count, $policyStateResponses.Count)

        # Map deviceId -> array of nonCompliant/Error policy state ids (preserving the existing <=10 guard).
        # Graph occasionally returns the same policy-state id twice for a single device (e.g. when the
        # same policy resolves via more than one assignment path). Adding both would produce duplicate
        # sub-request ids in the batch, which Graph rejects with 400 "Request Id ... has to be unique
        # in a batch." Dedupe the state list per device by id BEFORE the <=10 guard so duplicates
        # don't accidentally push a device past the cap and silently drop its compliance details.
        $settingStatePairs = [System.Collections.Generic.List[object]]::new()
        foreach ($resp in $policyStateResponses) {
            if ($resp.Status -ne 200 -or -not $resp.Body) { continue }
            $deviceId = $resp.Id -replace '^ps:', ''
            $rawStates = @($resp.Body.value | Where-Object { $_.State -eq 'nonCompliant' -or $_.State -eq 'Error' })

            $seenStateIds = [System.Collections.Generic.HashSet[string]]::new()
            $states = [System.Collections.Generic.List[object]]::new()
            foreach ($s in $rawStates) {
                if ($seenStateIds.Add($s.id)) { $states.Add($s) }
            }

            if ($states.Count -eq 0 -or $states.Count -gt 10) { continue }
            $ComplianceRulesByDevice[$deviceId] = [System.Collections.Generic.List[string]]::new()
            foreach ($s in $states) {
                $settingStatePairs.Add([PSCustomObject]@{
                        Id  = "ss:$deviceId|$($s.id)"
                        Url = "beta/deviceManagement/managedDevices/$deviceId/deviceCompliancePolicyStates/$($s.id)/settingStates"
                    })
            }
        }

        if ($settingStatePairs.Count -gt 0) {
            $swStep.Restart()
            $settingResponses = Invoke-RKCoreGraphBatch -Requests @($settingStatePairs)
            $swStep.Stop()
            Write-Verbose ("[Get-AllDeviceData] Compliance setting-states batch: {0:N2}s ({1} pairs)" -f $swStep.Elapsed.TotalSeconds, $settingStatePairs.Count)

            $failedPairs = [System.Collections.Generic.List[object]]::new()
            $pairIndex = @{}
            foreach ($p in $settingStatePairs) { $pairIndex[$p.Id] = $p }

            foreach ($resp in $settingResponses) {
                if ($resp.Status -ne 200 -or -not $resp.Body) {
                    if ($pairIndex.ContainsKey($resp.Id)) { $failedPairs.Add($pairIndex[$resp.Id]) }
                    continue
                }
                $deviceId = ($resp.Id -replace '^ss:', '') -split '\|' | Select-Object -First 1
                if (-not $ComplianceRulesByDevice.ContainsKey($deviceId)) { continue }
                $details = @($resp.Body.value | Where-Object { $_.state -match 'nonCompliant' })
                foreach ($det in $details) {
                    # settingName (e.g. Windows10CompliancePolicy.RtpEnabled) is always set and maps to a friendly name; `setting` is empty for Default Device Compliance Policy rows.
                    $rule = if ($det.settingName) { $det.settingName } else { $det.setting }
                    if ($rule) { $ComplianceRulesByDevice[$deviceId].Add($rule) }
                }
            }

            # Fallback: sub-requests that didn't return 200 from $batch get retried as
            # per-request GETs. Graph's $batch endpoint occasionally rejects this
            # setting-states sub-route with 400 even though the individual GET works.
            if ($failedPairs.Count -gt 0) {
                $swStep.Restart()
                foreach ($pair in $failedPairs) {
                    $deviceId = ($pair.Id -replace '^ss:', '') -split '\|' | Select-Object -First 1
                    if (-not $ComplianceRulesByDevice.ContainsKey($deviceId)) { continue }
                    try {
                        $direct = Invoke-RKCoreGraphRequest -Uri $pair.Url
                        $details = @($direct.value | Where-Object { $_.state -match 'nonCompliant' })
                        foreach ($det in $details) {
                            $rule = if ($det.settingName) { $det.settingName } else { $det.setting }
                            if ($rule) { $ComplianceRulesByDevice[$deviceId].Add($rule) }
                        }
                    } catch {
                        Write-Verbose "Compliance setting-states fallback: per-request GET failed for $($pair.Id): $($_.Exception.Message)"
                    }
                }
                $swStep.Stop()
                Write-Verbose ("[Get-AllDeviceData] Compliance setting-states fallback: {0:N2}s ({1} pairs retried individually)" -f $swStep.Elapsed.TotalSeconds, $failedPairs.Count)
            }
        }

        # Best effort: a noncompliant device without rules shows as Unknown on the Noncompliant tab, so count them.
        $devicesWithoutRules = @($NonCompliantDevices | Where-Object { -not ($ComplianceRulesByDevice.ContainsKey($_.id) -and $ComplianceRulesByDevice[$_.id].Count -gt 0) }).Count
        if ($devicesWithoutRules -gt 0) { $Warnings.Add([pscustomobject]@{ Area = 'Compliance'; Message = "The compliance reasons of $devicesWithoutRules device(s) are not available (they could not be read, or the device fails more than 10 policies); they show as Unknown." }) }
    }

    # Loop through all devices for device data
    $results = [System.Collections.Generic.List[PSObject]]::new()
    $totalDevices = $AllDeviceData.Count

    Write-Verbose "Processing $totalDevices devices..."
    $swStep.Restart()

    for ($i = 0; $i -lt $AllDeviceData.Count; $i++) {
        $DeviceData = $AllDeviceData[$i]

        try {
            # Use bulk-fetched data directly (no per-device re-fetch needed - same $select)
            $DeviceProperties = $DeviceData

            # Process Autopilot information via pre-built hashtable
            $AutopilotInfo = $AutopilotLookup[[string]$DeviceData.SerialNumber]
            $HashUploaded = $AutopilotLookup.ContainsKey([string]$DeviceData.SerialNumber)

            # Compliance rules were prefetched via $batch above; just look up.
            $FilteredForAlerting = @("DefaultDeviceCompliancePolicy.RequireDeviceCompliancePolicyAssigned", "DefaultDeviceCompliancePolicy.RequireRemainContact")
            $uniqueRules = @()
            if ($DeviceData.complianceState -eq 'noncompliant' -and $ComplianceRulesByDevice.ContainsKey($DeviceData.id)) {
                $uniqueRules = @($ComplianceRulesByDevice[$DeviceData.id] | Select-Object -Unique)
            }

            # Check if all logged in user ID's still exist in Microsoft Entra ID
            $LoggedInUsers = $DeviceProperties.usersLoggedOn.userId | Select-Object -Unique
            $ExistingLoggedInUsers = [System.Collections.Generic.List[string]]::new()

            if ($LoggedInUsers) {
                foreach ($user in $LoggedInUsers) {
                    if ($UserLookup.ContainsKey($user)) {
                        $ExistingLoggedInUsers.Add($UserLookup[$user])
                    }
                }
            }

            # Handle storage calculations with null checking
            $TotalStorageFormatted = if ($DeviceProperties.TotalStorageSpaceInBytes -and $DeviceProperties.TotalStorageSpaceInBytes -gt 0) {
                Convert-Size -From bytes -To GB -Value $DeviceProperties.TotalStorageSpaceInBytes -Precision 2
            } else {
                "N/A"
            }

            $FreeStorageFormatted = if ($DeviceProperties.FreeStorageSpaceInBytes -and $DeviceProperties.FreeStorageSpaceInBytes -gt 0) {
                Convert-Size -From bytes -To GB -Value $DeviceProperties.FreeStorageSpaceInBytes -Precision 2
            } else {
                "N/A"
            }

            # Access hardware information with null checking
            $hardwareInfo = $DeviceProperties.hardwareInformation

            # Read the Azure AD device id defensively. Graph normalises the property to
            # `azureADDeviceId` (capital AD); some object types lose that on dot-access,
            # and on a few records only the legacy `azureActiveDirectoryDeviceId` is set.
            $resolvedAadDeviceId = $null
            foreach ($propName in 'azureADDeviceId','azureAdDeviceId','AzureAdDeviceId','azureActiveDirectoryDeviceId','AzureActiveDirectoryDeviceId') {
                $p = $DeviceProperties.PSObject.Properties[$propName]
                if ($p -and $p.Value) { $resolvedAadDeviceId = [string]$p.Value; break }
            }

            $results.Add([PSCustomObject][ordered]@{
                Customer                   = $TenantName
                DeviceName                 = $DeviceProperties.DeviceName
                AzureAdDeviceId            = $resolvedAadDeviceId
                DeviceOwnership            = $DeviceProperties.ManagedDeviceOwnerType
                PrimaryUser                = if ($DeviceProperties.UserPrincipalName) { $DeviceProperties.UserPrincipalName } else { "None" }
                Serialnumber               = $DeviceProperties.SerialNumber
                DeviceManufacturer         = $DeviceProperties.Manufacturer
                DeviceModel                = $DeviceProperties.Model
                ProcessorArchitecture      = if ($hardwareInfo.processorArchitecture) { $hardwareInfo.processorArchitecture } else { $DeviceProperties.processorArchitecture }
                TPMversion                 = if ($hardwareInfo.tpmVersion) { $hardwareInfo.tpmVersion } else { "Unknown" }
                tpmSpecificationVersion    = if ($hardwareInfo.tpmSpecificationVersion) { $hardwareInfo.tpmSpecificationVersion } else { "Unknown" }
                WiFiMAC                    = $DeviceProperties.WiFiMacAddress
                EthernetMAC                = $DeviceProperties.EthernetMacAddress
                TotalStorage               = $TotalStorageFormatted
                FreeStorage                = $FreeStorageFormatted
                EnrolledDate               = $DeviceProperties.EnrolledDateTime
                LastContact                = $DeviceProperties.LastSyncDateTime
                DeviceHashUploaded         = $HashUploaded
                DeviceAssociated           = [bool]($DeviceData.SerialNumber -and $CorporateIdentifierSerials.Contains($DeviceData.SerialNumber))
                # Only looked up for devices that BitLocker or Autopilot would flag: $false means "not a Cloud PC, or not looked up".
                IsCloudPC                  = $CloudPcIds.Contains([string]$DeviceData.id)
                AutopilotGroupTag          = $AutopilotInfo.groupTag
                AutopilotAssignedUser      = if ($AutopilotInfo.userprincipalname) { $AutopilotInfo.userprincipalname } else { $null }
                EnrollmentProfile          = $DeviceProperties.EnrollmentProfileName
                Encrypted                  = $DeviceProperties.IsEncrypted
                DeviceEnrollmentType       = $DeviceProperties.DeviceEnrollmentType
                usersLoggedOnIds           = if ($ExistingLoggedInUsers) { $ExistingLoggedInUsers -join ', ' } else { "" }
                usersLoggedOnCount         = if ($LoggedInUsers) { $LoggedInUsers.Count } else { 0 }
                Operatingsystem            = $DeviceProperties.OperatingSystem
                OperatingSystemVersion     = $DeviceProperties.OSVersion
                OSFriendlyname             = Get-OSFriendlyName -OperatingSystemVersion $DeviceProperties.OSVersion
                OperatingSystemLanguage    = if ($hardwareInfo.operatingSystemLanguage) { $hardwareInfo.operatingSystemLanguage } else { "Unknown" }
                OperatingSystemEdition     = if ($DeviceProperties.skuFamily) { $DeviceProperties.skuFamily } elseif ($hardwareInfo.operatingSystemEdition) { $hardwareInfo.operatingSystemEdition } else { "Unknown" }
                operatingSystemProductType = if ($hardwareInfo.operatingSystemProductType) { Get-OperatingSystemProductType -Customer "$($hardwareInfo.operatingSystemProductType)" } else { "Unknown" }
                BiosVersion                = if ($hardwareInfo.systemManagementBIOSVersion) { $hardwareInfo.systemManagementBIOSVersion } else { "Unknown" }
                ComplianceStatus           = $DeviceProperties.ComplianceState
                # **FIX**: Use unique rules to prevent duplicates
                NoncompliantBasedOn        = (@($uniqueRules | ForEach-Object { ConvertTo-ComplianceSettingFriendlyName $_ }) | Select-Object -Unique) -join ', '
                NoncompliantAlert          = (@($uniqueRules | Where-Object { $_ -notin $FilteredForAlerting } | ForEach-Object { ConvertTo-ComplianceSettingFriendlyName $_ }) | Select-Object -Unique) -join ', '
            })
        } catch {
            $Warnings.Add([pscustomobject]@{ Area = 'General'; Message = "Device $($DeviceData.DeviceName) could not be processed and is left out of every category: $($_.Exception.Message)" })
            continue
        }
    }

    $swStep.Stop()
    Write-Verbose ("[Get-AllDeviceData] Per-device projection loop: {0:N2}s" -f $swStep.Elapsed.TotalSeconds)
    $swTotal.Stop()
    Write-Verbose ("[Get-AllDeviceData] TOTAL: {0:N2}s" -f $swTotal.Elapsed.TotalSeconds)

    Write-Verbose "Device processing completed!"
    Write-Verbose "Processed $($results.Count) devices out of $totalDevices total devices"
    return $results

}

function Get-IntuneAppData {
    <#
    .SYNOPSIS
        The Windows apps of the tenant with their install summary. Readable is $false (with Error) when the app list
        cannot be read for lack of permission (DeviceManagementApps.Read.All).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $TenantName,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.Collections.Generic.List[object]] $Warnings
    )

    $types = 'win32CatalogApp', 'windowsStoreApp', 'microsoftStoreForBusinessApp', 'officeSuiteApp', 'windowsMicrosoftEdgeApp',
        'windowsPhone81AppX', 'windowsPhone81StoreApp', 'windowsPhoneXAP', 'windowsAppX', 'windowsMobileMSI', 'windowsUniversalAppX',
        'webApp', 'windowsWebApp', 'winGetApp'
    $typeFilter = (@($types | ForEach-Object { "isof('microsoft.graph.$_')" }) + "(isof('microsoft.graph.win32LobApp') and not(isof('microsoft.graph.win32CatalogApp')))") -join ' or '
    $filter = "($typeFilter) and (microsoft.graph.managedApp/appAvailability eq null or microsoft.graph.managedApp/appAvailability eq 'lineOfBusiness' or isAssigned eq true)"
    try {
        $apps = @(Invoke-RKCoreGraphRequest -Uri "beta/deviceAppManagement/mobileApps?`$filter=$filter&`$orderby=displayName" -All)
    }
    catch {
        if ((Get-RKCoreGraphErrorStatus -ErrorRecord $_) -in 401, 403) {
            return [pscustomobject]@{ Readable = $false; Error = $_.Exception.Message; Apps = @() }
        }
        throw
    }

    $requests = @(foreach ($a in $apps) { [pscustomobject]@{ Id = [string]$a.id; Url = "beta/deviceAppManagement/mobileApps/$($a.id)/installSummary" } })
    $summaries = @{}
    $unread = 0
    foreach ($resp in @(Invoke-RKCoreGraphBatch -Requests $requests)) {
        # Per-app summaries are best effort: one app that answers 400 must not stop every test.
        if ($resp.Status -eq 200 -and $resp.Body) { $summaries[$resp.Id] = $resp.Body } else { $unread++ }
    }
    if ($unread -gt 0) { $Warnings.Add([pscustomobject]@{ Area = 'Apps'; Message = "The install summary of $unread app(s) could not be read; they are left out." }) }

    $rows = foreach ($a in $apps) {
        $s = $summaries[[string]$a.id]
        if (-not $s) { continue }
        $failed = [int]$s.failedDeviceCount
        $targeted = $failed + [int]$s.installedDeviceCount + [int]$s.notInstalledDeviceCount + [int]$s.pendingInstallDeviceCount
        [pscustomobject][ordered]@{
            Customer               = $TenantName
            Application            = $a.displayName
            Platform               = 'Windows'
            Version                = [string]$a.displayVersion
            AssignmentStatus       = [bool]$a.isAssigned
            FailedUserCount        = [int]$s.failedUserCount
            FailedDeviceCount      = $failed
            FailedDevicePercentage = if ($targeted -gt 0) { [math]::Round(100 * $failed / $targeted, 2) } else { 0 }
        }
    }
    [pscustomobject]@{ Readable = $true; Error = $null; Apps = @($rows) }
}

function Get-AutopilotProfilesInformation {
    @(Invoke-RKCoreGraphRequest -Uri 'beta/deviceManagement/windowsAutopilotDeploymentProfiles' -All)
}

# ----------------------------------------------------------------------------
# BitLocker key escrow & Windows LAPS backup discovery.
# Determines per-device whether a BitLocker policy and a Windows LAPS (Entra-
# backed) policy are actually applied (via assignment resolution incl. groups,
# all users/devices, and assignment filters), and cross-references with the
# recovery key / local credential directory endpoints to flag escrow gaps.
# Reads Graph through Invoke-RKCoreGraphRequest and Invoke-RKCoreGraphBatch, fills the
# $script:AllFilters cache that Initialize-RKIntuneCoreState creates, and uses
# Get-DetailedPolicyAssignments and Test-IntuneFilter from Core/IntunePolicyAssignment.ps1.
# ----------------------------------------------------------------------------

function Get-BitLockerLapsAssignmentContext {
    <#
        Bulk-fetches all Settings Catalog and legacy device configuration policies that
        target BitLocker or Windows LAPS, resolves their assignments, fetches assignment
        filters, primes group transitive member sets for groups referenced by those
        assignments, and pulls the BitLocker recovery key and LAPS local credential
        directories. Returns a single context object consumed by the per-device evaluators.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)] [switch] $DebugMode,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.Collections.Generic.List[object]] $Warnings
    )

    $context = [PSCustomObject]@{
        BitLockerPolicies        = [System.Collections.Generic.List[object]]::new()
        LapsPolicies             = [System.Collections.Generic.List[object]]::new()
        GroupDeviceAadIds        = @{}    # groupId -> HashSet<string> of azureAdDeviceId
        GroupUserUpns            = @{}    # groupId -> HashSet<string> of lowercased UPN
        GroupUserIds             = @{}    # groupId -> HashSet<string> of user object id
        BitLockerKeyDeviceIds    = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        # Both LAPS deviceLocalCredentialInfo.id and bitlockerRecoveryKey.deviceId are
        # the Azure AD device id (= managedDevice.azureADDeviceId). Verified against
        # live Graph data. Both joins go through that single GUID.
        LapsCredentialByDeviceId   = @{}  # azureAdDeviceId -> LAPS credential record
        LapsCredentialByDeviceName = @{}  # lowercased deviceName -> LAPS credential record (fallback join)
        # Safety-net for device records that lack azureADDeviceId in Intune: lookup
        # the Entra device by displayName, recover its deviceId, retry the join.
        EntraDeviceIdByName      = @{}    # lowercased deviceName -> azureAdDeviceId
        Filters                  = @{}    # filterId -> filter object (mirror of $script:AllFilters)
        # configurationPolicies with their settings (and settingDefinitions) attached,
        # retained so the deprecation walker can iterate every Settings Catalog policy
        # without a second Graph round-trip.
        ConfigurationPolicies    = @()
        # Scope-gated reads: $false (with the error) when the connection lacks the scope.
        BitLockerKeysReadable    = $true
        BitLockerKeysError       = $null
        LapsCredentialsReadable  = $true
        LapsCredentialsError     = $null
    }

    # 1. Prime assignment filters cache (used by Get-DetailedPolicyAssignments + Test-IntuneFilter).
    foreach ($f in @(Invoke-RKCoreGraphRequest -Uri 'beta/deviceManagement/assignmentFilters' -All)) { $script:AllFilters[$f.id] = $f }
    $context.Filters = $script:AllFilters

    # 2. Discover Settings Catalog policies that contain BitLocker or LAPS settings.
    #    Two-step pattern: list policies (small payload) then fetch /settings per policy. Avoids Graph
    #    hard-capping page size when $expand=settings is used on the list (observed ~1-2 policies per
    #    page on cosmos-backed tenants). The settings requests go through Invoke-RKCoreGraphBatch: the
    #    module batches them, the pack sends them one by one.
    Write-Verbose '  Discovering Settings Catalog policies (BitLocker / LAPS)...'
    $swCfg = [System.Diagnostics.Stopwatch]::StartNew()
    $configPolicies = @(Invoke-RKCoreGraphRequest -Uri 'beta/deviceManagement/configurationPolicies' -All)
    $swCfg.Stop()
    Write-Verbose ("[BitLockerLaps] configurationPolicies list: {0:N2}s ({1} policies)" -f $swCfg.Elapsed.TotalSeconds, $configPolicies.Count)

    if ($configPolicies.Count -gt 0) {
        $swSet = [System.Diagnostics.Stopwatch]::StartNew()
        $settingsRequests = foreach ($p in $configPolicies) {
            [PSCustomObject]@{
                Id  = "cps:$($p.id)"
                Url = "beta/deviceManagement/configurationPolicies/$($p.id)/settings?`$expand=settingDefinitions"
            }
        }
        $settingsUrlByPolicyId = @{}
        foreach ($r in $settingsRequests) { $settingsUrlByPolicyId[($r.Id -replace '^cps:', '')] = $r.Url }
        $settingsResponses = Invoke-RKCoreGraphBatch -Requests @($settingsRequests)
        $settingsByPolicyId = @{}
        foreach ($resp in $settingsResponses) {
            $policyId = $resp.Id -replace '^cps:', ''
            # Graph's $batch sometimes answers 400 for this sub-route although the single GET works: retry it directly (a failure there throws).
            if ($resp.Status -notin 200, 404) { $settingsByPolicyId[$policyId] = @(Invoke-RKCoreGraphRequest -Uri $settingsUrlByPolicyId[$policyId] -All); continue }
            $body = Get-RKCoreBatchBody -Response $resp -What "the settings of policy $policyId" -Area 'Settings', 'BitLocker', 'Laps' -Warnings $Warnings; if (-not $body) { continue }
            $settings = @($body.value)
            # $expand can make Graph return smaller pages: read the rest, so a policy with many settings is complete.
            if ($body.'@odata.nextLink') { $settings += @(Invoke-RKCoreGraphRequest -Uri $body.'@odata.nextLink' -All) }
            $settingsByPolicyId[$policyId] = $settings
        }
        # Attach settings back onto each policy so the existing loop below stays unchanged.
        foreach ($p in $configPolicies) {
            $s = $null
            if ($settingsByPolicyId.ContainsKey($p.id)) { $s = $settingsByPolicyId[$p.id] }
            if ($p.PSObject.Properties['settings']) { $p.settings = $s }
            else { $p | Add-Member -NotePropertyName settings -NotePropertyValue $s -Force }
        }
        $swSet.Stop()
        Write-Verbose ("[BitLockerLaps] Settings batch-fetch: {0:N2}s ({1} policies)" -f $swSet.Elapsed.TotalSeconds, $configPolicies.Count)
    }
    $context.ConfigurationPolicies = $configPolicies

    foreach ($policy in $configPolicies) {
        if (-not $policy.settings) { continue }
        $hasBitLocker = $false; $hasLaps = $false; $lapsBackupValue = $null
        foreach ($s in $policy.settings) {
            $defId = $null
            if ($s.settingInstance -and $s.settingInstance.settingDefinitionId) { $defId = [string]$s.settingInstance.settingDefinitionId }
            if (-not $defId) { continue }
            $defIdLower = $defId.ToLowerInvariant()
            if ($defIdLower.StartsWith($script:BitLockerSettingPrefix)) { $hasBitLocker = $true }
            if ($defIdLower.StartsWith($script:LapsSettingPrefix))      { $hasLaps = $true }
            if ($defIdLower -eq $script:LapsBackupDirectorySetting) {
                # Choice settings expose the chosen value via choiceSettingValue.value, suffixed with the choice integer.
                $val = $null
                if ($s.settingInstance.choiceSettingValue -and $s.settingInstance.choiceSettingValue.value) {
                    $val = [string]$s.settingInstance.choiceSettingValue.value
                }
                if ($val) {
                    if     ($val -match '_1$') { $lapsBackupValue = 1 }   # Azure AD (Entra)
                    elseif ($val -match '_2$') { $lapsBackupValue = 2 }   # On-prem AD
                    elseif ($val -match '_0$') { $lapsBackupValue = 0 }   # Disabled
                }
            }
        }
        $displayName = if ($policy.name) { $policy.name } else { $policy.displayName }
        if ($hasBitLocker) {
            $assignments = Get-DetailedPolicyAssignments -EntityType 'configurationPolicies' -EntityId $policy.id -PolicyName $displayName -DebugMode:$DebugMode -ThrowOnError
            $context.BitLockerPolicies.Add([PSCustomObject]@{
                Id          = $policy.id
                DisplayName = $displayName
                Source      = 'SettingsCatalog'
                Assignments = @($assignments)
            })
        }
        if ($hasLaps -and $lapsBackupValue -eq 1) {
            # Only Entra-backed LAPS policies are relevant to "can we read it from Entra?"
            $assignments = Get-DetailedPolicyAssignments -EntityType 'configurationPolicies' -EntityId $policy.id -PolicyName $displayName -DebugMode:$DebugMode -ThrowOnError
            $context.LapsPolicies.Add([PSCustomObject]@{
                Id              = $policy.id
                DisplayName     = $displayName
                Source          = 'SettingsCatalog'
                BackupDirectory = $lapsBackupValue
                Assignments     = @($assignments)
            })
        }
    }

    # 3. Discover legacy Endpoint Protection device configurations that include BitLocker settings.
    Write-Verbose '  Discovering legacy device configuration policies (BitLocker)...'
    $deviceConfigs = @(Invoke-RKCoreGraphRequest -Uri 'beta/deviceManagement/deviceConfigurations' -All)
    foreach ($cfg in $deviceConfigs) {
        $odataType = [string]$cfg.'@odata.type'
        if ($odataType -eq '#microsoft.graph.windows10EndpointProtectionConfiguration') {
            $assignments = Get-DetailedPolicyAssignments -EntityType 'deviceConfigurations' -EntityId $cfg.id -PolicyName $cfg.displayName -DebugMode:$DebugMode -ThrowOnError
            $context.BitLockerPolicies.Add([PSCustomObject]@{
                Id          = $cfg.id
                DisplayName = $cfg.displayName
                Source      = 'DeviceConfiguration'
                Assignments = @($assignments)
            })
        }
    }

    # 4. Discover Endpoint Security intents (BitLocker disk encryption template).
    #    The intent list is best effort; the assignments of a matching intent are read like every other policy's.
    Write-Verbose '  Discovering Endpoint Security intents (BitLocker)...'
    $intents = @()
    try { $intents = @(Invoke-RKCoreGraphRequest -Uri 'beta/deviceManagement/intents' -All) }
    catch { $Warnings.Add([pscustomobject]@{ Area = 'BitLocker'; Message = "The Endpoint Security intents could not be read, so a BitLocker policy made from an intent is not counted: $($_.Exception.Message)" }) }
    foreach ($intent in $intents) {
        $name = [string]$intent.displayName
        # No reliable filter for BitLocker template; match by name as a pragmatic discovery hint.
        if ($name -match '(?i)bitlocker' -or $name -match '(?i)disk\s*encryption') {
            $assignments = Get-DetailedPolicyAssignments -EntityType 'deviceManagement/intents' -EntityId $intent.id -PolicyName $name -DebugMode:$DebugMode -ThrowOnError
            $context.BitLockerPolicies.Add([PSCustomObject]@{
                Id          = $intent.id
                DisplayName = $name
                Source      = 'Intent'
                Assignments = @($assignments)
            })
        }
    }

    # 5. Collect every groupId referenced by any of those policies' assignments, then
    #    fetch transitive members for each group exactly once.
    $referencedGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($p in @($context.BitLockerPolicies + $context.LapsPolicies)) {
        foreach ($a in $p.Assignments) {
            if ($a.GroupId) { [void]$referencedGroupIds.Add([string]$a.GroupId) }
        }
    }
    Write-Verbose "  Resolving transitive members for $($referencedGroupIds.Count) group(s) used by BitLocker/LAPS assignments..."
    # Initialise empty sets up front so every referenced group has entries even if it was not found.
    foreach ($gid in $referencedGroupIds) {
        $context.GroupDeviceAadIds[$gid] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $context.GroupUserUpns[$gid]     = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $context.GroupUserIds[$gid]      = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }
    if ($referencedGroupIds.Count -gt 0) {
        $memberRequests = [System.Collections.Generic.List[object]]::new()
        foreach ($gid in $referencedGroupIds) {
            $memberRequests.Add([PSCustomObject]@{
                    Id  = "gd:$gid"
                    Url = "v1.0/groups/$gid/transitiveMembers/microsoft.graph.device?`$select=id,deviceId"
                })
            $memberRequests.Add([PSCustomObject]@{
                    Id  = "gu:$gid"
                    Url = "v1.0/groups/$gid/transitiveMembers/microsoft.graph.user?`$select=id,userPrincipalName"
                })
        }
        # transitiveMembers lives on v1.0, so the request URLs name that version.
        $memberResponses = Invoke-RKCoreGraphBatch -Requests @($memberRequests)
        # Any sub-response carrying @odata.nextLink overflows the first page (999 members).
        # Fall back to sequential paging for those groups only — rare in practice.
        $nextLinkFollowups = [System.Collections.Generic.List[object]]::new()
        foreach ($resp in $memberResponses) {
            $body = Get-RKCoreBatchBody -Response $resp -What "the members of group $($resp.Id.Substring(3))" -Area 'BitLocker', 'Laps' -Warnings $Warnings; if (-not $body) { continue }
            $isUser = $resp.Id.StartsWith('gu:')
            $gid    = $resp.Id.Substring(3)
            foreach ($m in @($body.value)) {
                if ($isUser) {
                    if ($m.id)                { [void]$context.GroupUserIds[$gid].Add([string]$m.id) }
                    if ($m.userPrincipalName) { [void]$context.GroupUserUpns[$gid].Add(([string]$m.userPrincipalName).ToLowerInvariant()) }
                } else {
                    if ($m.deviceId) { [void]$context.GroupDeviceAadIds[$gid].Add([string]$m.deviceId) }
                }
            }
            if ($body.'@odata.nextLink') {
                $nextLinkFollowups.Add([PSCustomObject]@{ Gid = $gid; IsUser = $isUser; NextUri = [string]$body.'@odata.nextLink' })
            }
        }
        foreach ($f in $nextLinkFollowups) {
            $rest = @(Invoke-RKCoreGraphRequest -Uri $f.NextUri -All)
            foreach ($m in $rest) {
                if ($f.IsUser) {
                    if ($m.id)                { [void]$context.GroupUserIds[$f.Gid].Add([string]$m.id) }
                    if ($m.userPrincipalName) { [void]$context.GroupUserUpns[$f.Gid].Add(([string]$m.userPrincipalName).ToLowerInvariant()) }
                } else {
                    if ($m.deviceId) { [void]$context.GroupDeviceAadIds[$f.Gid].Add([string]$m.deviceId) }
                }
            }
        }
    }

    # 6. BitLocker recovery keys. OS-volume key (volumeType == 1) is what protects the system drive.
    #    A missing scope (BitlockerKey.ReadBasic.All, 401/403) sets BitLockerKeysReadable instead of
    #    reading as an empty key list; any other error throws.
    Write-Verbose '  Fetching BitLocker recovery key index (OS volume)...'
    $bitLockerFetched = 0
    try {
        foreach ($k in @(Invoke-RKCoreGraphRequest -Uri 'v1.0/informationProtection/bitlocker/recoveryKeys?$select=id,deviceId,volumeType,createdDateTime' -All)) {
            $bitLockerFetched++
            if ([int]$k.volumeType -eq 1 -and $k.deviceId) { [void]$context.BitLockerKeyDeviceIds.Add([string]$k.deviceId) }
        }
        Write-Verbose "    Indexed $($context.BitLockerKeyDeviceIds.Count) device(s) with OS-volume BitLocker keys (from $bitLockerFetched key rotations)."
    }
    catch {
        if ((Get-RKCoreGraphErrorStatus -ErrorRecord $_) -in 401, 403) {
            $context.BitLockerKeysReadable = $false
            $context.BitLockerKeysError = $_.Exception.Message
        }
        else { throw }
    }

    # 7. LAPS local credentials directory. Entry's `id` field IS the Azure AD device id
    #    (verified against live Graph). Key the hashtable by it directly.
    #    A missing scope (DeviceLocalCredential.ReadBasic.All, 401/403) sets LapsCredentialsReadable.
    Write-Verbose '  Fetching directory/deviceLocalCredentials (Windows LAPS)...'
    $lapsFetched = 0
    try {
        foreach ($e in @(Invoke-RKCoreGraphRequest -Uri 'v1.0/directory/deviceLocalCredentials?$select=id,deviceName,lastBackupDateTime,refreshDateTime' -All)) {
            $lapsFetched++
            if ($e.id) { $context.LapsCredentialByDeviceId[[string]$e.id] = $e }
            if ($e.deviceName) {
                $context.LapsCredentialByDeviceName[([string]$e.deviceName).ToLowerInvariant()] = $e
            }
        }
        Write-Verbose "    Indexed $lapsFetched LAPS local credential record(s)."
    }
    catch {
        if ((Get-RKCoreGraphErrorStatus -ErrorRecord $_) -in 401, 403) {
            $context.LapsCredentialsReadable = $false
            $context.LapsCredentialsError = $_.Exception.Message
        }
        else { throw }
    }

    # 8. Safety-net deviceName -> deviceId map. Only used when a managedDevice record
    #    has no azureADDeviceId of its own (rare, but seen on partially-enrolled devices).
    Write-Verbose '  Building Entra device name -> deviceId safety map...'
    try {
        $entraDevices = @(Invoke-RKCoreGraphRequest -Uri 'v1.0/devices?$top=999&$select=deviceId,displayName' -All)
        foreach ($ed in $entraDevices) {
            if ($ed.deviceId -and $ed.displayName) {
                $context.EntraDeviceIdByName[([string]$ed.displayName).ToLowerInvariant()] = [string]$ed.deviceId
            }
        }
        Write-Verbose "    Mapped $($context.EntraDeviceIdByName.Count) Entra device(s) by displayName."
    }
    catch {
        foreach ($area in 'BitLocker', 'Laps') { $Warnings.Add([pscustomobject]@{ Area = $area; Message = "The Entra device list could not be read, so a device without an Entra device ID in Intune is not matched by name: $($_.Exception.Message)" }) }
    }

    return $context
}

function Get-RKIntuneAnomalyData {
    <#
    .SYNOPSIS
        Collects everything the Intune anomaly resolvers need, once. Required reads throw; reads that need an
        optional scope set a *Readable flag; best-effort reads add Warnings (Area, Message).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    Initialize-RKIntuneCoreState
    $warnings = [System.Collections.Generic.List[object]]::new()
    Write-Verbose 'Collecting Intune anomaly data'

    $organization = @(Invoke-RKCoreGraphRequest -Uri 'v1.0/organization?$select=id,displayName' -All)[0]
    $tenantName = [string]$organization.displayName
    $users = @(Invoke-RKCoreGraphRequest -Uri 'beta/users?$top=999&$select=id,userPrincipalName,userType,accountEnabled' -All | Where-Object { $_.userType -eq 'Member' })
    $devices = @(Get-AllDeviceData -Users $users -TenantName $tenantName -Warnings $warnings)
    $profiles = @(Get-AutopilotProfilesInformation)
    $context = Get-BitLockerLapsAssignmentContext -Warnings $warnings
    $apps = Get-IntuneAppData -TenantName $tenantName -Warnings $warnings

    [pscustomobject]@{
        TenantId                = [string]$organization.id
        TenantName              = $tenantName
        CollectedAt             = Get-Date
        Devices                 = $devices
        Users                   = $users
        AutopilotProfiles       = $profiles
        SecurityKeyContext      = $context
        Apps                    = $apps.Apps
        AppsReadable            = $apps.Readable
        AppsError               = $apps.Error
        BitLockerKeysReadable   = $context.BitLockerKeysReadable
        BitLockerKeysError      = $context.BitLockerKeysError
        LapsCredentialsReadable = $context.LapsCredentialsReadable
        LapsCredentialsError    = $context.LapsCredentialsError
        Warnings                = $warnings.ToArray()
    }
}
