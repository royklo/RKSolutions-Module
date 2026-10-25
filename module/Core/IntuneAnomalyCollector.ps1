function Get-AllDeviceData {
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
    $AllDeviceData = Invoke-graphRequestWithPaging -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'&`$select=$($Properties -join ',')"
    #filter out managed by MDE
    $AllDeviceData = $AllDeviceData | Where-Object { $_.managementAgent -ne "msSense" }
    $swStep.Stop()
    Write-Verbose ("[Get-AllDeviceData] Managed devices fetch: {0:N2}s ({1} devices)" -f $swStep.Elapsed.TotalSeconds, $AllDeviceData.Count)

    # The list endpoint always reports deviceType 'windowsRT'; only a per-device GET returns 'cloudPC'
    # (Windows 365 and Dev Box). A failed lookup counts as a physical device so nothing is hidden by mistake.
    $swStep.Restart()
    $CloudPcIds = [System.Collections.Generic.HashSet[string]]::new()
    $deviceTypeRequests = foreach ($d in $AllDeviceData) {
        [PSCustomObject]@{ Id = "dt:$($d.id)"; Url = "/deviceManagement/managedDevices/$($d.id)?`$select=id,deviceType" }
    }
    foreach ($resp in (Invoke-RKGraphBatch -Requests @($deviceTypeRequests) -Activity 'Device types')) {
        if ($resp.Status -eq 200 -and $resp.Body.deviceType -eq 'cloudPC') { [void]$CloudPcIds.Add(($resp.Id -replace '^dt:', '')) }
    }
    $swStep.Stop()
    Write-Verbose ("[Get-AllDeviceData] Device type lookup: {0:N2}s ({1} Cloud PCs)" -f $swStep.Elapsed.TotalSeconds, $CloudPcIds.Count)

    # Get all AutoPilot registered devices under "Enrollment"
    Write-Host "Fetching Autopilot devices..." -ForegroundColor Yellow
    $swStep.Restart()
    $AutopilotDevices = (Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities")
    $swStep.Stop()
    Write-Verbose ("[Get-AllDeviceData] Autopilot devices fetch: {0:N2}s ({1} devices)" -f $swStep.Elapsed.TotalSeconds, $AutopilotDevices.Count)

    # Pre-build Autopilot lookup hashtable (serialNumber -> device object) for O(1) lookups
    $AutopilotLookup = @{}
    foreach ($ap in $AutopilotDevices) {
        if ($ap.serialNumber) { $AutopilotLookup[$ap.serialNumber] = $ap }
    }

    # Autopilot device preparation (v2) associates devices via corporate identifiers
    # ("Manufacturer,Model,Serial" or a bare serial) instead of a hardware hash.
    $CorporateIdentifierSerials = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $CorporateIdentifiers = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/beta/deviceManagement/importedDeviceIdentities"
    foreach ($ci in $CorporateIdentifiers) {
        if ($ci.importedDeviceIdentifier) { [void]$CorporateIdentifierSerials.Add(($ci.importedDeviceIdentifier -split ',')[-1].Trim()) }
    }

    # Pre-build user lookup hashtable (id -> UPN) for O(1) lookups
    $UserLookup = @{}
    foreach ($u in $AllEntraIDUsers) {
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
            $checkinDays = (Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/beta/deviceManagement/settings' -OutputType PSObject -ErrorAction Stop).deviceComplianceCheckinThresholdDays
            if ($checkinDays) { $script:ComplianceSettingFriendlyNames['RequireRemainContact'] = "Is active (no compliance check-in in the last $checkinDays days)" }
        } catch {
            Write-Verbose "Compliance validity period lookup failed: $($_.Exception.Message)"
        }
        Write-Host "Fetching compliance details for $($NonCompliantDevices.Count) noncompliant devices..." -ForegroundColor Yellow

        $policyStateRequests = foreach ($d in $NonCompliantDevices) {
            [PSCustomObject]@{
                Id  = "ps:$($d.id)"
                Url = "/deviceManagement/managedDevices/$($d.id)/deviceCompliancePolicyStates"
            }
        }

        $swStep.Restart()
        $policyStateResponses = Invoke-RKGraphBatch -Requests @($policyStateRequests) -Activity "Compliance policy states"
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
                        Url = "/deviceManagement/managedDevices/$deviceId/deviceCompliancePolicyStates/$($s.id)/settingStates"
                    })
            }
        }

        if ($settingStatePairs.Count -gt 0) {
            $swStep.Restart()
            $settingResponses = Invoke-RKGraphBatch -Requests @($settingStatePairs) -Activity "Compliance setting states"
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
                        $direct = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta$($pair.Url)" -ErrorAction Stop
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
    }

    # Loop through all devices for device data
    $results = [System.Collections.Generic.List[PSObject]]::new()
    $totalDevices = $AllDeviceData.Count

    Write-Host "Processing $totalDevices devices..." -ForegroundColor Yellow
    $swStep.Restart()

    for ($i = 0; $i -lt $AllDeviceData.Count; $i++) {
        $DeviceData = $AllDeviceData[$i]
        $currentIndex = $i + 1

        # Calculate progress percentage
        $progressPercent = [math]::Round(($currentIndex / $totalDevices) * 100, 1)

        # Show progress bar instead of Write-Host
        Write-Progress -Activity "Processing Intune Devices" -Status "Processing device: $($DeviceData.DeviceName)" -CurrentOperation "$currentIndex of $totalDevices devices processed" -PercentComplete $progressPercent

        try {
            # Use bulk-fetched data directly (no per-device re-fetch needed - same $select)
            $DeviceProperties = $DeviceData

            # Process Autopilot information via pre-built hashtable
            $AutopilotInfo = $AutopilotLookup[$DeviceData.SerialNumber]
            $HashUploaded = $AutopilotLookup.ContainsKey($DeviceData.SerialNumber)

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
            Write-Warning "Error processing device $($DeviceData.DeviceName): $_"
            continue
        }
    }

    # Clear the progress bar when done
    Write-Progress -Activity "Processing Intune Devices" -Completed
    $swStep.Stop()
    Write-Verbose ("[Get-AllDeviceData] Per-device projection loop: {0:N2}s" -f $swStep.Elapsed.TotalSeconds)
    $swTotal.Stop()
    Write-Verbose ("[Get-AllDeviceData] TOTAL: {0:N2}s" -f $swTotal.Elapsed.TotalSeconds)

    Write-Host "Device processing completed!" -ForegroundColor Green
    Write-Host "Processed $($results.Count) devices out of $totalDevices total devices" -ForegroundColor Green

    # Debug output for hardware information availability
    $devicesWithoutHardwareInfo = $results | Where-Object { $_.TPMversion -eq "Unknown" }

    if ($devicesWithoutHardwareInfo.Count -gt 0) {
        Write-Host "Some devices are missing hardware information:" -ForegroundColor Yellow
        foreach ($device in $devicesWithoutHardwareInfo) {
            Write-Host " - $($device.DeviceName) (Serial: $($device.Serialnumber))" -ForegroundColor Gray
        }
    }
    return $results

}

function Get-ApplicationFailures {

    # Cross-platform temporary file path
    # Detect OS and set appropriate temp path
    $detectedWindows = $false

    # Check if automatic variables exist (PowerShell Core 6.0+)
    if (Get-Variable -Name "IsWindows" -ErrorAction SilentlyContinue) {
        $detectedWindows = $IsWindows
    }
    # Fallback for older PowerShell versions
    else {
        $osInfo = [System.Environment]::OSVersion.Platform
        switch ($osInfo) {
            "Win32NT" { $detectedWindows = $true }
            "Unix" {
                # macOS/Linux detected but not used in this context
            }
            default {
                try {
                    if ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)) {
                        $detectedWindows = $true
                    }
                    else {
                        # Non-Windows OS detected
                    }
                } catch {
                    $detectedWindows = $true
                }
            }
        }
    }

    # Use a unique temporary file to avoid race conditions
    $Data = [System.IO.Path]::GetTempFileName()

    $apps = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=(isof(%27microsoft.graph.win32CatalogApp%27)%20or%20isof(%27microsoft.graph.windowsStoreApp%27)%20or%20isof(%27microsoft.graph.microsoftStoreForBusinessApp%27)%20or%20isof(%27microsoft.graph.officeSuiteApp%27)%20or%20(isof(%27microsoft.graph.win32LobApp%27)%20and%20not(isof(%27microsoft.graph.win32CatalogApp%27)))%20or%20isof(%27microsoft.graph.windowsMicrosoftEdgeApp%27)%20or%20isof(%27microsoft.graph.windowsPhone81AppX%27)%20or%20isof(%27microsoft.graph.windowsPhone81StoreApp%27)%20or%20isof(%27microsoft.graph.windowsPhoneXAP%27)%20or%20isof(%27microsoft.graph.windowsAppX%27)%20or%20isof(%27microsoft.graph.windowsMobileMSI%27)%20or%20isof(%27microsoft.graph.windowsUniversalAppX%27)%20or%20isof(%27microsoft.graph.webApp%27)%20or%20isof(%27microsoft.graph.windowsWebApp%27)%20or%20isof(%27microsoft.graph.winGetApp%27))%20and%20(microsoft.graph.managedApp/appAvailability%20eq%20null%20or%20microsoft.graph.managedApp/appAvailability%20eq%20%27lineOfBusiness%27%20or%20isAssigned%20eq%20true)&`$orderby=displayName&").value

    $params = @{
        Select  = @(
            "DisplayName"
            "Publisher"
            "Platform"
            "AppVersion"
            "FailedDevicePercentage"
            "FailedDeviceCount"
            "FailedUserCount"
            "ApplicationId"
        )
        Skip    = 0
        Top     = 50
        Filter  = "(FailedDeviceCount gt '0')"
        OrderBy = @(
            "FailedDeviceCount desc"
        )
    }
    Invoke-MgGraphRequest -Body $params -Uri "https://graph.microsoft.com/beta/deviceManagement/reports/getAppsInstallSummaryReport" -Method POST -OutputFilePath $Data

    $DataFile = Get-Content $Data
    # Fix char encoding to UTF-8
    $Response = [system.Text.Encoding]::UTF8.GetString(($DataFile).ToCharArray()) | ConvertFrom-Json

    # Pre-build app lookup (id -> app) so the per-row joins are O(1) instead of O(N) Where-Object scans.
    $appLookup = @{}
    foreach ($a in $apps) { if ($a.Id) { $appLookup[[string]$a.Id] = $a } }

    # First pass: materialise the schema-flattened rows that have a matching app,
    # collecting unique ApplicationIds so we can batch-fetch their assignments.
    $rows = [System.Collections.Generic.List[hashtable]]::new()
    $expandIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($value in $Response.Values) {
        $LineObject = @{ }
        foreach ($prop in $Response.Schema) {
            $LineObject[$prop.Column] = $value[$Response.Schema.IndexOf($prop)]
        }
        $appId = [string]$LineObject.ApplicationId
        if ($appLookup.ContainsKey($appId)) {
            $rows.Add($LineObject)
            [void]$expandIds.Add($appId)
        }
    }

    # Batch-fetch ?$expand=assignments for every failed app in parallel sub-requests.
    $assignmentStatusByAppId = @{}
    if ($expandIds.Count -gt 0) {
        $expandRequests = foreach ($id in $expandIds) {
            [PSCustomObject]@{
                Id  = "app:$id"
                Url = "/deviceAppManagement/mobileApps/$id/?`$expand=assignments"
            }
        }
        $expandResponses = Invoke-RKGraphBatch -Requests @($expandRequests) -Activity 'App assignments'
        foreach ($resp in $expandResponses) {
            if ($resp.Status -ne 200 -or -not $resp.Body) { continue }
            $appId = $resp.Id -replace '^app:', ''
            $hasAssignments = $resp.Body.assignments -and (@($resp.Body.assignments).Count -gt 0)
            $assignmentStatusByAppId[$appId] = [bool]$hasAssignments
        }
    }

    # Build result array from the prepared rows using O(1) lookups.
    $ReturnObject = New-Object System.Collections.ArrayList
    foreach ($LineObject in $rows) {
        $appId = [string]$LineObject.ApplicationId
        $PlatformName = if ($LineObject.Platform_loc) { $LineObject.Platform_loc } else { $LineObject.Platform }
        $AssignmentStatus = $false
        if ($assignmentStatusByAppId.ContainsKey($appId)) { $AssignmentStatus = $assignmentStatusByAppId[$appId] }

        $ReturnObject.Add([PSCustomObject][ordered]@{
                Customer               = $tenantname
                Application            = $appLookup[$appId].displayName
                Platform               = $PlatformName
                Version                = $LineObject.AppVersion
                AssignmentStatus       = $AssignmentStatus
                FailedUserCount        = $LineObject.FailedUserCount
                FailedDeviceCount      = $LineObject.FailedDeviceCount
                FailedDevicePercentage = [math]::Round([double]$LineObject.FailedDevicePercentage, 2)
            }) | Out-Null
    }

    # Clean up temporary data file
    if (Test-Path -Path $Data) {
        Remove-Item -Path $Data -Force
    }

    return $ReturnObject | Sort-Object -Property FailedDeviceCount -Descending
}

function Get-AutopilotProfilesInformation {
(Invoke-GraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles/" -OutputType PSObject).value
}

# ----------------------------------------------------------------------------
# BitLocker key escrow & Windows LAPS backup discovery.
# Determines per-device whether a BitLocker policy and a Windows LAPS (Entra-
# backed) policy are actually applied (via assignment resolution incl. groups,
# all users/devices, and assignment filters), and cross-references with the
# recovery key / local credential directory endpoints to flag escrow gaps.
# Reuses Invoke-GraphRequestWithPaging, the $script:AllFilters cache established by
# IntuneEnrollmentFlows.ps1 (load order is enforced in RKSolutions.psm1), and
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
        [Parameter(Mandatory = $false)] [switch] $DebugMode
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
        # Raw configurationPolicies?$expand=settings result, retained so the
        # deprecation walker can iterate every Settings Catalog policy without
        # a second Graph round-trip.
        ConfigurationPolicies    = @()
    }

    # 1. Prime assignment filters cache (used by Get-DetailedPolicyAssignments + Test-IntuneFilter).
    if ($script:AllFilters.Count -eq 0) {
        try {
            $rawFilters = Invoke-GraphRequestWithPaging -Uri 'https://graph.microsoft.com/beta/deviceManagement/assignmentFilters'
            if ($rawFilters) { foreach ($f in $rawFilters) { $script:AllFilters[$f.id] = $f } }
        }
        catch { Write-Verbose "Failed to load assignment filters: $($_.Exception.Message)" }
    }
    $context.Filters = $script:AllFilters

    # 2. Discover Settings Catalog policies that contain BitLocker or LAPS settings.
    #    Two-step pattern: list policies (small payload, fits in 1-2 pages) then $batch-fetch
    #    /settings per policy. Avoids Graph hard-capping page size when $expand=settings is used
    #    (observed ~1-2 policies per page on cosmos-backed tenants); batch sub-requests execute
    #    server-side in parallel, so ~75 policies finish in ~4 batched round trips.
    Write-Host '  Discovering Settings Catalog policies (BitLocker / LAPS)...' -ForegroundColor Cyan
    $configPolicies = @()
    $swCfg = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $configPolicies = Invoke-GraphRequestWithPaging -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$top=999'
    }
    catch { Write-Warning "configurationPolicies list fetch failed: $($_.Exception.Message)" }
    $swCfg.Stop()
    Write-Verbose ("[BitLockerLaps] configurationPolicies list: {0:N2}s ({1} policies)" -f $swCfg.Elapsed.TotalSeconds, $configPolicies.Count)

    if ($configPolicies.Count -gt 0) {
        $swSet = [System.Diagnostics.Stopwatch]::StartNew()
        $settingsRequests = foreach ($p in $configPolicies) {
            [PSCustomObject]@{
                Id  = "cps:$($p.id)"
                Url = "/deviceManagement/configurationPolicies/$($p.id)/settings"
            }
        }
        $settingsResponses = Invoke-RKGraphBatch -Requests @($settingsRequests) -Activity 'Settings Catalog policy settings'
        $settingsByPolicyId = @{}
        foreach ($resp in $settingsResponses) {
            if ($resp.Status -ne 200 -or -not $resp.Body) { continue }
            $pid = $resp.Id -replace '^cps:', ''
            $settingsByPolicyId[$pid] = @($resp.Body.value)
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
            $assignments = Get-DetailedPolicyAssignments -EntityType 'configurationPolicies' -EntityId $policy.id -PolicyName $displayName -DebugMode:$DebugMode
            $context.BitLockerPolicies.Add([PSCustomObject]@{
                Id          = $policy.id
                DisplayName = $displayName
                Source      = 'SettingsCatalog'
                Assignments = @($assignments)
            })
        }
        if ($hasLaps -and $lapsBackupValue -eq 1) {
            # Only Entra-backed LAPS policies are relevant to "can we read it from Entra?"
            $assignments = Get-DetailedPolicyAssignments -EntityType 'configurationPolicies' -EntityId $policy.id -PolicyName $displayName -DebugMode:$DebugMode
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
    Write-Host '  Discovering legacy device configuration policies (BitLocker)...' -ForegroundColor Cyan
    try {
        $deviceConfigs = Invoke-GraphRequestWithPaging -Uri 'https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations'
        foreach ($cfg in $deviceConfigs) {
            $odataType = [string]$cfg.'@odata.type'
            if ($odataType -eq '#microsoft.graph.windows10EndpointProtectionConfiguration') {
                $assignments = Get-DetailedPolicyAssignments -EntityType 'deviceConfigurations' -EntityId $cfg.id -PolicyName $cfg.displayName -DebugMode:$DebugMode
                $context.BitLockerPolicies.Add([PSCustomObject]@{
                    Id          = $cfg.id
                    DisplayName = $cfg.displayName
                    Source      = 'DeviceConfiguration'
                    Assignments = @($assignments)
                })
            }
        }
    }
    catch { Write-Warning "deviceConfigurations fetch failed: $($_.Exception.Message)" }

    # 4. Discover Endpoint Security intents (BitLocker disk encryption template).
    Write-Host '  Discovering Endpoint Security intents (BitLocker)...' -ForegroundColor Cyan
    try {
        $intents = Invoke-GraphRequestWithPaging -Uri 'https://graph.microsoft.com/beta/deviceManagement/intents'
        foreach ($intent in $intents) {
            $name = [string]$intent.displayName
            # No reliable filter for BitLocker template; match by name as a pragmatic discovery hint.
            if ($name -match '(?i)bitlocker' -or $name -match '(?i)disk\s*encryption') {
                $assignments = Get-DetailedPolicyAssignments -EntityType 'deviceManagement/intents' -EntityId $intent.id -PolicyName $name -DebugMode:$DebugMode
                $context.BitLockerPolicies.Add([PSCustomObject]@{
                    Id          = $intent.id
                    DisplayName = $name
                    Source      = 'Intent'
                    Assignments = @($assignments)
                })
            }
        }
    }
    catch { Write-Verbose "intents fetch skipped: $($_.Exception.Message)" }

    # 5. Collect every groupId referenced by any of those policies' assignments, then
    #    fetch transitive members for each group exactly once.
    $referencedGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($p in @($context.BitLockerPolicies + $context.LapsPolicies)) {
        foreach ($a in $p.Assignments) {
            if ($a.GroupId) { [void]$referencedGroupIds.Add([string]$a.GroupId) }
        }
    }
    Write-Host "  Resolving transitive members for $($referencedGroupIds.Count) group(s) used by BitLocker/LAPS assignments..." -ForegroundColor Cyan
    # Initialise empty sets up front so every referenced group has entries even if a fetch fails.
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
                    Url = "/groups/$gid/transitiveMembers/microsoft.graph.device?`$select=id,deviceId"
                })
            $memberRequests.Add([PSCustomObject]@{
                    Id  = "gu:$gid"
                    Url = "/groups/$gid/transitiveMembers/microsoft.graph.user?`$select=id,userPrincipalName"
                })
        }
        # transitiveMembers lives on v1.0; route the batch there to keep sub-request URLs consistent.
        $memberResponses = Invoke-RKGraphBatch -Requests @($memberRequests) -GraphVersion 'v1.0' -Activity 'Group transitive members'
        # Any sub-response carrying @odata.nextLink overflows the first page (999 members).
        # Fall back to sequential paging for those groups only — rare in practice.
        $nextLinkFollowups = [System.Collections.Generic.List[object]]::new()
        foreach ($resp in $memberResponses) {
            if ($resp.Status -ne 200 -or -not $resp.Body) { continue }
            $isUser = $resp.Id.StartsWith('gu:')
            $gid    = $resp.Id.Substring(3)
            foreach ($m in @($resp.Body.value)) {
                if ($isUser) {
                    if ($m.id)                { [void]$context.GroupUserIds[$gid].Add([string]$m.id) }
                    if ($m.userPrincipalName) { [void]$context.GroupUserUpns[$gid].Add(([string]$m.userPrincipalName).ToLowerInvariant()) }
                } else {
                    if ($m.deviceId) { [void]$context.GroupDeviceAadIds[$gid].Add([string]$m.deviceId) }
                }
            }
            if ($resp.Body.'@odata.nextLink') {
                $nextLinkFollowups.Add([PSCustomObject]@{ Gid = $gid; IsUser = $isUser; NextUri = [string]$resp.Body.'@odata.nextLink' })
            }
        }
        foreach ($f in $nextLinkFollowups) {
            try {
                $rest = Invoke-GraphRequestWithPaging -Uri $f.NextUri
                foreach ($m in $rest) {
                    if ($f.IsUser) {
                        if ($m.id)                { [void]$context.GroupUserIds[$f.Gid].Add([string]$m.id) }
                        if ($m.userPrincipalName) { [void]$context.GroupUserUpns[$f.Gid].Add(([string]$m.userPrincipalName).ToLowerInvariant()) }
                    } else {
                        if ($m.deviceId) { [void]$context.GroupDeviceAadIds[$f.Gid].Add([string]$m.deviceId) }
                    }
                }
            }
            catch { Write-Verbose "Group $($f.Gid) nextLink follow-up failed: $($_.Exception.Message)" }
        }
    }

    # 6. BitLocker recovery keys. OS-volume key (volumeType == 1) is what protects the system drive.
    #    Use Invoke-MgGraphRequest directly (not the swallow-on-error paging helper) so a
    #    missing scope surfaces as a clear warning rather than an empty result set.
    Write-Host '  Fetching BitLocker recovery key index (OS volume)...' -ForegroundColor Cyan
    $bitLockerFetched = 0
    try {
        $bitlockerUri = 'https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys?$select=id,deviceId,volumeType,createdDateTime'
        do {
            $resp = Invoke-MgGraphRequest -Uri $bitlockerUri -Method GET -OutputType PSObject -ErrorAction Stop
            if ($resp.value) {
                foreach ($k in $resp.value) {
                    $bitLockerFetched++
                    if ([int]$k.volumeType -eq 1 -and $k.deviceId) { [void]$context.BitLockerKeyDeviceIds.Add([string]$k.deviceId) }
                }
            }
            $bitlockerUri = $resp.'@odata.nextLink'
        } while ($bitlockerUri)
        Write-Host "    Indexed $($context.BitLockerKeyDeviceIds.Count) device(s) with OS-volume BitLocker keys (from $bitLockerFetched key rotations)." -ForegroundColor Green
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match '403|Forbidden|Authorization|scopes') {
            Write-Warning "BitLocker recoveryKeys fetch DENIED. Most likely cause: the connected app is missing BitlockerKey.ReadBasic.All. Without it, every encrypted device will be flagged 'no key in Entra'. Reconnect with: Disconnect-RKGraph; Connect-RKGraph"
        }
        else { Write-Warning "BitLocker recoveryKeys fetch failed: $msg" }
    }

    # 7. LAPS local credentials directory. Entry's `id` field IS the Azure AD device id
    #    (verified against live Graph). Key the hashtable by it directly.
    Write-Host '  Fetching directory/deviceLocalCredentials (Windows LAPS)...' -ForegroundColor Cyan
    $lapsFetched = 0
    try {
        $lapsUri = 'https://graph.microsoft.com/v1.0/directory/deviceLocalCredentials?$select=id,deviceName,lastBackupDateTime,refreshDateTime'
        do {
            $resp = Invoke-MgGraphRequest -Uri $lapsUri -Method GET -OutputType PSObject -ErrorAction Stop
            if ($resp.value) {
                foreach ($e in $resp.value) {
                    $lapsFetched++
                    if ($e.id) { $context.LapsCredentialByDeviceId[[string]$e.id] = $e }
                    if ($e.deviceName) {
                        $context.LapsCredentialByDeviceName[([string]$e.deviceName).ToLowerInvariant()] = $e
                    }
                }
            }
            $lapsUri = $resp.'@odata.nextLink'
        } while ($lapsUri)
        Write-Host "    Indexed $lapsFetched LAPS local credential record(s)." -ForegroundColor Green
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match '403|Forbidden|Authorization|scopes') {
            Write-Warning "deviceLocalCredentials fetch DENIED. Most likely cause: the connected app is missing DeviceLocalCredential.ReadBasic.All. Without it, every device will be flagged 'no LAPS backup'. Reconnect with: Disconnect-RKGraph; Connect-RKGraph"
        }
        else { Write-Warning "deviceLocalCredentials fetch failed: $msg" }
    }

    # 8. Safety-net deviceName -> deviceId map. Only used when a managedDevice record
    #    has no azureADDeviceId of its own (rare, but seen on partially-enrolled devices).
    Write-Host '  Building Entra device name -> deviceId safety map...' -ForegroundColor Cyan
    try {
        $entraDevices = Invoke-GraphRequestWithPaging -Uri 'https://graph.microsoft.com/v1.0/devices?$select=deviceId,displayName'
        foreach ($ed in $entraDevices) {
            if ($ed.deviceId -and $ed.displayName) {
                $context.EntraDeviceIdByName[([string]$ed.displayName).ToLowerInvariant()] = [string]$ed.deviceId
            }
        }
        Write-Host "    Mapped $($context.EntraDeviceIdByName.Count) Entra device(s) by displayName." -ForegroundColor Green
    }
    catch { Write-Warning "/v1.0/devices fetch failed: $($_.Exception.Message)" }

    return $context
}
