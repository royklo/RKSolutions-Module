function ConvertTo-ComplianceSettingFriendlyName {
    param([string] $SettingName)
    if ($SettingName -match '^(Windows10CompliancePolicy|DefaultDeviceCompliancePolicy)\.(.+)$' -and $script:ComplianceSettingFriendlyNames.ContainsKey($Matches[2])) {
        return $script:ComplianceSettingFriendlyNames[$Matches[2]]
    }
    return $SettingName
}

function Resolve-DeviceAzureAdDeviceId {
    <#
        Return the Azure AD device id to use for joining with BitLocker recovery keys
        and LAPS credentials. Prefer the value already on the device record; fall back
        to the deviceName -> deviceId map populated by Get-BitLockerLapsAssignmentContext.
    #>
    param([Parameter(Mandatory)] [PSCustomObject] $Device, [Parameter(Mandatory)] [PSCustomObject] $Context)
    if ($Device.AzureAdDeviceId) { return [string]$Device.AzureAdDeviceId }
    if ($Device.DeviceName) {
        $key = ([string]$Device.DeviceName).ToLowerInvariant()
        if ($Context.EntraDeviceIdByName.ContainsKey($key)) { return $Context.EntraDeviceIdByName[$key] }
    }
    return $null
}

function ConvertTo-FilterDeviceProperties {
    <#
        Map a Get-AllDeviceData record onto the property names that Test-IntuneFilter
        recognises (deviceName, operatingSystem, osVersion, manufacturer, model,
        isEncrypted, ownerType, serialNumber, enrollmentProfileName, complianceState).
    #>
    param([Parameter(Mandatory)] [PSCustomObject] $Device)
    [PSCustomObject]@{
        DeviceName            = $Device.DeviceName
        OperatingSystem       = $Device.Operatingsystem
        OSVersion             = $Device.OperatingSystemVersion
        Manufacturer          = $Device.DeviceManufacturer
        Model                 = $Device.DeviceModel
        IsEncrypted           = $Device.Encrypted
        OwnerType             = $Device.DeviceOwnership
        SerialNumber          = $Device.Serialnumber
        EnrollmentProfileName = $Device.EnrollmentProfile
        ComplianceState       = $Device.ComplianceStatus
        ProcessorArchitecture = $Device.ProcessorArchitecture
        DeviceType            = 'desktop'
        UserPrincipalName     = $Device.PrimaryUser
        AzureAdDeviceId       = $Device.AzureAdDeviceId
    }
}

function Test-AssignmentAppliesToManagedDevice {
    <#
        Evaluates a single Get-DetailedPolicyAssignments row against a managed device
        record using the pre-resolved group membership sets in the context. Returns
        @{ Applies = $bool; FilterDecision = '...' }. Assignment filters are honored
        via Test-IntuneFilter (provided by Core/IntunePolicyAssignment.ps1).
    #>
    param(
        [Parameter(Mandatory)] [PSCustomObject] $Assignment,
        [Parameter(Mandatory)] [PSCustomObject] $Device,
        [Parameter(Mandatory)] [PSCustomObject] $Context
    )

    $aadDeviceId = if ($Device.AzureAdDeviceId) { [string]$Device.AzureAdDeviceId } else { '' }
    $primaryUpn  = if ($Device.PrimaryUser)     { ([string]$Device.PrimaryUser).ToLowerInvariant() } else { '' }
    $baseApplies = $false

    switch ($Assignment.AssignmentType) {
        'All Devices'      { $baseApplies = $true }
        'All Users'        { $baseApplies = -not [string]::IsNullOrEmpty($primaryUpn) }
        'Group (Include)'  {
            if ($Assignment.GroupId) {
                $gid = [string]$Assignment.GroupId
                $inDeviceGroup = $false; $inUserGroup = $false
                if ($Context.GroupDeviceAadIds.ContainsKey($gid) -and $aadDeviceId) {
                    $inDeviceGroup = $Context.GroupDeviceAadIds[$gid].Contains($aadDeviceId)
                }
                if ($Context.GroupUserUpns.ContainsKey($gid) -and $primaryUpn) {
                    $inUserGroup = $Context.GroupUserUpns[$gid].Contains($primaryUpn)
                }
                $baseApplies = $inDeviceGroup -or $inUserGroup
            }
        }
        'Group (Exclude)'  { $baseApplies = $false }    # Handled explicitly by caller as an exclusion signal.
        default            { $baseApplies = $false }
    }

    if (-not $baseApplies) { return [PSCustomObject]@{ Applies = $false; FilterDecision = 'N/A' } }

    # Assignment filter handling. Filter rules use property names like deviceName,
    # operatingSystem, manufacturer, model, osVersion - all surfaced by Test-IntuneFilter.
    $filterDecision = 'N/A'
    $hasFilter = $Assignment.FilterId -and $Assignment.FilterId -ne '00000000-0000-0000-0000-000000000000' `
                 -and $Assignment.FilterType -and $Assignment.FilterType -ne 'None' -and $Assignment.FilterType -ne 'none'
    if ($hasFilter) {
        $filter = $Context.Filters[$Assignment.FilterId]
        if ($filter -and $filter.rule) {
            $filterDeviceProps = ConvertTo-FilterDeviceProperties -Device $Device
            $matched = Test-IntuneFilter -FilterRule $filter.rule -DeviceProperties $filterDeviceProps
            $filterDecision = if ($matched) { 'Matched' } else { 'NotMatched' }
            $ft = ([string]$Assignment.FilterType).ToLowerInvariant()
            if     ($ft -eq 'include') { if (-not $matched) { return [PSCustomObject]@{ Applies = $false; FilterDecision = $filterDecision } } }
            elseif ($ft -eq 'exclude') { if ($matched)      { return [PSCustomObject]@{ Applies = $false; FilterDecision = $filterDecision } } }
        }
        else { $filterDecision = 'FilterNotFound' }
    }

    return [PSCustomObject]@{ Applies = $true; FilterDecision = $filterDecision }
}

function Get-DeviceAppliedPolicies {
    <#
        Walks a list of policies (each carrying its full assignment set) and decides,
        for the given device, whether each policy:
          - applies                       -> goes in .Applied
          - was explicitly excluded       -> goes in .Excluded with a reason
                                             ('ExcludeGroup' | 'AssignmentFilter')
          - simply doesn't target it      -> ignored (not anomalous in itself)

        A policy that targeted the device's group/user but was then rejected by an
        assignment filter, or by an exclusion group, is an *intentional* exclusion -
        callers can hide those from anomaly output by default and surface them under
        an opt-in flag.
    #>
    param(
        [Parameter(Mandatory)] [array]      $Policies,
        [Parameter(Mandatory)] [PSCustomObject] $Device,
        [Parameter(Mandatory)] [PSCustomObject] $Context
    )

    $applied  = [System.Collections.Generic.List[object]]::new()
    $excluded = [System.Collections.Generic.List[object]]::new()
    $aadDeviceId = if ($Device.AzureAdDeviceId) { [string]$Device.AzureAdDeviceId } else { '' }
    $primaryUpn  = if ($Device.PrimaryUser)     { ([string]$Device.PrimaryUser).ToLowerInvariant() } else { '' }

    foreach ($policy in $Policies) {
        # First pass: any exclude group that names this device's identity wins
        # outright, regardless of include scope.
        $excludeGroupHit = $false
        foreach ($a in $policy.Assignments) {
            if ($a.AssignmentType -ne 'Group (Exclude)' -or -not $a.GroupId) { continue }
            $gid = [string]$a.GroupId
            if ($aadDeviceId -and $Context.GroupDeviceAadIds.ContainsKey($gid) -and $Context.GroupDeviceAadIds[$gid].Contains($aadDeviceId)) { $excludeGroupHit = $true; break }
            if ($primaryUpn  -and $Context.GroupUserUpns.ContainsKey($gid)     -and $Context.GroupUserUpns[$gid].Contains($primaryUpn))      { $excludeGroupHit = $true; break }
        }
        if ($excludeGroupHit) {
            $excluded.Add([PSCustomObject]@{ PolicyName = $policy.DisplayName; Source = $policy.Source; Reason = 'ExcludeGroup' })
            continue
        }

        # Second pass: walk include-style assignments. Track whether any base-level
        # match happened (device was *targeted*) so we can distinguish a filter
        # rejection ("intentional exclusion") from a plain non-match ("never in scope").
        $includeMatch = $null
        $baseTargetedButFiltered = $false
        foreach ($a in $policy.Assignments) {
            if ($a.AssignmentType -eq 'Group (Exclude)' -or $a.AssignmentType -eq 'Not Assigned') { continue }
            $result = Test-AssignmentAppliesToManagedDevice -Assignment $a -Device $Device -Context $Context
            if ($result.Applies) { $includeMatch = $result; break }
            # Applies=$false with a non-N/A FilterDecision == base matched, filter rejected.
            if ($result.FilterDecision -in 'Matched','NotMatched','FilterNotFound') { $baseTargetedButFiltered = $true }
        }
        if ($includeMatch) {
            $applied.Add([PSCustomObject]@{ PolicyName = $policy.DisplayName; Source = $policy.Source; FilterDecision = $includeMatch.FilterDecision })
        }
        elseif ($baseTargetedButFiltered) {
            $excluded.Add([PSCustomObject]@{ PolicyName = $policy.DisplayName; Source = $policy.Source; Reason = 'AssignmentFilter' })
        }
        # else: policy didn't target this device at all. Not surfaced.
    }

    return [PSCustomObject]@{ Applied = @($applied); Excluded = @($excluded) }
}

function Resolve-IntuneBitLockerAnomalies {
    <#
        For each Windows managed device, decide whether BitLocker is governed by an
        assigned Intune policy and whether the OS-volume recovery key is escrowed to
        Entra. Emits one row per anomalous device with a severity bucket.

        A device that was deliberately excluded (via an exclude group or an assignment
        filter that rejected it) is *not* an anomaly by default - that's the admin's
        explicit intent. Pass -ShowExcludedDevices to surface those rows as Info.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [array]          $Devices,
        [Parameter(Mandatory)] [PSCustomObject] $Context,
        [Parameter(Mandatory)] [string]         $TenantName,
        [switch]                                $ShowExcludedDevices
    )
    $out = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($d in $Devices) {
        $aadId = Resolve-DeviceAzureAdDeviceId -Device $d -Context $Context
        $isEncrypted = [bool]$d.Encrypted

        # No Entra device id: we can't evaluate policy assignment or key escrow
        # for this device, but we still flag "not encrypted" since that's the
        # one anomaly we can detect from the managed-device record alone.
        if (-not $aadId) {
            if (-not $isEncrypted) {
                $out.Add([PSCustomObject][ordered]@{
                    Customer        = $TenantName
                    DeviceName      = $d.DeviceName
                    PrimaryUser     = $d.PrimaryUser
                    Serialnumber    = $d.Serialnumber
                    IsEncrypted     = 'No'
                    AppliedPolicies = ''
                    KeyEscrowed     = 'Unknown'
                    Severity        = 'Warning'
                    Status          = 'Device not encrypted (no Azure AD device id to verify policy / key state)'
                })
            }
            continue
        }

        $eval     = Get-DeviceAppliedPolicies -Policies $Context.BitLockerPolicies -Device $d -Context $Context
        $applied  = $eval.Applied
        $excluded = $eval.Excluded
        $hasKey   = $Context.BitLockerKeyDeviceIds.Contains($aadId)

        $status = $null; $severity = $null
        if ($applied.Count -eq 0 -and $excluded.Count -gt 0) {
            # Intentional exclusion. Skip unless caller asked to surface them.
            if (-not $ShowExcludedDevices) { continue }
            $reasons = ($excluded | Select-Object -ExpandProperty Reason -Unique) -join ', '
            $status = "Intentionally excluded from BitLocker policies ($reasons)"
            $severity = 'Info'
        }
        elseif (-not $isEncrypted -and $applied.Count -eq 0)  { $status = 'Device not encrypted and no BitLocker policy assigned';     $severity = 'Critical' }
        elseif ($applied.Count -eq 0)                         { $status = 'No BitLocker policy assigned';                              $severity = 'Warning' }
        elseif ($isEncrypted -and -not $hasKey)               { $status = 'Encrypted but no OS-volume key in Entra';                   $severity = 'Critical' }
        elseif (-not $isEncrypted)                            { $status = 'BitLocker policy assigned but device not encrypted';        $severity = 'Critical' }
        else                                                  { continue }   # Healthy.

        $out.Add([PSCustomObject][ordered]@{
            Customer           = $TenantName
            DeviceName         = $d.DeviceName
            PrimaryUser        = $d.PrimaryUser
            Serialnumber       = $d.Serialnumber
            IsEncrypted        = if ($isEncrypted) { 'Yes' } else { 'No' }
            AppliedPolicies    = if ($applied.Count -gt 0) { ($applied | ForEach-Object { $_.PolicyName }) -join '; ' } else { ($excluded | ForEach-Object { "$($_.PolicyName) [$($_.Reason)]" }) -join '; ' }
            KeyEscrowed        = if ($hasKey) { 'Yes' } else { 'No' }
            Severity           = $severity
            Status             = $status
        })
    }
    return $out
}

function Resolve-IntuneLapsAnomalies {
    <#
        For each Windows managed device, decide whether an Entra-backed Windows LAPS
        policy is assigned and whether a local admin credential is actually backed up.
        A backup older than -MaxBackupAgeDays is reported as a rotation gap.

        Devices that were explicitly excluded (exclude group or assignment filter) are
        suppressed by default; pass -ShowExcludedDevices to surface them as Info rows.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [array]          $Devices,
        [Parameter(Mandatory)] [PSCustomObject] $Context,
        [Parameter(Mandatory)] [string]         $TenantName,
        [int]    $MaxBackupAgeDays = 60,
        [switch] $ShowExcludedDevices
    )
    $out = [System.Collections.Generic.List[PSObject]]::new()
    $cutoff = (Get-Date).AddDays(-$MaxBackupAgeDays)

    foreach ($d in $Devices) {
        $aadId = Resolve-DeviceAzureAdDeviceId -Device $d -Context $Context
        if (-not $aadId) { continue }
        $eval     = Get-DeviceAppliedPolicies -Policies $Context.LapsPolicies -Device $d -Context $Context
        $applied  = $eval.Applied
        $excluded = $eval.Excluded

        # Primary join: LAPS credentials are keyed by the Azure AD device id.
        # Secondary join (deviceName) covers the rare case where the device id we
        # have differs from the one Entra stored the LAPS entry against.
        $entry = $Context.LapsCredentialByDeviceId[$aadId]
        if (-not $entry -and $d.DeviceName) {
            $nameKey = ([string]$d.DeviceName).ToLowerInvariant()
            if ($Context.LapsCredentialByDeviceName.ContainsKey($nameKey)) {
                $entry = $Context.LapsCredentialByDeviceName[$nameKey]
            }
        }

        $status = $null; $severity = $null
        $lastBackup = $null; $ageDays = $null
        if ($entry -and $entry.lastBackupDateTime) {
            try { $lastBackup = [datetime]$entry.lastBackupDateTime; $ageDays = [int](((Get-Date) - $lastBackup).TotalDays) } catch { }
        }

        if ($applied.Count -eq 0 -and $excluded.Count -gt 0) {
            if (-not $ShowExcludedDevices) { continue }
            $reasons = ($excluded | Select-Object -ExpandProperty Reason -Unique) -join ', '
            $status = "Intentionally excluded from LAPS policies ($reasons)"
            $severity = 'Info'
        }
        elseif ($applied.Count -eq 0)                                    { $status = 'No Entra-backed LAPS policy assigned';                          $severity = 'Warning' }
        elseif (-not $entry)                                             { $status = 'LAPS policy assigned but no credential backed up to Entra';     $severity = 'Critical' }
        elseif ($lastBackup -and $lastBackup -lt $cutoff)                { $status = "LAPS backup stale (> $MaxBackupAgeDays days, rotation may be stalled)"; $severity = 'Warning' }
        else { continue }   # Healthy.

        $out.Add([PSCustomObject][ordered]@{
            Customer           = $TenantName
            DeviceName         = $d.DeviceName
            PrimaryUser        = $d.PrimaryUser
            Serialnumber       = $d.Serialnumber
            OwnerType          = $d.DeviceOwnership
            AppliedPolicies    = if ($applied.Count -gt 0) { ($applied | ForEach-Object { $_.PolicyName }) -join '; ' } else { ($excluded | ForEach-Object { "$($_.PolicyName) [$($_.Reason)]" }) -join '; ' }
            LastBackupDateTime = if ($lastBackup) { $lastBackup.ToString('yyyy-MM-dd HH:mm') } else { '' }
            BackupAgeDays      = if ($null -ne $ageDays) { $ageDays } else { '' }
            Severity           = $severity
            Status             = $status
        })
    }
    return $out
}

# ----------------------------------------------------------------------------
# Deprecated Intune settings discovery.
# Walks every Settings Catalog policy already fetched into the context by
# Get-BitLockerLapsAssignmentContext and flags individual settings that
# Microsoft has marked as deprecated. Detection sources, in order of precedence:
#   1. Catalog DisplayName for the settingDefinitionId contains "deprecated"
#   2. The setting's raw definition id contains "deprecated"
#   3. The configured value string contains "deprecated"
# Catalog data comes from Get-RKIntuneSettingsCatalog (cached download from
# github.com/royklo/IntuneSettingsCatalogData); when the catalog is empty the
# walker still detects the subset that carries "deprecated" in the raw id or
# value string.
# ----------------------------------------------------------------------------

function Test-IsSettingDeprecated {
    <#
        Returns @{ IsDeprecated; Source; DisplayName } where Source is one of
        'CatalogDisplayName' | 'DefinitionIdMatch' | 'ValueMatch' on a hit.
    #>
    [CmdletBinding()]
    param(
        [string]   $DefinitionId,
        [string]   $Value,
        [hashtable]$Catalog
    )

    if ($Catalog -and $Catalog.Count -gt 0 -and -not [string]::IsNullOrEmpty($DefinitionId) -and $Catalog.ContainsKey($DefinitionId)) {
        $dn = [string]$Catalog[$DefinitionId].DisplayName
        if ($dn -match '(?i)deprecated') {
            return [PSCustomObject]@{ IsDeprecated = $true; Source = 'CatalogDisplayName'; DisplayName = $dn }
        }
    }
    if (-not [string]::IsNullOrEmpty($DefinitionId) -and $DefinitionId -match '(?i)deprecated') {
        $dn = if ($Catalog -and $Catalog.ContainsKey($DefinitionId)) { [string]$Catalog[$DefinitionId].DisplayName } else { $DefinitionId }
        return [PSCustomObject]@{ IsDeprecated = $true; Source = 'DefinitionIdMatch'; DisplayName = $dn }
    }
    if (-not [string]::IsNullOrEmpty($Value) -and $Value -match '(?i)deprecated') {
        $dn = if ($Catalog -and $Catalog.ContainsKey($DefinitionId)) { [string]$Catalog[$DefinitionId].DisplayName } else { $DefinitionId }
        return [PSCustomObject]@{ IsDeprecated = $true; Source = 'ValueMatch'; DisplayName = $dn }
    }
    return [PSCustomObject]@{ IsDeprecated = $false }
}

function Get-IntuneSettingValueString {
    <#
        Extract the configured value(s) from a settingInstance, regardless of OData
        sub-type (choice / simple / collection). Returns a string suitable for
        display and for the deprecation regex check. Group / collection instances
        return empty here; their children are walked separately by the caller.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)] $SettingInstance)

    $type = [string]$SettingInstance.'@odata.type'
    switch -Wildcard ($type) {
        '*choiceSettingInstance*'           { return [string]$SettingInstance.choiceSettingValue.value }
        '*simpleSettingInstance*'           { return [string]$SettingInstance.simpleSettingValue.value }
        '*choiceSettingCollectionInstance*' {
            $vals = @()
            foreach ($cv in @($SettingInstance.choiceSettingCollectionValue)) { if ($cv.value) { $vals += [string]$cv.value } }
            return ($vals -join '; ')
        }
        '*simpleSettingCollectionInstance*' {
            $vals = @()
            foreach ($sv in @($SettingInstance.simpleSettingCollectionValue)) { if ($sv.value) { $vals += [string]$sv.value } }
            return ($vals -join '; ')
        }
        default                              { return '' }
    }
}

function Get-IntuneSettingDeprecations {
    <#
        Recursively walk a settingInstance (and any nested group children) and
        return one record per deprecated setting found:
            { DefinitionId; DisplayName; Value; Source }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $SettingInstance,
        [hashtable]            $Catalog
    )

    $hits = [System.Collections.Generic.List[object]]::new()
    if (-not $SettingInstance) { return $hits }

    $defId = [string]$SettingInstance.settingDefinitionId
    $value = Get-IntuneSettingValueString -SettingInstance $SettingInstance
    $check = Test-IsSettingDeprecated -DefinitionId $defId -Value $value -Catalog $Catalog
    if ($check.IsDeprecated) {
        $hits.Add([PSCustomObject]@{
            DefinitionId = $defId
            DisplayName  = if ([string]::IsNullOrWhiteSpace($check.DisplayName)) { $defId } else { $check.DisplayName }
            Value        = $value
            Source       = $check.Source
        })
    }

    $type = [string]$SettingInstance.'@odata.type'
    if ($type -like '*groupSettingCollectionInstance*') {
        foreach ($child in @($SettingInstance.groupSettingCollectionValue)) {
            foreach ($childSetting in @($child.children)) {
                foreach ($h in (Get-IntuneSettingDeprecations -SettingInstance $childSetting -Catalog $Catalog)) { $hits.Add($h) }
            }
        }
    }
    elseif ($type -like '*groupSettingInstance*') {
        foreach ($childSetting in @($SettingInstance.groupSettingValue.children)) {
            foreach ($h in (Get-IntuneSettingDeprecations -SettingInstance $childSetting -Catalog $Catalog)) { $hits.Add($h) }
        }
    }
    return $hits
}

function Resolve-IntuneDeprecatedAnomalies {
    <#
        Iterates the Settings Catalog policies already on the context, scans
        every setting against the cached settings-catalog lookup, and emits one
        row per deprecated setting in use. Loads / refreshes the catalog on
        first call; subsequent calls in the same session reuse it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [PSCustomObject] $Context,
        [Parameter(Mandatory)] [string]         $TenantName
    )

    Write-Host '  Scanning Settings Catalog policies for deprecated settings...' -ForegroundColor Cyan
    $catalog = Get-RKIntuneSettingsCatalog
    $out = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($policy in $Context.ConfigurationPolicies) {
        if (-not $policy.settings) { continue }
        $policyName   = if ($policy.name) { $policy.name } elseif ($policy.displayName) { $policy.displayName } else { '(unnamed)' }
        $platform     = if ($policy.platforms) { [string]$policy.platforms } else { '' }
        $technologies = if ($policy.technologies) { [string]$policy.technologies } else { '' }

        foreach ($s in @($policy.settings)) {
            if (-not $s.settingInstance) { continue }
            foreach ($hit in (Get-IntuneSettingDeprecations -SettingInstance $s.settingInstance -Catalog $catalog)) {
                $out.Add([PSCustomObject]@{
                    Customer            = $TenantName
                    PolicyName          = $policyName
                    PolicyId            = $policy.id
                    Platform            = $platform
                    Technologies        = $technologies
                    SettingDisplayName  = $hit.DisplayName
                    SettingDefinitionId = $hit.DefinitionId
                    ConfiguredValue     = $hit.Value
                    DetectionSource     = $hit.Source
                })
            }
        }
    }

    Write-Host "    Found $($out.Count) deprecated setting instance(s) across $($Context.ConfigurationPolicies.Count) Settings Catalog policies." -ForegroundColor Gray
    return $out
}
