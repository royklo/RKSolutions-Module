function Get-GroupInfo {
    param([Parameter(Mandatory)][string]$GroupId)
    if ($script:AllGroups.Count -gt 0) {
        $group = $script:AllGroups[$GroupId]
        if ($group) { return @{ Id = $group.id; DisplayName = $group.displayName; Success = $true } }
    }
    try {
        $group = Invoke-RKCoreGraphRequest -Uri "v1.0/groups/$GroupId"
        return @{ Id = $group.id; DisplayName = $group.displayName; Success = $true }
    }
    catch {
        return @{ Id = $GroupId; DisplayName = "Unknown Group"; Success = $false }
    }
}

function Get-DetailedPolicyAssignments {
    param([Parameter(Mandatory)][string]$EntityType, [string]$EntityId, [string]$PolicyName, [switch]$DebugMode, [switch]$ThrowOnError)
    
    $assignmentsUri = $null
    if ($EntityType -eq "deviceAppManagement/managedAppPolicies") {
        try {
            $r = Invoke-RKCoreGraphRequest -Uri "beta/deviceAppManagement/managedAppPolicies/$EntityId"
            $path = switch ($r.'@odata.type') {
                "#microsoft.graph.androidManagedAppProtection" { "androidManagedAppProtections" }
                "#microsoft.graph.iosManagedAppProtection" { "iosManagedAppProtections" }
                "#microsoft.graph.windowsManagedAppProtection" { "windowsManagedAppProtections" }
                default { $null }
            }
            if ($path) { $assignmentsUri = "beta/deviceAppManagement/$path('$EntityId')/assignments" }
        }
        catch { return @() }
    }
    elseif ($EntityType -eq "mobileAppConfigurations") {
        $assignmentsUri = "beta/deviceAppManagement/mobileAppConfigurations('$EntityId')/assignments"
    }
    elseif ($EntityType -eq "deviceAppManagement/mobileApps") {
        $assignmentsUri = "beta/deviceAppManagement/mobileApps('$EntityId')/assignments"
    }
    elseif ($EntityType -like "deviceAppManagement/*ManagedAppProtections") {
        $assignmentsUri = "beta/$EntityType('$EntityId')/assignments"
    }
    elseif ($EntityType -eq "deviceManagement/intents") {
        $assignmentsUri = "beta/deviceManagement/intents/$EntityId/assignments"
    }
    elseif ($EntityType -eq "deviceManagement/templates") {
        $assignmentsUri = "beta/deviceManagement/templates/$EntityId/assignments"
    }
    elseif ($EntityType -like "virtualEndpoint/*") {
        $assignmentsUri = "beta/deviceManagement/$EntityType/$EntityId/assignments"
    }
    elseif ($EntityType -eq "configurationPolicies") {
        $assignmentsUri = "beta/deviceManagement/configurationPolicies/$EntityId/assignments"
    }
    else {
        $assignmentsUri = "beta/deviceManagement/$EntityType('$EntityId')/assignments"
    }
    
    if (-not $assignmentsUri) { return @() }
    
    try {
        Write-Verbose "        Querying assignments for: $PolicyName ($EntityType, $EntityId)"
        Write-Verbose "        Assignments URI: $assignmentsUri"
        $assignments = @(Invoke-RKCoreGraphRequest -Uri $assignmentsUri -All)
        Write-Verbose "        Assignments result: $($assignments.Count) assignments returned"
        
        $detailedAssignments = @()
        foreach ($assignment in $assignments) {
            $targetType = $null
            $targetName = $null
            $targetId = $null
            $groupId = $null
            
            if ($assignment.target) {
                $targetType = $assignment.target.'@odata.type'
                $groupId = $assignment.target.groupId
                $targetId = $assignment.target.deviceAndAppManagementAssignmentFilterId
                $targetFilterType = $assignment.target.deviceAndAppManagementAssignmentFilterType
                
                if ($targetType -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
                    $targetName = "All Users"
                }
                elseif ($targetType -eq '#microsoft.graph.allDevicesAssignmentTarget') {
                    $targetName = "All Devices"
                }
                elseif ($targetType -eq '#microsoft.graph.groupAssignmentTarget' -and $groupId) {
                    $groupInfo = Get-GroupInfo -GroupId $groupId
                    $targetName = $groupInfo.DisplayName
                }
                elseif ($targetType -eq '#microsoft.graph.exclusionGroupAssignmentTarget' -and $groupId) {
                    $groupInfo = Get-GroupInfo -GroupId $groupId
                    $targetName = $groupInfo.DisplayName
                }
            }
            
            $assignmentIntent = "Apply"
            if ($assignment.intent) {
                switch ($assignment.intent) {
                    "required" { $assignmentIntent = "Required" }
                    "available" { $assignmentIntent = "Available" }
                    "uninstall" { $assignmentIntent = "Uninstall" }
                    "availableWithoutEnrollment" { $assignmentIntent = "Available (No Enrollment)" }
                    default { $assignmentIntent = $assignment.intent }
                }
            }
            
            # Filter enrichment
            $filterId = $null
            $filterName = "No Filter"
            $filterType = "None"
            $filterRule = $null
            $filterPlatform = $null
            
            if ($targetId) {
                $filter = $script:AllFilters[$targetId]
                if ($filter) {
                    $filterId = $filter.id
                    $filterName = $filter.displayName
                    # Use filterType from assignment target, not from filter object
                    if ($targetFilterType) {
                        $filterType = $targetFilterType
                    } else {
                        $filterType = $filter.assignmentFilterManagementType
                    }
                    $filterRule = $filter.rule
                    $filterPlatform = $filter.platform
                }
            }
            
            # Convert raw OData type to friendly AssignmentType for categorization
            $friendlyAssignmentType = if ($script:ODataToFriendly.ContainsKey($targetType)) { $script:ODataToFriendly[$targetType] } else { $targetType }
            
            $detailedAssignments += [PSCustomObject]@{
                PolicyName = $PolicyName
                PolicyId = $EntityId
                PolicyType = $EntityType
                AssignmentType = $friendlyAssignmentType
                AssignmentIntent = $assignmentIntent
                TargetName = $targetName
                TargetId = $targetId
                GroupId = $groupId
                FilterId = $filterId
                FilterName = $filterName
                FilterType = $filterType
                FilterRule = $filterRule
                FilterPlatform = $filterPlatform
                AssignmentId = $assignment.id
            }
        }
        return $detailedAssignments
    }
    catch {
        # The anomaly core passes -ThrowOnError: an unread assignment would make devices look unprotected.
        # A 404 means the policy was deleted while we read it, which is the same as no assignments.
        if ($ThrowOnError -and (Get-RKCoreGraphErrorStatus -ErrorRecord $_) -ne 404) { throw }
        return @()
    }
}

function Test-IntuneFilter {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$FilterRule, [Parameter(Mandatory)][PSCustomObject]$DeviceProperties)
    $propertyMapping = @{
        'cpuArchitecture' = 'ProcessorArchitecture'; 'deviceCategory' = 'DeviceCategoryDisplayName'; 'deviceName' = 'DeviceName'; 'deviceOwnership' = 'OwnerType'; 'deviceTrustType' = 'JoinType'
        'enrollmentProfileName' = 'EnrollmentProfileName'; 'isRooted' = 'JailBroken'; 'manufacturer' = 'Manufacturer'; 'model' = 'Model'; 'operatingSystemVersion' = 'OSVersion'; 'osVersion' = 'OSVersion'
        'deviceType' = 'DeviceType'; 'operatingSystem' = 'OperatingSystem'; 'complianceState' = 'ComplianceState'; 'managementAgent' = 'ManagementAgent'; 'autopilotEnrolled' = 'AutopilotEnrolled'
        'deviceEnrollmentType' = 'DeviceEnrollmentType'; 'chassisType' = 'ChassisType'; 'managementState' = 'ManagementState'; 'deviceRegistrationState' = 'DeviceRegistrationState'
        'azureAdRegistered' = 'AzureAdRegistered'; 'isEncrypted' = 'IsEncrypted'; 'isSupervised' = 'IsSupervised'; 'jailBroken' = 'JailBroken'; 'serialNumber' = 'SerialNumber'; 'processorArchitecture' = 'ProcessorArchitecture'
    }
    function Test-Condition {
        param([string]$PropertyPath, [string]$Operator, [string]$Value)
        $property = $PropertyPath.Trim(); if ($property.StartsWith("device.")) { $property = $property.Substring(7) }
        $actualPropertyName = if ($propertyMapping.ContainsKey($property.ToLower())) { $propertyMapping[$property.ToLower()] } else { $property }
        $actualValue = $null; if ($DeviceProperties.PSObject.Properties.Name -contains $actualPropertyName) { $actualValue = $DeviceProperties.$actualPropertyName }
        if ($null -eq $actualValue) { return $false }
        $actualValueStr = $actualValue.ToString(); $targetValueStr = $Value.ToString()
        switch -Regex ($Operator.Trim()) {
            "^-?eq$" { return $actualValueStr -eq $targetValueStr }
            "^-?ne$" { return $actualValueStr -ne $targetValueStr }
            "^-?contains$" { return $actualValueStr.Contains($targetValueStr) }
            "^-?notContains$" { return -not $actualValueStr.Contains($targetValueStr) }
            "^-?startsWith$" { return $actualValueStr.StartsWith($targetValueStr) }
            "^-?notStartsWith$" { return -not $actualValueStr.StartsWith($targetValueStr) }
            "^-?endsWith$" { return $actualValueStr.EndsWith($targetValueStr) }
            "^-?notEndsWith$" { return -not $actualValueStr.EndsWith($targetValueStr) }
            "^-?match$" { return $actualValueStr -match $targetValueStr }
            "^-?notMatch$" { return $actualValueStr -notmatch $targetValueStr }
            "^-?in$" { $cleanValue = $targetValueStr -replace '^\[|\]$', ''; $valueArray = $cleanValue -split ',' | ForEach-Object { $_.Trim().Trim('"').Trim("'") }; return $valueArray -contains $actualValueStr }
            "^-?notIn$" { $cleanValue = $targetValueStr -replace '^\[|\]$', ''; $valueArray = $cleanValue -split ',' | ForEach-Object { $_.Trim().Trim('"').Trim("'") }; return $valueArray -notcontains $actualValueStr }
            "^-?gt$" { try { return [System.Version]$actualValueStr -gt [System.Version]$targetValueStr } catch { return $actualValueStr -gt $targetValueStr } }
            "^-?ge$" { try { return [System.Version]$actualValueStr -ge [System.Version]$targetValueStr } catch { return $actualValueStr -ge $targetValueStr } }
            "^-?lt$" { try { return [System.Version]$actualValueStr -lt [System.Version]$targetValueStr } catch { return $actualValueStr -lt $targetValueStr } }
            "^-?le$" { try { return [System.Version]$actualValueStr -le [System.Version]$targetValueStr } catch { return $actualValueStr -le $targetValueStr } }
            default { return $false }
        }
    }
    $filterRule = $FilterRule -replace '\s+', ' '
    $orConditions = @(); $sections = $filterRule -split '\s+or\s+', 0, 'IgnoreCase'
    foreach ($section in $sections) {
        $andConditions = @(); $andSections = $section -split '\s+and\s+', 0, 'IgnoreCase'
        foreach ($andSection in $andSections) {
            $cleanSection = $andSection.Trim()
            if ($cleanSection.StartsWith("(") -and $cleanSection.EndsWith(")")) { $cleanSection = $cleanSection.Substring(1, $cleanSection.Length - 2).Trim() }
            if ($cleanSection -match '^(.+?)\s+(-?\w+)\s+(.+)$') {
                $propertyPath = $matches[1].Trim(); if ($propertyPath.StartsWith("device.")) { $propertyPath = $propertyPath.Substring(7) }
                $operator = $matches[2].Trim(); $value = $matches[3].Trim()
                $cleanValue = $value
                if (-not ($value.StartsWith('[') -and $value.EndsWith(']'))) {
                    $cleanValue = $value.Trim().Trim('"').Trim("'")
                }
                $andConditions += @{ PropertyPath = $propertyPath; Operator = $operator; Value = $cleanValue }
            }
        }
        $orConditions += @{ AndConditions = $andConditions }
    }
    foreach ($orCondition in $orConditions) {
        $allTrue = $true
        foreach ($andCondition in $orCondition.AndConditions) {
            if (-not (Test-Condition -PropertyPath $andCondition.PropertyPath -Operator $andCondition.Operator -Value $andCondition.Value)) { $allTrue = $false; break }
        }
        if ($allTrue) { return $true }
    }
    return $false
}
