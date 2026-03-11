# Intune Enrollment Flows - Private helpers

$ErrorActionPreference = 'Stop'
$script:AllFilters = @{}
$script:AllGroups = @{}
$script:GroupMemberTypeCache = $null

# Shared helper: convert boolean/null to Yes/No/dash
$script:ToYesNo = { param($v) if ($null -eq $v) { "-" } elseif ($v -eq $true) { "Yes" } elseif ($v -eq $false) { "No" } else { $v.ToString() } }

# Shared lookup tables for OData assignment type <-> friendly name
$script:ODataToFriendly = @{
    '#microsoft.graph.allDevicesAssignmentTarget'         = 'All Devices'
    '#microsoft.graph.allLicensedUsersAssignmentTarget'   = 'All Users'
    '#microsoft.graph.groupAssignmentTarget'              = 'Group (Include)'
    '#microsoft.graph.exclusionGroupAssignmentTarget'     = 'Group (Exclude)'
}

# Single source of truth for Graph device $select (Section 8.1)
$script:ManagedDeviceSelectProperties = @(
    'AadRegistered', 'hardwareInformation', 'ActivationLockBypassCode', 'AndroidSecurityPatchLevel', 'AssignmentFilterEvaluationStatusDetails',
    'AutopilotEnrolled', 'AzureActiveDirectoryDeviceId', 'AzureAdDeviceId', 'AzureAdRegistered', 'BootstrapTokenEscrowed', 'ChassisType',
    'ChromeOSDeviceInfo', 'ComplianceGracePeriodExpirationDateTime', 'ComplianceState', 'ConfigurationManagerClientEnabledFeatures',
    'ConfigurationManagerClientHealthState', 'ConfigurationManagerClientInformation', 'DetectedApps', 'DeviceActionResults', 'DeviceCategory',
    'DeviceCategoryDisplayName', 'DeviceCompliancePolicyStates', 'DeviceConfigurationStates', 'DeviceEnrollmentType', 'DeviceFirmwareConfigurationInterfaceManaged',
    'DeviceHealthAttestationState', 'DeviceName', 'DeviceRegistrationState', 'DeviceType', 'EasActivated', 'EasActivationDateTime', 'EasDeviceId',
    'EmailAddress', 'EnrolledDateTime', 'EnrollmentProfileName', 'EthernetMacAddress', 'ExchangeAccessState', 'ExchangeAccessStateReason',
    'ExchangeLastSuccessfulSyncDateTime', 'FreeStorageSpaceInBytes', 'Iccid', 'Id', 'Imei', 'IsEncrypted', 'IsSupervised', 'JailBroken', 'JoinType',
    'LastSyncDateTime', 'LogCollectionRequests', 'LostModeState', 'ManagedDeviceMobileAppConfigurationStates', 'ManagedDeviceName', 'ManagedDeviceOwnerType',
    'ManagementAgent', 'ManagementCertificateExpirationDate', 'ManagementFeatures', 'ManagementState', 'Manufacturer', 'Meid', 'Model', 'Notes',
    'OSVersion', 'OperatingSystem', 'OwnerType', 'PartnerReportedThreatState', 'PhoneNumber', 'PhysicalMemoryInBytes', 'PreferMdmOverGroupPolicyAppliedDateTime',
    'ProcessorArchitecture', 'RemoteAssistanceSessionErrorDetails', 'RemoteAssistanceSessionUrl', 'RequireUserEnrollmentApproval', 'RetireAfterDateTime',
    'RoleScopeTagIds', 'SecurityBaselineStates', 'SerialNumber', 'SkuFamily', 'SkuNumber', 'SpecificationVersion', 'SubscriberCarrier',
    'TotalStorageSpaceInBytes', 'Udid', 'UserDisplayName', 'UserId', 'UserPrincipalName', 'Users', 'UsersLoggedOn', 'WiFiMacAddress',
    'WindowsActiveMalwareCount', 'WindowsProtectionState', 'WindowsRemediatedMalwareCount'
)

function Get-GroupMemberTargetType {
    param([Parameter(Mandatory)][string]$GroupId)
    if (-not $script:GroupMemberTypeCache) { $script:GroupMemberTypeCache = @{} }
    if ($script:GroupMemberTypeCache.ContainsKey($GroupId)) { return $script:GroupMemberTypeCache[$GroupId] }
    $result = 'Unknown'
    # 1. For dynamic groups, analyze the membershipRule
    if ($script:AllGroups.Count -gt 0) {
        $group = $script:AllGroups[$GroupId]
        if ($group) {
            $isDynamic = $group.groupTypes -and (@($group.groupTypes) -contains 'DynamicMembership')
            if ($isDynamic -and $group.membershipRule) {
                $rule = $group.membershipRule.Trim()
                if ($rule -match '\bdevice\.') { $result = 'Device' }
                elseif ($rule -match '\buser\.') { $result = 'User' }
            }
        }
    }
    # 2. For static groups or unparsable rules, check the first member's type
    if ($result -eq 'Unknown') {
        try {
            $membersResp = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members?`$top=1&`$select=id" -Method GET -OutputType PSObject -ErrorAction Stop
            if ($membersResp.value -and $membersResp.value.Count -gt 0) {
                $odataType = $membersResp.value[0].'@odata.type'
                if ($odataType -eq '#microsoft.graph.device') { $result = 'Device' }
                elseif ($odataType -eq '#microsoft.graph.user') { $result = 'User' }
            }
        }
        catch { }
    }
    $script:GroupMemberTypeCache[$GroupId] = $result
    return $result
}

function Get-GroupMemberTargetTypeFromRow {
    param([Parameter(Mandatory)][string]$AssignmentType, [string]$GroupId, [array]$DeviceGroupIds, [array]$UserGroupIds)
    if ($AssignmentType -eq 'All Devices') { return 'Device' }
    if ($AssignmentType -eq 'All Users') { return 'User' }
    if ($AssignmentType -match 'Group' -and $GroupId) {
        $memberType = Get-GroupMemberTargetType -GroupId $GroupId
        if ($memberType -ne 'Unknown') { return $memberType }
        # Fallback: check which membership list contains this group
        $gidStr = [string]$GroupId
        $inDevice = @($DeviceGroupIds | Where-Object { [string]$_ -eq $gidStr }).Count -gt 0
        $inUser = @($UserGroupIds | Where-Object { [string]$_ -eq $gidStr }).Count -gt 0
        if ($inDevice -and -not $inUser) { return 'Device' }
        if ($inUser -and -not $inDevice) { return 'User' }
        if ($inDevice) { return 'Device' }
    }
    return 'Unknown'
}

function Get-IntuneEntities {
    param([Parameter(Mandatory)][string]$EntityType, [string]$Filter = "", [string]$Select = "", [string]$Expand = "", [switch]$DebugMode)
    if ($EntityType -like "deviceAppManagement/*" -or $EntityType -eq "deviceManagement/templates" -or $EntityType -eq "deviceManagement/intents") {
        $baseUri = "https://graph.microsoft.com/beta"; $actualEntityType = $EntityType
    }
    else {
        $baseUri = "https://graph.microsoft.com/beta/deviceManagement"; $actualEntityType = $EntityType
    }
    $currentUri = "$baseUri/$actualEntityType"
    if ($Filter) { $currentUri += "?`$filter=$Filter" }
    if ($Select) { $currentUri += $(if ($Filter) { "&" } else { "?" }) + "`$select=$Select" }
    if ($Expand) { $currentUri += $(if ($Filter -or $Select) { "&" } else { "?" }) + "`$expand=$Expand" }
    $entities = Invoke-GraphRequestWithPaging -Uri $currentUri -Method "GET" -DebugMode:$DebugMode
    if ($entities) { return $entities } else { return @() }
}

function Get-AutopilotProfileConfigByDisplayName {
    param([Parameter(Mandatory)][string]$DisplayName)
    try {
        $profiles = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles" -Method GET
        $autopilotProfile = $profiles | Where-Object { $_.displayName -eq $DisplayName } | Select-Object -First 1
        if (-not $autopilotProfile) { return $null }
        return @{
            DeviceNameTemplate = if ($autopilotProfile.deviceNameTemplate) { $autopilotProfile.deviceNameTemplate } else { "-" }
            Language           = if ($autopilotProfile.language) { $autopilotProfile.language } else { "-" }
            Locale             = if ($autopilotProfile.locale) { $autopilotProfile.locale } else { "-" }
        }
    }
    catch { return $null }
}

function Get-MobileAppDisplayName {
    param([Parameter(Mandatory)][string]$AppId)
    try {
        $app = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps('$AppId')?`$select=displayName" -Method GET -OutputType PSObject -ErrorAction Stop
        if ($app -and $app.displayName) { return $app.displayName }
    }
    catch {
        Write-Verbose "Get-MobileAppDisplayName: Could not resolve app '$AppId'; using ID as display name."
    }
    return $AppId
}

function Get-EspConfigByDisplayName {
    param([Parameter(Mandatory)][string]$DisplayName)
    try {
        $configs = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations" -Method GET
        $displayNameNorm = ($DisplayName -replace '\s+', ' ').Trim()
        $espMatch = $configs | Where-Object {
            $_.'@odata.type' -eq '#microsoft.graph.windows10EnrollmentCompletionPageConfiguration' -and
            $_.displayName -and (($_.displayName -replace '\s+', ' ').Trim() -eq $displayNameNorm)
        } | Select-Object -First 1
        if (-not $espMatch) { return $null }
        $espId = $espMatch.id
        $getUri = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations('$espId')"
        $rawEsp = Invoke-MgGraphRequest -Uri $getUri -Method GET -OutputType PSObject -ErrorAction Stop
        $esp = if ($rawEsp.PSObject.Properties['value']) { $rawEsp.value } else { $rawEsp }
        if (-not $esp) { return $null }
        $toYesNo = $script:ToYesNo
        $desc = if ($esp.description) { $esp.description } else { "-" }
        $timeoutVal = $esp.installProgressTimeoutInMinutes; $timeoutMinutes = if ($null -ne $timeoutVal) { [int]$timeoutVal } else { $null }; $timeoutDisplay = if ($null -ne $timeoutMinutes) { "$timeoutMinutes" } else { "-" }
        $showProgressVal = $esp.showInstallationProgress; $showProgress = & $toYesNo $showProgressVal
        $trackVal = $esp.trackInstallProgressForAutopilotOnly; $trackAutopilotOnly = & $toYesNo $trackVal
        $qualityVal = $esp.installQualityUpdates; $installQuality = & $toYesNo $qualityVal
        $allowResetVal = $esp.allowDeviceResetOnInstallFailure; $allowReset = & $toYesNo $allowResetVal
        $allowUseVal = $esp.allowDeviceUseOnInstallFailure; $allowUse = & $toYesNo $allowUseVal
        $allowLogVal = $esp.allowLogCollectionOnInstallFailure; $allowLog = & $toYesNo $allowLogVal
        $allowNonBlockVal = $esp.allowNonBlockingAppInstallation; $allowNonBlocking = & $toYesNo $allowNonBlockVal
        $customErr = if ($esp.customErrorMessage) { $esp.customErrorMessage.Trim() } else { "" }; $customErrorMessage = if ($customErr) { $customErr } else { "-" }
        $showCustomMessage = if ($customErr) { "Yes" } else { "No" }
        $selectedIds = [System.Collections.ArrayList]::new()
        $rawSelected = if ($esp.PSObject.Properties['selectedMobileAppIds']) { $esp.selectedMobileAppIds } elseif ($esp.PSObject.Properties['SelectedMobileAppIds']) { $esp.SelectedMobileAppIds } else { $null }
        if ($rawSelected) {
            foreach ($item in @($rawSelected)) {
                $idStr = [string]$item
                if ($idStr -and $idStr.Trim().Length -gt 0) { [void]$selectedIds.Add($idStr.Trim()) }
            }
        }
        $selectedIds = @($selectedIds)
        return @{
            Description                     = $desc
            InstallProgressTimeout          = $timeoutDisplay
            ShowInstallationProgress        = $showProgress
            ShowCustomMessageWhenError      = $showCustomMessage
            CustomErrorMessage              = $customErrorMessage
            TrackAutopilotOnly              = $trackAutopilotOnly
            InstallQualityUpdates           = $installQuality
            AllowDeviceResetOnFailure       = $allowReset
            AllowDeviceUseOnFailure         = $allowUse
            AllowLogCollectionOnFailure     = $allowLog
            AllowNonBlockingAppInstallation = $allowNonBlocking
            SelectedMobileAppIds            = $selectedIds
        }
    }
    catch { return $null }
}

function Get-GroupParentGroupNames {
    param([Parameter(Mandatory)][string]$GroupId)
    try {
        $memberOf = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/memberOf?`$select=id" -Method GET
        if (-not $memberOf -or $memberOf.Count -eq 0) { return @() }
        $parentNames = @()
        foreach ($parent in $memberOf) {
            $parentId = $parent.id
            if ($script:AllGroups.Count -gt 0) {
                $p = $script:AllGroups[$parentId]
                if ($p -and $p.displayName) { $parentNames += $p.displayName }
            }
        }
        return $parentNames
    }
    catch { return @() }
}

function Get-GroupDirectMemberGroupIds {
    param([Parameter(Mandatory)][string]$GroupId)
    try {
        $uri = "https://graph.microsoft.com/beta/groups/$GroupId/members"
        $members = Invoke-GraphRequestWithPaging -Uri $uri -Method GET
        if (-not $members) { return @() }
        $membersArray = @($members)
        if ($membersArray.Count -eq 0) { return @() }
        $groupIds = @()
        foreach ($m in $membersArray) {
            $mid = $m.id
            if (-not $mid) { continue }
            $otype = $null
            if ($m.PSObject.Properties['@odata.type']) { $otype = $m.'@odata.type' }
            if (-not $otype -and $m.PSObject.Properties['odata.type']) { $otype = $m.'odata.type' }
            $isGroup = ($otype -eq '#microsoft.graph.group') -or ($script:AllGroups.Count -gt 0 -and $script:AllGroups.ContainsKey($mid))
            if ($isGroup) { $groupIds += $mid }
        }
        return $groupIds
    }
    catch { return @() }
}

function Get-NestedGroupChainNames {
    param([Parameter(Mandatory)][string]$GroupId, [Parameter(Mandatory)][array]$DeviceGroupIdsStr)
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $levels = [System.Collections.ArrayList]::new()
    $current = Get-GroupDirectMemberGroupIds -GroupId $GroupId | Where-Object { $DeviceGroupIdsStr -contains [string]$_ }
    while ($current -and $current.Count -gt 0) {
        $levelIds = @()
        foreach ($id in $current) {
            $sid = [string]$id
            if ($seen.Contains($sid)) { continue }
            [void]$seen.Add($sid)
            $levelIds += $id
        }
        if ($levelIds.Count -eq 0) { break }
        [void]$levels.Add($levelIds)
        $next = @()
        foreach ($lid in $levelIds) {
            $memberIds = Get-GroupDirectMemberGroupIds -GroupId $lid
            foreach ($mid in $memberIds) {
                if ($DeviceGroupIdsStr -contains [string]$mid -and -not $seen.Contains([string]$mid)) { $next += $mid }
            }
        }
        $current = $next
    }
    $orderedIds = @()
    for ($i = $levels.Count - 1; $i -ge 0; $i--) { $orderedIds += $levels[$i] }
    $names = @()
    foreach ($nid in $orderedIds) {
        $ng = $script:AllGroups[$nid]
        if ($ng -and $ng.displayName) { $names += $ng.displayName }
    }
    return $names
}

function Get-GroupInfo {
    param([Parameter(Mandatory)][string]$GroupId)
    if ($script:AllGroups.Count -gt 0) {
        $group = $script:AllGroups[$GroupId]
        if ($group) { return @{ Id = $group.id; DisplayName = $group.displayName; Success = $true } }
    }
    try {
        $group = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId" -Method Get
        return @{ Id = $group.id; DisplayName = $group.displayName; Success = $true }
    }
    catch {
        return @{ Id = $GroupId; DisplayName = "Unknown Group"; Success = $false }
    }
}

#region Assignment collection
function Get-DetailedPolicyAssignments {
    param([Parameter(Mandatory)][string]$EntityType, [string]$EntityId, [string]$PolicyName, [switch]$DebugMode)
    
    $assignmentsUri = $null
    if ($EntityType -eq "deviceAppManagement/managedAppPolicies") {
        try {
            $r = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/managedAppPolicies/$EntityId" -Method Get
            $path = switch ($r.'@odata.type') {
                "#microsoft.graph.androidManagedAppProtection" { "androidManagedAppProtections" }
                "#microsoft.graph.iosManagedAppProtection" { "iosManagedAppProtections" }
                "#microsoft.graph.windowsManagedAppProtection" { "windowsManagedAppProtections" }
                default { $null }
            }
            if ($path) { $assignmentsUri = "https://graph.microsoft.com/beta/deviceAppManagement/$path('$EntityId')/assignments" }
        }
        catch { return @() }
    }
    elseif ($EntityType -eq "mobileAppConfigurations") {
        $assignmentsUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations('$EntityId')/assignments"
    }
    elseif ($EntityType -eq "deviceAppManagement/mobileApps") {
        $assignmentsUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps('$EntityId')/assignments"
    }
    elseif ($EntityType -like "deviceAppManagement/*ManagedAppProtections") {
        $assignmentsUri = "https://graph.microsoft.com/beta/$EntityType('$EntityId')/assignments"
    }
    elseif ($EntityType -eq "deviceManagement/intents") {
        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/intents/$EntityId/assignments"
    }
    elseif ($EntityType -eq "deviceManagement/templates") {
        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/templates/$EntityId/assignments"
    }
    elseif ($EntityType -like "virtualEndpoint/*") {
        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/$EntityType/$EntityId/assignments"
    }
    elseif ($EntityType -eq "configurationPolicies") {
        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$EntityId/assignments"
    }
    else {
        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/$EntityType('$EntityId')/assignments"
    }
    
    if (-not $assignmentsUri) { return @() }
    
    try {
        Write-Verbose "        Querying assignments for: $PolicyName ($EntityType, $EntityId)"
        Write-Verbose "        Assignments URI: $assignmentsUri"
        $assignments = Invoke-GraphRequestWithPaging -Uri $assignmentsUri -Method "GET" -DebugMode:$DebugMode
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
        return @()
    }
}

function Get-DetailedPolicyAssignmentsFromExpanded {
    param([Parameter(Mandatory)]$Policy, [Parameter(Mandatory)][string]$PolicyName, [Parameter(Mandatory)]$PolicyType, [switch]$DebugMode)
    
    $detailedAssignments = @()
    if (-not $Policy.assignments) { return @() }
    
    foreach ($assignment in $Policy.assignments) {
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
            PolicyId = $Policy.id
            PolicyType = $PolicyType.EntityType
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

# Cloud PC policies: Use the same approach as manual calls: (1) GET policy?$expand=assignments with default SDK response so assignments is Hashtable like { 28f33e7b-... }; (2) each key is the group ID - GET groups/{id} for displayName/rule. No -OutputType Json so we get the native Hashtable.
# Normalize a group so membershipRule is only set for Dynamic groups (Assigned groups may have license info in membershipRule from API - we show empty).
function Normalize-EntraGroupForDisplay {
    param([Parameter(Mandatory)][object]$Group)
    if (-not $Group) { return $null }
    $id = $Group.id; $displayName = $Group.displayName; $groupTypes = $Group.groupTypes
    if ($Group -is [System.Collections.IDictionary]) {
        if ($Group.ContainsKey('id')) { $id = $Group['id'] }; if ($Group.ContainsKey('displayName')) { $displayName = $Group['displayName'] }; if ($Group.ContainsKey('groupTypes')) { $groupTypes = $Group['groupTypes'] }
    }
    $isDynamic = $groupTypes -and (@($groupTypes) -contains 'DynamicMembership')
    $membershipRule = if ($isDynamic) {
        $r = $Group.membershipRule; if ($Group -is [System.Collections.IDictionary] -and $Group.ContainsKey('membershipRule')) { $r = $Group['membershipRule'] }
        if (-not [string]::IsNullOrWhiteSpace($r)) { $r } else { $null }
    } else { $null }
    return [PSCustomObject]@{ id = $id; displayName = $displayName; groupTypes = $groupTypes; membershipRule = $membershipRule }
}
function Get-CloudPCPolicyWithAssignments {
    param([Parameter(Mandatory)][string]$EntityType, [Parameter(Mandatory)][string]$EntityId, [switch]$DebugMode)
    $baseUri = "https://graph.microsoft.com/beta/deviceManagement/$EntityType/$EntityId"
    $policy = $null
    # User's approach: GET with expand (no -OutputType Json) → assignments may be Hashtable { groupId } = value, or PSCustomObject with GUID-named properties.
    try {
        $policy = Invoke-MgGraphRequest -Uri ($baseUri + "?`$expand=assignments") -Method GET -ErrorAction Stop
        if (-not $policy) { }
        else {
            $assignmentsRaw = $policy.assignments; if ($null -eq $assignmentsRaw) { $assignmentsRaw = $policy.Assignments }
            $groupIds = @()
            if ($assignmentsRaw -is [System.Collections.IDictionary]) {
                $keys = @($assignmentsRaw.Keys)
                foreach ($key in $keys) {
                    $keyStr = [string]$key
                    if ($keyStr -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') { $groupIds += $keyStr }
                }
                if ($groupIds.Count -eq 0) {
                    foreach ($val in $assignmentsRaw.Values) {
                        $gid = $null
                        if ($val -is [System.Collections.IDictionary]) {
                            if ($val.ContainsKey('target') -and $val['target']) { $t = $val['target']; if ($t -is [System.Collections.IDictionary] -and $t.ContainsKey('groupId')) { $gid = $t['groupId'] } }
                            if (-not $gid -and $val.ContainsKey('groupId')) { $gid = $val['groupId'] }
                        } else { if ($val.target) { $gid = $val.target.groupId }; if (-not $gid) { $gid = $val.groupId } }
                        if ($gid -and $groupIds -notcontains $gid) { $groupIds += $gid }
                    }
                }
            }
            elseif ($assignmentsRaw -is [PSCustomObject] -and $assignmentsRaw.PSObject -and $assignmentsRaw.PSObject.Properties) {
                foreach ($p in $assignmentsRaw.PSObject.Properties) {
                    $keyStr = [string]$p.Name
                    if ($keyStr -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                        if ($groupIds -notcontains $keyStr) { $groupIds += $keyStr }
                    } else {
                        $v = $p.Value
                        if ($v) {
                            $gid = $null
                            if ($v -is [System.Collections.IDictionary]) {
                                if ($v.ContainsKey('target') -and $v['target']) { $t = $v['target']; if ($t -is [System.Collections.IDictionary] -and $t.ContainsKey('groupId')) { $gid = $t['groupId'] } }
                                if (-not $gid -and $v.ContainsKey('groupId')) { $gid = $v['groupId'] }
                            } else { if ($v.target) { $gid = $v.target.groupId }; if (-not $gid) { $gid = $v.groupId } }
                            if ($gid -and $groupIds -notcontains $gid) { $groupIds += $gid }
                        }
                    }
                }
            }
            if ($groupIds.Count -gt 0) {
                $assignmentsList = [System.Collections.ArrayList]::new()
                foreach ($groupId in $groupIds) {
                    try {
                        $groupJson = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$groupId`?`$select=id,displayName,membershipRule,groupTypes" -Method GET -OutputType Json -ErrorAction Stop
                        if (-not [string]::IsNullOrWhiteSpace($groupJson)) {
                            $group = $groupJson | ConvertFrom-Json
                            if ($group) {
                                $groupToAdd = Normalize-EntraGroupForDisplay -Group $group
                                if ($groupToAdd) {
                                    $script:AllGroups[$groupToAdd.id] = $groupToAdd
                                }
                            }
                        }
                    }
                    catch { if ($DebugMode) { Write-Warning "Get-CloudPCPolicyWithAssignments group $groupId : $($_.Exception.Message)" } }
                    $null = $assignmentsList.Add([PSCustomObject]@{ id = $groupId; target = [PSCustomObject]@{ groupId = $groupId; '@odata.type' = '#microsoft.graph.cloudPcManagementGroupAssignmentTarget' } })
                }
                try { if ($policy.PSObject -and $policy.PSObject.Properties['assignments']) { $policy.PSObject.Properties.Remove('assignments') } } catch { }
                try { if ($policy -is [System.Collections.IDictionary] -and $policy.ContainsKey('assignments')) { $policy.Remove('assignments') } } catch { }
                $policy | Add-Member -NotePropertyName 'assignments' -NotePropertyValue $assignmentsList.ToArray() -Force
                return $policy
            }
        }
    }
    catch { if ($DebugMode) { Write-Warning "Get-CloudPCPolicyWithAssignments (expand default) $EntityType $EntityId : $($_.Exception.Message)" } }
    # Fallback: Json expand (for when first GET failed or returned no policy)
    try {
        $expandJson = Invoke-MgGraphRequest -Uri ($baseUri + "?`$expand=assignments") -Method GET -OutputType Json -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($expandJson)) {
            $policy = $expandJson | ConvertFrom-Json -NoEnumerate
            if ($policy -and $policy.assignments) {
                $a = $policy.assignments
                $groupIds = @()
                if ($a -is [System.Collections.IDictionary]) {
                    foreach ($key in $a.Keys) {
                        $keyStr = [string]$key
                        if ($keyStr -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') { $groupIds += $keyStr }
                    }
                }
                elseif ($a -is [PSCustomObject] -and $a.PSObject -and $a.PSObject.Properties) {
                    foreach ($p in $a.PSObject.Properties) {
                        $keyStr = [string]$p.Name
                        if ($keyStr -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                            if ($groupIds -notcontains $keyStr) { $groupIds += $keyStr }
                        } else {
                            $v = $p.Value; if ($v) { $gid = $v.target.groupId; if ($v.groupId) { $gid = $v.groupId }; if ($gid -and $groupIds -notcontains $gid) { $groupIds += $gid } }
                        }
                    }
                }
                if ($groupIds.Count -gt 0) {
                    $assignmentsList = [System.Collections.ArrayList]::new()
                    foreach ($groupId in $groupIds) {
                        try {
                            $groupJson = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$groupId`?`$select=id,displayName,membershipRule,groupTypes" -Method GET -OutputType Json -ErrorAction Stop
                            if (-not [string]::IsNullOrWhiteSpace($groupJson)) {
                                $group = $groupJson | ConvertFrom-Json
                                if ($group) {
                                    $groupToAdd = Normalize-EntraGroupForDisplay -Group $group
                                    if ($groupToAdd) {
                                        $script:AllGroups[$groupToAdd.id] = $groupToAdd
                                    }
                                }
                            }
                        } catch { }
                        $null = $assignmentsList.Add([PSCustomObject]@{ id = $groupId; target = [PSCustomObject]@{ groupId = $groupId; '@odata.type' = '#microsoft.graph.cloudPcManagementGroupAssignmentTarget' } })
                    }
                    try { if ($policy.PSObject.Properties['assignments']) { $policy.PSObject.Properties.Remove('assignments') } } catch { }
                    $policy | Add-Member -NotePropertyName 'assignments' -NotePropertyValue $assignmentsList.ToArray() -Force
                    return $policy
                }
                $assignmentsList = if ($null -eq $a) { @() } elseif ($a -is [System.Collections.IEnumerable] -and $a -isnot [string]) { @($a) } else { @($a) }
                if ($assignmentsList.Count -gt 0) {
                    try { if ($policy.PSObject.Properties['assignments']) { $policy.PSObject.Properties.Remove('assignments') } } catch { }
                    $policy | Add-Member -NotePropertyName 'assignments' -NotePropertyValue $assignmentsList -Force
                    return $policy
                }
            }
        }
    }
    catch { if ($DebugMode) { Write-Warning "Get-CloudPCPolicyWithAssignments (expand Json) $EntityType $EntityId : $($_.Exception.Message)" } }
    # Fallback: GET policy without expand, then try GET .../assignments (may 404 for Cloud PC)
    try {
        $jsonString = Invoke-MgGraphRequest -Uri $baseUri -Method GET -OutputType Json -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($jsonString)) {
            $policy = $jsonString | ConvertFrom-Json
        }
    }
    catch { if ($DebugMode) { Write-Warning "Get-CloudPCPolicyWithAssignments (policy GET) $EntityType $EntityId : $($_.Exception.Message)" } }
    if (-not $policy) {
        try {
            $policy = Invoke-MgGraphRequest -Uri $baseUri -Method GET -OutputType PSObject -ErrorAction Stop
        }
        catch {
            if ($DebugMode) { Write-Warning "Get-CloudPCPolicyWithAssignments $EntityType $EntityId : $($_.Exception.Message)" }
            return $null
        }
    }
    if (-not $policy) { return $null }
    $assignmentsList = $null
    try {
        $assignmentsUri = "$baseUri/assignments"
        $assignmentsJson = Invoke-MgGraphRequest -Uri $assignmentsUri -Method GET -OutputType Json -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($assignmentsJson)) {
            $assignmentsResp = $assignmentsJson | ConvertFrom-Json -NoEnumerate
            if ($assignmentsResp.value) { $assignmentsList = @($assignmentsResp.value) }
            elseif ($assignmentsResp -is [Array]) { $assignmentsList = @($assignmentsResp) }
        }
    }
    catch { if ($DebugMode) { Write-Warning "Get-CloudPCPolicyWithAssignments GET assignments (Json) $EntityId : $($_.Exception.Message)" } }
    if ($null -eq $assignmentsList -or $assignmentsList.Count -eq 0) {
        try {
            $assignmentsRaw = Invoke-MgGraphRequest -Uri "$baseUri/assignments" -Method GET -ErrorAction Stop
            if ($assignmentsRaw) {
                if ($assignmentsRaw -is [System.Collections.IDictionary]) {
                    if ($assignmentsRaw.ContainsKey('value')) { $assignmentsList = @($assignmentsRaw['value']) }
                    elseif ($assignmentsRaw.ContainsKey('Value')) { $assignmentsList = @($assignmentsRaw['Value']) }
                }
                elseif ($assignmentsRaw.PSObject.Properties['value']) { $assignmentsList = @($assignmentsRaw.value) }
                elseif ($assignmentsRaw.PSObject.Properties['Value']) { $assignmentsList = @($assignmentsRaw.Value) }
            }
        }
        catch { if ($DebugMode) { Write-Warning "Get-CloudPCPolicyWithAssignments GET assignments (default) $EntityId : $($_.Exception.Message)" } }
    }
    if ($null -eq $assignmentsList -or $assignmentsList.Count -eq 0) {
        $assignmentsFromPolicy = $policy.assignments
        if ($null -eq $assignmentsFromPolicy) { $assignmentsFromPolicy = $policy.Assignments }
        if ($assignmentsFromPolicy -is [System.Collections.IDictionary]) {
            if ($assignmentsFromPolicy.ContainsKey('target') -or $assignmentsFromPolicy.ContainsKey('Target')) {
                $assignmentsList = @($assignmentsFromPolicy)
            }
            elseif ($assignmentsFromPolicy.ContainsKey('value')) {
                $assignmentsList = @($assignmentsFromPolicy['value'])
            }
            else {
                $assignmentsList = @($assignmentsFromPolicy.Values)
            }
        }
        elseif ($assignmentsFromPolicy -ne $null) {
            $assignmentsList = @($assignmentsFromPolicy)
        }
    }
    if ($assignmentsList -and $assignmentsList.Count -gt 0) {
        try {
            if ($policy.PSObject.Properties['assignments']) { $policy.PSObject.Properties.Remove('assignments') }
        } catch { }
        try {
            if ($policy -is [System.Collections.IDictionary] -and $policy.ContainsKey('assignments')) { $policy.Remove('assignments') }
        } catch { }
        $policy | Add-Member -NotePropertyName 'assignments' -NotePropertyValue $assignmentsList -Force
    }
    return $policy
}

function Get-CloudPCProvisioningConfigByDisplayName {
    param([Parameter(Mandatory)][string]$DisplayName)
    try {
        $policies = Get-IntuneEntities -EntityType "virtualEndpoint/provisioningPolicies" -DebugMode:$false
        $displayNameNorm = ($DisplayName -replace '\s+', ' ').Trim()
        $match = $policies | Where-Object { $_.displayName -and (($_.displayName -replace '\s+', ' ').Trim() -eq $displayNameNorm) } | Select-Object -First 1
        if (-not $match -or -not $match.id) { return $null }
        $policy = Get-CloudPCPolicyWithAssignments -EntityType "virtualEndpoint/provisioningPolicies" -EntityId $match.id -DebugMode:$false
        return $policy
    }
    catch { return $null }
}

function Get-CloudPCUserSettingsConfigByDisplayName {
    param([Parameter(Mandatory)][string]$DisplayName)
    try {
        $policies = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings" -Method GET
        $displayNameNorm = ($DisplayName -replace '\s+', ' ').Trim()
        $match = $policies | Where-Object { $_.displayName -and (($_.displayName -replace '\s+', ' ').Trim() -eq $displayNameNorm) } | Select-Object -First 1
        if (-not $match) { return $null }
        $policyId = $match.id
        $getUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings/$policyId"
        $policy = Invoke-MgGraphRequest -Uri $getUri -Method GET -OutputType PSObject -ErrorAction Stop
        if (-not $policy) { return $null }
        
        $toYesNo = $script:ToYesNo
        
        $name = if ($policy.displayName) { $policy.displayName } else { "-" }
        $selfServiceEnabled = & $toYesNo $policy.selfServiceEnabled
        $localAdminEnabled = & $toYesNo $policy.localAdminEnabled
        $resetEnabled = & $toYesNo $policy.resetEnabled
        
        # Restore Point Settings
        $restorePointSetting = $policy.restorePointSetting
        $frequencyType = "-"
        if ($restorePointSetting -and $restorePointSetting.frequencyType) {
            $freqType = $restorePointSetting.frequencyType
            $frequencyType = switch ($freqType) {
                "default" { "Default" }
                "everyFourHours" { "Every 4 hours" }
                "fourHours" { "Every 4 hours" }
                "everySixHours" { "Every 6 hours" }
                "sixHours" { "Every 6 hours" }
                "everyTwelveHours" { "Every 12 hours" }
                "twelveHours" { "Every 12 hours" }
                "everyTwentyFourHours" { "Every 24 hours" }
                "twentyFourHours" { "Every 24 hours" }
                default { $freqType }
            }
        }
        $userRestoreEnabled = if ($restorePointSetting -and $null -ne $restorePointSetting.userRestoreEnabled) { & $toYesNo $restorePointSetting.userRestoreEnabled } else { "-" }
        
        # Cross Region Disaster Recovery Settings
        $drSetting = $policy.crossRegionDisasterRecoverySetting
        $drEnabled = "None"
        $userInitiatedDRAllowed = "-"
        if ($drSetting) {
            if ($drSetting.crossRegionDisasterRecoveryEnabled -eq $true) {
                $drEnabled = "Enabled"
            }
            $userInitiatedDRAllowed = & $toYesNo $drSetting.userInitiatedDisasterRecoveryAllowed
        }
        
        # Notification Settings
        $notificationSetting = $policy.notificationSetting
        $restartPromptsDisabled = if ($notificationSetting -and $null -ne $notificationSetting.restartPromptsDisabled) { & $toYesNo $notificationSetting.restartPromptsDisabled } else { "-" }
        
        return @{
            Name = $name
            SelfServiceEnabled = $selfServiceEnabled
            LocalAdminEnabled = $localAdminEnabled
            ResetEnabled = $resetEnabled
            RestorePointFrequency = $frequencyType
            UserRestoreEnabled = $userRestoreEnabled
            DisasterRecoveryEnabled = $drEnabled
            UserInitiatedDRAllowed = $userInitiatedDRAllowed
            RestartPromptsDisabled = $restartPromptsDisabled
        }
    }
    catch { return $null }
}

function Get-CloudPCPolicyGroupInfoInternal {
    [CmdletBinding()]
    param(
        [Parameter(ParameterSetName = 'ByName')][string]$PolicyName,
        [Parameter(ParameterSetName = 'ById')][string]$PolicyId,
        [Parameter(ParameterSetName = 'ById')][string]$EntityType = 'virtualEndpoint/provisioningPolicies',
        [Parameter(ParameterSetName = 'ById')][switch]$UpdateAllGroups,
        [switch]$DebugMode
    )
    if ($PSCmdlet.ParameterSetName -eq 'ByName' -and $PolicyName) {
        try { $list = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies' -Method GET -ErrorAction Stop }
        catch { Write-Error "Failed to list provisioning policies: $_"; return }
        $value = if ($list.value) { $list.value } else { @($list) }
        $policy = $value | Where-Object { $_.displayName -eq $PolicyName } | Select-Object -First 1
        if (-not $policy) { Write-Error "Provisioning policy not found: '$PolicyName'"; return }
        $PolicyId = $policy.id
        $EntityType = 'virtualEndpoint/provisioningPolicies'
    }
    if (-not $PolicyId) { Write-Error "Provide -PolicyName or -PolicyId."; return }
    if (-not $EntityType) { $EntityType = 'virtualEndpoint/provisioningPolicies' }
    $baseUri = "https://graph.microsoft.com/beta/deviceManagement/$EntityType/$PolicyId"
    try { $policyWithAssignments = Invoke-MgGraphRequest -Uri ($baseUri + "?`$expand=assignments") -Method GET -ErrorAction Stop }
    catch { Write-Error "Failed to get policy: $_"; return }
    
    # Debug: Check what we received
    if ($DebugMode) {
        Write-Host "          [DEBUG] Retrieved policy with assignments. Checking assignments property..." -ForegroundColor DarkGray
    }
    
    $groupIds = @()
    $assignments = $policyWithAssignments.assignments
    if ($null -eq $assignments) { $assignments = $policyWithAssignments.Assignments }
    
    # Debug: Log assignments state
    if ($DebugMode) {
        if ($null -eq $assignments) {
            Write-Host "          [DEBUG] Assignments is NULL" -ForegroundColor DarkYellow
        } else {
            $assignmentsType = $assignments.GetType().FullName
            $assignmentsCount = if ($assignments -is [System.Collections.IEnumerable] -and $assignments -isnot [string]) { @($assignments).Count } else { "N/A (not enumerable)" }
            Write-Host "          [DEBUG] Assignments type: $assignmentsType, Count: $assignmentsCount" -ForegroundColor DarkGray
        }
    }
    
    # Extract group IDs from assignments (handle IDictionary, Array, or PSCustomObject formats)
    if ($assignments -is [System.Collections.IDictionary]) {
        if ($DebugMode) { Write-Host "          [DEBUG] Processing as IDictionary (Hashtable)" -ForegroundColor DarkGray }
        foreach ($key in $assignments.Keys) {
            $keyStr = [string]$key
            if ($keyStr -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') { $groupIds += $keyStr }
        }
        if ($groupIds.Count -eq 0) {
            foreach ($val in $assignments.Values) {
                $gid = $null
                if ($val -is [System.Collections.IDictionary]) {
                    if ($val.ContainsKey('target') -and $val['target']) { $t = $val['target']; if ($t -is [System.Collections.IDictionary] -and $t.ContainsKey('groupId')) { $gid = $t['groupId'] } }
                    if (-not $gid -and $val.ContainsKey('groupId')) { $gid = $val['groupId'] }
                } else { if ($val.target) { $gid = $val.target.groupId }; if (-not $gid) { $gid = $val.groupId } }
                if ($gid -and $groupIds -notcontains $gid) { $groupIds += $gid }
            }
        }
    } elseif ($assignments -is [System.Collections.IEnumerable] -and $assignments -isnot [string]) {
        if ($DebugMode) { Write-Host "          [DEBUG] Processing as IEnumerable (Array)" -ForegroundColor DarkGray }
        foreach ($a in @($assignments)) {
            $gid = $null
            if ($a -is [System.Collections.IDictionary]) {
                if ($a.ContainsKey('target')) { $t = $a['target']; if ($t -is [System.Collections.IDictionary] -and $t.ContainsKey('groupId')) { $gid = $t['groupId'] } }
                if (-not $gid -and $a.ContainsKey('groupId')) { $gid = $a['groupId'] }
            } else { 
                # PSCustomObject with properties
                if ($a.target) { $gid = $a.target.groupId }
                if (-not $gid) { $gid = $a.groupId }
            }
            if ($gid) {
                if ($DebugMode) { Write-Host "          [DEBUG] Found groupId: $gid" -ForegroundColor DarkGreen }
                if ($groupIds -notcontains $gid) { $groupIds += $gid }
            }
        }
    }
    elseif ($assignments -is [PSCustomObject] -and $assignments.PSObject -and $assignments.PSObject.Properties) {
        if ($DebugMode) { Write-Host "          [DEBUG] Processing as PSCustomObject" -ForegroundColor DarkGray }
        foreach ($p in $assignments.PSObject.Properties) {
            $keyStr = [string]$p.Name
            if ($keyStr -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                if ($groupIds -notcontains $keyStr) { $groupIds += $keyStr }
            } else {
                $v = $p.Value; if ($v) { $gid = $null; if ($v.target) { $gid = $v.target.groupId }; if (-not $gid -and $v.groupId) { $gid = $v.groupId }; if ($gid -and $groupIds -notcontains $gid) { $groupIds += $gid } }
            }
        }
    } else {
        if ($DebugMode) { 
            $actualType = if ($assignments) { $assignments.GetType().FullName } else { "null" }
            Write-Host "          [DEBUG] Assignments type not recognized: $actualType" -ForegroundColor DarkYellow 
        }
    }
    
    if ($DebugMode) {
        Write-Host "          [DEBUG] Extracted $($groupIds.Count) group ID(s): $($groupIds -join ', ')" -ForegroundColor DarkGray
    }
    $policyDisplayName = $null
    if ($policyWithAssignments -is [System.Collections.IDictionary]) {
        if ($policyWithAssignments.ContainsKey('displayName')) { $policyDisplayName = $policyWithAssignments['displayName'] }
        elseif ($policyWithAssignments.ContainsKey('DisplayName')) { $policyDisplayName = $policyWithAssignments['DisplayName'] }
    }
    if (-not $policyDisplayName -and $policyWithAssignments.PSObject.Properties['displayName']) { $policyDisplayName = $policyWithAssignments.displayName }
    if (-not $policyDisplayName) { $policyDisplayName = $policyWithAssignments.displayName }
    if (-not $policyDisplayName) { $policyDisplayName = "Policy $PolicyId" }
    if ($groupIds.Count -eq 0) { 
        if ($DebugMode) { Write-Host "          [DEBUG] No group IDs found. Returning 'Not Assigned'" -ForegroundColor DarkYellow }
        return [PSCustomObject]@{ PolicyId = $PolicyId; PolicyName = $policyDisplayName; GroupId = $null; GroupName = 'Not Assigned'; GroupType = '-'; MembershipRule = '-' } 
    }
    
    $results = @()
    foreach ($gid in $groupIds) {
        if ($DebugMode) { Write-Host "          [DEBUG] Looking up group: $gid" -ForegroundColor DarkGray }
        try { 
            $groupJson = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$gid`?`$select=id,displayName,membershipRule,groupTypes" -Method GET -OutputType Json -ErrorAction Stop
            $group = if ($groupJson) { $groupJson | ConvertFrom-Json } else { $null }
            if ($DebugMode -and $group) { Write-Host "          [DEBUG] Group found: $($group.displayName)" -ForegroundColor DarkGreen }
        }
        catch { 
            if ($DebugMode) { Write-Host "          [DEBUG] Group lookup failed: $($_.Exception.Message)" -ForegroundColor DarkYellow }
            $group = $null 
        }
        if (-not $group) { 
            if ($DebugMode) { Write-Host "          [DEBUG] Adding assignment with group ID only (lookup failed)" -ForegroundColor DarkGray }
            $results += [PSCustomObject]@{ PolicyId = $PolicyId; PolicyName = $policyDisplayName; GroupId = $gid; GroupName = "Group (ID: $gid)"; GroupType = '-'; MembershipRule = '-' }
            continue 
        }
        $normalized = Normalize-EntraGroupForDisplay -Group $group
        $displayName = $normalized.displayName; if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = "Group (ID: $gid)" }
        if ($UpdateAllGroups) {
            if (-not $script:AllGroups.ContainsKey($gid)) {
                $script:AllGroups[$gid] = [PSCustomObject]@{ id = $gid; displayName = $displayName; groupTypes = $normalized.groupTypes; membershipRule = $normalized.membershipRule }
            }
        }
        $isDynamic = $normalized.groupTypes -and (@($normalized.groupTypes) -contains 'DynamicMembership')
        $groupType = if ($isDynamic) { 'Dynamic' } else { 'Static' }
        $rule = if ($isDynamic -and $normalized.membershipRule) { $normalized.membershipRule } else { '' }
        if ($DebugMode) { Write-Host "          [DEBUG] Adding assignment for group: $displayName (ID: $gid)" -ForegroundColor DarkGreen }
        $results += [PSCustomObject]@{ PolicyId = $PolicyId; PolicyName = $policyDisplayName; GroupId = $gid; GroupName = $displayName; GroupType = $groupType; MembershipRule = $rule }
    }
    if ($DebugMode) { Write-Host "          [DEBUG] Returning $($results.Count) result(s)" -ForegroundColor DarkGray }
    if ($results.Count -eq 1) { return $results[0] }; return $results
}

function Get-AllPolicyAssignmentsBatch {
    param([Parameter(Mandatory)][array]$PolicyTypes, [switch]$DebugMode)
    $allPolicyAssignments = [System.Collections.ArrayList]::new()
    $GraphEndpoint = "https://graph.microsoft.com"
    if ($script:AllFilters.Count -eq 0) {
        $rawFilters = Invoke-GraphRequestWithPaging -Uri "$GraphEndpoint/beta/deviceManagement/assignmentFilters" -Method "GET" -DebugMode:$DebugMode
        if ($rawFilters) { foreach ($f in $rawFilters) { $script:AllFilters[$f.id] = $f } }
    }
    foreach ($policyType in $PolicyTypes) {
        Write-Host "  Processing $($policyType.DisplayName)..." -ForegroundColor Cyan
        try {
            $selectPlatforms = ""
            if ($policyType.EntityType -eq "deviceManagement/intents") {
                # deviceManagementIntent has no 'name' property; include templateId to detect Security Baselines (created from template)
                $selectPlatforms = "id,displayName,platforms,createdDateTime,lastModifiedDateTime,templateId"
            }
            # configurationPolicies (Settings Catalog): fetch without $select to avoid API property differences
            $policies = if ($policyType.Name -eq "Applications") {
                Get-IntuneEntities -EntityType $policyType.EntityType -Expand "assignments" -DebugMode:$DebugMode
            }
            elseif ($policyType.Name -eq "Cloud PC Provisioning" -or $policyType.Name -eq "Cloud PC User Settings") {
                # Cloud PC: list does not return assignments; get each policy by ID with $expand=assignments
                Get-IntuneEntities -EntityType $policyType.EntityType -DebugMode:$DebugMode
            }
            elseif ($selectPlatforms) {
                Get-IntuneEntities -EntityType $policyType.EntityType -Select $selectPlatforms -DebugMode:$DebugMode
            }
            else {
                Get-IntuneEntities -EntityType $policyType.EntityType -DebugMode:$DebugMode
            }
            Write-Host "    Found $($policies.Count) policies for $($policyType.DisplayName)" -ForegroundColor Yellow
            if ($policies.Count -gt 0) {
                foreach ($policy in $policies) {
                    $policyName = if ($policy.displayName) { $policy.displayName } elseif ($policy.name) { $policy.name } else { "Unnamed Policy" }
                    $policyPlatform = Get-PolicyPlatformFromPolicy -Policy $policy -EntityType $policyType.EntityType
                    # Intents created from a template (templateId set) = Security Baseline; others = Endpoint Security
                    $effectiveCategory = $policyType.Name
                    if ($policyType.EntityType -eq "deviceManagement/intents" -and $policy.templateId) { $effectiveCategory = "Security Baselines" }
                    # Settings Catalog policies that are security baseline / version monitoring → show under Security Baselines in Overview
                    if ($policyType.EntityType -eq "configurationPolicies") {
                        $nameForMatch = $policyName
                        if (($nameForMatch -match 'Baseline' -and $nameForMatch -match 'Version Monitoring') -or ($nameForMatch -match 'Security Baseline')) { $effectiveCategory = "Security Baselines" }
                    }
                    Write-Host "      Processing policy: $policyName (ID: $($policy.id))" -ForegroundColor Gray
                    try {
                        $detailedAssignments = if ($policyType.Name -eq "Applications" -and $policy.assignments) {
                            Get-DetailedPolicyAssignmentsFromExpanded -Policy $policy -PolicyName $policyName -PolicyType $policyType -DebugMode:$DebugMode
                        }
                        elseif ($policyType.Name -eq "Cloud PC Provisioning" -or $policyType.Name -eq "Cloud PC User Settings") {
                            $policyId = $policy.id
                            if ($null -eq $policyId -and $policy.PSObject.Properties) { $policyId = ($policy.PSObject.Properties | Where-Object { $_.Name -eq 'id' } | Select-Object -First 1).Value }
                            if ($null -eq $policyId -and $policy -is [System.Collections.IDictionary]) {
                                if ($policy.ContainsKey('id')) { $policyId = $policy['id'] } elseif ($policy.ContainsKey('Id')) { $policyId = $policy['Id'] }
                            }
                            $policyId = [string]$policyId
                            $cloudPCAssignments = @()
                            if (-not [string]::IsNullOrWhiteSpace($policyId)) {
                                try {
                                    $raw = Get-CloudPCPolicyGroupInfoInternal -PolicyId $policyId -EntityType $policyType.EntityType -UpdateAllGroups -DebugMode:$DebugMode
                                    if ($DebugMode) { 
                                        $rawType = if ($raw) { $raw.GetType().FullName } else { "null" }
                                        Write-Host "          [DEBUG] Get-CloudPCPolicyGroupInfoInternal returned: Type=$rawType" -ForegroundColor DarkGray 
                                    }
                                    $infos = if ($raw) { @($raw) } else { @() }
                                    if ($DebugMode) { Write-Host "          [DEBUG] Processing $($infos.Count) info object(s)" -ForegroundColor DarkGray }
                                    foreach ($info in $infos) {
                                        if ($DebugMode) { 
                                            Write-Host "          [DEBUG] Info object: GroupId=$($info.GroupId), GroupName=$($info.GroupName)" -ForegroundColor DarkGray 
                                        }
                                        if ($info.GroupId) {
                                            $cloudPCAssignments += [PSCustomObject]@{
                                                PolicyName = $info.PolicyName; PolicyId = $info.PolicyId; PolicyType = $policyType.EntityType
                                                AssignmentType = "Group (Include)"; AssignmentIntent = $null; TargetName = $info.GroupName; TargetId = $info.GroupId; GroupId = $info.GroupId
                                                FilterId = $null; FilterName = "No Filter"; FilterType = "None"; FilterRule = $null; FilterPlatform = $null; AssignmentId = $null
                                            }
                                            if ($DebugMode) { Write-Host "          [DEBUG] Added assignment for: $($info.GroupName)" -ForegroundColor DarkGreen }
                                        } else {
                                            if ($DebugMode) { Write-Host "          [DEBUG] Skipped (no GroupId): $($info.GroupName)" -ForegroundColor DarkYellow }
                                        }
                                    }
                                } catch { if ($DebugMode) { Write-Warning "Cloud PC assignments $policyName : $($_.Exception.Message)" } }
                            }
                            # Return the assignments array so it gets assigned to $detailedAssignments
                            $cloudPCAssignments
                        }
                        else {
                            Get-DetailedPolicyAssignments -EntityType $policyType.EntityType -EntityId $policy.id -PolicyName $policyName -DebugMode:$DebugMode
                        }
                        Write-Host "        Found $($detailedAssignments.Count) assignments for $policyName" -ForegroundColor Magenta
                        if ($detailedAssignments.Count -eq 0) {
                            $null = $allPolicyAssignments.Add([PSCustomObject]@{
                                    PolicyCategory = $effectiveCategory; PolicyName = $policyName; PolicyId = $policy.id; PolicyType = $policyType.EntityType; PolicyPlatform = $policyPlatform
                                    AssignmentType = "Not Assigned"; TargetName = "Not Assigned"; TargetId = $null; GroupId = $null; FilterId = $null; FilterName = "No Filter"; FilterType = "None"; FilterRule = $null; FilterPlatform = $null; AssignmentId = $null; CreatedDateTime = $policy.createdDateTime; LastModifiedDateTime = $policy.lastModifiedDateTime
                                })
                        }
                        else {
                            foreach ($assignment in $detailedAssignments) {
                                $null = $allPolicyAssignments.Add([PSCustomObject]@{
                                        PolicyCategory = $effectiveCategory; PolicyName = $assignment.PolicyName; PolicyId = $assignment.PolicyId; PolicyType = $assignment.PolicyType; PolicyPlatform = $policyPlatform
                                        AssignmentType = $assignment.AssignmentType; AssignmentIntent = if ($assignment.AssignmentIntent) { $assignment.AssignmentIntent } else { $null }; TargetName = $assignment.TargetName; TargetId = $assignment.TargetId; GroupId = $assignment.GroupId
                                        FilterId = $assignment.FilterId; FilterName = $assignment.FilterName; FilterType = $assignment.FilterType; FilterRule = $assignment.FilterRule; FilterPlatform = $assignment.FilterPlatform; AssignmentId = $assignment.AssignmentId
                                        CreatedDateTime = $policy.createdDateTime; LastModifiedDateTime = $policy.lastModifiedDateTime
                                    })
                            }
                        }
                    }
                    catch {
                        Write-Warning "        Error processing assignments for ${policyName}: $($_.Exception.Message)"
                        $null = $allPolicyAssignments.Add([PSCustomObject]@{
                                PolicyCategory = $effectiveCategory; PolicyName = $policyName; PolicyId = $policy.id; PolicyType = $policyType.EntityType; PolicyPlatform = $policyPlatform
                                AssignmentType = "Not Assigned"; TargetName = "Not Assigned"; TargetId = $null; GroupId = $null; FilterId = $null; FilterName = "No Filter"; FilterType = "None"; FilterRule = $null; FilterPlatform = $null; AssignmentId = $null; CreatedDateTime = $policy.createdDateTime; LastModifiedDateTime = $policy.lastModifiedDateTime
                            })
                    }
                }
            }
        }
        catch { Write-Warning "Error processing $($policyType.Name): $($_.Exception.Message)" }
        Write-Host "  ✓ $($policyType.DisplayName)" -ForegroundColor Green
    }
    $allPolicyAssignments.ToArray()
}

function Get-AllIntunePoliciesWithAssignments {
    param([switch]$DebugMode)
    $GraphEndpoint = "https://graph.microsoft.com"
    $policyTypes = @(
        @{ Name = "Device Configuration"; EntityType = "deviceConfigurations"; DisplayName = "Device Configuration Profiles" },
        @{ Name = "Settings Catalog"; EntityType = "configurationPolicies"; DisplayName = "Settings Catalog Policies" },
        @{ Name = "Administrative Templates"; EntityType = "groupPolicyConfigurations"; DisplayName = "Administrative Templates" },
        @{ Name = "Compliance Policies"; EntityType = "deviceCompliancePolicies"; DisplayName = "Compliance Policies" },
        @{ Name = "App Protection"; EntityType = "deviceAppManagement/managedAppPolicies"; DisplayName = "App Protection Policies" },
        @{ Name = "App Configuration"; EntityType = "mobileAppConfigurations"; DisplayName = "App Configuration Policies" },
        @{ Name = "Applications"; EntityType = "deviceAppManagement/mobileApps"; DisplayName = "Applications" },
        @{ Name = "Platform Scripts"; EntityType = "deviceManagementScripts"; DisplayName = "Platform Scripts" },
        @{ Name = "Remediation Scripts"; EntityType = "deviceHealthScripts"; DisplayName = "Proactive Remediation Scripts" },
        @{ Name = "Autopilot Profile"; EntityType = "windowsAutopilotDeploymentProfiles"; DisplayName = "Autopilot Profile" },
        @{ Name = "Enrollment Status Page"; EntityType = "deviceEnrollmentConfigurations"; DisplayName = "Enrollment Status Page" },
        @{ Name = "Endpoint Security"; EntityType = "deviceManagement/intents"; DisplayName = "Endpoint Security Policies" },
        @{ Name = "Cloud PC Provisioning"; EntityType = "virtualEndpoint/provisioningPolicies"; DisplayName = "Cloud PC Provisioning Policies" },
        @{ Name = "Cloud PC User Settings"; EntityType = "virtualEndpoint/userSettings"; DisplayName = "Cloud PC User Settings" }
    )
    Write-Host "Retrieving assignment filters and groups..." -ForegroundColor Yellow
    $rawFilters = Invoke-GraphRequestWithPaging -Uri "$GraphEndpoint/beta/deviceManagement/assignmentFilters" -Method "GET" -DebugMode:$DebugMode
    $script:AllFilters = @{}
    if ($rawFilters) { foreach ($f in $rawFilters) { $script:AllFilters[$f.id] = $f } }
    $rawGroups = Invoke-GraphRequestWithPaging -Uri "$GraphEndpoint/v1.0/groups?`$select=id,displayName,membershipRule,groupTypes" -Method "GET" -DebugMode:$DebugMode
    $script:AllGroups = @{}
    if ($rawGroups) { foreach ($g in $rawGroups) { $n = Normalize-EntraGroupForDisplay -Group $g; if ($n -and $n.id) { $script:AllGroups[$n.id] = $n } } }
    Write-Host "  ✓ Assignment filters and groups" -ForegroundColor Green
    Write-Host "Collecting policy assignments from $($policyTypes.Count) categories..." -ForegroundColor Yellow
    Get-AllPolicyAssignmentsBatch -PolicyTypes $policyTypes -DebugMode:$DebugMode
}
#endregion

#region Device context and filter evaluation
function Get-PolicyPlatformFromPolicy {
    param(
        [Parameter(Mandatory)]
        [object]$Policy,
        [Parameter(Mandatory)]
        [string]$EntityType
    )
    $odataKey = '@odata.type'
    $odataType = $Policy.PSObject.Properties[$odataKey].Value
    if (-not $odataType) {
        if ($Policy.PSObject.Properties['operatingSystem'].Value) {
            $os = $Policy.operatingSystem
            if ($os -is [array]) { $os = $os[0] }
            $osStr = [string]$os
            if ($osStr -match 'windows') { return 'Windows' }
            if ($osStr -match 'macOS') { return 'macOS' }
            if ($osStr -match 'mac') { return 'macOS' }
            if ($osStr -match 'ios') { return 'iOS' }
            if ($osStr -match 'ipad') { return 'iOS' }
            if ($osStr -match 'android') { return 'Android' }
        }
        if ($Policy.PSObject.Properties['platformType'].Value) {
            $pt = [string]$Policy.platformType
            if ($pt -match 'windows') { return 'Windows' }
            if ($pt -match 'macOS') { return 'macOS' }
            if ($pt -match 'mac') { return 'macOS' }
            if ($pt -match 'ios') { return 'iOS' }
            if ($pt -match 'iPados') { return 'iOS' }
            if ($pt -match 'android') { return 'Android' }
        }
        $platsRaw = $null
        if ($Policy.PSObject.Properties['platforms']) { $platsRaw = $Policy.PSObject.Properties['platforms'].Value }
        if ($null -eq $platsRaw -and $null -ne $Policy.platforms) { $platsRaw = $Policy.platforms }
        $platsStr = ''
        if ($null -ne $platsRaw) { $platsStr = [string]$platsRaw }
        if ($platsStr.Trim().Length -gt 0) {
            $plats = $platsRaw
            if ($plats -is [array]) { $plats = $plats -join ',' }
            $platStr = [string]$plats
            $platLower = $platStr.ToLowerInvariant()
            if ($platLower.IndexOf('windows') -ge 0) { return 'Windows' }
            if ($platLower.IndexOf('macos') -ge 0 -or $platLower.IndexOf('mac') -ge 0) { return 'macOS' }
            if ($platLower.IndexOf('ios') -ge 0 -or $platLower.IndexOf('ipados') -ge 0) { return 'iOS' }
            if ($platLower.IndexOf('android') -ge 0) { return 'Android' }
        }
        if ($EntityType -match 'windowsAutopilotDeploymentProfiles') { return 'Windows' }
        if ($EntityType -match 'groupPolicyConfigurations') { return 'Windows' }
        if ($EntityType -match 'virtualEndpoint') { return 'Windows' }
        return $null
    }
    $prefix = '#microsoft.graph.'
    $escapedPrefix = [regex]::Escape($prefix)
    $typeName = ($odataType -replace $escapedPrefix, '').ToLowerInvariant()
    if ($typeName -match '^windows') { return 'Windows' }
    if ($typeName -match 'cloudpc') { return 'Windows' }
    if ($typeName -match '^macos') { return 'macOS' }
    if ($typeName -match '^mac') { return 'macOS' }
    if ($typeName -match '^ios') { return 'iOS' }
    if ($typeName -match '^ipad') { return 'iOS' }
    if ($typeName -match '^android') { return 'Android' }
    return $null
}

function Get-NormalizedDevicePlatform {
    param([Parameter(Mandatory = $false)][string]$OperatingSystem)
    if ([string]::IsNullOrWhiteSpace($OperatingSystem)) { return $null }
    $os = $OperatingSystem.Trim().ToLowerInvariant()
    if ($os -match '^windows') { return 'Windows' }
    if ($os -match '^macos') { return 'macOS' }
    if ($os -match '^mac') { return 'macOS' }
    if ($os -match '^ios') { return 'iOS' }
    if ($os -match '^ipad') { return 'iOS' }
    if ($os -match '^android') { return 'Android' }
    return $null
}

# Infer policy platform from display name when API does not return platforms (e.g. intents, or $select omits it).
function Get-PolicyPlatformFromPolicyName {
    param([Parameter(Mandatory = $false)][string]$PolicyName)
    if ([string]::IsNullOrWhiteSpace($PolicyName)) { return $null }
    $n = $PolicyName.Trim().ToLowerInvariant()
    if ($n.StartsWith('macos ') -or $n.StartsWith('macos-') -or $n.StartsWith('macos:')) { return 'macOS' }
    if ($n.StartsWith('mac os')) { return 'macOS' }
    if ($n.StartsWith('ios ') -or $n.StartsWith('ios-') -or $n.StartsWith('ipad')) { return 'iOS' }
    if ($n -like '*- ios') { return 'iOS' }
    if ($n.StartsWith('android ') -or $n.StartsWith('android-') -or $n -like '*- android') { return 'Android' }
    if ($n.StartsWith('windows ') -or $n.Contains('windows 11') -or $n.Contains('windows 365') -or $n.Contains('cloud pc') -or $n -like '*- windows*') { return 'Windows' }
    return $null
}

# Returns $true if the assignment row's policy applies to the given device platform (for Mermaid/flow view). When devicePlatform is null, all applied rows are shown.
function Test-AssignmentMatchesDevicePlatform {
    param([Parameter(Mandatory)][object]$AssignmentRow, [Parameter(Mandatory = $false)][string]$DevicePlatform)
    if ([string]::IsNullOrWhiteSpace($DevicePlatform)) { return $true }
    $policyPlat = $AssignmentRow.PolicyPlatform
    if ($null -eq $policyPlat -or ([string]$policyPlat).Trim().Length -eq 0) {
        $policyPlat = Get-PolicyPlatformFromPolicyName -PolicyName $AssignmentRow.PolicyName
    }
    $policyPlatStr = if ($null -ne $policyPlat) { [string]$policyPlat } else { '' }
    if ($policyPlatStr.Trim().Length -eq 0) { return $true }
    $pNorm = $policyPlatStr.Trim().ToLowerInvariant()
    $dNorm = ([string]$DevicePlatform).Trim().ToLowerInvariant()
    $pIsWin = ($pNorm -eq 'windows' -or $pNorm -like 'windows*')
    $dIsWin = ($dNorm -eq 'windows' -or $dNorm -like 'windows*')
    $pIsMac = ($pNorm -eq 'macos' -or $pNorm -like 'mac*')
    $dIsMac = ($dNorm -eq 'macos' -or $dNorm -like 'mac*')
    $pIsIos = ($pNorm -eq 'ios' -or $pNorm -like 'ios*' -or $pNorm -like '*ipad*')
    $dIsIos = ($dNorm -eq 'ios' -or $dNorm -like 'ios*')
    $pIsAndroid = ($pNorm -eq 'android' -or $pNorm -like 'android*')
    $dIsAndroid = ($dNorm -eq 'android' -or $dNorm -like 'android*')
    $samePlatform = ($pIsWin -and $dIsWin) -or ($pIsMac -and $dIsMac) -or ($pIsIos -and $dIsIos) -or ($pIsAndroid -and $dIsAndroid)
    return $samePlatform
}

function ConvertTo-DevicePropertiesForFilter {
    param([Parameter(Mandatory)][object]$RawDevice)
    $azureAdDeviceId = $null
    if ($RawDevice.PSObject.Properties['azureAdDeviceId']) { $azureAdDeviceId = $RawDevice.azureAdDeviceId }
    if ($null -eq $azureAdDeviceId -and $RawDevice.PSObject.Properties['AzureAdDeviceId']) { $azureAdDeviceId = $RawDevice.AzureAdDeviceId }
    [PSCustomObject]@{
        Id = $RawDevice.id; DeviceName = $RawDevice.deviceName; UserPrincipalName = $RawDevice.userPrincipalName; AzureAdDeviceId = $azureAdDeviceId
        OperatingSystem = $RawDevice.operatingSystem; OSVersion = $RawDevice.osVersion; DeviceType = $RawDevice.deviceType; ComplianceState = $RawDevice.complianceState
        JoinType = $RawDevice.joinType; ManagementAgent = $RawDevice.managementAgent; OwnerType = $RawDevice.ownerType; EnrollmentProfileName = $RawDevice.enrollmentProfileName
        AutopilotEnrolled = $RawDevice.autopilotEnrolled; Manufacturer = $RawDevice.manufacturer; Model = $RawDevice.model; SerialNumber = $RawDevice.serialNumber
        ProcessorArchitecture = $RawDevice.processorArchitecture; EthernetMacAddress = $RawDevice.ethernetMacAddress; WiFiMacAddress = $RawDevice.wiFiMacAddress
        TotalStorageSpaceInBytes = $RawDevice.totalStorageSpaceInBytes; FreeStorageSpaceInBytes = $RawDevice.freeStorageSpaceInBytes; PhysicalMemoryInBytes = $RawDevice.physicalMemoryInBytes
        IsEncrypted = $RawDevice.isEncrypted; IsSupervised = $RawDevice.isSupervised; JailBroken = $RawDevice.jailBroken; AzureAdRegistered = $RawDevice.azureAdRegistered
        DeviceEnrollmentType = $RawDevice.deviceEnrollmentType; ChassisType = $RawDevice.chassisType; EnrolledDateTime = $RawDevice.enrolledDateTime; LastSyncDateTime = $RawDevice.lastSyncDateTime
        ManagementState = $RawDevice.managementState; DeviceRegistrationState = $RawDevice.deviceRegistrationState; DeviceCategory = $RawDevice.deviceCategory
        DeviceCategoryDisplayName = $RawDevice.deviceCategoryDisplayName; EmailAddress = $RawDevice.emailAddress; UserDisplayName = $RawDevice.userDisplayName; UserId = $RawDevice.userId
        DevicePhysicalIds = $RawDevice.devicePhysicalIds  # Added for debugging dynamic group membership
    }
}

function Get-DeviceEvaluationContext {
    param([Parameter(Mandatory)][string]$DeviceNameOrId, [switch]$DebugMode)
    $isGuid = $DeviceNameOrId -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
    $deviceId = $null
    if ($isGuid) {
        try {
            $response = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$DeviceNameOrId')" -Method GET -OutputType PSObject -ErrorAction Stop
            $deviceId = $response.id
        }
        catch {
            $is404 = $_.Exception.Response?.StatusCode -eq 404 -or ($_.ErrorDetails?.Message -and ($_.ErrorDetails.Message -match '404' -or $_.ErrorDetails.Message -match 'Not Found'))
            if (-not $is404 -and $DebugMode) { Write-Warning "Managed device lookup: $($_.Exception.Message)" }
            if ($is404 -or $null -eq $deviceId) {
                try {
                    $entraDevice = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/devices/$DeviceNameOrId" -Method GET -OutputType PSObject -ErrorAction Stop
                    $azureDeviceId = $entraDevice.deviceId
                    if ($azureDeviceId) {
                        $matching = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=azureAdDeviceId eq '$azureDeviceId'&`$select=id" -Method GET -OutputType PSObject -ErrorAction Stop
                        if ($matching.value -and $matching.value.Count -gt 0) { $deviceId = $matching.value[0].id }
                    }
                }
                catch {
                    if ($DebugMode) { Write-Warning "Entra device lookup: $($_.Exception.Message)" }
                }
            }
            if (-not $deviceId) { throw }
        }
    }
    else {
        $escaped = $DeviceNameOrId -replace "'", "''"
        $deviceUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=deviceName eq '$escaped'&`$select=id"
        $response = Invoke-MgGraphRequest -Uri $deviceUri -Method GET -OutputType PSObject -ErrorAction Stop
        if ($response.PSObject.Properties['value'] -and $response.value -and $response.value.Count -gt 0) {
            $deviceId = $response.value[0].id
        }
    }
    if (-not $deviceId) { return $null }
    $props = $script:ManagedDeviceSelectProperties -join ','
    $rawDevice = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$deviceId`?`$select=$props" -OutputType PSObject
    $deviceProperties = ConvertTo-DevicePropertiesForFilter -RawDevice $rawDevice
    $userGroupIds = @(); $userDirectGroupIds = @(); $deviceGroupIds = @(); $deviceDirectGroupIds = @(); $entraDeviceObjectId = $null
    if ($deviceProperties.UserPrincipalName) {
        try {
            $userUri = "https://graph.microsoft.com/v1.0/users/$($deviceProperties.UserPrincipalName)"
            $userResponse = Invoke-MgGraphRequest -Method GET -Uri $userUri -ErrorAction Stop
            if ($userResponse) {
                $userGroupsData = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/v1.0/users/$($userResponse.id)/transitiveMemberOf?`$select=id" -Method GET
                $userGroupIds = @($userGroupsData | Select-Object -ExpandProperty id)
                $userDirectData = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/v1.0/users/$($userResponse.id)/memberOf?`$select=id" -Method GET
                $userDirectGroupIds = @($userDirectData | Select-Object -ExpandProperty id)
            }
        }
        catch { Write-Warning "User group resolution failed: $($_.Exception.Message)" }
    }
    if ($deviceProperties.AzureAdDeviceId) {
        try {
            $azureResp = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$($deviceProperties.AzureAdDeviceId)'&`$select=id,physicalIds" -Method GET
            if ($azureResp -and $azureResp.Count -gt 0) {
                $entraDeviceObjectId = $azureResp[0].id
                # Store physicalIds for debugging dynamic group membership (even if empty)
                if ($null -ne $azureResp[0].physicalIds) {
                    $deviceProperties | Add-Member -MemberType NoteProperty -Name 'PhysicalIds' -Value $azureResp[0].physicalIds -Force
                }
                $deviceGroupsData = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/v1.0/devices/$($azureResp[0].id)/transitiveMemberOf?`$select=id" -Method GET
                $deviceGroupIds = @($deviceGroupsData | Select-Object -ExpandProperty id)
                $deviceDirectData = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/v1.0/devices/$($azureResp[0].id)/memberOf?`$select=id" -Method GET
                $deviceDirectGroupIds = @($deviceDirectData | Select-Object -ExpandProperty id)
            }
        }
        catch { Write-Warning "Device group resolution failed: $($_.Exception.Message)" }
    }
    # Fallback: when AzureAdDeviceId is missing, try Entra object ID (AzureActiveDirectoryDeviceId) to resolve device group membership
    if (($deviceGroupIds.Count -eq 0) -and $rawDevice) {
        $entraObjId = $null
        if ($rawDevice.PSObject.Properties['azureActiveDirectoryDeviceId']) { $entraObjId = $rawDevice.azureActiveDirectoryDeviceId }
        elseif ($rawDevice.PSObject.Properties['AzureActiveDirectoryDeviceId']) { $entraObjId = $rawDevice.AzureActiveDirectoryDeviceId }
        if ($entraObjId -and [string]$entraObjId -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
            try {
                # entraObjId here is azureActiveDirectoryDeviceId (= Entra deviceId), NOT the directory object ID.
                # We must first resolve the Entra directory object ID, then use that for transitiveMemberOf.
                $fallbackResp = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$entraObjId'&`$select=id" -Method GET
                if ($fallbackResp -and $fallbackResp.Count -gt 0) {
                    $resolvedObjectId = $fallbackResp[0].id
                    $deviceGroupsData = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/v1.0/devices/$resolvedObjectId/transitiveMemberOf?`$select=id" -Method GET
                    $deviceGroupIds = @($deviceGroupsData | Select-Object -ExpandProperty id)
                    $deviceDirectData = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/v1.0/devices/$resolvedObjectId/memberOf?`$select=id" -Method GET
                    $deviceDirectGroupIds = @($deviceDirectData | Select-Object -ExpandProperty id)
                    if (-not $entraDeviceObjectId) { $entraDeviceObjectId = $resolvedObjectId }
                }
            }
            catch { Write-Warning "Device group resolution (fallback): $($_.Exception.Message)" }
        }
    }
    # Diagnostic: always show group membership counts so failures are visible
    Write-Host "  Device group IDs found: $($deviceGroupIds.Count)" -ForegroundColor $(if ($deviceGroupIds.Count -eq 0) { 'DarkYellow' } else { 'Green' })
    Write-Host "  User group IDs found: $($userGroupIds.Count)" -ForegroundColor $(if ($userGroupIds.Count -eq 0 -and $deviceProperties.UserPrincipalName) { 'DarkYellow' } else { 'Green' })
    if ($deviceGroupIds.Count -eq 0) { Write-Warning "No device group memberships found. Group-based assignments will not be evaluated for this device." }
    [PSCustomObject]@{
        ManagedDeviceId = $deviceId; EntraDeviceObjectId = $entraDeviceObjectId; DeviceProperties = $deviceProperties; UserGroupIds = $userGroupIds; UserDirectGroupIds = $userDirectGroupIds; DeviceGroupIds = $deviceGroupIds; DeviceDirectGroupIds = $deviceDirectGroupIds; RawDevice = $rawDevice
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

function ConvertTo-CanonicalAssignment {
    param([Parameter(Mandatory)][PSCustomObject]$AssignmentRow)
    # Map assignment type strings to Graph @odata.type. Unknown types default to All Devices.
    $odataType = "#microsoft.graph.allDevicesAssignmentTarget"
    if ($AssignmentRow.AssignmentType -eq "All Users" -or $AssignmentRow.AssignmentType -eq "All Users Assignment") { $odataType = "#microsoft.graph.allLicensedUsersAssignmentTarget" }
    elseif ($AssignmentRow.AssignmentType -eq "All Devices" -or $AssignmentRow.AssignmentType -eq "All Devices Assignment") { $odataType = "#microsoft.graph.allDevicesAssignmentTarget" }
    elseif ($AssignmentRow.AssignmentType -eq "Group (Include)" -or $AssignmentRow.AssignmentType -eq "Group") { $odataType = "#microsoft.graph.groupAssignmentTarget" }
    elseif ($AssignmentRow.AssignmentType -eq "Group (Exclude)" -or $AssignmentRow.AssignmentType -eq "Exclude Group") { $odataType = "#microsoft.graph.exclusionGroupAssignmentTarget" }
    $filterIdVal = $null
    if ($AssignmentRow.FilterId -and $AssignmentRow.FilterId -ne "00000000-0000-0000-0000-000000000000") { $filterIdVal = $AssignmentRow.FilterId }
    $filterTypeVal = $null
    if ($AssignmentRow.FilterType -and $AssignmentRow.FilterType -ne "none") { $filterTypeVal = $AssignmentRow.FilterType.ToLower() }
    [PSCustomObject]@{ TargetType = $odataType; GroupId = $AssignmentRow.GroupId; FilterId = $filterIdVal; FilterType = $filterTypeVal }
}

function Test-AssignmentAppliesToDevice {
    param([Parameter(Mandatory)][PSCustomObject]$CanonicalAssignment, [Parameter(Mandatory)][PSCustomObject]$DeviceContext, [hashtable]$Filters)
    $dp = $DeviceContext.DeviceProperties
    $userGroupIds = if ($DeviceContext.UserGroupIds) { @($DeviceContext.UserGroupIds) } else { @() }
    $deviceGroupIds = if ($DeviceContext.DeviceGroupIds) { @($DeviceContext.DeviceGroupIds) } else { @() }
    $targetType = $CanonicalAssignment.TargetType; $filterId = $CanonicalAssignment.FilterId; $filterType = $CanonicalAssignment.FilterType
    $baseApplicable = $false
    if ($targetType -eq "#microsoft.graph.allDevicesAssignmentTarget") {
        $baseApplicable = $true
    }
    elseif ($targetType -eq "#microsoft.graph.allLicensedUsersAssignmentTarget") {
        $hasDeviceFilter = $filterId -and $filterId -ne "00000000-0000-0000-0000-000000000000"
        $baseApplicable = ($null -ne $dp.UserPrincipalName) -or $hasDeviceFilter
    }
    elseif ($targetType -eq "#microsoft.graph.groupAssignmentTarget") {
        $groupId = $CanonicalAssignment.GroupId
        $groupIdStr = [string]$groupId
        if ($groupIdStr.Trim().Length -gt 0) {
            $baseApplicable = @($deviceGroupIds | Where-Object { [string]$_ -eq $groupIdStr }).Count -gt 0 -or @($userGroupIds | Where-Object { [string]$_ -eq $groupIdStr }).Count -gt 0
        }
    }
    elseif ($targetType -eq "#microsoft.graph.exclusionGroupAssignmentTarget") {
        $groupId = $CanonicalAssignment.GroupId
        $groupIdStr = [string]$groupId
        if ($groupIdStr.Trim().Length -gt 0) {
            $baseApplicable = @($deviceGroupIds | Where-Object { [string]$_ -eq $groupIdStr }).Count -gt 0 -or @($userGroupIds | Where-Object { [string]$_ -eq $groupIdStr }).Count -gt 0
        }
    }
    $filterResult = 'N/A'; $applies = $baseApplicable
    $hasRealFilter = $filterId -and $filterId -ne "00000000-0000-0000-0000-000000000000" -and $filterType -and $filterType -ne "none"
    if ($hasRealFilter -and $baseApplicable) {
        $filterObj = $Filters[$filterId]
        if ($filterObj) {
            $filterMatched = Test-IntuneFilter -FilterRule $filterObj.rule -DeviceProperties $dp
            if ($filterMatched) { $filterResult = 'Matched' } else { $filterResult = 'NotMatched' }
            if ($filterType -eq "include") { $applies = $filterMatched } elseif ($filterType -eq "exclude") { $applies = -not $filterMatched }
        }
        else { $filterResult = 'NotFound'; $applies = $false }
    }
    if ($targetType -eq "#microsoft.graph.exclusionGroupAssignmentTarget") { $applies = $false }
    [PSCustomObject]@{ AppliesToDevice = $applies; FilterResult = $filterResult }
}

function Invoke-EvaluateAssignmentsForDevice {
    param([Parameter(Mandatory)][array]$PolicyAssignments, [Parameter(Mandatory)][PSCustomObject]$DeviceContext, [switch]$ApplyPlatformFilter, [switch]$DebugMode)
    $filters = $script:AllFilters; if (-not $filters) { $filters = @{} }
    $deviceOs = $DeviceContext.DeviceProperties.OperatingSystem
    $devicePlatform = Get-NormalizedDevicePlatform -OperatingSystem $deviceOs
    $out = [System.Collections.ArrayList]::new()
    foreach ($row in $PolicyAssignments) {
        $evalResult = $null
        if ($row.AssignmentType -eq "Not Assigned") {
            $evalResult = [PSCustomObject]@{ AppliesToDevice = $false; FilterResult = 'N/A' }
        }
        else {
            $canon = ConvertTo-CanonicalAssignment -AssignmentRow $row
            $evalResult = Test-AssignmentAppliesToDevice -CanonicalAssignment $canon -DeviceContext $DeviceContext -Filters $filters
        }
        $applies = $evalResult.AppliesToDevice
        $filterResult = $evalResult.FilterResult
        # Optional: Intune scopes "All Devices" / "All Users" by policy platform. Only exclude when -ApplyPlatformFilter is set (default: off to match legacy behavior).
        if ($ApplyPlatformFilter -and $applies -and $devicePlatform) {
            $policyPlat = $row.PolicyPlatform
            $policyPlatStr = ''
            if ($null -ne $policyPlat) { $policyPlatStr = [string]$policyPlat }
            if ($policyPlatStr.Trim().Length -eq 0) {
                $policyPlat = Get-PolicyPlatformFromPolicyName -PolicyName $row.PolicyName
            }
            $policyPlatStr = ''
            if ($null -ne $policyPlat) { $policyPlatStr = [string]$policyPlat }
            if ($policyPlatStr.Trim().Length -gt 0) {
                $pNorm = $policyPlatStr.Trim().ToLowerInvariant()
                $dNorm = ([string]$devicePlatform).Trim().ToLowerInvariant()
                $pIsWin = ($pNorm -eq 'windows' -or $pNorm -like 'windows*')
                $dIsWin = ($dNorm -eq 'windows' -or $dNorm -like 'windows*')
                $pIsMac = ($pNorm -eq 'macos' -or $pNorm -like 'mac*')
                $dIsMac = ($dNorm -eq 'macos' -or $dNorm -like 'mac*')
                $pIsIos = ($pNorm -eq 'ios' -or $pNorm -like 'ios*' -or $pNorm -like '*ipad*')
                $dIsIos = ($dNorm -eq 'ios' -or $dNorm -like 'ios*')
                $pIsAndroid = ($pNorm -eq 'android' -or $pNorm -like 'android*')
                $dIsAndroid = ($dNorm -eq 'android' -or $dNorm -like 'android*')
                $samePlatform = ($pIsWin -and $dIsWin) -or ($pIsMac -and $dIsMac) -or ($pIsIos -and $dIsIos) -or ($pIsAndroid -and $dIsAndroid)
                if (-not $samePlatform) {
                    $applies = $false
                    $filterResult = 'PlatformMismatch'
                }
            }
        }
        $ht = @{}; $row.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }; $ht['FilterResult'] = $filterResult; $ht['AppliesToDevice'] = $applies
        $null = $out.Add([PSCustomObject]$ht)
    }
    $out.ToArray()
}
#endregion

#region HTML Reports
function Get-FriendlyPlatformDisplayName {
    param([string]$Platform)
    if ([string]::IsNullOrWhiteSpace($Platform)) { return "Unknown" }
    $p = $Platform.Trim().ToLowerInvariant()
    $map = @{
        'windows10andlater'                  = 'Windows'
        'windows81andlater'                  = 'Windows'
        'windowsphone81'                     = 'Windows'
        'macos'                              = 'macOS'
        'ios'                                = 'iOS/iPadOS'
        'iosmobileapplicationmanagement'     = 'iOS/iPadOS'
        'android'                            = 'Android'
        'androidforwork'                     = 'Android'
        'androidworkprofile'                 = 'Android'
        'androidaosp'                        = 'Android'
        'androidmobileapplicationmanagement' = 'Android'
    }
    $key = $p -replace '\s', ''
    if ($map.ContainsKey($key)) { return $map[$key] }
    return $Platform
}

function Get-AssignmentOverviewTabFragment {
    param([Parameter(Mandatory)][array]$PolicyAssignments, [string]$TenantName = "Intune Tenant")
    $totalAssignments = if ($PolicyAssignments) { $PolicyAssignments.Count } else { 0 }
    # Count all unique policies, regardless of assignment status
    $totalPolicies = if ($PolicyAssignments) { ($PolicyAssignments | Select-Object -Property PolicyName -Unique).Count } else { 0 }
    $unassignedPolicies = if ($PolicyAssignments) { ($PolicyAssignments | Where-Object { $_.AssignmentType -eq "Not Assigned" } | Select-Object -Property PolicyName -Unique).Count } else { 0 }
    $filtersUsed = $PolicyAssignments | Where-Object { $_.FilterName -and $_.FilterName -ne "No Filter" -and $_.FilterName -ne "Filter Not Found" }
    $filterStats = if ($filtersUsed) { $filtersUsed | Group-Object -Property FilterName | Sort-Object Count -Descending } else { @() }
    $hasCategory = ($PolicyAssignments | Select-Object -First 1).PSObject.Properties.Name -contains 'PolicyCategory'
    $tableRows = ""
    $seenKey = @{}
    if ($PolicyAssignments -and $PolicyAssignments.Count -gt 0) {
        foreach ($a in $PolicyAssignments) {
            $targetName = if ($a.AssignmentType -match "^(All Devices|All Users)$") { $a.AssignmentType } else { [System.Net.WebUtility]::HtmlEncode($a.TargetName) }
            if ([string]::IsNullOrWhiteSpace($targetName) -and $a.GroupId -and $script:AllGroups.Count -gt 0) {
                $grp = $script:AllGroups[$a.GroupId]
                if ($grp -and $grp.displayName) { $targetName = [System.Net.WebUtility]::HtmlEncode($grp.displayName) }
            }
            if ([string]::IsNullOrWhiteSpace($targetName)) { $targetName = [System.Net.WebUtility]::HtmlEncode([string]$a.TargetName) }
            if ([string]::IsNullOrWhiteSpace($targetName) -and $a.AssignmentType -eq "Not Assigned") { $targetName = "Not Assigned" }
            # Show filter name with type: "Filter Name (Include)" or "Filter Name (Exclude)"
            $filterDisplay = "No Filter"
            if ($a.FilterName -and $a.FilterName -ne "No Filter" -and $a.FilterName -ne "Filter Not Found") {
                $filterNameEnc = [System.Net.WebUtility]::HtmlEncode($a.FilterName)
                if ($a.FilterType -and $a.FilterType -ne "None" -and $a.FilterType -ne "none") {
                    $ft = $a.FilterType.ToString()
                    $filterTypeLabel = "(" + $ft.Substring(0, 1).ToUpper() + $ft.Substring(1) + ")"
                    $filterDisplay = "$filterNameEnc $filterTypeLabel"
                } else {
                    $filterDisplay = $filterNameEnc
                }
            }
            $pc = if ($hasCategory -and $a.PolicyCategory) { [System.Net.WebUtility]::HtmlEncode($a.PolicyCategory) } else { "" }
            $dedupeKey = "$($a.PolicyName)|$($a.PolicyCategory)|$($a.AssignmentType)|$($a.TargetName)|$($a.FilterName)"
            if ($seenKey[$dedupeKey]) { continue }
            $seenKey[$dedupeKey] = $true
            $tableRows += " <tr><td>$([System.Net.WebUtility]::HtmlEncode($a.PolicyName))</td><td>$pc</td><td>$targetName</td><td>$filterDisplay</td></tr>`n"
        }
    }
    if ([string]::IsNullOrWhiteSpace($tableRows)) {
        $tableRows = " <tr><td colspan=`"4`" class=`"text-muted text-center py-4`"><i class=`"fas fa-inbox me-2`"></i>No assignments</td></tr>`n"
    }
    $filterTableRows = ""
    foreach ($f in $filterStats) {
        $firstWithFilter = $PolicyAssignments | Where-Object { $_.FilterName -eq $f.Name } | Select-Object -First 1
        $filterRule = if ($firstWithFilter -and $firstWithFilter.FilterRule) { [System.Net.WebUtility]::HtmlEncode($firstWithFilter.FilterRule) } else { "Not Available" }
        $filterPlatform = if ($firstWithFilter -and $firstWithFilter.FilterPlatform) { [System.Net.WebUtility]::HtmlEncode((Get-FriendlyPlatformDisplayName -Platform $firstWithFilter.FilterPlatform)) } else { "Unknown" }
        $filterTableRows += " <tr><td><strong>$([System.Net.WebUtility]::HtmlEncode($f.Name))</strong></td><td><code style=`"font-size: 0.85em; word-break: break-all;`">$filterRule</code></td><td><span class=`"badge bg-secondary`">$filterPlatform</span></td><td><span class=`"badge bg-primary`">$($f.Count)</span></td></tr>`n"
    }
    $filterTableBody = "<tbody>" + $filterTableRows + "</tbody>"
    $filtersSectionHtml = if ($filterStats.Count -gt 0) {
        "<div class=`"row`"><div class=`"col-12`"><div class=`"modern-table-container`"><div class=`"modern-table-header`"><div style=`"display: flex; justify-content: space-between; align-items: center; width: 100%;`"><div><h5><i class=`"fas fa-filter me-2`"></i>Intune Assignment Filters</h5><small>Filter usage statistics with actual filtering syntax</small></div><div class=`"form-check form-switch`"><input class=`"form-check-input`" type=`"checkbox`" role=`"switch`" id=`"showAllFilters`"><label class=`"form-check-label`" for=`"showAllFilters`">Show All Results</label></div></div></div><div class=`"modern-table-body`"><div class=`"table-responsive`"><table class=`"table modern-table table-hover`" id=`"filtersTable`"><thead><tr><th>Filter Name</th><th>Filter Rule/Syntax</th><th>Platform</th><th>Usage Count</th></tr></thead>" + $filterTableBody + "</table></div></div></div></div></div>"
    }
    else {
        "<div class=`"row`"><div class=`"col-12`"><div class=`"modern-table-container`"><div class=`"modern-table-header`"><h5><i class=`"fas fa-info-circle me-2`"></i>Intune Assignment Filters</h5><small>No assignment filters are currently in use</small></div><div class=`"modern-table-body`"><div class=`"alert alert-info m-3`" role=`"alert`"><i class=`"fas fa-info-circle me-2`"></i>No assignment filters are currently in use in your tenant.</div></div></div></div></div>"
    }
    $overviewTemplate = @'
                    <div class="overview-container">
                        <div class="overview-header">
                            <h2><i class="fas fa-chart-pie me-3"></i>Assignment Overview</h2>
                            <p>Complete view of all policy assignments across your tenant</p>
                        </div>
                        <div class="row g-4 mb-5 overview-tiles-row">
                            <div class="col-md-6">
                                <div class="summary-card border-primary text-center">
                                    <div class="card-icon"><i class="fas fa-file-alt"></i></div>
                                    <h3 class="card-title">__PH_TP__</h3>
                                    <p class="card-text">Total Policies, Apps and (remediation) Scripts</p>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="summary-card border-danger text-center">
                                    <div class="card-icon"><i class="fas fa-exclamation-triangle"></i></div>
                                    <h3 class="card-title">__PH_UP__</h3>
                                    <p class="card-text">Unassigned Policies</p>
                                </div>
                            </div>
                        </div>
                        <div class="row mb-5">
                            <div class="col-12">
                                <div class="modern-table-container">
                                    <div class="modern-table-header">
                                        <div class="d-flex justify-content-between align-items-center">
                                            <div>
                                                <h5><i class="fas fa-table me-2"></i>All Policy Assignments Overview</h5>
                                                <small>Complete view of all policy assignments across your tenant</small>
                                            </div>
                                            <div class="form-check form-switch">
                                                <input class="form-check-input" type="checkbox" id="showAllAssignments">
                                                <label class="form-check-label" for="showAllAssignments"><i class="fas fa-list me-1"></i>Show All Results</label>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="assignment-filters assignment-filters-modern" id="overviewFiltersBar">
                                        <div class="assignment-filters-inner">
                                            <div class="filter-group dropdown">
                                                <label class="filter-label"><i class="fas fa-tag fa-fw"></i>Category</label>
                                                <button class="btn btn-sm btn-outline-secondary dropdown-toggle filter-dropdown-btn" type="button" id="overviewFilterCategoryBtn" data-bs-toggle="dropdown" data-bs-auto-close="outside" aria-expanded="false">Select...</button>
                                                <div class="dropdown-menu filter-checkbox-dropdown" id="overviewFilterCategoryMenu"></div>
                                            </div>
                                            <div class="filter-group dropdown">
                                                <label class="filter-label"><i class="fas fa-bullseye fa-fw"></i>Target</label>
                                                <button class="btn btn-sm btn-outline-secondary dropdown-toggle filter-dropdown-btn" type="button" id="overviewFilterTargetBtn" data-bs-toggle="dropdown" data-bs-auto-close="outside" aria-expanded="false">Select...</button>
                                                <div class="dropdown-menu filter-checkbox-dropdown" id="overviewFilterTargetMenu"></div>
                                            </div>
                                            <div class="filter-group dropdown">
                                                <label class="filter-label"><i class="fas fa-filter fa-fw"></i>Filter</label>
                                                <button class="btn btn-sm btn-outline-secondary dropdown-toggle filter-dropdown-btn" type="button" id="overviewFilterFilterBtn" data-bs-toggle="dropdown" data-bs-auto-close="outside" aria-expanded="false">Select...</button>
                                                <div class="dropdown-menu filter-checkbox-dropdown" id="overviewFilterFilterMenu"></div>
                                            </div>
                                            <div class="filter-group"><label class="filter-label"><i class="fas fa-eye-slash fa-fw"></i>Not assigned</label><select class="filter-select" id="overviewFilterHideNotAssigned"><option value="">Show</option><option value="hide">Hide</option></select></div>
                                            <button type="button" class="filter-reset-btn" id="overviewFiltersReset" title="Reset filters"><i class="fas fa-times-circle"></i> Reset</button>
                                        </div>
                                    </div>
                                    <div class="modern-table-body">
                                        <div class="table-responsive">
                                            <table class="table modern-table table-hover text-start" id="allAssignmentsTable">
                                                <thead><tr><th class="text-start">Policy Name</th><th class="text-start">Category</th><th class="text-start">Target</th><th class="text-start">Filter</th></tr></thead>
                                                <tbody>
__PH_TABLEROWS__
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
__PH_FILTERS__
                    </div>
'@
    $overviewTemplate.Replace('__PH_TP__', [string]$totalPolicies).Replace('__PH_UP__', [string]$unassignedPolicies).Replace('__PH_TABLEROWS__', $tableRows).Replace('__PH_FILTERS__', $filtersSectionHtml)
}

function New-AssignmentOverviewHtmlReport {
    param([Parameter(Mandatory)][array]$PolicyAssignments, [string]$TenantName = "Intune Tenant", [Parameter(Mandatory)][string]$OutputPath)
    $fragment = Get-AssignmentOverviewTabFragment -PolicyAssignments $PolicyAssignments -TenantName $TenantName
    # Script in literal here-string so $ in JavaScript are not expanded by PowerShell (same approach as device report).
    $assignmentOverviewScript = @'
document.addEventListener('DOMContentLoaded', function() {
  var themeToggle = document.getElementById('themeToggle');
  var prefersDark = window.matchMedia('(prefers-color-scheme: dark)');
  function applyTheme(isDark) {
    document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light');
    if (themeToggle) themeToggle.checked = isDark;
  }
  var saved = localStorage.getItem('theme');
  if (saved === 'dark' || saved === 'light') {
    applyTheme(saved === 'dark');
  } else {
    applyTheme(prefersDark.matches);
  }
  if (themeToggle) themeToggle.addEventListener('change', function() {
    var isDark = this.checked;
    document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light');
    localStorage.setItem('theme', isDark ? 'dark' : 'light');
  });
  prefersDark.addEventListener('change', function(e) {
    if (localStorage.getItem('theme') === null) {
      applyTheme(e.matches);
    }
  });
  if (jQuery && jQuery.fn.DataTable && jQuery('#allAssignmentsTable').length > 0) {
    var $allTbl = jQuery('#allAssignmentsTable');
    var preCategories = [], preTargets = [], preFilters = [];
    function normFilter(s) { return String(s||'').replace(/\s*\((?:Include|Exclude)$/i, '').trim() || String(s||''); }
    $allTbl.find('tbody tr').each(function() {
      var $row = jQuery(this);
      var $cells = $row.find('td');
      if ($cells.length >= 4) {
        var c = jQuery.trim($cells.eq(1).text()); if (c && preCategories.indexOf(c) === -1) preCategories.push(c);
        var t = jQuery.trim($cells.eq(2).text()); if (t && preTargets.indexOf(t) === -1) preTargets.push(t);
        var f = normFilter(jQuery.trim($cells.eq(3).text())); if (f && preFilters.indexOf(f) === -1) preFilters.push(f);
      }
    });
    preCategories.sort(); preTargets.sort(); preFilters.sort();
    var at = $allTbl.DataTable({
      responsive: true, pageLength: 25, order: [[1,'asc'],[0,'asc']], dom: 'Bfrtip', buttons: ['copy','csv','excel','pdf','print'],
            columnDefs: [
                { targets: [0,1,2,3], className: 'text-start' },
                { targets: 0, width: '25%' }, { targets: 1, width: '15%' }, { targets: 2, width: '15%' }, { targets: 3, width: '20%' }
            ],
      initComplete: function() {
        var api = this.api();
        function escOv(v){ return (v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
        function getOvChecked(menuId){ return jQuery('#'+menuId+' input.filter-cb:checked').map(function(){ return jQuery(this).attr('data-value'); }).get(); }
        function updateOvBtn(btnId, menuId, label){ var n = getOvChecked(menuId).length; jQuery('#'+btnId).text(n ? label + ' (' + n + ')' : 'Select...'); }
        function normalizeFilterName(v) { var s = String(v||''); return s.replace(/\s*\((?:Include|Exclude)\)$/i, '').trim() || s; }
        function fillOvDropdownFromArray(menuId, btnId, values, label) {
          var menu = jQuery('#'+menuId); if (!menu.length) return; menu.empty();
          jQuery.each(values, function(i,v){ menu.append('<label class="dropdown-item"><input type="checkbox" class="filter-cb" data-value="'+escOv(v)+'"> '+escOv(v)+'</label>'); });
          menu.find('input.filter-cb').on('change', function(e){ e.stopPropagation(); updateOvBtn(btnId, menuId, label); api.draw(); });
          menu.on('click', function(e){ e.stopPropagation(); });
          updateOvBtn(btnId, menuId, label);
        }
        fillOvDropdownFromArray('overviewFilterCategoryMenu','overviewFilterCategoryBtn', preCategories, 'Category');
        fillOvDropdownFromArray('overviewFilterTargetMenu','overviewFilterTargetBtn', preTargets, 'Target');
        fillOvDropdownFromArray('overviewFilterFilterMenu','overviewFilterFilterBtn', preFilters, 'Filter');
        var overviewSearchFn = function(settings, data, dataIndex) {
          if (settings.nTable && settings.nTable.id !== 'allAssignmentsTable') return true;
          if (jQuery('#overviewFilterHideNotAssigned').length && jQuery('#overviewFilterHideNotAssigned').val() === 'hide' && data[2] === 'Not Assigned') return false;
          var searchStr = ''; try { var searchApi = new jQuery.fn.dataTable.Api(settings); searchStr = (searchApi.search() || '').trim(); } catch (e) {}
          if (searchStr) { var found = false; for (var i = 0; i < data.length; i++) { if (data[i] && data[i].toString().toLowerCase().indexOf(searchStr.toLowerCase()) !== -1) { found = true; break; } } if (!found) return false; }
          var c = getOvChecked('overviewFilterCategoryMenu'); if (c.length && jQuery.inArray(data[1], c) === -1) return false;
          var t = getOvChecked('overviewFilterTargetMenu'); if (t.length && jQuery.inArray(data[2], t) === -1) return false;
          var rowFilterNorm = normalizeFilterName(data[3]); var f = getOvChecked('overviewFilterFilterMenu'); if (f.length && jQuery.inArray(rowFilterNorm, f) === -1) return false;
          return true;
        };
        jQuery.fn.dataTable.ext.search.push(overviewSearchFn);
        var hideNotAssigned = jQuery('#overviewFilterHideNotAssigned');
        if (hideNotAssigned.length) hideNotAssigned.on('change', function(){ api.draw(); });
        var resetBtn = jQuery('#overviewFiltersReset');
        if (resetBtn.length) resetBtn.on('click', function(){
          jQuery('#overviewFilterCategoryMenu,#overviewFilterTargetMenu,#overviewFilterFilterMenu').find('input.filter-cb').prop('checked', false);
          updateOvBtn('overviewFilterCategoryBtn','overviewFilterCategoryMenu','Category'); updateOvBtn('overviewFilterTargetBtn','overviewFilterTargetMenu','Target'); updateOvBtn('overviewFilterFilterBtn','overviewFilterFilterMenu','Filter');
          hideNotAssigned.val('');
          api.draw();
        });
      }
    });
    jQuery('#showAllAssignments').prop('checked', false);
    jQuery('#showAllAssignments').on('change', function() {
      var dt = jQuery('#allAssignmentsTable').DataTable();
      var paginateControls = jQuery('#allAssignmentsTable').closest('.dataTables_wrapper').find('.dataTables_paginate, .dataTables_info');
      if (this.checked) { dt.page.len(-1); paginateControls.hide(); } else { dt.page.len(25); paginateControls.show(); }
      dt.draw();
    });
  }
  if (jQuery && jQuery.fn.DataTable && jQuery('#filtersTable').length > 0) {
    var filtersTable = jQuery('#filtersTable').DataTable({
      responsive: true, pageLength: 10, order: [[3, 'desc']],
      dom: 'Bfrtip', buttons: ['copy', 'csv', 'excel', 'pdf', 'print'],
            columnDefs: [
                { targets: [0,1,2,3], className: 'text-start' },
                { targets: 0, width: '20%' }, { targets: 1, width: '40%' }, { targets: 2, width: '15%' },
                { targets: 3, width: '25%' }
            ]
    });
    jQuery('#showAllFilters').prop('checked', false);
    jQuery('#showAllFilters').on('change', function() {
      var isShowAll = this.checked;
      var paginateControls = jQuery('#filtersTable').closest('.dataTables_wrapper').find('.dataTables_paginate, .dataTables_info');
      if (isShowAll) { filtersTable.page.len(-1); paginateControls.hide(); }
      else { filtersTable.page.len(10); paginateControls.show(); }
      filtersTable.draw();
    });
  }
});
'@
    $html = @"
<!DOCTYPE html>
<html lang=`"en`">
<head>
    <meta charset=`"UTF-8`">
    <meta name=`"viewport`" content=`"width=device-width, initial-scale=1.0`">
    <title>$TenantName Intune Enrollment Flow Visualization</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/5.3.0/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.4.1/css/buttons.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://code.jquery.com/jquery-3.7.0.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/dataTables.buttons.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.bootstrap5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.html5.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <style>
        :root { --primary-color: #0078d4; --secondary-color: #2b88d8; --accent-color: #107c10; --bg-color: #e2e8f0; --card-bg: #eef2f7; --text-color: #334155; --text-secondary: #64748b; --border-color: #cbd5e1; --shadow-sm: 0 1px 2px 0 rgb(0 0 0 / 0.05); --shadow-lg: 0 10px 15px -3px rgb(0 0 0 / 0.1), 0 4px 6px -4px rgb(0 0 0 / 0.1); --bg-secondary: #e8ecf1; --bg-tertiary: #e2e8f0; --table-hover-bg: rgba(0,0,0,0.04); }
        [data-theme="dark"] { --bg-color: #0d0d0d; --card-bg: #1a1a1a; --text-color: #e5e5e5; --text-secondary: #a3a3a3; --border-color: #404040; --bg-secondary: #1a1a1a; --bg-tertiary: #262626; --table-hover-bg: rgba(255,255,255,0.05); }
        body { background: var(--bg-color); color: var(--text-color); min-height: 100vh; font-family: Segoe UI, Roboto, sans-serif; }
        .app-container { min-height: 100vh; padding: 0; display: flex; flex-direction: column; }
        .container { max-width: 1200px; margin: 0 auto; padding: 0 20px; background: var(--card-bg); flex: 1; }
        .theme-toggle { position: absolute; top: 10px; right: 10px; z-index: 1050; display: flex; align-items: center; gap: 6px; background-color: var(--card-bg); padding: 6px 10px; border-radius: 30px; box-shadow: var(--shadow-lg); border: 1px solid var(--border-color); }
        .theme-toggle-switch { position: relative; display: inline-block; width: 40px; height: 20px; }
        .theme-toggle-switch input { opacity: 0; width: 0; height: 0; }
        .theme-toggle-slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #cbd5e1; transition: .4s; border-radius: 20px; }
        .theme-toggle-slider:before { position: absolute; content: ""; height: 16px; width: 16px; left: 2px; bottom: 2px; background-color: white; transition: .4s; border-radius: 50%; box-shadow: 0 1px 2px rgba(0,0,0,0.2); }
        input:checked + .theme-toggle-slider { background-color: #334155; }
        input:checked + .theme-toggle-slider:before { transform: translateX(20px); }
        .theme-icon { font-size: 12px; color: var(--text-color); }
        .dashboard-header { background: linear-gradient(180deg, #1e293b 0%, #334155 100%); color: #fff; padding: 3rem 2rem; text-align: center; position: relative; overflow: hidden; box-shadow: 0 4px 24px rgba(0,0,0,0.08); border-bottom: 3px solid #475569; }
        .dashboard-title { position: relative; z-index: 2; display: flex; align-items: center; justify-content: center; gap: 1.5rem; margin-bottom: 1rem; }
        .dashboard-title h1 { margin: 0; font-size: 2.5rem; font-weight: 700; color: #fff; letter-spacing: -0.02em; }
        .logo { height: 50px; width: 50px; filter: drop-shadow(0 6px 16px rgba(0,0,0,0.3)); }
        .report-date { font-size: 0.9375rem; color: rgba(255,255,255,0.88); position: relative; z-index: 2; background: rgba(255,255,255,0.08); padding: 0.6rem 1.25rem; border-radius: 999px; border: 1px solid rgba(255,255,255,0.15); display: inline-flex; align-items: center; gap: 0.6rem; font-weight: 500; }
        .overview-container { background: linear-gradient(135deg, var(--bg-color) 0%, var(--bg-secondary) 100%); border-radius: 16px; padding: 2rem; box-shadow: var(--shadow-lg); position: relative; overflow: hidden; }
        .overview-container::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: #475569; }
        .overview-header { text-align: center; margin-bottom: 2.5rem; background: linear-gradient(180deg, #1e293b 0%, #334155 100%); color: #fff; margin-left: -2rem; margin-right: -2rem; margin-top: -2rem; padding: 2rem 2rem 1.5rem; border-radius: 16px 16px 0 0; border-bottom: 2px solid rgba(255,255,255,0.06); }
        .overview-header h2 { font-size: 2.5rem; font-weight: 700; margin-bottom: 0.5rem; color: #fff; letter-spacing: -0.02em; }
        .overview-header p { font-size: 1.125rem; color: rgba(255,255,255,0.9); margin: 0; }
        .summary-card { background: var(--card-bg); border: 1px solid var(--border-color); border-radius: 16px; padding: 1.5rem; transition: all 0.3s ease; box-shadow: var(--shadow-sm); min-height: 160px; display: flex; flex-direction: column; justify-content: center; align-items: center; }
        .overview-tiles-row .summary-card { min-height: 100px; padding: 0.75rem 1rem; }
        .overview-tiles-row .summary-card .card-icon { width: 36px; height: 36px; font-size: 1rem; margin: 0 auto 0.4rem; }
        .overview-tiles-row .summary-card .card-title { font-size: 1.5rem; margin: 0.2rem 0; }
        .overview-tiles-row .summary-card .card-text { min-height: 2em; font-size: 0.8rem; display: flex; align-items: center; justify-content: center; text-align: center; }
        .summary-card .card-icon { width: 60px; height: 60px; border-radius: 16px; display: flex; align-items: center; justify-content: center; margin: 0 auto 1rem; font-size: 1.5rem; color: white; }
        .summary-card.border-primary .card-icon { background: linear-gradient(135deg, #0078d4, #106ebe); }
        .summary-card.border-success .card-icon { background: linear-gradient(135deg, #107c10, #0e6e0e); }
        .summary-card.border-warning .card-icon { background: linear-gradient(135deg, #ffc107, #e0a800); }
        .summary-card.border-danger .card-icon { background: linear-gradient(135deg, #d83b01, #c13401); }
        .summary-card .card-title { font-size: 2.5rem; font-weight: 700; margin: 0.5rem 0; }
        .summary-card .card-text { font-size: 0.9rem; font-weight: 500; color: var(--text-secondary); margin: 0; line-height: 1.4; text-align: center; }
        .modern-table-container { background: var(--card-bg); border-radius: 16px; box-shadow: var(--shadow-lg); overflow: hidden; border: 1px solid var(--border-color); }
        .modern-table-header { background: linear-gradient(180deg, #1e293b 0%, #334155 100%); padding: 1.5rem; border-bottom: 2px solid #475569; color: #fff; }
        .modern-table-header h5 { margin: 0; font-weight: 600; font-size: 1.25rem; color: #fff; }
        .modern-table-header small { color: rgba(255,255,255,0.85); font-size: 0.875rem; }
        .modern-table-body { padding: 0; }
        .modern-table { margin: 0; }
        .modern-table thead th { background: var(--bg-secondary); border: none; font-weight: 600; font-size: 0.875rem; padding: 1rem 0.75rem; text-transform: uppercase; letter-spacing: 0.5px; }
        .modern-table tbody tr { border-bottom: 1px solid var(--border-color); transition: background-color 0.2s ease; }
        .modern-table tbody tr:hover { background-color: var(--table-hover-bg); }
        .modern-table tbody td { padding: 1rem 0.75rem; border: none; vertical-align: middle; }
        .assignment-filters-modern { background: linear-gradient(180deg, var(--card-bg) 0%, var(--bg-secondary) 100%); border-bottom: 1px solid var(--border-color); padding: 1rem 1.25rem; }
        .assignment-filters-modern .assignment-filters-inner { display: flex; flex-wrap: wrap; align-items: flex-end; gap: 1rem; }
        .assignment-filters-modern .filter-group { display: flex; flex-direction: column; gap: 0.25rem; }
        .assignment-filters-modern .filter-label { font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.03em; color: var(--text-secondary); margin: 0; }
        .assignment-filters-modern .filter-select { min-width: 140px; padding: 0.5rem 0.75rem; border-radius: 8px; border: 1px solid var(--border-color); background: var(--card-bg); font-size: 0.875rem; color: var(--text-color); }
        .assignment-filters-modern .filter-reset-btn { display: inline-flex; align-items: center; gap: 0.35rem; padding: 0.5rem 0.9rem; border-radius: 8px; border: 1px solid var(--border-color); background: var(--card-bg); font-size: 0.875rem; color: var(--text-secondary); cursor: pointer; height: 2.15rem; }
        .assignment-filters-modern .filter-multiselect { min-height: 80px; min-width: 160px; }
        .assignment-filters-modern .filter-checkbox-group { display: flex; flex-direction: column; justify-content: flex-end; }
        .assignment-filters-modern .filter-checkbox-label { font-size: 0.875rem; color: var(--text-color); margin: 0; cursor: pointer; display: inline-flex; align-items: center; gap: 0.4rem; white-space: nowrap; }
        [data-theme="dark"] .assignment-filters-modern { background: linear-gradient(180deg, var(--bg-tertiary) 0%, var(--card-bg) 100%); border-color: var(--border-color); }
        [data-theme="dark"] .assignment-filters-modern .filter-select, [data-theme="dark"] .assignment-filters-modern .filter-reset-btn { background: var(--card-bg); color: var(--text-color); border-color: var(--border-color); }
        [data-theme="dark"] .assignment-filters-modern .filter-dropdown-btn { background: var(--card-bg); border-color: var(--border-color); color: var(--text-color); }
        [data-theme="dark"] .assignment-filters-modern .filter-dropdown-btn:hover, [data-theme="dark"] .assignment-filters-modern .filter-dropdown-btn:focus, [data-theme="dark"] .assignment-filters-modern .filter-dropdown-btn.show { background: var(--bg-tertiary); border-color: var(--border-color); color: var(--text-color); }
        [data-theme="dark"] .assignment-filters-modern .filter-checkbox-dropdown { background: var(--card-bg); border-color: var(--border-color); }
        [data-theme="dark"] .assignment-filters-modern .filter-checkbox-dropdown .dropdown-item { color: var(--text-color); }
        [data-theme="dark"] .assignment-filters-modern .filter-checkbox-dropdown .dropdown-item:hover { background: var(--bg-tertiary); color: var(--text-color); }
        [data-theme="dark"] .assignment-filters-modern .filter-checkbox-dropdown label { color: var(--text-color); }
        [data-theme="dark"] .assignment-filters-modern .filter-checkbox-dropdown input.filter-cb { accent-color: #a3a3a3; }
        [data-theme="dark"] .dashboard-header { background: linear-gradient(180deg, #171717 0%, #262626 100%); border-bottom-color: var(--border-color); box-shadow: none; }
        [data-theme="dark"] .dashboard-title h1 { color: var(--text-color); }
        [data-theme="dark"] .report-date { background: rgba(255,255,255,0.06); border-color: var(--border-color); color: var(--text-secondary); }
        [data-theme="dark"] .overview-header { background: var(--bg-tertiary); color: var(--text-color); border-bottom-color: var(--border-color); }
        [data-theme="dark"] .overview-header h2 { color: var(--text-color); }
        [data-theme="dark"] .overview-header p { color: var(--text-secondary); opacity: 1; }
        [data-theme="dark"] .table, [data-theme="dark"] .modern-table { background: var(--card-bg); color: var(--text-color); border-color: var(--border-color); }
        [data-theme="dark"] .table thead th, [data-theme="dark"] .modern-table thead th { background: var(--bg-tertiary); color: var(--text-color); border-color: var(--border-color); }
        [data-theme="dark"] .table tbody td, [data-theme="dark"] .modern-table tbody td { background: var(--card-bg); color: var(--text-color); border-color: var(--border-color); }
        [data-theme="dark"] .table tbody tr:hover td, [data-theme="dark"] .modern-table tbody tr:hover td { background: var(--table-hover-bg); }
        [data-theme="dark"] .table-striped tbody tr:nth-of-type(odd) td { background: var(--bg-secondary); }
        [data-theme="dark"] .dataTables_wrapper { color: var(--text-color); background: var(--card-bg); border-color: var(--border-color); }
        [data-theme="dark"] .dataTables_wrapper .table { background: var(--card-bg); color: var(--text-color); }
        [data-theme="dark"] .modern-table-container .modern-table-body, [data-theme="dark"] .modern-table-container .table-responsive { background: var(--card-bg); }
        [data-theme="dark"] .modern-table-container { background: var(--card-bg); }
        [data-theme="dark"] .modern-table-header { background: var(--bg-tertiary); color: var(--text-color); border-bottom-color: var(--border-color); }
        [data-theme="dark"] .modern-table-header h5, [data-theme="dark"] .modern-table-header small { color: var(--text-color); }
        [data-theme="dark"] footer { background: linear-gradient(180deg, #171717 0%, #262626 100%); border-top-color: var(--border-color); }
        [data-theme="dark"] .dataTables_filter input, [data-theme="dark"] .dataTables_length select { background: var(--bg-tertiary); color: var(--text-color); border-color: var(--border-color); }
        [data-theme="dark"] .dataTables_info { color: var(--text-secondary); }
        [data-theme="dark"] .dataTables_paginate .page-link { background: var(--card-bg); color: var(--text-color); border-color: var(--border-color); }
        [data-theme="dark"] .dataTables_paginate .page-link:hover { background: var(--bg-tertiary); color: var(--text-color); }
        [data-theme="dark"] .dataTables_paginate .page-item.active .page-link { background: var(--bg-tertiary); border-color: var(--border-color); color: var(--text-color); }
        [data-theme="dark"] .dataTables_paginate .page-item.disabled .page-link { background: var(--bg-secondary); color: var(--text-secondary); }
        [data-theme="dark"] .dt-button { background: var(--bg-tertiary); color: var(--text-color); border-color: var(--border-color); }
        [data-theme="dark"] .dt-button:hover { background: var(--table-hover-bg); color: var(--text-color); }
        footer { background: linear-gradient(180deg, #1e293b 0%, #334155 100%); color: rgba(255,255,255,0.9); padding: 2rem 0 1rem 0; margin-top: 3rem; border-top: 3px solid #475569; }
        .footer-content { max-width: 1200px; margin: 0 auto; padding: 0 1rem; display: grid; grid-template-columns: 1fr auto 1fr; align-items: center; gap: 2rem; }
        .footer-brand { font-size: 1.1rem; font-weight: 600; margin-bottom: 0.5rem; }
        .footer-description { font-size: 0.85rem; opacity: 0.9; margin: 0; }
        .footer-center { text-align: center; }
        .footer-logo { width: 40px; height: 40px; background: rgba(255,255,255,0.1); border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 0.5rem; font-size: 1.2rem; font-weight: bold; }
        .footer-tagline { font-size: 0.8rem; opacity: 0.8; margin: 0; }
        .footer-links { text-align: right; }
        .footer-social { display: flex; gap: 1rem; justify-content: flex-end; margin-bottom: 0.5rem; }
        .footer-social a { color: rgba(255,255,255,0.8); transition: all 0.3s ease; }
        .footer-copyright { font-size: 0.75rem; opacity: 0.7; margin: 0; }
        .footer-bottom { margin-top: 1.5rem; padding-top: 1rem; border-top: 1px solid rgba(255,255,255,0.1); text-align: center; }
        .footer-tech-stack { font-size: 0.7rem; opacity: 0.6; margin: 0.5rem 0 0 0; }
        .tech-badge { display: inline-block; background: rgba(255,255,255,0.15); padding: 0.2rem 0.5rem; border-radius: 6px; margin: 0 0.2rem; }
    </style>
</head>
<body>
    <div class="theme-toggle">
        <div class="theme-icon"><i class="fas fa-sun"></i></div>
        <label class="theme-toggle-switch">
            <input type="checkbox" id="themeToggle">
            <span class="theme-toggle-slider"></span>
        </label>
        <div class="theme-icon"><i class="fas fa-moon"></i></div>
    </div>
    <div class="app-container">
        <div class="container">
            <div class="dashboard-header">
                <div class="dashboard-title">
                    <svg class="logo" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48">
                        <path fill="#ff5722" d="M6 6H22V22H6z" transform="rotate(-180 14 14)"/>
                        <path fill="#4caf50" d="M26 6H42V22H26z" transform="rotate(-180 34 14)"/>
                        <path fill="#ffc107" d="M6 26H22V42H6z" transform="rotate(-180 14 34)"/>
                        <path fill="#03a9f4" d="M26 26H42V42H26z" transform="rotate(-180 34 34)"/>
                    </svg>
                    <h1>$TenantName Enrollment Flow Visualization</h1>
                </div>
                <div class="report-date">
                    <i class="fas fa-calendar-alt me-2"></i>
                    Report generated on: $(Get-Date -Format 'MMMM dd, yyyy \a\t HH:mm')
                </div>
            </div>
            <ul class="nav nav-tabs" id="flowTabs" role="tablist">
                <li class="nav-item" role="presentation">
                    <button class="nav-link active" id="assignment-overview-tab" data-bs-toggle="tab" data-bs-target="#assignment-overview" type="button" role="tab" aria-controls="assignment-overview" aria-selected="true">
                        <i class="fas fa-chart-pie me-2"></i>Assignment Overview
                    </button>
                </li>
            </ul>
            <div class="tab-content" id="flowTabContent">
                <div class="tab-pane fade show active" id="assignment-overview" role="tabpanel" aria-labelledby="assignment-overview-tab" tabindex="0">
$fragment
                </div>
            </div>
        </div>
    </div>
    <footer>
        <div class="footer-content">
            <div class="footer-info">
                <div class="footer-brand">RK Solutions</div>
                <p class="footer-description">Helping IT professionals with Microsoft 365 and Intune reporting tools and insights</p>
            </div>
            <div class="footer-center">
                <div class="footer-logo">RK</div>
                <p class="footer-tagline">Practical IT Solutions & Insights</p>
            </div>
            <div class="footer-links">
                <div class="footer-social">
                    <a href="https://rksolutions.nl" title="Visit RK Solutions Blog" target="_blank">🌐</a>
                    <a href="https://github.com/royklo" title="GitHub" target="_blank">🔗</a>
                    <a href="https://linkedin.com/in/roy-klooster" title="LinkedIn" target="_blank">💼</a>
                </div>
                <p class="footer-copyright">© 2025 Roy Klooster - RK Solutions</p>
            </div>
        </div>
        <div class="footer-bottom">
            <p class="footer-copyright">Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
            <p class="footer-tech-stack">
                <span class="tech-badge">PowerShell</span>
                <span class="tech-badge">Microsoft Graph</span>
                <span class="tech-badge">Intune</span>
                <span class="tech-badge">HTML5</span>
            </p>
        </div>
    </footer>
    <script>
$assignmentOverviewScript
    </script>
</body>
</html>
"@
    $outDir = Split-Path -Parent $OutputPath
    if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    if (-not [System.IO.Path]::HasExtension($OutputPath)) { $OutputPath = $OutputPath + ".html" }
    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($false))
    $OutputPath
}

function Get-ArchitectureDiagramFragment {
    param([Parameter(Mandatory)][array]$EvaluatedAssignments, [Parameter(Mandatory)][string]$DeviceName, [array]$DeviceGroupDetails = @(), [string]$IntuneDeviceId = "", [string]$EntraDeviceId = "", [Parameter(Mandatory = $false)][string]$DevicePlatform = $null, [Parameter(Mandatory = $false)][switch]$IsCloudPC = $false, [array]$DeviceGroupIds = @(), [array]$UserGroupIds = @())
    $applied = $EvaluatedAssignments | Where-Object { $_.AppliesToDevice -eq $true }
    if ($DevicePlatform) {
        $applied = $applied | Where-Object { Test-AssignmentMatchesDevicePlatform -AssignmentRow $_ -DevicePlatform $DevicePlatform }
    }
    $hasCategory = ($EvaluatedAssignments | Select-Object -First 1).PSObject.Properties.Name -contains 'PolicyCategory'
    $configCategories = @("Device Configuration", "Settings Catalog", "Administrative Templates", "Endpoint Security", "Platform Scripts", "Remediation Scripts")
    $isDeviceAssignment = { param($at, $gid) (Get-GroupMemberTargetTypeFromRow -AssignmentType $at -GroupId $gid -DeviceGroupIds $DeviceGroupIds -UserGroupIds $UserGroupIds) -eq 'Device' }
    $isUserAssignment = { param($at, $gid) (Get-GroupMemberTargetTypeFromRow -AssignmentType $at -GroupId $gid -DeviceGroupIds $DeviceGroupIds -UserGroupIds $UserGroupIds) -eq 'User' }
    $autopilotPolicies = [System.Collections.ArrayList]::new()
    $espPolicies = [System.Collections.ArrayList]::new()
    $configDevicePolicies = [System.Collections.ArrayList]::new()
    $configUserPolicies = [System.Collections.ArrayList]::new()
    $securityBaselineDevicePolicies = [System.Collections.ArrayList]::new()
    $securityBaselineUserPolicies = [System.Collections.ArrayList]::new()
    $complianceDevicePolicies = [System.Collections.ArrayList]::new()
    $complianceUserPolicies = [System.Collections.ArrayList]::new()
    $appDevicePolicies = [System.Collections.ArrayList]::new()
    $appUserPolicies = [System.Collections.ArrayList]::new()
    $cloudPCProvisioningPolicies = [System.Collections.ArrayList]::new()
    $cloudPCUserSettingsPolicies = [System.Collections.ArrayList]::new()
    foreach ($row in $applied) {
        $cat = if ($hasCategory -and $row.PolicyCategory) { $row.PolicyCategory } else { $null }
        $name = $row.PolicyName
        $at = $row.AssignmentType
        $gid = $row.GroupId
        if ($cat -eq "Autopilot Profile") { if (-not $autopilotPolicies.Contains($name)) { [void]$autopilotPolicies.Add($name) } }
        elseif ($cat -eq "Enrollment Status Page") { if (-not $espPolicies.Contains($name)) { [void]$espPolicies.Add($name) } }
        elseif ($cat -eq "Security Baselines") {
            if (& $isDeviceAssignment $at $gid) { if (-not $securityBaselineDevicePolicies.Contains($name)) { [void]$securityBaselineDevicePolicies.Add($name) } }
            if (& $isUserAssignment $at $gid) { if (-not $securityBaselineUserPolicies.Contains($name)) { [void]$securityBaselineUserPolicies.Add($name) } }
        }
        elseif ($cat -and $configCategories -contains $cat) {
            if (& $isDeviceAssignment $at $gid) { if (-not $configDevicePolicies.Contains($name)) { [void]$configDevicePolicies.Add($name) } }
            if (& $isUserAssignment $at $gid) { if (-not $configUserPolicies.Contains($name)) { [void]$configUserPolicies.Add($name) } }
        }
        elseif ($cat -eq "Compliance Policies") {
            if (& $isDeviceAssignment $at $gid) { if (-not $complianceDevicePolicies.Contains($name)) { [void]$complianceDevicePolicies.Add($name) } }
            if (& $isUserAssignment $at $gid) { if (-not $complianceUserPolicies.Contains($name)) { [void]$complianceUserPolicies.Add($name) } }
        }
        elseif ($cat -eq "Cloud PC Provisioning") { if (-not $cloudPCProvisioningPolicies.Contains($name)) { [void]$cloudPCProvisioningPolicies.Add($name) } }
        elseif ($cat -eq "Cloud PC User Settings") { if (-not $cloudPCUserSettingsPolicies.Contains($name)) { [void]$cloudPCUserSettingsPolicies.Add($name) } }
        elseif ($cat -eq "Applications") {
            if (& $isDeviceAssignment $at $gid) { if (-not $appDevicePolicies.Contains($name)) { [void]$appDevicePolicies.Add($name) } }
            if (& $isUserAssignment $at $gid) { if (-not $appUserPolicies.Contains($name)) { [void]$appUserPolicies.Add($name) } }
        }
    }
    # Cloud PC: also include from all assignments so architecture shows policy names/counts even when not applied to this device
    if ($IsCloudPC) {
        foreach ($row in $EvaluatedAssignments) {
            $cat = if ($row.PolicyCategory) { [string]$row.PolicyCategory } else { '' }
            if ($cat.Trim().Length -eq 0) { continue }
            $catNorm = $cat.Trim()
            if ($catNorm -eq "Cloud PC Provisioning" -and -not $cloudPCProvisioningPolicies.Contains($row.PolicyName)) { [void]$cloudPCProvisioningPolicies.Add($row.PolicyName) }
            if ($catNorm -eq "Cloud PC User Settings" -and -not $cloudPCUserSettingsPolicies.Contains($row.PolicyName)) { [void]$cloudPCUserSettingsPolicies.Add($row.PolicyName) }
        }
    }
    $deviceEnc = [System.Net.WebUtility]::HtmlEncode($DeviceName)
    $intuneIdEnc = if ($IntuneDeviceId) { [System.Net.WebUtility]::HtmlEncode($IntuneDeviceId) } else { "-" }
    $entraIdEnc = if ($EntraDeviceId) { [System.Net.WebUtility]::HtmlEncode($EntraDeviceId) } else { "-" }
    $deviceTableRow = "<tr><td class=`"arch-group-cell`">$deviceEnc</td><td class=`"arch-group-cell`">$intuneIdEnc</td><td class=`"arch-group-cell`">$entraIdEnc</td></tr>"
    $groupTableRows = ""
    if (-not $DeviceGroupDetails -or $DeviceGroupDetails.Count -eq 0) {
        $groupTableRows = "<tr><td class=`"arch-group-cell`" colspan=`"3`">-</td></tr>"
    }
    else {
        foreach ($gd in $DeviceGroupDetails) {
            $encName = [System.Net.WebUtility]::HtmlEncode($gd.DisplayName)
            $encType = [System.Net.WebUtility]::HtmlEncode($gd.GroupType)
            $ruleCell = if ($gd.MembershipRule) { $gd.MembershipRule } else { "-" }
            $groupTableRows += "<tr><td class=`"arch-group-cell`">$encName</td><td class=`"arch-group-type`">$encType</td><td class=`"arch-group-rule`">$ruleCell</td></tr>"
        }
    }
    $autopilotName = if ($autopilotPolicies.Count -gt 0) { [System.Net.WebUtility]::HtmlEncode($autopilotPolicies[0]) } else { "- None -" }
    $espName = if ($espPolicies.Count -gt 0) { [System.Net.WebUtility]::HtmlEncode($espPolicies[0]) } else { "- None -" }
    $cloudPCSuffix = if ($IsCloudPC) { ", Cloud PC Prov($($cloudPCProvisioningPolicies.Count)), Cloud PC User($($cloudPCUserSettingsPolicies.Count))" } else { "" }
    $deviceSetupSub = "Configuration($($configDevicePolicies.Count)), Security Baselines($($securityBaselineDevicePolicies.Count)), Compliance($($complianceDevicePolicies.Count)), Apps($($appDevicePolicies.Count))$cloudPCSuffix"
    $accountSetupSub = "Configuration($($configUserPolicies.Count)), Security Baselines($($securityBaselineUserPolicies.Count)), Compliance($($complianceUserPolicies.Count)), Apps($($appUserPolicies.Count))$cloudPCSuffix"
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('<div class="arch-diagram-wrapper arch-diagram-html">')
    $archStep = { param($title, $subtitle, $fill) @"
<div class="arch-step" style="--arch-fill:$fill;">
<div class="arch-step-title">$([System.Net.WebUtility]::HtmlEncode($title))</div>
<div class="arch-step-sub">$subtitle</div>
</div>
"@ }
    [void]$sb.AppendLine('<div class="arch-flow">')
    [void]$sb.AppendLine(@"
<div class="arch-step arch-step-device" style="--arch-fill:#0f766e;">
<div class="arch-step-title">Device</div>
<div class="arch-group-table-wrap">
<table class="arch-group-table">
<thead><tr><th class="arch-group-th-name">Intune Device Name</th><th class="arch-group-th-type">Intune Object ID</th><th class="arch-group-th-rule">Entra Object ID</th></tr></thead>
<tbody>$deviceTableRow</tbody>
</table>
</div>
</div>
"@)
    [void]$sb.AppendLine('<div class="arch-arrow" aria-hidden="true"><i class="fas fa-chevron-right"></i></div>')
    [void]$sb.AppendLine(@"
<div class="arch-step arch-step-memberof" style="--arch-fill:#115e59;">
<div class="arch-step-title">Entra ID Groups</div>
<div class="arch-group-table-wrap">
<table class="arch-group-table">
<thead><tr><th class="arch-group-th-name">Group name</th><th class="arch-group-th-type">Assignment</th><th class="arch-group-th-rule">Membership rule</th></tr></thead>
<tbody>$groupTableRows</tbody>
</table>
</div>
</div>
"@)
    [void]$sb.AppendLine('<div class="arch-arrow" aria-hidden="true"><i class="fas fa-chevron-right"></i></div>')
    if ($IsCloudPC) {
        $cloudPCProvName = if ($cloudPCProvisioningPolicies.Count -gt 0) { [System.Net.WebUtility]::HtmlEncode($cloudPCProvisioningPolicies[0]) } else { "- None -" }
        [void]$sb.AppendLine(@"
<div class="arch-step arch-step-cloudpc-provisioning" style="--arch-fill:#0369a1;">
<div class="arch-step-title">Cloud PC Provisioning</div>
<div class="arch-step-sub">$cloudPCProvName</div>
</div>
"@)
        [void]$sb.AppendLine('<div class="arch-arrow" aria-hidden="true"><i class="fas fa-chevron-right"></i></div>')
    }
    [void]$sb.AppendLine(@"
<div class="arch-step arch-step-autopilot-esp" style="--arch-fill:#0f766e;">
<div class="arch-autopilot-esp-inner">
<div class="arch-esp-block" style="--arch-fill:#0f766e;">
<div class="arch-step-title">Autopilot Profile</div>
<div class="arch-step-sub">$autopilotName</div>
</div>
<div class="arch-arrow-down" aria-hidden="true"><i class="fas fa-chevron-down"></i></div>
<div class="arch-esp-block" style="--arch-fill:#115e59;">
<div class="arch-step-title">Enrollment Status Page</div>
<div class="arch-step-sub">$espName</div>
</div>
</div>
</div>
"@)
    [void]$sb.AppendLine('<div class="arch-arrow" aria-hidden="true"><i class="fas fa-chevron-right"></i></div>')
    [void]$sb.AppendLine(@"
<div class="arch-step arch-step-enrollment" style="--arch-fill:#475569;">
<div class="arch-step-title">Autopilot Enrollment</div>
<div class="arch-enrollment-inner">
<div class="arch-enrollment-row">
<div class="arch-inner-title">Device Setup</div>
<div class="arch-inner-sub">$([System.Net.WebUtility]::HtmlEncode($deviceSetupSub))</div>
</div>
<div class="arch-enrollment-row">
<div class="arch-inner-title">Account Setup</div>
<div class="arch-inner-sub">$([System.Net.WebUtility]::HtmlEncode($accountSetupSub))</div>
</div>
</div>
</div>
"@)
    [void]$sb.AppendLine('<div class="arch-arrow" aria-hidden="true"><i class="fas fa-chevron-right"></i></div>')
    [void]$sb.AppendLine((& $archStep "Enrolled" $deviceEnc "#047857"))
    [void]$sb.AppendLine('</div>')
    [void]$sb.AppendLine('</div>')
    $sb.ToString()
}

function Get-AppliedFlowHtmlFragment {
    param([Parameter(Mandatory)][array]$EvaluatedAssignments, [Parameter(Mandatory)][string]$DeviceName, [array]$DeviceGroupDetails = @(), [Parameter(Mandatory = $false)][string]$DevicePlatform = $null, [Parameter(Mandatory = $false)][switch]$IsCloudPC = $false, [array]$DeviceGroupIds = @(), [array]$UserGroupIds = @())
    # Filter to only show assignments that actually apply to this device
    $applied = $EvaluatedAssignments | Where-Object { $_.AppliesToDevice -eq $true }
    if ($DevicePlatform) {
        $applied = $applied | Where-Object { Test-AssignmentMatchesDevicePlatform -AssignmentRow $_ -DevicePlatform $DevicePlatform }
    }
    $hasCategory = ($EvaluatedAssignments | Select-Object -First 1).PSObject.Properties.Name -contains 'PolicyCategory'
    $configCategories = @("Device Configuration", "Settings Catalog", "Administrative Templates", "Endpoint Security", "Platform Scripts", "Remediation Scripts")
    # Determine if an assignment targets a device group or user group by checking the group's actual member types
    $isDeviceAssignment = { param($at, $gid) (Get-GroupMemberTargetTypeFromRow -AssignmentType $at -GroupId $gid -DeviceGroupIds $DeviceGroupIds -UserGroupIds $UserGroupIds) -eq 'Device' }
    $isUserAssignment = { param($at, $gid) (Get-GroupMemberTargetTypeFromRow -AssignmentType $at -GroupId $gid -DeviceGroupIds $DeviceGroupIds -UserGroupIds $UserGroupIds) -eq 'User' }
    $autopilotPolicies = [System.Collections.ArrayList]::new()
    $espPolicies = [System.Collections.ArrayList]::new()
    $configDevicePolicies = [System.Collections.ArrayList]::new()
    $configUserPolicies = [System.Collections.ArrayList]::new()
    $securityBaselineDevicePolicies = [System.Collections.ArrayList]::new()
    $securityBaselineUserPolicies = [System.Collections.ArrayList]::new()
    $complianceDevicePolicies = [System.Collections.ArrayList]::new()
    $complianceUserPolicies = [System.Collections.ArrayList]::new()
    $appDevicePolicies = [System.Collections.ArrayList]::new()
    $appUserPolicies = [System.Collections.ArrayList]::new()
    $cloudPCProvisioningDevicePolicies = [System.Collections.ArrayList]::new()
    $cloudPCProvisioningUserPolicies = [System.Collections.ArrayList]::new()
    $cloudPCUserSettingsDevicePolicies = [System.Collections.ArrayList]::new()
    $cloudPCUserSettingsUserPolicies = [System.Collections.ArrayList]::new()
    foreach ($row in $applied) {
        $cat = if ($hasCategory -and $row.PolicyCategory) { $row.PolicyCategory } else { $null }
        $name = $row.PolicyName
        $at = $row.AssignmentType
        $gid = $row.GroupId
        if ($cat -eq "Autopilot Profile") {
            if (-not $autopilotPolicies.Contains($name)) { [void]$autopilotPolicies.Add($name) }
        }
        elseif ($cat -eq "Enrollment Status Page") {
            if (-not $espPolicies.Contains($name)) { [void]$espPolicies.Add($name) }
        }
        elseif ($cat -eq "Security Baselines") {
            if (& $isDeviceAssignment $at $gid) { if (-not $securityBaselineDevicePolicies.Contains($name)) { [void]$securityBaselineDevicePolicies.Add($name) } }
            if (& $isUserAssignment $at $gid) { if (-not $securityBaselineUserPolicies.Contains($name)) { [void]$securityBaselineUserPolicies.Add($name) } }
        }
        elseif ($cat -and $configCategories -contains $cat) {
            if (& $isDeviceAssignment $at $gid) { if (-not $configDevicePolicies.Contains($name)) { [void]$configDevicePolicies.Add($name) } }
            if (& $isUserAssignment $at $gid) { if (-not $configUserPolicies.Contains($name)) { [void]$configUserPolicies.Add($name) } }
        }
        elseif ($cat -eq "Compliance Policies") {
            if (& $isDeviceAssignment $at $gid) { if (-not $complianceDevicePolicies.Contains($name)) { [void]$complianceDevicePolicies.Add($name) } }
            if (& $isUserAssignment $at $gid) { if (-not $complianceUserPolicies.Contains($name)) { [void]$complianceUserPolicies.Add($name) } }
        }
        elseif ($cat -eq "Cloud PC Provisioning") {
            if (& $isDeviceAssignment $at $gid) { if (-not $cloudPCProvisioningDevicePolicies.Contains($name)) { [void]$cloudPCProvisioningDevicePolicies.Add($name) } }
            if (& $isUserAssignment $at $gid) { if (-not $cloudPCProvisioningUserPolicies.Contains($name)) { [void]$cloudPCProvisioningUserPolicies.Add($name) } }
        }
        elseif ($cat -eq "Cloud PC User Settings") {
            if (& $isDeviceAssignment $at $gid) { if (-not $cloudPCUserSettingsDevicePolicies.Contains($name)) { [void]$cloudPCUserSettingsDevicePolicies.Add($name) } }
            if (& $isUserAssignment $at $gid) { if (-not $cloudPCUserSettingsUserPolicies.Contains($name)) { [void]$cloudPCUserSettingsUserPolicies.Add($name) } }
        }
        elseif ($cat -eq "Applications") {
            if (& $isDeviceAssignment $at $gid) { if (-not $appDevicePolicies.Contains($name)) { [void]$appDevicePolicies.Add($name) } }
            if (& $isUserAssignment $at $gid) { if (-not $appUserPolicies.Contains($name)) { [void]$appUserPolicies.Add($name) } }
        }
    }
    foreach ($p in $securityBaselineDevicePolicies) { if (-not $configDevicePolicies.Contains($p)) { [void]$configDevicePolicies.Add($p) } }
    foreach ($p in $securityBaselineUserPolicies) { if (-not $configUserPolicies.Contains($p)) { [void]$configUserPolicies.Add($p) } }
    # Cloud PC: single list (always device-assigned conceptually), no device/user columns
    $cloudPCProvisioningMerged = [System.Collections.ArrayList]::new()
    foreach ($p in $cloudPCProvisioningDevicePolicies) { if (-not $cloudPCProvisioningMerged.Contains($p)) { [void]$cloudPCProvisioningMerged.Add($p) } }
    foreach ($p in $cloudPCProvisioningUserPolicies) { if (-not $cloudPCProvisioningMerged.Contains($p)) { [void]$cloudPCProvisioningMerged.Add($p) } }
    $cloudPCUserSettingsMerged = [System.Collections.ArrayList]::new()
    foreach ($p in $cloudPCUserSettingsDevicePolicies) { if (-not $cloudPCUserSettingsMerged.Contains($p)) { [void]$cloudPCUserSettingsMerged.Add($p) } }
    foreach ($p in $cloudPCUserSettingsUserPolicies) { if (-not $cloudPCUserSettingsMerged.Contains($p)) { [void]$cloudPCUserSettingsMerged.Add($p) } }
    # Cloud PC: also include policies from all assignments (not only applied) so flow shows policy name and assignment even when device isn't in target group
    if ($IsCloudPC) {
        foreach ($row in $EvaluatedAssignments) {
            $cat = if ($row.PolicyCategory) { [string]$row.PolicyCategory } else { '' }
            if ($cat.Trim().Length -eq 0) { continue }
            $catNorm = $cat.Trim()
            if ($catNorm -eq "Cloud PC Provisioning") {
                $pname = if ($row.PolicyName) { [string]$row.PolicyName } else { '' }; if ($pname.Trim().Length -gt 0 -and -not $cloudPCProvisioningMerged.Contains($pname)) { [void]$cloudPCProvisioningMerged.Add($pname) }
            }
            if ($catNorm -eq "Cloud PC User Settings") {
                $pname = if ($row.PolicyName) { [string]$row.PolicyName } else { '' }; if ($pname.Trim().Length -gt 0 -and -not $cloudPCUserSettingsMerged.Contains($pname)) { [void]$cloudPCUserSettingsMerged.Add($pname) }
            }
        }
    }
    $deviceEnc = [System.Net.WebUtility]::HtmlEncode($DeviceName)
    function Get-FirstAssignmentAndFilter($policyName, [array]$eval) {
        if (-not $eval -or $eval.Count -eq 0) { return @("-", "-") }
        # Prefer a row with a real assignment (so Cloud PC / others show target when both "Not Assigned" and assigned rows exist)
        $a = $eval | Where-Object { $_.PolicyName -eq $policyName -and $_.AssignmentType -ne "Not Assigned" } | Select-Object -First 1
        if (-not $a) { $a = $eval | Where-Object { $_.PolicyName -eq $policyName } | Select-Object -First 1 }
        if (-not $a) { return @("-", "-") }
        $target = if ($a.AssignmentType -match "^(All Devices|All Users)$") { $a.AssignmentType } else { $a.TargetName }
        $filterD = "No Filter"
        if ($a.FilterName -and $a.FilterName -ne "No Filter" -and $a.FilterName -ne "Filter Not Found") {
            # Show filter name with type: "Filter Name (Include)" or "Filter Name (Exclude)"
            if ($a.FilterType -eq "include") { 
                $filterD = "$($a.FilterName) (include)" 
            } elseif ($a.FilterType -eq "exclude") { 
                $filterD = "$($a.FilterName) (exclude)" 
            } else { 
                $filterD = $a.FilterName 
            }
        }
        # For group assignments: show type (Dynamic/Assigned) like Entra ID groups, and membership rule when dynamic
        if ($a.GroupId -and $script:AllGroups.Count -gt 0) {
            $grp = $script:AllGroups[$a.GroupId]
            if ($grp) {
                $isDynamic = $grp.groupTypes -and (@($grp.groupTypes) -contains 'DynamicMembership')
                # Only override filter display with membership rule if it's a dynamic group AND no Intune filter is set
                if ($isDynamic -and $grp.membershipRule -and ($a.FilterName -eq "No Filter" -or -not $a.FilterName)) { $filterD = $grp.membershipRule }
            }
        }
        return @($target, $filterD)
    }
    function Get-FirstAssignmentGroupTypeRule($policyName, [array]$eval) {
        if (-not $eval -or $eval.Count -eq 0) { return @("Not Assigned", "-", "No Filter") }
        $a = $eval | Where-Object { $_.PolicyName -eq $policyName -and $_.AssignmentType -ne "Not Assigned" } | Select-Object -First 1
        if (-not $a) { $a = $eval | Where-Object { $_.PolicyName -eq $policyName } | Select-Object -First 1 }
        if (-not $a) { return @("Not Assigned", "-", "No Filter") }
        $groupName = if ($a.AssignmentType -match "^(All Devices|All Users)$") { $a.AssignmentType } else { $a.TargetName }
        $typeLabel = "-"
        $rule = if ($a.FilterName -and $a.FilterName -ne "No Filter" -and $a.FilterName -ne "Filter Not Found") { $a.FilterName } else { "No Filter" }
        if ($a.GroupId -and $script:AllGroups.Count -gt 0) {
            $grp = $script:AllGroups[$a.GroupId]
            if ($grp) {
                $groupName = $grp.displayName
                $isDynamic = $grp.groupTypes -and (@($grp.groupTypes) -contains 'DynamicMembership')
                $typeLabel = if ($isDynamic) { "Dynamic" } else { "Static" }
                # Only show membership rule for Dynamic groups; for Static use empty
                if ($isDynamic -and $grp.membershipRule) { $rule = $grp.membershipRule } else { $rule = "-" }
            }
        }
        return @($groupName, $typeLabel, $rule)
    }
    function Write-FlowPolicyTable($categoryTitle, $col1Header, $col2Header, $col3Header, [System.Collections.ArrayList]$list, [array]$eval, [bool]$hideCategoryTitle = $false) {
        $sb = [System.Text.StringBuilder]::new()
        if ($list.Count -eq 0) {
            [void]$sb.AppendLine('<div class="policy-category policy-empty"><p class="policy-empty-msg text-muted mb-0"><i class="fas fa-inbox me-2"></i>No assignments</p></div>')
            return $sb.ToString()
        }
        $titleEnc = [System.Net.WebUtility]::HtmlEncode($categoryTitle)
        $c1 = [System.Net.WebUtility]::HtmlEncode($col1Header)
        $c2 = [System.Net.WebUtility]::HtmlEncode($col2Header)
        $c3 = [System.Net.WebUtility]::HtmlEncode($col3Header)
        [void]$sb.AppendLine('<div class="policy-category">')
        if (-not $hideCategoryTitle) {
            [void]$sb.AppendLine('<h5>' + $titleEnc + ' <span class="policy-count">' + $list.Count + '</span></h5>')
        }
        [void]$sb.AppendLine('<div class="policy-table-header" style="grid-template-columns: 3fr 1.5fr 2fr;">')
        [void]$sb.AppendLine('<div class="header-policy-name">' + $c1 + '</div><div class="header-assignment">' + $c2 + '</div><div class="header-filter">' + $c3 + '</div>')
        [void]$sb.AppendLine('</div>')
        foreach ($p in $list) {
            $pEnc = [System.Net.WebUtility]::HtmlEncode($p)
            $assignFilter = Get-FirstAssignmentAndFilter $p $eval
            $aEnc = [System.Net.WebUtility]::HtmlEncode($assignFilter[0])
            $fEnc = [System.Net.WebUtility]::HtmlEncode($assignFilter[1])
            [void]$sb.AppendLine('<div class="policy-item" style="grid-template-columns: 3fr 1.5fr 2fr;"><div class="policy-name">' + $pEnc + '</div><div class="policy-assignment">' + $aEnc + '</div><div class="policy-filter">' + $fEnc + '</div></div>')
        }
        [void]$sb.AppendLine('</div>')
        $sb.ToString()
    }
    function Write-FlowPolicyTable4Col($categoryTitle, $col1Header, $col2Header, $col3Header, $col4Header, [System.Collections.ArrayList]$list, [array]$eval, [bool]$hideCategoryTitle = $false) {
        $sb = [System.Text.StringBuilder]::new()
        if ($list.Count -eq 0) {
            [void]$sb.AppendLine('<div class="policy-category policy-empty"><p class="policy-empty-msg text-muted mb-0"><i class="fas fa-inbox me-2"></i>No assignments</p></div>')
            return $sb.ToString()
        }
        $titleEnc = [System.Net.WebUtility]::HtmlEncode($categoryTitle)
        $c1 = [System.Net.WebUtility]::HtmlEncode($col1Header)
        $c2 = [System.Net.WebUtility]::HtmlEncode($col2Header)
        $c3 = [System.Net.WebUtility]::HtmlEncode($col3Header)
        $c4 = [System.Net.WebUtility]::HtmlEncode($col4Header)
        [void]$sb.AppendLine('<div class="policy-category">')
        if (-not $hideCategoryTitle) {
            [void]$sb.AppendLine('<h5>' + $titleEnc + ' <span class="policy-count">' + $list.Count + '</span></h5>')
        }
        [void]$sb.AppendLine('<div class="policy-table-header policy-table-header-4col" style="grid-template-columns: 2fr 2fr 1fr 2fr;">')
        [void]$sb.AppendLine('<div class="header-policy-name">' + $c1 + '</div><div class="header-assignment">' + $c2 + '</div><div class="header-type">' + $c3 + '</div><div class="header-filter">' + $c4 + '</div>')
        [void]$sb.AppendLine('</div>')
        foreach ($p in $list) {
            $pEnc = [System.Net.WebUtility]::HtmlEncode($p)
            $gtr = Get-FirstAssignmentGroupTypeRule $p $eval
            $gEnc = [System.Net.WebUtility]::HtmlEncode($gtr[0])
            $tEnc = [System.Net.WebUtility]::HtmlEncode($gtr[1])
            $rEnc = [System.Net.WebUtility]::HtmlEncode($gtr[2])
            [void]$sb.AppendLine('<div class="policy-item policy-item-4col" style="grid-template-columns: 2fr 2fr 1fr 2fr;"><div class="policy-name">' + $pEnc + '</div><div class="policy-assignment">' + $gEnc + '</div><div class="policy-type">' + $tEnc + '</div><div class="policy-filter">' + $rEnc + '</div></div>')
        }
        [void]$sb.AppendLine('</div>')
        $sb.ToString()
    }
    function Write-FlowStepLikeReference($stepId, $stepClass, $title, $iconClass, $list, [array]$eval, $stepNameLabel, $categoryTitle, $col1Header, [string]$ConfigSectionHtml = "", [switch]$UseGroupTypeRuleTable) {
        $t = [System.Net.WebUtility]::HtmlEncode($title)
        $n = $list.Count
        $stepName = if ($list.Count -gt 0) { [System.Net.WebUtility]::HtmlEncode($list[0]) } else { "- None -" }
        # Don't show step-name for any policy type - it's redundant with the header and button
        $stepNameLine = ""
        $btnText = if ($stepId -eq "flow-autopilot") { "Expand Autopilot Profile Details" } elseif ($stepId -eq "flow-esp") { "Expand Enrollment Status Page Configuration" } else { "Expand " + $title + " (" + $n + ")" }
        $tableHtml = if ($UseGroupTypeRuleTable) { Write-FlowPolicyTable4Col $categoryTitle "Policy Name" "Group" "Type" "Membership rule" $list $eval } else { Write-FlowPolicyTable $categoryTitle $col1Header "Assignment" "Filter" $list $eval }
        @"
<div class=`"flow-step $stepClass`">
<div class=`"step-header`"><i class=`"fas $iconClass`"></i><span>$t</span></div>
<div class=`"step-content`">
<button class=`"btn btn-outline-primary btn-sm step-toggle`" type=`"button`" data-bs-toggle=`"collapse`" data-bs-target=`"#$stepId`" aria-expanded=`"true`" aria-controls=`"$stepId`"><i class=`"fas fa-chevron-down`"></i>$btnText</button>
<div class=`"collapse show`" id=`"$stepId`">$stepNameLine
$tableHtml
$ConfigSectionHtml
</div>
</div>
</div>
"@
    }
    $autopilotConfig = if ($autopilotPolicies.Count -gt 0) { Get-AutopilotProfileConfigByDisplayName -DisplayName $autopilotPolicies[0] } else { $null }
    $apDeviceName = if ($autopilotConfig) { [System.Net.WebUtility]::HtmlEncode($autopilotConfig.DeviceNameTemplate) } else { "-" }
    $apLanguage = if ($autopilotConfig) { [System.Net.WebUtility]::HtmlEncode($autopilotConfig.Language) } else { "-" }
    $apLocale = if ($autopilotConfig) { [System.Net.WebUtility]::HtmlEncode($autopilotConfig.Locale) } else { "-" }
    $autopilotConfigSection = @"
<div class=`"autopilot-config-section`">
<button class=`"btn btn-outline-secondary btn-sm config-toggle`" type=`"button`" data-bs-toggle=`"collapse`" data-bs-target=`"#autopilotConfig-0`" aria-expanded=`"false`" aria-controls=`"autopilotConfig-0`"><i class=`"fas fa-chevron-down me-2`"></i>Show Configuration Details</button>
<div class=`"collapse mt-3`" id=`"autopilotConfig-0`">
<div class=`"autopilot-settings`">
<dt>Device Name Template:</dt><dd>$apDeviceName</dd>
<dt>Language:</dt><dd>$apLanguage</dd>
<dt>Locale:</dt><dd>$apLocale</dd>
</div>
</div>
</div>
"@
    $espConfig = if ($espPolicies.Count -gt 0) { Get-EspConfigByDisplayName -DisplayName $espPolicies[0] } else { $null }
    $espDesc = if ($espConfig) { [System.Net.WebUtility]::HtmlEncode($espConfig.Description) } else { "-" }
    $espShowProgress = if ($espConfig) { [System.Net.WebUtility]::HtmlEncode($espConfig.ShowInstallationProgress) } else { "-" }
    $espRequiredAppsMode = "All"
    $espSelectedAppsHtml = ""
    $selectedAppIds = @()
    if ($espConfig -and $espConfig.SelectedMobileAppIds) {
        foreach ($aid in $espConfig.SelectedMobileAppIds) {
            $s = [string]$aid
            if ($s -and $s.Trim().Length -gt 0) { $selectedAppIds += $s.Trim() }
        }
    }
    if ($selectedAppIds.Count -gt 0) {
        $espRequiredAppsMode = "Selected"
        $appNames = @()
        foreach ($appId in $selectedAppIds) {
            $appNames += Get-MobileAppDisplayName -AppId $appId
        }
        $espSelectedAppsHtml = "<dt>Selected apps:</dt><dd>" + '<ul class="esp-app-list">' + (($appNames | ForEach-Object { "<li>" + [System.Net.WebUtility]::HtmlEncode($_) + "</li>" }) -join "") + "</ul></dd>"
    }
    $espRequiredAppsModeEnc = [System.Net.WebUtility]::HtmlEncode($espRequiredAppsMode)
    $espBlockDeviceUseHtml = "<dt>Block device use until required apps are installed if they are assigned to the user/device:</dt><dd>$espRequiredAppsModeEnc</dd>"
    if ($espRequiredAppsMode -eq "Selected") { $espBlockDeviceUseHtml += $espSelectedAppsHtml }
    $espTimeout = if ($espConfig) { [System.Net.WebUtility]::HtmlEncode($espConfig.InstallProgressTimeout) } else { "-" }
    $espShowCustomMessage = if ($espConfig -and $espConfig.ShowCustomMessageWhenError) { [System.Net.WebUtility]::HtmlEncode($espConfig.ShowCustomMessageWhenError) } else { "-" }
    $espCustomErrorMessage = if ($espConfig -and $espConfig.CustomErrorMessage) { [System.Net.WebUtility]::HtmlEncode($espConfig.CustomErrorMessage) } else { "-" }
    $espAllowLog = if ($espConfig) { [System.Net.WebUtility]::HtmlEncode($espConfig.AllowLogCollectionOnFailure) } else { "-" }
    $espTrackAutopilot = if ($espConfig) { [System.Net.WebUtility]::HtmlEncode($espConfig.TrackAutopilotOnly) } else { "-" }
    $espQuality = if ($espConfig) { [System.Net.WebUtility]::HtmlEncode($espConfig.InstallQualityUpdates) } else { "-" }
    $espBlockUntilAll = if ($espConfig -and $espConfig.ShowInstallationProgress) { [System.Net.WebUtility]::HtmlEncode($espConfig.ShowInstallationProgress) } else { "-" }
    $espAllowReset = if ($espConfig) { [System.Net.WebUtility]::HtmlEncode($espConfig.AllowDeviceResetOnFailure) } else { "-" }
    $espAllowUse = if ($espConfig) { [System.Net.WebUtility]::HtmlEncode($espConfig.AllowDeviceUseOnFailure) } else { "-" }
    $espAllowNonBlock = if ($espConfig) { [System.Net.WebUtility]::HtmlEncode($espConfig.AllowNonBlockingAppInstallation) } else { "-" }
    $espRestHtml = ""
    if ($espShowProgress -ne "No") {
        $espBlockDeviceUseHtmlSafe = $espBlockDeviceUseHtml -replace '"', '`"'
        $espRestHtml = @"
<dt>Show an error when installation takes longer than specified number of minutes</dt><dd>$espTimeout</dd>
<dt>Show custom message when time limit or error occurs</dt><dd>$espShowCustomMessage</dd>
<dt>Error message</dt><dd>$espCustomErrorMessage</dd>
<dt>Turn on log collection and diagnostics page for end users</dt><dd>$espAllowLog</dd>
<dt>Only show page to devices provisioned by out-of-box experience (OOBE)</dt><dd>$espTrackAutopilot</dd>
<dt>Install Windows updates (might restart the device)</dt><dd>$espQuality</dd>
<dt>Block device use until all apps and profiles are installed</dt><dd>$espBlockUntilAll</dd>
<dt>Allow users to reset device if installation error occurs</dt><dd>$espAllowReset</dd>
<dt>Allow users to use device if installation error occurs</dt><dd>$espAllowUse</dd>
<dt>Only fail selected blocking apps in technician phase</dt><dd>$espAllowNonBlock</dd>
$espBlockDeviceUseHtmlSafe
"@
    }
    $espConfigSection = @"
<div class=`"esp-config-section`">
<button class=`"btn btn-outline-secondary btn-sm config-toggle`" type=`"button`" data-bs-toggle=`"collapse`" data-bs-target=`"#espConfig-0`" aria-expanded=`"false`" aria-controls=`"espConfig-0`"><i class=`"fas fa-chevron-down me-2`"></i>Show Configuration Details</button>
<div class=`"collapse mt-3`" id=`"espConfig-0`">
<div class=`"autopilot-settings`">
<dt>Description</dt><dd>$espDesc</dd>
<dt>Show app and profile configuration progress</dt><dd>$espShowProgress</dd>
$espRestHtml
</div>
</div>
</div>
"@
    $cloudPCProvisioningConfigSection = ""
    if ($cloudPCProvisioningMerged.Count -gt 0) {
        $cloudPCPolicy = Get-CloudPCProvisioningConfigByDisplayName -DisplayName $cloudPCProvisioningMerged[0]
        if ($cloudPCPolicy) {
            $toYesNo = $script:ToYesNo
            $cpName = if ($cloudPCPolicy.displayName) { [System.Net.WebUtility]::HtmlEncode($cloudPCPolicy.displayName) } else { "-" }
            $cpExperience = if ($cloudPCPolicy.userExperienceType -eq 'cloudPc') { "Access a full Cloud PC desktop" } else { [System.Net.WebUtility]::HtmlEncode([string]$cloudPCPolicy.userExperienceType) }
            $cpLicense = if ($cloudPCPolicy.managedBy -eq 'windows365') { "Enterprise" } else { [System.Net.WebUtility]::HtmlEncode([string]$cloudPCPolicy.managedBy) }
            $cpSso = & $toYesNo $cloudPCPolicy.enableSingleSignOn
            $cpJoinType = "-"
            $cpGeo = "-"
            if ($cloudPCPolicy.domainJoinConfigurations -and @($cloudPCPolicy.domainJoinConfigurations).Count -gt 0) {
                $djc = @($cloudPCPolicy.domainJoinConfigurations)[0]
                $cpJoinType = if ($djc.type -eq 'azureADJoin') { "Microsoft Entra Join" } else { [System.Net.WebUtility]::HtmlEncode([string]$djc.type) }
                if ($djc.geographicLocationType) {
                    $cap = [string]$djc.geographicLocationType
                    $capEnc = if ($cap.Length -gt 0) { $cap.Substring(0, 1).ToUpper() + $cap.Substring(1) } else { $cap }
                    $cpGeo = [System.Net.WebUtility]::HtmlEncode($capEnc)
                }
            }
            $cpImageType = if ($cloudPCPolicy.imageType) {
                $it = [string]$cloudPCPolicy.imageType
                $itEnc = if ($it.Length -gt 0) { $it.Substring(0, 1).ToUpper() + $it.Substring(1) } else { $it }
                [System.Net.WebUtility]::HtmlEncode($itEnc)
            } else { "-" }
            $cpImageName = if ($cloudPCPolicy.imageDisplayName) { [System.Net.WebUtility]::HtmlEncode($cloudPCPolicy.imageDisplayName) } else { "-" }
            $cpLang = "-"
            if ($cloudPCPolicy.windowsSettings -and $cloudPCPolicy.windowsSettings.language) { $cpLang = [System.Net.WebUtility]::HtmlEncode($cloudPCPolicy.windowsSettings.language) }
            elseif ($cloudPCPolicy.windowsSetting -and $cloudPCPolicy.windowsSetting.locale) { $cpLang = [System.Net.WebUtility]::HtmlEncode($cloudPCPolicy.windowsSetting.locale) }
            $cpApplyNameTemplate = & $toYesNo ($cloudPCPolicy.cloudPcNamingTemplate -and [string]$cloudPCPolicy.cloudPcNamingTemplate.Trim().Length -gt 0)
            $cpNameTemplate = if ($cloudPCPolicy.cloudPcNamingTemplate) { [System.Net.WebUtility]::HtmlEncode($cloudPCPolicy.cloudPcNamingTemplate) } else { "-" }
            $cpScopeIds = if ($cloudPCPolicy.scopeIds -and @($cloudPCPolicy.scopeIds).Count -gt 0) {
                $scopeParts = $cloudPCPolicy.scopeIds | ForEach-Object { if ($_ -eq '0') { 'Default' } else { $_ } }
                [System.Net.WebUtility]::HtmlEncode(($scopeParts -join ', '))
            } else { "Default" }
            $cloudPCProvisioningConfigSection = @"
<div class=`"esp-config-section cloudpc-config-section`">
<button class=`"btn btn-outline-secondary btn-sm config-toggle`" type=`"button`" data-bs-toggle=`"collapse`" data-bs-target=`"#cloudpcConfig-0`" aria-expanded=`"false`" aria-controls=`"cloudpcConfig-0`"><i class=`"fas fa-chevron-down me-2`"></i>Show Configuration Details</button>
<div class=`"collapse mt-3`" id=`"cloudpcConfig-0`">
<div class=`"autopilot-settings`">
<h6 class=`"mb-2 mt-2`">General</h6>
<dt>Name</dt><dd>$cpName</dd>
<h6 class=`"mb-2 mt-2`">Experience</h6>
<dt>Experience type</dt><dd>$cpExperience</dd>
<dt>License type</dt><dd>$cpLicense</dd>
<dt>Use Microsoft Entra single sign-on</dt><dd>$cpSso</dd>
<dt>Join type</dt><dd>$cpJoinType</dd>
<h6 class=`"mb-2 mt-2`">Geography</h6>
<dt>Region</dt><dd>$cpGeo</dd>
<h6 class=`"mb-2 mt-2`">Image</h6>
<dt>Image type</dt><dd>$cpImageType</dd>
<dt>Image</dt><dd>$cpImageName</dd>
<h6 class=`"mb-2 mt-2`">Configuration</h6>
<dt>Language &amp; Region</dt><dd>$cpLang</dd>
<dt>Apply device name template</dt><dd>$cpApplyNameTemplate</dd>
<dt>Device name template</dt><dd>$cpNameTemplate</dd>
<h6 class=`"mb-2 mt-2`">Scope tags</h6>
<dt>Scope tags</dt><dd>$cpScopeIds</dd>
</div>
</div>
</div>
"@
        }
        if ($cloudPCProvisioningMerged.Count -gt 0) {
            $cloudPCProvisioningDeviceGroupsHtml = ""
            $cloudPCProvisioningDeviceGroupsCount = $cloudPCProvisioningMerged.Count
            $cloudPCTableHeader = '<div class="policy-table-header policy-table-header-devicegroups" style="grid-template-columns: 2fr 2fr 1fr 3fr;"><div class="header-policy-name">Policy name</div><div class="header-assignment">Group name</div><div class="header-type">Type</div><div class="header-filter">Membership rule</div></div>'
            foreach ($p in $cloudPCProvisioningMerged) {
                $gtr = Get-FirstAssignmentGroupTypeRule $p $EvaluatedAssignments
                $pEnc = [System.Net.WebUtility]::HtmlEncode($p)
                $gEnc = [System.Net.WebUtility]::HtmlEncode($gtr[0])
                $tEnc = [System.Net.WebUtility]::HtmlEncode($gtr[1])
                $rEnc = [System.Net.WebUtility]::HtmlEncode($gtr[2])
                $cloudPCProvisioningDeviceGroupsHtml += '<div class="policy-item policy-item-devicegroups" style="grid-template-columns: 2fr 2fr 1fr 3fr;">'
                $cloudPCProvisioningDeviceGroupsHtml += '<div class="policy-name"><span class="flow-field-label">Policy name</span><br><span class="flow-field-value">' + $pEnc + '</span></div>'
                $cloudPCProvisioningDeviceGroupsHtml += '<div class="policy-assignment"><span class="flow-field-label">Group name</span><br><span class="flow-field-value">' + $gEnc + '</span></div>'
                $cloudPCProvisioningDeviceGroupsHtml += '<div class="policy-type"><span class="flow-field-label">Type</span><br><span class="flow-field-value">' + $tEnc + '</span></div>'
                $cloudPCProvisioningDeviceGroupsHtml += '<div class="policy-filter entra-rule"><span class="flow-field-label">Membership rule</span><br><span class="flow-field-value">' + $rEnc + '</span></div>'
                $cloudPCProvisioningDeviceGroupsHtml += '</div>'
            }
            $cloudPCProvisioningDeviceGroupsHtml = $cloudPCTableHeader + $cloudPCProvisioningDeviceGroupsHtml
        }
        else {
            $cloudPCProvisioningDeviceGroupsCount = 0
            $cloudPCProvisioningDeviceGroupsHtml = '<div class="policy-table-header policy-table-header-devicegroups" style="grid-template-columns: 2fr 2fr 1fr 3fr;"><div class="header-policy-name">Policy name</div><div class="header-assignment">Group name</div><div class="header-type">Type</div><div class="header-filter">Membership rule</div></div><div class="policy-item policy-item-devicegroups policy-item-stacked" style="grid-template-columns: 2fr 2fr 1fr 3fr;"><div class="policy-name"><span class="flow-field-label">Policy name</span><br><span class="flow-field-value">- None -</span></div><div class="policy-assignment"><span class="flow-field-label">Group name</span><br><span class="flow-field-value">-</span></div><div class="policy-type"><span class="flow-field-label">Type</span><br><span class="flow-field-value">-</span></div><div class="policy-filter"><span class="flow-field-label">Membership rule</span><br><span class="flow-field-value">-</span></div></div>'
        }
    }
    
    # Cloud PC User Settings Configuration Section
    $cloudPCUserSettingsConfigSection = ""
    if ($cloudPCUserSettingsMerged.Count -gt 0) {
        $cloudPCUserSettings = Get-CloudPCUserSettingsConfigByDisplayName -DisplayName $cloudPCUserSettingsMerged[0]
        if ($cloudPCUserSettings) {
            $usName = if ($cloudPCUserSettings.Name) { [System.Net.WebUtility]::HtmlEncode($cloudPCUserSettings.Name) } else { "-" }
            $usSelfService = if ($cloudPCUserSettings.SelfServiceEnabled) { [System.Net.WebUtility]::HtmlEncode($cloudPCUserSettings.SelfServiceEnabled) } else { "-" }
            $usLocalAdmin = if ($cloudPCUserSettings.LocalAdminEnabled) { [System.Net.WebUtility]::HtmlEncode($cloudPCUserSettings.LocalAdminEnabled) } else { "-" }
            $usReset = if ($cloudPCUserSettings.ResetEnabled) { [System.Net.WebUtility]::HtmlEncode($cloudPCUserSettings.ResetEnabled) } else { "-" }
            $usRestoreFreq = if ($cloudPCUserSettings.RestorePointFrequency) { [System.Net.WebUtility]::HtmlEncode($cloudPCUserSettings.RestorePointFrequency) } else { "-" }
            $usUserRestore = if ($cloudPCUserSettings.UserRestoreEnabled) { [System.Net.WebUtility]::HtmlEncode($cloudPCUserSettings.UserRestoreEnabled) } else { "-" }
            $usDisasterRecovery = if ($cloudPCUserSettings.DisasterRecoveryEnabled) { [System.Net.WebUtility]::HtmlEncode($cloudPCUserSettings.DisasterRecoveryEnabled) } else { "-" }
            $usUserInitiatedDR = if ($cloudPCUserSettings.UserInitiatedDRAllowed) { [System.Net.WebUtility]::HtmlEncode($cloudPCUserSettings.UserInitiatedDRAllowed) } else { "-" }
            $usRestartPrompts = if ($cloudPCUserSettings.RestartPromptsDisabled) { [System.Net.WebUtility]::HtmlEncode($cloudPCUserSettings.RestartPromptsDisabled) } else { "-" }
            $cloudPCUserSettingsConfigSection = @"
<div class=`"esp-config-section cloudpc-usersettings-config-section`">
<button class=`"btn btn-outline-secondary btn-sm config-toggle`" type=`"button`" data-bs-toggle=`"collapse`" data-bs-target=`"#cloudpcUserSettingsConfig-0`" aria-expanded=`"false`" aria-controls=`"cloudpcUserSettingsConfig-0`"><i class=`"fas fa-chevron-down me-2`"></i>Show Configuration Details</button>
<div class=`"collapse mt-3`" id=`"cloudpcUserSettingsConfig-0`">
<div class=`"autopilot-settings`">
<h6 class=`"mb-2 mt-2`">General</h6>
<dt>Name</dt><dd>$usName</dd>
<h6 class=`"mb-2 mt-2`">User Settings</h6>
<dt>Self-service upgrades and resizes</dt><dd>$usSelfService</dd>
<dt>Local admin</dt><dd>$usLocalAdmin</dd>
<dt>Reset</dt><dd>$usReset</dd>
<h6 class=`"mb-2 mt-2`">System Restore Point</h6>
<dt>Frequency</dt><dd>$usRestoreFreq</dd>
<dt>User-initiated restore</dt><dd>$usUserRestore</dd>
<h6 class=`"mb-2 mt-2`">Cross Region Disaster Recovery</h6>
<dt>Disaster recovery enabled</dt><dd>$usDisasterRecovery</dd>
<dt>User-initiated disaster recovery</dt><dd>$usUserInitiatedDR</dd>
<h6 class=`"mb-2 mt-2`">Notifications</h6>
<dt>Restart prompts</dt><dd>$usRestartPrompts</dd>
</div>
</div>
</div>
"@
        }
    }
    
    if ($IsCloudPC) {
        if ($cloudPCUserSettingsMerged.Count -gt 0) {
            $cloudPCUserSettingsDeviceGroupsHtml = ""
            $cloudPCUserSettingsDeviceGroupsCount = $cloudPCUserSettingsMerged.Count
            $cloudPCUserSettingsTableHeader = '<div class="policy-table-header policy-table-header-devicegroups" style="grid-template-columns: 2fr 2fr 1fr 3fr;"><div class="header-policy-name">Policy name</div><div class="header-assignment">Group name</div><div class="header-type">Type</div><div class="header-filter">Membership rule</div></div>'
            foreach ($p in $cloudPCUserSettingsMerged) {
                $gtr = Get-FirstAssignmentGroupTypeRule $p $EvaluatedAssignments
                $pEnc = [System.Net.WebUtility]::HtmlEncode($p)
                $gEnc = [System.Net.WebUtility]::HtmlEncode($gtr[0])
                $tEnc = [System.Net.WebUtility]::HtmlEncode($gtr[1])
                $rEnc = [System.Net.WebUtility]::HtmlEncode($gtr[2])
                $cloudPCUserSettingsDeviceGroupsHtml += '<div class="policy-item policy-item-devicegroups" style="grid-template-columns: 2fr 2fr 1fr 3fr;">'
                $cloudPCUserSettingsDeviceGroupsHtml += '<div class="policy-name"><span class="flow-field-label">Policy name</span><br><span class="flow-field-value">' + $pEnc + '</span></div>'
                $cloudPCUserSettingsDeviceGroupsHtml += '<div class="policy-assignment"><span class="flow-field-label">Group name</span><br><span class="flow-field-value">' + $gEnc + '</span></div>'
                $cloudPCUserSettingsDeviceGroupsHtml += '<div class="policy-type"><span class="flow-field-label">Type</span><br><span class="flow-field-value">' + $tEnc + '</span></div>'
                $cloudPCUserSettingsDeviceGroupsHtml += '<div class="policy-filter entra-rule"><span class="flow-field-label">Membership rule</span><br><span class="flow-field-value">' + $rEnc + '</span></div>'
                $cloudPCUserSettingsDeviceGroupsHtml += '</div>'
            }
            $cloudPCUserSettingsDeviceGroupsHtml = $cloudPCUserSettingsTableHeader + $cloudPCUserSettingsDeviceGroupsHtml
        }
        else {
            $cloudPCUserSettingsDeviceGroupsCount = 0
            $cloudPCUserSettingsDeviceGroupsHtml = '<div class="policy-table-header policy-table-header-devicegroups" style="grid-template-columns: 2fr 2fr 1fr 3fr;"><div class="header-policy-name">Policy name</div><div class="header-assignment">Group name</div><div class="header-type">Type</div><div class="header-filter">Membership rule</div></div><div class="policy-item policy-item-devicegroups policy-item-stacked" style="grid-template-columns: 2fr 2fr 1fr 3fr;"><div class="policy-name"><span class="flow-field-label">Policy name</span><br><span class="flow-field-value">- None -</span></div><div class="policy-assignment"><span class="flow-field-label">Group name</span><br><span class="flow-field-value">-</span></div><div class="policy-type"><span class="flow-field-label">Type</span><br><span class="flow-field-value">-</span></div><div class="policy-filter"><span class="flow-field-label">Membership rule</span><br><span class="flow-field-value">-</span></div></div>'
        }
        $cloudPCDeviceGroupsCount = $cloudPCProvisioningDeviceGroupsCount
        $cloudPCDeviceGroupsHtml = $cloudPCProvisioningDeviceGroupsHtml
    }
    function Write-FlowStepTwoCol($stepId, $title, $iconClass, $deviceList, $userList, [array]$eval) {
        $t = [System.Net.WebUtility]::HtmlEncode($title)
        $dc = $deviceList.Count; $uc = $userList.Count
        $total = $dc + $uc
        $deviceTable = Write-FlowPolicyTable "Device assigned" "Policy Name" "Assignment" "Filter" $deviceList $eval -hideCategoryTitle $true
        $userTable = Write-FlowPolicyTable "User assigned" "Policy Name" "Assignment" "Filter" $userList $eval -hideCategoryTitle $true
        @"
<div class=`"flow-step flow-step-twocol $stepId`">
<div class=`"step-header`"><i class=`"fas $iconClass`"></i><span>$t</span></div>
<div class=`"step-content`">
<button class=`"btn btn-outline-primary btn-sm step-toggle`" type=`"button`" data-bs-toggle=`"collapse`" data-bs-target=`"#$stepId`" aria-expanded=`"true`" aria-controls=`"$stepId`"><i class=`"fas fa-chevron-down`"></i>Expand ($total)</button>
<div class=`"collapse show`" id=`"$stepId`">
<div class=`"row g-3`">
<div class=`"col-md-6 flow-col-device`">
<div class=`"flow-col-header`">Device assigned</div>
$deviceTable
</div>
<div class=`"col-md-6 flow-col-user`">
<div class=`"flow-col-header`">User assigned</div>
$userTable
</div>
</div>
</div>
</div>
</div>
"@
    }
    $entraCount = if ($DeviceGroupDetails) { $DeviceGroupDetails.Count } else { 0 }
    $entraGroupListHtml = ""
    $entraGroupTableHeader = '<div class="policy-table-header policy-table-header-devicegroups" style="grid-template-columns: 2fr 1fr 3fr;"><div class="header-policy-name">Group name</div><div class="header-assignment">Type</div><div class="header-filter">Membership rule</div></div>'
    if ($entraCount -gt 0) {
        foreach ($gd in $DeviceGroupDetails) {
            $name = if ($gd.DisplayName) { $gd.DisplayName } else { "-" }
            $gType = if ($gd.GroupType) { $gd.GroupType } else { "Static" }
            $ruleHtml = if ($gd.MembershipRule) { $gd.MembershipRule } else { "-" }
            $encName = [System.Net.WebUtility]::HtmlEncode($name)
            $encType = [System.Net.WebUtility]::HtmlEncode($gType)
            $entraGroupListHtml += '<div class="policy-item policy-item-devicegroups" style="grid-template-columns: 2fr 1fr 3fr;">'
            $entraGroupListHtml += '<div class="policy-name"><span class="flow-field-label">Group name</span><br><span class="flow-field-value">' + $encName + '</span></div>'
            $entraGroupListHtml += '<div class="policy-assignment"><span class="flow-field-label">Type</span><br><span class="flow-field-value">' + $encType + '</span></div>'
            $entraGroupListHtml += '<div class="policy-filter entra-rule"><span class="flow-field-label">Membership rule</span><br><span class="flow-field-value">' + $ruleHtml + '</span></div>'
            $entraGroupListHtml += '</div>'
        }
    }
    else {
        $entraGroupListHtml = '<div class="policy-item policy-item-devicegroups policy-item-stacked" style="grid-template-columns: 2fr 1fr 3fr;"><div class="policy-name"><span class="flow-field-label">Group name</span><br><span class="flow-field-value">- None -</span></div><div class="policy-assignment"><span class="flow-field-label">Type</span><br><span class="flow-field-value">-</span></div><div class="policy-filter"><span class="flow-field-label">Membership rule</span><br><span class="flow-field-value">-</span></div></div>'
    }
    $entraGroupListHtml = $entraGroupTableHeader + $entraGroupListHtml
    $entraStepHtml = @"
<div class=`"flow-step step-entra`">
<div class=`"step-header`"><i class=`"fas fa-users`"></i><span>Entra ID groups</span></div>
<div class=`"step-content`">
<button class=`"btn btn-outline-primary btn-sm step-toggle`" type=`"button`" data-bs-toggle=`"collapse`" data-bs-target=`"#flow-entra`" aria-expanded=`"true`" aria-controls=`"flow-entra`"><i class=`"fas fa-chevron-down`"></i>Device group membership</button>
<div class=`"collapse show`" id=`"flow-entra`">
<div class=`"policy-category`"><h5>Device groups <span class=`"policy-count`">$entraCount</span></h5>
$entraGroupListHtml
</div>
</div>
</div>
</div>
"@
    $cloudPCStepN = $cloudPCProvisioningMerged.Count
    $cloudPCStepHtml = @"
<div class=`"flow-step step-cloudpc-provisioning`">
<div class=`"step-header`"><i class=`"fas fa-cloud`"></i><span>Cloud PC Provisioning</span></div>
<div class=`"step-content`">
<button class=`"btn btn-outline-primary btn-sm step-toggle`" type=`"button`" data-bs-toggle=`"collapse`" data-bs-target=`"#flow-cloudpc-provisioning`" aria-expanded=`"true`" aria-controls=`"flow-cloudpc-provisioning`"><i class=`"fas fa-chevron-down`"></i>Expand Cloud PC Provisioning ($cloudPCStepN)</button>
<div class=`"collapse show`" id=`"flow-cloudpc-provisioning`">
<div class=`"policy-category`"><h5>Device groups <span class=`"policy-count`">$cloudPCDeviceGroupsCount</span></h5>
$cloudPCDeviceGroupsHtml
</div>
$cloudPCProvisioningConfigSection
</div>
</div>
</div>
"@
    $cloudPCUserSettingsStepN = $cloudPCUserSettingsMerged.Count
    $cloudPCUserSettingsStepHtml = @"
<div class=`"flow-step step-cloudpc-usersettings`">
<div class=`"step-header`"><i class=`"fas fa-user-cog`"></i><span>Cloud PC User Settings</span></div>
<div class=`"step-content`">
<button class=`"btn btn-outline-primary btn-sm step-toggle`" type=`"button`" data-bs-toggle=`"collapse`" data-bs-target=`"#flow-cloudpc-usersettings`" aria-expanded=`"true`" aria-controls=`"flow-cloudpc-usersettings`"><i class=`"fas fa-chevron-down`"></i>Expand Cloud PC User Settings ($cloudPCUserSettingsStepN)</button>
<div class=`"collapse show`" id=`"flow-cloudpc-usersettings`">
<div class=`"policy-category`"><h5>Device groups <span class=`"policy-count`">$cloudPCUserSettingsDeviceGroupsCount</span></h5>
$cloudPCUserSettingsDeviceGroupsHtml
</div>
$cloudPCUserSettingsConfigSection
</div>
</div>
</div>
"@
    $parts = @()
    $parts += '<div class="flow-container">'
    $parts += '<div class="flow-node flow-device"><div class="flow-device-label">Device</div><div class="flow-device-name">' + $deviceEnc + '</div></div>'
    $parts += '<div class="flow-arrow" aria-hidden="true"><i class="fas fa-chevron-down"></i></div>'
    $parts += $entraStepHtml
    $parts += '<div class="flow-arrow" aria-hidden="true"><i class="fas fa-chevron-down"></i></div>'
    if ($IsCloudPC) {
        $parts += $cloudPCUserSettingsStepHtml
        $parts += '<div class="flow-arrow" aria-hidden="true"><i class="fas fa-chevron-down"></i></div>'
        $parts += $cloudPCStepHtml
        $parts += '<div class="flow-arrow" aria-hidden="true"><i class="fas fa-chevron-down"></i></div>'
    }
    $parts += Write-FlowStepLikeReference "flow-autopilot" "step-autopilot" "Autopilot Profile" "fa-rocket" $autopilotPolicies $EvaluatedAssignments "Profile" "Autopilot Assignment" "Autopilot Profile" $autopilotConfigSection
    $parts += '<div class="flow-arrow" aria-hidden="true"><i class="fas fa-chevron-down"></i></div>'
    $parts += Write-FlowStepLikeReference "flow-esp" "step-esp" "Enrollment Status Page" "fa-shield-alt" $espPolicies $EvaluatedAssignments "ESP" "ESP Assignment" "Enrollment Status Page" $espConfigSection
    $parts += '<div class="flow-arrow" aria-hidden="true"><i class="fas fa-chevron-down"></i></div>'
    $parts += Write-FlowStepTwoCol "flow-config" "Configuration Profiles" "fa-cogs" $configDevicePolicies $configUserPolicies $EvaluatedAssignments
    $parts += '<div class="flow-arrow" aria-hidden="true"><i class="fas fa-chevron-down"></i></div>'
    $parts += Write-FlowStepTwoCol "flow-compliance" "Compliance policies" "fa-check-circle" $complianceDevicePolicies $complianceUserPolicies $EvaluatedAssignments
    $parts += '<div class="flow-arrow" aria-hidden="true"><i class="fas fa-chevron-down"></i></div>'
    $parts += Write-FlowStepTwoCol "flow-apps" "Apps" "fa-box" $appDevicePolicies $appUserPolicies $EvaluatedAssignments
    $parts += '</div>'
    '<div class="flow-diagram">' + ($parts -join '') + '</div>'
}

function New-DeviceVisualizationHtmlReport {
    param(
        [Parameter(Mandatory)][array]$EvaluatedAssignments,
        [Parameter(Mandatory)][string]$DeviceName,
        [string]$TenantName = "Intune Tenant",
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory = $false)][string]$AssignmentOverviewFragment = "",
        [Parameter(Mandatory = $false)][string]$AppliedFlowHtml = "",
        [Parameter(Mandatory = $false)][array]$DeviceGroupDetails = @(),
        [Parameter(Mandatory = $false)][string]$IntuneDeviceId = "",
        [Parameter(Mandatory = $false)][string]$EntraDeviceId = "",
        [Parameter(Mandatory = $false)][string]$DevicePlatform = $null,
        [Parameter(Mandatory = $false)][switch]$IsCloudPC = $false,
        [Parameter(Mandatory = $false)][array]$DeviceGroupIds = @(),
        [Parameter(Mandatory = $false)][array]$UserGroupIds = @()
    )
    if (-not $AppliedFlowHtml) { $AppliedFlowHtml = Get-AppliedFlowHtmlFragment -EvaluatedAssignments $EvaluatedAssignments -DeviceName $DeviceName -DeviceGroupDetails $DeviceGroupDetails -DevicePlatform $DevicePlatform -IsCloudPC:$IsCloudPC -DeviceGroupIds $DeviceGroupIds -UserGroupIds $UserGroupIds }
    $architectureFragment = Get-ArchitectureDiagramFragment -EvaluatedAssignments $EvaluatedAssignments -DeviceName $DeviceName -DeviceGroupDetails $DeviceGroupDetails -IntuneDeviceId $IntuneDeviceId -EntraDeviceId $EntraDeviceId -DevicePlatform $DevicePlatform -IsCloudPC:$IsCloudPC -DeviceGroupIds $DeviceGroupIds -UserGroupIds $UserGroupIds
    $overviewTabNav = ""
    $overviewTabPane = ""
    $diagramTabActive = "active"
    $diagramPaneActive = "show active"
    $architectureTabNav = ""
    if ($AssignmentOverviewFragment) {
        $overviewTabNav = "<li class=`"nav-item`"><button class=`"nav-link active`" id=`"overview-tab`" data-bs-toggle=`"tab`" data-bs-target=`"#overview`" type=`"button`" role=`"tab`"><i class=`"fas fa-chart-pie me-2`"></i>Assignment Overview</button></li>"
        $overviewTabPane = "<div class=`"tab-pane fade show active`" id=`"overview`" role=`"tabpanel`">" + $AssignmentOverviewFragment + "</div>"
        $diagramTabActive = ""
        $diagramPaneActive = ""
    }
    $reportDate = Get-Date -Format 'MMMM dd, yyyy HH:mm'
    $titleDevice = [System.Net.WebUtility]::HtmlEncode($DeviceName)
    $bannerDevice = [System.Net.WebUtility]::HtmlEncode($DeviceName)
    $tenantDisplay = [System.Net.WebUtility]::HtmlEncode($TenantName) + " · " + $reportDate
    $htmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Intune Device Visualization - __PH_TITLE_DEVICE__</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/5.3.0/css/bootstrap.min.css">
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css">
<link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.4.1/css/buttons.bootstrap5.min.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<style>
:root{--primary-color:#0078d4;--bg-color:#e2e8f0;--card-bg:#eef2f7;--text-color:#1e293b;--text-secondary:#64748b;--border-color:#cbd5e1;--gradient-primary:linear-gradient(135deg,#0078d4 0%,#107c10 100%);}
[data-theme="dark"]{--bg-color:#0d0d0d;--card-bg:#1a1a1a;--text-color:#e5e5e5;--text-secondary:#a3a3a3;--border-color:#404040;--bg-secondary:#1a1a1a;--bg-tertiary:#262626;}
body{background:var(--bg-color);color:var(--text-color);min-height:100vh;}
.app-container{max-width:1200px;margin:0 auto;padding:20px;}
.main-tabs-nav{display:flex;width:100%;}
.main-tabs-nav .nav-item{flex:1;text-align:center;}
.main-tabs-nav .nav-link{width:100%;text-align:center;}
.dashboard-header{background:linear-gradient(180deg,#1e293b 0%,#334155 100%);color:#fff;padding:2rem;text-align:center;box-shadow:0 4px 24px rgba(0,0,0,0.08);border-bottom:3px solid #475569;}
.overview-tiles-row .summary-card{min-height:100px;padding:0.75rem 1rem;display:flex;flex-direction:column;justify-content:center;align-items:center;}
.overview-tiles-row .summary-card .card-icon{width:36px;height:36px;font-size:1rem;margin:0 auto 0.4rem;}
.overview-tiles-row .summary-card .card-title{font-size:1.5rem;margin:0.2rem 0;}
.overview-tiles-row .summary-card .card-text{min-height:2em;font-size:0.8rem;display:flex;align-items:center;justify-content:center;text-align:center;}
.device-banner{background:var(--card-bg);border:1px solid var(--border-color,#e5e7eb);border-radius:8px;padding:1rem;margin-bottom:1rem;}
.summary-card{background:var(--card-bg);border-radius:12px;padding:1.5rem;text-align:center;}
.modern-table-container{margin:1rem 0;border-radius:12px;overflow:hidden;}
.modern-table-header{background:linear-gradient(180deg,#1e293b 0%,#334155 100%);padding:1rem;color:#fff;}
.modern-table-header h5,.modern-table-header small{color:#fff;}
.modern-table-header small{opacity:0.88;font-size:0.875rem;}
.nav-tabs .nav-link{font-weight:500;}
.nav-tabs .nav-link.active{background:#334155;color:#fff;border-color:#334155;}
[data-theme="dark"] .nav-tabs .nav-link{background:var(--bg-tertiary);color:var(--text-secondary);border-color:var(--border-color);}
[data-theme="dark"] .nav-tabs .nav-link:hover{background:#333;color:var(--text-color);border-color:var(--border-color);}
[data-theme="dark"] .nav-tabs .nav-link.active{background:#333;color:var(--text-color);border-color:var(--border-color);}
[data-theme="dark"] .table,[data-theme="dark"] .modern-table{background:var(--card-bg);color:var(--text-color);border-color:var(--border-color);}
[data-theme="dark"] .table thead th,[data-theme="dark"] .modern-table thead th{background:var(--bg-tertiary);color:var(--text-color);border-color:var(--border-color);}
[data-theme="dark"] .table tbody td,[data-theme="dark"] .modern-table tbody td{background:var(--card-bg);color:var(--text-color);border-color:var(--border-color);}
[data-theme="dark"] .table tbody tr:hover td,[data-theme="dark"] .modern-table tbody tr:hover td{background:rgba(255,255,255,0.05);}
[data-theme="dark"] .table-striped tbody tr:nth-of-type(odd) td{background:var(--bg-tertiary);}
[data-theme="dark"] .dataTables_wrapper{color:var(--text-color);background:var(--card-bg);border-color:var(--border-color);}
[data-theme="dark"] .modern-table-container .table-responsive,[data-theme="dark"] .modern-table-body{background:var(--card-bg);}
[data-theme="dark"] .dataTables_filter input,[data-theme="dark"] .dataTables_length select{background:var(--bg-tertiary);color:var(--text-color);border-color:var(--border-color);}
[data-theme="dark"] .dataTables_info{color:var(--text-secondary);}
[data-theme="dark"] .dataTables_paginate .page-link{background:var(--card-bg);color:var(--text-color);border-color:var(--border-color);}
[data-theme="dark"] .dataTables_paginate .page-link:hover{background:var(--bg-tertiary);color:var(--text-color);}
[data-theme="dark"] .dataTables_paginate .page-item.active .page-link{background:var(--bg-tertiary);border-color:var(--border-color);color:var(--text-color);}
[data-theme="dark"] .dataTables_paginate .page-item.disabled .page-link{background:var(--card-bg);color:var(--text-secondary);}
[data-theme="dark"] .dt-button{background:var(--bg-tertiary);color:var(--text-color);border-color:var(--border-color);}
[data-theme="dark"] .dt-button:hover{background:#333;color:var(--text-color);}
[data-theme="dark"] .modern-table-container{background:var(--card-bg);border-color:var(--border-color);}
[data-theme="dark"] .modern-table-header{background:var(--bg-tertiary);color:var(--text-color);}
[data-theme="dark"] .device-banner{border-color:var(--border-color);}
.alert-info{background-color:#e7f3ff;border-color:#b3d9ff;color:#004085;}
.alert-info code{background-color:rgba(0,0,0,0.05);padding:2px 6px;border-radius:3px;font-family:'Cascadia Code','Source Code Pro',Menlo,Consolas,monospace;}
.alert-info a{color:#004085;font-weight:600;}
.alert-info a:hover{color:#002752;}
[data-theme="dark"] .alert-info{background-color:#1a3a52;border-color:#2c6089;color:#a8d1ff;}
[data-theme="dark"] .alert-info code{background-color:rgba(255,255,255,0.1);}
[data-theme="dark"] .alert-info a{color:#a8d1ff;}
[data-theme="dark"] .alert-info a:hover{color:#d4e9ff;}
[data-theme="dark"] .dashboard-header{background:linear-gradient(180deg,#171717 0%,#262626 100%);border-bottom-color:var(--border-color);box-shadow:none;}
.assignment-filters-modern{background:linear-gradient(180deg,var(--card-bg) 0%,#e8ecf1 100%);border-bottom:1px solid var(--border-color);padding:1rem 1.25rem;}
.assignment-filters-modern .assignment-filters-inner{display:flex;flex-wrap:wrap;align-items:flex-end;gap:1rem;}
.assignment-filters-modern .filter-group{display:flex;flex-direction:column;gap:0.25rem;}
.assignment-filters-modern .filter-label{font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.03em;color:var(--text-secondary,#64748b);margin:0;}
.assignment-filters-modern .filter-select{min-width:140px;padding:0.5rem 0.75rem;border-radius:8px;border:1px solid var(--border-color);background:var(--card-bg);font-size:0.875rem;color:var(--text-color);transition:border-color 0.2s,box-shadow 0.2s;}
.assignment-filters-modern .filter-select:focus{border-color:#475569;outline:none;box-shadow:0 0 0 3px rgba(71,85,105,0.2);}
.assignment-filters-modern .filter-dropdown-btn{min-width:140px;text-align:left;}
.assignment-filters-modern .filter-checkbox-dropdown{max-height:260px;overflow-y:auto;padding:0.25rem;}
.assignment-filters-modern .filter-checkbox-dropdown .dropdown-item{white-space:nowrap;cursor:pointer;display:flex;align-items:center;gap:0.5rem;}
.assignment-filters-modern .filter-checkbox-dropdown .dropdown-item input{margin:0;cursor:pointer;}
.assignment-filters-modern .filter-checkbox-dropdown label{margin:0;cursor:pointer;width:100%;}
.assignment-filters-modern .filter-reset-btn{display:inline-flex;align-items:center;gap:0.35rem;padding:0.5rem 0.9rem;border-radius:8px;border:1px solid var(--border-color);background:var(--card-bg);font-size:0.875rem;color:var(--text-secondary);cursor:pointer;transition:all 0.2s;height:2.15rem;}
.assignment-filters-modern .filter-reset-btn:hover{background:#e2e8f0;border-color:#cbd5e1;color:#475569;}
[data-theme="dark"] .assignment-filters-modern{background:linear-gradient(180deg,var(--bg-tertiary) 0%,var(--card-bg) 100%);border-color:var(--border-color);}
[data-theme="dark"] .assignment-filters-modern .filter-label{color:var(--text-secondary);}
[data-theme="dark"] .assignment-filters-modern .filter-select{background:var(--card-bg);color:var(--text-color);border-color:var(--border-color);}
[data-theme="dark"] .assignment-filters-modern .filter-select:focus{border-color:#737373;box-shadow:0 0 0 3px rgba(115,115,115,0.2);}
[data-theme="dark"] .assignment-filters-modern .filter-reset-btn{background:var(--bg-tertiary);border-color:var(--border-color);color:var(--text-secondary);}
[data-theme="dark"] .assignment-filters-modern .filter-reset-btn:hover{background:#333;color:var(--text-color);}
[data-theme="dark"] .assignment-filters-modern .filter-dropdown-btn{background:var(--card-bg);border-color:var(--border-color);color:var(--text-color);}
[data-theme="dark"] .assignment-filters-modern .filter-dropdown-btn:hover,[data-theme="dark"] .assignment-filters-modern .filter-dropdown-btn:focus,[data-theme="dark"] .assignment-filters-modern .filter-dropdown-btn.show{background:var(--bg-tertiary);border-color:var(--border-color);color:var(--text-color);}
[data-theme="dark"] .assignment-filters-modern .filter-checkbox-dropdown{background:var(--card-bg);border-color:var(--border-color);}
[data-theme="dark"] .assignment-filters-modern .filter-checkbox-dropdown .dropdown-item{color:var(--text-color);}
[data-theme="dark"] .assignment-filters-modern .filter-checkbox-dropdown .dropdown-item:hover{background:var(--bg-tertiary);color:var(--text-color);}
[data-theme="dark"] .assignment-filters-modern .filter-checkbox-dropdown label{color:var(--text-color);}
[data-theme="dark"] .assignment-filters-modern .filter-checkbox-dropdown input.filter-cb{accent-color:#a3a3a3;}
[data-theme="dark"] .summary-card{background:var(--card-bg);border-color:var(--border-color);color:var(--text-color) !important;}
[data-theme="dark"] .summary-card .card-title,[data-theme="dark"] .summary-card div.card-title{color:var(--text-color) !important;}
[data-theme="dark"] .summary-card .text-muted,[data-theme="dark"] .summary-card p.text-muted{color:var(--text-secondary) !important;}
[data-theme="dark"] .text-muted{color:var(--text-secondary);}
[data-theme="dark"] .row .summary-card,.row .summary-card *{color:inherit;}
[data-theme="dark"] .row .summary-card .card-title{color:var(--text-color) !important;}
[data-theme="dark"] .row .summary-card .text-muted,.row .summary-card p{color:var(--text-secondary) !important;}
html[data-theme="dark"] #viz .summary-card,html[data-theme="dark"] #viz .summary-card .card-title,html[data-theme="dark"] #viz .summary-card div,html[data-theme="dark"] #viz .summary-card p{color:var(--text-color) !important;}
html[data-theme="dark"] #viz .summary-card p.text-muted,html[data-theme="dark"] #viz .summary-card .text-muted{color:var(--text-secondary) !important;}
[data-theme="dark"] .overview-container{background:linear-gradient(135deg,var(--bg-color) 0%,var(--card-bg) 100%);border-color:var(--border-color);}
[data-theme="dark"] .overview-header{background:var(--bg-tertiary);color:var(--text-color);}
[data-theme="dark"] .overview-header h2{color:var(--text-color);}
[data-theme="dark"] .overview-header p{color:var(--text-secondary);opacity:1;}
.overview-container{background:linear-gradient(135deg,var(--card-bg) 0%,#e8ecf1 100%);border-radius:16px;padding:2rem;margin:1rem 0;}
.overview-header{text-align:center;margin-bottom:2rem;background:linear-gradient(180deg,#1e293b 0%,#334155 100%);color:#fff;margin-left:-2rem;margin-right:-2rem;margin-top:-2rem;padding:2rem 2rem 1.5rem;border-radius:16px 16px 0 0;border-bottom:2px solid rgba(255,255,255,0.06);}
.overview-header h2{font-size:1.75rem;font-weight:700;margin-bottom:0.5rem;color:#fff;letter-spacing:-0.02em;}
.overview-header p{font-size:1rem;color:rgba(255,255,255,0.9);margin:0;}
.summary-card .card-icon{width:50px;height:50px;border-radius:12px;display:flex;align-items:center;justify-content:center;margin:0 auto 0.75rem;font-size:1.25rem;color:#fff;}
.summary-card.border-primary .card-icon{background:linear-gradient(135deg,#0078d4,#106ebe);}
.summary-card.border-success .card-icon{background:linear-gradient(135deg,#107c10,#0e6e0e);}
.summary-card.border-warning .card-icon{background:linear-gradient(135deg,#ffc107,#e0a800);}
.summary-card.border-danger .card-icon{background:linear-gradient(135deg,#d83b01,#c13401);}
.device-theme-toggle{position:fixed;top:10px;right:10px;z-index:1050;display:flex;align-items:center;gap:6px;background:var(--card-bg);padding:6px 10px;border-radius:30px;box-shadow:0 4px 12px rgba(0,0,0,0.1);border:1px solid var(--border-color,#e5e7eb);}
.device-theme-toggle .theme-icon{font-size:12px;color:var(--text-color);}
.device-theme-switch{position:relative;display:inline-block;width:40px;height:20px;}
.device-theme-switch input{opacity:0;width:0;height:0;}
.device-theme-slider{position:absolute;cursor:pointer;top:0;left:0;right:0;bottom:0;background:#cbd5e1;transition:.4s;border-radius:20px;}
.device-theme-slider:before{position:absolute;content:"";height:16px;width:16px;left:2px;bottom:2px;background:#fff;transition:.4s;border-radius:50%;box-shadow:0 1px 2px rgba(0,0,0,0.2);}
.device-theme-switch input:checked + .device-theme-slider{background:#334155;}
.device-theme-switch input:checked + .device-theme-slider:before{transform:translateX(20px);}
[data-theme="dark"] .device-theme-toggle{border-color:var(--border-color);}
.flow-diagram{display:flex;flex-direction:column;flex-wrap:nowrap;align-items:stretch;gap:1rem;padding:1.5rem;font-size:1rem;line-height:1.5;width:100%;max-width:100%;}
.flow-container{width:100%;max-width:100%;}
.flow-node{flex-shrink:0;}
.flow-device{background:linear-gradient(180deg,#0f766e 0%,#115e59 100%);color:#fff;padding:1rem 1.5rem;border-radius:12px;min-width:160px;max-width:100%;align-self:center;text-align:left;}
.flow-arrow{color:#0d9488;font-size:1.5rem;flex-shrink:0;align-self:center;display:flex;justify-content:center;}
.flow-device-label{display:block;font-size:0.75rem;text-transform:uppercase;letter-spacing:0.05em;opacity:0.9;margin-bottom:0.35rem;}
.flow-device-name{display:block;font-weight:600;font-size:1.1rem;word-break:break-word;}
.flow-step{background:var(--card-bg);border:1px solid var(--border-color,#e5e7eb);border-radius:12px;margin-bottom:1rem;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.06);}
.flow-step .step-header{display:flex;align-items:center;gap:0.5rem;padding:0.75rem 1rem;font-weight:600;font-size:1rem;color:#fff;background:linear-gradient(180deg,#0f766e 0%,#115e59 100%);}
.flow-step .step-content{padding:1rem;}
.flow-step .step-toggle{display:inline-flex;align-items:center;gap:0.4rem;margin-bottom:0.75rem;}
.flow-step .step-toggle i{transition:transform 0.2s ease;}
.flow-step .step-toggle[aria-expanded="true"] i{transform:rotate(180deg);}
.step-name{font-weight:600;font-size:0.9rem;margin-bottom:0.5rem;color:var(--text-color);}
.policy-category{background:var(--card-bg);border:1px solid var(--border-color,#e5e7eb);border-radius:8px;padding:1rem;margin-bottom:0.75rem;}
.policy-category h5{color:var(--text-color);margin-bottom:0.6rem;font-size:0.9rem;font-weight:600;display:flex;justify-content:space-between;align-items:center;}
.policy-count{background:#0f766e;color:#fff;padding:4px 10px;border-radius:20px;font-size:0.7rem;font-weight:600;}
.policy-table-header{display:grid;gap:0.75rem;padding:0.75rem 1rem;background:var(--bg-color);border:1px solid var(--border-color,#e5e7eb);border-radius:8px;margin-bottom:0.5rem;font-size:0.7rem;font-weight:700;color:var(--text-color);text-transform:uppercase;letter-spacing:0.5px;}
.policy-item{display:grid;gap:0.75rem;padding:0.6rem 1rem;background:var(--card-bg);border:1px solid var(--border-color,#e5e7eb);border-radius:6px;margin-bottom:0.5rem;font-size:0.8rem;}
.policy-item:hover{background:rgba(0,0,0,0.02);}
.policy-item-stacked{display:flex;flex-direction:column;gap:0.75rem;grid-template-columns:none !important;}
.policy-item-stacked .flow-field{margin-bottom:0.25rem;}
.policy-item-stacked .flow-field:last-child{margin-bottom:0;}
.flow-field-label{font-size:0.7rem;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;color:var(--text-secondary,#6c757d);margin-bottom:0.2rem;}
.flow-field-value{font-size:0.9rem;color:var(--text-color);word-wrap:break-word;overflow-wrap:break-word;}
.flow-field-value.entra-rule{font-family:ui-monospace,'Cascadia Code','Source Code Pro',Menlo,Consolas,monospace;font-size:0.75rem;word-break:break-all;background:rgba(0,0,0,0.06);padding:0.4rem 0.5rem;border-radius:4px;line-height:1.4;}
.flow-field-value.entra-rule code{background:rgba(0,0,0,0.1);padding:0.15rem 0.4rem;border-radius:3px;font-size:0.7rem;}
.policy-name{font-weight:600;color:var(--text-color);word-wrap:break-word;overflow-wrap:break-word;}
.policy-assignment{font-weight:500;color:#0f766e;line-height:1.4;}
.policy-filter{font-size:0.8rem;color:var(--text-secondary,#6c757d);line-height:1.4;word-wrap:break-word;}
.entra-group-row .entra-rule{font-family:ui-monospace,monospace;font-size:0.75rem;word-break:break-all;}
.entra-group-row .entra-rule code{background:rgba(0,0,0,0.06);padding:0.15rem 0.4rem;border-radius:4px;font-size:0.75rem;}
[data-theme="dark"] .entra-group-row .entra-rule code{background:rgba(255,255,255,0.1);}
.autopilot-config-section,.esp-config-section{margin-top:0.75rem;}
.config-toggle{font-size:0.8rem;border-color:var(--border-color,#e5e7eb);color:var(--text-color);}
.config-toggle:hover{background:#0f766e;border-color:#0f766e;color:#fff;}
.config-toggle i{transition:transform 0.2s ease;}
.config-toggle[aria-expanded="true"] i{transform:rotate(180deg);}
.config-toggle[aria-expanded="true"]{background:#0f766e;border-color:#0f766e;color:#fff;}
.flow-step .step-toggle.btn-outline-primary,.flow-step .btn-outline-primary.step-toggle{border-color:#0f766e;color:#0f766e;}
.flow-step .step-toggle.btn-outline-primary:hover,.flow-step .btn-outline-primary.step-toggle:hover{background:#0f766e;border-color:#0f766e;color:#fff;}
.autopilot-settings{background:rgba(0,0,0,0.03);border:1px solid var(--border-color,#e5e7eb);border-radius:6px;padding:0.75rem 1rem;margin-top:0.5rem;font-size:0.85rem;}
.autopilot-settings dt{font-weight:600;color:var(--text-color);margin:0.4rem 0 0.1rem 0;}
.autopilot-settings dd{margin:0 0 0.25rem 0;color:var(--text-secondary,#6c757d);}
.autopilot-settings dd:last-child{margin-bottom:0;}
.esp-app-list{margin:0.25rem 0 0 0;padding-left:1.25rem;}
.esp-app-list li{margin-bottom:0.15rem;}
.flow-step-twocol .flow-col-header{font-weight:600;font-size:0.9rem;margin-bottom:0.5rem;color:var(--text-color);}
.flow-step-twocol .flow-col-device,.flow-step-twocol .flow-col-user{background:rgba(0,0,0,0.02);border-radius:8px;padding:1rem;border:1px solid var(--border-color,#e5e7eb);}
[data-theme="dark"] .flow-step-twocol .flow-col-device,[data-theme="dark"] .flow-step-twocol .flow-col-user{background:rgba(255,255,255,0.04);border-color:var(--border-color);}
[data-theme="dark"] .flow-step{background:var(--card-bg);border-color:var(--border-color);}
[data-theme="dark"] .flow-step .step-header{background:var(--bg-tertiary);color:var(--text-color);}
[data-theme="dark"] .flow-device{background:var(--bg-tertiary);color:var(--text-color);}
[data-theme="dark"] .flow-arrow{color:var(--text-secondary);}
[data-theme="dark"] .flow-step-twocol .flow-col-header{color:var(--text-color);}
[data-theme="dark"] .step-name{color:var(--text-color);}
[data-theme="dark"] .policy-category{background:var(--card-bg);border-color:var(--border-color);}
[data-theme="dark"] .policy-category h5{color:var(--text-color);}
[data-theme="dark"] .policy-table-header{background:var(--bg-tertiary);border-color:var(--border-color);color:var(--text-color);}
[data-theme="dark"] .policy-item{background:var(--bg-tertiary);border-color:var(--border-color);}
[data-theme="dark"] .policy-item:hover{background:rgba(255,255,255,0.05);}
[data-theme="dark"] .policy-name{color:var(--text-color);}
[data-theme="dark"] .policy-assignment{color:#a3a3a3;}
[data-theme="dark"] .policy-filter{color:var(--text-secondary);}
[data-theme="dark"] .flow-field-label{color:var(--text-secondary);}
[data-theme="dark"] .flow-field-value{color:var(--text-color);}
[data-theme="dark"] .flow-field-value.entra-rule{background:rgba(255,255,255,0.05);}
[data-theme="dark"] .flow-field-value.entra-rule code{background:rgba(255,255,255,0.1);}
[data-theme="dark"] .policy-empty .policy-empty-msg,[data-theme="dark"] .policy-empty-msg.text-muted{color:var(--text-secondary) !important;}
[data-theme="dark"] .autopilot-settings{background:rgba(0,0,0,0.2);border-color:var(--border-color);}
[data-theme="dark"] .autopilot-settings dt{color:var(--text-color);}
[data-theme="dark"] .autopilot-settings dd{color:var(--text-secondary);}
.flow-diagram-wrapper{overflow-y:auto;overflow-x:auto;padding:0.5rem 0;}
@media (max-width:576px){.flow-step-twocol .row{flex-direction:column;}}
.arch-diagram-wrapper{background:linear-gradient(180deg,var(--bg-color) 0%,#cbd5e1 100%);border:1px solid var(--border-color);border-radius:16px;padding:2rem;margin:1rem 0;overflow:auto;min-height:520px;}
.arch-flow-horizontal .arch-diagram-wrapper{min-height:auto;padding:1rem 1rem;}
.arch-diagram-svg{width:100%;height:auto;display:block;}
.arch-diagram-svg .arch-box{filter:drop-shadow(0 2px 6px rgba(0,0,0,0.15));}
.arch-diagram-svg .arch-box:hover{filter:drop-shadow(0 4px 12px rgba(0,0,0,0.2));}
[data-theme="dark"] .arch-diagram-wrapper{background:linear-gradient(180deg,var(--bg-color) 0%,var(--card-bg) 100%);border-color:var(--border-color);}
[data-theme="dark"] .arch-diagram-svg .arch-box{filter:drop-shadow(0 2px 8px rgba(0,0,0,0.4));}
#architecture h4,#architecture h4 i{color:var(--text-color) !important;}
#architecture .text-muted,#architecture p.text-muted{color:var(--text-secondary) !important;}
[data-theme="dark"] #diagram .p-4 > p.text-muted{color:var(--text-color) !important;}
#architecture .p-4{color:var(--text-color);}
.arch-diagram-html .arch-flow{display:flex;flex-direction:column;align-items:center;gap:0;max-width:860px;margin:0 auto;}
.arch-flow-horizontal .arch-flow{flex-direction:row;align-items:flex-start;flex-wrap:nowrap;gap:0;max-width:100%;margin:0;padding:0.5rem 0;overflow-x:auto;overflow-y:visible;justify-content:flex-start;}
.arch-flow-horizontal .arch-step{flex:0 0 auto;min-width:90px;max-width:140px;padding:0.5rem 0.35rem;}
.arch-flow-horizontal .arch-step-device,.arch-flow-horizontal .arch-step-memberof{flex:1 1 0;min-width:180px;max-width:280px;padding:0.5rem 0.35rem;}
.arch-flow-horizontal .arch-step-enrollment{flex:0 0 auto;min-width:120px;max-width:180px;}
.arch-flow-horizontal .arch-arrow{display:flex;align-items:center;justify-content:center;padding:0 0.2rem;flex-shrink:0;font-size:1rem;align-self:flex-start;margin-top:1.25rem;}
.arch-flow-horizontal .arch-step-autopilot-esp + .arch-arrow{margin-top:6.5rem;}
.arch-flow-horizontal .arch-diagram-wrapper{min-height:auto;}
.arch-flow-horizontal .arch-step-title{font-size:0.85rem;margin-bottom:0.4rem;}
.arch-flow-horizontal .arch-step-sub{font-size:0.78rem;}
.arch-flow-horizontal .arch-group-table-wrap{margin-top:0.4rem;}
.arch-flow-horizontal .arch-diagram-html .arch-group-table{font-size:0.75rem;display:block;}
.arch-flow-horizontal .arch-diagram-html .arch-group-table thead{display:none;}
.arch-flow-horizontal .arch-diagram-html .arch-group-table tbody{display:block;}
.arch-flow-horizontal .arch-diagram-html .arch-group-table tbody tr{display:block;border-bottom:1px solid rgba(255,255,255,0.15);padding:0.25rem 0;}
.arch-flow-horizontal .arch-diagram-html .arch-group-table tbody tr:last-child{border-bottom:none;}
.arch-flow-horizontal .arch-diagram-html .arch-group-table td{display:block;padding:0.2rem 0.4rem;font-size:0.72rem;word-break:break-all;}
.arch-flow-horizontal .arch-diagram-html .arch-step-device .arch-group-table td,.arch-flow-horizontal .arch-diagram-html .arch-step-memberof .arch-group-table td{display:block;}
.arch-flow-horizontal .arch-diagram-html .arch-step-device .arch-group-table td::before,.arch-flow-horizontal .arch-diagram-html .arch-step-memberof .arch-group-table td::before{display:block;white-space:nowrap;margin-bottom:0.15rem;font-weight:700;}
.arch-flow-horizontal .arch-diagram-html .arch-step-device .arch-group-table td:nth-child(1)::before{content:"Name";opacity:0.9;}
.arch-flow-horizontal .arch-diagram-html .arch-step-device .arch-group-table td:nth-child(2)::before{content:"Intune ID";opacity:0.9;}
.arch-flow-horizontal .arch-diagram-html .arch-step-device .arch-group-table td:nth-child(3)::before{content:"Entra ID";opacity:0.9;}
.arch-flow-horizontal .arch-diagram-html .arch-step-memberof .arch-group-table td:nth-child(1)::before{content:"Group";opacity:0.9;}
.arch-flow-horizontal .arch-diagram-html .arch-step-memberof .arch-group-table td:nth-child(2)::before{content:"Assignment";opacity:0.9;}
.arch-flow-horizontal .arch-diagram-html .arch-step-memberof .arch-group-table td:nth-child(3)::before{content:"Rule";opacity:0.9;}
.arch-flow-horizontal .arch-diagram-html .arch-step-memberof .arch-group-table td:nth-child(3){font-family:ui-monospace,'Cascadia Code','Source Code Pro',Menlo,Consolas,monospace;font-size:0.68rem;background:rgba(0,0,0,0.2);padding:0.4rem 0.5rem;border-radius:4px;word-break:break-all;line-height:1.4;}
.arch-flow-horizontal .arch-diagram-html .arch-group-table .arch-group-rule .arch-rule-inner{display:flex;flex-direction:column;align-items:start;gap:0.15rem;}
.arch-flow-horizontal .arch-diagram-html .arch-group-table .arch-group-rule .arch-rule-label{white-space:nowrap;}
.arch-flow-horizontal .arch-diagram-html .arch-group-table .arch-group-rule .arch-rule-value{word-break:break-word;}
.arch-flow-horizontal .arch-diagram-html .arch-group-table .arch-group-rule code{font-size:0.65rem;padding:0.1rem 0.2rem;background:rgba(255,255,255,0.12);border-radius:3px;font-weight:500;color:#fff;}
.arch-diagram-html .arch-step{background:var(--arch-fill);color:#fff;border-radius:12px;padding:1rem 1.5rem;min-width:280px;max-width:100%;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.15);border:1px solid rgba(255,255,255,0.3);}
.arch-diagram-html .arch-step-title{font-weight:600;font-size:1rem;}
.arch-diagram-html .arch-step-sub{font-size:0.875rem;opacity:0.95;margin-top:0.25rem;word-break:break-word;}
.arch-diagram-html .arch-arrow{color:var(--text-secondary,#64748b);font-size:1.25rem;padding:0.35rem 0;}
.arch-diagram-html .arch-step-device,.arch-diagram-html .arch-step-memberof{text-align:left;min-width:100%;width:100%;max-width:860px;}
.arch-diagram-html .arch-step-device .arch-step-title,.arch-diagram-html .arch-step-memberof .arch-step-title{text-align:center;}
.arch-diagram-html .arch-step-device .arch-group-table-wrap,.arch-diagram-html .arch-step-memberof .arch-group-table-wrap{margin-left:0;margin-right:0;width:100%;min-width:100%;}
.arch-diagram-html .arch-group-table-wrap{margin-top:0.75rem;background:rgba(0,0,0,0.12);border-radius:8px;overflow:hidden;border:1px solid rgba(255,255,255,0.2);}
.arch-diagram-html .arch-group-table{width:100%;border-collapse:collapse;font-size:0.875rem;}
.arch-diagram-html .arch-group-table thead{background:rgba(0,0,0,0.2);}
.arch-diagram-html .arch-group-table th{padding:0.5rem 0.75rem;text-align:left;font-weight:600;font-size:0.75rem;text-transform:uppercase;letter-spacing:0.03em;color:rgba(255,255,255,0.95);border-bottom:1px solid rgba(255,255,255,0.25);}
.arch-diagram-html .arch-group-table .arch-group-th-name{text-align:left;}
.arch-diagram-html .arch-group-table .arch-group-th-type{white-space:nowrap;}
.arch-diagram-html .arch-group-table .arch-group-th-rule{min-width:8rem;}
.arch-diagram-html .arch-group-table td{padding:0.5rem 0.75rem;color:rgba(255,255,255,0.95);border-bottom:1px solid rgba(255,255,255,0.15);text-align:left;}
.arch-diagram-html .arch-group-table tbody tr:last-child td{border-bottom:none;}
.arch-diagram-html .arch-group-table tbody tr:nth-child(even){background:rgba(255,255,255,0.06);}
.arch-diagram-html .arch-group-cell{word-break:break-word;}
.arch-diagram-html .arch-group-type{white-space:nowrap;font-size:0.8rem;}
.arch-diagram-html .arch-group-rule{word-break:break-word;font-size:0.8rem;font-family:ui-monospace,monospace;}
.arch-diagram-html .arch-group-table .arch-group-rule .arch-rule-inner{display:flex;flex-direction:column;align-items:start;gap:0.15rem;}
.arch-diagram-html .arch-group-table .arch-group-rule .arch-rule-label{white-space:nowrap;}
.arch-diagram-html .arch-group-table .arch-group-rule .arch-rule-value{word-break:break-word;}
.arch-diagram-html .arch-group-rule code{background:rgba(0,0,0,0.2);padding:0.15rem 0.35rem;border-radius:4px;font-size:0.75rem;color:#fff;}
.arch-diagram-html .arch-step-enrollment{padding:0;min-width:320px;}
.arch-diagram-html .arch-step-enrollment .arch-step-title{padding:1rem 1.5rem;border-bottom:1px solid rgba(255,255,255,0.25);}
.arch-diagram-html .arch-enrollment-inner{display:flex;flex-direction:column;gap:0;padding:1rem 1.5rem;text-align:left;}
.arch-diagram-html .arch-enrollment-row{margin-bottom:0.5rem;}
.arch-diagram-html .arch-enrollment-row:last-child{margin-bottom:0;}
.arch-diagram-html .arch-inner-box{padding:1rem 1.5rem;text-align:left;border-bottom:1px solid rgba(255,255,255,0.2);background:rgba(0,0,0,0.12);}
.arch-diagram-html .arch-inner-box:last-child{border-bottom:none;}
.arch-diagram-html .arch-inner-title{font-weight:600;font-size:0.95rem;color:#fff;margin-bottom:0.15rem;}
.arch-diagram-html .arch-inner-sub{font-size:0.8rem;opacity:0.9;margin-top:0;color:rgba(255,255,255,0.95);}
.arch-diagram-html .arch-step-autopilot-esp{padding:0;display:flex;flex-direction:column;background:transparent;box-shadow:none;border:none;}
.arch-diagram-html .arch-autopilot-esp-inner{display:flex;flex-direction:column;align-items:stretch;gap:0;}
.arch-diagram-html .arch-esp-block{background:var(--arch-fill);color:#fff;padding:0.75rem 1rem;text-align:center;border-radius:8px;border:1px solid rgba(255,255,255,0.25);box-shadow:0 2px 6px rgba(0,0,0,0.15);}
.arch-diagram-html .arch-esp-block .arch-step-title{font-weight:600;font-size:0.95rem;margin-bottom:0.25rem;}
.arch-diagram-html .arch-esp-block .arch-step-sub{font-size:0.8rem;opacity:0.95;word-break:break-word;}
.arch-diagram-html .arch-arrow-down{display:flex;align-items:center;justify-content:center;padding:0.2rem 0;color:rgba(255,255,255,0.8);font-size:0.85rem;}
.arch-flow-horizontal .arch-diagram-html .arch-step{flex:0 0 auto;min-width:90px;max-width:140px;padding:0.5rem 0.35rem;}
.arch-flow-horizontal .arch-diagram-html .arch-step-device,.arch-flow-horizontal .arch-diagram-html .arch-step-memberof{flex:1 1 0;min-width:180px;max-width:280px;padding:0.5rem 0.35rem;}
.arch-flow-horizontal .arch-diagram-html .arch-step-enrollment{flex:0 0 auto;min-width:120px;max-width:180px;}
.arch-flow-horizontal .arch-diagram-html .arch-step-autopilot-esp{flex:0 0 auto;min-width:100px;max-width:160px;padding:0.4rem 0.35rem;}
.arch-flow-horizontal .arch-diagram-html .arch-autopilot-esp-inner{gap:0;}
.arch-flow-horizontal .arch-diagram-html .arch-esp-block{padding:0.4rem 0.5rem;border-radius:6px;}
.arch-flow-horizontal .arch-diagram-html .arch-esp-block .arch-step-title{font-size:0.8rem;}
.arch-flow-horizontal .arch-diagram-html .arch-esp-block .arch-step-sub{font-size:0.72rem;}
.arch-flow-horizontal .arch-diagram-html .arch-arrow-down{font-size:0.7rem;padding:0.15rem 0;}
.arch-flow-horizontal .arch-diagram-html .arch-step-title{font-size:0.85rem;line-height:1.25;word-break:break-word;}
.arch-flow-horizontal .arch-diagram-html .arch-step-sub{font-size:0.78rem;line-height:1.25;}
.arch-flow-horizontal .arch-diagram-html .arch-arrow{padding:0 0.2rem;font-size:1rem;}
.arch-flow-horizontal .arch-diagram-html .arch-enrollment-inner{padding:0.5rem 0.6rem;}
.arch-flow-horizontal .arch-diagram-html .arch-enrollment-row{margin-bottom:0.35rem;}
.arch-flow-horizontal .arch-diagram-html .arch-inner-box{padding:0.5rem 0.6rem;}
.arch-flow-horizontal .arch-diagram-html .arch-inner-title{font-size:0.8rem;}
.arch-flow-horizontal .arch-diagram-html .arch-inner-sub{font-size:0.75rem;margin-top:0;}
.arch-flow-horizontal .arch-diagram-html .arch-step-enrollment .arch-step-title{padding:0.5rem 0.6rem;}
[data-theme="dark"] .arch-diagram-html .arch-arrow{color:var(--text-secondary);}
[data-theme="dark"] .arch-flow-horizontal .arch-diagram-html .arch-step-memberof .arch-group-table td:nth-child(3){background:rgba(255,255,255,0.05);}
</style>
</head>
<body>
<div class="device-theme-toggle" title="Toggle dark/light mode (follows system when auto)">
<div class="theme-icon"><i class="fas fa-sun"></i></div>
<label class="device-theme-switch">
<input type="checkbox" id="deviceThemeToggle" aria-label="Dark mode">
<span class="device-theme-slider"></span>
</label>
<div class="theme-icon"><i class="fas fa-moon"></i></div>
</div>
<div class="app-container">
<div class="dashboard-header"><h1><i class="fas fa-laptop me-2"></i>Intune Device Visualization</h1><p class="mb-0">__PH_TENANT_DATE__</p></div>
<div class="device-banner"><strong>Device:</strong> __PH_DEVICE_BANNER__</div>
<ul class="nav nav-tabs main-tabs-nav" id="mainTabs" role="tablist">
__PH_OVERVIEW_NAV__
<li class="nav-item"><button class="nav-link __PH_DIAGRAM_TAB_ACTIVE__" id="diagram-tab" data-bs-toggle="tab" data-bs-target="#diagram" type="button" role="tab">Enrollment flow</button></li>
__PH_ARCHITECTURE_NAV__
</ul>
<div class="tab-content" id="mainTabContent">
__PH_OVERVIEW_PANE__
<div class="tab-pane fade __PH_DIAGRAM_PANE_ACTIVE__" id="diagram" role="tabpanel">
<div class="p-4 w-100">
<h4 class="mb-3"><i class="fas fa-drafting-compass me-2"></i>Architecture overview</h4>
<p class="text-muted mb-4">High-level view of identity, management, policies, and assignments for this device.</p>
<h5 class="mb-3">Autopilot Enrollment</h5>
<div class="arch-diagram-wrapper arch-diagram-html arch-flow-horizontal">__PH_ARCHITECTURE__</div>
<h4 class="mt-5 mb-3"><i class="fas fa-project-diagram me-2"></i>Enrollment flow</h4>
<p class="text-muted mb-4">Policies and assignments applied to this device at each stage.</p>
<div class="flow-diagram-wrapper">__PH_APPLIED_FLOW__</div>
</div>
</div>
</div>
</div>
<script src="https://code.jquery.com/jquery-3.7.0.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
<script src="https://cdn.datatables.net/buttons/2.4.1/js/dataTables.buttons.min.js"></script>
<script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.bootstrap5.min.js"></script>
<script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.html5.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script type="module">
import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.esm.min.mjs';
const isDarkTheme = () => document.documentElement.getAttribute('data-theme') === 'dark';
function initMermaid() {
  const isDark = isDarkTheme();
  mermaid.initialize({
    startOnLoad: false,
    theme: isDark ? 'dark' : 'base',
    themeVariables: isDark ? {
      primaryColor: '#0d9488',
      primaryTextColor: '#fff',
      primaryBorderColor: '#14b8a6',
      lineColor: '#94a3b8',
      secondaryColor: '#134e4a',
      tertiaryColor: '#1a1a1a',
      background: '#0d0d0d',
      mainBkg: '#1a1a1a',
      secondBkg: '#262626',
      fontSize: '16px',
      fontFamily: 'Segoe UI, system-ui, sans-serif'
    } : {
      primaryColor: '#0f766e',
      primaryTextColor: '#fff',
      primaryBorderColor: '#115e59',
      lineColor: '#64748b',
      secondaryColor: '#ccfbf1',
      tertiaryColor: '#ffffff',
      background: '#ffffff',
      mainBkg: '#f8fafc',
      secondBkg: '#e8ecf1',
      fontSize: '16px',
      fontFamily: 'Segoe UI, system-ui, sans-serif'
    },
    flowchart: {
      useMaxWidth: true,
      htmlLabels: true,
      curve: 'basis',
      padding: 20
    },
    securityLevel: 'loose'
  });
}
initMermaid();
window.mermaid = mermaid;
</script>
<script>
document.addEventListener('DOMContentLoaded', function() {
  var themeToggle = document.getElementById('deviceThemeToggle');
  var prefersDark = window.matchMedia('(prefers-color-scheme: dark)');
  function applyTheme(isDark) {
    document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light');
    if (themeToggle) themeToggle.checked = isDark;
  }
  var saved = localStorage.getItem('theme');
  if (saved === 'dark' || saved === 'light') {
    applyTheme(saved === 'dark');
  } else {
    applyTheme(prefersDark.matches);
  }
  if (themeToggle) themeToggle.addEventListener('change', function() {
    var isDark = this.checked;
    document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light');
    localStorage.setItem('theme', isDark ? 'dark' : 'light');
  });
  prefersDark.addEventListener('change', function(e) {
    if (localStorage.getItem('theme') === null) {
      applyTheme(e.matches);
    }
  });
  if (jQuery && jQuery.fn.DataTable && jQuery('#allAssignmentsTable').length > 0) {
    var $allTbl = jQuery('#allAssignmentsTable');
    var preCategories = [], preTargets = [], preFilters = [];
    function normFilter(s) { return String(s||'').replace(/\s*\((?:Include|Exclude)$/i, '').trim() || String(s||''); }
    $allTbl.find('tbody tr').each(function() {
      var $row = jQuery(this);
      var $cells = $row.find('td');
      if ($cells.length >= 4) {
        var c = jQuery.trim($cells.eq(1).text()); if (c && preCategories.indexOf(c) === -1) preCategories.push(c);
        var t = jQuery.trim($cells.eq(2).text()); if (t && preTargets.indexOf(t) === -1) preTargets.push(t);
        var f = normFilter(jQuery.trim($cells.eq(3).text())); if (f && preFilters.indexOf(f) === -1) preFilters.push(f);
      }
    });
    preCategories.sort(); preTargets.sort(); preFilters.sort();
    var at = $allTbl.DataTable({
      responsive: true, pageLength: 25, order: [[1,'asc'],[0,'asc']], dom: 'Bfrtip', buttons: ['copy','csv','excel','pdf','print'],
      initComplete: function() {
        var api = this.api();
        function escOv(v){ return (v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
        function getOvChecked(menuId){ return jQuery('#'+menuId+' input.filter-cb:checked').map(function(){ return jQuery(this).attr('data-value'); }).get(); }
        function updateOvBtn(btnId, menuId, label){ var n = getOvChecked(menuId).length; jQuery('#'+btnId).text(n ? label + ' (' + n + ')' : 'Select...'); }
        function normalizeFilterName(v) { var s = String(v||''); return s.replace(/\s*\((?:Include|Exclude)\)$/i, '').trim() || s; }
        function fillOvDropdownFromArray(menuId, btnId, values, label) {
          var menu = jQuery('#'+menuId); if (!menu.length) return; menu.empty();
          jQuery.each(values, function(i,v){ menu.append('<label class="dropdown-item"><input type="checkbox" class="filter-cb" data-value="'+escOv(v)+'"> '+escOv(v)+'</label>'); });
          menu.find('input.filter-cb').on('change', function(e){ e.stopPropagation(); updateOvBtn(btnId, menuId, label); api.draw(); });
          menu.on('click', function(e){ e.stopPropagation(); });
          updateOvBtn(btnId, menuId, label);
        }
        fillOvDropdownFromArray('overviewFilterCategoryMenu','overviewFilterCategoryBtn', preCategories, 'Category');
        fillOvDropdownFromArray('overviewFilterTargetMenu','overviewFilterTargetBtn', preTargets, 'Target');
        fillOvDropdownFromArray('overviewFilterFilterMenu','overviewFilterFilterBtn', preFilters, 'Filter');
        var overviewSearchFn = function(settings, data, dataIndex) {
          if (settings.nTable && settings.nTable.id !== 'allAssignmentsTable') return true;
          if (jQuery('#overviewFilterHideNotAssigned').length && jQuery('#overviewFilterHideNotAssigned').val() === 'hide' && data[2] === 'Not Assigned') return false;
          var searchStr = ''; try { var searchApi = new jQuery.fn.dataTable.Api(settings); searchStr = (searchApi.search() || '').trim(); } catch (e) {}
          if (searchStr) { var found = false; for (var i = 0; i < data.length; i++) { if (data[i] && data[i].toString().toLowerCase().indexOf(searchStr.toLowerCase()) !== -1) { found = true; break; } } if (!found) return false; }
          var c = getOvChecked('overviewFilterCategoryMenu'); if (c.length && jQuery.inArray(data[1], c) === -1) return false;
          var t = getOvChecked('overviewFilterTargetMenu'); if (t.length && jQuery.inArray(data[2], t) === -1) return false;
          var rowFilterNorm = normalizeFilterName(data[3]); var f = getOvChecked('overviewFilterFilterMenu'); if (f.length && jQuery.inArray(rowFilterNorm, f) === -1) return false;
          return true;
        };
        jQuery.fn.dataTable.ext.search.push(overviewSearchFn);
        var hideNotAssigned = jQuery('#overviewFilterHideNotAssigned');
        if (hideNotAssigned.length) hideNotAssigned.on('change', function(){ api.draw(); });
        var resetBtn = jQuery('#overviewFiltersReset');
        if (resetBtn.length) resetBtn.on('click', function(){
          jQuery('#overviewFilterCategoryMenu,#overviewFilterTargetMenu,#overviewFilterFilterMenu').find('input.filter-cb').prop('checked', false);
          updateOvBtn('overviewFilterCategoryBtn','overviewFilterCategoryMenu','Category'); updateOvBtn('overviewFilterTargetBtn','overviewFilterTargetMenu','Target'); updateOvBtn('overviewFilterFilterBtn','overviewFilterFilterMenu','Filter');
          hideNotAssigned.val('');
          api.draw();
        });
      }
    });
    jQuery('#showAllAssignments').prop('checked', false);
    jQuery('#showAllAssignments').on('change', function() {
      var dt = jQuery('#allAssignmentsTable').DataTable();
      var paginateControls = jQuery('#allAssignmentsTable').closest('.dataTables_wrapper').find('.dataTables_paginate, .dataTables_info');
      if (this.checked) { dt.page.len(-1); paginateControls.hide(); } else { dt.page.len(25); paginateControls.show(); }
      dt.draw();
    });
  }
});
if (jQuery && jQuery.fn.DataTable) {
  if (jQuery('#filtersTable').length > 0) {
    var ft = jQuery('#filtersTable').DataTable({responsive:true,pageLength:10,order:[[3,'desc']],dom:'Bfrtip',buttons:['copy','csv','excel','pdf','print']});
    jQuery('#showAllFilters').prop('checked', false);
    jQuery('#showAllFilters').on('change', function() {
      var paginateControls = jQuery('#filtersTable').closest('.dataTables_wrapper').find('.dataTables_paginate, .dataTables_info');
      if (this.checked) { ft.page.len(-1); paginateControls.hide(); } else { ft.page.len(10); paginateControls.show(); }
      ft.draw();
    });
  }
}
</script>
</body>
</html>
'@
    $outDir = Split-Path -Parent $OutputPath
    if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    if (-not [System.IO.Path]::HasExtension($OutputPath)) { $OutputPath = $OutputPath + ".html" }
    $html = $htmlTemplate.Replace('__PH_TITLE_DEVICE__', $titleDevice).Replace('__PH_TENANT_DATE__', $tenantDisplay).Replace('__PH_DEVICE_BANNER__', $bannerDevice).Replace('__PH_OVERVIEW_NAV__', $overviewTabNav).Replace('__PH_OVERVIEW_PANE__', $overviewTabPane).Replace('__PH_DIAGRAM_TAB_ACTIVE__', $diagramTabActive).Replace('__PH_DIAGRAM_PANE_ACTIVE__', $diagramPaneActive).Replace('__PH_APPLIED_FLOW__', $AppliedFlowHtml).Replace('__PH_ARCHITECTURE_NAV__', $architectureTabNav).Replace('__PH_ARCHITECTURE__', $architectureFragment)
    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($false))
    $OutputPath
}
