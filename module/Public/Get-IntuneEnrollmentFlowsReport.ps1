<#
.SYNOPSIS
    Generates Intune assignment overview and/or device visualization report.
.DESCRIPTION
    Two modes: (1) Assignment overview only (no device). (2) Device visualization (device required):
    outputs Assignment Overview tab and Diagram tab (flow with policy details).
    Device is mandatory for the visualization flow so filters can be validated.
    Connect first with Connect-RKGraph, or pass auth parameters to this cmdlet.
.PARAMETER AssignmentOverviewOnly
    Run assignment collection only and generate the same HTML as the standalone assignment overview (no device).
.PARAMETER Device
    Device identifier: display name, Intune managed device ID (GUID), or Entra device object ID (GUID).
.PARAMETER MermaidOverview
    When using `-Device, also export the mermaid diagram to a standalone .mmd file.
.PARAMETER ApplyPlatformFilter
    When using `-Device, exclude policies whose platform does not match the device. Default is off to match legacy report behavior.
#>
function Get-IntuneEnrollmentFlowsReport {
    [CmdletBinding(DefaultParameterSetName = 'Device')]
    param(
        [Parameter(Mandatory = $false)]
        [string[]] $RequiredScopes = @('User.Read.All', 'Group.Read.All', 'DeviceManagementConfiguration.Read.All', 'DeviceManagementApps.Read.All', 'DeviceManagementManagedDevices.Read.All', 'Device.Read.All', 'CloudPC.Read.All'),
        [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AssignmentOnly')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Device')]
        [string] $TenantId,
        [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AssignmentOnly')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Device')]
        [string] $ClientId,
        [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')]
        [object] $ClientSecret,
        [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')]
        [string] $CertificateThumbprint,
        [Parameter(Mandatory = $true, ParameterSetName = 'Identity')]
        [switch] $Identity,
        [Parameter(Mandatory = $true, ParameterSetName = 'AccessToken')]
        [object] $AccessToken,
        [Parameter(Mandatory = $false)]
        [string] $OutputPath,
        [Parameter(Mandatory = $false)]
        [switch] $ExportToCsv,
        [Parameter(Mandatory = $false)]
        [string] $ExportFolder = '',
        [Parameter(Mandatory = $true, ParameterSetName = 'AssignmentOnly')]
        [switch] $AssignmentOverviewOnly,
        [Parameter(Mandatory = $true, ParameterSetName = 'Device')]
        [ValidateNotNullOrEmpty()]
        [string] $Device,
        [Parameter(Mandatory = $false, ParameterSetName = 'Device')]
        [switch] $MermaidOverview,
        [Parameter(Mandatory = $false, ParameterSetName = 'Device')]
        [switch] $ApplyPlatformFilter,
        [Parameter(Mandatory = $false)]
        [switch] $DebugMode
    )

$ErrorActionPreference = 'Stop'

if ($ClientSecret -is [string]) { $ClientSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force }
if ($AccessToken -is [string]) { $AccessToken = ConvertTo-SecureString $AccessToken -AsPlainText -Force }

try {
    Write-Host 'Intune Assignment Overview (RKSolutions)' -ForegroundColor White
    Write-Host ''

    if (-not $AssignmentOverviewOnly -and [string]::IsNullOrWhiteSpace($Device)) {
        Write-Error "Device is required for the visualization flow. Use -Device 'DeviceName', an Intune managed device ID (GUID), or an Entra device object ID (GUID). Use -AssignmentOverviewOnly for assignment-only report."
        return
    }

    Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Yellow
    $connected = Invoke-RKSolutionsWithConnection -RequiredScopes $RequiredScopes -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -Identity:$Identity -AccessToken $AccessToken -DebugMode:$DebugMode -ParameterSetName $PSCmdlet.ParameterSetName
    if (-not $connected) { throw 'Failed to connect to Microsoft Graph.' }

    $tenantInfo = Invoke-MgGraphRequest -Uri 'beta/organization' -Method Get -OutputType PSObject
    $tenantName = $tenantInfo.value[0].displayName
    Write-Host "Connected to tenant: $tenantName" -ForegroundColor Green
    Write-Host ''

    if ($AssignmentOverviewOnly) {
        $assignments = Get-AllIntunePoliciesWithAssignments -DebugMode:$DebugMode
        Write-Host "  Collected $($assignments.Count) assignment records" -ForegroundColor Green
        Write-Host ''
        if ($ExportToCsv) {
            Write-Host 'Exporting to CSV...' -ForegroundColor Yellow
            $csvPath = Export-Results -Results $assignments -FileName 'IntuneAssignmentOverview' -Extension 'csv' -OutputFolder $ExportFolder -IncludeTimestamp $true -DebugMode:$DebugMode
            Write-Host "CSV saved: $csvPath" -ForegroundColor Green
        }
        $outPath = if ($OutputPath) {
            if (-not [System.IO.Path]::IsPathRooted($OutputPath)) { Join-Path (Get-Location) $OutputPath } else { $OutputPath }
        } else {
            Join-Path (Get-Location) "IntuneAssignmentOverview_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        }
        $generated = New-AssignmentOverviewHtmlReport -PolicyAssignments $assignments -TenantName $tenantName -OutputPath $outPath
        Write-Host "Assignment overview report saved: $generated" -ForegroundColor Green
        try { if ($IsWindows) { Invoke-Item -LiteralPath $generated } else { & /usr/bin/open -- $generated } } catch { Write-Host "Report saved at: $generated" -ForegroundColor Green }
        Write-Host 'Done.' -ForegroundColor Green
        return
    }

    Write-Host "Resolving device: $Device..." -ForegroundColor Yellow
    $deviceContext = Get-DeviceEvaluationContext -DeviceNameOrId $Device -DebugMode:$DebugMode
    if (-not $deviceContext) {
        Write-Error "Device '$Device' not found in Intune. Script stopped."
        return
    }
    $deviceDisplayName = $deviceContext.DeviceProperties.DeviceName
    Write-Host "  Device found: $deviceDisplayName" -ForegroundColor Green
    Write-Host ''

    Write-Host 'Collecting policy assignments...' -ForegroundColor Yellow
    $assignments = Get-AllIntunePoliciesWithAssignments -DebugMode:$DebugMode
    Write-Host "  Collected $($assignments.Count) assignment records" -ForegroundColor Green
    Write-Host ''

    Write-Host 'Evaluating assignments for device...' -ForegroundColor Yellow
    $evaluatedAssignments = Invoke-EvaluateAssignmentsForDevice -PolicyAssignments $assignments -DeviceContext $deviceContext -ApplyPlatformFilter:$ApplyPlatformFilter -DebugMode:$DebugMode
    Write-Host '  Evaluated assignments' -ForegroundColor Green
    $devicePlatform = Get-NormalizedDevicePlatform -OperatingSystem $deviceContext.DeviceProperties.OperatingSystem
    $modelStr = if ($deviceContext.DeviceProperties.Model) { [string]$deviceContext.DeviceProperties.Model } else { '' }
    $isCloudPC = $modelStr.Trim().ToLowerInvariant().StartsWith('cloud pc')
    Write-Host 'Building assignment overview...' -ForegroundColor Yellow
    $overviewFragment = Get-AssignmentOverviewTabFragment -PolicyAssignments $assignments -TenantName $tenantName
    Write-Host '  Assignment overview complete' -ForegroundColor Green

    $deviceGroupDetails = @()
    $directIds = if ($deviceContext.DeviceDirectGroupIds) { @($deviceContext.DeviceDirectGroupIds) } else { @() }
    $allGroupIds = if ($deviceContext.DeviceGroupIds) { @($deviceContext.DeviceGroupIds) } else { @() }
    $allGroupIdsStr = $allGroupIds | ForEach-Object { [string]$_ }
    $idsToHide = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($gid in $allGroupIds) {
        $memberGroupIds = Get-GroupDirectMemberGroupIds -GroupId $gid
        foreach ($mid in $memberGroupIds) { [void]$idsToHide.Add([string]$mid) }
    }
    $groupIdsToShow = $allGroupIds | Where-Object { -not $idsToHide.Contains([string]$_) }
    if ($groupIdsToShow -and $groupIdsToShow.Count -gt 0 -and $script:AllGroups) {
        foreach ($gid in $groupIdsToShow) {
            $g = $script:AllGroups | Where-Object { $_.id -eq $gid } | Select-Object -First 1
            if (-not $g -or -not $g.displayName) { continue }
            $isDynamic = $g.groupTypes -and ($g.groupTypes -contains 'DynamicMembership')
            $isDirect = ($directIds | ForEach-Object { [string]$_ }) -contains [string]$gid
            
            # Validate dynamic group membership: check if device actually matches the rule
            if ($isDynamic -and $g.membershipRule -and $deviceContext.DeviceProperties) {
                $rule = $g.membershipRule
                # Check for [OrderID]:xxx pattern which uses physicalIds
                if ($rule -match '\[OrderID\]:(\w+)') {
                    $orderIdTag = "[OrderID]:$($matches[1])"
                    $physicalIds = $deviceContext.DeviceProperties.PhysicalIds
                    $hasTag = $false
                    if ($physicalIds) {
                        $hasTag = @($physicalIds) -contains $orderIdTag
                    }
                    # Skip this group if the device doesn't have the required OrderID tag
                    if (-not $hasTag) {
                        Write-Host "  ℹ️  Skipping group '$($g.displayName)' - device doesn't match membership rule (missing $orderIdTag)" -ForegroundColor DarkYellow
                        continue
                    }
                }
            }
            
            $groupTypeLabel = if ($isDynamic) { 'Dynamic' } else { 'Assigned' }
            if (-not $isDirect) { $groupTypeLabel += ' (nested)' }
            $ruleDisplay = $null
            if ($isDirect) {
                if ($isDynamic -and $g.membershipRule) {
                    $ruleDisplay = [System.Net.WebUtility]::HtmlEncode($g.membershipRule)
                } else {
                    $parentNames = Get-GroupParentGroupNames -GroupId $gid
                    if ($parentNames -and $parentNames.Count -gt 0) {
                        $nestedValue = ($parentNames | ForEach-Object { '<code>' + [System.Net.WebUtility]::HtmlEncode($_) + '</code>' }) -join '<br>'
                        $ruleDisplay = "<span class=`"arch-rule-inner`"><span class=`"arch-rule-label`">Nested groups:</span><span class=`"arch-rule-value`">$nestedValue</span></span>"
                    }
                }
            } else {
                $nestedFromNames = Get-NestedGroupChainNames -GroupId $gid -DeviceGroupIdsStr $allGroupIdsStr
                if ($nestedFromNames -and $nestedFromNames.Count -gt 0) {
                    $nestedValue = ($nestedFromNames | ForEach-Object { '<code>' + [System.Net.WebUtility]::HtmlEncode($_) + '</code>' }) -join '<br>'
                    $ruleDisplay = "<span class=`"arch-rule-inner`"><span class=`"arch-rule-label`">Nested groups:</span><span class=`"arch-rule-value`">$nestedValue</span></span>"
                }
            }
            $deviceGroupDetails += [PSCustomObject]@{
                DisplayName    = $g.displayName
                GroupType      = $groupTypeLabel
                MembershipRule = $ruleDisplay
            }
        }
    }
    $outPath = if ($OutputPath) {
        if (-not [System.IO.Path]::IsPathRooted($OutputPath)) { Join-Path (Get-Location) $OutputPath } else { $OutputPath }
    } else {
        Join-Path (Get-Location) "IntuneDeviceVisualization_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    }
    $intuneId = if ($deviceContext.ManagedDeviceId) { [string]$deviceContext.ManagedDeviceId } else { '' }
    $entraId = if ($deviceContext.EntraDeviceObjectId) { [string]$deviceContext.EntraDeviceObjectId } else { '' }
    $generated = New-DeviceVisualizationHtmlReport -EvaluatedAssignments $evaluatedAssignments -DeviceName $deviceDisplayName -TenantName $tenantName -OutputPath $outPath -AssignmentOverviewFragment $overviewFragment -DeviceGroupDetails $deviceGroupDetails -IntuneDeviceId $intuneId -EntraDeviceId $entraId -DevicePlatform $devicePlatform -IsCloudPC:$isCloudPC
    Write-Host "Device visualization report saved: $generated" -ForegroundColor Green

    if ($MermaidOverview) {
        $mmdPath = [System.IO.Path]::ChangeExtension($generated, '.mmd')
        [System.IO.File]::WriteAllText($mmdPath, $mermaidDiagram, [System.Text.UTF8Encoding]::new($false))
        Write-Host "Mermaid file saved: $mmdPath" -ForegroundColor Green
    }

    if ($ExportToCsv) {
        $csvPath = Export-Results -Results $evaluatedAssignments -FileName 'IntuneDeviceAssignments' -Extension 'csv' -OutputFolder $ExportFolder -IncludeTimestamp $true -DebugMode:$DebugMode
        Write-Host "CSV saved: $csvPath" -ForegroundColor Green
    }

    try { if ($IsWindows) { Invoke-Item -LiteralPath $generated } else { & /usr/bin/open -- $generated } } catch { Write-Host "Report saved at: $generated" -ForegroundColor Green }
    Write-Host 'Done.' -ForegroundColor Green
}
catch {
    Write-Error "Error: $_"; throw $_
}
finally {
    # Session left connected; use Disconnect-RKGraph when done.
}
}
