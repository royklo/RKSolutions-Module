# Intune Anomalies - Private helpers

function New-IntuneAnomaliesHTMLReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantName,

        [Parameter(Mandatory = $false)]
        [array]$Report_ApplicationFailureReport,
        [Parameter(Mandatory = $false)]
        [array]$Report_DevicesWithMultipleUsers,
        [Parameter(Mandatory = $false)]
        [array]$Report_DevicesWithoutAutopilotHash,
        [Parameter(Mandatory = $false)]
        [array]$Report_InactiveDevices,
        [Parameter(Mandatory = $false)]
        [array]$Report_OperatingSystemEditionOverview,
        [Parameter(Mandatory = $false)]
        [array]$Report_NoncompliantDevices,
        [Parameter(Mandatory = $false)]
        [array]$Report_DisabledPrimaryUsers,
        [Parameter(Mandatory = $false)]
        [array]$Report_BitLockerStatus,
        [Parameter(Mandatory = $false)]
        [array]$Report_LapsStatus,
        [Parameter(Mandatory = $false)]
        [array]$Report_DeprecatedSettings,
        [Parameter(Mandatory = $false)]
        [string]$ExportPath
    )

    # Default ExportPath to current folder if not provided
    if (-not $ExportPath) {
        $safeTenantName = $TenantName -replace '[\\/:*?"<>|]', '_'
        $ExportPath = Join-Path (Get-Location).Path "$safeTenantName-IntuneAnomaliesReport.html"
    }

    # Calculate counts for dashboard statistics
    $Report_ApplicationFailureReport_Count = $Report_ApplicationFailureReport | Measure-Object | Select-Object -ExpandProperty Count
    $Report_DevicesWithMultipleUsers_Count = $Report_DevicesWithMultipleUsers | Measure-Object | Select-Object -ExpandProperty Count
    $Report_DevicesWithoutAutopilotHash_Count = $Report_DevicesWithoutAutopilotHash | Measure-Object | Select-Object -ExpandProperty Count
    $Report_InactiveDevices_Count = $Report_InactiveDevices | Measure-Object | Select-Object -ExpandProperty Count
    $Report_NoncompliantDevices_Count = ($Report_NoncompliantDevices | Select-Object -Property DeviceName -Unique | Measure-Object).Count
    $Report_OperatingSystemEditionOverview_Count = $Report_OperatingSystemEditionOverview | Measure-Object | Select-Object -ExpandProperty Count
    $Report_DisabledPrimaryUsers_Count = $Report_DisabledPrimaryUsers | Measure-Object | Select-Object -ExpandProperty Count
    $Report_BitLockerStatus_Count = $Report_BitLockerStatus | Measure-Object | Select-Object -ExpandProperty Count
    $Report_LapsStatus_Count = $Report_LapsStatus | Measure-Object | Select-Object -ExpandProperty Count
    $Report_DeprecatedSettings_Count = $Report_DeprecatedSettings | Measure-Object | Select-Object -ExpandProperty Count
    $Report_DeprecatedPolicies_Count = ($Report_DeprecatedSettings | Select-Object -ExpandProperty PolicyId -Unique | Measure-Object).Count

    # Get the current date and time for the report header
    $CurrentDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

    # Generate table rows for all application failures
    $applicationFailureRows = ""
    foreach ($item in $Report_ApplicationFailureReport) {
        $applicationFailureRows += @"
        <tr>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Customer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Application))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Platform))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Version))</td>
            <td>$($item.FailedDeviceCount)</td>
            <td>$($item.FailedDevicePercentage)%</td>
        </tr>
"@
    }

    # Generate table rows for devices with multiple users
    $multipleUsersRows = ""
    foreach ($item in $Report_DevicesWithMultipleUsers) {
        $multipleUsersRows += @"
        <tr>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Customer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceName))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.PrimaryUser))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.EnrollmentProfile))</td>
            <td><span class="rk-badge rk-badge-warn">$($item.usersLoggedOnCount)</span></td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.usersLoggedOnIds))</td>
        </tr>
"@
    }

    # Generate table rows for Non-company owned devices
    $noAutopilotHashRows = ""
    foreach ($item in $Report_DevicesWithoutAutopilotHash) {
        $noAutopilotHashRows += @"
        <tr>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Customer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceName))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.PrimaryUser))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Serialnumber))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceManufacturer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceModel))</td>
        </tr>
"@
    }

    # Generate table rows for inactive devices
    $inactiveDevicesRows = ""
    foreach ($item in $Report_InactiveDevices) {
        $inactiveDevicesRows += @"
        <tr>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Customer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceName))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.PrimaryUser))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Serialnumber))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceManufacturer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceModel))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.LastContact))</td>
        </tr>
"@
    }

    # Generate table rows for noncompliant devices
    $noncompliantDevicesRows = ""
    foreach ($item in $Report_NoncompliantDevices) {
        $statusBadge = '<span class="rk-badge rk-badge-error">Noncompliant</span>'

        $noncompliantDevicesRows += @"
        <tr>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Customer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceName))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.PrimaryUser))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Serialnumber))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceManufacturer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceModel))</td>
            <td>$statusBadge</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.NoncompliantBasedOn))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.NoncompliantAlert))</td>
        </tr>
"@
    }

    # Generate table rows for OS Edition Overview
    $osEditionOverviewRows = ""
    foreach ($item in $Report_OperatingSystemEditionOverview) {
        $osEditionOverviewRows += @"
        <tr>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Customer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceName))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.PrimaryUser))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.OperatingSystemEdition))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.OSFriendlyname))</td>
        </tr>
"@
    }

    # Generate table rows for BitLocker key escrow anomalies
    $bitLockerStatusRows = ""
    foreach ($item in $Report_BitLockerStatus) {
        $sevClass = switch ($item.Severity) { 'Critical' { 'rk-badge-error' } 'Warning' { 'rk-badge-warn' } default { 'rk-badge' } }
        $bitLockerStatusRows += @"
        <tr>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Customer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceName))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.PrimaryUser))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Serialnumber))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceManufacturer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceModel))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.IsEncrypted))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.PolicyAssigned))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.AppliedPolicies))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.KeyEscrowed))</td>
            <td><span class="rk-badge $sevClass">$([System.Net.WebUtility]::HtmlEncode($item.Severity))</span></td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Status))</td>
        </tr>
"@
    }

    # Generate table rows for Windows LAPS backup anomalies
    $lapsStatusRows = ""
    foreach ($item in $Report_LapsStatus) {
        $sevClass = switch ($item.Severity) { 'Critical' { 'rk-badge-error' } 'Warning' { 'rk-badge-warn' } default { 'rk-badge' } }
        $lapsStatusRows += @"
        <tr>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Customer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceName))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.PrimaryUser))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Serialnumber))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceManufacturer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceModel))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.OwnerType))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.PolicyAssigned))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.AppliedPolicies))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.LastBackupDateTime))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.BackupAgeDays))</td>
            <td><span class="rk-badge $sevClass">$([System.Net.WebUtility]::HtmlEncode($item.Severity))</span></td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Status))</td>
        </tr>
"@
    }

    # Generate table rows for deprecated Settings Catalog settings.
    # Customer / ConfiguredValue / DetectionSource intentionally omitted - they're
    # either constant (Customer) or redundant with the Setting Definition ID for
    # the common choice-setting case where Value == DefId + "_<n>". The full
    # record still carries those fields for downstream consumers.
    $deprecatedSettingsRows = ""
    foreach ($item in $Report_DeprecatedSettings) {
        $deprecatedSettingsRows += @"
        <tr>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.PolicyName))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Platform))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.SettingDisplayName))</td>
            <td><code>$([System.Net.WebUtility]::HtmlEncode($item.SettingDefinitionId))</code></td>
        </tr>
"@
    }

    # Generate table rows for disabled primary users
    $disabledPrimaryUsersRows = ""
    foreach ($item in $Report_DisabledPrimaryUsers) {
        $disabledPrimaryUsersRows += @"
        <tr>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Customer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceName))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.PrimaryUser))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Serialnumber))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceManufacturer))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DeviceModel))</td>
        </tr>
"@
    }

    # Build stat tiles HTML (8 cards, each with a unique color, 5-column grid)
    $statsCardsHtml = @"
            <div class="rk-stat-tile t-rust">
                <div class="rk-stat-eyebrow">APPLICATION FAILURES</div>
                <div class="rk-stat-number">$Report_ApplicationFailureReport_Count</div>
                <div class="rk-stat-caption">Failed app installations</div>
            </div>
            <div class="rk-stat-tile t-olive">
                <div class="rk-stat-eyebrow">MULTIPLE USERS</div>
                <div class="rk-stat-number">$Report_DevicesWithMultipleUsers_Count</div>
                <div class="rk-stat-caption">Non-shared devices</div>
            </div>
            <div class="rk-stat-tile t-rose">
                <div class="rk-stat-eyebrow">NO AUTOPILOT HASH</div>
                <div class="rk-stat-number">$Report_DevicesWithoutAutopilotHash_Count</div>
                <div class="rk-stat-caption">Missing hardware hash</div>
            </div>
            <div class="rk-stat-tile t-amber">
                <div class="rk-stat-eyebrow">INACTIVE DEVICES</div>
                <div class="rk-stat-number">$Report_InactiveDevices_Count</div>
                <div class="rk-stat-caption">90+ days inactive</div>
            </div>
            <div class="rk-stat-tile t-violet">
                <div class="rk-stat-eyebrow">NONCOMPLIANT</div>
                <div class="rk-stat-number">$Report_NoncompliantDevices_Count</div>
                <div class="rk-stat-caption">Unique noncompliant devices</div>
            </div>
            <div class="rk-stat-tile t-teal">
                <div class="rk-stat-eyebrow">OS EDITIONS</div>
                <div class="rk-stat-number">$Report_OperatingSystemEditionOverview_Count</div>
                <div class="rk-stat-caption">OS edition entries</div>
            </div>
            <div class="rk-stat-tile t-slate">
                <div class="rk-stat-eyebrow">DISABLED USERS</div>
                <div class="rk-stat-number">$Report_DisabledPrimaryUsers_Count</div>
                <div class="rk-stat-caption">Disabled primary users</div>
            </div>
            <div class="rk-stat-tile t-rust">
                <div class="rk-stat-eyebrow">BITLOCKER</div>
                <div class="rk-stat-number">$Report_BitLockerStatus_Count</div>
                <div class="rk-stat-caption">Encryption / escrow gaps</div>
            </div>
            <div class="rk-stat-tile t-amber">
                <div class="rk-stat-eyebrow">LAPS BACKUP</div>
                <div class="rk-stat-number">$Report_LapsStatus_Count</div>
                <div class="rk-stat-caption">LAPS coverage gaps</div>
            </div>
            <div class="rk-stat-tile t-violet">
                <div class="rk-stat-eyebrow">DEPRECATED SETTINGS</div>
                <div class="rk-stat-number">$Report_DeprecatedSettings_Count</div>
                <div class="rk-stat-caption">Across $Report_DeprecatedPolicies_Count polic(ies)</div>
            </div>
"@

    # Build body content HTML (tabs + panels + filter containers + tables + script)
    $bodyContentHtml = @"
    <!-- Tab Navigation -->
    <div class="rk-tabs">
        <button class="rk-tab active" data-target="panel-app-failures">Application Failures</button>
        <button class="rk-tab" data-target="panel-multiple-users">Multiple Users</button>
        <button class="rk-tab" data-target="panel-no-autopilot">No Autopilot Hash</button>
        <button class="rk-tab" data-target="panel-inactive-devices">Inactive Devices</button>
        <button class="rk-tab" data-target="panel-noncompliant">Noncompliant</button>
        <button class="rk-tab" data-target="panel-os-edition">OS Edition Overview</button>
        <button class="rk-tab" data-target="panel-disabled-users">Disabled Primary Users</button>
        <button class="rk-tab" data-target="panel-bitlocker-status">BitLocker</button>
        <button class="rk-tab" data-target="panel-laps-status">Windows LAPS</button>
        <button class="rk-tab" data-target="panel-deprecated-settings">Deprecated Settings</button>
    </div>

    <!-- Application Failures Panel -->
    <div id="panel-app-failures" class="rk-panel active">
        <div class="rk-filter-bar">
            <span>Filters:</span>
            <select id="appFailuresCustomerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Customers</option>
            </select>
            <select id="appFailuresAppFilter" class="form-select" style="max-width:180px;">
                <option value="">All Applications</option>
            </select>
            <select id="appFailuresPlatformFilter" class="form-select" style="max-width:180px;">
                <option value="">All Platforms</option>
            </select>
            <select id="appFailuresVersionFilter" class="form-select" style="max-width:180px;">
                <option value="">All Versions</option>
            </select>
            <select id="appFailuresPercentageFilter" class="form-select" style="max-width:180px;">
                <option value="">All Percentages</option>
                <option value="0-20">0-20%</option>
                <option value="20-40">20-40%</option>
                <option value="40-60">40-60%</option>
                <option value="60-80">60-80%</option>
                <option value="80-100">80-100%</option>
            </select>
            <button class="rk-filter-chip" onclick="clearAppFailuresFilters()">Clear</button>
        </div>
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Application Failures</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="appFailuresShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="appFailuresTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Customer</th>
                            <th>Application</th>
                            <th>Platform</th>
                            <th>Version</th>
                            <th>Failed Device Count</th>
                            <th>Failed Device Percentage</th>
                        </tr>
                    </thead>
                    <tbody>
                        $applicationFailureRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Multiple Users Panel -->
    <div id="panel-multiple-users" class="rk-panel">
        <div class="rk-filter-bar">
            <span>Filters:</span>
            <select id="multipleUsersCustomerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Customers</option>
            </select>
            <select id="multipleUsersDeviceFilter" class="form-select" style="max-width:180px;">
                <option value="">All Devices</option>
            </select>
            <select id="multipleUsersPrimaryUserFilter" class="form-select" style="max-width:180px;">
                <option value="">All Users</option>
            </select>
            <select id="multipleUsersProfileFilter" class="form-select" style="max-width:180px;">
                <option value="">All Profiles</option>
            </select>
            <select id="multipleUsersCountFilter" class="form-select" style="max-width:180px;">
                <option value="">All Counts</option>
                <option value="2">2 Users</option>
                <option value="3">3 Users</option>
                <option value="4+">4+ Users</option>
            </select>
            <button class="rk-filter-chip" onclick="clearMultipleUsersFilters()">Clear</button>
        </div>
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Devices with Multiple Users</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="multipleUsersShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="multipleUsersTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Customer</th>
                            <th>Device Name</th>
                            <th>Primary User</th>
                            <th>Enrollment Profile</th>
                            <th>User Count</th>
                            <th>Logged On User IDs</th>
                        </tr>
                    </thead>
                    <tbody>
                        $multipleUsersRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- No Autopilot Hash Panel -->
    <div id="panel-no-autopilot" class="rk-panel">
        <div class="rk-filter-bar">
            <span>Filters:</span>
            <select id="noAutopilotCustomerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Customers</option>
            </select>
            <select id="noAutopilotDeviceFilter" class="form-select" style="max-width:180px;">
                <option value="">All Devices</option>
            </select>
            <select id="noAutopilotUserFilter" class="form-select" style="max-width:180px;">
                <option value="">All Users</option>
            </select>
            <select id="noAutopilotManufacturerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Manufacturers</option>
            </select>
            <select id="noAutopilotModelFilter" class="form-select" style="max-width:180px;">
                <option value="">All Models</option>
            </select>
            <button class="rk-filter-chip" onclick="clearNoAutopilotFilters()">Clear</button>
        </div>
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Non-company owned devices</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="noAutopilotShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="noAutopilotTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Customer</th>
                            <th>Device Name</th>
                            <th>Primary User</th>
                            <th>Serial Number</th>
                            <th>Manufacturer</th>
                            <th>Model</th>
                        </tr>
                    </thead>
                    <tbody>
                        $noAutopilotHashRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Inactive Devices Panel -->
    <div id="panel-inactive-devices" class="rk-panel">
        <div class="rk-filter-bar">
            <span>Filters:</span>
            <select id="inactiveCustomerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Customers</option>
            </select>
            <select id="inactiveDeviceFilter" class="form-select" style="max-width:180px;">
                <option value="">All Devices</option>
            </select>
            <select id="inactiveUserFilter" class="form-select" style="max-width:180px;">
                <option value="">All Users</option>
            </select>
            <select id="inactiveManufacturerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Manufacturers</option>
            </select>
            <select id="inactiveModelFilter" class="form-select" style="max-width:180px;">
                <option value="">All Models</option>
            </select>
            <select id="inactiveInactivityFilter" class="form-select" style="max-width:180px;">
                <option value="">All Periods</option>
                <option value="90-180">90-180 days</option>
                <option value="180+">180+ days</option>
            </select>
            <button class="rk-filter-chip" onclick="clearInactiveFilters()">Clear</button>
        </div>
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Inactive Devices (90+ days)</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="inactiveDevicesShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="inactiveDevicesTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Customer</th>
                            <th>Device Name</th>
                            <th>Primary User</th>
                            <th>Serial Number</th>
                            <th>Manufacturer</th>
                            <th>Model</th>
                            <th>Last Contact</th>
                        </tr>
                    </thead>
                    <tbody>
                        $inactiveDevicesRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Noncompliant Devices Panel -->
    <div id="panel-noncompliant" class="rk-panel">
        <div class="rk-filter-bar">
            <span>Filters:</span>
            <select id="noncompliantCustomerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Customers</option>
            </select>
            <select id="noncompliantDeviceFilter" class="form-select" style="max-width:180px;">
                <option value="">All Devices</option>
            </select>
            <select id="noncompliantUserFilter" class="form-select" style="max-width:180px;">
                <option value="">All Users</option>
            </select>
            <select id="noncompliantManufacturerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Manufacturers</option>
            </select>
            <select id="noncompliantModelFilter" class="form-select" style="max-width:180px;">
                <option value="">All Models</option>
            </select>
            <select id="noncompliantReasonFilter" class="form-select" style="max-width:180px;">
                <option value="">All Reasons</option>
            </select>
            <button class="rk-filter-chip" onclick="clearNoncompliantFilters()">Clear</button>
        </div>
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Noncompliant Devices</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="noncompliantShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="noncompliantTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Customer</th>
                            <th>Device Name</th>
                            <th>Primary User</th>
                            <th>Serial Number</th>
                            <th>Manufacturer</th>
                            <th>Model</th>
                            <th>Compliance Status</th>
                            <th>Noncompliant Based On</th>
                            <th>Noncompliant Alert</th>
                        </tr>
                    </thead>
                    <tbody>
                        $noncompliantDevicesRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- OS Edition Overview Panel -->
    <div id="panel-os-edition" class="rk-panel">
        <div class="rk-filter-bar">
            <span>Filters:</span>
            <select id="osEditionCustomerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Customers</option>
            </select>
            <select id="osEditionDeviceFilter" class="form-select" style="max-width:180px;">
                <option value="">All Devices</option>
            </select>
            <select id="osEditionUserFilter" class="form-select" style="max-width:180px;">
                <option value="">All Users</option>
            </select>
            <select id="osEditionEditionFilter" class="form-select" style="max-width:180px;">
                <option value="">All Editions</option>
            </select>
            <select id="osEditionFriendlyNameFilter" class="form-select" style="max-width:180px;">
                <option value="">All OS Versions</option>
            </select>
            <button class="rk-filter-chip" onclick="clearOSEditionFilters()">Clear</button>
        </div>
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Operating System Edition Overview</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="osEditionShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="osEditionTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Customer</th>
                            <th>Device Name</th>
                            <th>Primary User</th>
                            <th>Operating System Edition</th>
                            <th>OS Friendly Name</th>
                        </tr>
                    </thead>
                    <tbody>
                        $osEditionOverviewRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Disabled Primary Users Panel -->
    <div id="panel-disabled-users" class="rk-panel">
        <div class="rk-filter-bar">
            <span>Filters:</span>
            <select id="disabledUsersCustomerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Customers</option>
            </select>
            <select id="disabledUsersDeviceFilter" class="form-select" style="max-width:180px;">
                <option value="">All Devices</option>
            </select>
            <select id="disabledUsersUserFilter" class="form-select" style="max-width:180px;">
                <option value="">All Users</option>
            </select>
            <select id="disabledUsersManufacturerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Manufacturers</option>
            </select>
            <select id="disabledUsersModelFilter" class="form-select" style="max-width:180px;">
                <option value="">All Models</option>
            </select>
            <button class="rk-filter-chip" onclick="clearDisabledUsersFilters()">Clear</button>
        </div>
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Devices with Disabled Primary Users</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="disabledUsersShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="disabledUsersTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Customer</th>
                            <th>Device Name</th>
                            <th>Primary User</th>
                            <th>Serial Number</th>
                            <th>Manufacturer</th>
                            <th>Model</th>
                        </tr>
                    </thead>
                    <tbody>
                        $disabledPrimaryUsersRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- BitLocker Panel -->
    <div id="panel-bitlocker-status" class="rk-panel">
        <div class="rk-filter-bar">
            <span>Filters:</span>
            <select id="bitLockerCustomerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Customers</option>
            </select>
            <select id="bitLockerDeviceFilter" class="form-select" style="max-width:180px;">
                <option value="">All Devices</option>
            </select>
            <select id="bitLockerUserFilter" class="form-select" style="max-width:180px;">
                <option value="">All Users</option>
            </select>
            <select id="bitLockerPolicyAssignedFilter" class="form-select" style="max-width:180px;">
                <option value="">Any Policy State</option>
            </select>
            <select id="bitLockerKeyEscrowedFilter" class="form-select" style="max-width:180px;">
                <option value="">Any Key State</option>
            </select>
            <select id="bitLockerSeverityFilter" class="form-select" style="max-width:180px;">
                <option value="">All Severities</option>
            </select>
            <button class="rk-filter-chip" onclick="clearBitLockerFilters()">Clear</button>
        </div>
        <div class="rk-card">
            <div class="rk-card-header">
                <span>BitLocker Anomalies</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="bitLockerShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="bitLockerTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Customer</th>
                            <th>Device Name</th>
                            <th>Primary User</th>
                            <th>Serial Number</th>
                            <th>Manufacturer</th>
                            <th>Model</th>
                            <th>Encrypted</th>
                            <th>Policy Assigned</th>
                            <th>Applied Policies</th>
                            <th>Key Escrowed</th>
                            <th>Severity</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        $bitLockerStatusRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Windows LAPS Panel -->
    <div id="panel-laps-status" class="rk-panel">
        <div class="rk-filter-bar">
            <span>Filters:</span>
            <select id="lapsCustomerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Customers</option>
            </select>
            <select id="lapsDeviceFilter" class="form-select" style="max-width:180px;">
                <option value="">All Devices</option>
            </select>
            <select id="lapsUserFilter" class="form-select" style="max-width:180px;">
                <option value="">All Users</option>
            </select>
            <select id="lapsOwnerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Ownerships</option>
            </select>
            <select id="lapsPolicyAssignedFilter" class="form-select" style="max-width:180px;">
                <option value="">Any Policy State</option>
            </select>
            <select id="lapsSeverityFilter" class="form-select" style="max-width:180px;">
                <option value="">All Severities</option>
            </select>
            <button class="rk-filter-chip" onclick="clearLapsFilters()">Clear</button>
        </div>
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Windows LAPS Backup Anomalies</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="lapsShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="lapsTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Customer</th>
                            <th>Device Name</th>
                            <th>Primary User</th>
                            <th>Serial Number</th>
                            <th>Manufacturer</th>
                            <th>Model</th>
                            <th>Ownership</th>
                            <th>Policy Assigned</th>
                            <th>Applied Policies</th>
                            <th>Last Backup</th>
                            <th>Backup Age (days)</th>
                            <th>Severity</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        $lapsStatusRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Deprecated Settings Panel -->
    <div id="panel-deprecated-settings" class="rk-panel">
        <div class="rk-filter-bar">
            <span>Filters:</span>
            <select id="deprecatedPlatformFilter" class="form-select" style="max-width:180px;">
                <option value="">All Platforms</option>
            </select>
            <button class="rk-filter-chip" onclick="clearDeprecatedFilters()">Clear</button>
        </div>
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Deprecated Intune Settings</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="deprecatedShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="deprecatedTable" class="table table-bordered">
                    <thead>
                        <tr>
                            <th>Policy Name</th>
                            <th>Platform</th>
                            <th>Setting Display Name</th>
                            <th>Setting Definition ID</th>
                        </tr>
                    </thead>
                    <tbody>
                        $deprecatedSettingsRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
    `$(document).ready(function() {
        // Initialize all tables using the shared helper
        var appFailuresTable = initRKTable('#appFailuresTable');
        var multipleUsersTable = initRKTable('#multipleUsersTable', { order: [[4, 'desc']] });
        var noAutopilotTable = initRKTable('#noAutopilotTable');
        var inactiveDevicesTable = initRKTable('#inactiveDevicesTable', { order: [[6, 'asc']] });
        var noncompliantTable = initRKTable('#noncompliantTable');
        var osEditionTable = initRKTable('#osEditionTable');
        var disabledUsersTable = initRKTable('#disabledUsersTable');
        var bitLockerTable = initRKTable('#bitLockerTable');
        var lapsTable = initRKTable('#lapsTable');
        var deprecatedTable = initRKTable('#deprecatedTable');

        // Populate filter dropdowns
        function populateFilters() {
            populateSelectFromColumn('appFailuresCustomerFilter', appFailuresTable, 0);
            populateSelectFromColumn('appFailuresAppFilter', appFailuresTable, 1);
            populateSelectFromColumn('appFailuresPlatformFilter', appFailuresTable, 2);
            populateSelectFromColumn('appFailuresVersionFilter', appFailuresTable, 3);

            populateSelectFromColumn('multipleUsersCustomerFilter', multipleUsersTable, 0);
            populateSelectFromColumn('multipleUsersDeviceFilter', multipleUsersTable, 1);
            populateSelectFromColumn('multipleUsersPrimaryUserFilter', multipleUsersTable, 2);
            populateSelectFromColumn('multipleUsersProfileFilter', multipleUsersTable, 3);

            populateSelectFromColumn('noAutopilotCustomerFilter', noAutopilotTable, 0);
            populateSelectFromColumn('noAutopilotDeviceFilter', noAutopilotTable, 1);
            populateSelectFromColumn('noAutopilotUserFilter', noAutopilotTable, 2);
            populateSelectFromColumn('noAutopilotManufacturerFilter', noAutopilotTable, 4);
            populateSelectFromColumn('noAutopilotModelFilter', noAutopilotTable, 5);

            populateSelectFromColumn('inactiveCustomerFilter', inactiveDevicesTable, 0);
            populateSelectFromColumn('inactiveDeviceFilter', inactiveDevicesTable, 1);
            populateSelectFromColumn('inactiveUserFilter', inactiveDevicesTable, 2);
            populateSelectFromColumn('inactiveManufacturerFilter', inactiveDevicesTable, 4);
            populateSelectFromColumn('inactiveModelFilter', inactiveDevicesTable, 5);

            populateSelectFromColumn('noncompliantCustomerFilter', noncompliantTable, 0);
            populateSelectFromColumn('noncompliantDeviceFilter', noncompliantTable, 1);
            populateSelectFromColumn('noncompliantUserFilter', noncompliantTable, 2);
            populateSelectFromColumn('noncompliantManufacturerFilter', noncompliantTable, 4);
            populateSelectFromColumn('noncompliantModelFilter', noncompliantTable, 5);
            populateSelectFromColumn('noncompliantReasonFilter', noncompliantTable, 7);

            populateSelectFromColumn('osEditionCustomerFilter', osEditionTable, 0);
            populateSelectFromColumn('osEditionDeviceFilter', osEditionTable, 1);
            populateSelectFromColumn('osEditionUserFilter', osEditionTable, 2);
            populateSelectFromColumn('osEditionEditionFilter', osEditionTable, 3);
            populateSelectFromColumn('osEditionFriendlyNameFilter', osEditionTable, 4);

            populateSelectFromColumn('disabledUsersCustomerFilter', disabledUsersTable, 0);
            populateSelectFromColumn('disabledUsersDeviceFilter', disabledUsersTable, 1);
            populateSelectFromColumn('disabledUsersUserFilter', disabledUsersTable, 2);
            populateSelectFromColumn('disabledUsersManufacturerFilter', disabledUsersTable, 4);
            populateSelectFromColumn('disabledUsersModelFilter', disabledUsersTable, 5);

            populateSelectFromColumn('bitLockerCustomerFilter', bitLockerTable, 0);
            populateSelectFromColumn('bitLockerDeviceFilter', bitLockerTable, 1);
            populateSelectFromColumn('bitLockerUserFilter', bitLockerTable, 2);
            populateSelectFromColumn('bitLockerPolicyAssignedFilter', bitLockerTable, 7);
            populateSelectFromColumn('bitLockerKeyEscrowedFilter', bitLockerTable, 9);
            populateSelectFromColumn('bitLockerSeverityFilter', bitLockerTable, 10);

            populateSelectFromColumn('lapsCustomerFilter', lapsTable, 0);
            populateSelectFromColumn('lapsDeviceFilter', lapsTable, 1);
            populateSelectFromColumn('lapsUserFilter', lapsTable, 2);
            populateSelectFromColumn('lapsOwnerFilter', lapsTable, 6);
            populateSelectFromColumn('lapsPolicyAssignedFilter', lapsTable, 7);
            populateSelectFromColumn('lapsSeverityFilter', lapsTable, 11);

            populateSelectFromColumn('deprecatedPlatformFilter', deprecatedTable, 1);
        }

        function populateSelectFromColumn(selectId, table, columnIndex) {
            var values = [...new Set(table.column(columnIndex).data().toArray())].sort();
            var select = `$('#' + selectId);
            values.forEach(function(value) {
                if (value && value.toString().trim() !== '') {
                    select.append(`$('<option>').val(value).text(value));
                }
            });
        }

        // Application Failures filter functions
        window.applyAppFailuresFilters = function() {
            var customerFilter = `$('#appFailuresCustomerFilter').val();
            var appFilter = `$('#appFailuresAppFilter').val();
            var platformFilter = `$('#appFailuresPlatformFilter').val();
            var versionFilter = `$('#appFailuresVersionFilter').val();
            var percentageFilter = `$('#appFailuresPercentageFilter').val();

            appFailuresTable.columns().search('').draw();

            if (customerFilter) appFailuresTable.column(0).search('^' + customerFilter + '`$', true, false);
            if (appFilter) appFailuresTable.column(1).search('^' + appFilter + '`$', true, false);
            if (platformFilter) appFailuresTable.column(2).search('^' + platformFilter + '`$', true, false);
            if (versionFilter) appFailuresTable.column(3).search('^' + versionFilter + '`$', true, false);
            if (percentageFilter) {
                var regex = '';
                if (percentageFilter === '0-20') regex = '^(0|[1-9]|1[0-9]|20)%`$';
                else if (percentageFilter === '20-40') regex = '^(2[0-9]|3[0-9]|40)%`$';
                else if (percentageFilter === '40-60') regex = '^(4[0-9]|5[0-9]|60)%`$';
                else if (percentageFilter === '60-80') regex = '^(6[0-9]|7[0-9]|80)%`$';
                else if (percentageFilter === '80-100') regex = '^(8[0-9]|9[0-9]|100)%`$';
                if (regex) appFailuresTable.column(5).search(regex, true, false);
            }

            appFailuresTable.draw();
        };

        window.clearAppFailuresFilters = function() {
            `$('#appFailuresCustomerFilter, #appFailuresAppFilter, #appFailuresPlatformFilter, #appFailuresVersionFilter, #appFailuresPercentageFilter').val('');
            appFailuresTable.search('').columns().search('').draw();
        };

        // Multiple Users filter functions
        window.applyMultipleUsersFilters = function() {
            var customerFilter = `$('#multipleUsersCustomerFilter').val();
            var deviceFilter = `$('#multipleUsersDeviceFilter').val();
            var userFilter = `$('#multipleUsersPrimaryUserFilter').val();
            var profileFilter = `$('#multipleUsersProfileFilter').val();
            var countFilter = `$('#multipleUsersCountFilter').val();

            multipleUsersTable.columns().search('').draw();

            if (customerFilter) multipleUsersTable.column(0).search('^' + customerFilter + '`$', true, false);
            if (deviceFilter) multipleUsersTable.column(1).search('^' + deviceFilter + '`$', true, false);
            if (userFilter) multipleUsersTable.column(2).search('^' + userFilter + '`$', true, false);
            if (profileFilter) multipleUsersTable.column(3).search('^' + profileFilter + '`$', true, false);
            if (countFilter) {
                if (countFilter === '2') multipleUsersTable.column(4).search('^2`$', true, false);
                else if (countFilter === '3') multipleUsersTable.column(4).search('^3`$', true, false);
                else if (countFilter === '4+') multipleUsersTable.column(4).search('[4-9]|[1-9][0-9]+', true, false);
            }

            multipleUsersTable.draw();
        };

        window.clearMultipleUsersFilters = function() {
            `$('#multipleUsersCustomerFilter, #multipleUsersDeviceFilter, #multipleUsersPrimaryUserFilter, #multipleUsersProfileFilter, #multipleUsersCountFilter').val('');
            multipleUsersTable.search('').columns().search('').draw();
        };

        // No Autopilot filter functions
        window.applyNoAutopilotFilters = function() {
            var customerFilter = `$('#noAutopilotCustomerFilter').val();
            var deviceFilter = `$('#noAutopilotDeviceFilter').val();
            var userFilter = `$('#noAutopilotUserFilter').val();
            var manufacturerFilter = `$('#noAutopilotManufacturerFilter').val();
            var modelFilter = `$('#noAutopilotModelFilter').val();

            noAutopilotTable.columns().search('').draw();

            if (customerFilter) noAutopilotTable.column(0).search('^' + customerFilter + '`$', true, false);
            if (deviceFilter) noAutopilotTable.column(1).search('^' + deviceFilter + '`$', true, false);
            if (userFilter) noAutopilotTable.column(2).search('^' + userFilter + '`$', true, false);
            if (manufacturerFilter) noAutopilotTable.column(4).search('^' + manufacturerFilter + '`$', true, false);
            if (modelFilter) noAutopilotTable.column(5).search('^' + modelFilter + '`$', true, false);

            noAutopilotTable.draw();
        };

        window.clearNoAutopilotFilters = function() {
            `$('#noAutopilotCustomerFilter, #noAutopilotDeviceFilter, #noAutopilotUserFilter, #noAutopilotManufacturerFilter, #noAutopilotModelFilter').val('');
            noAutopilotTable.search('').columns().search('').draw();
        };

        // Inactive Devices filter functions
        window.applyInactiveFilters = function() {
            var customerFilter = `$('#inactiveCustomerFilter').val();
            var deviceFilter = `$('#inactiveDeviceFilter').val();
            var userFilter = `$('#inactiveUserFilter').val();
            var manufacturerFilter = `$('#inactiveManufacturerFilter').val();
            var modelFilter = `$('#inactiveModelFilter').val();
            var inactivityFilter = `$('#inactiveInactivityFilter').val();

            inactiveDevicesTable.columns().search('').draw();

            if (customerFilter) inactiveDevicesTable.column(0).search('^' + customerFilter + '`$', true, false);
            if (deviceFilter) inactiveDevicesTable.column(1).search('^' + deviceFilter + '`$', true, false);
            if (userFilter) inactiveDevicesTable.column(2).search('^' + userFilter + '`$', true, false);
            if (manufacturerFilter) inactiveDevicesTable.column(4).search('^' + manufacturerFilter + '`$', true, false);
            if (modelFilter) inactiveDevicesTable.column(5).search('^' + modelFilter + '`$', true, false);

            inactiveDevicesTable.draw();
        };

        window.clearInactiveFilters = function() {
            `$('#inactiveCustomerFilter, #inactiveDeviceFilter, #inactiveUserFilter, #inactiveManufacturerFilter, #inactiveModelFilter, #inactiveInactivityFilter').val('');
            inactiveDevicesTable.search('').columns().search('').draw();
        };

        // Noncompliant Devices filter functions
        window.applyNoncompliantFilters = function() {
            var customerFilter = `$('#noncompliantCustomerFilter').val();
            var deviceFilter = `$('#noncompliantDeviceFilter').val();
            var userFilter = `$('#noncompliantUserFilter').val();
            var manufacturerFilter = `$('#noncompliantManufacturerFilter').val();
            var modelFilter = `$('#noncompliantModelFilter').val();
            var reasonFilter = `$('#noncompliantReasonFilter').val();

            noncompliantTable.columns().search('').draw();

            if (customerFilter) noncompliantTable.column(0).search('^' + customerFilter + '`$', true, false);
            if (deviceFilter) noncompliantTable.column(1).search('^' + deviceFilter + '`$', true, false);
            if (userFilter) noncompliantTable.column(2).search('^' + userFilter + '`$', true, false);
            if (manufacturerFilter) noncompliantTable.column(4).search('^' + manufacturerFilter + '`$', true, false);
            if (modelFilter) noncompliantTable.column(5).search('^' + modelFilter + '`$', true, false);
            if (reasonFilter) noncompliantTable.column(7).search('^' + reasonFilter + '`$', true, false);

            noncompliantTable.draw();
        };

        window.clearNoncompliantFilters = function() {
            `$('#noncompliantCustomerFilter, #noncompliantDeviceFilter, #noncompliantUserFilter, #noncompliantManufacturerFilter, #noncompliantModelFilter, #noncompliantReasonFilter').val('');
            noncompliantTable.search('').columns().search('').draw();
        };

        // OS Edition Overview filter functions
        window.applyOSEditionFilters = function() {
            var customerFilter = `$('#osEditionCustomerFilter').val();
            var deviceFilter = `$('#osEditionDeviceFilter').val();
            var userFilter = `$('#osEditionUserFilter').val();
            var editionFilter = `$('#osEditionEditionFilter').val();
            var friendlyNameFilter = `$('#osEditionFriendlyNameFilter').val();

            osEditionTable.columns().search('').draw();

            if (customerFilter) osEditionTable.column(0).search('^' + customerFilter + '`$', true, false);
            if (deviceFilter) osEditionTable.column(1).search('^' + deviceFilter + '`$', true, false);
            if (userFilter) osEditionTable.column(2).search('^' + userFilter + '`$', true, false);
            if (editionFilter) osEditionTable.column(3).search('^' + editionFilter + '`$', true, false);
            if (friendlyNameFilter) osEditionTable.column(4).search('^' + friendlyNameFilter + '`$', true, false);

            osEditionTable.draw();
        };

        window.clearOSEditionFilters = function() {
            `$('#osEditionCustomerFilter, #osEditionDeviceFilter, #osEditionUserFilter, #osEditionEditionFilter, #osEditionFriendlyNameFilter').val('');
            osEditionTable.search('').columns().search('').draw();
        };

        // Disabled Primary Users filter functions
        window.applyDisabledUsersFilters = function() {
            var customerFilter = `$('#disabledUsersCustomerFilter').val();
            var deviceFilter = `$('#disabledUsersDeviceFilter').val();
            var userFilter = `$('#disabledUsersUserFilter').val();
            var manufacturerFilter = `$('#disabledUsersManufacturerFilter').val();
            var modelFilter = `$('#disabledUsersModelFilter').val();

            disabledUsersTable.columns().search('').draw();

            if (customerFilter) disabledUsersTable.column(0).search('^' + customerFilter + '`$', true, false);
            if (deviceFilter) disabledUsersTable.column(1).search('^' + deviceFilter + '`$', true, false);
            if (userFilter) disabledUsersTable.column(2).search('^' + userFilter + '`$', true, false);
            if (manufacturerFilter) disabledUsersTable.column(4).search('^' + manufacturerFilter + '`$', true, false);
            if (modelFilter) disabledUsersTable.column(5).search('^' + modelFilter + '`$', true, false);

            disabledUsersTable.draw();
        };

        window.clearDisabledUsersFilters = function() {
            `$('#disabledUsersCustomerFilter, #disabledUsersDeviceFilter, #disabledUsersUserFilter, #disabledUsersManufacturerFilter, #disabledUsersModelFilter').val('');
            disabledUsersTable.search('').columns().search('').draw();
        };

        // BitLocker filter functions
        window.applyBitLockerFilters = function() {
            var customerFilter = `$('#bitLockerCustomerFilter').val();
            var deviceFilter = `$('#bitLockerDeviceFilter').val();
            var userFilter = `$('#bitLockerUserFilter').val();
            var policyAssignedFilter = `$('#bitLockerPolicyAssignedFilter').val();
            var keyEscrowedFilter = `$('#bitLockerKeyEscrowedFilter').val();
            var severityFilter = `$('#bitLockerSeverityFilter').val();

            bitLockerTable.columns().search('').draw();

            if (customerFilter) bitLockerTable.column(0).search('^' + customerFilter + '`$', true, false);
            if (deviceFilter) bitLockerTable.column(1).search('^' + deviceFilter + '`$', true, false);
            if (userFilter) bitLockerTable.column(2).search('^' + userFilter + '`$', true, false);
            if (policyAssignedFilter) bitLockerTable.column(7).search('^' + policyAssignedFilter + '`$', true, false);
            if (keyEscrowedFilter) bitLockerTable.column(9).search('^' + keyEscrowedFilter + '`$', true, false);
            if (severityFilter) bitLockerTable.column(10).search(severityFilter, true, false);

            bitLockerTable.draw();
        };

        window.clearBitLockerFilters = function() {
            `$('#bitLockerCustomerFilter, #bitLockerDeviceFilter, #bitLockerUserFilter, #bitLockerPolicyAssignedFilter, #bitLockerKeyEscrowedFilter, #bitLockerSeverityFilter').val('');
            bitLockerTable.search('').columns().search('').draw();
        };

        // Windows LAPS filter functions
        window.applyLapsFilters = function() {
            var customerFilter = `$('#lapsCustomerFilter').val();
            var deviceFilter = `$('#lapsDeviceFilter').val();
            var userFilter = `$('#lapsUserFilter').val();
            var ownerFilter = `$('#lapsOwnerFilter').val();
            var policyAssignedFilter = `$('#lapsPolicyAssignedFilter').val();
            var severityFilter = `$('#lapsSeverityFilter').val();

            lapsTable.columns().search('').draw();

            if (customerFilter) lapsTable.column(0).search('^' + customerFilter + '`$', true, false);
            if (deviceFilter) lapsTable.column(1).search('^' + deviceFilter + '`$', true, false);
            if (userFilter) lapsTable.column(2).search('^' + userFilter + '`$', true, false);
            if (ownerFilter) lapsTable.column(6).search('^' + ownerFilter + '`$', true, false);
            if (policyAssignedFilter) lapsTable.column(7).search('^' + policyAssignedFilter + '`$', true, false);
            if (severityFilter) lapsTable.column(11).search(severityFilter, true, false);

            lapsTable.draw();
        };

        window.clearLapsFilters = function() {
            `$('#lapsCustomerFilter, #lapsDeviceFilter, #lapsUserFilter, #lapsOwnerFilter, #lapsPolicyAssignedFilter, #lapsSeverityFilter').val('');
            lapsTable.search('').columns().search('').draw();
        };

        // Deprecated Settings filter functions
        window.applyDeprecatedFilters = function() {
            var platformFilter = `$('#deprecatedPlatformFilter').val();
            deprecatedTable.columns().search('').draw();
            if (platformFilter) deprecatedTable.column(1).search('^' + platformFilter + '`$', true, false);
            deprecatedTable.draw();
        };

        window.clearDeprecatedFilters = function() {
            `$('#deprecatedPlatformFilter').val('');
            deprecatedTable.search('').columns().search('').draw();
        };

        // Auto-apply filters on change - Application Failures
        `$('#appFailuresCustomerFilter, #appFailuresAppFilter, #appFailuresPlatformFilter, #appFailuresVersionFilter, #appFailuresPercentageFilter').on('change', function() {
            applyAppFailuresFilters();
        });

        // Auto-apply filters on change - Multiple Users
        `$('#multipleUsersCustomerFilter, #multipleUsersDeviceFilter, #multipleUsersPrimaryUserFilter, #multipleUsersProfileFilter, #multipleUsersCountFilter').on('change', function() {
            applyMultipleUsersFilters();
        });

        // Auto-apply filters on change - No Autopilot
        `$('#noAutopilotCustomerFilter, #noAutopilotDeviceFilter, #noAutopilotUserFilter, #noAutopilotManufacturerFilter, #noAutopilotModelFilter').on('change', function() {
            applyNoAutopilotFilters();
        });

        // Auto-apply filters on change - Inactive Devices
        `$('#inactiveCustomerFilter, #inactiveDeviceFilter, #inactiveUserFilter, #inactiveManufacturerFilter, #inactiveModelFilter, #inactiveInactivityFilter').on('change', function() {
            applyInactiveFilters();
        });

        // Auto-apply filters on change - Noncompliant Devices
        `$('#noncompliantCustomerFilter, #noncompliantDeviceFilter, #noncompliantUserFilter, #noncompliantManufacturerFilter, #noncompliantModelFilter, #noncompliantReasonFilter').on('change', function() {
            applyNoncompliantFilters();
        });

        // Auto-apply filters on change - OS Edition Overview
        `$('#osEditionCustomerFilter, #osEditionDeviceFilter, #osEditionUserFilter, #osEditionEditionFilter, #osEditionFriendlyNameFilter').on('change', function() {
            applyOSEditionFilters();
        });

        // Auto-apply filters on change - Disabled Primary Users
        `$('#disabledUsersCustomerFilter, #disabledUsersDeviceFilter, #disabledUsersUserFilter, #disabledUsersManufacturerFilter, #disabledUsersModelFilter').on('change', function() {
            applyDisabledUsersFilters();
        });

        // Auto-apply filters on change - BitLocker
        `$('#bitLockerCustomerFilter, #bitLockerDeviceFilter, #bitLockerUserFilter, #bitLockerPolicyAssignedFilter, #bitLockerKeyEscrowedFilter, #bitLockerSeverityFilter').on('change', function() {
            applyBitLockerFilters();
        });

        // Auto-apply filters on change - Windows LAPS
        `$('#lapsCustomerFilter, #lapsDeviceFilter, #lapsUserFilter, #lapsOwnerFilter, #lapsPolicyAssignedFilter, #lapsSeverityFilter').on('change', function() {
            applyLapsFilters();
        });

        // Auto-apply filters on change - Deprecated Settings
        `$('#deprecatedPlatformFilter').on('change', function() {
            applyDeprecatedFilters();
        });

        // Show all toggle functions for each table
        `$('#appFailuresShowAllToggle').on('change', function() {
            appFailuresTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        `$('#multipleUsersShowAllToggle').on('change', function() {
            multipleUsersTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        `$('#noAutopilotShowAllToggle').on('change', function() {
            noAutopilotTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        `$('#inactiveDevicesShowAllToggle').on('change', function() {
            inactiveDevicesTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        `$('#noncompliantShowAllToggle').on('change', function() {
            noncompliantTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        `$('#osEditionShowAllToggle').on('change', function() {
            osEditionTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        `$('#disabledUsersShowAllToggle').on('change', function() {
            disabledUsersTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        `$('#bitLockerShowAllToggle').on('change', function() {
            bitLockerTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        `$('#lapsShowAllToggle').on('change', function() {
            lapsTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        `$('#deprecatedShowAllToggle').on('change', function() {
            deprecatedTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        // Populate filters after tables are initialized
        setTimeout(function() {
            populateFilters();
        }, 100);
    });
    </script>
"@

    # Report-specific CSS (minimal -- filter bar inline styles only)
    $customCss = @"
    .rk-filter-bar .form-select {
        font-family: 'Geist Mono', ui-monospace, monospace;
        font-size: 0.75rem;
        padding: 4px 8px;
        border-radius: 6px;
    }
"@

    # Generate the full HTML report using the shared template
    $htmlContent = Get-RKSolutionsReportTemplate `
        -TenantName $TenantName `
        -ReportTitle 'Anomalies' `
        -ReportSlug 'intune-anomalies' `
        -Eyebrow 'INTUNE ANOMALIES' `
        -Lede 'Device compliance overview with flagged anomalies across encryption, activity, and application health.' `
        -StatsCardsHtml $statsCardsHtml `
        -BodyContentHtml $bodyContentHtml `
        -CustomCss $customCss `
        -ReportDate $CurrentDate `
        -Tags @('Intune', 'Compliance', 'Security') `
        -StatsClass 'rk-stats-5'

    # Export to HTML file
    $htmlContent | Out-File -FilePath $ExportPath -Encoding utf8

    # Set script-scoped variable for email attachment
    $script:ExportPath = $ExportPath

    Write-Host "INFO: All actions completed successfully."
    Write-Host "INFO: Intune Anomalies Report saved to: $ExportPath" -ForegroundColor Cyan

    # Open the HTML file only if we're not sending email
    if (-not $SendEmail) {
        try { Invoke-Item $ExportPath -ErrorAction Stop }
        catch { Write-Host "Report saved to: $ExportPath (could not open automatically)." -ForegroundColor Yellow }
    }
}


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

        # Map deviceId -> array of nonCompliant/Error policy state ids (preserving the existing <=10 guard)
        $settingStatePairs = [System.Collections.Generic.List[object]]::new()
        foreach ($resp in $policyStateResponses) {
            if ($resp.Status -ne 200 -or -not $resp.Body) { continue }
            $deviceId = $resp.Id -replace '^ps:', ''
            $states = @($resp.Body.value | Where-Object { $_.State -eq 'nonCompliant' -or $_.State -eq 'Error' })
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
            foreach ($resp in $settingResponses) {
                if ($resp.Status -ne 200 -or -not $resp.Body) { continue }
                $deviceId = ($resp.Id -replace '^ss:', '') -split '\|' | Select-Object -First 1
                if (-not $ComplianceRulesByDevice.ContainsKey($deviceId)) { continue }
                $details = @($resp.Body.value | Where-Object { $_.state -match 'nonCompliant' })
                foreach ($det in $details) {
                    if ($det.setting) { $ComplianceRulesByDevice[$deviceId].Add($det.setting) }
                }
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
                NoncompliantBasedOn        = if ($uniqueRules) { $uniqueRules -join ', ' } else { "" }
                NoncompliantAlert          = if ($uniqueRules) { ($uniqueRules | Where-Object { $_ -notin $FilteredForAlerting }) -join ', ' } else { "" }
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

    # Build result array from response values
    $ReturnObject = New-Object System.Collections.ArrayList

    # For each value set in the response
    foreach ($value in $Response.Values) {
        # Create a new line object (hashtable)
        $LineObject = @{ }

        # For each property in the schema
        foreach ($prop in $Response.Schema) {
            $LineObject[$prop.Column] = $value[$Response.Schema.IndexOf($prop)]
        }
        # Check if $LineObject.ApplicationId can be found in $apps
        if ($apps | Where-Object { $_.Id -eq $LineObject.ApplicationId }) {
            $AppAssignment = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($LineObject.ApplicationId)/?`$expand=assignments").assignments
            $AssignmentStatus = $false
            if ($AppAssignment) {
                $AssignmentStatus = $true
            }

            # Use Platform_loc instead of Platform for better readability
            $PlatformName = if ($LineObject.Platform_loc) { $LineObject.Platform_loc } else { $LineObject.Platform }

            $ReturnObject.Add([PSCustomObject][ordered]@{
                    Customer               = $tenantname
                    Application            = ($apps | Where-Object { $_.Id -eq $LineObject.ApplicationId }).displayName
                    Platform               = $PlatformName
                    Version                = $LineObject.AppVersion
                    AssignmentStatus       = $AssignmentStatus
                    FailedUserCount        = $LineObject.FailedUserCount
                    FailedDeviceCount      = $LineObject.FailedDeviceCount
                    FailedDevicePercentage = [double]($LineObject.FailedDevicePercentage / 100).toString('0.00')
                }) | Out-Null
        }
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
# Reuses Invoke-GraphRequestWithPaging, Get-DetailedPolicyAssignments,
# Test-IntuneFilter, and the $script:AllFilters cache established by
# IntuneEnrollmentFlows.ps1 - load order is enforced in RKSolutions.psm1.
# ----------------------------------------------------------------------------

$script:BitLockerSettingPrefix = 'device_vendor_msft_bitlocker_'
$script:LapsSettingPrefix      = 'device_vendor_msft_laps_policies_'
$script:LapsBackupDirectorySetting = 'device_vendor_msft_laps_policies_backupdirectory'

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
        LapsCredentialByDeviceId = @{}    # azureAdDeviceId -> LAPS credential record
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
    foreach ($gid in $referencedGroupIds) {
        $deviceSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $userUpnSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $userIdSet  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        try {
            $devMembers = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/v1.0/groups/$gid/transitiveMembers/microsoft.graph.device?`$select=id,deviceId"
            foreach ($m in $devMembers) { if ($m.deviceId) { [void]$deviceSet.Add([string]$m.deviceId) } }
        }
        catch { Write-Verbose "Group $gid device-member fetch failed: $($_.Exception.Message)" }
        try {
            $userMembers = Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/v1.0/groups/$gid/transitiveMembers/microsoft.graph.user?`$select=id,userPrincipalName"
            foreach ($m in $userMembers) {
                if ($m.id) { [void]$userIdSet.Add([string]$m.id) }
                if ($m.userPrincipalName) { [void]$userUpnSet.Add(([string]$m.userPrincipalName).ToLowerInvariant()) }
            }
        }
        catch { Write-Verbose "Group $gid user-member fetch failed: $($_.Exception.Message)" }
        $context.GroupDeviceAadIds[$gid] = $deviceSet
        $context.GroupUserUpns[$gid]     = $userUpnSet
        $context.GroupUserIds[$gid]      = $userIdSet
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
        via Test-IntuneFilter (provided by IntuneEnrollmentFlows.ps1).
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
                $out.Add([PSCustomObject]@{
                    Customer = $TenantName; DeviceName = $d.DeviceName; PrimaryUser = $d.PrimaryUser
                    Serialnumber = $d.Serialnumber; DeviceManufacturer = $d.DeviceManufacturer; DeviceModel = $d.DeviceModel
                    IsEncrypted = 'No'; PolicyAssigned = 'Unknown'; AppliedPolicies = ''
                    KeyEscrowed = 'Unknown'
                    Status = 'Device not encrypted (no Azure AD device id to verify policy / key state)'
                    Severity = 'Warning'
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

        $out.Add([PSCustomObject]@{
            Customer           = $TenantName
            DeviceName         = $d.DeviceName
            PrimaryUser        = $d.PrimaryUser
            Serialnumber       = $d.Serialnumber
            DeviceManufacturer = $d.DeviceManufacturer
            DeviceModel        = $d.DeviceModel
            IsEncrypted        = if ($isEncrypted) { 'Yes' } else { 'No' }
            PolicyAssigned     = if ($applied.Count -gt 0) { 'Yes' } elseif ($excluded.Count -gt 0) { 'Excluded' } else { 'No' }
            AppliedPolicies    = if ($applied.Count -gt 0) { ($applied | ForEach-Object { $_.PolicyName }) -join '; ' } else { ($excluded | ForEach-Object { "$($_.PolicyName) [$($_.Reason)]" }) -join '; ' }
            KeyEscrowed        = if ($hasKey) { 'Yes' } else { 'No' }
            Status             = $status
            Severity           = $severity
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
            $entry = $Context.LapsCredentialByDeviceId.Values | Where-Object { $_.deviceName -eq $d.DeviceName } | Select-Object -First 1
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

        $out.Add([PSCustomObject]@{
            Customer           = $TenantName
            DeviceName         = $d.DeviceName
            PrimaryUser        = $d.PrimaryUser
            Serialnumber       = $d.Serialnumber
            DeviceManufacturer = $d.DeviceManufacturer
            DeviceModel        = $d.DeviceModel
            OwnerType          = $d.OwnerType
            PolicyAssigned     = if ($applied.Count -gt 0) { 'Yes' } elseif ($excluded.Count -gt 0) { 'Excluded' } else { 'No' }
            AppliedPolicies    = if ($applied.Count -gt 0) { ($applied | ForEach-Object { $_.PolicyName }) -join '; ' } else { ($excluded | ForEach-Object { "$($_.PolicyName) [$($_.Reason)]" }) -join '; ' }
            LastBackupDateTime = if ($lastBackup) { $lastBackup.ToString('yyyy-MM-dd HH:mm') } else { '' }
            BackupAgeDays      = if ($null -ne $ageDays) { $ageDays } else { '' }
            Status             = $status
            Severity           = $severity
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
