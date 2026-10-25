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

    # Generate table rows for devices not registered for Autopilot
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
            <td>$([System.Net.WebUtility]::HtmlEncode($item.IsEncrypted))</td>
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
            <td>$([System.Net.WebUtility]::HtmlEncode($item.OwnerType))</td>
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
                <div class="rk-stat-eyebrow">NOT IN AUTOPILOT</div>
                <div class="rk-stat-number">$Report_DevicesWithoutAutopilotHash_Count</div>
                <div class="rk-stat-caption">No hash or corporate identifier</div>
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
            <div class="rk-stat-tile t-steel">
                <div class="rk-stat-eyebrow">BITLOCKER</div>
                <div class="rk-stat-number">$Report_BitLockerStatus_Count</div>
                <div class="rk-stat-caption">Encryption / escrow gaps</div>
            </div>
            <div class="rk-stat-tile t-indigo">
                <div class="rk-stat-eyebrow">LAPS BACKUP</div>
                <div class="rk-stat-number">$Report_LapsStatus_Count</div>
                <div class="rk-stat-caption">LAPS coverage gaps</div>
            </div>
            <div class="rk-stat-tile t-pink">
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
        <button class="rk-tab" data-target="panel-no-autopilot">Not in Autopilot</button>
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
        <div class="rk-tab-intro"><strong>What we check:</strong> Windows apps in Intune with at least one failed installation, taken from the Intune app install status report and sorted by number of failed devices.</div>
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
        <div class="rk-tab-intro"><strong>What we check:</strong> Devices enrolled with a user-driven (single-user) Autopilot profile where more than one user has signed in. These devices are meant for one person. Multiple users usually means the device was handed over without a reset, or that it is really used as a shared device, in which case a shared-device (self-deploying) enrollment is a better fit.</div>
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

    <!-- Not in Autopilot Panel -->
    <div id="panel-no-autopilot" class="rk-panel">
        <div class="rk-tab-intro"><strong>What we check:</strong> Windows devices that are not registered for Autopilot at all: no hardware hash (Autopilot) and no corporate identifier (Autopilot device preparation). Cloud PCs and Dev Boxes are left out.</div>
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
                <span>Devices not registered for Autopilot</span>
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
        <div class="rk-tab-intro"><strong>What we check:</strong> Devices that have not synced with Intune in the last 90 days.</div>
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
        <div class="rk-tab-intro"><strong>What we check:</strong> Devices that Intune marks as noncompliant, with one row per compliance rule that failed. The Noncompliant Alert column leaves out the <em>device has not checked in</em> and <em>no compliance policy assigned</em> rules.</div>
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
        <div class="rk-tab-intro"><strong>What we check:</strong> Overview of every Windows device with its OS edition and Windows version. Use it to spot unexpected editions: a device running Pro where you expect Enterprise cannot use Enterprise-only features such as Credential Guard, so policies that rely on them silently do nothing.</div>
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
        <div class="rk-tab-intro"><strong>What we check:</strong> Devices whose primary user has a disabled account in Entra ID.</div>
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
        <div class="rk-tab-intro"><strong>What we check:</strong> Windows devices that are not encrypted, have no BitLocker policy assigned, or are encrypted without an OS-volume recovery key stored in Entra ID. Cloud PCs and Dev Boxes are left out. Devices deliberately excluded from BitLocker policies are hidden unless the report runs with <code>-ShowExcludedDevices</code>.</div>
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
                            <th>Encrypted</th>
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
        <div class="rk-tab-intro"><strong>What we check:</strong> Windows devices without an Entra-backed Windows LAPS policy, with a LAPS policy but no password backed up to Entra ID, or with a backup older than 60 days. Devices deliberately excluded from LAPS policies are hidden unless the report runs with <code>-ShowExcludedDevices</code>.</div>
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
                            <th>Ownership</th>
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
        <div class="rk-tab-intro"><strong>What we check:</strong> Settings in your Settings Catalog policies that Microsoft has marked as deprecated. These may stop working or be removed in a future Intune release.</div>
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
            populateSelectFromColumn('bitLockerKeyEscrowedFilter', bitLockerTable, 6);
            populateSelectFromColumn('bitLockerSeverityFilter', bitLockerTable, 7);

            populateSelectFromColumn('lapsCustomerFilter', lapsTable, 0);
            populateSelectFromColumn('lapsDeviceFilter', lapsTable, 1);
            populateSelectFromColumn('lapsUserFilter', lapsTable, 2);
            populateSelectFromColumn('lapsOwnerFilter', lapsTable, 4);
            populateSelectFromColumn('lapsSeverityFilter', lapsTable, 8);

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
            var keyEscrowedFilter = `$('#bitLockerKeyEscrowedFilter').val();
            var severityFilter = `$('#bitLockerSeverityFilter').val();

            bitLockerTable.columns().search('').draw();

            if (customerFilter) bitLockerTable.column(0).search('^' + customerFilter + '`$', true, false);
            if (deviceFilter) bitLockerTable.column(1).search('^' + deviceFilter + '`$', true, false);
            if (userFilter) bitLockerTable.column(2).search('^' + userFilter + '`$', true, false);
            if (keyEscrowedFilter) bitLockerTable.column(6).search('^' + keyEscrowedFilter + '`$', true, false);
            if (severityFilter) bitLockerTable.column(7).search(severityFilter, true, false);

            bitLockerTable.draw();
        };

        window.clearBitLockerFilters = function() {
            `$('#bitLockerCustomerFilter, #bitLockerDeviceFilter, #bitLockerUserFilter, #bitLockerKeyEscrowedFilter, #bitLockerSeverityFilter').val('');
            bitLockerTable.search('').columns().search('').draw();
        };

        // Windows LAPS filter functions
        window.applyLapsFilters = function() {
            var customerFilter = `$('#lapsCustomerFilter').val();
            var deviceFilter = `$('#lapsDeviceFilter').val();
            var userFilter = `$('#lapsUserFilter').val();
            var ownerFilter = `$('#lapsOwnerFilter').val();
            var severityFilter = `$('#lapsSeverityFilter').val();

            lapsTable.columns().search('').draw();

            if (customerFilter) lapsTable.column(0).search('^' + customerFilter + '`$', true, false);
            if (deviceFilter) lapsTable.column(1).search('^' + deviceFilter + '`$', true, false);
            if (userFilter) lapsTable.column(2).search('^' + userFilter + '`$', true, false);
            if (ownerFilter) lapsTable.column(4).search('^' + ownerFilter + '`$', true, false);
            if (severityFilter) lapsTable.column(8).search(severityFilter, true, false);

            lapsTable.draw();
        };

        window.clearLapsFilters = function() {
            `$('#lapsCustomerFilter, #lapsDeviceFilter, #lapsUserFilter, #lapsOwnerFilter, #lapsSeverityFilter').val('');
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
        `$('#bitLockerCustomerFilter, #bitLockerDeviceFilter, #bitLockerUserFilter, #bitLockerKeyEscrowedFilter, #bitLockerSeverityFilter').on('change', function() {
            applyBitLockerFilters();
        });

        // Auto-apply filters on change - Windows LAPS
        `$('#lapsCustomerFilter, #lapsDeviceFilter, #lapsUserFilter, #lapsOwnerFilter, #lapsSeverityFilter').on('change', function() {
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
