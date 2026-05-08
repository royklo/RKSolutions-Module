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
        [array]$Report_NotEncryptedDevices,
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
    $Report_NotEncryptedDevices_Count = $Report_NotEncryptedDevices | Measure-Object | Select-Object -ExpandProperty Count
    $Report_DevicesWithoutAutopilotHash_Count = $Report_DevicesWithoutAutopilotHash | Measure-Object | Select-Object -ExpandProperty Count
    $Report_InactiveDevices_Count = $Report_InactiveDevices | Measure-Object | Select-Object -ExpandProperty Count
    $Report_NoncompliantDevices_Count = ($Report_NoncompliantDevices | Select-Object -Property DeviceName -Unique | Measure-Object).Count
    $Report_OperatingSystemEditionOverview_Count = $Report_OperatingSystemEditionOverview | Measure-Object | Select-Object -ExpandProperty Count
    $Report_DisabledPrimaryUsers_Count = $Report_DisabledPrimaryUsers | Measure-Object | Select-Object -ExpandProperty Count

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

    # Generate table rows for not encrypted devices
    $notEncryptedRows = ""
    foreach ($item in $Report_NotEncryptedDevices) {
        $notEncryptedRows += @"
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
            <div class="rk-stat-tile t-steel">
                <div class="rk-stat-eyebrow">NOT ENCRYPTED</div>
                <div class="rk-stat-number">$Report_NotEncryptedDevices_Count</div>
                <div class="rk-stat-caption">Unencrypted devices</div>
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
"@

    # Build body content HTML (tabs + panels + filter containers + tables + script)
    $bodyContentHtml = @"
    <!-- Tab Navigation -->
    <div class="rk-tabs">
        <button class="rk-tab active" data-target="panel-app-failures">Application Failures</button>
        <button class="rk-tab" data-target="panel-multiple-users">Multiple Users</button>
        <button class="rk-tab" data-target="panel-not-encrypted">Not Encrypted</button>
        <button class="rk-tab" data-target="panel-no-autopilot">No Autopilot Hash</button>
        <button class="rk-tab" data-target="panel-inactive-devices">Inactive Devices</button>
        <button class="rk-tab" data-target="panel-noncompliant">Noncompliant</button>
        <button class="rk-tab" data-target="panel-os-edition">OS Edition Overview</button>
        <button class="rk-tab" data-target="panel-disabled-users">Disabled Primary Users</button>
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

    <!-- Not Encrypted Panel -->
    <div id="panel-not-encrypted" class="rk-panel">
        <div class="rk-filter-bar">
            <span>Filters:</span>
            <select id="notEncryptedCustomerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Customers</option>
            </select>
            <select id="notEncryptedDeviceFilter" class="form-select" style="max-width:180px;">
                <option value="">All Devices</option>
            </select>
            <select id="notEncryptedUserFilter" class="form-select" style="max-width:180px;">
                <option value="">All Users</option>
            </select>
            <select id="notEncryptedManufacturerFilter" class="form-select" style="max-width:180px;">
                <option value="">All Manufacturers</option>
            </select>
            <select id="notEncryptedModelFilter" class="form-select" style="max-width:180px;">
                <option value="">All Models</option>
            </select>
            <button class="rk-filter-chip" onclick="clearNotEncryptedFilters()">Clear</button>
        </div>
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Not Encrypted Devices</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="notEncryptedShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="notEncryptedTable" class="table table-bordered" style="width:100%">
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
                        $notEncryptedRows
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

    <script>
    `$(document).ready(function() {
        // Initialize all tables using the shared helper
        var appFailuresTable = initRKTable('#appFailuresTable');
        var multipleUsersTable = initRKTable('#multipleUsersTable', { order: [[4, 'desc']] });
        var notEncryptedTable = initRKTable('#notEncryptedTable');
        var noAutopilotTable = initRKTable('#noAutopilotTable');
        var inactiveDevicesTable = initRKTable('#inactiveDevicesTable', { order: [[6, 'asc']] });
        var noncompliantTable = initRKTable('#noncompliantTable');
        var osEditionTable = initRKTable('#osEditionTable');
        var disabledUsersTable = initRKTable('#disabledUsersTable');

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

            populateSelectFromColumn('notEncryptedCustomerFilter', notEncryptedTable, 0);
            populateSelectFromColumn('notEncryptedDeviceFilter', notEncryptedTable, 1);
            populateSelectFromColumn('notEncryptedUserFilter', notEncryptedTable, 2);
            populateSelectFromColumn('notEncryptedManufacturerFilter', notEncryptedTable, 4);
            populateSelectFromColumn('notEncryptedModelFilter', notEncryptedTable, 5);

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
        }

        function populateSelectFromColumn(selectId, table, columnIndex) {
            var values = [...new Set(table.column(columnIndex).data().toArray())].sort();
            var select = `$('#' + selectId);
            values.forEach(function(value) {
                if (value && value.toString().trim() !== '') {
                    select.append($('<option>').val(value).text(value));
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

        // Not Encrypted filter functions
        window.applyNotEncryptedFilters = function() {
            var customerFilter = `$('#notEncryptedCustomerFilter').val();
            var deviceFilter = `$('#notEncryptedDeviceFilter').val();
            var userFilter = `$('#notEncryptedUserFilter').val();
            var manufacturerFilter = `$('#notEncryptedManufacturerFilter').val();
            var modelFilter = `$('#notEncryptedModelFilter').val();

            notEncryptedTable.columns().search('').draw();

            if (customerFilter) notEncryptedTable.column(0).search('^' + customerFilter + '`$', true, false);
            if (deviceFilter) notEncryptedTable.column(1).search('^' + deviceFilter + '`$', true, false);
            if (userFilter) notEncryptedTable.column(2).search('^' + userFilter + '`$', true, false);
            if (manufacturerFilter) notEncryptedTable.column(4).search('^' + manufacturerFilter + '`$', true, false);
            if (modelFilter) notEncryptedTable.column(5).search('^' + modelFilter + '`$', true, false);

            notEncryptedTable.draw();
        };

        window.clearNotEncryptedFilters = function() {
            `$('#notEncryptedCustomerFilter, #notEncryptedDeviceFilter, #notEncryptedUserFilter, #notEncryptedManufacturerFilter, #notEncryptedModelFilter').val('');
            notEncryptedTable.search('').columns().search('').draw();
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

        // Auto-apply filters on change - Application Failures
        `$('#appFailuresCustomerFilter, #appFailuresAppFilter, #appFailuresPlatformFilter, #appFailuresVersionFilter, #appFailuresPercentageFilter').on('change', function() {
            applyAppFailuresFilters();
        });

        // Auto-apply filters on change - Multiple Users
        `$('#multipleUsersCustomerFilter, #multipleUsersDeviceFilter, #multipleUsersPrimaryUserFilter, #multipleUsersProfileFilter, #multipleUsersCountFilter').on('change', function() {
            applyMultipleUsersFilters();
        });

        // Auto-apply filters on change - Not Encrypted
        `$('#notEncryptedCustomerFilter, #notEncryptedDeviceFilter, #notEncryptedUserFilter, #notEncryptedManufacturerFilter, #notEncryptedModelFilter').on('change', function() {
            applyNotEncryptedFilters();
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

        // Show all toggle functions for each table
        `$('#appFailuresShowAllToggle').on('change', function() {
            appFailuresTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        `$('#multipleUsersShowAllToggle').on('change', function() {
            multipleUsersTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        `$('#notEncryptedShowAllToggle').on('change', function() {
            notEncryptedTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
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

    # Get all Windows Devices from Microsoft Intune
    $AllDeviceData = Invoke-graphRequestWithPaging -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'&`$select=$($Properties -join ',')"
    #filter out managed by MDE
    $AllDeviceData = $AllDeviceData | Where-Object { $_.managementAgent -ne "msSense" }

    # Get all AutoPilot registered devices under "Enrollment"
    Write-Host "Fetching Autopilot devices..." -ForegroundColor Yellow
    $AutopilotDevices = (Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities")

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

    # Loop through all devices for device data
    $results = [System.Collections.Generic.List[PSObject]]::new()
    $totalDevices = $AllDeviceData.Count

    Write-Host "Processing $totalDevices devices..." -ForegroundColor Yellow

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

            # Initialize compliance rule variables
            $allRules = [System.Collections.Generic.List[string]]::new()
            $uniqueRules = @()

            # Check if device is compliant or not. If not compliant, get compliance rule details
            $FilteredForAlerting = @("DefaultDeviceCompliancePolicy.RequireDeviceCompliancePolicyAssigned", "DefaultDeviceCompliancePolicy.RequireRemainContact")

            if ($DeviceData.complianceState -eq "noncompliant") {
                try {
                    $ComplianceRules = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($DeviceData.id)/deviceCompliancePolicyStates" -ErrorAction SilentlyContinue).value | Where-Object { $_.State -eq "nonCompliant" -or $_.State -eq "Error" }

                    if ($ComplianceRules -and $ComplianceRules.count -le 10) {
                        foreach ($ComplianceRule in $ComplianceRules) {
                            try {
                                $ruleDetails = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($DeviceData.id)/deviceCompliancePolicyStates/$($ComplianceRule.id)/settingStates" -ErrorAction SilentlyContinue).value | Where-Object { $_.state -match 'nonCompliant' }

                                if ($ruleDetails) {
                                    # Add individual rule settings to the collection
                                    foreach ($ruleDetail in $ruleDetails) {
                                        if ($ruleDetail.setting) {
                                            $allRules.Add($ruleDetail.setting)
                                        }
                                    }
                                }
                            } catch {
                                Write-Verbose "Failed to get compliance rule details for $($DeviceData.DeviceName): $_"
                            }
                        }

                        # **FIX**: Get unique values only to eliminate duplicates
                        $uniqueRules = $allRules | Select-Object -Unique
                    }
                } catch {
                    Write-Verbose "Failed to get compliance details for $($DeviceData.DeviceName): $_"
                }
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

            $results.Add([PSCustomObject][ordered]@{
                Customer                   = $TenantName
                DeviceName                 = $DeviceProperties.DeviceName
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
