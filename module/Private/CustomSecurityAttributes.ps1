# Custom Security Attributes - Private helpers

function Get-CustomSecurityAttributeData {
    param(
        [Parameter(Mandatory = $false)]
        [string]$AttributeSet,

        [Parameter(Mandatory = $false)]
        [switch]$DebugMode
    )

    # Helper: query entities with paging and extract custom security attributes
    function Get-EntitiesWithAttributes {
        param([string]$Uri, [string]$EntityType, [hashtable]$Headers = @{})
        $all = [System.Collections.Generic.List[object]]::new()
        try {
            $results = Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers $Headers -OutputType PSObject
            if ($results.value) { $all.AddRange($results.value) }
            while ($results.'@odata.nextLink') {
                $results = Invoke-MgGraphRequest -Method GET -Uri $results.'@odata.nextLink' -Headers $Headers -OutputType PSObject
                if ($results.value) { $all.AddRange($results.value) }
            }
        } catch {
            Write-Host "  WARNING: Could not query $EntityType - $($_.Exception.Message)" -ForegroundColor Yellow
        }
        return $all
    }

    # ===== Step 1: Query all three entity types =====
    Write-Host "Querying users..." -ForegroundColor Cyan
    $allUsers = Get-EntitiesWithAttributes `
        -Uri "https://graph.microsoft.com/v1.0/users?`$count=true&`$select=id,displayName,userPrincipalName,customSecurityAttributes&`$top=999" `
        -EntityType "users" -Headers @{ ConsistencyLevel = 'eventual' }
    Write-Host "  Found $($allUsers.Count) user(s)" -ForegroundColor Green

    Write-Host "Querying devices..." -ForegroundColor Cyan
    $allDevices = Get-EntitiesWithAttributes `
        -Uri "https://graph.microsoft.com/beta/devices?`$select=id,displayName,operatingSystem,customSecurityAttributes&`$top=999" `
        -EntityType "devices"
    Write-Host "  Found $($allDevices.Count) device(s)" -ForegroundColor Green

    Write-Host "Querying enterprise applications..." -ForegroundColor Cyan
    $allApps = Get-EntitiesWithAttributes `
        -Uri "https://graph.microsoft.com/beta/servicePrincipals?`$select=id,displayName,appId,customSecurityAttributes&`$top=999" `
        -EntityType "service principals"
    Write-Host "  Found $($allApps.Count) app(s)" -ForegroundColor Green

    # ===== Step 2: Discover all attribute sets from all entities =====
    Write-Host "Discovering attribute sets..." -ForegroundColor Cyan
    $allAttributeSets = [System.Collections.Generic.HashSet[string]]::new()
    $allEntities = @($allUsers) + @($allDevices) + @($allApps)

    foreach ($entity in $allEntities) {
        if ($entity.customSecurityAttributes) {
            foreach ($prop in $entity.customSecurityAttributes.PSObject.Properties) {
                if ($prop.Name -ne '@odata.type' -and $prop.Value -is [System.Management.Automation.PSCustomObject]) {
                    [void]$allAttributeSets.Add($prop.Name)
                }
            }
        }
    }

    $sortedSets = @($allAttributeSets | Sort-Object)
    if ($sortedSets.Count -eq 0) {
        throw "No custom security attribute sets found in your tenant."
    }
    Write-Host "  Found attribute sets: $($sortedSets -join ', ')" -ForegroundColor Green

    # If a specific set was requested, validate it exists
    if ($AttributeSet -and $AttributeSet -notin $sortedSets) {
        throw "Attribute set '$AttributeSet' not found. Available sets: $($sortedSets -join ', ')"
    }

    # ===== Step 3: Discover attributes per set =====
    $setData = [ordered]@{}
    $setsToProcess = if ($AttributeSet) { @($AttributeSet) } else { $sortedSets }

    foreach ($setName in $setsToProcess) {
        $discoveredAttrs = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($entity in $allEntities) {
            $attrData = $entity.customSecurityAttributes.$setName
            if ($attrData) {
                foreach ($prop in $attrData.PSObject.Properties) {
                    if ($prop.Name -ne '@odata.type') {
                        [void]$discoveredAttrs.Add($prop.Name)
                    }
                }
            }
        }
        $attrNames = @($discoveredAttrs | Sort-Object)
        if ($attrNames.Count -eq 0) { continue }

        # Process each entity type for this set
        $usersForSet = [System.Collections.Generic.List[PSObject]]::new()
        $devicesForSet = [System.Collections.Generic.List[PSObject]]::new()
        $appsForSet = [System.Collections.Generic.List[PSObject]]::new()

        foreach ($user in $allUsers) {
            $attrData = $user.customSecurityAttributes.$setName
            if (-not $attrData) { continue }
            $obj = [ordered]@{ DisplayName = $user.displayName; Identifier = $user.userPrincipalName }
            foreach ($a in $attrNames) { $obj[$a] = if ($attrData.$a) { $attrData.$a } else { '-' } }
            $obj['ObjectId'] = $user.id
            $usersForSet.Add([PSCustomObject]$obj)
        }

        foreach ($device in $allDevices) {
            $attrData = $device.customSecurityAttributes.$setName
            if (-not $attrData) { continue }
            $obj = [ordered]@{ DisplayName = $device.displayName; Identifier = if ($device.operatingSystem) { $device.operatingSystem } else { '-' } }
            foreach ($a in $attrNames) { $obj[$a] = if ($attrData.$a) { $attrData.$a } else { '-' } }
            $obj['ObjectId'] = $device.id
            $devicesForSet.Add([PSCustomObject]$obj)
        }

        foreach ($app in $allApps) {
            $attrData = $app.customSecurityAttributes.$setName
            if (-not $attrData) { continue }
            $obj = [ordered]@{ DisplayName = $app.displayName; Identifier = if ($app.appId) { $app.appId } else { '-' } }
            foreach ($a in $attrNames) { $obj[$a] = if ($attrData.$a) { $attrData.$a } else { '-' } }
            $obj['ObjectId'] = $app.id
            $appsForSet.Add([PSCustomObject]$obj)
        }

        $setData[$setName] = @{
            AttributeNames = $attrNames
            Users          = $usersForSet
            Devices        = $devicesForSet
            Apps           = $appsForSet
        }

        Write-Host "  $setName : $($attrNames.Count) attributes, $($usersForSet.Count) users, $($devicesForSet.Count) devices, $($appsForSet.Count) apps" -ForegroundColor Yellow
    }

    # Build overview data (coverage matrix)
    $overviewData = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($user in $allUsers) {
        if (-not $user.customSecurityAttributes) { continue }
        $hasSets = @{}
        foreach ($s in $setsToProcess) { $hasSets[$s] = if ($user.customSecurityAttributes.$s) { $true } else { $false } }
        if (-not ($hasSets.Values -contains $true)) { continue }
        $overviewData.Add(@{ Type = 'User'; Name = $user.displayName; Identifier = $user.userPrincipalName; Sets = $hasSets })
    }
    foreach ($device in $allDevices) {
        if (-not $device.customSecurityAttributes) { continue }
        $hasSets = @{}
        foreach ($s in $setsToProcess) { $hasSets[$s] = if ($device.customSecurityAttributes.$s) { $true } else { $false } }
        if (-not ($hasSets.Values -contains $true)) { continue }
        $overviewData.Add(@{ Type = 'Device'; Name = $device.displayName; Identifier = if ($device.operatingSystem) { $device.operatingSystem } else { '-' }; Sets = $hasSets })
    }
    foreach ($app in $allApps) {
        if (-not $app.customSecurityAttributes) { continue }
        $hasSets = @{}
        foreach ($s in $setsToProcess) { $hasSets[$s] = if ($app.customSecurityAttributes.$s) { $true } else { $false } }
        if (-not ($hasSets.Values -contains $true)) { continue }
        $overviewData.Add(@{ Type = 'App'; Name = $app.displayName; Identifier = if ($app.appId) { $app.appId } else { '-' }; Sets = $hasSets })
    }

    return @{
        SetData       = $setData
        AttributeSets = $setsToProcess
        OverviewData  = $overviewData
        Counts        = @{
            Users   = ($overviewData | Where-Object { $_.Type -eq 'User' }).Count
            Devices = ($overviewData | Where-Object { $_.Type -eq 'Device' }).Count
            Apps    = ($overviewData | Where-Object { $_.Type -eq 'App' }).Count
            Sets    = $setsToProcess.Count
        }
    }
}

function New-CustomSecurityAttributesHTMLReport {
    param(
        [Parameter(Mandatory = $true)] [string]$TenantName,
        [Parameter(Mandatory = $true)] [hashtable]$ReportData,
        [Parameter(Mandatory = $false)] [string]$ExportPath
    )

    if (-not $ExportPath) {
        $ExportPath = Join-Path (Get-Location).Path "$TenantName-CustomSecurityAttributes.html"
    }

    $exportDir = Split-Path -Path $ExportPath -Parent
    if ($exportDir -and -not (Test-Path $exportDir)) {
        New-Item -Path $exportDir -ItemType Directory -Force | Out-Null
    }

    $reportDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $sets = $ReportData.AttributeSets
    $setData = $ReportData.SetData
    $overview = $ReportData.OverviewData
    $counts = $ReportData.Counts

    # Stat tiles
    $statsCardsHtml = @"
            <div class="rk-stat-tile t-rust">
                <div class="rk-stat-eyebrow">USERS</div>
                <div class="rk-stat-number">$($counts.Users)</div>
                <div class="rk-stat-caption">With attributes assigned</div>
            </div>
            <div class="rk-stat-tile t-olive">
                <div class="rk-stat-eyebrow">DEVICES</div>
                <div class="rk-stat-number">$($counts.Devices)</div>
                <div class="rk-stat-caption">With attributes assigned</div>
            </div>
            <div class="rk-stat-tile t-steel">
                <div class="rk-stat-eyebrow">ENTERPRISE APPS</div>
                <div class="rk-stat-number">$($counts.Apps)</div>
                <div class="rk-stat-caption">With attributes assigned</div>
            </div>
            <div class="rk-stat-tile t-rose">
                <div class="rk-stat-eyebrow">ATTRIBUTE SETS</div>
                <div class="rk-stat-number">$($counts.Sets)</div>
                <div class="rk-stat-caption">Active in tenant</div>
            </div>
"@

    # ===== Build tab buttons =====
    $tabButtonsHtml = '        <button class="rk-tab active" data-target="panel-overview">Overview</button>'
    foreach ($setName in $sets) {
        $tabButtonsHtml += "`n        <button class=`"rk-tab`" data-target=`"panel-$($setName.ToLower())`">$setName</button>"
    }

    # ===== Build Overview panel =====
    $overviewHeaders = "                            <th>Type</th>`n                            <th>Name</th>"
    foreach ($s in $sets) { $overviewHeaders += "`n                            <th>$s</th>" }

    $overviewRows = ''
    foreach ($item in $overview) {
        $typeLabel = switch ($item.Type) {
            'User'   { 'User' }
            'Device' { 'Device' }
            'App'    { 'Enterprise App' }
        }
        $overviewRows += "                        <tr>`n"
        $overviewRows += "                            <td>$typeLabel</td>`n"
        $overviewRows += "                            <td>$([System.Net.WebUtility]::HtmlEncode($item.Name))</td>`n"
        foreach ($s in $sets) {
            $check = if ($item.Sets[$s]) { '<span class="rk-check">&#10003;</span>' } else { '<span class="rk-dash">&mdash;</span>' }
            $overviewRows += "                            <td style=`"text-align:center`">$check</td>`n"
        }
        $overviewRows += "                        </tr>`n"
    }

    # Calculate coverage metrics
    $totalEntities = $overview.Count
    $totalSlots = $totalEntities * $sets.Count
    $assignedSlots = 0
    $fullCoverageCount = 0
    foreach ($item in $overview) {
        $assignedCount = ($sets | Where-Object { $item.Sets[$_] }).Count
        $assignedSlots += $assignedCount
        if ($assignedCount -eq $sets.Count) { $fullCoverageCount++ }
    }
    $avgCoverage = if ($totalSlots -gt 0) { [Math]::Round(($assignedSlots / $totalSlots) * 100) } else { 0 }
    $noCoverageCount = $totalSlots - $assignedSlots

    # Coverage per set
    $setCoverageHtml = ''
    $barColors = @('orange', 'blue', 'green', 'purple', 'red', 'orange', 'blue', 'green')
    $setIdx = 0
    foreach ($s in $sets) {
        $entitiesWithSet = ($overview | Where-Object { $_.Sets[$s] }).Count
        $pct = if ($totalEntities -gt 0) { [Math]::Round(($entitiesWithSet / $totalEntities) * 100) } else { 0 }
        $color = $barColors[$setIdx % $barColors.Count]
        $setCoverageHtml += "<div class=`"rk-cov-row`"><span class=`"rk-cov-name`">$s</span><div class=`"rk-cov-track`"><div class=`"rk-cov-fill rk-cov-$color`" style=`"width:${pct}%`"></div></div><span class=`"rk-cov-pct`">${pct}%</span></div>`n"
        $setIdx++
    }

    # Coverage per entity type
    $userCount = ($overview | Where-Object { $_.Type -eq 'User' }).Count
    $deviceCount = ($overview | Where-Object { $_.Type -eq 'Device' }).Count
    $appCount = ($overview | Where-Object { $_.Type -eq 'App' }).Count
    $entityCoverageHtml = @"
        <div class="rk-cov-row"><span class="rk-cov-name">Users</span><div class="rk-cov-track"><div class="rk-cov-fill rk-cov-orange" style="width:$(if ($totalEntities -gt 0) { [Math]::Round(($userCount / $totalEntities) * 100) } else { 0 })%"></div></div><span class="rk-cov-pct">$userCount</span></div>
        <div class="rk-cov-row"><span class="rk-cov-name">Devices</span><div class="rk-cov-track"><div class="rk-cov-fill rk-cov-blue" style="width:$(if ($totalEntities -gt 0) { [Math]::Round(($deviceCount / $totalEntities) * 100) } else { 0 })%"></div></div><span class="rk-cov-pct">$deviceCount</span></div>
        <div class="rk-cov-row"><span class="rk-cov-name">Enterprise Apps</span><div class="rk-cov-track"><div class="rk-cov-fill rk-cov-green" style="width:$(if ($totalEntities -gt 0) { [Math]::Round(($appCount / $totalEntities) * 100) } else { 0 })%"></div></div><span class="rk-cov-pct">$appCount</span></div>
"@

    $overviewPanelHtml = @"
    <div id="panel-overview" class="rk-panel active">
        <!-- Mini Stats -->
        <div class="rk-mini-stats">
            <div class="rk-mini-stat"><div class="rk-mini-number">$totalEntities</div><div class="rk-mini-label">Total Entities</div></div>
            <div class="rk-mini-stat"><div class="rk-mini-number">$($sets.Count)</div><div class="rk-mini-label">Attribute Sets</div></div>
            <div class="rk-mini-stat"><div class="rk-mini-number">${avgCoverage}%</div><div class="rk-mini-label">Avg Coverage</div></div>
            <div class="rk-mini-stat"><div class="rk-mini-number">$fullCoverageCount</div><div class="rk-mini-label">Full Coverage</div></div>
            <div class="rk-mini-stat"><div class="rk-mini-number">$noCoverageCount</div><div class="rk-mini-label">Unassigned Slots</div></div>
        </div>

        <!-- Coverage Bars -->
        <div class="rk-cov-grid">
            <div class="rk-cov-card">
                <div class="rk-cov-title">Coverage by Attribute Set</div>
$setCoverageHtml
            </div>
            <div class="rk-cov-card">
                <div class="rk-cov-title">Entities by Type</div>
$entityCoverageHtml
            </div>
        </div>

        <div class="rk-card">
            <div class="rk-card-header">
                <span>Attribute Set Coverage Matrix</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="overviewShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="overviewTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
$overviewHeaders
                        </tr>
                    </thead>
                    <tbody>
$overviewRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>
"@

    # ===== Build attribute set panels =====
    $setPanelsHtml = ''
    $setScriptHtml = ''
    $tableCounter = 0

    foreach ($setName in $sets) {
        $sd = $setData[$setName]
        $attrNames = $sd.AttributeNames
        $panelId = "panel-$($setName.ToLower())"

        $panelContent = "    <div id=`"$panelId`" class=`"rk-panel`">`n"

        # Build filter bar with dropdowns per attribute
        $filterId = "filter_$($setName.ToLower())"
        $allEntitiesForSet = @($sd.Users) + @($sd.Devices) + @($sd.Apps)
        $filterDropdowns = "<span>Filters:</span>"
        foreach ($a in $attrNames) {
            $uniqueVals = @($allEntitiesForSet | ForEach-Object { $_.$a } | Where-Object { $_ -and $_ -ne '-' } | Sort-Object -Unique)
            $options = ($uniqueVals | ForEach-Object { "<option value=`"$([System.Net.WebUtility]::HtmlEncode($_))`">$([System.Net.WebUtility]::HtmlEncode($_))</option>" }) -join ''
            $filterDropdowns += "`n            <select class=`"form-select ${filterId}-filter`" style=`"max-width:180px;`"><option value=`"`">All $a</option>$options</select>"
        }
        $filterDropdowns += "`n            <button class=`"rk-filter-chip`" onclick=`"document.querySelectorAll('.${filterId}-filter').forEach(f=>f.value='');document.querySelectorAll('.${filterId}-filter').forEach(f=>f.dispatchEvent(new Event('change')))`">Clear</button>"

        $panelContent += @"
        <div class="rk-filter-bar">
            $filterDropdowns
        </div>
"@

        # Single combined table: Type | Name | Identifier | attributes...
        $tableCounter++
        $tableId = "table_${tableCounter}"
        $toggleId = "toggle_${tableCounter}"

        $combinedHeaders = "                            <th>Type</th>`n                            <th>Name</th>`n                            <th>Identifier</th>"
        foreach ($a in $attrNames) { $combinedHeaders += "`n                            <th>$a</th>" }

        $combinedRows = ''
        # Users
        foreach ($u in $sd.Users) {
            $combinedRows += "                        <tr>`n"
            $combinedRows += "                            <td>User</td>`n"
            $combinedRows += "                            <td>$([System.Net.WebUtility]::HtmlEncode($u.DisplayName))</td>`n"
            $combinedRows += "                            <td class=`"rk-mono`">$([System.Net.WebUtility]::HtmlEncode($u.Identifier))</td>`n"
            foreach ($a in $attrNames) {
                $val = $u.$a
                $display = if ($val -ne '-') { [System.Net.WebUtility]::HtmlEncode($val) } else { '<span style="color:var(--text-dim);font-style:italic">-</span>' }
                $combinedRows += "                            <td>$display</td>`n"
            }
            $combinedRows += "                        </tr>`n"
        }
        # Devices
        foreach ($d in $sd.Devices) {
            $combinedRows += "                        <tr>`n"
            $combinedRows += "                            <td>Device</td>`n"
            $combinedRows += "                            <td>$([System.Net.WebUtility]::HtmlEncode($d.DisplayName))</td>`n"
            $combinedRows += "                            <td>$([System.Net.WebUtility]::HtmlEncode($d.Identifier))</td>`n"
            foreach ($a in $attrNames) {
                $val = $d.$a
                $display = if ($val -ne '-') { [System.Net.WebUtility]::HtmlEncode($val) } else { '<span style="color:var(--text-dim);font-style:italic">-</span>' }
                $combinedRows += "                            <td>$display</td>`n"
            }
            $combinedRows += "                        </tr>`n"
        }
        # Apps
        foreach ($app in $sd.Apps) {
            $combinedRows += "                        <tr>`n"
            $combinedRows += "                            <td>Enterprise App</td>`n"
            $combinedRows += "                            <td>$([System.Net.WebUtility]::HtmlEncode($app.DisplayName))</td>`n"
            $combinedRows += "                            <td class=`"rk-mono`">$([System.Net.WebUtility]::HtmlEncode($app.Identifier))</td>`n"
            foreach ($a in $attrNames) {
                $val = $app.$a
                $display = if ($val -ne '-') { [System.Net.WebUtility]::HtmlEncode($val) } else { '<span style="color:var(--text-dim);font-style:italic">-</span>' }
                $combinedRows += "                            <td>$display</td>`n"
            }
            $combinedRows += "                        </tr>`n"
        }

        $totalCount = $sd.Users.Count + $sd.Devices.Count + $sd.Apps.Count
        $panelContent += @"
        <div class="rk-card">
            <div class="rk-card-header">
                <span>$setName ($totalCount)</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch"><input type="checkbox" id="$toggleId"><span class="rk-toggle-slider"></span></label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="$tableId" class="table table-bordered" style="width:100%">
                    <thead><tr>
$combinedHeaders
                    </tr></thead>
                    <tbody>
$combinedRows
                    </tbody>
                </table>
            </div>
        </div>
"@
        $setScriptHtml += "        var t$tableCounter = initRKTable('#$tableId');`n"
        $setScriptHtml += "        `$('#$toggleId').on('change', function() { t$tableCounter.page.len(`$(this).is(':checked') ? -1 : 10).draw(); });`n"

        $panelContent += "    </div>`n"
        $setPanelsHtml += $panelContent
    }

    # ===== Assemble body content =====
    $bodyContentHtml = @"
    <!-- Tab Navigation -->
    <div class="rk-tabs">
$tabButtonsHtml
    </div>

$overviewPanelHtml
$setPanelsHtml

    <script>
    `$(document).ready(function() {
        var overviewTable = initRKTable('#overviewTable');
        `$('#overviewShowAllToggle').on('change', function() {
            overviewTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });
$setScriptHtml

        // Generic filter logic: each .rk-filter-bar filters all DataTables in the same .rk-panel
        `$.fn.dataTable.ext.search.push(function(settings, data) {
            var table = `$(settings.nTable);
            var panel = table.closest('.rk-panel');
            if (panel.length === 0) return true;
            var filters = panel.find('.rk-filter-bar select');
            if (filters.length === 0) return true;
            for (var i = 0; i < filters.length; i++) {
                var val = `$(filters[i]).val();
                if (val && data[3 + i].indexOf(val) === -1) return false;
            }
            return true;
        });

        // Redraw all tables in the panel when a filter changes
        `$(document).on('change', '.rk-filter-bar select', function() {
            var panel = `$(this).closest('.rk-panel');
            panel.find('table.dataTable').each(function() {
                `$(this).DataTable().draw();
            });
        });
    });
    </script>
"@

    # Custom CSS
    $customCss = @"
    .rk-check { display: inline-flex; align-items: center; justify-content: center; width: 24px; height: 24px; border-radius: 6px; background: rgba(22,163,74,0.1); color: var(--success); font-size: 0.85rem; font-weight: 700; }
    .rk-dash { display: inline-flex; align-items: center; justify-content: center; width: 24px; height: 24px; border-radius: 6px; background: rgba(163,163,163,0.08); color: var(--text-dim); font-size: 0.8rem; }
    [data-theme="dark"] .rk-check { background: rgba(74,222,128,0.1); }
    [data-theme="dark"] .rk-dash { background: rgba(82,82,82,0.15); }
    .rk-mono { font-family: 'Geist Mono', ui-monospace, monospace; font-size: 0.82rem; }
    .table tbody td, .table thead th { word-wrap: break-word; overflow-wrap: break-word; max-width: 300px; }
    .rk-filter-bar .form-select { font-family: 'Geist', -apple-system, sans-serif; font-size: 0.8rem; padding: 4px 8px; border-radius: 6px; }

    /* Mini Stats */
    .rk-mini-stats { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 20px; }
    .rk-mini-stat { background: var(--bg-elevated); border: 1px solid var(--border); border-radius: 10px; padding: 16px; text-align: center; }
    .rk-mini-number { font-family: 'Geist', sans-serif; font-size: 1.6rem; font-weight: 700; color: var(--text); }
    .rk-mini-label { font-family: 'Geist Mono', monospace; font-size: 0.62rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.08em; margin-top: 2px; }

    /* Coverage Bars */
    .rk-cov-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 20px; }
    .rk-cov-card { background: var(--bg-elevated); border: 1px solid var(--border); border-radius: 10px; padding: 16px 20px; }
    .rk-cov-title { font-family: 'Geist Mono', monospace; font-size: 0.68rem; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 12px; }
    .rk-cov-row { display: flex; align-items: center; gap: 12px; margin-bottom: 8px; }
    .rk-cov-row:last-child { margin-bottom: 0; }
    .rk-cov-name { font-size: 0.78rem; color: var(--text-body); min-width: 130px; }
    .rk-cov-track { flex: 1; height: 8px; background: var(--bg-warm); border-radius: 4px; overflow: hidden; }
    .rk-cov-fill { height: 100%; border-radius: 4px; }
    .rk-cov-orange { background: linear-gradient(90deg, #ea580c, #fb923c); }
    .rk-cov-blue { background: linear-gradient(90deg, #0284c7, #38bdf8); }
    .rk-cov-green { background: linear-gradient(90deg, #16a34a, #4ade80); }
    .rk-cov-purple { background: linear-gradient(90deg, #9333ea, #c084fc); }
    .rk-cov-red { background: linear-gradient(90deg, #dc2626, #f87171); }
    .rk-cov-pct { font-family: 'Geist Mono', monospace; font-size: 0.72rem; color: var(--text-muted); min-width: 36px; text-align: right; }
"@

    # Build tags from attribute set names
    $tags = @('Entra ID', 'Security') + $sets

    # Generate final HTML
    $htmlContent = Get-RKSolutionsReportTemplate `
        -TenantName $TenantName `
        -ReportTitle 'Custom Security Attributes' `
        -ReportSlug 'custom-security-attributes' `
        -Eyebrow 'CUSTOM SECURITY ATTRIBUTES' `
        -Lede 'Custom security attribute assignments across users, devices, and enterprise applications.' `
        -StatsCardsHtml $statsCardsHtml `
        -BodyContentHtml $bodyContentHtml `
        -CustomCss $customCss `
        -ReportDate $reportDate `
        -Tags $tags

    $htmlContent | Out-File -FilePath $ExportPath -Encoding utf8

    $script:ExportPath = $ExportPath
    Write-Host "Report saved to: $ExportPath" -ForegroundColor Cyan

    try { Invoke-Item $ExportPath -ErrorAction Stop }
    catch { Write-Host "Report saved to: $ExportPath (could not open automatically)." -ForegroundColor Yellow }

    return $ExportPath
}
