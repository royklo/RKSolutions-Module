# Custom Security Attributes - Private helpers

function Get-CustomSecurityAttributeData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AttributeSet,

        [Parameter(Mandatory = $false)]
        [string[]]$AttributeNames,

        [Parameter(Mandatory = $false)]
        [hashtable]$Filters,

        [Parameter(Mandatory = $false)]
        [switch]$DebugMode
    )

    Write-Host "Retrieving attribute definitions for '$AttributeSet'..." -ForegroundColor Cyan
    $attributeFilter = [uri]::EscapeDataString("attributeSet eq '$AttributeSet'")
    $attrDefUri = "https://graph.microsoft.com/beta/directory/customSecurityAttributeDefinitions?`$filter=$attributeFilter"
    $attributeDefinitions = Invoke-MgGraphRequest -Method GET -Uri $attrDefUri -OutputType PSObject

    if ($attributeDefinitions.value.Count -eq 0) {
        throw "No custom security attribute definitions found for attribute set '$AttributeSet'"
    }

    # If AttributeNames not specified, use all attributes from the set
    if (-not $AttributeNames) {
        $AttributeNames = $attributeDefinitions.value | ForEach-Object { $_.name }
        Write-Host "Using all attributes: $($AttributeNames -join ', ')" -ForegroundColor Yellow
    } else {
        Write-Host "Selected attributes: $($AttributeNames -join ', ')" -ForegroundColor Cyan
    }

    # Build filter query
    $filterConditions = @()
    if ($Filters) {
        foreach ($key in $Filters.Keys) {
            $value = $Filters[$key]
            $filterConditions += "customSecurityAttributes/$AttributeSet/$key eq '$value'"
        }
    }

    if ($filterConditions.Count -eq 0) {
        $firstAttr = $AttributeNames[0]
        $filterQuery = "customSecurityAttributes/$AttributeSet/$firstAttr ne null"
    } else {
        $filterQuery = $filterConditions -join ' and '
    }

    # Query users
    $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$filterQuery&`$count=true&`$select=id,displayName,userPrincipalName,customSecurityAttributes"

    Write-Host 'Querying users with custom security attributes...' -ForegroundColor Cyan
    $results = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers @{
        ConsistencyLevel = 'eventual'
    } -OutputType PSObject

    # Handle paging
    $allUsers = [System.Collections.Generic.List[object]]::new()
    if ($results.value) { $allUsers.AddRange($results.value) }
    while ($results.'@odata.nextLink') {
        $results = Invoke-MgGraphRequest -Method GET -Uri $results.'@odata.nextLink' -Headers @{ ConsistencyLevel = 'eventual' } -OutputType PSObject
        if ($results.value) { $allUsers.AddRange($results.value) }
    }

    # Process results
    $userData = @()
    foreach ($user in $allUsers) {
        $attributeData = $user.customSecurityAttributes.$AttributeSet

        $userObject = [ordered]@{
            DisplayName       = $user.displayName
            UserPrincipalName = $user.userPrincipalName
        }

        foreach ($attrName in $AttributeNames) {
            $userObject[$attrName] = if ($attributeData.$attrName) { $attributeData.$attrName } else { '-' }
        }

        $userObject['UserId'] = $user.id
        $userData += [PSCustomObject]$userObject
    }

    return @{
        UserData       = $userData
        AttributeNames = $AttributeNames
    }
}

function New-CustomSecurityAttributesHTMLReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantName,

        [Parameter(Mandatory = $true)]
        [array]$UserData,

        [Parameter(Mandatory = $true)]
        [string]$AttributeSet,

        [Parameter(Mandatory = $true)]
        [string[]]$AttributeNames,

        [Parameter(Mandatory = $false)]
        [string]$ExportPath
    )

    if (-not $ExportPath) {
        $ExportPath = Join-Path (Get-Location).Path "$TenantName-CustomSecurityAttributes.html"
    }

    $totalUsers = $UserData.Count
    $reportDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

    # Calculate unique counts per attribute
    $attributeStats = @{}
    foreach ($attrName in $AttributeNames) {
        $uniqueCount = ($UserData | Select-Object -ExpandProperty $attrName -Unique | Where-Object { $_ -ne '-' }).Count
        $attributeStats[$attrName] = $uniqueCount
    }

    # Build stat tiles
    $tileColors = @('t-rust', 't-olive', 't-steel', 't-rose')
    $statsCardsHtml = @"
        <div class="rk-stat-tile t-rust">
            <div class="rk-stat-eyebrow">TOTAL USERS</div>
            <div class="rk-stat-number">$totalUsers</div>
            <div class="rk-stat-caption">With $AttributeSet attributes</div>
        </div>
"@
    $colorIdx = 1
    foreach ($attrName in ($AttributeNames | Select-Object -First 3)) {
        $count = $attributeStats[$attrName]
        $color = $tileColors[$colorIdx]
        $statsCardsHtml += @"

        <div class="rk-stat-tile $color">
            <div class="rk-stat-eyebrow">$($attrName.ToUpper())</div>
            <div class="rk-stat-number">$count</div>
            <div class="rk-stat-caption">Unique values</div>
        </div>
"@
        $colorIdx++
    }

    # Build filter dropdowns
    $colSize = [Math]::Max(12 / [Math]::Min($AttributeNames.Count, 4), 3)
    $filterDropdownsHtml = ''
    foreach ($attrName in $AttributeNames) {
        $uniqueValues = $UserData | Select-Object -ExpandProperty $attrName -Unique | Where-Object { $_ -ne '-' } | Sort-Object
        $optionsHtml = ($uniqueValues | ForEach-Object { "<option value=`"$_`">$_</option>" }) -join "`n"
        $filterDropdownsHtml += @"
                <div class="col-md-$colSize">
                    <div class="mb-3">
                        <label for="${attrName}Filter" class="form-label">$attrName</label>
                        <select id="${attrName}Filter" class="form-select attribute-filter">
                            <option value="">All</option>
                            $optionsHtml
                        </select>
                    </div>
                </div>
"@
    }

    # Build table headers
    $tableHeaders = "                            <th>Display Name</th>`n                            <th>User Principal Name</th>"
    foreach ($attrName in $AttributeNames) {
        $tableHeaders += "`n                            <th>$attrName</th>"
    }
    $tableHeaders += "`n                            <th>User ID</th>"

    # Build table rows
    $tableRows = ''
    foreach ($user in $UserData) {
        $tableRows += "                        <tr>`n"
        $tableRows += "                            <td>$($user.DisplayName)</td>`n"
        $tableRows += "                            <td class=`"rk-mono`">$($user.UserPrincipalName)</td>`n"
        foreach ($attrName in $AttributeNames) {
            $value = $user.$attrName
            $displayValue = if ($value -ne '-') { $value } else { '<span style="color: var(--text-dim); font-style: italic;">Not Set</span>' }
            $tableRows += "                            <td>$displayValue</td>`n"
        }
        $tableRows += "                            <td class=`"rk-mono`">$($user.UserId)</td>`n"
        $tableRows += "                        </tr>`n"
    }

    # Build body content
    $bodyContentHtml = @"
    <div class="rk-filter-bar">
        <i class="fas fa-filter"></i>
        <div class="row w-100">
$filterDropdownsHtml
        </div>
    </div>

    <div class="rk-card">
        <div class="rk-card-header">
            <span><i class="fas fa-table me-2"></i>User Custom Security Attributes ($AttributeSet)</span>
            <div class="rk-show-all">
                <label class="rk-toggle-switch">
                    <input type="checkbox" id="showAllToggle">
                    <span class="rk-toggle-slider"></span>
                </label>
                <span>Show all</span>
            </div>
        </div>
        <div class="rk-card-body">
            <div class="table-responsive">
                <table id="usersTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
$tableHeaders
                        </tr>
                    </thead>
                    <tbody>
$tableRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
    `$(document).ready(function() {
        var usersTable = initRKTable('#usersTable');

        `$('#showAllToggle').on('change', function() {
            usersTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        `$.fn.dataTable.ext.search.push(function(settings, data) {
            if (settings.nTable.id !== 'usersTable') return true;
            var filters = `$('.attribute-filter');
            for (var i = 0; i < filters.length; i++) {
                var val = `$(filters[i]).val();
                if (val && !data[2 + i].includes(val)) return false;
            }
            return true;
        });

        `$('.attribute-filter').on('change', function() { usersTable.draw(); });
    });
    </script>
"@

    # Generate final HTML using shared template
    $htmlContent = New-RKSolutionsReportTemplate `
        -TenantName $tenantName `
        -ReportTitle 'Security Attributes' `
        -ReportSlug 'custom-security-attributes' `
        -Eyebrow 'CUSTOM SECURITY ATTRIBUTES' `
        -Lede "Users with custom security attribute assignments across the $AttributeSet attribute set." `
        -StatsCardsHtml $statsCardsHtml `
        -BodyContentHtml $bodyContentHtml `
        -ReportDate $reportDate `
        -Tags @($AttributeSet, 'Entra ID', 'Security')

    $htmlContent | Out-File -FilePath $ExportPath -Encoding utf8
    Write-Host "HTML report saved to: $ExportPath" -ForegroundColor Green

    if (-not $SendEmail) {
        if ($IsWindows -or (-not (Get-Variable -Name 'IsWindows' -ErrorAction SilentlyContinue))) {
            Invoke-Item $ExportPath
        } else {
            & open $ExportPath
        }
    }

    return $ExportPath
}
