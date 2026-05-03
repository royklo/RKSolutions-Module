# M365 License - Private helpers

function New-HTMLReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Organization,

        [Parameter(Mandatory = $true)]
        [array]$Report,

        [Parameter(Mandatory = $true)]
        [array]$SubscriptionOverview,

        [Parameter(Mandatory = $false)]
        [string]$ExportPath
    )

    # Default ExportPath to current folder if not provided
    if (-not $ExportPath) {
        $safeOrganization = $Organization -replace '[\\/:*?"<>|]', '_'
        $ExportPath = Join-Path (Get-Location).Path "$safeOrganization-M365LicensingReport.html"
    }

    $exportDir = Split-Path -Path $ExportPath -Parent
    if ($exportDir -and -not (Test-Path $exportDir)) {
        New-Item -Path $exportDir -ItemType Directory -Force | Out-Null
    }

    # Calculate license counts for dashboard statistics
    $directLicenses = ($Report | Where-Object { $_.AssignmentType -eq "Direct" }).Count
    $inheritedLicenses = ($Report | Where-Object { $_.AssignmentType -eq "Inherited" }).Count
    $bothLicenses = ($Report | Where-Object { $_.AssignmentType -eq "Both" }).Count
    $DisabledUsersWithLicenses = ($Report | Where-Object { $_.AccountEnabled -eq "No" }).Count

    # Generate table rows for user licenses
    $tableRows = ""
    foreach ($item in $Report) {
        $accountStatusClass = if ($item.AccountEnabled -eq "No") { 'class="table-danger"' } else { '' }

        $assignmentTypeBadge = switch ($item.AssignmentType) {
            "Direct" { '<span class="rk-badge rk-badge-ok">Direct</span>' }
            "Inherited" { '<span class="rk-badge rk-badge-accent">Inherited</span>' }
            "Both" { '<span class="rk-badge rk-badge-warn">Both</span>' }
            default { '<span class="rk-badge rk-badge-na">Unknown</span>' }
        }

        $accountStatus = if ($item.AccountEnabled -eq "Yes") {
            '<span class="rk-badge rk-badge-ok">Enabled</span>'
        }
        else {
            '<span class="rk-badge rk-badge-error">Disabled</span>'
        }

        $tableRows += @"
    <tr $accountStatusClass>
        <td>$($item.DisplayName)</td>
        <td>$($item.UserPrincipalName)</td>
        <td>$accountStatus</td>
        <td>$($item.LastSuccessfulSignIn)</td>
        <td>$($item.AssignedLicensesFriendlyName)</td>
        <td>$assignmentTypeBadge</td>
        <td>$($item.Inheritance)</td>
    </tr>
"@
    }

    # Generate table rows for subscription overview
    $subscriptionRows = ""
    foreach ($item in $SubscriptionOverview) {
        $availabilityPercentage = if ($item.TotalLicenses -ne 0) {
            [Math]::Round(($item.AvailableLicenses / $item.TotalLicenses) * 100)
        }
        else {
            0
        }

        $availabilityBadge = if ($availabilityPercentage -lt 10) {
            '<span class="rk-badge rk-badge-error">' + $item.AvailableLicenses + ' (' + $availabilityPercentage + '%)</span>'
        }
        elseif ($availabilityPercentage -lt 20) {
            '<span class="rk-badge rk-badge-warn">' + $item.AvailableLicenses + ' (' + $availabilityPercentage + '%)</span>'
        }
        else {
            '<span class="rk-badge rk-badge-ok">' + $item.AvailableLicenses + ' (' + $availabilityPercentage + '%)</span>'
        }

        $licenseStatusBadge = if ($item.LicenseStatus -eq "Enabled") {
            '<span class="rk-badge rk-badge-ok">Enabled</span>'
        }
        else {
            '<span class="rk-badge rk-badge-error">Disabled</span>'
        }

        $subscriptionRows += @"
    <tr>
        <td>$($item.FriendlyName)</td>
        <td>$($item.CreatedDate)</td>
        <td>$($item.EndDate)</td>
        <td>$licenseStatusBadge</td>
        <td>$($item.ConsumedUnits)</td>
        <td>$($item.TotalLicenses)</td>
        <td>$availabilityBadge</td>
    </tr>
"@
    }

    # Generate table rows for disabled users (users with AccountEnabled = No that have licenses)
    $disabledUsersRows = ""
    $disabledUsers = $Report | Where-Object { $_.AccountEnabled -eq "No" }
    foreach ($item in $disabledUsers) {
        $assignmentTypeBadge = switch ($item.AssignmentType) {
            "Direct" { '<span class="rk-badge rk-badge-ok">Direct</span>' }
            "Inherited" { '<span class="rk-badge rk-badge-accent">Inherited</span>' }
            "Both" { '<span class="rk-badge rk-badge-warn">Both</span>' }
            default { '<span class="rk-badge rk-badge-na">Unknown</span>' }
        }

        $disabledUsersRows += @"
    <tr>
        <td>$($item.DisplayName)</td>
        <td>$($item.UserPrincipalName)</td>
        <td><span class="rk-badge rk-badge-error">Disabled</span></td>
        <td>$($item.LastSuccessfulSignIn)</td>
        <td>$($item.AssignedLicensesFriendlyName)</td>
        <td>$assignmentTypeBadge</td>
        <td>$($item.Inheritance)</td>
    </tr>
"@
    }

    # Get current date for report
    $CurrentDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

    # Build stat tiles HTML
    $statsCardsHtml = @"
            <div class="rk-stat-tile t-rust">
                <div class="rk-stat-eyebrow">DIRECT</div>
                <div class="rk-stat-number">$directLicenses</div>
                <div class="rk-stat-caption">Direct license assignments</div>
            </div>
            <div class="rk-stat-tile t-olive">
                <div class="rk-stat-eyebrow">INHERITED</div>
                <div class="rk-stat-number">$inheritedLicenses</div>
                <div class="rk-stat-caption">Group-based assignments</div>
            </div>
            <div class="rk-stat-tile t-steel">
                <div class="rk-stat-eyebrow">BOTH</div>
                <div class="rk-stat-number">$bothLicenses</div>
                <div class="rk-stat-caption">Direct + Inherited</div>
            </div>
            <div class="rk-stat-tile t-rose">
                <div class="rk-stat-eyebrow">DISABLED</div>
                <div class="rk-stat-number">$DisabledUsersWithLicenses</div>
                <div class="rk-stat-caption">Disabled users with licenses</div>
            </div>
"@

    # Build body content HTML (tabs + panels + filter containers + tables + script)
    $bodyContentHtml = @"
    <!-- Tab Navigation -->
    <div class="rk-tabs">
        <button class="rk-tab active" data-target="panel-license-assignment">License Assignment</button>
        <button class="rk-tab" data-target="panel-subscription-overview">Subscription Overview</button>
        <button class="rk-tab" data-target="panel-disabled-users">Disabled Users</button>
    </div>

    <!-- License Assignment Panel -->
    <div id="panel-license-assignment" class="rk-panel active">
        <div class="rk-filter-bar">
            <span>Filters:</span>
            <select id="accountStatusFilter" class="form-select" style="max-width:180px;">
                <option value="">All Accounts</option>
                <option value="Enabled">Enabled</option>
                <option value="Disabled">Disabled</option>
            </select>
            <select id="assignmentTypeFilter" class="form-select" style="max-width:180px;">
                <option value="">All Types</option>
                <option value="Direct">Direct</option>
                <option value="Inherited">Inherited</option>
                <option value="Both">Both</option>
            </select>
            <select id="licenseNameFilter" class="form-select" style="max-width:220px;">
                <option value="">All Licenses</option>
            </select>
            <button class="rk-filter-chip" onclick="clearLicenseFilters()">Clear</button>
        </div>
        <div class="rk-card">
            <div class="rk-card-header">
                <span>License Assignment</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="licensesShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="licensesTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Display Name</th>
                            <th>User Principal Name</th>
                            <th>Account Status</th>
                            <th>Last Successful Sign In</th>
                            <th>License</th>
                            <th>Assignment Type</th>
                            <th>Inheritance Details</th>
                        </tr>
                    </thead>
                    <tbody>
                        $tableRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Subscription Overview Panel -->
    <div id="panel-subscription-overview" class="rk-panel">
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Subscription Overview</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="subscriptionShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="subscriptionTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Subscription</th>
                            <th>Created Date</th>
                            <th>End Date</th>
                            <th>License Status</th>
                            <th>Consumed Units</th>
                            <th>Total Licenses</th>
                            <th>Available Licenses</th>
                        </tr>
                    </thead>
                    <tbody>
                        $subscriptionRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Disabled Users Panel -->
    <div id="panel-disabled-users" class="rk-panel">
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Disabled Users with Licenses</span>
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
                            <th>Display Name</th>
                            <th>User Principal Name</th>
                            <th>Account Status</th>
                            <th>Last Successful Sign In</th>
                            <th>License</th>
                            <th>Assignment Type</th>
                            <th>Inheritance Details</th>
                        </tr>
                    </thead>
                    <tbody>
                        $disabledUsersRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
    `$(document).ready(function() {
        // Initialize all tables using the shared helper
        var licensesTable = initRKTable('#licensesTable');
        var subscriptionTable = initRKTable('#subscriptionTable');
        var disabledUsersTable = initRKTable('#disabledUsersTable');

        // Populate license name filter dropdown
        function populateLicenseFilter() {
            var values = [...new Set(licensesTable.column(4).data().toArray())].sort();
            var select = `$('#licenseNameFilter');
            values.forEach(function(value) {
                if (value && value.toString().trim() !== '') {
                    // Strip HTML tags for display
                    var text = value.replace(/<[^>]*>/g, '').trim();
                    if (text) {
                        select.append('<option value="' + text + '">' + text + '</option>');
                    }
                }
            });
        }

        // Custom filtering for the licenses table
        `$.fn.dataTable.ext.search.push(function(settings, data, dataIndex) {
            if (settings.nTable.id !== 'licensesTable') return true;

            var accountStatus = `$('#accountStatusFilter').val();
            var assignmentType = `$('#assignmentTypeFilter').val();
            var licenseName = `$('#licenseNameFilter').val();

            var rowAccountStatus = data[2];
            var rowAssignmentType = data[5];
            var rowLicenseName = data[4];

            if (accountStatus && rowAccountStatus.indexOf(accountStatus) === -1) return false;
            if (assignmentType && rowAssignmentType.indexOf(assignmentType) === -1) return false;
            if (licenseName && rowLicenseName.indexOf(licenseName) === -1) return false;

            return true;
        });

        // Apply filters on change
        `$('#accountStatusFilter, #assignmentTypeFilter, #licenseNameFilter').on('change', function() {
            licensesTable.draw();
        });

        // Clear filters
        window.clearLicenseFilters = function() {
            `$('#accountStatusFilter, #assignmentTypeFilter, #licenseNameFilter').val('');
            licensesTable.search('').columns().search('').draw();
        };

        // Show all toggle for licenses table
        `$('#licensesShowAllToggle').on('change', function() {
            licensesTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        // Show all toggle for subscription table
        `$('#subscriptionShowAllToggle').on('change', function() {
            subscriptionTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        // Show all toggle for disabled users table
        `$('#disabledUsersShowAllToggle').on('change', function() {
            disabledUsersTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        // Populate filters after tables are initialized
        setTimeout(function() {
            populateLicenseFilter();
        }, 100);
    });
    </script>
"@

    # Report-specific CSS
    $customCss = @"
    .rk-filter-bar .form-select {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.75rem;
        padding: 4px 8px;
        border-radius: 6px;
    }
    .table-danger td {
        background-color: rgba(192, 57, 43, 0.08) !important;
    }
    [data-theme="dark"] .table-danger td {
        background-color: rgba(224, 96, 80, 0.1) !important;
    }
"@

    # Generate the full HTML report using the shared template
    $htmlContent = New-RKSolutionsReportTemplate `
        -TenantName $Organization `
        -ReportTitle 'License' `
        -ReportSlug 'm365-licenses' `
        -Eyebrow 'M365 LICENSE ASSIGNMENT' `
        -Lede 'License assignment overview including direct, inherited, and disabled user assignments.' `
        -StatsCardsHtml $statsCardsHtml `
        -BodyContentHtml $bodyContentHtml `
        -CustomCss $customCss `
        -ReportDate $CurrentDate `
        -Tags @('M365', 'Licensing', 'Entra ID')

    # Export to HTML file
    $htmlContent | Out-File -FilePath $ExportPath -Encoding utf8

    # Set script-scoped variable for email attachment
    $script:ExportPath = $ExportPath

    Write-Host "All actions completed successfully." -ForegroundColor Cyan
    Write-Host "Report saved to: $ExportPath" -ForegroundColor Cyan

    # Open the HTML file (cross-platform: Invoke-Item uses default handler; fallback so script does not fail in headless env)
    if (-not $SendEmail) {
        try {
            Invoke-Item $ExportPath -ErrorAction Stop
        } catch {
            Write-Host "Report saved to: $ExportPath (could not open automatically)." -ForegroundColor Yellow
        }
    }
}

function Get-LicenseIdentifiers {
    $header = 'Product_Display_Name', 'String_Id', 'GUID', 'Service_Plan_Name', 'Service_Plan_Id', 'Service_Plans_Included_Friendly_Names'
    $params = @{
        Method = 'Get'
        Uri    = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
    }
    $Identifiers = Invoke-RestMethod @params | ConvertFrom-Csv -Header $header |
    ForEach-Object {
        [PSCustomObject]@{
            GUID                 = $_.GUID
            String_Id            = $_.String_Id
            Product_Display_Name = $_.Product_Display_Name
        }
    }
    return $Identifiers | Select-Object -Skip 1
}


function Invoke-M365LicenseReportCore {
    param(
        [Parameter(Mandatory=$false)] [switch] $SendEmail,
        [Parameter(Mandatory=$false)] [string[]] $Recipient,
        [Parameter(Mandatory=$false)] [string] $From,
        [Parameter(Mandatory=$false)] [string] $ExportPath
    )
        # CODE

        # Get Organization Name
        $Organization = Invoke-MgGraphRequest -Uri "beta/organization" -OutputType PSObject | Select-Object -Expand Value | Select-Object -ExpandProperty DisplayName

        # Get product identifiers
        $Identifiers = Get-LicenseIdentifiers

        # Select all SKUs with friendly display name
        [array]$SKU_friendly = $Identifiers | Select-Object GUID, String_Id, Product_Display_Name -Unique

        # NEW CLOUD LICENSING API: Get allotments with subscription details in one call (Beta API)
        # This replaces the previous two separate calls to subscribedSkus and directory/subscriptions
        Write-Host "INFO: Retrieving allotment and subscription data using Cloud Licensing API..." -ForegroundColor Cyan

        try {
            # Try new Cloud Licensing API first
            # Note: subscriptions is included by default, no need to expand
            [Array]$allotments = Invoke-GraphRequestWithPaging -Uri "beta/admin/cloudLicensing/allotments?`$select=id,allottedUnits,consumedUnits,skuId,skuPartNumber,assignableTo,subscriptions"

            if (-not $allotments -or $allotments.Count -eq 0) {
                throw "Allotments API returned empty results"
            }

            $useCloudLicensingAPI = $true
            Write-Host "INFO: Successfully retrieved data from Cloud Licensing API" -ForegroundColor Green

            # Diagnostic: Show what properties are available in first subscription (if verbose)
            if ($VerbosePreference -eq 'Continue' -and $allotments.Count -gt 0) {
                $firstAllotment = $allotments[0]
                if ($firstAllotment.subscriptions -and $firstAllotment.subscriptions.Count -gt 0) {
                    $firstSub = $firstAllotment.subscriptions[0]
                    Write-Verbose "Subscription properties available: $($firstSub.PSObject.Properties.Name -join ', ')"
                }
            }

            # Always show diagnostic info about subscription structure (helps with troubleshooting)
            Write-Host "INFO: Found $($allotments.Count) allotments" -ForegroundColor Cyan
            $totalSubscriptions = ($allotments | ForEach-Object { if ($_.subscriptions) { $_.subscriptions.Count } else { 0 } } | Measure-Object -Sum).Sum
            Write-Host "INFO: Total subscriptions across all allotments: $totalSubscriptions" -ForegroundColor Cyan

            # Supplementary call to get dates from legacy API if needed
            # The allotments API includes startDate and nextLifecycleDate, but we fetch
            # from legacy API as a fallback in case subscription IDs don't match perfectly
            Write-Host "INFO: Retrieving subscription dates as fallback..." -ForegroundColor Cyan
            [Array]$LegacySubscriptions = Invoke-MgGraphRequest -Uri "beta/directory/subscriptions?`$select=id,createdDateTime,nextLifecycleDateTime,skuId" -OutputType PSObject |
                Select-Object -ExpandProperty Value

            # Create lookup table for quick access to both created and end dates
            $subscriptionDateLookup = @{}
            foreach ($legacySub in $LegacySubscriptions) {
                if ($legacySub.id) {
                    $subscriptionDateLookup[$legacySub.id] = @{
                        CreatedDate = $legacySub.createdDateTime
                        EndDate = $legacySub.nextLifecycleDateTime
                    }
                }
            }

            Write-Host "INFO: Created lookup table with $($subscriptionDateLookup.Count) subscription dates" -ForegroundColor Cyan
            if ($VerbosePreference -eq 'Continue' -and $subscriptionDateLookup.Count -gt 0) {
                Write-Verbose "Sample lookup IDs: $(($subscriptionDateLookup.Keys | Select-Object -First 3) -join ', ')"
                $firstId = $subscriptionDateLookup.Keys | Select-Object -First 1
                if ($firstId) {
                    Write-Verbose "Sample data for ID $firstId - CreatedDate: $($subscriptionDateLookup[$firstId].CreatedDate), EndDate: $($subscriptionDateLookup[$firstId].EndDate)"
                }
            }

            # Show how many subscriptions have end dates
            $subsWithEndDates = ($LegacySubscriptions | Where-Object { $_.nextLifecycleDateTime }).Count
            $subsWithoutEndDates = $LegacySubscriptions.Count - $subsWithEndDates
            Write-Host "INFO: Subscriptions with end dates: $subsWithEndDates, without end dates: $subsWithoutEndDates" -ForegroundColor Cyan
        }
        catch {
            # Fallback to legacy API if Cloud Licensing API fails
            Write-Host "WARNING: Cloud Licensing API failed, falling back to legacy API. Error: $($_.Exception.Message)" -ForegroundColor Yellow
            $useCloudLicensingAPI = $false

            # Legacy API calls
            [Array]$Skus = Invoke-MgGraphRequest -Uri "Beta/subscribedSkus" -OutputType PSObject |
            Select-Object -ExpandProperty Value
            [Array]$Subscriptions = Invoke-MgGraphRequest -Uri "beta/directory/subscriptions" -OutputType PSObject |
            Select-Object -ExpandProperty Value
        }

        # Create an overview of subscriptions with their end date
        $SubscriptionOverview = @()

        if ($useCloudLicensingAPI) {
            # NEW: Process allotments from Cloud Licensing API
            $datesFoundCount = 0
            $datesNotFoundCount = 0

            # Group allotments by SKU to combine duplicate licenses
            $allotmentsBySkuId = @{}

            foreach ($allotment in $allotments) {
                # Get friendly name
                $friendlyName = $SKU_friendly | Where-Object { $_.GUID -eq $allotment.skuId } |
                    Select-Object -ExpandProperty Product_Display_Name -ErrorAction SilentlyContinue

                if (-not $friendlyName) {
                    $friendlyName = if ($allotment.skuPartNumber) { $allotment.skuPartNumber } else { "Unknown License ($($allotment.skuId))" }
                }

                # Initialize SKU group if not exists
                if (-not $allotmentsBySkuId.ContainsKey($allotment.skuId)) {
                    $allotmentsBySkuId[$allotment.skuId] = @{
                        FriendlyName = $friendlyName
                        SKUPartNumber = $allotment.skuPartNumber
                        AssignableTo = $allotment.assignableTo
                        TotalLicenses = 0
                        ConsumedUnits = 0
                        CreatedDates = @()
                        EndDates = @()
                        SubscriptionIds = @()
                    }
                }

                # Aggregate license counts
                $allotmentsBySkuId[$allotment.skuId].TotalLicenses += if ($allotment.allottedUnits) { $allotment.allottedUnits } else { 0 }
                $allotmentsBySkuId[$allotment.skuId].ConsumedUnits += if ($allotment.consumedUnits) { $allotment.consumedUnits } else { 0 }

                # Process subscriptions to collect dates
                if ($allotment.subscriptions -and $allotment.subscriptions.Count -gt 0) {
                    foreach ($subscription in $allotment.subscriptions) {
                        if ($subscription.id) {
                            $allotmentsBySkuId[$allotment.skuId].SubscriptionIds += $subscription.id
                        }

                        # Resolve created/start date
                        $subCreated = $null
                        if ($subscription.startDate) { $subCreated = $subscription.startDate }
                        elseif ($subscription.createdDateTime) { $subCreated = $subscription.createdDateTime }
                        elseif ($subscription.createdDate) { $subCreated = $subscription.createdDate }
                        elseif ($subscriptionDateLookup -and $subscription.id -and $subscriptionDateLookup.ContainsKey($subscription.id)) {
                            $subCreated = $subscriptionDateLookup[$subscription.id].CreatedDate
                        }

                        if ($subCreated) {
                            $datesFoundCount++
                            $d = try { [DateTime]$subCreated } catch { $null }
                            if ($d) {
                                $allotmentsBySkuId[$allotment.skuId].CreatedDates += $d
                            }
                        } else {
                            $datesNotFoundCount++
                        }

                        # Resolve end/lifecycle date
                        $subEnd = $null
                        if ($subscription.nextLifecycleDate) { $subEnd = $subscription.nextLifecycleDate }
                        elseif ($subscription.nextLifecycleDateTime) { $subEnd = $subscription.nextLifecycleDateTime }
                        elseif ($subscription.endDate) { $subEnd = $subscription.endDate }
                        elseif ($subscription.expiryDate) { $subEnd = $subscription.expiryDate }
                        elseif ($subscriptionDateLookup -and $subscription.id -and $subscriptionDateLookup.ContainsKey($subscription.id)) {
                            $subEnd = $subscriptionDateLookup[$subscription.id].EndDate
                        }

                        if ($subEnd -and $subEnd -ne "No end date found") {
                            $e = try { [DateTime]$subEnd } catch { $null }
                            if ($e) {
                                $allotmentsBySkuId[$allotment.skuId].EndDates += $e
                            }
                        }
                    }
                }
            }

            # Now create subscription overview with one row per SKU
            foreach ($skuId in $allotmentsBySkuId.Keys) {
                $skuData = $allotmentsBySkuId[$skuId]

                # Get earliest created date
                $createdDate = if ($skuData.CreatedDates.Count -gt 0) {
                    ($skuData.CreatedDates | Measure-Object -Minimum).Minimum
                } else { $null }

                $formattedCreatedDate = if ($createdDate) {
                    try { Get-Date $createdDate -Format "dd-MM-yyyy HH:mm" }
                    catch { $createdDate.ToString() }
                } else { "Unknown" }

                # Get latest end date
                $endDate = if ($skuData.EndDates.Count -gt 0) {
                    ($skuData.EndDates | Measure-Object -Maximum).Maximum
                } else { $null }

                if (-not $endDate) {
                    $endDate = "No end date found"
                }

                $formattedEndDate = if ($endDate -ne "No end date found") {
                    try { Get-Date $endDate -Format "dd-MM-yyyy HH:mm" }
                    catch { $endDate }
                } else { $endDate }

                # Determine license status
                $licenseStatus = "Enabled"
                if ($endDate -ne "No end date found") {
                    try {
                        $dateObj = [DateTime]$endDate
                        $licenseStatus = if ($dateObj -gt (Get-Date)) { "Enabled" } else { "Disabled" }
                    } catch {
                        $licenseStatus = "Unknown"
                    }
                }

                $availableLicenses = $skuData.TotalLicenses - $skuData.ConsumedUnits

                $SubscriptionOverview += [PSCustomObject]@{
                    SubscriptionId    = ($skuData.SubscriptionIds | Select-Object -First 1)
                    FriendlyName      = $skuData.FriendlyName
                    SKUPartNumber     = $skuData.SKUPartNumber
                    CreatedDate       = $formattedCreatedDate
                    EndDate           = $formattedEndDate
                    LicenseStatus     = $licenseStatus
                    ConsumedUnits     = $skuData.ConsumedUnits
                    TotalLicenses     = $skuData.TotalLicenses
                    AvailableLicenses = $availableLicenses
                    AssignableTo      = $skuData.AssignableTo
                }
            }

            # Show summary of date matching
            Write-Host "INFO: Created dates - Found: $datesFoundCount, Not Found: $datesNotFoundCount" -ForegroundColor Cyan
        }
        else {
            # LEGACY: Process subscriptions from old API
            foreach ($subscription in $Subscriptions) {
                $sku = $Skus | Where-Object { $_.SkuId -eq $subscription.SkuId }
                $friendlyName = $SKU_friendly | Where-Object { $_.GUID -eq $sku.SkuId } |
                Select-Object -ExpandProperty Product_Display_Name -ErrorAction SilentlyContinue

                if (-not $friendlyName) {
                    $friendlyName = "Unknown License ($($sku.SkuId))"
                }

                $endDate = if ($null -eq $subscription.NextLifecycleDateTime) {
                    "No end date found"
                } else {
                    $subscription.NextLifecycleDateTime
                }

                # Format dates
                $formattedCreatedDate = if ($subscription.CreatedDateTime -is [DateTime]) {
                    Get-Date $subscription.CreatedDateTime -Format "dd-MM-yyyy HH:mm"
                } elseif ($subscription.CreatedDateTime) {
                    try {
                        Get-Date $subscription.CreatedDateTime -Format "dd-MM-yyyy HH:mm"
                    } catch {
                        $subscription.CreatedDateTime
                    }
                } else {
                    "Unknown"
                }

                $formattedEndDate = if ($endDate -is [DateTime]) {
                    Get-Date $endDate -Format "dd-MM-yyyy HH:mm"
                } elseif ($endDate -and $endDate -ne "No end date found") {
                    try {
                        Get-Date $endDate -Format "dd-MM-yyyy HH:mm"
                    } catch {
                        $endDate
                    }
                } else {
                    $endDate
                }

                # Determine license status
                $licenseStatus = if ($endDate -eq "No end date found") {
                    "Enabled"
                } elseif ($endDate -is [DateTime] -and $endDate -gt (Get-Date)) {
                    "Enabled"
                } elseif ($endDate -ne "No end date found") {
                    try {
                        $dateObj = [DateTime]$endDate
                        if ($dateObj -gt (Get-Date)) {
                            "Enabled"
                        } else {
                            "Disabled"
                        }
                    } catch {
                        "Unknown"
                    }
                } else {
                    "Unknown"
                }

                # Calculate available licenses
                $totalLicenses = if ($subscription.TotalLicenses) { $subscription.TotalLicenses } else { 0 }
                $consumedUnits = if ($sku.ConsumedUnits) { $sku.ConsumedUnits } else { 0 }
                $availableLicenses = $totalLicenses - $consumedUnits

                $SubscriptionOverview += [PSCustomObject]@{
                    SubscriptionId    = $subscription.Id
                    FriendlyName      = $friendlyName
                    CreatedDate       = $formattedCreatedDate
                    EndDate           = $formattedEndDate
                    LicenseStatus     = $licenseStatus
                    ConsumedUnits     = $consumedUnits
                    TotalLicenses     = $totalLicenses
                    AvailableLicenses = $availableLicenses
                }
            }
        }

        # Output the overview
        Write-Host "INFO: Generating subscription overview..." -ForegroundColor Cyan

        # Get all users with licenses - using paging to ensure all results are retrieved
        Write-Host "INFO: Retrieving user license data..." -ForegroundColor Cyan
        $users = Invoke-GraphRequestWithPaging -Uri "beta/users?`$select=UserPrincipalName,LicenseAssignmentStates,DisplayName,AccountEnabled,AssignedLicenses,signInActivity&`$top=999"

        # Get all groups with their licenses
        Write-Host "INFO: Retrieving group license data..." -ForegroundColor Cyan
        $Groups = Invoke-GraphRequestWithPaging -Uri "beta/groups?`$select=id,displayName,assignedLicenses&`$top=999"
        $groupsWithLicenses = @()

        # Loop through each group and check if it has any licenses assigned
        Write-Host "INFO: Checking groups for licenses..." -ForegroundColor Cyan
        foreach ($group in $Groups) {
            if ($group.assignedLicenses -and $group.assignedLicenses.Count -gt 0) {
                $groupData = [PSCustomObject]@{
                    ObjectId    = $group.id
                    DisplayName = $group.displayName
                    Licenses    = $group.assignedLicenses
                }
                $groupsWithLicenses += $groupData
            }
        }

        # Initialize the report array
        $Report = @()

        # Process user license data
        $totalUsers = $users.Count
        $currentIndex = 0

        foreach ($user in $users) {
            $currentIndex++
            Write-Progress -Activity "Processing users" -Status "Processing $currentIndex of $totalUsers" -PercentComplete (($currentIndex / $totalUsers) * 100)

            # Skip users with no license assignment states
            if (-not $user.LicenseAssignmentStates) {
                continue
            }

            # Group licenses by SkuId to detect both direct and inherited assignments
            $licensesBySkuId = @{}

            foreach ($license in $user.LicenseAssignmentStates) {
                $SkuId = $license.SkuId
                $AssignedByGroup = $license.AssignedByGroup

                if (-not $licensesBySkuId.ContainsKey($SkuId)) {
                    $licensesBySkuId[$SkuId] = @{
                        DirectAssignment = $false
                        GroupAssignments = @()
                    }
                }

                if ($null -eq $AssignedByGroup) {
                    $licensesBySkuId[$SkuId].DirectAssignment = $true
                }
                else {
                    $licensesBySkuId[$SkuId].GroupAssignments += $AssignedByGroup
                }
            }

            # Process each unique license
            foreach ($SkuId in $licensesBySkuId.Keys) {
                $licenseInfo = $licensesBySkuId[$SkuId]
                $isDirect = $licenseInfo.DirectAssignment
                $isInherited = ($licenseInfo.GroupAssignments.Count -gt 0)

                # Determine assignment type
                $assignmentType = if ($isDirect -and $isInherited) {
                    "Both"
                }
                elseif ($isDirect) {
                    "Direct"
                }
                elseif ($isInherited) {
                    "Inherited"
                }
                else {
                    "Unknown"
                }

                # Get friendly name for the license
                $friendlyName = $SKU_friendly | Where-Object { $_.GUID -eq $SkuId } |
                Select-Object -ExpandProperty Product_Display_Name -ErrorAction SilentlyContinue

                if (-not $friendlyName) {
                    $friendlyName = "Unknown License ($SkuId)"
                }

                # Get group names if inherited
                $groupNames = ""
                if ($isInherited) {
                    $groupNamesList = @()
                    foreach ($groupId in $licenseInfo.GroupAssignments) {
                        $group = $groupsWithLicenses | Where-Object { $_.ObjectId -eq $groupId }
                        if ($group) {
                            $groupNamesList += $group.DisplayName
                        }
                        else {
                            $groupNamesList += "Unknown Group ($groupId)"
                        }
                    }
                    $groupNames = $groupNamesList -join ", "
                }

                # Determine inheritance description
                if ($isDirect -and -not $groupNames) {
                    $inheritance = "Direct"
                }
                elseif (-not $isDirect -and $groupNames) {
                    $inheritance = $groupNames
                }
                elseif ($isDirect -and $groupNames) {
                    $inheritance = "Direct, $groupNames"
                }
                else {
                    $inheritance = "Unknown"
                }

                # Last Login Activity (robust handling of null/invalid values)
                $lastSignIn = ConvertTo-DateString -Value $user.signInActivity.lastSignInDateTime
                if ($lastSignIn -eq "No sign-in activity" -or $lastSignIn -eq "Invalid date value") {
                    $lastSignIn = ConvertTo-DateString -Value $user.signInActivity.lastSuccessfulSignInDateTime
                }

                # Create the license data object
                $licenseData = [PSCustomObject]@{
                    UserPrincipalName            = $user.UserPrincipalName
                    DisplayName                  = $user.DisplayName
                    AccountEnabled               = if ($user.AccountEnabled) { "Yes" } else { "No" }
                    LastSuccessfulSignIn = $lastSignIn
                    AssignedLicenses             = $SkuId
                    AssignedLicensesFriendlyName = $friendlyName
                    Inheritance                  = $inheritance
                    AssignmentType               = $assignmentType
                    IsDirect                     = $isDirect
                    IsInherited                  = $isInherited
                }

                # Add to the report
                $Report += $licenseData
            }
        }


        # Calculate metrics for summary boxes
        $script:directLicenses = ($Report | Where-Object { $_.IsDirect -eq $true -and $_.IsInherited -eq $false }).Count
        $script:inheritedLicenses = ($Report | Where-Object { $_.IsInherited -eq $true -and $_.IsDirect -eq $false }).Count
        $script:bothLicenses = ($Report | Where-Object { $_.IsDirect -eq $true -and $_.IsInherited -eq $true }).Count
        $script:DisabledUsersWithLicenses = ($Report | Where-Object { $_.AccountEnabled -eq "No" } | Select-Object -Unique UserPrincipalName).Count

        # Output summary information
        Write-Host "INFO: License Summary:" -ForegroundColor Cyan
        Write-Host "Total users processed: $totalUsers" -ForegroundColor White
        Write-Host "Users with licenses: $($Report | Select-Object -Unique UserPrincipalName | Measure-Object | Select-Object -ExpandProperty Count)" -ForegroundColor White
        Write-Host "Direct license assignments: $script:directLicenses" -ForegroundColor White
        Write-Host "Inherited license assignments: $script:inheritedLicenses" -ForegroundColor White
        Write-Host "Both direct and inherited: $script:bothLicenses" -ForegroundColor White
        Write-Host "Disabled users with licenses: $script:DisabledUsersWithLicenses" -ForegroundColor White

        # Export to HTML
        New-HTMLReport -Organization $Organization -Report $Report -SubscriptionOverview $SubscriptionOverview -ExportPath $ExportPath

        # Send email with the report
        $emailSent = $false
        if ($SendEmail) {
            $subject = "$Organization - Microsoft 365 License Assignment Report"
            $bodyHtml = "<html><body style='font-family: Segoe UI, Arial, sans-serif;'><h2>Microsoft 365 License Assignment Report</h2><p>Attached is the latest Microsoft 365 license assignment report for $Organization.</p><p>Open the attached HTML in a browser for the full report.</p><p style='color:#666;'>Generated by RKSolutions - please do not reply.</p></body></html>"
            $emailSent = Send-EmailWithAttachment -Recipient $Recipient -AttachmentPath $script:ExportPath -From $From -Subject $subject -BodyHtml $bodyHtml

            if ($emailSent) {
                Write-Host "INFO: Email sent successfully." -ForegroundColor Green
            }
            else {
                Write-Host "ERROR: Failed to send email." -ForegroundColor Red
            }
        }
        else {
            Write-Host "INFO: Email sending is disabled. Set -SendEmail to `$true to enable." -ForegroundColor Yellow
        }

        # Clean up the report file
        if ($SendEmail -and $emailSent) {
            if (Test-Path -Path $script:ExportPath) {
                Remove-Item -Path $script:ExportPath -Force
                Write-Host "INFO: Temporary report file deleted." -ForegroundColor Green
            } else {
                Write-Host "INFO: No temporary report file found to delete." -ForegroundColor Yellow
            }
        }
}
