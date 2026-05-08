# Entra Admin Roles - Private helpers

function New-AdminRoleHTMLReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantName,

        [Parameter(Mandatory = $true)]
        [array]$Report,

        [Parameter(Mandatory = $false)]
        [array]$GroupAssignmentReport,

        [Parameter(Mandatory = $false)]
        [array]$ServicePrincipalReport,

        [Parameter(Mandatory = $false)]
        [array]$UserAssignmentReport,

        [Parameter(Mandatory = $false)]
        [array]$GroupMembershipOverviewReport,

        [Parameter(Mandatory = $false)]
        [array]$PIMAuditLogsReport,

        [Parameter(Mandatory = $false)]
        [string]$ExportPath
    )

    # Default ExportPath to current folder if not provided
    if (-not $ExportPath) {
        $ExportPath = Join-Path (Get-Location).Path "$TenantName-AdminRolesReport.html"
    }


    # Calculate roles counts for dashboard statistics
    $permanentRoles = ($Report | Where-Object { $_.AssignmentType -eq "Permanent" }).Count
    $eligibleRoles = ($Report | Where-Object { $_.AssignmentType -like "Eligible*" }).Count
    $groupAssignedRoles = $GroupAssignmentReport.Count
    $servicePrincipalRoles = $ServicePrincipalReport.Count

    # Get the current date and time for the report header
    $CurrentDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

    # Generate table rows for all role assignments
    $allRolesRows = ""
    foreach ($item in $Report) {
        $assignmentTypeBadge = switch ($item.AssignmentType) {
            "Permanent" { "<span class=`"rk-badge rk-badge-error`">Permanent</span>" }
            "Eligible" { "<span class=`"rk-badge rk-badge-ok`">Eligible</span>" }
            "Eligible (Active)" {
                if ($item.PrincipalType -eq "group") {
                    # Create a safe ID from the principal name and role
                    $safeId = ($item.Principal + "-" + $item.'Assigned Role').Replace(" ", "-").Replace("@", "-").Replace(".", "-")
                    "<span class=`"rk-badge rk-badge-ok badge-eligible-active group-jump-link`" data-group-id=`"$safeId`" style=`"cursor: pointer;`" title=`"Click to view group details in Group Assignments tab`">Eligible (Active) <i class=`"fas fa-external-link-alt`" style=`"font-size: 10px; margin-left: 4px;`"></i></span>"
                } else {
                    "<span class=`"rk-badge rk-badge-ok badge-eligible-active`">Eligible (Active)</span>"
                }
            }
            default { "<span class=`"rk-badge rk-badge-na`">Unknown</span>" }
        }

        $allRolesRows += @"
        <tr>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Principal))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DisplayName))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.PrincipalType))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.AccountStatus))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.'Assigned Role'))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.AssignedRoleScopeName))</td>
            <td>$assignmentTypeBadge</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.AssignmentStartDate))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.AssignmentEndDate))</td>
        </tr>
"@
    }

    # Generate table rows for user role assignments
    $userRolesRows = ""
    foreach ($item in $UserAssignmentReport) {
        $assignmentTypeBadge = switch ($item.AssignmentType) {
            "Permanent" { "<span class=`"rk-badge rk-badge-error`">Permanent</span>" }
            "Eligible" { "<span class=`"rk-badge rk-badge-ok`">Eligible</span>" }
            "Eligible (Active)" { "<span class=`"rk-badge rk-badge-ok badge-eligible-active`">Eligible (Active)</span>" }
            default { "<span class=`"rk-badge rk-badge-na`">Unknown</span>" }
        }

        $userRolesRows += @"
    <tr>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.Principal))</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.DisplayName))</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.PrincipalType))</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.AccountStatus))</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.'Assigned Role'))</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.AssignedRoleScopeName))</td>
        <td>$assignmentTypeBadge</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.AssignmentStartDate))</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.AssignmentEndDate))</td>
    </tr>
"@
    }

    # Generate table rows for group role assignments
    $groupRolesRows = ""
    foreach ($item in $GroupAssignmentReport) {
        $assignmentTypeBadge = switch ($item.AssignmentType) {
            "Permanent" { "<span class=`"rk-badge rk-badge-error`">Permanent</span>" }
            "Eligible" { "<span class=`"rk-badge rk-badge-ok`">Eligible</span>" }
            "Eligible (Active)" { "<span class=`"rk-badge rk-badge-ok badge-eligible-active`">Eligible (Active)</span>" }
            default { "<span class=`"rk-badge rk-badge-na`">Unknown</span>" }
        }

        # Get group members from the overview report
        $groupMembers = ($GroupMembershipOverviewReport | Where-Object { $_.Principal -eq $item.Principal }).Members
        if (-not $groupMembers) {
            $groupMembers = "None"
        }

        # Format activated members information (simplified - UserPrincipalName only)
        $activatedMembersText = "None"
        if ($item.ActivatedMembers -and @($item.ActivatedMembers).Count -gt 0) {
            $activatedList = @()
            foreach ($activatedMember in $item.ActivatedMembers) {
                $activatedList += [System.Net.WebUtility]::HtmlEncode($activatedMember.UserPrincipalName)
            }
            $activatedMembersText = $activatedList -join "<br/>"
        }

        # Create a safe ID from the principal name and role for targeting
        $safeId = ($item.Principal + "-" + $item.'Assigned Role').Replace(" ", "-").Replace("@", "-").Replace(".", "-")

        $groupRolesRows += @"
        <tr id="group-$safeId" class="group-assignment-row">
            <td>$([System.Net.WebUtility]::HtmlEncode($item.Principal))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.DisplayName))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.PrincipalType))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.AccountStatus))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.'Assigned Role'))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.AssignedRoleScopeName))</td>
            <td>$assignmentTypeBadge</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.AssignmentStartDate))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($item.AssignmentEndDate))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($groupMembers))</td>
            <td style="max-width: 300px; word-wrap: break-word;">$activatedMembersText</td>
        </tr>
"@
    }

    # Generate table rows for service principal role assignments
    $spRolesRows = ""
    foreach ($item in $ServicePrincipalReport) {
        $assignmentTypeBadge = switch ($item.AssignmentType) {
            "Permanent" { "<span class=`"rk-badge rk-badge-error`">Permanent</span>" }
            "Eligible" { "<span class=`"rk-badge rk-badge-ok`">Eligible</span>" }
            "Eligible (Active)" { "<span class=`"rk-badge rk-badge-ok badge-eligible-active`">Eligible (Active)</span>" }
            default { "<span class=`"rk-badge rk-badge-na`">Unknown</span>" }
        }

        $spRolesRows += @"
    <tr>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.Principal))</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.DisplayName))</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.PrincipalType))</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.AccountStatus))</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.'Assigned Role'))</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.AssignedRoleScopeName))</td>
        <td>$assignmentTypeBadge</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.AssignmentStartDate))</td>
        <td>$([System.Net.WebUtility]::HtmlEncode($item.AssignmentEndDate))</td>
    </tr>
"@
    }

    # Generate table rows for PIM audit logs
    $pimAuditLogsRows = ""
    if ($PIMAuditLogsReport -and $PIMAuditLogsReport.Count -gt 0) {
        foreach ($log in $PIMAuditLogsReport) {
            $resultBadge = switch ($log.Result) {
                "Success" { "<span class=`"rk-badge rk-badge-ok`">Success</span>" }
                "Failure" { "<span class=`"rk-badge rk-badge-error`">Failure</span>" }
                default { "<span class=`"rk-badge rk-badge-na`">Unknown</span>" }
            }

            $pimAuditLogsRows += @"
        <tr>
            <td>$([System.Net.WebUtility]::HtmlEncode($log.DateTime))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($log.InitiatedBy))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($log.OperationType))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($log.InitiatedByType))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($log.Role))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($log.Target))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($log.Operation))</td>
            <td>$resultBadge</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($log.RoleProperties))</td>
            <td>$([System.Net.WebUtility]::HtmlEncode($log.Justification))</td>
        </tr>
"@
        }
    }

    # Build stat tiles HTML
    $statsCardsHtml = @"
            <div class="rk-stat-tile t-rust">
                <div class="rk-stat-eyebrow">PERMANENT</div>
                <div class="rk-stat-number">$permanentRoles</div>
                <div class="rk-stat-caption">Permanent assignments</div>
            </div>
            <div class="rk-stat-tile t-olive">
                <div class="rk-stat-eyebrow">ELIGIBLE</div>
                <div class="rk-stat-number">$eligibleRoles</div>
                <div class="rk-stat-caption">Eligible assignments</div>
            </div>
            <div class="rk-stat-tile t-steel">
                <div class="rk-stat-eyebrow">GROUP</div>
                <div class="rk-stat-number">$groupAssignedRoles</div>
                <div class="rk-stat-caption">Group assignments</div>
            </div>
            <div class="rk-stat-tile t-rose">
                <div class="rk-stat-eyebrow">SERVICE PRINCIPAL</div>
                <div class="rk-stat-number">$servicePrincipalRoles</div>
                <div class="rk-stat-caption">Service principal assignments</div>
            </div>
"@

    # Build dynamic tab buttons
    $tabButtons = @"
        <button class="rk-tab active" data-target="panel-all-roles">All Assignments</button>
"@

    if ($UserAssignmentReport.Count -gt 0) {
        $tabButtons += @"
        <button class="rk-tab" data-target="panel-user-roles">User Assignments</button>
"@
    }

    if ($GroupAssignmentReport.Count -gt 0) {
        $tabButtons += @"
        <button class="rk-tab" data-target="panel-group-roles">Group Assignments</button>
"@
    }

    if ($ServicePrincipalReport.Count -gt 0) {
        $tabButtons += @"
        <button class="rk-tab" data-target="panel-sp-roles">Service Principal Assignments</button>
"@
    }

    if ($PIMAuditLogsReport -and $PIMAuditLogsReport.Count -gt 0) {
        $tabButtons += @"
        <button class="rk-tab" data-target="panel-pim-audit-logs">PIM Audit Logs</button>
"@
    }

    # Build body content HTML
    $bodyContentHtml = @"
    <!-- Filter Bar -->
    <div class="rk-filter-bar" id="general-filter-section">
        <span>Filters:</span>
        <select id="principalTypeFilter" class="form-select" style="max-width:180px;">
            <option value="">All Principal Types</option>
            <option value="user">User</option>
            <option value="group">Group</option>
            <option value="service Principal">Service Principal</option>
        </select>
        <select id="assignmentTypeFilter" class="form-select" style="max-width:180px;">
            <option value="">All Assignment Types</option>
            <option value="Permanent">Permanent</option>
            <option value="Eligible">Eligible</option>
        </select>
        <input type="text" id="roleNameFilter" class="form-control" placeholder="Search role names..." style="max-width:200px;">
        <select id="scopeFilter" class="form-select" style="max-width:180px;">
            <option value="">All Scopes</option>
            <option value="Tenant-Wide">Tenant-Wide</option>
            <option value="AU/">Administrative Unit</option>
        </select>
        <button class="rk-filter-chip" onclick="clearAllFilters()">Clear</button>
        <div class="rk-filter-tags" id="enabledFilters"></div>
    </div>

    <!-- PIM Audit Logs Filter Bar (hidden by default) -->
    <div class="rk-filter-bar" id="pim-filter-section" style="display:none;">
        <span>PIM Filters:</span>
        <select id="pimOperationTypeFilter" class="form-select" style="max-width:180px;">
            <option value="">All Operations</option>
            <option value="Add">Add</option>
            <option value="Update">Update</option>
            <option value="Delete">Delete</option>
            <option value="Activate">Activate</option>
        </select>
        <input type="text" id="pimInitiatorFilter" class="form-control" placeholder="Filter by user..." style="max-width:180px;">
        <select id="pimResultFilter" class="form-select" style="max-width:180px;">
            <option value="">All Results</option>
            <option value="Success">Success</option>
            <option value="Failure">Failure</option>
        </select>
        <input type="text" id="pimRoleFilter" class="form-control" placeholder="Filter by role..." style="max-width:180px;">
        <input type="text" id="pimTargetFilter" class="form-control" placeholder="Filter by target..." style="max-width:180px;">
        <div style="display:flex;align-items:center;gap:4px;">
            <input type="date" id="pimStartDateFilter" class="form-control" style="max-width:140px;">
            <span>to</span>
            <input type="date" id="pimEndDateFilter" class="form-control" style="max-width:140px;">
        </div>
        <button class="rk-filter-chip" onclick="clearPimFilters()">Clear</button>
        <div class="rk-filter-tags" id="pimEnabledFilters"></div>
    </div>

    <!-- Tab Navigation -->
    <div class="rk-tabs">
$tabButtons
    </div>

    <!-- All Roles Panel -->
    <div id="panel-all-roles" class="rk-panel active">
        <div class="rk-card">
            <div class="rk-card-header">
                <span>All Role Assignments</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="allShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="allRolesTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Principal</th>
                            <th>Display Name</th>
                            <th>Principal Type</th>
                            <th>Account Status</th>
                            <th>Assigned Role</th>
                            <th>Role Scope</th>
                            <th>Assignment Type</th>
                            <th>Start Date</th>
                            <th>End Date</th>
                        </tr>
                    </thead>
                    <tbody>
                        $allRolesRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- User Roles Panel -->
    <div id="panel-user-roles" class="rk-panel">
        <div class="rk-card">
            <div class="rk-card-header">
                <span>User Role Assignments</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="userShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="userRolesTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Principal</th>
                            <th>Display Name</th>
                            <th>Principal Type</th>
                            <th>Account Status</th>
                            <th>Assigned Role</th>
                            <th>Role Scope</th>
                            <th>Assignment Type</th>
                            <th>Start Date</th>
                            <th>End Date</th>
                        </tr>
                    </thead>
                    <tbody>
                        $userRolesRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Group Roles Panel -->
    <div id="panel-group-roles" class="rk-panel">
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Group Role Assignments</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="groupShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="groupRolesTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Principal</th>
                            <th>Display Name</th>
                            <th>Principal Type</th>
                            <th>Account Status</th>
                            <th>Assigned Role</th>
                            <th>Role Scope</th>
                            <th>Assignment Type</th>
                            <th>Start Date</th>
                            <th>End Date</th>
                            <th>Members</th>
                            <th>Activated Members</th>
                        </tr>
                    </thead>
                    <tbody>
                        $groupRolesRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Service Principal Roles Panel -->
    <div id="panel-sp-roles" class="rk-panel">
        <div class="rk-card">
            <div class="rk-card-header">
                <span>Service Principal Role Assignments</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="spShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="spRolesTable" class="table table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th>Principal</th>
                            <th>Display Name</th>
                            <th>Principal Type</th>
                            <th>Account Status</th>
                            <th>Assigned Role</th>
                            <th>Role Scope</th>
                            <th>Assignment Type</th>
                            <th>Start Date</th>
                            <th>End Date</th>
                        </tr>
                    </thead>
                    <tbody>
                        $spRolesRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- PIM Audit Logs Panel -->
    <div id="panel-pim-audit-logs" class="rk-panel">
        <div class="rk-card">
            <div class="rk-card-header">
                <span>PIM Audit Logs</span>
                <div class="rk-show-all">
                    <label class="rk-toggle-switch">
                        <input type="checkbox" id="pimAuditLogsShowAllToggle">
                        <span class="rk-toggle-slider"></span>
                    </label>
                    <span>Show all</span>
                </div>
            </div>
            <div class="rk-card-body">
                <table id="pimAuditLogsTable" class="table table-bordered" style="width:100%; table-layout:fixed;">
                    <thead>
                        <tr>
                            <th style="width:10%">Date/Time</th>
                            <th style="width:14%">Initiated By</th>
                            <th style="width:7%">Operation Type</th>
                            <th style="width:7%">Initiator Type</th>
                            <th style="width:12%">Role</th>
                            <th style="width:14%">Target</th>
                            <th style="width:12%">Operation</th>
                            <th style="width:6%">Result</th>
                            <th style="width:10%">Role Properties</th>
                            <th style="width:8%">Justification</th>
                        </tr>
                    </thead>
                    <tbody>
                        $pimAuditLogsRows
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
    `$(document).ready(function() {
        // Initialize all tables using the shared helper
        var allRolesTable = initRKTable('#allRolesTable', { order: [[3, 'asc']] });
        var userRolesTable = initRKTable('#userRolesTable', { order: [[3, 'asc']] });
        var groupRolesTable = initRKTable('#groupRolesTable', { order: [[3, 'asc']] });
        var spRolesTable = initRKTable('#spRolesTable', { order: [[3, 'asc']] });
        var pimAuditLogsTable = initRKTable('#pimAuditLogsTable', { order: [[0, 'desc']] });

        // Toggle filter sections based on active tab
        `$(document).on('click', '.rk-tab', function() {
            var target = `$(this).data('target');
            if (target === 'panel-pim-audit-logs') {
                `$('#general-filter-section').hide();
                `$('#pim-filter-section').show();
            } else {
                `$('#general-filter-section').show();
                `$('#pim-filter-section').hide();
            }
            // Adjust DataTables columns when switching tabs
            setTimeout(function() {
                `$.fn.dataTable.tables({ visible: true, api: true }).columns.adjust();
            }, 10);
        });

        // Custom filtering function for all tables (general filters)
        `$.fn.dataTable.ext.search.push(
            function(settings, data, dataIndex) {
                // Skip PIM audit logs table for general filters
                if (settings.nTable.id === 'pimAuditLogsTable') {
                    return true;
                }

                var principalType = `$('#principalTypeFilter').val().toLowerCase();
                var assignmentType = `$('#assignmentTypeFilter').val();
                var roleName = `$('#roleNameFilter').val().toLowerCase();
                var scopeFilter = `$('#scopeFilter').val();

                var colPrincipalType = data[2].toLowerCase();
                var colRole = data[4].toLowerCase();
                var colAssignmentType = data[6];
                var colScope = data[5];

                if (principalType && !colPrincipalType.includes(principalType)) return false;
                if (assignmentType && !colAssignmentType.includes(assignmentType)) return false;
                if (roleName && !colRole.includes(roleName)) return false;
                if (scopeFilter && !colScope.toLowerCase().includes(scopeFilter.toLowerCase())) return false;

                return true;
            }
        );

        // Custom filtering function for PIM audit logs
        `$.fn.dataTable.ext.search.push(
            function(settings, data, dataIndex) {
                if (settings.nTable.id !== 'pimAuditLogsTable') return true;

                var operationType = `$('#pimOperationTypeFilter').val().toLowerCase();
                var initiator = `$('#pimInitiatorFilter').val().toLowerCase();
                var result = `$('#pimResultFilter').val().toLowerCase();
                var role = `$('#pimRoleFilter').val().toLowerCase();
                var target = `$('#pimTargetFilter').val().toLowerCase();
                var startDate = `$('#pimStartDateFilter').val();
                var endDate = `$('#pimEndDateFilter').val();

                var colDateTime = data[0];
                var colInitiator = data[1].toLowerCase();
                var colOperation = data[2].toLowerCase();
                var colRole = data[4].toLowerCase();
                var colTarget = data[5].toLowerCase();
                var colResult = data[7].toLowerCase();

                if (operationType && !colOperation.includes(operationType)) return false;
                if (initiator && !colInitiator.includes(initiator)) return false;
                if (result && !colResult.includes(result)) return false;
                if (role && !colRole.includes(role)) return false;
                if (target && !colTarget.includes(target)) return false;

                // Date range filtering
                if (startDate || endDate) {
                    var rowDate = null;
                    try {
                        var dateMatch = colDateTime.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
                        if (dateMatch) {
                            rowDate = new Date(dateMatch[3], dateMatch[1] - 1, dateMatch[2]);
                        } else {
                            rowDate = new Date(colDateTime);
                        }
                    } catch (e) { rowDate = null; }

                    if (rowDate) {
                        if (startDate) {
                            var filterStartDate = new Date(startDate);
                            if (rowDate < filterStartDate) return false;
                        }
                        if (endDate) {
                            var filterEndDate = new Date(endDate);
                            filterEndDate.setHours(23, 59, 59, 999);
                            if (rowDate > filterEndDate) return false;
                        }
                    }
                }

                return true;
            }
        );

        // Apply general filters
        function applyFilters() {
            allRolesTable.draw();
            userRolesTable.draw();
            groupRolesTable.draw();
            spRolesTable.draw();
        }

        // Apply PIM filters
        function applyPimFilters() {
            pimAuditLogsTable.draw();
        }

        // General filter change handlers
        `$('#principalTypeFilter, #assignmentTypeFilter, #scopeFilter').on('change', function() {
            var filterType = `$(this).attr('id');
            var filterValue = `$(this).val();
            if (filterValue) {
                if (filterType === 'principalTypeFilter') updateEnabledFilters('Principal Type', filterValue);
                else if (filterType === 'assignmentTypeFilter') updateEnabledFilters('Assignment Type', filterValue);
                else if (filterType === 'scopeFilter') updateEnabledFilters('Scope', filterValue);
            } else {
                if (filterType === 'principalTypeFilter') removeEnabledFilter('Principal Type');
                else if (filterType === 'assignmentTypeFilter') removeEnabledFilter('Assignment Type');
                else if (filterType === 'scopeFilter') removeEnabledFilter('Scope');
            }
            applyFilters();
        });

        `$('#roleNameFilter').on('input', function() {
            var filterValue = `$(this).val();
            if (filterValue) {
                updateEnabledFilters('Role Name', filterValue);
            } else {
                removeEnabledFilter('Role Name');
            }
            applyFilters();
        });

        // PIM filter change handlers
        `$('#pimOperationTypeFilter, #pimResultFilter').on('change', function() {
            var filterValue = `$(this).val();
            var filterLabel = `$(this).attr('id') === 'pimOperationTypeFilter' ? 'Operation Type' : 'Result';
            if (filterValue) {
                updatePimEnabledFilter(filterLabel, filterValue);
            } else {
                removePimEnabledFilter(filterLabel);
            }
            applyPimFilters();
        });

        `$('#pimInitiatorFilter, #pimRoleFilter, #pimTargetFilter').on('input', function() {
            var filterValue = `$(this).val();
            var filterLabel;
            if (`$(this).attr('id') === 'pimInitiatorFilter') filterLabel = 'Initiated By';
            else if (`$(this).attr('id') === 'pimRoleFilter') filterLabel = 'Role';
            else filterLabel = 'Target User';
            if (filterValue) {
                updatePimEnabledFilter(filterLabel, filterValue);
            } else {
                removePimEnabledFilter(filterLabel);
            }
            clearTimeout(`$(this).data('timeout'));
            `$(this).data('timeout', setTimeout(function() { applyPimFilters(); }, 300));
        });

        `$('#pimStartDateFilter, #pimEndDateFilter').on('change', function() {
            var startDate = `$('#pimStartDateFilter').val();
            var endDate = `$('#pimEndDateFilter').val();
            if (startDate || endDate) {
                var dateRangeText = '';
                if (startDate && endDate) dateRangeText = startDate + ' to ' + endDate;
                else if (startDate) dateRangeText = 'From ' + startDate;
                else if (endDate) dateRangeText = 'Until ' + endDate;
                updatePimEnabledFilter('Date Range', dateRangeText);
            } else {
                removePimEnabledFilter('Date Range');
            }
            applyPimFilters();
        });

        // Clear all general filters
        window.clearAllFilters = function() {
            `$('#principalTypeFilter').val('');
            `$('#assignmentTypeFilter').val('');
            `$('#roleNameFilter').val('');
            `$('#scopeFilter').val('');
            `$('#enabledFilters').empty();
            allRolesTable.search('').columns().search('').draw();
            userRolesTable.search('').columns().search('').draw();
            groupRolesTable.search('').columns().search('').draw();
            spRolesTable.search('').columns().search('').draw();
            applyFilters();
            // Switch back to All Assignments tab
            `$('.rk-tab[data-target="panel-all-roles"]').click();
        };

        // Clear PIM filters
        window.clearPimFilters = function() {
            `$('#pimOperationTypeFilter').val('');
            `$('#pimInitiatorFilter').val('');
            `$('#pimResultFilter').val('');
            `$('#pimRoleFilter').val('');
            `$('#pimTargetFilter').val('');
            `$('#pimStartDateFilter').val('');
            `$('#pimEndDateFilter').val('');
            `$('#pimEnabledFilters').empty();
            applyPimFilters();
        };

        // Enabled filter tag management (general)
        function updateEnabledFilters(filterType, filterValue) {
            removeEnabledFilter(filterType);
            var filterTag = '<div class="rk-filter-tag" data-filter-type="' + filterType + '">' +
                '<span>' + filterType + ': ' + filterValue + '</span>' +
                '<i class="fas fa-times-circle remove-filter" data-filter-type="' + filterType + '"></i>' +
                '</div>';
            `$('#enabledFilters').append(filterTag);
            `$('.remove-filter').off('click').on('click', function() {
                var ft = `$(this).data('filter-type');
                if (ft === 'Principal Type') `$('#principalTypeFilter').val('');
                else if (ft === 'Assignment Type') `$('#assignmentTypeFilter').val('');
                else if (ft === 'Role Name') `$('#roleNameFilter').val('');
                else if (ft === 'Scope') `$('#scopeFilter').val('');
                `$(this).closest('.rk-filter-tag').remove();
                applyFilters();
            });
        }

        function removeEnabledFilter(filterType) {
            `$('.rk-filter-tag[data-filter-type="' + filterType + '"]').remove();
        }

        // Enabled filter tag management (PIM)
        function updatePimEnabledFilter(filterType, filterValue) {
            removePimEnabledFilter(filterType);
            var filterTag = '<div class="rk-filter-tag" data-pim-filter-type="' + filterType + '">' +
                '<span>' + filterType + ': ' + filterValue + '</span>' +
                '<i class="fas fa-times-circle remove-pim-filter" data-pim-filter-type="' + filterType + '"></i>' +
                '</div>';
            `$('#pimEnabledFilters').append(filterTag);
            `$('.remove-pim-filter').off('click').on('click', function() {
                var ft = `$(this).data('pim-filter-type');
                if (ft === 'Operation Type') `$('#pimOperationTypeFilter').val('');
                else if (ft === 'Initiated By') `$('#pimInitiatorFilter').val('');
                else if (ft === 'Result') `$('#pimResultFilter').val('');
                else if (ft === 'Role') `$('#pimRoleFilter').val('');
                else if (ft === 'Target User') `$('#pimTargetFilter').val('');
                else if (ft === 'Date Range') { `$('#pimStartDateFilter').val(''); `$('#pimEndDateFilter').val(''); }
                `$(this).closest('.rk-filter-tag').remove();
                applyPimFilters();
            });
        }

        function removePimEnabledFilter(filterType) {
            `$('.rk-filter-tag[data-pim-filter-type="' + filterType + '"]').remove();
        }

        // Show all toggle functionality
        `$('#allShowAllToggle').on('change', function() {
            allRolesTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });
        `$('#userShowAllToggle').on('change', function() {
            userRolesTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });
        `$('#groupShowAllToggle').on('change', function() {
            groupRolesTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });
        `$('#spShowAllToggle').on('change', function() {
            spRolesTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });
        `$('#pimAuditLogsShowAllToggle').on('change', function() {
            pimAuditLogsTable.page.len(`$(this).is(':checked') ? -1 : 10).draw();
        });

        // Group jump functionality - handle clicks on group links in main table
        `$(document).on('click', '.group-jump-link', function(e) {
            e.preventDefault();

            var groupId = `$(this).data('group-id');
            var targetRow = `$('#group-' + groupId);

            if (targetRow.length > 0) {
                // Switch to Group Assignments tab
                `$('.rk-tab').removeClass('active');
                `$('.rk-panel').removeClass('active');
                `$('.rk-tab[data-target="panel-group-roles"]').addClass('active');
                `$('#panel-group-roles').addClass('active');

                // Show general filter section, hide PIM filter section
                `$('#general-filter-section').show();
                `$('#pim-filter-section').hide();

                // Adjust DataTables columns for the newly visible tab
                setTimeout(function() {
                    groupRolesTable.columns.adjust();

                    // Scroll to the target row and highlight it
                    var targetRowElement = targetRow[0];
                    if (targetRowElement) {
                        targetRowElement.scrollIntoView({
                            behavior: 'smooth',
                            block: 'center'
                        });

                        // Add highlight effect
                        targetRow.addClass('highlight-row');

                        // Remove highlight after 3 seconds
                        setTimeout(function() {
                            targetRow.removeClass('highlight-row');
                        }, 3000);
                    }
                }, 100);
            } else {
                console.warn('Target group row not found:', groupId);
            }
        });
    });
    </script>
"@

    # Report-specific CSS (shimmer animation for eligible-active badges, highlight-pulse for group jumps)
    $customCss = @"
    /* Shimmer animation for eligible-active badges */
    .badge-eligible-active {
        box-shadow: 0 0 8px rgba(45, 122, 58, 0.6), 0 0 16px rgba(45, 122, 58, 0.4);
        border: 2px solid rgba(255, 255, 255, 0.3);
        animation: shimmer 2s infinite;
    }

    .badge-eligible-active.group-jump-link {
        transition: all 0.3s ease;
        cursor: pointer;
        position: relative;
    }

    .badge-eligible-active.group-jump-link:hover {
        box-shadow: 0 0 12px rgba(45, 122, 58, 0.8), 0 0 24px rgba(45, 122, 58, 0.6);
        transform: translateY(-1px);
        border-color: rgba(255, 255, 255, 0.5);
    }

    .badge-eligible-active.group-jump-link:active {
        transform: translateY(0);
        box-shadow: 0 0 6px rgba(45, 122, 58, 0.4), 0 0 12px rgba(45, 122, 58, 0.3);
    }

    @keyframes shimmer {
        0% { box-shadow: 0 0 8px rgba(45, 122, 58, 0.6), 0 0 16px rgba(45, 122, 58, 0.4); }
        50% { box-shadow: 0 0 12px rgba(45, 122, 58, 0.8), 0 0 24px rgba(45, 122, 58, 0.6); }
        100% { box-shadow: 0 0 8px rgba(45, 122, 58, 0.6), 0 0 16px rgba(45, 122, 58, 0.4); }
    }

    /* Highlight effect for group rows when jumping from main table */
    .highlight-row {
        background-color: rgba(255, 193, 7, 0.3) !important;
        border: 2px solid #ffc107 !important;
        animation: highlightPulse 3s ease-in-out;
        transition: all 0.3s ease;
        position: relative;
        z-index: 10;
    }

    .highlight-row td {
        position: relative;
        z-index: 10;
    }

    @keyframes highlightPulse {
        0% {
            background-color: rgba(255, 193, 7, 0.6);
            box-shadow: 0 0 15px rgba(255, 193, 7, 0.8);
        }
        50% {
            background-color: rgba(255, 193, 7, 0.4);
            box-shadow: 0 0 25px rgba(255, 193, 7, 0.6);
        }
        100% {
            background-color: rgba(255, 193, 7, 0.3);
            box-shadow: 0 0 10px rgba(255, 193, 7, 0.4);
        }
    }

    /* Filter bar form controls */
    .rk-filter-bar .form-select,
    .rk-filter-bar .form-control {
        font-family: 'Geist Mono', ui-monospace, monospace;
        font-size: 0.75rem;
        padding: 4px 8px;
        border-radius: 6px;
    }

    /* PIM Audit Logs: constrain long-content columns */
    #pimAuditLogsTable td {
        word-break: break-word;
        overflow-wrap: break-word;
        vertical-align: top;
    }
    #pimAuditLogsTable td:nth-child(9),
    #pimAuditLogsTable td:nth-child(10) {
        font-size: 0.7rem;
        max-height: 6em;
        overflow: hidden;
        text-overflow: ellipsis;
    }
"@

    # Generate the full HTML report using the shared template
    $htmlContent = Get-RKSolutionsReportTemplate `
        -TenantName $TenantName `
        -ReportTitle 'Admin Roles' `
        -ReportSlug 'entra-admin-roles' `
        -Eyebrow 'ENTRA ADMIN ROLES' `
        -Lede 'Privileged role assignments including permanent, eligible, group-based, and service principal assignments.' `
        -StatsCardsHtml $statsCardsHtml `
        -BodyContentHtml $bodyContentHtml `
        -CustomCss $customCss `
        -ReportDate $CurrentDate `
        -Tags @('Entra ID', 'PIM', 'Security')

    # Export to HTML file
    $htmlContent | Out-File -FilePath $ExportPath -Encoding utf8

    # Set script-scoped variable for email attachment
    $script:ExportPath = $ExportPath

    Write-Host "INFO: All actions completed successfully."
    Write-Host "INFO: Admin Roles Report saved to: $ExportPath" -ForegroundColor Cyan

    # Open the HTML file only if we're not sending email
    if (-not $SendEmail) {
        Invoke-Item $ExportPath
    }
}


function Get-SecurityGroups {
    param (
        [switch]$Verbose
    )

    $securityGroups = Invoke-GraphRequestWithPaging -Uri "beta/groups?`$filter=isassignabletorole eq true" -Method Get
    if ($Verbose) {
        Write-Verbose "Found $($securityGroups.Count) security groups that are assignable to roles"
    } Else {
        Write-Host "INFO: Found $($securityGroups.Count) security groups that are assignable to roles"
    }

    # Collect members for each security group for later reference
    $securityGroupMembers = @{}
    foreach ($group in $securityGroups) {
        if ($Verbose) {
            Write-Verbose "Collecting members for security group: $($group.displayName)"
        } Else {
            Write-Host "INFO: Collecting members for security group: $($group.displayName)"
        }

        try {
            $members = Invoke-GraphRequestWithPaging -Uri "beta/groups/$($group.id)/transitiveMembers?`$select=id,displayName,userPrincipalName" -Method Get

            # Create member list with useful information
            $memberList = @()
            foreach ($member in $members) {
                # Handle different member types
                if ($member.'@odata.type' -eq '#microsoft.graph.user') {
                    $memberList += [PSCustomObject]@{
                        Type              = "User"
                        Id                = $member.id
                        DisplayName       = $member.displayName
                        UserPrincipalName = $member.userPrincipalName
                    }
                } Elseif ($member.'@odata.type' -eq '#microsoft.graph.group') {
                    $memberList += [PSCustomObject]@{
                        Type        = "Group"
                        Id          = $member.id
                        DisplayName = $member.displayName
                    }
                } Else {
                    $memberList += [PSCustomObject]@{
                        Type        = $member.'@odata.type'.Replace('#microsoft.graph.', '')
                        Id          = $member.id
                        DisplayName = $member.displayName
                    }
                }
            }

            # Store the members in the hashtable
            $securityGroupMembers[$group.id] = @{
                GroupDisplayName = $group.displayName
                GroupId          = $group.id
                Members          = $memberList
                MemberCount      = $memberList.Count
            }

        } catch {
            Write-Error "ERROR: Collecting members for group $($group.displayName): $_"
            continue
        }
    }

    # Return the security group members
    return $securityGroupMembers
}

Function Get-PIMAuditLogs {
    # Get PIM audit logs
    $PIMAudits = Invoke-GraphRequestWithPaging "beta/auditLogs/directoryAudits?`$filter=loggedByService eq 'PIM'"

    $results = @()
    foreach ($PIMaudit in $PIMAudits) {
        # Extract user who initiated the action
        $initiatedByUser = $null
        if ($PIMaudit.InitiatedBy.user) {
            if ($PIMaudit.InitiatedBy.user.userPrincipalName) {
                $initiatedByUser = $PIMaudit.InitiatedBy.user.userPrincipalName
            } elseif ($PIMaudit.InitiatedBy.user.displayName) {
                $initiatedByUser = $PIMaudit.InitiatedBy.user.displayName
            } else {
                $initiatedByUser = "Unknown"
            }
        }
        # If not a user, check if it was initiated by an app or service principal
        if (-not $initiatedByUser) {
            if ($PIMaudit.InitiatedBy.app) {
                $initiatedByUser = $PIMaudit.InitiatedBy.app.displayName
            } elseif ($PIMaudit.InitiatedBy.servicePrincipal) {
                $initiatedByUser = $PIMaudit.InitiatedBy.servicePrincipal.displayName
            } else {
                $initiatedByUser = "Unknown"
            }
        }

        # Determine initiator type
        $initiatorType = "Unknown"
        if ($PIMaudit.InitiatedBy.user) {
            $initiatorType = "User"
        } elseif ($PIMaudit.InitiatedBy.app) {
            $initiatorType = "Application"
        } elseif ($PIMaudit.InitiatedBy.servicePrincipal) {
            $initiatorType = "Service Principal"
        }

        # Get role information
        $roleResource = $PIMaudit.TargetResources | Where-Object { $_.Type -eq "Role" }
        $roleName = $roleResource.DisplayName
        $roleId = $roleResource.id

        # Extract modified properties information
        $roleProperties = @()
        if ($roleResource.modifiedProperties) {
            foreach ($prop in $roleResource.modifiedProperties) {
                # Clean up the values
                $oldValue = $prop.oldValue -replace "^'|'$", ""
                $newValue = $prop.newValue -replace "^'|'$", ""

                # Make property name more readable
                $propName = $prop.displayName
                # Handle common PIM property names
                switch -Wildcard ($propName) {
                    "*ExpirationTime*" { $propName = "Expiration" }
                    "*ActivationTime*" { $propName = "Activation" }
                    "*StartTime*" { $propName = "Start Time" }
                    "*Justification*" { $propName = "Reason" }
                    "*MemberType*" { $propName = "Member Type" }
                    "*AssignmentState*" { $propName = "Assignment" }
                }

                # Format datetime values as before...

                # For empty values
                if ([string]::IsNullOrWhiteSpace($oldValue)) { $oldValue = "(none)" }
                if ([string]::IsNullOrWhiteSpace($newValue)) { $newValue = "(none)" }

                # Create property change format - using HTML entity for arrow instead of Unicode
                if ($oldValue -eq "(none)" -and $newValue -ne "(none)") {
                    $roleProperties += "$($propName): $newValue"
                } elseif ($oldValue -ne "(none)" -and $newValue -eq "(none)") {
                    $roleProperties += "$($propName): Removed"
                } elseif ($oldValue -eq $newValue) {
                    $roleProperties += "$($propName): $newValue"
                } else {
                    $roleProperties += "$($propName): $oldValue → $newValue"
                }
            }
        }
        $rolePropertiesText = $roleProperties -join " | "

        # Get request information
        $requestResource = $PIMaudit.TargetResources | Where-Object { $_.type -eq "Request" }
        $requestId = $requestResource.id

        # Extract target user details with enhanced group context
        $userDetails = "N/A"
        $targetUserId = $null

        # Check target resources for user
        $userResource = $PIMaudit.TargetResources | Where-Object { $_.type -eq "User" }
        if ($userResource -and $userResource.userPrincipalName) {
            $userDetails = $userResource.userPrincipalName
            $targetUserId = $userResource.id
        }

        # Check if this is a group-based activation
        $isGroupBasedActivation = $false
        $groupInfo = $null

        # Look for group information in additional details
        $groupDetail = $PIMaudit.AdditionalDetails | Where-Object { $_.key -eq "GroupId" -or $_.key -eq "Group" -or $_.key -eq "MemberType" }
        if ($groupDetail -and $groupDetail.value -eq "Group") {
            $isGroupBasedActivation = $true
        }

        # Check if the operation indicates group-based activation
        if ($PIMaudit.ActivityDisplayName -like "*group*" -or $rolePropertiesText -like "*Group*") {
            $isGroupBasedActivation = $true
        }

        # Get directory information
        $directoryResource = $PIMaudit.TargetResources | Where-Object { $_.type -eq "Directory" }
        $directoryName = $directoryResource.displayName

        # Get reason for the action
        $reason = $PIMaudit.ResultReason
        if ([string]::IsNullOrWhiteSpace($reason)) {
            $reason = "N/A"
        }

        # Extract start time and duration for better activation tracking
        $startTime = "N/A"
        $duration = "N/A"
        $startTimeDetail = $PIMaudit.AdditionalDetails | Where-Object { $_.key -eq "StartTime" }
        if ($startTimeDetail) {
            try {
                $parsedStartTime = [DateTime]$startTimeDetail.value
                $startTime = $parsedStartTime.ToString('dd/MM/yyyy hh:mm:ss tt')
            } catch {
                $startTime = $startTimeDetail.value
            }
        }

        $durationDetail = $PIMaudit.AdditionalDetails | Where-Object { $_.key -eq "Duration" }
        if ($durationDetail) {
            $duration = $durationDetail.value
        }

        # Create custom object with enhanced information for group activation tracking
        $results += [PSCustomObject]@{
            "DateTime"               = $PIMaudit.ActivityDateTime
            "InitiatedBy"            = $initiatedByUser
            "OperationType"          = $PIMaudit.OperationType
            "InitiatedByType"        = $initiatorType
            "Role"                   = $roleName
            "RoleID"                 = $roleId
            "RoleProperties"         = $rolePropertiesText
            "Target"                 = $userDetails
            "TargetUserId"           = $targetUserId
            "IsGroupBasedActivation" = $isGroupBasedActivation
            "GroupInfo"              = $groupInfo
            "Directory"              = $directoryName
            "RequestID"              = $requestId
            "Operation"              = $PIMaudit.ActivityDisplayName
            "Result"                 = $PIMaudit.Result
            "Justification"          = $reason
            "StartTime"              = $startTime
            "Duration"               = $duration
        }
    }

    # Sort the results by DateTime in descending order
    $results = $results | Sort-Object -Property DateTime -Descending
    # Return the results
    return $results
}


function Get-GroupActivationDetails {
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$ActivatedMembers,
        [Parameter(Mandatory = $true)]
        [array]$PIMAuditLogs,
        [Parameter(Mandatory = $true)]
        [string]$RoleName
    )

    $enrichedActivations = @()

    # Return empty array if no activated members (PowerShell 5 compatible)
    if (-not $ActivatedMembers -or @($ActivatedMembers).Count -eq 0) {
        return $enrichedActivations
    }

    foreach ($member in $ActivatedMembers) {
        # Find corresponding PIM audit logs for this member and role
        $memberAuditLogs = $PIMAuditLogs | Where-Object {
            ($_.Target -eq $member.UserPrincipalName -or $_.TargetUserId -eq $member.UserId) -and
            $_.Role -eq $RoleName -and
            $_.IsGroupBasedActivation -eq $true -and
            $_.Result -eq "Success" -and
            ($_.Operation -like "*Activate*" -or $_.OperationType -eq "Assign")
        }

        # Get the most recent activation for this member
        $recentActivation = $memberAuditLogs | Sort-Object DateTime -Descending | Select-Object -First 1

        if ($recentActivation) {
            $enrichedActivations += [PSCustomObject]@{
                UserPrincipalName = $member.UserPrincipalName
                DisplayName       = $member.DisplayName
                UserId            = $member.UserId
                ActivationTime    = $member.ActivationTime
                StartTime         = $member.StartTime
                EndTime           = $member.EndTime
                AssignmentState   = $member.AssignmentState
                MemberType        = $member.MemberType
                # Enhanced with audit log data
                AuditLogDateTime  = $recentActivation.DateTime
                ActivatedBy       = $recentActivation.InitiatedBy
                Justification     = $recentActivation.Justification
                Duration          = $recentActivation.Duration
                RequestID         = $recentActivation.RequestID
            }
        } else {
            # Include member even without audit log correlation
            $enrichedActivations += [PSCustomObject]@{
                UserPrincipalName = $member.UserPrincipalName
                DisplayName       = $member.DisplayName
                UserId            = $member.UserId
                ActivationTime    = $member.ActivationTime
                StartTime         = $member.StartTime
                EndTime           = $member.EndTime
                AssignmentState   = $member.AssignmentState
                MemberType        = $member.MemberType
                # No audit log data found
                AuditLogDateTime  = "N/A"
                ActivatedBy       = "N/A"
                Justification     = "N/A"
                Duration          = "N/A"
                RequestID         = "N/A"
            }
        }
    }

    return $enrichedActivations
}

function Invoke-EntraAdminRolesReportCore {
    param(
        [Parameter(Mandatory=$false)] [switch] $SendEmail,
        [Parameter(Mandatory=$false)] [string[]] $Recipient,
        [Parameter(Mandatory=$false)] [string] $From,
        [Parameter(Mandatory=$false)] [string] $ExportPath,
        [Parameter(Mandatory=$false)] [switch] $DebugMode
    )
    $tenantInfo = Invoke-MgGraphRequest -Uri 'beta/organization' -Method Get -OutputType PSObject
    $tenantname = $tenantInfo.value[0].displayName

        # Call the function to get the security group members
        $securityGroupMembers = Get-SecurityGroups

        # Add debug output for group structure
        if ($DebugMode) { Write-Host "debug: Security group members structure:" -ForegroundColor Yellow }
        foreach ($groupId in $securityGroupMembers.Keys) {
            $group = $securityGroupMembers[$groupId]
            if ($DebugMode) { Write-Host "debug: Group $groupId ($($group.GroupDisplayName)) has $($group.MemberCount) members" -ForegroundColor Yellow }
            if ($group.Members) {
                foreach ($member in $group.Members) {
                    if ($DebugMode) { Write-Host "debug:   Member: $($member.UserPrincipalName) (ID: $($member.Id))" -ForegroundColor Cyan }
                }
            }
        }

        $adminUnits = Invoke-GraphRequestWithPaging -Uri "beta/directory/administrativeUnits" -Method Get
        $auLookup = @{}
        foreach ($au in $adminUnits) {
            # The directoryScopeId format is: "/administrativeUnits/{id}"
            $auId = "/administrativeUnits/$($au.id)"
            $auLookup[$auId] = $au.displayName
        }

        if ($adminUnits) {
            Write-Host "INFO: Found administrative units in the tenant."
        } else {
            Write-Host "INFO: No administrative units found in the tenant."
        }

        # Get role assignments with principal expansion
        $rolesWithPrincipal = Invoke-GraphRequestWithPaging -Uri "beta/roleManagement/directory/roleAssignments?`$expand=principal" -Method Get
        # Get role assignments with roleDefinition expansion
        $rolesWithDefinition = Invoke-GraphRequestWithPaging -Uri "beta/roleManagement/directory/roleAssignments?`$expand=roleDefinition" -Method Get

        # Merge the data for complete role assignment information
        $defLookup = @{}
        foreach ($d in $rolesWithDefinition) { $defLookup[$d.id] = $d.roleDefinition }
        $roles = [System.Collections.Generic.List[object]]::new($rolesWithPrincipal.Count)
        foreach ($role in $rolesWithPrincipal) {
            $roleDefinition = $defLookup[$role.id]
            $role | Add-Member -MemberType NoteProperty -Name roleDefinition1 -Value $roleDefinition -Force
            $roles.Add($role)
        }

        Write-Host "INFO: Found $($rolesWithPrincipal.Count) role assignments in the tenant."

        try {
            Write-Host "INFO: Collecting PIM eligible role assignments..." -ForegroundColor Cyan
            $eligibleRoles = Invoke-GraphRequestWithPaging -Uri "beta/roleManagement/directory/roleEligibilitySchedules?`$expand=roleDefinition,principal" -Method Get

            if ($DebugMode) { Write-Host "debug: Eligible roles API call completed" -ForegroundColor Magenta }
            if ($DebugMode) { Write-Host "debug: eligibleRoles type: $($eligibleRoles.GetType().Name)" -ForegroundColor Magenta }
            if ($DebugMode) { Write-Host "debug: eligibleRoles is null: $($null -eq $eligibleRoles)" -ForegroundColor Magenta }
            if ($DebugMode) { Write-Host "debug: eligibleRoles count: $($eligibleRoles.Count)" -ForegroundColor Magenta }

            if ($null -eq $eligibleRoles) {
                Write-Warning "Unable to collect PIM eligible role assignments. This MAY be due to missing Microsoft Entra ID Premium P2 license."
                Write-Host "INFO: Continuing without PIM eligible role assignments..." -ForegroundColor Yellow
                $eligibleRoles = @() # Set to empty array so the code can continue
            } else {
                Write-Host "INFO: Found $($eligibleRoles.Count) eligible role assignments." -ForegroundColor Green
            }
        } catch {
            Write-Warning "Unable to collect PIM eligible role assignments. This may be due to missing Microsoft Entra ID Premium P2 license."
            Write-Host "INFO: Continuing without PIM eligible role assignments..." -ForegroundColor Yellow
            if ($DebugMode) { Write-Host "debug: Exception details: $($_.Exception.Message)" -ForegroundColor Red }
            $eligibleRoles = @() # Set to empty array so the code can continue
        }

        foreach ($eligibleRole in $eligibleRoles) {
            $eligibleRole | Add-Member -MemberType NoteProperty -Name roleDefinition1 -Value $eligibleRole.roleDefinition -Force
            $roles += $eligibleRole
        }

        # Get activated PIM role assignments
        $roleActivations = @()
        # Always try to collect activated assignments, regardless of eligible roles
        # because users might have activations even without current eligible assignments
        try {
            Write-Host "INFO: Collecting activated PIM role assignments..." -ForegroundColor Cyan
            if ($DebugMode) { Write-Host "debug: Using API endpoint: beta/roleManagement/directory/roleAssignmentScheduleInstances?`$filter=assignmentType eq 'Activated'" -ForegroundColor Magenta }

            $roleActivations = Invoke-GraphRequestWithPaging -Uri "beta/roleManagement/directory/roleAssignmentScheduleInstances?`$filter=assignmentType eq 'Activated'" -Method Get

            if ($DebugMode) { Write-Host "debug: API call completed. Checking results..." -ForegroundColor Magenta }
            if ($DebugMode) { Write-Host "debug: roleActivations type: $($roleActivations.GetType().Name)" -ForegroundColor Magenta }
            if ($DebugMode) { Write-Host "debug: roleActivations is null: $($null -eq $roleActivations)" -ForegroundColor Magenta }
            if ($DebugMode) { Write-Host "debug: roleActivations count: $($roleActivations.Count)" -ForegroundColor Magenta }

            if ($roleActivations -and $roleActivations.Count -gt 0) {
                Write-Host "INFO: Found $($roleActivations.Count) activated PIM role assignments." -ForegroundColor Green
                # Add debug info about the activations
                if ($DebugMode) { Write-Host "debug: Detailed activation analysis:" -ForegroundColor Magenta }
                foreach ($activation in ($roleActivations | Select-Object -First 3)) {
                    if ($DebugMode) { Write-Host "debug ACTIVATION SAMPLE:" -ForegroundColor Cyan }
                    if ($DebugMode) { Write-Host "  Principal ID: $($activation.principalId)" -ForegroundColor Cyan }
                    if ($DebugMode) { Write-Host "  Role Definition ID: $($activation.roleDefinitionId)" -ForegroundColor Cyan }
                    if ($DebugMode) { Write-Host "  Directory Scope ID: $($activation.directoryScopeId)" -ForegroundColor Cyan }
                    if ($DebugMode) { Write-Host "  Assignment Type: $($activation.assignmentType)" -ForegroundColor Cyan }
                    if ($DebugMode) { Write-Host "  Member Type: $($activation.memberType)" -ForegroundColor Cyan }
                    if ($DebugMode) { Write-Host "  Start Time: $($activation.startDateTime)" -ForegroundColor Cyan }
                    if ($DebugMode) { Write-Host "  End Time: $($activation.endDateTime)" -ForegroundColor Cyan }
                    if ($DebugMode) { Write-Host "  Created DateTime: $($activation.createdDateTime)" -ForegroundColor Cyan }
                    if ($DebugMode) { Write-Host "  ---" -ForegroundColor Cyan }
                }
            } else {
                if ($DebugMode) { Write-Host "INFO: No activated PIM role assignments found." -ForegroundColor Yellow }
                if ($DebugMode) { Write-Host "debug: This could mean:" -ForegroundColor Yellow }
                if ($DebugMode) { Write-Host "  1. No users have currently activated PIM roles" -ForegroundColor Yellow }
                if ($DebugMode) { Write-Host "  2. Permission issues preventing access to activation data" -ForegroundColor Yellow }
                if ($DebugMode) { Write-Host "  3. API endpoint or filter syntax issues" -ForegroundColor Yellow }

                # Try alternative API call without filter to see if we get any data
                if ($DebugMode) { Write-Host "debug: Trying alternative API call without filter..." -ForegroundColor Magenta }
                try {
                    $allInstances = Invoke-GraphRequestWithPaging -Uri "beta/roleManagement/directory/roleAssignmentScheduleInstances" -Method Get | Select-Object -First 10
                    if ($DebugMode) { Write-Host "debug: Alternative call returned $($allInstances.Count) items (showing first 10)" -ForegroundColor Magenta }

                    if ($allInstances -and $allInstances.Count -gt 0) {
                        if ($DebugMode) { Write-Host "debug: Sample of all roleAssignmentScheduleInstances:" -ForegroundColor Magenta }
                        if ($DebugMode) {
                            foreach ($instance in $allInstances) {
                                Write-Host "  Instance - Assignment Type: $($instance.assignmentType), Principal: $($instance.principalId), Role: $($instance.roleDefinitionId)" -ForegroundColor Cyan
                            }
                        }

                        # Check if any of these are activated
                        $activatedFromAll = $allInstances | Where-Object { $_.assignmentType -eq "Activated" }
                        if ($DebugMode) { Write-Host "debug: Found $($activatedFromAll.Count) activated assignments in unfiltered results" -ForegroundColor Magenta }

                        if ($activatedFromAll.Count -gt 0) {
                            if ($DebugMode) { Write-Host "debug: Using activated assignments from unfiltered results" -ForegroundColor Green }
                            $roleActivations = $activatedFromAll
                        }
                    }

                    # Also try another endpoint that might contain activation data
                    if ($DebugMode) { Write-Host "debug: Trying roleAssignmentSchedules endpoint..." -ForegroundColor Magenta }
                    $schedules = Invoke-GraphRequestWithPaging -Uri "beta/roleManagement/directory/roleAssignmentSchedules" -Method Get | Select-Object -First 5
                    if ($DebugMode) { Write-Host "debug: roleAssignmentSchedules returned $($schedules.Count) items" -ForegroundColor Magenta }
                    if ($DebugMode) {
                        foreach ($schedule in $schedules) {
                            Write-Host "  Schedule - Assignment Type: $($schedule.assignmentType), Principal: $($schedule.principalId), Status: $($schedule.status)" -ForegroundColor Cyan
                        }
                    }

                } catch {
                    if ($DebugMode) { Write-Host "debug: Alternative API calls also failed: $_" -ForegroundColor Red }
                }

                $roleActivations = @()
            }
        } catch {
            Write-Warning "Unable to collect activated PIM role assignments: $_"
            if ($DebugMode) { Write-Host "debug: Exception details: $($_.Exception.Message)" -ForegroundColor Red }
            if ($DebugMode) { Write-Host "debug: Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red }
            $roleActivations = @()
        }

        if (!$roles) {
            if ($verbose) {
                Write-Verbose "No role assignments found, exiting..."
            } else {
                Write-Host "INFO: No role assignments found" -ForegroundColor Red
            }
            return
        }

        $Report = @()

        # Create a hashtable to track role assignments and detect duplicates
        $roleAssignmentTracker = @{}

        # Process roles in specific order: permanent assignments first, then eligible assignments
        # This ensures eligible assignments can override permanent ones
        $permanentRoles = $roles | Where-Object { -not $_.status }
        $eligibleRolesFromCollection = $roles | Where-Object { $_.status }

        Write-Host "INFO: Processing $($permanentRoles.Count) permanent role assignments..." -ForegroundColor Cyan
        Write-Host "INFO: Processing $($eligibleRolesFromCollection.Count) eligible role assignments..." -ForegroundColor Cyan

        # Combine in the right order: permanent first, then eligible
        $orderedRoles = $permanentRoles + $eligibleRolesFromCollection

        foreach ($role in $orderedRoles) {
            # Decide the principal type based on the '@odata.type' property
            Switch ($role.principal.'@odata.type') {
                '#microsoft.graph.user' {
                    $principalType = "User"
                    $Principal = $role.principal.userPrincipalName

                    if ($null -eq $role.principal.accountEnabled) {
                        $AccountStatus = "N/A"
                    } elseif ($role.principal.accountEnabled -eq $true) {
                        $AccountStatus = "Enabled"
                    } else {
                        $AccountStatus = "Disabled"
                    }
                }
                '#microsoft.graph.group' {
                    $principalType = "Group"
                    $Principal = $role.principal.id
                    if ($null -eq $role.principal.accountEnabled) {
                        $AccountStatus = "N/A"
                    } elseif ($role.principal.accountEnabled -eq $true) {
                        $AccountStatus = "Enabled"
                    } else {
                        $AccountStatus = "Disabled"
                    }
                }
                '#microsoft.graph.servicePrincipal' {
                    $principalType = "Service Principal"
                    $Principal = $role.principal.id

                    if ($null -eq $role.principal.accountEnabled) {
                        $AccountStatus = "N/A"
                    } elseif ($role.principal.accountEnabled -eq $true) {
                        $AccountStatus = "Enabled"
                    } else {
                        $AccountStatus = "Disabled"
                    }
                }
            }

            # Check Assigned Role FIRST (we need this for debug messages)
            if ($role.roleDefinition1.displayName) {
                $assignedRole = $role.roleDefinition1.displayName
            } elseif ($role.roleDefinition.displayName) {
                $assignedRole = $role.roleDefinition.displayName
            } else {
                $assignedRole = "Unknown"
            }

            # Decide the role assignment type based on the role eligibility schedule
            if ($role.status) {
                $status = "Eligible"

                # Initialize activated members array (only for groups with actual activations)
                $activatedMembers = @()

                # Check for activated assignments - only for individual users, not groups
                if ($roleActivations.Count -gt 0 -and $principalType -eq "User") {
                    # Look for user activations in the roleActivations data
                    $userActivation = $roleActivations | Where-Object {
                        $_.roleDefinitionId -eq $role.roleDefinitionId -and
                        $_.principalId -eq $role.principalId -and
                        $_.directoryScopeId -eq $role.directoryScopeId
                    }

                    if ($userActivation) {
                        $status = "Eligible (Active)"
                    }
                }
                # For groups, collect which members have activated this role using roleActivations data
                elseif ($roleActivations.Count -gt 0 -and $principalType -eq "Group") {
                    # First check if there are any activations for this specific role+scope combination
                    $roleSpecificActivationsPreCheck = $roleActivations | Where-Object {
                        $_.roleDefinitionId -eq $role.roleDefinitionId -and
                        $_.directoryScopeId -eq $role.directoryScopeId
                    }

                    # Only proceed if there are potential activations for this role
                    if ($roleSpecificActivationsPreCheck.Count -gt 0) {
                        if ($DebugMode) { Write-Host "debug: Processing group $($role.principal.displayName) for role $assignedRole" -ForegroundColor Yellow }
                        if ($DebugMode) { Write-Host "debug: Group ID: $($role.principal.id)" -ForegroundColor Yellow }
                        if ($DebugMode) { Write-Host "debug: Role Definition ID: $($role.roleDefinitionId)" -ForegroundColor Yellow }
                        if ($DebugMode) { Write-Host "debug: Directory Scope ID: $($role.directoryScopeId)" -ForegroundColor Yellow }

                        # Get members of this specific group
                        $groupMembers = $securityGroupMembers[$role.principal.id]
                        if ($DebugMode) { Write-Host "debug: Group members found: $($null -ne $groupMembers)" -ForegroundColor Yellow }

                        if ($groupMembers -and $groupMembers.Members) {
                            if ($DebugMode) { Write-Host "debug: Group has $($groupMembers.Members.Count) members" -ForegroundColor Yellow }

                            # Show first few group members for debugging
                            if ($DebugMode) {
                                foreach ($member in ($groupMembers.Members | Select-Object -First 3)) {
                                    Write-Host "debug: Group member: $($member.UserPrincipalName) (ID: $($member.Id))" -ForegroundColor Cyan
                                }
                            }

                            # Look for activations for this specific role definition
                            $roleSpecificActivations = $roleActivations | Where-Object {
                                $_.roleDefinitionId -eq $role.roleDefinitionId -and
                                $_.directoryScopeId -eq $role.directoryScopeId
                                # Remove the memberType filter to catch all activations
                            }

                            if ($DebugMode) { Write-Host "debug: Found $($roleSpecificActivations.Count) role activations for role $assignedRole" -ForegroundColor Yellow }

                            # If no activations found, try broader search
                            if ($roleSpecificActivations.Count -eq 0) {
                                if ($DebugMode) { Write-Host "debug: No activations found with exact match, trying broader search..." -ForegroundColor Yellow }

                                # Try without directory scope restriction
                                $broadActivations = $roleActivations | Where-Object {
                                    $_.roleDefinitionId -eq $role.roleDefinitionId
                                }
                                if ($DebugMode) { Write-Host "debug: Found $($broadActivations.Count) activations for this role (any scope)" -ForegroundColor Yellow }

                                # Try looking for any activations involving group members
                                $memberActivations = $roleActivations | Where-Object {
                                    $activation = $_
                                    $groupMembers.Members | Where-Object { $_.Id -eq $activation.principalId }
                                }
                                if ($DebugMode) { Write-Host "debug: Found $($memberActivations.Count) activations involving group members (any role)" -ForegroundColor Yellow }
                            }

                            foreach ($activation in $roleSpecificActivations) {
                                if ($DebugMode) { Write-Host "debug: Checking activation for principal $($activation.principalId)" -ForegroundColor Cyan }
                                if ($DebugMode) { Write-Host "debug: Activation start time: $($activation.startDateTime) (type: $($activation.startDateTime.GetType().Name))" -ForegroundColor Cyan }
                                if ($DebugMode) { Write-Host "debug: Activation end time: $($activation.endDateTime) (type: $($activation.endDateTime.GetType().Name))" -ForegroundColor Cyan }

                                # Check if the activation's principal is a member of this group
                                $targetUser = $groupMembers.Members | Where-Object {
                                    $_.Id -eq $activation.principalId
                                }

                                if ($targetUser) {
                                    if ($DebugMode) { Write-Host "debug: Found matching user: $($targetUser.UserPrincipalName)" -ForegroundColor Green }

                                    # Check if we already have this activation recorded
                                    $existingActivation = $activatedMembers | Where-Object {
                                        $_.UserId -eq $targetUser.Id
                                    }

                                    if (-not $existingActivation) {
                                        # Calculate end time and duration with robust datetime parsing
                                        $startTime = $null
                                        $endTime = $null

                                        if ($activation.startDateTime) {
                                            try {
                                                # Try different parsing approaches
                                                if ($activation.startDateTime -is [DateTime]) {
                                                    $startTime = $activation.startDateTime
                                                } elseif ($activation.startDateTime -is [string]) {
                                                    # Handle different datetime formats
                                                    if ($activation.startDateTime -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}') {
                                                        $startTime = [DateTime]::Parse($activation.startDateTime)
                                                    } else {
                                                        $startTime = Get-Date $activation.startDateTime
                                                    }
                                                }
                                            } catch {
                                                if ($DebugMode) { Write-Host "debug: Failed to parse start time: $($activation.startDateTime)" -ForegroundColor Red }
                                                $startTime = $null
                                            }
                                        }

                                        if ($activation.endDateTime) {
                                            try {
                                                # Try different parsing approaches
                                                if ($activation.endDateTime -is [DateTime]) {
                                                    $endTime = $activation.endDateTime
                                                } elseif ($activation.endDateTime -is [string]) {
                                                    # Handle different datetime formats
                                                    if ($activation.endDateTime -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}') {
                                                        $endTime = [DateTime]::Parse($activation.endDateTime)
                                                    } else {
                                                        $endTime = Get-Date $activation.endDateTime
                                                    }
                                                }
                                            } catch {
                                                if ($DebugMode) { Write-Host "debug: Failed to parse end time: $($activation.endDateTime)" -ForegroundColor Red }
                                                $endTime = $null
                                            }
                                        }

                                        $duration = if ($startTime -and $endTime) {
                                            $timeSpan = $endTime - $startTime
                                            if ($timeSpan.TotalDays -ge 1) {
                                                "$([math]::Floor($timeSpan.TotalDays)) days, $($timeSpan.Hours) hours"
                                            } else {
                                                "$($timeSpan.Hours) hours, $($timeSpan.Minutes) minutes"
                                            }
                                        } else { "N/A" }

                                        $activatedMembers += [PSCustomObject]@{
                                            UserPrincipalName = $targetUser.UserPrincipalName
                                            DisplayName       = $targetUser.DisplayName
                                            UserId            = $targetUser.Id
                                            ActivationTime    = if ($startTime) { $startTime.ToString('dd-MM-yyyy HH:mm') } else { "N/A" }
                                            StartTime         = if ($startTime) { $startTime.ToString('dd-MM-yyyy HH:mm') } else { "N/A" }
                                            EndTime           = if ($endTime) { $endTime.ToString('dd-MM-yyyy HH:mm') } else { "N/A" }
                                            Duration          = $duration
                                            AssignmentState   = "Active"
                                            MemberType        = "Direct"
                                            # Try to get additional details from the activation if available
                                            ActivatedBy       = if ($activation.createdBy) { $activation.createdBy.user.displayName } else { "N/A" }
                                            Justification     = if ($activation.justification) { $activation.justification } else { "N/A" }
                                        }
                                        if ($DebugMode) { Write-Host "debug: Added activated member $($targetUser.UserPrincipalName) for role $assignedRole" -ForegroundColor Green }
                                    } else {
                                        if ($DebugMode) { Write-Host "debug: Activation already recorded for $($targetUser.UserPrincipalName)" -ForegroundColor Yellow }
                                    }
                                } else {
                                    if ($DebugMode) { Write-Host "debug: Principal $($activation.principalId) not found in group members" -ForegroundColor Red }
                                }
                            }
                            if ($DebugMode) { Write-Host "debug: Final activated members count: $(@($activatedMembers).Count)" -ForegroundColor Magenta }

                            # Update status if we found activated members
                            if (@($activatedMembers).Count -gt 0) {
                                $status = "Eligible (Active)"
                                if ($DebugMode) { Write-Host "debug: Group marked as Eligible (Active) due to activated members" -ForegroundColor Green }
                            }

                            # If no activated members found through roleActivations, try using PIM audit logs as fallback
                            if (@($activatedMembers).Count -eq 0) {
                                if ($DebugMode) { Write-Host "debug: No activated members found via roleActivations, trying enhanced PIM audit logs search..." -ForegroundColor Yellow }

                                # Get PIM audit logs if not already collected
                                if (-not $PIMAuditLogsReport) {
                                    if ($DebugMode) { Write-Host "debug: Collecting PIM audit logs for activation detection..." -ForegroundColor Yellow }
                                    $PIMAuditLogsReport = Get-PIMAuditLogs
                                }

                                if ($PIMAuditLogsReport) {
                                    if ($DebugMode) { Write-Host "debug: PIM audit logs available: $($PIMAuditLogsReport.Count) entries" -ForegroundColor Yellow }

                                    # Look for recent activations in PIM audit logs for this role with enhanced search
                                    $roleActivationLogs = $PIMAuditLogsReport | Where-Object {
                            ($_.Operation -like "*Activate*" -or $_.Operation -like "*activate*" -or $_.OperationType -eq "Assign") -and
                            ($_.Role -eq $assignedRole -or $_.RoleID -eq $role.roleDefinitionId) -and
                                        $_.Result -eq "Success" -and
                                        $_.DateTime -gt (Get-Date).AddDays(-7)  # Look at last 7 days for more recent data
                                    }

                                    if ($DebugMode) { Write-Host "debug: Found $($roleActivationLogs.Count) activation logs for role $assignedRole in last 7 days" -ForegroundColor Yellow }

                                    # If still no results, try broader search criteria
                                    if ($roleActivationLogs.Count -eq 0) {
                                        if ($DebugMode) { Write-Host "debug: No activations in last 7 days, trying last 30 days..." -ForegroundColor Yellow }
                                        $roleActivationLogs = $PIMAuditLogsReport | Where-Object {
                                ($_.Operation -like "*Activate*" -or $_.Operation -like "*activate*" -or $_.OperationType -eq "Assign") -and
                                ($_.Role -eq $assignedRole -or $_.RoleID -eq $role.roleDefinitionId) -and
                                            $_.Result -eq "Success" -and
                                            $_.DateTime -gt (Get-Date).AddDays(-30)  # Look at last 30 days
                                        }
                                        if ($DebugMode) { Write-Host "debug: Found $($roleActivationLogs.Count) activation logs for role $assignedRole in last 30 days" -ForegroundColor Yellow }
                                    }

                                    # If still no results, try even broader search without role name restriction
                                    if ($roleActivationLogs.Count -eq 0) {
                                        if ($DebugMode) { Write-Host "debug: Trying even broader search for any activations by group members..." -ForegroundColor Yellow }
                                        $memberActivationLogs = $PIMAuditLogsReport | Where-Object {
                                ($_.Operation -like "*Activate*" -or $_.Operation -like "*activate*" -or $_.OperationType -eq "Assign") -and
                                            $_.Result -eq "Success" -and
                                            $_.DateTime -gt (Get-Date).AddDays(-7) -and
                                ($groupMembers.Members | Where-Object { $_.UserPrincipalName -eq $_.Target -or $_.Id -eq $_.TargetUserId })
                                        }
                                        if ($DebugMode) { Write-Host "debug: Found $($memberActivationLogs.Count) activation logs for any group members in last 7 days" -ForegroundColor Yellow }

                                        # Filter these to matching roles if possible
                                        $roleActivationLogs = $memberActivationLogs | Where-Object {
                                            $_.Role -eq $assignedRole -or $_.RoleID -eq $role.roleDefinitionId
                                        }
                                        if ($DebugMode) { Write-Host "debug: Of those, $($roleActivationLogs.Count) match the current role" -ForegroundColor Yellow }
                                    }

                                    foreach ($activationLog in $roleActivationLogs) {
                                        if ($DebugMode) { Write-Host "debug: Processing audit log activation: $($activationLog.Target) for role $($activationLog.Role)" -ForegroundColor Cyan }

                                        # Check if the target user is a member of this group
                                        $targetUser = $groupMembers.Members | Where-Object {
                                            $_.UserPrincipalName -eq $activationLog.Target -or
                                            $_.Id -eq $activationLog.TargetUserId
                                        }

                                        if ($targetUser) {
                                            if ($DebugMode) { Write-Host "debug: Found group member match: $($targetUser.UserPrincipalName)" -ForegroundColor Green }

                                            # Check if we already have this activation recorded
                                            $existingActivation = $activatedMembers | Where-Object {
                                                $_.UserPrincipalName -eq $targetUser.UserPrincipalName
                                            }

                                            if (-not $existingActivation) {
                                                $activatedMembers += [PSCustomObject]@{
                                                    UserPrincipalName = $targetUser.UserPrincipalName
                                                    DisplayName       = $targetUser.DisplayName
                                                    UserId            = $targetUser.Id
                                                    ActivationTime    = if ($activationLog.DateTime) { $activationLog.DateTime.ToString('dd-MM-yyyy HH:mm') } else { "N/A" }
                                                    StartTime         = if ($activationLog.StartTime -and $activationLog.StartTime -ne "N/A") { $activationLog.StartTime } else { "N/A" }
                                                    EndTime           = "N/A"
                                                    Duration          = if ($activationLog.Duration -and $activationLog.Duration -ne "N/A") { $activationLog.Duration } else { "N/A" }
                                                    AssignmentState   = "Active"
                                                    MemberType        = "Group"
                                                    ActivatedBy       = if ($activationLog.InitiatedBy) { $activationLog.InitiatedBy } else { "N/A" }
                                                    Justification     = if ($activationLog.Justification) { $activationLog.Justification } else { "N/A" }
                                                }
                                                if ($DebugMode) { Write-Host "debug: Added activated member from audit logs: $($targetUser.UserPrincipalName)" -ForegroundColor Green }
                                            } else {
                                                if ($DebugMode) { Write-Host "debug: Activation already recorded for $($targetUser.UserPrincipalName)" -ForegroundColor Yellow }
                                            }
                                        } else {
                                            if ($DebugMode) { Write-Host "debug: Target user $($activationLog.Target) not found in group members" -ForegroundColor Red }
                                        }
                                    }

                                    if ($DebugMode) { Write-Host "debug: Final activated members count after enhanced audit log check: $(@($activatedMembers).Count)" -ForegroundColor Magenta }

                                    # Update status if we found activated members through audit logs
                                    if (@($activatedMembers).Count -gt 0) {
                                        $status = "Eligible (Active)"
                                        if ($DebugMode) { Write-Host "debug: Group marked as Eligible (Active) due to activated members found in audit logs" -ForegroundColor Green }
                                    }
                                } else {
                                    if ($DebugMode) { Write-Host "debug: No PIM audit logs available for fallback detection" -ForegroundColor Red }
                                }
                            }
                        } else {
                            if ($DebugMode) { Write-Host "debug: No group members found for group $($role.principal.id)" -ForegroundColor Red }
                        }
                    } else {
                        if ($DebugMode) { Write-Host "debug: No activations found for group $($role.principal.displayName) and role $assignedRole - skipping activation processing" -ForegroundColor Yellow }
                    }
                }

                if ($role.scheduleInfo.startDateTime) {
                    $startDate = ($role.scheduleInfo.startDateTime).ToString("dd-MM-yyyy HH:mm")
                } else {
                    $startDate = "Permanent"
                }

                if ($role.scheduleInfo.expiration.endDateTime) {
                    $endDate = ($role.scheduleInfo.expiration.endDateTime).ToString("dd-MM-yyyy HH:mm")
                } else {
                    $endDate = "Permanent"
                }
            } else {
                $status = "Permanent"
                $StartDate = "Permanent"
                $endDate = "Permanent"
                $activatedMembers = @()
            }

            # Create a unique key for this role assignment to detect duplicates
            $uniqueKey = "$Principal|$assignedRole|$($role.directoryScopeId)"

            $Reportline = [PSCustomObject]@{
                "Principal"             = $Principal
                "DisplayName"           = $role.principal.displayName
                "AccountStatus"         = $AccountStatus
                "PrincipalType"         = $principalType
                "Assigned Role"         = $assignedRole
                "AssignedRoleScopeName" = if ($role.directoryScopeId -eq "/" -or $null -eq $role.directoryScopeId) { "Tenant-Wide" } else { "AU/$($auLookup[$role.directoryScopeId])" }
                "AssignmentType"        = $status
                "AssignmentStartDate"   = $startDate
                "AssignmentEndDate"     = $endDate
                "ActivatedMembers"      = if ($principalType -eq "Group" -and @($activatedMembers).Count -gt 0) {
                    $activatedMembers  # Return activated members for groups that have them
                } else {
                    @()  # Return empty array for groups without activated members or non-groups
                }
                "IsBuiltIn"             = if ($role.roleDefinition.isBuiltIn) { $role.roleDefinition.isBuiltIn } elseif ($role.roleDefinition1.isBuiltIn) { $role.roleDefinition1.isBuiltIn } else { $null }
            }

            # Check if we already have this role assignment
            if ($roleAssignmentTracker.ContainsKey($uniqueKey)) {
                $existingAssignment = $roleAssignmentTracker[$uniqueKey]

                if ($DebugMode) { Write-Host "debug: Found duplicate for $Principal - $assignedRole" -ForegroundColor Yellow }
                if ($DebugMode) { Write-Host "debug: Existing: $($existingAssignment.AssignmentType), Current: $status" -ForegroundColor Yellow }
                if ($DebugMode) { Write-Host "debug: Existing activated members: $($existingAssignment.ActivatedMembers.Count), Current: $($Reportline.ActivatedMembers.Count)" -ForegroundColor Yellow }

                # If the existing assignment is permanent and the current one is eligible, replace it
                # Eligible assignments take priority over permanent ones
                if ($existingAssignment.AssignmentType -eq "Permanent" -and ($status -eq "Eligible" -or $status -eq "Eligible (Active)")) {
                    if ($DebugMode) { Write-Host "REPLACING: permanent assignment with eligible assignment for: $Principal - $assignedRole" -ForegroundColor Green }
                    $roleAssignmentTracker[$uniqueKey] = $Reportline
                }
                # If both are eligible, merge the activated members to preserve all activations
                elseif ($existingAssignment.AssignmentType -eq "Eligible" -and ($status -eq "Eligible" -or $status -eq "Eligible (Active)")) {
                    Write-Host "MERGING: eligible assignments for: $Principal - $assignedRole" -ForegroundColor Green

                    # Merge activated members from both assignments
                    $mergedActivatedMembers = @()
                    $mergedActivatedMembers += $existingAssignment.ActivatedMembers

                    # Add new activated members that aren't already present
                    foreach ($newMember in $Reportline.ActivatedMembers) {
                        $existingMember = $mergedActivatedMembers | Where-Object { $_.UserId -eq $newMember.UserId }
                        if (-not $existingMember) {
                            $mergedActivatedMembers += $newMember
                        }
                    }

                    # Update the existing assignment with merged data
                    $existingAssignment.ActivatedMembers = $mergedActivatedMembers
                    $existingAssignment.AssignmentType = if (@($mergedActivatedMembers).Count -gt 0) { "Eligible (Active)" } else { "Eligible" }

                    if ($DebugMode) { Write-Host "debug: Merged assignment now has $(@($mergedActivatedMembers).Count) activated members" -ForegroundColor Green }
                }
                # If both are permanent, keep the first one
                else {
                    Write-Host "SKIPPING: duplicate assignment for: $Principal - $assignedRole (keeping existing $($existingAssignment.AssignmentType))" -ForegroundColor Yellow
                }
            } else {
                # First time seeing this role assignment, add it
                if ($DebugMode) { Write-Host "debug: Adding new assignment: $Principal - $assignedRole ($status) with $($Reportline.ActivatedMembers.Count) activated members" -ForegroundColor Cyan }
                $roleAssignmentTracker[$uniqueKey] = $Reportline
            }
        }

        # Convert the hashtable values back to an array and sort for consistent output
        $report = $roleAssignmentTracker.Values | Sort-Object Principal, "Assigned Role"

        # Instead of filtering out user assignments, use them to populate activated members
        Write-Host "INFO: Analyzing user assignments to identify activated members..." -ForegroundColor Cyan

        # Create a hashtable to track activated members by group and role
        $activatedMembersByGroup = @{}
        # Track which user assignments should be removed (they are activated PIM roles, not true permanent assignments)
        $userAssignmentsToRemove = @()

        # Look for user assignments that might be activated roles
        foreach ($assignment in $report) {
            if ($assignment.PrincipalType -eq "User") {
                # Check if there's an equivalent group assignment for the same role and scope
                $equivalentGroupAssignments = $report | Where-Object {
                    $_.PrincipalType -eq "Group" -and
                    $_."Assigned Role" -eq $assignment."Assigned Role" -and
                    $_.AssignedRoleScopeName -eq $assignment.AssignedRoleScopeName -and
            ($_.AssignmentType -eq "Eligible" -or $_.AssignmentType -eq "Eligible (Active)")
                }

                foreach ($groupAssignment in $equivalentGroupAssignments) {
                    # Check if the user is a member of this group
                    $groupMembers = $securityGroupMembers[$groupAssignment.Principal]
                    if ($groupMembers -and $groupMembers.Members) {
                        $userIsMember = $groupMembers.Members | Where-Object {
                            $_.UserPrincipalName -eq $assignment.Principal
                        }

                        if ($userIsMember) {
                            if ($DebugMode) { Write-Host "DETECTED ACTIVATION: User $($assignment.Principal) has active assignment for $($assignment."Assigned Role") through group $($groupAssignment.DisplayName)" -ForegroundColor Green }

                            # Create the group key
                            $groupKey = "$($groupAssignment.Principal)|$($assignment."Assigned Role")|$($assignment.AssignedRoleScopeName)"

                            if (-not $activatedMembersByGroup.ContainsKey($groupKey)) {
                                $activatedMembersByGroup[$groupKey] = @()
                            }

                            # Add this user as an activated member
                            $activatedMember = [PSCustomObject]@{
                                UserPrincipalName = $assignment.Principal
                                DisplayName       = $assignment.DisplayName
                                UserId            = $userIsMember.Id
                                ActivationTime    = if ($assignment.AssignmentStartDate -ne "Permanent") { $assignment.AssignmentStartDate } else { "N/A" }
                                StartTime         = if ($assignment.AssignmentStartDate -ne "Permanent") { $assignment.AssignmentStartDate } else { "N/A" }
                                EndTime           = if ($assignment.AssignmentEndDate -ne "Permanent") { $assignment.AssignmentEndDate } else { "N/A" }
                                Duration          = if ($assignment.AssignmentStartDate -ne "Permanent" -and $assignment.AssignmentEndDate -ne "Permanent") {
                                    try {
                                        $start = [DateTime]::Parse($assignment.AssignmentStartDate)
                                        $end = [DateTime]::Parse($assignment.AssignmentEndDate)
                                        $timeSpan = $end - $start
                                        if ($timeSpan.TotalDays -ge 1) {
                                            "$([math]::Floor($timeSpan.TotalDays)) days, $($timeSpan.Hours) hours"
                                        } else {
                                            "$($timeSpan.Hours) hours, $($timeSpan.Minutes) minutes"
                                        }
                                    } catch { "N/A" }
                                } else { "N/A" }
                                AssignmentState   = "Active"
                                MemberType        = "Group"
                                ActivatedBy       = "N/A"
                                Justification     = "N/A"
                            }

                            # Check if this user is already in the list for this group
                            $existingMember = $activatedMembersByGroup[$groupKey] | Where-Object {
                                $_.UserPrincipalName -eq $activatedMember.UserPrincipalName
                            }

                            if (-not $existingMember) {
                                $activatedMembersByGroup[$groupKey] += $activatedMember
                                # Mark this user assignment for removal since it's represented in the group
                                $userAssignmentsToRemove += $assignment
                                if ($DebugMode) { Write-Host "  Added $($assignment.Principal) as activated member for group $($groupAssignment.DisplayName) - marking user assignment for removal" -ForegroundColor Cyan }
                            }
                        }
                    }
                }

                # Also check for direct PIM activations (user assignments that appear but are actually activated PIM roles)
                # If this user has an "Permanent" assignment but there are role activations, it might be an activated PIM role
                if ($assignment.AssignmentType -eq "Permanent" -and $roleActivations.Count -gt 0) {
                    # Check if there's a corresponding activation in the roleActivations collection
                    $userActivation = $roleActivations | Where-Object {
                        $_.principalId -eq $assignment.Principal -and
                        $_.assignmentType -eq "Activated"
                    }

                    if ($userActivation) {
                        # This user assignment is actually an activated PIM role, mark it for removal
                        $userAssignmentsToRemove += $assignment
                        if ($DebugMode) { Write-Host "  Detected direct PIM activation for $($assignment.Principal) - $($assignment."Assigned Role") - marking for removal" -ForegroundColor Yellow }
                    }
                }
            }
        }

        # Now update the group assignments with the detected activated members
        foreach ($groupAssignment in $report | Where-Object { $_.PrincipalType -eq "Group" }) {
            $groupKey = "$($groupAssignment.Principal)|$($groupAssignment."Assigned Role")|$($groupAssignment.AssignedRoleScopeName)"

            if ($activatedMembersByGroup.ContainsKey($groupKey)) {
                $detectedMembers = $activatedMembersByGroup[$groupKey]

                # Merge with any existing activated members (from the original PIM API detection)
                $allActivatedMembers = @()

                # Add existing activated members first
                if ($groupAssignment.ActivatedMembers -and $groupAssignment.ActivatedMembers.Count -gt 0) {
                    $allActivatedMembers += $groupAssignment.ActivatedMembers
                }

                # Add newly detected members if they're not already there
                foreach ($detectedMember in $detectedMembers) {
                    $existingMember = $allActivatedMembers | Where-Object {
                        $_.UserPrincipalName -eq $detectedMember.UserPrincipalName
                    }
                    if (-not $existingMember) {
                        $allActivatedMembers += $detectedMember
                    }
                }

                # Update the group assignment
                $groupAssignment.ActivatedMembers = $allActivatedMembers
                $groupAssignment.AssignmentType = if (@($allActivatedMembers).Count -gt 0) { "Eligible (Active)" } else { $groupAssignment.AssignmentType }

                if ($DebugMode) { Write-Host "Updated group $($groupAssignment.DisplayName) with $(@($allActivatedMembers).Count) total activated members" -ForegroundColor Green }
            }
        }

        # Remove user assignments that are actually activated PIM roles (to avoid duplicates)
        if (@($userAssignmentsToRemove).Count -gt 0) {
            if ($DebugMode) { Write-Host "INFO: Removing $(@($userAssignmentsToRemove).Count) user assignments that are activated PIM roles (to avoid duplicates)..." -ForegroundColor Cyan }

            $originalCount = $report.Count
            $report = $report | Where-Object {
                $currentAssignment = $_
                -not ($userAssignmentsToRemove | Where-Object {
                        $_.Principal -eq $currentAssignment.Principal -and
                        $_."Assigned Role" -eq $currentAssignment."Assigned Role" -and
                        $_.AssignedRoleScopeName -eq $currentAssignment.AssignedRoleScopeName
                    })
            }
            $newCount = $report.Count

            if ($DebugMode) { Write-Host "debug: Removed $($originalCount - $newCount) user assignments from report (activated PIM roles)" -ForegroundColor Green }
        }

        # Collect subset of roles for each principal type
        $GroupAssignmentReport = $report | Where-Object { $_.PrincipalType -eq "group" }
        $ServicePrincipalReport = $report | Where-Object { $_.PrincipalType -eq "service Principal" } | Select-Object -ExcludeProperty Members
        $UserAssignmentReport = $report | Where-Object { $_.PrincipalType -eq "user" } | Select-Object -ExcludeProperty Members

        # Create a summary of the report
        $GroupMembershipOverviewReport = @()
        foreach ($group in $GroupAssignmentReport) {
            # Look up group members from our previously collected data
            $members = ($securityGroupMembers.Values | Where-Object { $_.groupid -eq $group.Principal }).members.userprincipalname -join ", "

            if (-not $members) {
                $members = "None"
            } else {
                $members = $members -join ", "
            }

            # Get activated members for this group
            $activatedMembersText = "None"
            if ($group.ActivatedMembers -and @($group.ActivatedMembers).Count -gt 0) {
                $activatedList = @()
                foreach ($activatedMember in $group.ActivatedMembers) {
                    $activationInfo = $activatedMember.UserPrincipalName
                    if ($activatedMember.ActivationTime -and $activatedMember.ActivationTime -ne "N/A") {
                        $activationInfo += " (Active since: $($activatedMember.ActivationTime)"
                        if ($activatedMember.EndTime -and $activatedMember.EndTime -ne "N/A") {
                            $activationInfo += " until $($activatedMember.EndTime)"
                        }
                        $activationInfo += ")"
                    }
                    $activatedList += $activationInfo
                }
                $activatedMembersText = $activatedList -join ", "
            }

            $Reportline = [PSCustomObject]@{
                Principal        = $group.Principal
                DisplayName      = $group.DisplayName
                Members          = $members
                ActivatedMembers = $activatedMembersText
            }

            $GroupMembershipOverviewReport += $Reportline
        }

        $GroupMembershipOverviewReport = $GroupMembershipOverviewReport | Select-Object -Property Principal, DisplayName, Members -Unique

        # Get PIM audit logs - move this earlier to use for activation tracking
        $PIMAuditLogsReport = Get-PIMAuditLogs
        if ($PIMAuditLogsReport) {
            Write-Host "INFO: Found $($PIMAuditLogsReport.Count) PIM audit logs." -ForegroundColor Green
        } else {
            Write-Host "INFO: No PIM audit logs found." -ForegroundColor Yellow
        }

        # DEBUG: Final summary of activated members across all groups
        Write-Host "=== FINAL ACTIVATED MEMBERS SUMMARY ===" -ForegroundColor Magenta
        $totalActivatedMembers = 0
        foreach ($groupAssignment in $GroupAssignmentReport) {
            if ($groupAssignment.ActivatedMembers -and @($groupAssignment.ActivatedMembers).Count -gt 0) {
                $totalActivatedMembers += @($groupAssignment.ActivatedMembers).Count
                Write-Host "Group: $($groupAssignment.DisplayName) - Role: $($groupAssignment.'Assigned Role') - Activated Members: $(@($groupAssignment.ActivatedMembers).Count)" -ForegroundColor Green
                foreach ($member in $groupAssignment.ActivatedMembers) {
                    Write-Host "  - $($member.UserPrincipalName)" -ForegroundColor Cyan
                }
            } else {
                Write-Host "Group: $($groupAssignment.DisplayName) - Role: $($groupAssignment.'Assigned Role') - Activated Members: 0" -ForegroundColor Yellow
            }
        }
        Write-Host "TOTAL ACTIVATED MEMBERS ACROSS ALL GROUPS: $totalActivatedMembers" -ForegroundColor Magenta

        # If no activated members found anywhere, provide suggestions
        if ($DebugMode -and $totalActivatedMembers -eq 0) {
            Write-Host "=== NO ACTIVATED MEMBERS FOUND - TROUBLESHOOTING ===" -ForegroundColor Red
            Write-Host "Possible reasons:" -ForegroundColor Yellow
            Write-Host "1. No users have currently activated PIM roles through group assignments" -ForegroundColor Yellow
            Write-Host "2. All role activations have expired" -ForegroundColor Yellow
            Write-Host "3. Missing permissions to read role activation data" -ForegroundColor Yellow
            Write-Host "4. The tenant doesn't have Azure AD P2 licensing for PIM" -ForegroundColor Yellow
            Write-Host "5. Role activations are happening through direct assignments, not group assignments" -ForegroundColor Yellow
            Write-Host "" -ForegroundColor Yellow
            Write-Host "To test if the feature works, try:" -ForegroundColor Cyan
            Write-Host "1. Activate a PIM role through a group assignment" -ForegroundColor Cyan
            Write-Host "2. Run this script while the activation is still active" -ForegroundColor Cyan
            Write-Host "3. Check the 'Activated Members' column in the Group Assignments section" -ForegroundColor Cyan
        }
        if ($DebugMode) { Write-Host "=============================================" -ForegroundColor Magenta }

        New-AdminRoleHTMLReport -TenantName $tenantname -Report $Report -UserAssignmentReport $UserAssignmentReport -GroupAssignmentReport $GroupAssignmentReport -ServicePrincipalReport $ServicePrincipalReport -GroupMembershipOverviewReport $GroupMembershipOverviewReport -PIMAuditLogsReport $PIMAuditLogsReport -ExportPath $ExportPath

        # Send email with the report
        if ($SendEmail) {
            $subject = "$tenantname - Microsoft Entra ID Admin Roles Report"
            $bodyHtml = "<html><body style='font-family: Segoe UI, Arial, sans-serif;'><h2>Microsoft Entra ID Admin Roles Report</h2><p>Attached is the latest Microsoft Entra ID administrative role assignments report for $tenantname.</p><p>Open the attached HTML in a browser for the full report.</p><p style='color:#666;'>Generated by RKSolutions - please do not reply.</p></body></html>"
            $emailSent = Send-EmailWithAttachment -Recipient $Recipient -AttachmentPath $script:ExportPath -From $From -Subject $subject -BodyHtml $bodyHtml

            if ($emailSent) {
                Write-Host "INFO: Email sent successfully." -ForegroundColor Green
            } else {
                Write-Host "ERROR: Failed to send email." -ForegroundColor Red
            }
        } else {
            Write-Host "INFO: Email sending is disabled. Set -SendEmail to $true to enable." -ForegroundColor Yellow
        }

        # Clean up the report file
        if ($SendEmail) {
            if (Test-Path -Path $script:ExportPath) {
                Remove-Item -Path $script:ExportPath -Force
                Write-Host "INFO: Temporary report file deleted." -ForegroundColor Green
            } else {
                Write-Host "INFO: No temporary report file found to delete." -ForegroundColor Yellow
            }
        }
}
