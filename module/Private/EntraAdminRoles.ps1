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

    # Create HTML Template with DataTables
    $htmlTemplate = @'
        <!DOCTYPE html>
        <html lang="en">
        <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>$TenantName Admin Roles Report</title>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/5.3.0/css/bootstrap.min.css">
        <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css">
        <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.4.1/css/buttons.bootstrap5.min.css">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <script src="https://code.jquery.com/jquery-3.7.0.js"></script>
        <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
        <script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
        <script src="https://cdn.datatables.net/buttons/2.4.1/js/dataTables.buttons.min.js"></script>
        <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.bootstrap5.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/pdfmake.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/vfs_fonts.js"></script>
        <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.html5.min.js"></script>
        <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.print.min.js"></script>
        <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.colVis.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
        <style>
        :root {
            /* Light mode variables (default) */
            --primary-color: #0078d4;
            --secondary-color: #2b88d8;
            --permanent-color: #d83b01;
            --eligible-color: #107c10;
            --group-color: #5c2d91;
            --disabled-color: #d9534f;
            --enabled-color: #5cb85c;
            --service-principal-color: #0078d4;
            --na-color: #6c757d;
            --bg-color: #f8f9fa;
            --card-bg: #ffffff;
            --text-color: #333333;
            --table-header-bg: #f5f5f5;
            --table-header-color: #333333;
            --table-stripe-bg: rgba(0,0,0,0.02);
            --table-hover-bg: rgba(0,0,0,0.04);
            --table-border-color: #dee2e6;
            --filter-tag-bg: #e9ecef;
            --filter-tag-color: #495057;
            --filter-bg: white;
            --btn-outline-color: #6c757d;
            --border-color: #dee2e6;
            --toggle-bg: #ccc;
            --button-bg: #f8f9fa;
            --button-color: #333;
            --button-border: #ddd;
            --button-hover-bg: #e9ecef;
            --footer-text: white;
            --input-bg: #fff;
            --input-color: #333;
            --input-border: #ced4da;
            --input-focus-border: #86b7fe;
            --input-focus-shadow: rgba(13, 110, 253, 0.25);
            --datatable-even-row-bg: #fff;
            --datatable
            --tab-active-color: #fff;
        }
        
        [data-theme="dark"] {
            /* Dark mode variables */
            --primary-color: #0078d4;
            --secondary-color: #2b88d8;
            --permanent-color: #d83b01;
            --eligible-color: #107c10;
            --group-color: #5c2d91;
            --disabled-color: #6c757d;
            --enabled-color: #0078d4;
            --service-principal-color: #0078d4;
            --bg-color: #121212;
            --card-bg: #1e1e1e;
            --text-color: #e0e0e0;
            --table-header-bg: #333333;
            --table-header-color: #e0e0e0;
            --table-header-bg: #333333;#e0e0e0;5,255,0.03);
            --table-header-color: #e0e0e0;
            --table-stripe-bg: rgba(255,255,255,0.03);
            --table-hover-bg: rgba(255,255,255,0.05);
            --table-border-color: #444444;
            --filter-bg: #252525;
            --btn-outline-color: #adb5bd;
            --border-color: #444444;
            --toggle-bg: #555555;
            --button-bg: #2a2a2a;
            --button-color: #e0e0e0;
            --button-border: #444;
            --button-hover-bg: #3a3a3a;
            --footer-text: white;
            --input-bg: #2a2a2a;
            --input-color: #e0e0e0;
            --input-border: #444444;
            --input-focus-border: #0078d4;
            --input-focus-shadow: rgba(0, 120, 212, 0.25);
            --datatable-even-row-bg: #1e1e1e;
            --datatable-odd-row-bg: #252525;
            --tab-active-bg: #0078d4;
            --tab-active-color: #fff;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background-color: var(--bg-color);
            color: var(--text-color);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            transition: background-color 0.3s ease, color 0.3s ease;
        }
        
        .container-fluid {
            max-width: 1600px;
            padding: 20px;
            flex: 1;
        }
        
        .dashboard-header {
            padding: 20px 0;
            margin-bottom: 30px;
            border-bottom: 1px solid rgba(128,128,128,0.2);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        
        .dashboard-title {
            display: flex;
            align-items: center;
            gap: 15px;
        }
        
        .dashboard-title h1 {
            margin: 0;
            font-size: 1.8rem;
            font-weight: 600;
            color: var(--primary-color);
        }
        
        .logo {
            height: 45px;
            width: 45px;
        }
        
        .report-date {
            font-size: 0.9rem;
            color: var(--text-color);
            opacity: 0.8;
        }
        
        .card {
            border: none;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.05);
            margin-bottom: 25px;
            transition: transform 0.2s, box-shadow 0.2s, background-color 0.3s ease;
            overflow: hidden;
            background-color: var(--card-bg);
        }
        
        .card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 16px rgba(0,0,0,0.1);
        }
        
        .card-header {
            background-color: var(--primary-color);
            color: white;
            font-weight: 600;
            padding: 15px 20px;
            border-bottom: none;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
        }
        
        .card-header i {
            font-size: 1.2rem;
        }
        
        .card-body {
            padding: 20px;
        }
        
        .stats-card {
            height: 100%;
            text-align: center;
            padding: 25px 15px;
            border-radius: 10px;
            color: white;
            position: relative;
            overflow: hidden;
            min-height: 160px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            cursor: pointer;
            transition: all 0.3s;
        }
        
        .stats-card.active {
            box-shadow: 0 0 0 4px rgba(255,255,255,0.6), 0 8px 16px rgba(0,0,0,0.2);
            transform: scale(1.05);
        }
        
        .stats-card::before {
            content: '';
            position: absolute;
            top: -20px;
            right: -20px;
            width: 100px;
            height: 100px;
            border-radius: 50%;
            background-color: rgba(255,255,255,0.1);
            z-index: 0;
        }
        
        .stats-card i {
            font-size: 2.5rem;
            margin-bottom: 15px;
            position: relative;
            z-index: 1;
        }
        
        .stats-card h3 {
            font-size: 1rem;
            font-weight: 500;
            margin-bottom: 10px;
            position: relative;
            z-index: 1;
        }
        
        .stats-card .number {
            font-size: 2.2rem;
            font-weight: 700;
            position: relative;
            z-index: 1;
        }
        
        .permanent-bg {
            background: linear-gradient(135deg, var(--permanent-color), #f25c05);
        }
        
        .eligible-bg {
            background: linear-gradient(135deg, var(--eligible-color), #2a9d2a);
        }
        
        .group-bg {
            background: linear-gradient(135deg, var(--group-color), #7b4db2);
        }

        .disabled-bg {
            background: linear-gradient(135deg, var(--disabled-color), #6c757d);
        }

        .enabled-bg {
            background: linear-gradient(135deg, var(--enabled-color), #0078d4);
        }
        
        .service-principal-bg {
            background: linear-gradient(135deg, var(--service-principal-color), #2b88d8);
        }
        
        /* DataTables Dark Mode Overrides */
        table.dataTable {
            border-collapse: collapse !important;
            width: 100% !important;
            color: var(--text-color) !important;
            border-color: var(--table-border-color) !important;
        }
        
        .table {
            color: var(--text-color) !important;
            border-color: var(--table-border-color) !important;
        }
        
        .table-striped>tbody>tr:nth-of-type(odd) {
            background-color: var(--datatable-odd-row-bg) !important;
        }
        
        .table-striped>tbody>tr:nth-of-type(even) {
            background-color: var(--datatable-even-row-bg) !important;
        }
        
        .table thead th {
            background-color: var(--table-header-bg) !important;
            color: var(--table-header-color) !important;
            font-weight: 600;
            border-top: none;
            padding: 12px;
            border-color: var(--table-border-color) !important;
        }
        
        .table tbody td {
            padding: 12px;
            vertical-align: middle;
            border-color: var(--table-border-color) !important;
            color: var(--text-color) !important;
        }
        
        .table.table-bordered {
            border-color: var(--table-border-color) !important;
        }
        
        .table-bordered td, .table-bordered th {
            border-color: var(--table-border-color) !important;
        }
        
        .table-hover tbody tr:hover {
            background-color: var(--table-hover-bg) !important;
        }
        
        .badge {
            padding: 6px 10px;
            font-weight: 500;
            border-radius: 6px;
        }
        
        .badge-permanent {
            background-color: var(--permanent-color);
            color: white;
        }
        
        .badge-eligible {
            background-color: var(--eligible-color);
            color: white;
        }
        
        .badge-eligible-active {
            background-color: var(--eligible-color);
            color: white;
            box-shadow: 0 0 8px rgba(16, 124, 16, 0.6), 0 0 16px rgba(16, 124, 16, 0.4);
            border: 2px solid rgba(255, 255, 255, 0.3);
            animation: shimmer 2s infinite;
        }
        
        /* Clickable badge specific styles */
        .badge-eligible-active.group-jump-link {
            transition: all 0.3s ease;
            cursor: pointer;
            position: relative;
        }
        
        .badge-eligible-active.group-jump-link:hover {
            background-color: #0a5a0a;
            box-shadow: 0 0 12px rgba(16, 124, 16, 0.8), 0 0 24px rgba(16, 124, 16, 0.6);
            transform: translateY(-1px);
            border-color: rgba(255, 255, 255, 0.5);
        }
        
        .badge-eligible-active.group-jump-link:active {
            transform: translateY(0);
            box-shadow: 0 0 6px rgba(16, 124, 16, 0.4), 0 0 12px rgba(16, 124, 16, 0.3);
        }
        
        @keyframes shimmer {
            0% { box-shadow: 0 0 8px rgba(16, 124, 16, 0.6), 0 0 16px rgba(16, 124, 16, 0.4); }
            50% { box-shadow: 0 0 12px rgba(16, 124, 16, 0.8), 0 0 24px rgba(16, 124, 16, 0.6); }
            100% { box-shadow: 0 0 8px rgba(16, 124, 16, 0.6), 0 0 16px rgba(16, 124, 16, 0.4); }
        }
        
        .badge-active {
            background-color: var(--service-principal-color);
            color: white;
        }
        
        .badge-group {
            background-color: var(--group-color);
            color: white;
        }
        
        /* Highlight effect for group rows when jumping from main table */
        .highlight-row {
            background-color: rgba(255, 193, 7, 0.3) !important;
            border: 2px solid #ffc107 !important;
            animation: highlightPulse 3s ease-in-out;
            transition: all 0.3s ease;
            position: relative;
            z-index: 10; /* Ensure highlight appears above table headers */
        }
        
        .highlight-row td {
            position: relative;
            z-index: 10; /* Ensure highlighted cells appear above headers */
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
        
        .dataTables_wrapper .dataTables_length, 
        .dataTables_wrapper .dataTables_filter, 
        .dataTables_wrapper .dataTables_info, 
        .dataTables_wrapper .dataTables_processing, 
        .dataTables_wrapper .dataTables_paginate {
            color: var(--text-color) !important;
        }
        
        .dataTables_wrapper .dataTables_paginate .paginate_button {
            padding: 0.3em 0.8em;
            border-radius: 4px;
            margin: 0 3px;
            color: var(--text-color) !important;
            border: 1px solid var(--border-color) !important;
            background-color: var(--button-bg) !important;
        }
        
        .dataTables_wrapper .dataTables_paginate .paginate_button.current {
            background: var(--primary-color) !important;
            border-color: var(--primary-color) !important;
            color: white !important;
        }
        
        .dataTables_wrapper .dataTables_paginate .paginate_button:hover {
            background: var(--button-hover-bg) !important;
            border-color: var(--border-color) !important;
            color: var(--text-color) !important;
        }
        
        .dataTables_wrapper .dataTables_length select,
        .dataTables_wrapper .dataTables_filter input {
            border: 1px solid var(--input-border);
            background-color: var(--input-bg);
            color: var(--input-color);
            border-radius: 4px;
            padding: 5px 10px;
        }
        
        .dataTables_wrapper .dataTables_filter input:focus {
            border-color: var(--input-focus-border);
            box-shadow: 0 0 0 0.25rem var(--input-focus-shadow);
        }
        
        .dataTables_info {
            padding-top: 10px;
            color: var(--text-color);
        }
        
        footer {
            background-color: var(--primary-color);
            color: var(--footer-text);
            text-align: center;
            padding: 15px 0;
            margin-top: auto;
        }
        
        footer p {
            margin: 0;
            font-weight: 500;
        }
        
        .filter-buttons {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-bottom: 20px;
        }
        
        .filter-button {
            padding: 8px 16px;
            border-radius: 20px;
            border: none;
            cursor: pointer;
            font-weight: 500;
            transition: all 0.2s;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .filter-button:hover {
            opacity: 0.9;
        }
        
        .filter-button.active {
            box-shadow: 0 0 0 2px rgba(128,128,128,0.2);
        }
        
        .toggle-container {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 15px;
        }
        
        .toggle-switch {
            position: relative;
            display: inline-block;
            width: 60px;
            height: 30px;
        }
        
        .toggle-switch input {
            opacity: 0;
            width: 0;
            height: 0;
        }
        
        .toggle-slider {
            position: absolute;
            cursor: pointer;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: var(--toggle-bg);
            transition: .4s;
            border-radius: 34px;
        }
        
        .toggle-slider:before {
            position: absolute;
            content: "";
            height: 22px;
            width: 22px;
            left: 4px;
            bottom: 4px;
            background-color: white;
            transition: .4s;
            border-radius: 50%;
        }
        
        input:checked + .toggle-slider {
            background-color: var(--primary-color);
        }
        
        input:checked + .toggle-slider:before {
            transform: translateX(30px);
        }
        
        .filter-section {
            background-color: var(--filter-bg);
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            margin-bottom: 20px;
            transition: background-color 0.3s ease;
        }
        
        .filter-section h5 {
            color: var(--primary-color);
            margin-bottom: 12px;
            font-weight: 600;
        }
        
        .filter-tags {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-top: 8px;
        }
        
        .filter-tag {
            background-color: var(--filter-tag-bg);
            padding: 4px 12px;
            border-radius: 16px;
            font-size: 0.85rem;
            color: var(--filter-tag-color);
            display: flex;
            align-items: center;
            gap: 6px;
            transition: background-color 0.3s ease, color 0.3s ease;
        }
        
        .filter-tag i {
            cursor: pointer;
            color: var(--filter-tag-color);
        }
        
        .filter-tag i:hover {
            color: var(--permanent-color);
        }
        
        .custom-search {
            display: flex;
            gap: 10px;
            margin-bottom: 15px;
        }
        
        .custom-search input {
            flex: 1;
            padding: 8px 12px;
            border: 1px solid var(--border-color);
            background-color: var(--bg-color);
            color: var(--text-color);
            border-radius: 4px;
        }
        
        .custom-search button {
            background-color: var(--primary-color);
            color: white;
            border: none;
            border-radius: 4px;
            padding: 8px 16px;
            cursor: pointer;
        }
        
        .enabled-filters-container {
            margin-bottom: 15px;
        }
        
        /* Show all entries toggle */
        .show-all-container {
            display: flex;
            align-items: center;
            gap: 12px;
            background-color: transparent;
            padding: 0;
            border: none;
            margin-left: 15px;
        }
        
        .show-all-text {
            font-weight: 500;
            margin: 0;
            color: white;
            font-size: 0.85rem;
        }
        
        /* Custom DataTable controls wrapper */
        .datatable-header {
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            margin-bottom: 1rem;
        }
        
        .datatable-controls {
            display: flex;
            align-items: center;
            gap: 15px;
            flex-wrap: wrap;
        }
        
        /* Theme toggle styles */
        .theme-toggle {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 1000;
            display: flex;
            align-items: center;
            gap: 10px;
            background-color: var(--card-bg);
            padding: 8px 12px;
            border-radius: 30px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            transition: background-color 0.3s ease;
        }
        
        .theme-toggle-switch {
            position: relative;
            display: inline-block;
            width: 50px;
            height: 26px;
        }
        
        .theme-toggle-switch input {
            opacity: 0;
            width: 0;
            height: 0;
        }
        
        .theme-toggle-slider {
            position: absolute;
            cursor: pointer;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: var(--toggle-bg);
            transition: .4s;
            border-radius: 34px;
        }
        
        .theme-toggle-slider:before {
            position: absolute;
            content: "";
            height: 18px;
            width: 18px;
            left: 4px;
            bottom: 4px;
            background-color: white;
            transition: .4s;
            border-radius: 50%;
        }
        
        input:checked + .theme-toggle-slider {
            background-color: var(--primary-color);
        }
        
        input:checked + .theme-toggle-slider:before {
            transform: translateX(24px);
        }
        
        .theme-icon {
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
            color: var(--text-color);
        }
        
        /* Form elements for dark mode */
        .form-select, .form-control {
            background-color: var(--input-bg) !important;
            color: var(--input-color) !important;
            border-color: var(--input-border) !important;
        }
        
        .form-select:focus, .form-control:focus {
            border-color: var(--input-focus-border) !important;
            box-shadow: 0 0 0 0.25rem var(--input-focus-shadow) !important;
        }
        
        .form-label {
            color: var(--text-color);
        }
        
        .btn-outline-secondary {
            color: var(--text-color);
            border-color: var(--border-color);
            background-color: transparent;
        }
        
        .btn-outline-secondary:hover {
            background-color: var(--filter-tag-bg);
            color: var(--text-color);
        }
        
        /* Override for dropdown menus and selects */
        .form-select option {
            background-color: var(--input-bg);
            color: var(--input-color);
        }
        
        /* Fix DataTables odd/even row striping */
        table.dataTable.stripe tbody tr.odd, 
        table.dataTable.display tbody tr.odd {
            background-color: var(--datatable-odd-row-bg) !important;
        }
        
        table.dataTable.stripe tbody tr.even, 
        table.dataTable.display tbody tr.even {
            background-color: var(--datatable-even-row-bg) !important;
        }
        
        /* Fix DataTables background color for hovered rows */
        table.dataTable.hover tbody tr:hover, 
        table.dataTable.display tbody tr:hover {
            background-color: var(--table-hover-bg) !important;
        }
        
        /* Fix DataTables border colors */
        table.dataTable.border-bottom,
        table.dataTable.border-top,
        table.dataTable thead th,
        table.dataTable tfoot th,
        table.dataTable thead td,
        table.dataTable tfoot td {
            border-color: var(--table-border-color) !important;
        }
        
        /* Bootstrap 5 DataTables specific overrides */
        .table-striped>tbody>tr:nth-of-type(odd)>* {
            --bs-table-accent-bg: var(--datatable-odd-row-bg) !important;
            color: var(--text-color) !important;
        }
        
        .table>:not(caption)>*>* {
            background-color: var(--card-bg) !important;
            color: var(--text-color) !important;
        }
        
        .table-striped>tbody>tr {
            background-color: var(--datatable-even-row-bg) !important;
        }
        
        /* Direct cell background colors */
        .table tbody tr td {
            background-color: transparent !important;
        }
        
        /* Force Bootstrap Tables to use the correct colors */
        .table-striped>tbody>tr:nth-of-type(odd) {
            --bs-table-accent-bg: var(--datatable-odd-row-bg) !important;
        }

        /* Report selector tabs */
        .report-tabs {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-bottom: 20px;
        }
        
        .report-tab {
            padding: 10px 20px;
            border-radius: 5px;
            background-color: var(--button-bg);
            color: var(--button-color);
            border: 1px solid var(--button-border);
            cursor: pointer;
            font-weight: 600;
            transition: all 0.2s;
        }
        
        .report-tab:hover {
            background-color: var(--button-hover-bg);
        }
        
        .report-tab.active {
            background-color: var(--tab-active-bg);
            color: var(--tab-active-color);
            border-color: var(--tab-active-bg);
        }
        
        .report-panel {
            display: none;
        }
        
        .report-panel.active {
            display: block;
        }
    </style>
    </head>
    <body>
        <!-- Dark Mode Toggle -->
        <div class="theme-toggle">
            <div class="theme-icon">
                <i class="fas fa-sun"></i>
            </div>
            <label class="theme-toggle-switch">
                <input type="checkbox" id="themeToggle">
                <span class="theme-toggle-slider"></span>
            </label>
            <div class="theme-icon">
                <i class="fas fa-moon"></i>
            </div>
        </div>
        
        <div class="container-fluid">
            <div class="dashboard-header">
                <div class="dashboard-title">
                    <svg class="logo" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48">
                        <path fill="#ff5722" d="M6 6H22V22H6z" transform="rotate(-180 14 14)"/>
                        <path fill="#4caf50" d="M26 6H42V22H26z" transform="rotate(-180 34 14)"/>
                        <path fill="#ffc107" d="M6 26H22V42H6z" transform="rotate(-180 14 34)"/>
                        <path fill="#03a9f4" d="M26 26H42V42H26z" transform="rotate(-180 34 34)"/>
                    </svg>
                    <h1>$TenantName Admin Roles Report</h1>
                </div>
                <div class="report-date">
                    <i class="fas fa-calendar-alt me-2"></i> Report generated on: $ReportDate
                </div>
            </div>
            
            <div class="row mb-4">
                <div class="col-md-3 mb-3">
                    <div class="stats-card permanent-bg" id="permanentFilter">
                        <i class="fas fa-key"></i>
                        <h3>Permanent Assignments</h3>
                        <div class="number">$permanentRoles</div>
                    </div>
                </div>
                <div class="col-md-3 mb-3">
                    <div class="stats-card eligible-bg" id="eligibleFilter">
                        <i class="fas fa-clock"></i>
                        <h3>Eligible Assignments</h3>
                        <div class="number">$eligibleRoles</div>
                    </div>
                </div>
                <div class="col-md-3 mb-3">
                    <div class="stats-card group-bg" id="groupFilter">
                        <i class="fas fa-users"></i>
                        <h3>Group Assignments</h3>
                        <div class="number">$groupAssignedRoles</div>
                    </div>
                </div>
                <div class="col-md-3 mb-3">
                    <div class="stats-card service-principal-bg" id="spFilter">
                        <i class="fas fa-robot"></i>
                        <h3>Service Principal Assignments</h3>
                        <div class="number">$servicePrincipalRoles</div>
                    </div>
                </div>
            </div>
            
            <div class="report-tabs">
            </div>

            <div class="filter-section" id="general-filter-section">
                <h5><i class="fas fa-filter me-2"></i>Filter Options</h5>
                
                <div class="row">
                    <div class="col-md-6">
                        <div class="mb-3">
                            <label for="principalTypeFilter" class="form-label">Principal Type</label>
                            <select id="principalTypeFilter" class="form-select">
                                <option value="">All Types</option>
                                <option value="user">User</option>
                                <option value="group">Group</option>
                                <option value="service Principal">Service Principal</option>
                            </select>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="mb-3">
                            <label for="assignmentTypeFilter" class="form-label">Assignment Type</label>
                            <select id="assignmentTypeFilter" class="form-select">
                                <option value="">All Types</option>
                                <option value="Permanent">Permanent</option>
                                <option value="Eligible">Eligible</option>
                            </select>
                        </div>
                    </div>
                </div>
                
                <div class="mb-3">
                    <label for="roleNameFilter" class="form-label">Role Name</label>
                    <input type="text" id="roleNameFilter" class="form-control" placeholder="Search for role names...">
                </div>

                <div class="mb-3">
                    <label for="scopeFilter" class="form-label">Role Scope</label>
                    <select id="scopeFilter" class="form-select">
                        <option value="">All Scopes</option>
                        <option value="Tenant-Wide">Tenant-Wide</option>
                        <option value="AU/">Administrative Unit</option>
                    </select>
                </div>
                
                <div class="enabled-filters-container">
                    <div class="d-flex justify-content-between align-items-center">
                        <label class="form-label mb-0">Enabled Filters:</label>
                        <button id="clearAllFilters" class="btn btn-sm btn-outline-secondary">Clear All</button>
                    </div>
                    <div class="filter-tags" id="enabledFilters">
                        <!-- Enabled filters will be displayed here -->
                    </div>
                </div>
            </div>
            
            <div id="all-report" class="report-panel active">
                <div class="card">
                    <div class="card-header">
                        <div>
                            <i class="fas fa-user-shield"></i> All Role Assignments
                        </div>
                        <div class="show-all-container">
                            <label class="toggle-switch">
                                <input type="checkbox" id="allShowAllToggle">
                                <span class="toggle-slider"></span>
                            </label>
                            <p class="show-all-text">Show all entries</p>
                        </div>
                    </div>
                    <div class="card-body">
                        <div class="table-responsive">
                            <table id="allRolesTable" class="table table-striped table-bordered" style="width:100%">
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
                                    {{ALL_ROLES_DATA}}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

            <div id="user-report" class="report-panel">
                <div class="card">
                    <div class="card-header">
                        <div>
                            <i class="fas fa-user"></i> User Role Assignments
                        </div>
                        <div class="show-all-container">
                            <label class="toggle-switch">
                                <input type="checkbox" id="userShowAllToggle">
                                <span class="toggle-slider"></span>
                            </label>
                            <p class="show-all-text">Show all entries</p>
                        </div>
                    </div>
                    <div class="card-body">
                        <div class="table-responsive">
                            <table id="userRolesTable" class="table table-striped table-bordered" style="width:100%">
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
                                    {{USER_ROLES_DATA}}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

            <div id="group-report" class="report-panel">
                <div class="card">
                    <div class="card-header">
                        <div>
                            <i class="fas fa-users"></i> Group Role Assignments
                        </div>
                        <div class="show-all-container">
                            <label class="toggle-switch">
                                <input type="checkbox" id="groupShowAllToggle">
                                <span class="toggle-slider"></span>
                            </label>
                            <p class="show-all-text">Show all entries</p>
                        </div>
                    </div>
                    <div class="card-body">
                        <div class="table-responsive">
                            <table id="groupRolesTable" class="table table-striped table-bordered" style="width:100%">
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
                                    {{GROUP_ROLES_DATA}}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

            <div id="service-principal-report" class="report-panel">
                <div class="card">
                    <div class="card-header">
                        <div>
                            <i class="fas fa-robot"></i> Service Principal Role Assignments
                        </div>
                        <div class="show-all-container">
                            <label class="toggle-switch">
                                <input type="checkbox" id="spShowAllToggle">
                                <span class="toggle-slider"></span>
                            </label>
                            <p class="show-all-text">Show all entries</p>
                        </div>
                    </div>
                    <div class="card-body">
                        <div class="table-responsive">
                            <table id="spRolesTable" class="table table-striped table-bordered" style="width:100%">
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
                                    {{SP_ROLES_DATA}}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

            <div id="pim-audit-logs-report" class="report-panel">
            <div class="filter-section" id="pim-filter-section">
                <h5><i class="fas fa-filter me-2"></i>PIM Audit Logs Filters</h5>
                
                <div class="row">
                    <div class="col-md-4">
                        <div class="mb-3">
                            <label for="pimOperationTypeFilter" class="form-label">Operation Type</label>
                            <select id="pimOperationTypeFilter" class="form-select">
                                <option value="">All Operations</option>
                                <option value="Add">Add</option>
                                <option value="Update">Update</option>
                                <option value="Delete">Delete</option>
                                <option value="Activate">Activate</option>
                            </select>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="mb-3">
                            <label for="pimInitiatorFilter" class="form-label">Initiated By</label>
                            <input type="text" id="pimInitiatorFilter" class="form-control" placeholder="Filter by user...">
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="mb-3">
                            <label for="pimResultFilter" class="form-label">Result</label>
                            <select id="pimResultFilter" class="form-select">
                                <option value="">All Results</option>
                                <option value="Success">Success</option>
                                <option value="Failure">Failure</option>
                            </select>
                        </div>
                    </div>
                </div>
                
                <div class="row">
                    <div class="col-md-4">
                        <div class="mb-3">
                            <label for="pimRoleFilter" class="form-label">Role</label>
                            <input type="text" id="pimRoleFilter" class="form-control" placeholder="Filter by role...">
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="mb-3">
                            <label for="pimTargetFilter" class="form-label">Target User</label>
                            <input type="text" id="pimTargetFilter" class="form-control" placeholder="Filter by target user...">
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="mb-3">
                            <label for="pimDateRangeFilter" class="form-label">Date Range</label>
                            <div class="input-group">
                                <input type="date" id="pimStartDateFilter" class="form-control">
                                <span class="input-group-text">to</span>
                                <input type="date" id="pimEndDateFilter" class="form-control">
                            </div>
                        </div>
                    </div>
                </div>
                
                <div class="enabled-filters-container">
                    <div class="d-flex justify-content-between align-items-center">
                        <label class="form-label mb-0">PIM Audit Logs Filters:</label>
                        <button id="clearPimFilters" class="btn btn-sm btn-outline-secondary">Clear Filters</button>
                    </div>
                    <div class="filter-tags" id="pimEnabledFilters">
                        <!-- Enabled PIM filters will be displayed here -->
                    </div>
                </div>
            </div>

            <div class="card">
                <div class="card-header">
                    <div>
                        <i class="fas fa-history"></i> PIM Audit Logs
                    </div>
                    <div class="show-all-container">
                        <label class="toggle-switch">
                            <input type="checkbox" id="pimAuditLogsShowAllToggle">
                            <span class="toggle-slider"></span>
                        </label>
                        <p class="show-all-text">Show all entries</p>
                    </div>
                </div>
                <div class="card-body">
                    <div class="table-responsive">
                        <table id="pimAuditLogsTable" class="table table-striped table-bordered" style="width:100%">
                            <thead>
                                <tr>
                                    <th>Date/Time</th>
                                    <th>Initiated By</th>
                                    <th>Operation Type</th>
                                    <th>Initiator Type</th>
                                    <th>Role</th>
                                    <th>Target</th>
                                    <th>Operation</th>
                                    <th>Result</th>
                                    <th>Role Properties</th>
                                    <th>Justification</th>
                                </tr>
                            </thead>
                            <tbody>
                                {{PIM_AUDIT_LOGS_DATA}}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
        <footer>
            <p>Generated by Roy Klooster - RK Solutions</p>
        </footer>
        
        <script>
            // Initialize DataTables
            $(document).ready(function() {
                // Theme toggling functionality
                const themeToggle = document.getElementById('themeToggle');
                const prefersDarkScheme = window.matchMedia('(prefers-color-scheme: dark)');
                
                // Function to update table colors for dark mode
                function updateTableColors() {
                    // Force all table cells to have the correct background
                    if (document.documentElement.getAttribute('data-theme') === 'dark') {
                        // Dark mode
                        $('table.dataTable tbody tr').css('background-color', 'var(--datatable-even-row-bg)');
                        $('table.dataTable tbody tr:nth-child(odd)').css('background-color', 'var(--datatable-odd-row-bg)');
                        $('table.dataTable tbody td').css('color', 'var(--text-color)');
                        $('table.dataTable thead th').css({
                            'background-color': 'var(--table-header-bg)',
                            'color': 'var(--table-header-color)'
                        });
                    } else {
                        // Light mode
                        $('table.dataTable tbody tr').css('background-color', 'var(--datatable-even-row-bg)');
                        $('table.dataTable tbody tr:nth-child(odd)').css('background-color', 'var(--datatable-odd-row-bg)');
                        $('table.dataTable tbody td').css('color', 'var(--text-color)');
                        $('table.dataTable thead th').css({
                            'background-color': 'var(--table-header-bg)',
                            'color': 'var(--table-header-color)'
                        });
                    }
                }
                
                // Check for saved user preference, or use system preference
                const savedTheme = localStorage.getItem('theme');
                if (savedTheme === 'dark' || (!savedTheme && prefersDarkScheme.matches)) {
                    document.documentElement.setAttribute('data-theme', 'dark');
                    themeToggle.checked = true;
                }
                
                // Add event listener for theme toggle
                themeToggle.addEventListener('change', function() {
                    if (this.checked) {
                        document.documentElement.setAttribute('data-theme', 'dark');
                        localStorage.setItem('theme', 'dark');
                    } else {
                        document.documentElement.setAttribute('data-theme', 'light');
                        localStorage.setItem('theme', 'light');
                    }
                    
                    // Apply the table color changes after theme switch
                    setTimeout(updateTableColors, 50);
                });
                
                // Tab switching functionality with filter section toggle
                $('.report-tab').on('click', function() {
                    // Remove active class from all tabs and panels
                    $('.report-tab').removeClass('active');
                    $('.report-panel').removeClass('active');
                    
                    // Add active class to clicked tab and corresponding panel
                    $(this).addClass('active');
                    const panelId = $(this).data('panel');
                    $(`#${panelId}`).addClass('active');
                    
                    // Toggle visibility of filter sections based on active tab
                    // Initialize filter visibility - hide PIM filter by default
                    $('#pim-filter-section').hide();
                    
                    // Adjust DataTables columns when switching tabs
                    setTimeout(function() {
                        $.fn.dataTable.tables({ visible: true, api: true }).columns.adjust();
                    }, 10);
                });
                
                // Initialize DataTable for all roles
                const allRolesTable = $('#allRolesTable').DataTable({
                    dom: 'Bfrtip',
                    buttons: [
                        {
                            extend: 'collection',
                            text: '<i class="fas fa-download"></i> Export',
                            buttons: [
                                {
                                    extend: 'excel',
                                    text: '<i class="fas fa-file-excel"></i> Excel',
                                    exportOptions: {
                                        columns: ':visible'
                                    }
                                },
                                {
                                    extend: 'csv',
                                    text: '<i class="fas fa-file-csv"></i> CSV',
                                    exportOptions: {
                                        columns: ':visible'
                                    }
                                },
                                {
                                    extend: 'pdf',
                                    text: '<i class="fas fa-file-pdf"></i> PDF',
                                    exportOptions: {
                                        columns: ':visible'
                                    }
                                },
                                {
                                    extend: 'print',
                                    text: '<i class="fas fa-print"></i> Print',
                                    exportOptions: {
                                        columns: ':visible'
                                    }
                                }
                            ]
                        },
                        {
                            extend: 'colvis',
                            text: '<i class="fas fa-columns"></i> Columns'
                        }
                    ],
                    paging: true,
                    searching: true,
                    ordering: true,
                    info: true,
                    responsive: true,
                    lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, "All"]],
                    order: [[3, 'asc']],
                    language: {
                        search: "<i class='fas fa-search'></i> _INPUT_",
                        searchPlaceholder: "Search records...",
                        lengthMenu: "Show _MENU_ entries",
                        info: "Showing _START_ to _END_ of _TOTAL_ entries",
                        paginate: {
                            first: "<i class='fas fa-angle-double-left'></i>",
                            last: "<i class='fas fa-angle-double-right'></i>",
                            next: "<i class='fas fa-angle-right'></i>",
                            previous: "<i class='fas fa-angle-left'></i>"
                        }
                    },
                    drawCallback: function() {
                        updateTableColors();
                    }
                });
                
                // Initialize DataTable for user roles
                const userRolesTable = $('#userRolesTable').DataTable({
                    dom: 'Bfrtip',
                    buttons: [
                        {
                            extend: 'collection',
                            text: '<i class="fas fa-download"></i> Export',
                            buttons: [
                                { extend: 'excel', text: '<i class="fas fa-file-excel"></i> Excel', exportOptions: { columns: ':visible' } },
                                { extend: 'csv', text: '<i class="fas fa-file-csv"></i> CSV', exportOptions: { columns: ':visible' } },
                                { extend: 'pdf', text: '<i class="fas fa-file-pdf"></i> PDF', exportOptions: { columns: ':visible' } },
                                { extend: 'print', text: '<i class="fas fa-print"></i> Print', exportOptions: { columns: ':visible' } }
                            ]
                        },
                        { extend: 'colvis', text: '<i class="fas fa-columns"></i> Columns' }
                    ],
                    paging: true, searching: true, ordering: true, info: true, responsive: true,
                    lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, "All"]],
                    order: [[3, 'asc']],
                    language: { search: "<i class='fas fa-search'></i> _INPUT_", searchPlaceholder: "Search records..." },
                    drawCallback: function() { updateTableColors(); }
                });

                // Initialize DataTable for group roles
                const groupRolesTable = $('#groupRolesTable').DataTable({
                    dom: 'Bfrtip',
                    buttons: [
                        {
                            extend: 'collection',
                            text: '<i class="fas fa-download"></i> Export',
                            buttons: [
                                { extend: 'excel', text: '<i class="fas fa-file-excel"></i> Excel', exportOptions: { columns: ':visible' } },
                                { extend: 'csv', text: '<i class="fas fa-file-csv"></i> CSV', exportOptions: { columns: ':visible' } },
                                { extend: 'pdf', text: '<i class="fas fa-file-pdf"></i> PDF', exportOptions: { columns: ':visible' } },
                                { extend: 'print', text: '<i class="fas fa-print"></i> Print', exportOptions: { columns: ':visible' } }
                            ]
                        },
                        { extend: 'colvis', text: '<i class="fas fa-columns"></i> Columns' }
                    ],
                    paging: true, searching: true, ordering: true, info: true, responsive: true,
                    lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, "All"]],
                    order: [[3, 'asc']],
                    language: { search: "<i class='fas fa-search'></i> _INPUT_", searchPlaceholder: "Search records..." },
                    drawCallback: function() { updateTableColors(); }
                });

                // Initialize DataTable for service principal roles
                const spRolesTable = $('#spRolesTable').DataTable({
                    dom: 'Bfrtip',
                    buttons: [
                        {
                            extend: 'collection',
                            text: '<i class="fas fa-download"></i> Export',
                            buttons: [
                                { extend: 'excel', text: '<i class="fas fa-file-excel"></i> Excel', exportOptions: { columns: ':visible' } },
                                { extend: 'csv', text: '<i class="fas fa-file-csv"></i> CSV', exportOptions: { columns: ':visible' } },
                                { extend: 'pdf', text: '<i class="fas fa-file-pdf"></i> PDF', exportOptions: { columns: ':visible' } },
                                { extend: 'print', text: '<i class="fas fa-print"></i> Print', exportOptions: { columns: ':visible' } }
                            ]
                        },
                        { extend: 'colvis', text: '<i class="fas fa-columns"></i> Columns' }
                    ],
                    paging: true, searching: true, ordering: true, info: true, responsive: true,
                    lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, "All"]],
                    order: [[3, 'asc']],
                    language: { search: "<i class='fas fa-search'></i> _INPUT_", searchPlaceholder: "Search records..." },
                    drawCallback: function() { updateTableColors(); }
                });

                // Initialize DataTable for PIM audit logs
                const pimAuditLogsTable = $('#pimAuditLogsTable').DataTable({
                    dom: 'Bfrtip',
                    buttons: [
                        {
                            extend: 'collection',
                            text: '<i class="fas fa-download"></i> Export',
                            buttons: [
                                { extend: 'excel', text: '<i class="fas fa-file-excel"></i> Excel', exportOptions: { columns: ':visible' } },
                                { extend: 'csv', text: '<i class="fas fa-file-csv"></i> CSV', exportOptions: { columns: ':visible' } },
                                { extend: 'pdf', text: '<i class="fas fa-file-pdf"></i> PDF', exportOptions: { columns: ':visible' } },
                                { extend: 'print', text: '<i class="fas fa-print"></i> Print', exportOptions: { columns: ':visible' } }
                            ]
                        },
                        { extend: 'colvis', text: '<i class="fas fa-columns"></i> Columns' }
                    ],
                    paging: true, 
                    searching: true, 
                    ordering: true, 
                    info: true, 
                    scrollX: true,
                    scrollCollapse: true,
                    fixedHeader: true,
                    responsive: true,
                    lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, "All"]],
                    order: [[0, 'desc']], // Sort by date/time descending
                    language: { search: "<i class='fas fa-search'></i> _INPUT_", searchPlaceholder: "Search records..." },
                    drawCallback: function() { 
                        updateTableColors();
                        
                        // Ensure the search and pagination controls maintain their position
                        $('.dataTables_filter, .dataTables_paginate').css({
                            'position': 'sticky',
                            'right': '0',
                            'background-color': 'var(--card-bg)',
                            'z-index': '10'
                        });
                    }
                });

                // Apply initial table colors
                setTimeout(updateTableColors, 100);

                // Show all toggle functionality
                $('#pimAuditLogsShowAllToggle').on('change', function() {
                    if ($(this).is(':checked')) {
                        pimAuditLogsTable.page.len(-1).draw();
                    } else {
                        pimAuditLogsTable.page.len(10).draw();
                    }
                });

                // Add these event handlers for the other toggles:
                $('#allShowAllToggle').on('change', function() {
                    if ($(this).is(':checked')) {
                        allRolesTable.page.len(-1).draw();
                    } else {
                        allRolesTable.page.len(10).draw();
                    }
                });

                $('#userShowAllToggle').on('change', function() {
                    if ($(this).is(':checked')) {
                        userRolesTable.page.len(-1).draw();
                    } else {
                        userRolesTable.page.len(10).draw();
                    }
                });

                $('#groupShowAllToggle').on('change', function() {
                    if ($(this).is(':checked')) {
                        groupRolesTable.page.len(-1).draw();
                    } else {
                        groupRolesTable.page.len(10).draw();
                    }
                });

                // Add this handler if you have a service principal toggle as well
                $('#spShowAllToggle').on('change', function() {
                    if ($(this).is(':checked')) {
                        spRolesTable.page.len(-1).draw();
                    } else {
                        spRolesTable.page.len(10).draw();
                    }
                });

                // Custom filtering function for PIM audit logs
                $.fn.dataTable.ext.search.push(
                    function(settings, data, dataIndex) {
                        // Only apply to PIM audit logs table
                        if (settings.nTable.id !== 'pimAuditLogsTable') {
                            return true;
                        }
                        
                        // Get PIM filter values
                        const operationType = $('#pimOperationTypeFilter').val().toLowerCase();
                        const initiator = $('#pimInitiatorFilter').val().toLowerCase();
                        const result = $('#pimResultFilter').val().toLowerCase();
                        const role = $('#pimRoleFilter').val().toLowerCase();
                        const target = $('#pimTargetFilter').val().toLowerCase();
                        const startDate = $('#pimStartDateFilter').val();
                        const endDate = $('#pimEndDateFilter').val();
                        
                        // Get row data with column indices for PIM audit logs table
                        const colDateTime = data[0]; // Date/Time column
                        const colInitiator = data[1].toLowerCase(); // Initiated By column
                        const colOperation = data[2].toLowerCase(); // Operation Type column
                        const colRole = data[4].toLowerCase(); // Role column
                        const colTarget = data[5].toLowerCase(); // Target column
                        const colResult = data[7].toLowerCase(); // Result column
                        
                        // Parse date from the colDateTime value
                        let rowDate = null;
                        try {
                            // Extract date part from "MM/DD/YYYY HH:MM:SS" format
                            const dateMatch = colDateTime.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
                            if (dateMatch) {
                                // Create date object (month is 0-indexed in JavaScript)
                                rowDate = new Date(dateMatch[3], dateMatch[1] - 1, dateMatch[2]);
                            } else {
                                // Try parsing as ISO date if not in the expected format
                                rowDate = new Date(colDateTime);
                            }
                        } catch (e) {
                            // If date parsing fails, skip date filtering for this row
                            console.warn("Failed to parse date:", colDateTime);
                        }
                        
                        // Filter by operation type
                        if (operationType && !colOperation.includes(operationType)) {
                            return false;
                        }
                        
                        // Filter by initiator
                        if (initiator && !colInitiator.includes(initiator)) {
                            return false;
                        }
                        
                        // Filter by result
                        if (result && !colResult.includes(result)) {
                            return false;
                        }
                        
                        // Filter by role
                        if (role && !colRole.includes(role)) {
                            return false;
                        }
                        
                        // Filter by target
                        if (target && !colTarget.includes(target)) {
                            return false;
                        }
                        
                        // Filter by date range
                        if (startDate && endDate && rowDate) {
                            const filterStartDate = new Date(startDate);
                            const filterEndDate = new Date(endDate);
                            // Set end date to end of day for inclusive filtering
                            filterEndDate.setHours(23, 59, 59, 999);
                            
                            if (rowDate < filterStartDate || rowDate > filterEndDate) {
                                return false;
                            }
                        } else if (startDate && rowDate) {
                            const filterStartDate = new Date(startDate);
                            if (rowDate < filterStartDate) {
                                return false;
                            }
                        } else if (endDate && rowDate) {
                            const filterEndDate = new Date(endDate);
                            filterEndDate.setHours(23, 59, 59, 999);
                            if (rowDate > filterEndDate) {
                                return false;
                            }
                        }
                        
                        return true;
                    }
                );

                // Apply PIM filters when select inputs change
                $('#pimOperationTypeFilter, #pimResultFilter').on('change', function() {
                    const filterType = $(this).attr('id');
                    const filterValue = $(this).val();
                    const filterLabel = $(this).prev('label').text();
                    
                    if (filterValue) {
                        updatePimEnabledFilter(filterLabel, filterValue);
                    } else {
                        removePimEnabledFilter(filterLabel);
                    }
                    
                    applyPimFilters();
                });

                // Apply PIM filters on text input
                $('#pimInitiatorFilter, #pimRoleFilter, #pimTargetFilter').on('input', function() {
                    const filterType = $(this).attr('id');
                    const filterValue = $(this).val();
                    const filterLabel = $(this).prev('label').text();
                    
                    if (filterValue) {
                        updatePimEnabledFilter(filterLabel, filterValue);
                    } else {
                        removePimEnabledFilter(filterLabel);
                    }
                    
                    // Slight delay for typing
                    clearTimeout($(this).data('timeout'));
                    $(this).data('timeout', setTimeout(function() {
                        applyPimFilters();
                    }, 300));
                });

                // Apply date range filter
                $('#pimStartDateFilter, #pimEndDateFilter').on('change', function() {
                    const startDate = $('#pimStartDateFilter').val();
                    const endDate = $('#pimEndDateFilter').val();
                    
                    if (startDate || endDate) {
                        let dateRangeText = '';
                        if (startDate && endDate) {
                            dateRangeText = `${startDate} to ${endDate}`;
                        } else if (startDate) {
                            dateRangeText = `From ${startDate}`;
                        } else if (endDate) {
                            dateRangeText = `Until ${endDate}`;
                        }
                        updatePimEnabledFilter('Date Range', dateRangeText);
                    } else {
                        removePimEnabledFilter('Date Range');
                    }
                    
                    applyPimFilters();
                });

                // Clear PIM filters button
                $('#clearPimFilters').on('click', function() {
                    // Reset all PIM filter inputs
                    $('#pimOperationTypeFilter').val('');
                    $('#pimInitiatorFilter').val('');
                    $('#pimResultFilter').val('');
                    $('#pimRoleFilter').val('');
                    $('#pimTargetFilter').val('');
                    $('#pimStartDateFilter').val('');
                    $('#pimEndDateFilter').val('');
                    
                    // Clear the PIM filter tags
                    $('#pimEnabledFilters').empty();
                    
                    // Apply the filters (with all values cleared)
                    applyPimFilters();
                });

                // Function to apply PIM filters
                function applyPimFilters() {
                    pimAuditLogsTable.draw();
                }

                // Function to update PIM enabled filters
                function updatePimEnabledFilter(filterType, filterValue) {
                    // Remove existing filter of the same type
                    removePimEnabledFilter(filterType);
                    
                    // Add new filter tag
                    const filterTag = `
                        <div class="filter-tag" data-pim-filter-type="${filterType}">
                            <span>${filterType}: ${filterValue}</span>
                            <i class="fas fa-times-circle remove-pim-filter" data-pim-filter-type="${filterType}"></i>
                        </div>
                    `;
                    
                    $('#pimEnabledFilters').append(filterTag);
                    
                    // Add click handler to remove filter
                    $('.remove-pim-filter').off('click').on('click', function() {
                        const filterTypeToRemove = $(this).data('pim-filter-type');
                        
                        if (filterTypeToRemove === 'Operation Type') {
                            $('#pimOperationTypeFilter').val('');
                        } else if (filterTypeToRemove === 'Initiated By') {
                            $('#pimInitiatorFilter').val('');
                        } else if (filterTypeToRemove === 'Result') {
                            $('#pimResultFilter').val('');
                        } else if (filterTypeToRemove === 'Role') {
                            $('#pimRoleFilter').val('');
                        } else if (filterTypeToRemove === 'Target User') {
                            $('#pimTargetFilter').val('');
                        } else if (filterTypeToRemove === 'Date Range') {
                            $('#pimStartDateFilter').val('');
                            $('#pimEndDateFilter').val('');
                        }
                        
                        $(this).closest('.filter-tag').remove();
                        applyPimFilters();
                    });
                }

                // Function to remove PIM enabled filter by type
                function removePimEnabledFilter(filterType) {
                    $('.filter-tag[data-pim-filter-type="' + filterType + '"]').remove();
                }

                // Make PIM filter options visible when tab is active
                $('.report-tab').on('click', function() {
                    // Remove active class from all tabs and panels
                    $('.report-tab').removeClass('active');
                    $('.report-panel').removeClass('active');
                    
                    // Add active class to clicked tab and corresponding panel
                    $(this).addClass('active');
                    const panelId = $(this).data('panel');
                    $(`#${panelId}`).addClass('active');
                    
                    // Toggle visibility of filter sections based on active tab
                    if (panelId === 'pim-audit-logs-report') {
                        // When PIM audit logs tab is active
                        $('#general-filter-section').hide(); // Hide general filter section
                        $('#pim-filter-section').show(); // Show PIM filter section
                        
                        // Refresh the table when switching to PIM tab to ensure correct column widths
                        setTimeout(function() {
                            pimAuditLogsTable.columns.adjust();
                        }, 10);
                    } else {
                        // For all other tabs
                        $('#general-filter-section').show(); // Show general filter section
                        $('#pim-filter-section').hide(); // Hide PIM filter section
                    }
                });

                
                // Custom filtering function for all tables
                $.fn.dataTable.ext.search.push(
                    function(settings, data, dataIndex) {
                        // Get filter values
                        const principalType = $('#principalTypeFilter').val().toLowerCase();
                        const assignmentType = $('#assignmentTypeFilter').val();
                        const roleName = $('#roleNameFilter').val().toLowerCase();
                        const scopeFilter = $('#scopeFilter').val();
                        
                        // Get row data with CORRECTED column indices
                        const colPrincipalType = data[2].toLowerCase(); // Principal Type column (index 2)
                        const colRole = data[4].toLowerCase(); // Assigned Role column (index 4)
                        const colAssignmentType = data[6]; // Assignment Type column (index 6) 
                        const colScope = data[5]; // Scope column (index 5)
                        
                        // Filter by principal type
                        if (principalType && !colPrincipalType.includes(principalType)) {
                            return false;
                        }
                        
                        // Filter by assignment type
                        if (assignmentType && !colAssignmentType.includes(assignmentType)) {
                            return false;
                        }
                        
                        // Filter by role name
                        if (roleName && !colRole.includes(roleName)) {
                            return false;
                        }
                        
                        // Filter by scope
                        if (scopeFilter && !colScope.toLowerCase().includes(scopeFilter.toLowerCase())) {
                        return false;
                        }
                        
                        return true;
                    }
                );
                
                // Stats card filtering
                $('#permanentFilter').on('click', function() {
                    $('#assignmentTypeFilter').val('Permanent');
                    updateEnabledFilters('Assignment Type', 'Permanent');
                    applyFilters();
                    toggleStatsCardEnabled('permanentFilter');
                });
                
                $('#eligibleFilter').on('click', function() {
                    $('#assignmentTypeFilter').val('Eligible');
                    updateEnabledFilters('Assignment Type', 'Eligible');
                    applyFilters();
                    toggleStatsCardEnabled('eligibleFilter');
                });
                
                $('#groupFilter').on('click', function() {
                    $('#principalTypeFilter').val('group');
                    updateEnabledFilters('Principal Type', 'Group');
                    applyFilters();
                    toggleStatsCardEnabled('groupFilter');
                    
                    // Switch to group report tab
                    $('.report-tab[data-panel="group-report"]').click();
                });
                
                $('#spFilter').on('click', function() {
                    $('#principalTypeFilter').val('service Principal');
                    updateEnabledFilters('Principal Type', 'Service Principal');
                    applyFilters();
                    toggleStatsCardEnabled('spFilter');
                    
                    // Switch to service principal report tab
                    $('.report-tab[data-panel="service-principal-report"]').click();
                });
                
                // Apply filters when select boxes change
                $('#principalTypeFilter, #assignmentTypeFilter, #scopeFilter').on('change', function() {
                    const filterType = $(this).attr('id');
                    const filterValue = $(this).val();
                    
                    if (filterValue) {
                        if (filterType === 'principalTypeFilter') {
                            updateEnabledFilters('Principal Type', filterValue);
                        } else if (filterType === 'assignmentTypeFilter') {
                            updateEnabledFilters('Assignment Type', filterValue);
                        } else if (filterType === 'scopeFilter') {
                            updateEnabledFilters('Scope', filterValue);
                        }
                    } else {
                        if (filterType === 'principalTypeFilter') {
                            removeEnabledFilter('Principal Type');
                        } else if (filterType === 'assignmentTypeFilter') {
                            removeEnabledFilter('Assignment Type');
                        } else if (filterType === 'scopeFilter') {
                            removeEnabledFilter('Scope');
                        }
                    }
                    
                    applyFilters();
                });
                
                // Apply filter when role name input changes
                $('#roleNameFilter').on('input', function() {
                    const filterValue = $(this).val();
                    
                    if (filterValue) {
                        updateEnabledFilters('Role Name', filterValue);
                    } else {
                        removeEnabledFilter('Role Name');
                    }
                    
                    applyFilters();
                });
                
                // Clear all filters button
                $('#clearAllFilters').on('click', function() {
                    // Reset all filter inputs
                    $('#principalTypeFilter').val('');
                    $('#assignmentTypeFilter').val('');
                    $('#roleNameFilter').val('');
                    $('#scopeFilter').val('');
                    
                    // Remove active state from stats cards
                    $('.stats-card').removeClass('active');
                    
                    // Clear the visible filter tags
                    clearEnabledFilters();
                    
                    // This is critical - force DataTables to redraw with cleared filters
                    allRolesTable.search('').columns().search('').draw();
                    userRolesTable.search('').columns().search('').draw();
                    groupRolesTable.search('').columns().search('').draw();
                    spRolesTable.search('').columns().search('').draw();
                    
                    // Apply the filters (with all values cleared)
                    applyFilters();

                    // IMPORTANT: Switch back to "All Assignments" tab
                    $('.report-tab[data-panel="all-report"]').click();
                }); 
                
                // Function to apply all filters
                function applyFilters() {
                    allRolesTable.draw();
                    userRolesTable.draw();
                    groupRolesTable.draw();
                    spRolesTable.draw();
                }
                
                // Function to toggle stats card active state
                function toggleStatsCardEnabled(cardId) {
                    $('.stats-card').removeClass('active');
                    $('#' + cardId).addClass('active');
                }
                
                // Function to update enabled filters
                function updateEnabledFilters(filterType, filterValue) {
                    // Remove existing filter of the same type
                    removeEnabledFilter(filterType);
                    
                    // Add new filter tag
                    const filterTag = `
                        <div class="filter-tag" data-filter-type="${filterType}">
                            <span>${filterType}: ${filterValue}</span>
                            <i class="fas fa-times-circle remove-filter" data-filter-type="${filterType}"></i>
                        </div>
                    `;
                    
                    $('#enabledFilters').append(filterTag);
                    
                    // Add click handler to remove filter
                    $('.remove-filter').off('click').on('click', function() {
                        const filterTypeToRemove = $(this).data('filter-type');
                        
                        if (filterTypeToRemove === 'Principal Type') {
                            $('#principalTypeFilter').val('');
                        } else if (filterTypeToRemove === 'Assignment Type') {
                            $('#assignmentTypeFilter').val('');
                        } else if (filterTypeToRemove === 'Role Name') {
                            $('#roleNameFilter').val('');
                        } else if (filterTypeToRemove === 'Scope') {
                            $('#scopeFilter').val('');
                        }
                        
                        $(this).closest('.filter-tag').remove();
                        
                        // Remove active state from stat cards
                        if (filterTypeToRemove === 'Assignment Type') {
                            if ($('#permanentFilter').hasClass('active') || $('#eligibleFilter').hasClass('active')) {
                                $('#permanentFilter, #eligibleFilter').removeClass('active');
                            }
                        } else if (filterTypeToRemove === 'Principal Type') {
                            if ($('#groupFilter').hasClass('active') || $('#spFilter').hasClass('active')) {
                                $('#groupFilter, #spFilter').removeClass('active');
                            }
                        }
                        
                        applyFilters();
                    });
                }
                
                // Function to remove enabled filter by type
                function removeEnabledFilter(filterType) {
                    $('.filter-tag[data-filter-type="' + filterType + '"]').remove();
                }
                
                // Function to clear all enabled filters
                function clearEnabledFilters() {
                    $('#enabledFilters').empty();
                }
                
                // Force dark mode to take effect on page elements
                $(window).on('load', function() {
                    setTimeout(updateTableColors, 200);
                });
                
                // Re-apply styles after DataTables operations
                allRolesTable.on('draw.dt', function() {
                    setTimeout(updateTableColors, 50);
                });
                
                userRolesTable.on('draw.dt', function() {
                    setTimeout(updateTableColors, 50);
                });
                
                groupRolesTable.on('draw.dt', function() {
                    setTimeout(updateTableColors, 50);
                });
                
                spRolesTable.on('draw.dt', function() {
                    setTimeout(updateTableColors, 50);
                });
                
                // Group jump functionality - handle clicks on group links in main table
                $(document).on('click', '.group-jump-link', function(e) {
                    e.preventDefault();
                    
                    const groupId = $(this).data('group-id');
                    const targetRow = $('#group-' + groupId);
                    
                    if (targetRow.length > 0) {
                        // Switch to Group Assignments tab
                        $('.report-tab').removeClass('active');
                        $('.report-panel').removeClass('active');
                        $('.report-tab[data-panel="group-report"]').addClass('active');
                        $('#group-report').addClass('active');
                        
                        // Show general filter section, hide PIM filter section
                        $('#general-filter-section').show();
                        $('#pim-filter-section').hide();
                        
                        // Adjust DataTables columns for the newly visible tab
                        setTimeout(function() {
                            groupRolesTable.columns.adjust();
                            
                            // Scroll to the target row and highlight it
                            const targetRowElement = targetRow[0];
                            if (targetRowElement) {
                                // Scroll to the row
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
    </body>
    </html>
'@
    
    # Generate table rows for all role assignments
    $allRolesRows = ""
    foreach ($item in $Report) {        
        $assignmentTypeBadge = switch ($item.AssignmentType) {
            "Permanent" { "<span class=`"badge badge-permanent`">Permanent</span>" }
            "Eligible" { "<span class=`"badge badge-eligible`">Eligible</span>" }
            "Eligible (Active)" { 
                if ($item.PrincipalType -eq "group") {
                    # Create a safe ID from the principal name and role
                    $safeId = ($item.Principal + "-" + $item.'Assigned Role').Replace(" ", "-").Replace("@", "-").Replace(".", "-")
                    "<span class=`"badge badge-eligible-active group-jump-link`" data-group-id=`"$safeId`" style=`"cursor: pointer;`" title=`"Click to view group details in Group Assignments tab`">Eligible (Active) <i class=`"fas fa-external-link-alt`" style=`"font-size: 10px; margin-left: 4px;`"></i></span>"
                } else {
                    "<span class=`"badge badge-eligible-active`">Eligible (Active)</span>"
                }
            }
            default { "<span class=`"badge bg-secondary`">Unknown</span>" }
        }
        
        $allRolesRows += @"
        <tr>
            <td>$($item.Principal)</td>
            <td>$($item.DisplayName)</td>
            <td>$($item.PrincipalType)</td>
            <td>$($item.AccountStatus)</td>
            <td>$($item.'Assigned Role')</td>
            <td>$($item.AssignedRoleScopeName)</td>
            <td>$assignmentTypeBadge</td>
            <td>$($item.AssignmentStartDate)</td>
            <td>$($item.AssignmentEndDate)</td>
        </tr>
"@
    }

    # Generate table rows for user role assignments
    $userRolesRows = ""
    foreach ($item in $UserAssignmentReport) {
        $assignmentTypeBadge = switch ($item.AssignmentType) {
            "Permanent" { "<span class=`"badge badge-permanent`">Permanent</span>" }
            "Eligible" { "<span class=`"badge badge-eligible`">Eligible</span>" }
            "Eligible (Active)" { "<span class=`"badge badge-eligible-active`">Eligible (Active)</span>" }
            default { "<span class=`"badge bg-secondary`">Unknown</span>" }
        }
        
        $userRolesRows += @"
    <tr>
        <td>$($item.Principal)</td>
        <td>$($item.DisplayName)</td>
        <td>$($item.PrincipalType)</td>
        <td>$($item.AccountStatus)</td>
        <td>$($item.'Assigned Role')</td>
        <td>$($item.AssignedRoleScopeName)</td>
        <td>$assignmentTypeBadge</td>
        <td>$($item.AssignmentStartDate)</td>
        <td>$($item.AssignmentEndDate)</td>
    </tr>
"@
    }

    # Generate table rows for group role assignments
    $groupRolesRows = ""
    foreach ($item in $GroupAssignmentReport) {
        $assignmentTypeBadge = switch ($item.AssignmentType) {
            "Permanent" { "<span class=`"badge badge-permanent`">Permanent</span>" }
            "Eligible" { "<span class=`"badge badge-eligible`">Eligible</span>" }
            "Eligible (Active)" { "<span class=`"badge badge-eligible-active`">Eligible (Active)</span>" }
            default { "<span class=`"badge bg-secondary`">Unknown</span>" }
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
                $activatedList += "$($activatedMember.UserPrincipalName)"
            }
            $activatedMembersText = $activatedList -join "<br/>"
        }
        
        # Create a safe ID from the principal name and role for targeting
        $safeId = ($item.Principal + "-" + $item.'Assigned Role').Replace(" ", "-").Replace("@", "-").Replace(".", "-")
        
        $groupRolesRows += @"
        <tr id="group-$safeId" class="group-assignment-row">
            <td>$($item.Principal)</td>
            <td>$($item.DisplayName)</td>
            <td>$($item.PrincipalType)</td>
            <td>$($item.AccountStatus)</td>
            <td>$($item.'Assigned Role')</td>
            <td>$($item.AssignedRoleScopeName)</td>
            <td>$assignmentTypeBadge</td>
            <td>$($item.AssignmentStartDate)</td>
            <td>$($item.AssignmentEndDate)</td>
            <td>$groupMembers</td>
            <td style="max-width: 300px; word-wrap: break-word;">$activatedMembersText</td>
        </tr>
"@
    }

    # Generate table rows for service principal role assignments
    $spRolesRows = ""
    foreach ($item in $ServicePrincipalReport) {
        $assignmentTypeBadge = switch ($item.AssignmentType) {
            "Permanent" { "<span class=`"badge badge-permanent`">Permanent</span>" }
            "Eligible" { "<span class=`"badge badge-eligible`">Eligible</span>" }
            "Eligible (Active)" { "<span class=`"badge badge-eligible-active`">Eligible (Active)</span>" }
            default { "<span class=`"badge bg-secondary`">Unknown</span>" }
        }
        
        $spRolesRows += @"
    <tr>
        <td>$($item.Principal)</td>
        <td>$($item.DisplayName)</td>
        <td>$($item.PrincipalType)</td>
        <td>$($item.AccountStatus)</td>
        <td>$($item.'Assigned Role')</td>
        <td>$($item.AssignedRoleScopeName)</td>
        <td>$assignmentTypeBadge</td>
        <td>$($item.AssignmentStartDate)</td>
        <td>$($item.AssignmentEndDate)</td>
    </tr>
"@
    }

    # Generate table rows for PIM audit logs
    $pimAuditLogsRows = ""
    if ($PIMAuditLogsReport -and $PIMAuditLogsReport.Count -gt 0) {
        foreach ($log in $PIMAuditLogsReport) {
            $resultBadge = switch ($log.Result) {
                "Success" { "<span class=`"badge badge-eligible`">Success</span>" }
                "Failure" { "<span class=`"badge badge-permanent`">Failure</span>" }
                default { "<span class=`"badge bg-secondary`">Unknown</span>" }
            }
            
            $pimAuditLogsRows += @"
        <tr>
            <td>$($log.DateTime)</td>
            <td>$($log.InitiatedBy)</td>
            <td>$($log.OperationType)</td>
            <td>$($log.InitiatedByType)</td>
            <td>$($log.Role)</td>
            <td>$($log.Target)</td>
            <td>$($log.Operation)</td>
            <td>$resultBadge</td>
            <td>$($log.RoleProperties)</td>
            <td>$($log.Justification)</td>
        </tr>
"@
        }
    }

    # Replace placeholders in template with actual values
    $htmlContent = $htmlTemplate
    $htmlContent = $htmlContent.Replace('$TenantName', $TenantName)
    $htmlContent = $htmlContent.Replace('$ReportDate', $currentDate)
    $htmlContent = $htmlContent.Replace('$permanentRoles', $permanentRoles)
    $htmlContent = $htmlContent.Replace('$eligibleRoles', $eligibleRoles)
    $htmlContent = $htmlContent.Replace('$groupAssignedRoles', $groupAssignedRoles)
    $htmlContent = $htmlContent.Replace('$servicePrincipalRoles', $servicePrincipalRoles)
    $htmlContent = $htmlContent.Replace('{{PIM_AUDIT_LOGS_DATA}}', $pimAuditLogsRows)

    # Add report tabs based on available data
    $reportTabs = @"
<div class="report-tab active" data-panel="all-report">Assignments</div>
"@

    if ($UserAssignmentReport.Count -gt 0) {
        $reportTabs += @"
<div class="report-tab" data-panel="user-report" style="display:none;">User Assignments</div>
"@
    }
    
    if ($GroupAssignmentReport.Count -gt 0) {
        $reportTabs += @"
<div class="report-tab" data-panel="group-report" style="display:none;">Group Assignments</div>
"@
    }
    
    if ($ServicePrincipalReport.Count -gt 0) {
        $reportTabs += @"
<div class="report-tab" data-panel="service-principal-report" style="display:none;">Service Principal Assignments</div>
"@
    }


    if ($PIMAuditLogsReport -and $PIMAuditLogsReport.Count -gt 0) {
        $reportTabs += @"
<div class="report-tab" data-panel="pim-audit-logs-report">PIM Audit Logs</div>
"@
    }

    
    $htmlContent = $htmlContent.Replace("<div class=`"report-tabs`">", "<div class=`"report-tabs`">`n$reportTabs")

    # Replace table data placeholders
    $htmlContent = $htmlContent.Replace('{{ALL_ROLES_DATA}}', $allRolesRows)
    $htmlContent = $htmlContent.Replace('{{USER_ROLES_DATA}}', $userRolesRows)
    $htmlContent = $htmlContent.Replace('{{GROUP_ROLES_DATA}}', $groupRolesRows)
    $htmlContent = $htmlContent.Replace('{{SP_ROLES_DATA}}', $spRolesRows)

    # Add additional CSS for dark mode pagination
    $darkModePaginationCss = @'
    <style>
        [data-theme="dark"] .page-item.active .page-link {
            color: white !important;
            border-color: var(--primary-color) !important;
        }
        
        [data-theme="dark"] .page-item.disabled .page-link {
            background-color: var(--card-bg) !important;
            color: #6c757d !important;
            border-color: var(--border-color) !important;
        }
        
        [data-theme="dark"] .dataTables_paginate .paginate_button {
            background-color: var(--button-bg) !important;
            color: var(--text-color) !important;
            border-color: var(--border-color) !important;
        }
        
        [data-theme="dark"] .dataTables_paginate .paginate_button.current, 
        [data-theme="dark"] .dataTables_paginate .paginate_button.current:hover {
            background: var(--primary-color) !important;
            color: white !important;
            border-color: var(--primary-color) !important;
        }
        
        [data-theme="dark"] .dataTables_paginate .paginate_button:hover {
            background: var(--button-hover-bg) !important;
            color: var(--text-color) !important;
            border-color: var(--button-border) !important;
        }
        
        [data-theme="dark"] .dataTables_paginate .paginate_button.disabled, 
        [data-theme="dark"] .dataTables_paginate .paginate_button.disabled:hover {
            background-color: var(--card-bg) !important;
            color: #6c757d !important;
            border-color: var(--border-color) !important;
            opacity: 0.6;
        }
        
        .dataTables_wrapper .dataTables_filter,
        .dataTables_wrapper .dataTables_paginate {
            position: sticky;
            right: 0;
            background-color: var(--card-bg);
            padding: 5px;
            z-index: 10;
        }
        
        .dataTables_wrapper .dataTables_info,
        .dataTables_wrapper .dataTables_length {
            position: sticky;
            left: 0;
            background-color: var(--card-bg);
            padding: 5px;
            z-index: 10;
        }
    </style>
'@

    # Insert the dark mode pagination CSS before the </head> tag
    $htmlContent = $htmlContent.Replace("</head>", "$darkModePaginationCss`n</head>")

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
                    # Use HTML arrow entity instead of Unicode character
                    $roleProperties += "$($propName): $oldValue &rarr; $newValue"
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
