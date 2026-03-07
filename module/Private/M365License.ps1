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
    
    # Create HTML Template with DataTables
    $htmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$Organization M365 Licensing Report</title>
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
            --direct-color: #0078d4;
            --inherited-color: #107c10;
            --both-color: #5c2d91;
            --Disabled-color: #d83b01;
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
            --datatable-odd-row-bg: rgba(0,0,0,0.02);
        }
        
        [data-theme="dark"] {
            /* Dark mode variables */
            --primary-color: #0078d4;
            --secondary-color: #2b88d8;
            --direct-color: #0078d4;
            --inherited-color: #107c10;
            --both-color: #5c2d91;
            --Disabled-color: #d83b01;
            --bg-color: #121212;
            --card-bg: #1e1e1e;
            --text-color: #e0e0e0;
            --table-header-bg: #333333;
            --table-header-color: #e0e0e0;
            --table-stripe-bg: rgba(255,255,255,0.03);
            --table-hover-bg: rgba(255,255,255,0.05);
            --table-border-color: #444444;
            --filter-tag-bg: #2d2d2d;
            --filter-tag-color: #d0d0d0;
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
        
        .stats-card.Enabled {
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
        
        .direct-bg {
            background: linear-gradient(135deg, var(--direct-color), #2b88d8);
        }
        
        .inherited-bg {
            background: linear-gradient(135deg, var(--inherited-color), #2a9d2a);
        }
        
        .both-bg {
            background: linear-gradient(135deg, var(--both-color), #7b4db2);
        }
        
        .Disabled-bg {
            background: linear-gradient(135deg, var(--Disabled-color), #f25c05);
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
        
        .badge-direct {
            background-color: var(--direct-color);
            color: white;
        }
        
        .badge-inherited {
            background-color: var(--inherited-color);
            color: white;
        }
        
        .badge-both {
            background-color: var(--both-color);
            color: white;
        }
        
        .badge-Disabled {
            background-color: var(--Disabled-color);
            color: white;
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
        
        .filter-button.Enabled {
            box-shadow: 0 0 0 2px rgba(128,128,128,0.2);
        }
        
        .filter-button-direct {
            background-color: var(--direct-color);
            color: white;
        }
        
        .filter-button-inherited {
            background-color: var(--inherited-color);
            color: white;
        }
        
        .filter-button-both {
            background-color: var(--both-color);
            color: white;
        }
        
        .filter-button-Disabled {
            background-color: var(--Disabled-color);
            color: white;
        }
        
        .filter-button-all {
            background-color: var(--btn-outline-color);
            color: white;
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
        
        .license-filter-container {
            margin: 15px 0;
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
        }
        
        .license-badge {
            padding: 6px 12px;
            border-radius: 20px;
            background-color: var(--filter-tag-bg);
            color: var(--filter-tag-color);
            cursor: pointer;
            transition: all 0.2s;
            font-size: 0.85rem;
        }
        
        .license-badge.Enabled {
            background-color: var(--primary-color);
            color: white;
        }
        
        @media (max-width: 768px) {
            .dashboard-header {
                flex-direction: column;
                align-items: flex-start;
                gap: 10px;
            }
            
            .stats-card {
                min-height: 140px;
            }
            
            .filter-buttons {
                flex-direction: column;
            }
        }
        
        /* Export buttons styling */
        div.dt-buttons {
            margin-bottom: 1rem;
        }
        
        button.dt-button {
            background-color: var(--button-bg) !important;
            border-color: var(--button-border) !important;
            color: var(--button-color) !important;
            font-weight: 500 !important;
            padding: 8px 16px !important;
            border-radius: 4px !important;
            margin-right: 8px !important;
        }
        
        button.dt-button:hover {
            background-color: var(--button-hover-bg) !important;
            border-color: var(--border-color) !important;
        }
        
        button.dt-button.Enabled {
            background-color: var(--primary-color) !important;
            border-color: var(--primary-color) !important;
            color: white !important;
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
            color: var(--Disabled-color);
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
        
        .Enabled-filters-container {
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
                <h1>$Organization M365 Licensing Report</h1>
            </div>
            <div class="report-date">
                <i class="fas fa-calendar-alt me-2"></i> Report generated on: $ReportDate
            </div>
        </div>
        
        <div class="row mb-4">
            <div class="col-md-3 mb-3">
                <div class="stats-card direct-bg" id="directFilter">
                    <i class="fas fa-user-tag"></i>
                    <h3>Direct Licenses</h3>
                    <div class="number">$directLicenses</div>
                </div>
            </div>
            <div class="col-md-3 mb-3">
                <div class="stats-card inherited-bg" id="inheritedFilter">
                    <i class="fas fa-users-cog"></i>
                    <h3>Inherited Licenses</h3>
                    <div class="number">$inheritedLicenses</div>
                </div>
            </div>
            <div class="col-md-3 mb-3">
                <div class="stats-card both-bg" id="bothFilter">
                    <i class="fas fa-user-shield"></i>
                    <h3>Both (Direct + Inherited)</h3>
                    <div class="number">$bothLicenses</div>
                </div>
            </div>
            <div class="col-md-3 mb-3">
                <div class="stats-card Disabled-bg" id="DisabledFilter">
                    <i class="fas fa-user-slash"></i>
                    <h3>Disabled Users with Licenses</h3>
                    <div class="number">$DisabledUsersWithLicenses</div>
                </div>
            </div>
        </div>
        
        <div class="filter-section">
            <h5><i class="fas fa-filter me-2"></i>Filter Options</h5>
            
            <div class="row">
                <div class="col-md-6">
                    <div class="mb-3">
                        <label for="accountStatusFilter" class="form-label">Account Status</label>
                        <select id="accountStatusFilter" class="form-select">
                            <option value="">All Accounts</option>
                            <option value="Enabled">Enabled Accounts</option>
                            <option value="Disabled">Disabled Accounts</option>
                        </select>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="mb-3">
                        <label for="assignmentTypeFilter" class="form-label">Assignment Type</label>
                        <select id="assignmentTypeFilter" class="form-select">
                            <option value="">All Types</option>
                            <option value="Direct">Direct</option>
                            <option value="Inherited">Inherited</option>
                            <option value="Both">Both</option>
                        </select>
                    </div>
                </div>
            </div>
            
            <div class="mb-3">
                <label for="licenseNameFilter" class="form-label">License Name</label>
                <input type="text" id="licenseNameFilter" class="form-control" placeholder="Search for license names...">
            </div>
            
            <div class="mb-3">
                <label class="form-label">Quick Filters</label>
                <div class="filter-buttons">
                    <button class="filter-button filter-button-all" data-filter="all"><i class="fas fa-globe"></i> Show All</button>
                    <button class="filter-button filter-button-direct" data-filter="direct"><i class="fas fa-user-tag"></i> Direct Only</button>
                    <button class="filter-button filter-button-inherited" data-filter="inherited"><i class="fas fa-users-cog"></i> Inherited Only</button>
                    <button class="filter-button filter-button-both" data-filter="both"><i class="fas fa-user-shield"></i> Both</button>
                    <button class="filter-button filter-button-Disabled" data-filter="Disabled"><i class="fas fa-user-slash"></i> Disabled Users</button>
                </div>
            </div>
            
            <div class="Enabled-filters-container">
                <div class="d-flex justify-content-between align-items-center">
                    <label class="form-label mb-0">Enabled Filters:</label>
                    <button id="clearAllFilters" class="btn btn-sm btn-outline-secondary">Clear All</button>
                </div>
                <div class="filter-tags" id="EnabledFilters">
                    <!-- Enabled filters will be displayed here -->
                </div>
            </div>
        </div>
        
        <div class="card">
            <div class="card-header">
                <div>
                    <i class="fas fa-id-card"></i> License Assignment
                </div>
                <div class="show-all-container">
                    <label class="toggle-switch">
                        <input type="checkbox" id="licensesShowAllToggle">
                        <span class="toggle-slider"></span>
                    </label>
                    <p class="show-all-text">Show all entries</p>
                </div>
            </div>
            <div class="card-body">
                <div class="table-responsive">
                    <table id="licensesTable" class="table table-striped table-bordered" style="width:100%">
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
                            {{TABLE_DATA}}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <div class="card">
            <div class="card-header">
                <div>
                    <i class="fas fa-project-diagram"></i> Subscription Overview
                </div>
                <div class="show-all-container">
                    <label class="toggle-switch">
                        <input type="checkbox" id="subscriptionShowAllToggle">
                        <span class="toggle-slider"></span>
                    </label>
                    <p class="show-all-text">Show all entries</p>
                </div>
            </div>
            <div class="card-body">
                <div class="table-responsive">
                    <table id="subscriptionTable" class="table table-striped table-bordered" style="width:100%">
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
                            {{SUBSCRIPTION_DATA}}
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
            
            // Initialize DataTable for licenses
            const licensesTable = $('#licensesTable').DataTable({
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
                order: [[0, 'asc']],
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
                    // Enforce the correct colors after DataTables redraws
                    updateTableColors();
                }
            });
            
            // Initialize DataTable for subscriptions
            const subscriptionTable = $('#subscriptionTable').DataTable({
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
                lengthMenu: [[10, 25, 50, -1], [10, 25, 50, "All"]],
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
                    // Enforce the correct colors after DataTables redraws
                    updateTableColors();
                }
            });
            
            // Apply initial table colors
            setTimeout(updateTableColors, 100);
            
            // Show all toggle functionality for licenses table
            $('#licensesShowAllToggle').on('change', function() {
                if ($(this).is(':checked')) {
                    licensesTable.page.len(-1).draw();
                } else {
                    licensesTable.page.len(10).draw();
                }
            });
            
            // Show all toggle functionality for subscription table
            $('#subscriptionShowAllToggle').on('change', function() {
                if ($(this).is(':checked')) {
                    subscriptionTable.page.len(-1).draw();
                } else {
                    subscriptionTable.page.len(10).draw();
                }
            });
            
            // Custom filtering function
            $.fn.dataTable.ext.search.push(
                function(settings, data, dataIndex) {
                    // Only apply to licenses table
                    if (settings.nTable.id !== 'licensesTable') {
                        return true;
                    }
                    
                    // Get filter values
                    const accountStatus = $('#accountStatusFilter').val();
                    const assignmentType = $('#assignmentTypeFilter').val();
                    const licenseNameFilter = $('#licenseNameFilter').val().toLowerCase();
                    
                    // Get row data - Updated column indices after adding "Last Successful Sign In"
                    const rowAccountStatus = data[2]; // Account Status column
                    const rowAssignmentType = data[5]; // Assignment Type column (was 4, now 5)
                    const rowLicenseName = data[4].toLowerCase(); // License Name column (was 3, now 4)
                    
                    // Filter by account status
                    if (accountStatus && accountStatus === 'Enabled' && !rowAccountStatus.includes('Enabled')) {
                        return false;
                    }
                    if (accountStatus && accountStatus === 'Disabled' && !rowAccountStatus.includes('Disabled')) {
                        return false;
                    }
                    
                    // Filter by assignment type
                    if (assignmentType && !rowAssignmentType.includes(assignmentType)) {
                        return false;
                    }
                    
                    // Filter by license name
                    if (licenseNameFilter && !rowLicenseName.includes(licenseNameFilter)) {
                        return false;
                    }
                    
                    return true;
                }
            );
            
            // Stats card filtering
            $('#directFilter').on('click', function() {
                $('#assignmentTypeFilter').val('Direct');
                updateEnabledFilters('Assignment Type', 'Direct');
                applyFilters();
                toggleStatsCardEnabled('directFilter');
            });
            
            $('#inheritedFilter').on('click', function() {
                $('#assignmentTypeFilter').val('Inherited');
                updateEnabledFilters('Assignment Type', 'Inherited');
                applyFilters();
                toggleStatsCardEnabled('inheritedFilter');
            });
            
            $('#bothFilter').on('click', function() {
                $('#assignmentTypeFilter').val('Both');
                updateEnabledFilters('Assignment Type', 'Both');
                applyFilters();
                toggleStatsCardEnabled('bothFilter');
            });
            
            $('#DisabledFilter').on('click', function() {
                $('#accountStatusFilter').val('Disabled');
                updateEnabledFilters('Account Status', 'Disabled');
                applyFilters();
                toggleStatsCardEnabled('DisabledFilter');
            });
            
            // Button filtering
            $('.filter-button').on('click', function() {
                const filterType = $(this).data('filter');
                
                // Clear all filter button Enabled states
                $('.filter-button').removeClass('Enabled');
                $(this).addClass('Enabled');
                
                // Reset filters
                $('#accountStatusFilter').val('');
                $('#assignmentTypeFilter').val('');
                $('#licenseNameFilter').val('');
                clearEnabledFilters();
                
                // Apply selected filter
                switch(filterType) {
                    case 'direct':
                        $('#assignmentTypeFilter').val('Direct');
                        updateEnabledFilters('Assignment Type', 'Direct');
                        toggleStatsCardEnabled('directFilter');
                        break;
                    case 'inherited':
                        $('#assignmentTypeFilter').val('Inherited');
                        updateEnabledFilters('Assignment Type', 'Inherited');
                        toggleStatsCardEnabled('inheritedFilter');
                        break;
                    case 'both':
                        $('#assignmentTypeFilter').val('Both');
                        updateEnabledFilters('Assignment Type', 'Both');
                        toggleStatsCardEnabled('bothFilter');
                        break;
                    case 'Disabled':
                        $('#accountStatusFilter').val('Disabled');
                        updateEnabledFilters('Account Status', 'Disabled');
                        toggleStatsCardEnabled('DisabledFilter');
                        break;
                    case 'all':
                    default:
                        // Reset all filters
                        $('.stats-card').removeClass('Enabled');
                        break;
                }
                
                applyFilters();
            });
            
            // Apply filters when select boxes change
            $('#accountStatusFilter, #assignmentTypeFilter').on('change', function() {
                const filterType = $(this).attr('id');
                const filterValue = $(this).val();
                
                if (filterValue) {
                    if (filterType === 'accountStatusFilter') {
                        updateEnabledFilters('Account Status', filterValue);
                    } else if (filterType === 'assignmentTypeFilter') {
                        updateEnabledFilters('Assignment Type', filterValue);
                    }
                } else {
                    if (filterType === 'accountStatusFilter') {
                        removeEnabledFilter('Account Status');
                    } else if (filterType === 'assignmentTypeFilter') {
                        removeEnabledFilter('Assignment Type');
                    }
                }
                
                applyFilters();
            });
            
            // Apply filter when license name input changes
            $('#licenseNameFilter').on('input', function() {
                const filterValue = $(this).val();
                
                if (filterValue) {
                    updateEnabledFilters('License Name', filterValue);
                } else {
                    removeEnabledFilter('License Name');
                }
                
                applyFilters();
            });
            
            // Clear all filters button
            $('#clearAllFilters').on('click', function() {
                $('#accountStatusFilter').val('');
                $('#assignmentTypeFilter').val('');
                $('#licenseNameFilter').val('');
                $('.filter-button').removeClass('Enabled');
                $('.stats-card').removeClass('Enabled');
                clearEnabledFilters();
                applyFilters();
            });
            
            // Function to apply all filters
            function applyFilters() {
                licensesTable.draw();
            }
            
            // Function to toggle stats card Enabled state
            function toggleStatsCardEnabled(cardId) {
                $('.stats-card').removeClass('Enabled');
                $('#' + cardId).addClass('Enabled');
            }
            
            // Function to update Enabled filters
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
                
                $('#EnabledFilters').append(filterTag);
                
                // Add click handler to remove filter
                $('.remove-filter').off('click').on('click', function() {
                    const filterTypeToRemove = $(this).data('filter-type');
                    
                    if (filterTypeToRemove === 'Account Status') {
                        $('#accountStatusFilter').val('');
                    } else if (filterTypeToRemove === 'Assignment Type') {
                        $('#assignmentTypeFilter').val('');
                    } else if (filterTypeToRemove === 'License Name') {
                        $('#licenseNameFilter').val('');
                    }
                    
                    $(this).closest('.filter-tag').remove();
                    
                    // Remove Enabled state from stat cards and filter buttons
                    if (filterTypeToRemove === 'Account Status' && $('#DisabledFilter').hasClass('Enabled')) {
                        $('#DisabledFilter').removeClass('Enabled');
                    } else if (filterTypeToRemove === 'Assignment Type') {
                        $('#directFilter, #inheritedFilter, #bothFilter').removeClass('Enabled');
                    }
                    
                    $('.filter-button').removeClass('Enabled');
                    
                    applyFilters();
                });
            }
            
            // Function to remove Enabled filter by type
            function removeEnabledFilter(filterType) {
                $('.filter-tag[data-filter-type="' + filterType + '"]').remove();
            }
            
            // Function to clear all Enabled filters
            function clearEnabledFilters() {
                $('#EnabledFilters').empty();
            }
            
            // Force dark mode to take effect on page elements
            $(window).on('load', function() {
                setTimeout(updateTableColors, 200);
            });
            
            // Re-apply styles after DataTables operations
            licensesTable.on('draw.dt', function() {
                setTimeout(updateTableColors, 50);
            });
            
            subscriptionTable.on('draw.dt', function() {
                setTimeout(updateTableColors, 50);
            });
        });
    </script>
</body>
</html>
'@

    # Generate table rows for user licenses
    $tableRows = ""
    foreach ($item in $Report) {
        $accountStatusClass = if ($item.AccountEnabled -eq "No") { 'class="table-danger"' } else { '' }
        
        $assignmentTypeBadge = switch ($item.AssignmentType) {
            "Direct" { '<span class="badge badge-direct">Direct</span>' }
            "Inherited" { '<span class="badge badge-inherited">Inherited</span>' }
            "Both" { '<span class="badge badge-both">Both</span>' }
            default { '<span class="badge bg-secondary">Unknown</span>' }
        }
        
        $accountStatus = if ($item.AccountEnabled -eq "Yes") {
            '<span class="badge bg-success">Enabled</span>'
        }
        else {
            '<span class="badge badge-Disabled">Disabled</span>'
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
            '<span class="badge bg-danger">' + $item.AvailableLicenses + ' (' + $availabilityPercentage + '%)</span>'
        }
        elseif ($availabilityPercentage -lt 20) {
            '<span class="badge bg-warning text-dark">' + $item.AvailableLicenses + ' (' + $availabilityPercentage + '%)</span>'
        }
        else {
            '<span class="badge bg-success">' + $item.AvailableLicenses + ' (' + $availabilityPercentage + '%)</span>'
        }
        
        $licenseStatusBadge = if ($item.LicenseStatus -eq "Enabled") {
            '<span class="badge bg-success">Enabled</span>'
        }
        else {
            '<span class="badge bg-danger">Disabled</span>'
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

    # Get current date for report
    $currentDate = Get-Date -Format "dd-MM-yyyy HH:mm"

    # Replace placeholders in template with actual values
    $htmlContent = $htmlTemplate.Replace('$Organization', $Organization)
    $htmlContent = $htmlContent.Replace('$ReportDate', $currentDate)
    $htmlContent = $htmlContent.Replace('$directLicenses', $directLicenses)
    $htmlContent = $htmlContent.Replace('$inheritedLicenses', $inheritedLicenses)
    $htmlContent = $htmlContent.Replace('$bothLicenses', $bothLicenses)
    $htmlContent = $htmlContent.Replace('$DisabledUsersWithLicenses', $DisabledUsersWithLicenses)
    $htmlContent = $htmlContent.Replace('{{TABLE_DATA}}', $tableRows)
    $htmlContent = $htmlContent.Replace('{{SUBSCRIPTION_DATA}}', $subscriptionRows)

    # Add additional CSS for dark mode pagination
    $darkModePaginationCss = @"
    <style>
        /* Dark mode pagination buttons */
        [data-theme="dark"] .page-link {
            background-color: var(--button-bg) !important;
            color: var(--text-color) !important;
            border-color: var(--border-color) !important;
        }
        
        [data-theme="dark"] .page-item.Enabled .page-link {
            background-color: var(--primary-color) !important;
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
    </style>
"@
    
    # Insert the dark mode pagination CSS before the </head> tag
    $htmlContent = $htmlContent.Replace('</head>', "$darkModePaginationCss`n</head>")

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
            $emailSent =         # CODE

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
            Write-Host "INFO: Email sending is disabled. Set -SendEmail to $true to enable." -ForegroundColor Yellow
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
            Write-Host "INFO: Email sending is disabled. Set -SendEmail to $true to enable." -ForegroundColor Yellow
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
            Write-Host "INFO: Email sending is disabled. Set -SendEmail to $true to enable." -ForegroundColor Yellow
        }

        # Clean up the report file
        if ($SendEmail -and $emailSent) {
            if (Test-Path -Path $script:ExportPath) {
                Remove-Item -Path $script:ExportPath -Force
                Write-Host "INFO: Temporary report file deleted." -ForegroundColor Green
            } else {
                Write-Host "INFO: No temporary report file found to delete." -ForegroundColor Yellow
            }
        }subject -BodyHtml         # CODE

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
            Write-Host "INFO: Email sending is disabled. Set -SendEmail to $true to enable." -ForegroundColor Yellow
        }

        # Clean up the report file
        if ($SendEmail -and $emailSent) {
            if (Test-Path -Path $script:ExportPath) {
                Remove-Item -Path $script:ExportPath -Force
                Write-Host "INFO: Temporary report file deleted." -ForegroundColor Green
            } else {
                Write-Host "INFO: No temporary report file found to delete." -ForegroundColor Yellow
            }
        }body

            if ($emailSent) {
                Write-Host "INFO: Email sent successfully." -ForegroundColor Green
            }
            else {
                Write-Host "ERROR: Failed to send email." -ForegroundColor Red
            }
        }
        else {
            Write-Host "INFO: Email sending is disabled. Set -SendEmail to $true to enable." -ForegroundColor Yellow
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
