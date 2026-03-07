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
        [array]$Report_OperationSystemEdtionOverview,
        [Parameter(Mandatory = $false)]
        [array]$Report_NoncompliantDevices,
        [Parameter(Mandatory = $false)]
        [array]$Report_DisabledPrimaryUsers,
        [Parameter(Mandatory = $false)]
        [string]$ExportPath
    )

    # Default ExportPath to current folder if not provided
    if (-not $ExportPath) {
        $ExportPath = Join-Path (Get-Location).Path "$TenantName-IntuneAnomaliesReport.html"
    }

    # Calculate counts for dashboard statistics
    $Report_ApplicationFailureReport_Count = $Report_ApplicationFailureReport | Measure-Object | Select-Object -ExpandProperty Count
    $Report_DevicesWithMultipleUsers_Count = $Report_DevicesWithMultipleUsers | Measure-Object | Select-Object -ExpandProperty Count
    $Report_NotEncryptedDevices_Count = $Report_NotEncryptedDevices | Measure-Object | Select-Object -ExpandProperty Count
    $Report_DevicesWithoutAutopilotHash_Count = $Report_DevicesWithoutAutopilotHash | Measure-Object | Select-Object -ExpandProperty Count
    $Report_InactiveDevices_Count = $Report_InactiveDevices | Measure-Object | Select-Object -ExpandProperty Count
    $Report_NoncompliantDevices_Count = $NoncompliantDevicesRaw | Measure-Object | Select-Object -ExpandProperty Count
    $Report_OperationSystemEdtionOverview_Count = $Report_OperationSystemEdtionOverview | Measure-Object | Select-Object -ExpandProperty Count
    $Report_DisabledPrimaryUsers_Count = $Report_DisabledPrimaryUsers | Measure-Object | Select-Object -ExpandProperty Count
    
    # Get the current date and time for the report header
    $CurrentDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

    # Create HTML Template with DataTables
    $htmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>$TenantName Intune Anomalies Report</title>
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
    --datatable-odd-row-bg: #f9f9f9;
    --tab-active-bg: #0078d4;
    --tab-active-color: #fff;
}
 
[data-theme="dark"] {
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
    padding: 8px 12px;
    border-radius: 6px;
    color: white;
    position: relative;
    overflow: hidden;
    min-height: 70px;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    cursor: pointer;
    transition: all 0.3s;
}
 
.stats-card::before {
    content: '';
    position: absolute;
    top: -8px;
    right: -8px;
    width: 40px;
    height: 40px;
    border-radius: 50%;
    background-color: rgba(255,255,255,0.1);
    z-index: 0;
}
 
.stats-card i {
    font-size: 1.2rem;
    margin-bottom: 4px;
    position: relative;
    z-index: 1;
}
 
.stats-card h3 {
    font-size: 0.65rem;
    font-weight: 500;
    margin-bottom: 4px;
    position: relative;
    z-index: 1;
    line-height: 1.0;
    height: 24px;
    display: flex;
    align-items: center;
    justify-content: center;
}
 
.stats-card .number {
    font-size: 1.2rem;
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
 
.badge-active {
    background-color: var(--service-principal-color);
    color: white;
}
 
.badge-group {
    background-color: var(--group-color);
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
 
.filter-container {
    background-color: var(--card-bg);
    border: 1px solid var(--border-color);
    border-radius: 8px;
    padding: 20px;
    margin-bottom: 20px;
}
 
.filter-row {
    display: flex;
    flex-wrap: wrap;
    gap: 15px;
    align-items: end;
}
 
.filter-group {
    flex: 1;
    min-width: 200px;
}
 
.filter-group label {
    display: block;
    margin-bottom: 5px;
    font-weight: 500;
    color: var(--text-color);
}
 
.filter-buttons {
    display: flex;
    gap: 10px;
    align-items: end;
}
 
.btn-filter {
    padding: 8px 16px;
    border: none;
    border-radius: 4px;
    font-weight: 500;
    cursor: pointer;
    transition: all 0.2s;
}
 
.btn-primary {
    background-color: var(--primary-color);
    color: white;
}
 
.btn-primary:hover {
    background-color: var(--secondary-color);
}
 
.btn-secondary {
    background-color: var(--button-bg);
    color: var(--button-color);
    border: 1px solid var(--button-border);
}
 
.btn-secondary:hover {
    background-color: var(--button-hover-bg);
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
 
.form-select option {
    background-color: var(--input-bg);
    color: var(--input-color);
}
 
table.dataTable.stripe tbody tr.odd,
table.dataTable.display tbody tr.odd {
    background-color: var(--datatable-odd-row-bg) !important;
}
 
table.dataTable.stripe tbody tr.even,
table.dataTable.display tbody tr.even {
    background-color: var(--datatable-even-row-bg) !important;
}
 
table.dataTable.hover tbody tr:hover,
table.dataTable.display tbody tr:hover {
    background-color: var(--table-hover-bg) !important;
}
 
table.dataTable.border-bottom,
table.dataTable.border-top,
table.dataTable thead th,
table.dataTable tfoot th,
table.dataTable thead td,
table.dataTable tfoot td {
    border-color: var(--table-border-color) !important;
}
 
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
 
.table tbody tr td {
    background-color: transparent !important;
}
 
.table-striped>tbody>tr:nth-of-type(odd) {
    --bs-table-accent-bg: var(--datatable-odd-row-bg) !important;
}
 
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
                <h1>$TenantName Intune Anomalies Report</h1>
            </div>
            <div class="report-date">
                <i class="fas fa-calendar-alt me-2"></i> Report generated on: $ReportDate
            </div>
        </div>
         
        <div class="row mb-4">
            <div class="col-md-4 mb-3">
                <div class="stats-card permanent-bg">
                    <i class="fas fa-exclamation-triangle"></i>
                    <h3>Application Failures</h3>
                    <div class="number">$applicationFailures</div>
                </div>
            </div>
            <div class="col-md-4 mb-3">
                <div class="stats-card eligible-bg">
                    <i class="fas fa-users"></i>
                    <h3>Multiple users on non-shared devices</h3>
                    <div class="number">$multipleUsers</div>
                </div>
            </div>
            <div class="col-md-4 mb-3">
                <div class="stats-card group-bg">
                    <i class="fas fa-shield-alt"></i>
                    <h3>Not Encrypted Devices</h3>
                    <div class="number">$notEncrypted</div>
                </div>
            </div>
            <div class="col-md-4 mb-3">
                <div class="stats-card service-principal-bg">
                    <i class="fas fa-unlink"></i>
                    <h3>No Autopilot Hash</h3>
                    <div class="number">$noAutopilot</div>
                </div>
            </div>
            <div class="col-md-4 mb-3">
                <div class="stats-card disabled-bg">
                    <i class="fas fa-clock"></i>
                    <h3>Inactive Devices</h3>
                    <div class="number">$inactiveDevices</div>
                </div>
            </div>
            <div class="col-md-4 mb-3">
                <div class="stats-card enabled-bg">
                    <i class="fas fa-exclamation-circle"></i>
                    <h3>Noncompliant Devices</h3>
                    <div class="number">$noncompliantDevices</div>
                </div>
            </div>
            <div class="col-md-4 mb-3">
                <div class="stats-card" style="background: linear-gradient(135deg, #ff6b35, #f7931e);">
                    <i class="fas fa-desktop"></i>
                    <h3>OS Edition Overview</h3>
                    <div class="number">$osEditionOverview</div>
                </div>
            </div>
            <div class="col-md-4 mb-3">
                <div class="stats-card" style="background: linear-gradient(135deg, #dc3545, #c82333);">
                    <i class="fas fa-user-slash"></i>
                    <h3>Disabled Primary Users</h3>
                    <div class="number">$disabledPrimaryUsers</div>
                </div>
            </div>
        </div>
 
        <div class="report-tabs">
            <div class="report-tab active" data-panel="application-failures">Application Failures</div>
            <div class="report-tab" data-panel="multiple-users">Multiple Users</div>
            <div class="report-tab" data-panel="not-encrypted">Not Encrypted</div>
            <div class="report-tab" data-panel="no-autopilot">No Autopilot Hash</div>
            <div class="report-tab" data-panel="inactive-devices">Inactive Devices</div>
            <div class="report-tab" data-panel="noncompliant-devices">Noncompliant</div>
            <div class="report-tab" data-panel="os-edition-overview">OS Edition Overview</div>
            <div class="report-tab" data-panel="disabled-primary-users">Disabled Primary Users</div>
             
        </div>
 
        <div id="application-failures" class="report-panel active">
            <div class="alert alert-info mb-3" role="alert">
                <i class="fas fa-info-circle me-2"></i>
                <strong>Application Failures:</strong> Shows devices with failed application installations. If the failure count is high it might incidate a wrong packaged application.
            </div>
             
            <div class="filter-container">
                <div class="filter-row">
                    <div class="filter-group">
                        <label for="appFailuresCustomerFilter">Customer</label>
                        <select id="appFailuresCustomerFilter" class="form-select">
                            <option value="">All Customers</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="appFailuresAppFilter">Application</label>
                        <select id="appFailuresAppFilter" class="form-select">
                            <option value="">All Applications</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="appFailuresPlatformFilter">Platform</label>
                        <select id="appFailuresPlatformFilter" class="form-select">
                            <option value="">All Platforms</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="appFailuresVersionFilter">Version</label>
                        <select id="appFailuresVersionFilter" class="form-select">
                            <option value="">All Versions</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="appFailuresPercentageFilter">Failed Device Percentage</label>
                        <select id="appFailuresPercentageFilter" class="form-select">
                            <option value="">All Percentages</option>
                            <option value="0-20">0-20%</option>
                            <option value="20-40">20-40%</option>
                            <option value="40-60">40-60%</option>
                            <option value="60-80">60-80%</option>
                            <option value="80-100">80-100%</option>
                        </select>
                    </div>
                    <div class="filter-buttons">
                        <button class="btn-filter btn-secondary" onclick="clearAppFailuresFilters()">Clear</button>
                    </div>
                </div>
            </div>
             
            <div class="card">
                <div class="card-header">
                    <div>
                        <i class="fas fa-exclamation-triangle"></i> Application Failures
                    </div>
                    <div class="show-all-container">
                        <label class="toggle-switch">
                            <input type="checkbox" id="appFailuresShowAllToggle">
                            <span class="toggle-slider"></span>
                        </label>
                        <p class="show-all-text">Show all entries</p>
                    </div>
                </div>
                <div class="card-body">
                    <div class="table-responsive">
                        <table id="appFailuresTable" class="table table-striped table-bordered" style="width:100%">
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
                                {{APPLICATION_FAILURES_DATA}}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
 
        <div id="multiple-users" class="report-panel">
            <div class="alert alert-info mb-3" role="alert">
                <i class="fas fa-exclamation-triangle me-2"></i>
                <strong>Multiple Users on user-driven device:</strong> Lists user-driven devices that have multiple users logged on. These devices should typically be re-enrolled as shared devices to ensure proper configuration and security.
            </div>
 
        <div class="filter-container">
            <div class="filter-row">
                <div class="filter-group">
                    <label for="multipleUsersCustomerFilter">Customer</label>
                    <select id="multipleUsersCustomerFilter" class="form-select">
                        <option value="">All Customers</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="multipleUsersDeviceFilter">Device Name</label>
                    <select id="multipleUsersDeviceFilter" class="form-select">
                        <option value="">All Devices</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="multipleUsersPrimaryUserFilter">Primary User</label>
                    <select id="multipleUsersPrimaryUserFilter" class="form-select">
                        <option value="">All Users</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="multipleUsersProfileFilter">Enrollment Profile</label>
                    <select id="multipleUsersProfileFilter" class="form-select">
                        <option value="">All Profiles</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="multipleUsersCountFilter">User Count</label>
                    <select id="multipleUsersCountFilter" class="form-select">
                        <option value="">All Counts</option>
                        <option value="2">2 Users</option>
                        <option value="3">3 Users</option>
                        <option value="4+">4+ Users</option>
                    </select>
                </div>
                <div class="filter-buttons">
                    <button class="btn-filter btn-secondary" onclick="clearMultipleUsersFilters()">Clear</button>
                </div>
            </div>
        </div>
 
            <div class="card">
                <div class="card-header">
                    <div>
                        <i class="fas fa-users"></i> Devices with Multiple Users
                    </div>
                    <div class="show-all-container">
                        <label class="toggle-switch">
                            <input type="checkbox" id="multipleUsersShowAllToggle">
                            <span class="toggle-slider"></span>
                        </label>
                        <p class="show-all-text">Show all entries</p>
                    </div>
                </div>
                <div class="card-body">
                    <div class="table-responsive">
                        <table id="multipleUsersTable" class="table table-striped table-bordered" style="width:100%">
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
                                {{MULTIPLE_USERS_DATA}}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
 
        <div id="not-encrypted" class="report-panel">
            <div class="alert alert-info mb-3" role="alert">
                <i class="fas fa-shield-alt me-2"></i>
                <strong>Not Encrypted Devices:</strong> Displays devices that are not Bitlocker encrypted. This represents a significant security risk as data on these devices could be accessed if the device is lost or stolen.
            </div>
 
        <div class="filter-container">
            <div class="filter-row">
                <div class="filter-group">
                    <label for="notEncryptedCustomerFilter">Customer</label>
                    <select id="notEncryptedCustomerFilter" class="form-select">
                        <option value="">All Customers</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="notEncryptedDeviceFilter">Device Name</label>
                    <select id="notEncryptedDeviceFilter" class="form-select">
                        <option value="">All Devices</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="notEncryptedUserFilter">Primary User</label>
                    <select id="notEncryptedUserFilter" class="form-select">
                        <option value="">All Users</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="notEncryptedManufacturerFilter">Manufacturer</label>
                    <select id="notEncryptedManufacturerFilter" class="form-select">
                        <option value="">All Manufacturers</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="notEncryptedModelFilter">Model</label>
                    <select id="notEncryptedModelFilter" class="form-select">
                        <option value="">All Models</option>
                    </select>
                </div>
                <div class="filter-buttons">
                    <button class="btn-filter btn-secondary" onclick="clearNotEncryptedFilters()">Clear</button>
                </div>
            </div>
        </div>
 
            <div class="card">
                <div class="card-header">
                    <div>
                        <i class="fas fa-shield-alt"></i> Not Encrypted Devices
                    </div>
                    <div class="show-all-container">
                        <label class="toggle-switch">
                            <input type="checkbox" id="notEncryptedShowAllToggle">
                            <span class="toggle-slider"></span>
                        </label>
                        <p class="show-all-text">Show all entries</p>
                    </div>
                </div>
                <div class="card-body">
                    <div class="table-responsive">
                        <table id="notEncryptedTable" class="table table-striped table-bordered" style="width:100%">
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
                                {{NOT_ENCRYPTED_DATA}}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
 
        <div id="no-autopilot" class="report-panel">
            <div class="alert alert-info mb-3" role="alert">
                <i class="fas fa-plane me-2"></i>
                <strong>Non-company owned devices:</strong> Shows devices that are missing their hardware hash in Autopilot. These devices can get lost or stolen, and without the hash, they cannot be managed by Autopilot. This can lead to security risks and management challenges.
            </div>
 
        <div class="filter-container">
            <div class="filter-row">
                <div class="filter-group">
                    <label for="noAutopilotCustomerFilter">Customer</label>
                    <select id="noAutopilotCustomerFilter" class="form-select">
                        <option value="">All Customers</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="noAutopilotDeviceFilter">Device Name</label>
                    <select id="noAutopilotDeviceFilter" class="form-select">
                        <option value="">All Devices</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="noAutopilotUserFilter">Primary User</label>
                    <select id="noAutopilotUserFilter" class="form-select">
                        <option value="">All Users</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="noAutopilotManufacturerFilter">Manufacturer</label>
                    <select id="noAutopilotManufacturerFilter" class="form-select">
                        <option value="">All Manufacturers</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="noAutopilotModelFilter">Model</label>
                    <select id="noAutopilotModelFilter" class="form-select">
                        <option value="">All Models</option>
                    </select>
                </div>
                <div class="filter-buttons">
                    <button class="btn-filter btn-secondary" onclick="clearNoAutopilotFilters()">Clear</button>
                </div>
            </div>
        </div>
 
        <div class="card">
                <div class="card-header">
                    <div>
                        <i class="fas fa-plane"></i> Non-company owned devices
                    </div>
                    <div class="show-all-container">
                        <label class="toggle-switch">
                            <input type="checkbox" id="noAutopilotShowAllToggle">
                            <span class="toggle-slider"></span>
                        </label>
                        <p class="show-all-text">Show all entries</p>
                    </div>
                </div>
                <div class="card-body">
                    <div class="table-responsive">
                        <table id="noAutopilotTable" class="table table-striped table-bordered" style="width:100%">
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
                                {{NO_AUTOPILOT_DATA}}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
 
        <div id="inactive-devices" class="report-panel">
            <div class="alert alert-info mb-3" role="alert">
                <i class="fas fa-clock me-2"></i>
                <strong>Inactive Devices (90+ days):</strong> Lists devices that haven't contacted Intune in 90 or more days. These devices may be decommissioned, lost, or experiencing connectivity issues and should be reviewed for cleanup.
            </div>
 
        <div class="filter-container">
            <div class="filter-row">
                <div class="filter-group">
                    <label for="inactiveCustomerFilter">Customer</label>
                    <select id="inactiveCustomerFilter" class="form-select">
                        <option value="">All Customers</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="inactiveDeviceFilter">Device Name</label>
                    <select id="inactiveDeviceFilter" class="form-select">
                        <option value="">All Devices</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="inactiveUserFilter">Primary User</label>
                    <select id="inactiveUserFilter" class="form-select">
                        <option value="">All Users</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="inactiveManufacturerFilter">Manufacturer</label>
                    <select id="inactiveManufacturerFilter" class="form-select">
                        <option value="">All Manufacturers</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="inactiveModelFilter">Model</label>
                    <select id="inactiveModelFilter" class="form-select">
                        <option value="">All Models</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="inactiveInactivityFilter">Inactivity Period</label>
                    <select id="inactiveInactivityFilter" class="form-select">
                        <option value="">All Periods</option>
                        <option value="90-180">90-180 days</option>
                        <option value="180+">180+ days</option>
                    </select>
                </div>
                <div class="filter-buttons">
                    <button class="btn-filter btn-secondary" onclick="clearInactiveFilters()">Clear</button>
                </div>
            </div>
        </div>
 
            <div class="card">
                <div class="card-header">
                    <div>
                        <i class="fas fa-clock"></i> Inactive Devices (90+ days)
                    </div>
                    <div class="show-all-container">
                        <label class="toggle-switch">
                            <input type="checkbox" id="inactiveDevicesShowAllToggle">
                            <span class="toggle-slider"></span>
                        </label>
                        <p class="show-all-text">Show all entries</p>
                    </div>
                </div>
                <div class="card-body">
                    <div class="table-responsive">
                        <table id="inactiveDevicesTable" class="table table-striped table-bordered" style="width:100%">
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
                                {{INACTIVE_DEVICES_DATA}}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
 
        <div id="noncompliant-devices" class="report-panel">
            <div class="alert alert-info mb-3" role="alert">
                <i class="fas fa-exclamation-circle me-2"></i>
                <strong>Noncompliant Devices:</strong> Displays devices that don't meet your organization's compliance policies. These devices may have security vulnerabilities or configuration issues that need immediate attention to maintain security standards.
        </div>
 
        <div class="filter-container">
            <div class="filter-row">
                <div class="filter-group">
                    <label for="noncompliantCustomerFilter">Customer</label>
                    <select id="noncompliantCustomerFilter" class="form-select">
                        <option value="">All Customers</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="noncompliantDeviceFilter">Device Name</label>
                    <select id="noncompliantDeviceFilter" class="form-select">
                        <option value="">All Devices</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="noncompliantUserFilter">Primary User</label>
                    <select id="noncompliantUserFilter" class="form-select">
                        <option value="">All Users</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="noncompliantManufacturerFilter">Manufacturer</label>
                    <select id="noncompliantManufacturerFilter" class="form-select">
                        <option value="">All Manufacturers</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="noncompliantModelFilter">Model</label>
                    <select id="noncompliantModelFilter" class="form-select">
                        <option value="">All Models</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="noncompliantReasonFilter">Noncompliant Reason</label>
                    <select id="noncompliantReasonFilter" class="form-select">
                        <option value="">All Reasons</option>
                    </select>
                </div>
                <div class="filter-buttons">
                    <button class="btn-filter btn-secondary" onclick="clearNoncompliantFilters()">Clear</button>
                </div>
            </div>
        </div>
 
            <div class="card">
                <div class="card-header">
                    <div>
                        <i class="fas fa-exclamation-circle"></i> Noncompliant Devices
                    </div>
                    <div class="show-all-container">
                        <label class="toggle-switch">
                            <input type="checkbox" id="noncompliantShowAllToggle">
                            <span class="toggle-slider"></span>
                        </label>
                        <p class="show-all-text">Show all entries</p>
                    </div>
                </div>
                <div class="card-body">
                    <div class="table-responsive">
                        <table id="noncompliantTable" class="table table-striped table-bordered" style="width:100%">
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
                                {{NONCOMPLIANT_DEVICES_DATA}}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
 
        <div id="os-edition-overview" class="report-panel">
            <div class="alert alert-info mb-3" role="alert">
                <i class="fas fa-desktop me-2"></i>
                <strong>Operating System Edition Overview:</strong> Provides a comprehensive view of device operating system editions across your environment. This helps identify devices that may need OS edition upgrades (e.g., Pro to Enterprise) for enhanced management and security features.
            </div>
 
            <div class="filter-container">
                <div class="filter-row">
                    <div class="filter-group">
                        <label for="osEditionCustomerFilter">Customer</label>
                        <select id="osEditionCustomerFilter" class="form-select">
                            <option value="">All Customers</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="osEditionDeviceFilter">Device Name</label>
                        <select id="osEditionDeviceFilter" class="form-select">
                            <option value="">All Devices</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="osEditionUserFilter">Primary User</label>
                        <select id="osEditionUserFilter" class="form-select">
                            <option value="">All Users</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="osEditionEditionFilter">OS Edition</label>
                        <select id="osEditionEditionFilter" class="form-select">
                            <option value="">All Editions</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="osEditionFriendlyNameFilter">OS Friendly Name</label>
                        <select id="osEditionFriendlyNameFilter" class="form-select">
                            <option value="">All OS Versions</option>
                        </select>
                    </div>
                    <div class="filter-buttons">
                        <button class="btn-filter btn-secondary" onclick="clearOSEditionFilters()">Clear</button>
                    </div>
                </div>
            </div>
 
            <div class="card">
                <div class="card-header">
                    <div>
                        <i class="fas fa-desktop"></i> Operating System Edition Overview
                    </div>
                    <div class="show-all-container">
                        <label class="toggle-switch">
                            <input type="checkbox" id="osEditionShowAllToggle">
                            <span class="toggle-slider"></span>
                        </label>
                        <p class="show-all-text">Show all entries</p>
                    </div>
                </div>
                <div class="card-body">
                    <div class="table-responsive">
                        <table id="osEditionTable" class="table table-striped table-bordered" style="width:100%">
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
                                {{OS_EDITION_OVERVIEW_DATA}}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
 
        <div id="disabled-primary-users" class="report-panel">
            <div class="alert alert-info mb-3" role="alert">
                <i class="fas fa-user-slash me-2"></i>
                <strong>Devices with Disabled Primary Users:</strong> Shows devices where the primary user account has been disabled in Microsoft Entra ID. These devices may need to be reassigned to active users or cleaned up to maintain security and proper device management.
            </div>
 
            <div class="filter-container">
                <div class="filter-row">
                    <div class="filter-group">
                        <label for="disabledUsersCustomerFilter">Customer</label>
                        <select id="disabledUsersCustomerFilter" class="form-select">
                            <option value="">All Customers</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="disabledUsersDeviceFilter">Device Name</label>
                        <select id="disabledUsersDeviceFilter" class="form-select">
                            <option value="">All Devices</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="disabledUsersUserFilter">Primary User</label>
                        <select id="disabledUsersUserFilter" class="form-select">
                            <option value="">All Users</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="disabledUsersManufacturerFilter">Manufacturer</label>
                        <select id="disabledUsersManufacturerFilter" class="form-select">
                            <option value="">All Manufacturers</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="disabledUsersModelFilter">Model</label>
                        <select id="disabledUsersModelFilter" class="form-select">
                            <option value="">All Models</option>
                        </select>
                    </div>
                    <div class="filter-buttons">
                        <button class="btn-filter btn-secondary" onclick="clearDisabledUsersFilters()">Clear</button>
                    </div>
                </div>
            </div>
 
            <div class="card">
                <div class="card-header">
                    <div>
                        <i class="fas fa-user-slash"></i> Devices with Disabled Primary Users
                    </div>
                    <div class="show-all-container">
                        <label class="toggle-switch">
                            <input type="checkbox" id="disabledUsersShowAllToggle">
                            <span class="toggle-slider"></span>
                        </label>
                        <p class="show-all-text">Show all entries</p>
                    </div>
                </div>
                <div class="card-body">
                    <div class="table-responsive">
                        <table id="disabledUsersTable" class="table table-striped table-bordered" style="width:100%">
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
                                {{DISABLED_PRIMARY_USERS_DATA}}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
         
    </div>
    <footer>
        <p>Generated by Roy Klooster - RK Solutions</p>
    </footer>
     
    <script>
        $(document).ready(function() {
            const themeToggle = document.getElementById('themeToggle');
            const prefersDarkScheme = window.matchMedia('(prefers-color-scheme: dark)');
             
            function updateTableColors() {
                if (document.documentElement.getAttribute('data-theme') === 'dark') {
                    $('table.dataTable tbody tr').css('background-color', 'var(--datatable-even-row-bg)');
                    $('table.dataTable tbody tr:nth-child(odd)').css('background-color', 'var(--datatable-odd-row-bg)');
                    $('table.dataTable tbody td').css('color', 'var(--text-color)');
                    $('table.dataTable thead th').css({
                        'background-color': 'var(--table-header-bg)',
                        'color': 'var(--table-header-color)'
                    });
                } else {
                    $('table.dataTable tbody tr').css('background-color', 'var(--datatable-even-row-bg)');
                    $('table.dataTable tbody tr:nth-child(odd)').css('background-color', 'var(--datatable-odd-row-bg)');
                    $('table.dataTable tbody td').css('color', 'var(--text-color)');
                    $('table.dataTable thead th').css({
                        'background-color': 'var(--table-header-bg)',
                        'color': 'var(--table-header-color)'
                    });
                }
            }
             
            const savedTheme = localStorage.getItem('theme');
            if (savedTheme === 'dark' || (!savedTheme && prefersDarkScheme.matches)) {
                document.documentElement.setAttribute('data-theme', 'dark');
                themeToggle.checked = true;
            }
             
            themeToggle.addEventListener('change', function() {
                if (this.checked) {
                    document.documentElement.setAttribute('data-theme', 'dark');
                    localStorage.setItem('theme', 'dark');
                } else {
                    document.documentElement.setAttribute('data-theme', 'light');
                    localStorage.setItem('theme', 'light');
                }
                setTimeout(updateTableColors, 50);
            });
             
            $('.report-tab').on('click', function() {
                $('.report-tab').removeClass('active');
                $('.report-panel').removeClass('active');
                $(this).addClass('active');
                const panelId = $(this).data('panel');
                $(`#${panelId}`).addClass('active');
                setTimeout(function() {
                    $.fn.dataTable.tables({ visible: true, api: true }).columns.adjust();
                }, 10);
            });
             
            const tableOptions = {
                dom: 'Bfrtip',
                buttons: [
                    { extend: 'collection', text: '<i class="fas fa-download"></i> Export',
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
                order: [[1, 'asc']],
                language: { search: "<i class='fas fa-search'></i> _INPUT_", searchPlaceholder: "Search records..." },
                drawCallback: function() { updateTableColors(); }
            };
             
            const appFailuresTable = $('#appFailuresTable').DataTable(tableOptions);
            const multipleUsersTable = $('#multipleUsersTable').DataTable({...tableOptions, order: [[4, 'desc']]});
            const notEncryptedTable = $('#notEncryptedTable').DataTable(tableOptions);
            const noAutopilotTable = $('#noAutopilotTable').DataTable(tableOptions);
            const inactiveDevicesTable = $('#inactiveDevicesTable').DataTable({...tableOptions, order: [[6, 'asc']]});
            const noncompliantTable = $('#noncompliantTable').DataTable(tableOptions);
            const osEditionTable = $('#osEditionTable').DataTable(tableOptions);
            const disabledUsersTable = $('#disabledUsersTable').DataTable(tableOptions);
             
            // Populate filter dropdowns
            function populateFilters() {
                // Application Failures filters
                populateSelectFromColumn('appFailuresCustomerFilter', appFailuresTable, 0); // Customer
                populateSelectFromColumn('appFailuresAppFilter', appFailuresTable, 1); // Application
                populateSelectFromColumn('appFailuresPlatformFilter', appFailuresTable, 2); // Platform
                populateSelectFromColumn('appFailuresVersionFilter', appFailuresTable, 3); // Version
                // Note: Percentage filter uses predefined ranges, no need to populate from data
                 
                // Multiple Users filters
                populateSelectFromColumn('multipleUsersCustomerFilter', multipleUsersTable, 0);
                populateSelectFromColumn('multipleUsersDeviceFilter', multipleUsersTable, 1);
                populateSelectFromColumn('multipleUsersPrimaryUserFilter', multipleUsersTable, 2);
                populateSelectFromColumn('multipleUsersProfileFilter', multipleUsersTable, 3);
                 
                // Not Encrypted filters
                populateSelectFromColumn('notEncryptedCustomerFilter', notEncryptedTable, 0);
                populateSelectFromColumn('notEncryptedDeviceFilter', notEncryptedTable, 1);
                populateSelectFromColumn('notEncryptedUserFilter', notEncryptedTable, 2);
                populateSelectFromColumn('notEncryptedManufacturerFilter', notEncryptedTable, 4);
                populateSelectFromColumn('notEncryptedModelFilter', notEncryptedTable, 5);
                 
                // No Autopilot filters
                populateSelectFromColumn('noAutopilotCustomerFilter', noAutopilotTable, 0);
                populateSelectFromColumn('noAutopilotDeviceFilter', noAutopilotTable, 1);
                populateSelectFromColumn('noAutopilotUserFilter', noAutopilotTable, 2);
                populateSelectFromColumn('noAutopilotManufacturerFilter', noAutopilotTable, 4);
                populateSelectFromColumn('noAutopilotModelFilter', noAutopilotTable, 5);
                 
                // Inactive Devices filters
                populateSelectFromColumn('inactiveCustomerFilter', inactiveDevicesTable, 0);
                populateSelectFromColumn('inactiveDeviceFilter', inactiveDevicesTable, 1);
                populateSelectFromColumn('inactiveUserFilter', inactiveDevicesTable, 2);
                populateSelectFromColumn('inactiveManufacturerFilter', inactiveDevicesTable, 4);
                populateSelectFromColumn('inactiveModelFilter', inactiveDevicesTable, 5);
                 
                // Noncompliant Devices filters
                populateSelectFromColumn('noncompliantCustomerFilter', noncompliantTable, 0);
                populateSelectFromColumn('noncompliantDeviceFilter', noncompliantTable, 1);
                populateSelectFromColumn('noncompliantUserFilter', noncompliantTable, 2);
                populateSelectFromColumn('noncompliantManufacturerFilter', noncompliantTable, 4);
                populateSelectFromColumn('noncompliantModelFilter', noncompliantTable, 5);
                populateSelectFromColumn('noncompliantReasonFilter', noncompliantTable, 7);
 
                // OS Edition Overview filters
                populateSelectFromColumn('osEditionCustomerFilter', osEditionTable, 0);
                populateSelectFromColumn('osEditionDeviceFilter', osEditionTable, 1);
                populateSelectFromColumn('osEditionUserFilter', osEditionTable, 2);
                populateSelectFromColumn('osEditionEditionFilter', osEditionTable, 3);
                populateSelectFromColumn('osEditionFriendlyNameFilter', osEditionTable, 4);
 
                // Disabled Primary Users filters
                populateSelectFromColumn('disabledUsersCustomerFilter', disabledUsersTable, 0);
                populateSelectFromColumn('disabledUsersDeviceFilter', disabledUsersTable, 1);
                populateSelectFromColumn('disabledUsersUserFilter', disabledUsersTable, 2);
                populateSelectFromColumn('disabledUsersManufacturerFilter', disabledUsersTable, 4);
                populateSelectFromColumn('disabledUsersModelFilter', disabledUsersTable, 5);
            }
             
            function populateSelectFromColumn(selectId, table, columnIndex) {
                const values = [...new Set(table.column(columnIndex).data().toArray())].sort();
                const select = $(`#${selectId}`);
                values.forEach(value => {
                    if (value && value.trim() !== '') {
                        select.append(`<option value="${value}">${value}</option>`);
                    }
                });
            }
             
            // Application Failures filter functions
            window.applyAppFailuresFilters = function() {
                const customerFilter = $('#appFailuresCustomerFilter').val();
                const appFilter = $('#appFailuresAppFilter').val();
                const platformFilter = $('#appFailuresPlatformFilter').val();
                const versionFilter = $('#appFailuresVersionFilter').val();
                const percentageFilter = $('#appFailuresPercentageFilter').val();
                 
                appFailuresTable.columns().search('').draw();
                 
                if (customerFilter) appFailuresTable.column(0).search('^' + customerFilter + '$', true, false);
                if (appFilter) appFailuresTable.column(1).search('^' + appFilter + '$', true, false);
                if (platformFilter) appFailuresTable.column(2).search('^' + platformFilter + '$', true, false);
                if (versionFilter) appFailuresTable.column(3).search('^' + versionFilter + '$', true, false);
                if (percentageFilter) {
                    let regex = '';
                    if (percentageFilter === '0-20') {
                        regex = '^(0|[1-9]|1[0-9]|20)%$';
                    } else if (percentageFilter === '20-40') {
                        regex = '^(2[0-9]|3[0-9]|40)%$';
                    } else if (percentageFilter === '40-60') {
                        regex = '^(4[0-9]|5[0-9]|60)%$';
                    } else if (percentageFilter === '60-80') {
                        regex = '^(6[0-9]|7[0-9]|80)%$';
                    } else if (percentageFilter === '80-100') {
                        regex = '^(8[0-9]|9[0-9]|100)%$';
                    }
                    if (regex) appFailuresTable.column(5).search(regex, true, false);
                }
                 
                appFailuresTable.draw();
            };
 
            window.clearAppFailuresFilters = function() {
                $('#appFailuresCustomerFilter, #appFailuresAppFilter, #appFailuresPlatformFilter, #appFailuresVersionFilter, #appFailuresPercentageFilter').val('');
                appFailuresTable.search('').columns().search('').draw();
            };
 
            // Multiple Users filter functions
            window.applyMultipleUsersFilters = function() {
                const customerFilter = $('#multipleUsersCustomerFilter').val();
                const deviceFilter = $('#multipleUsersDeviceFilter').val();
                const userFilter = $('#multipleUsersPrimaryUserFilter').val();
                const profileFilter = $('#multipleUsersProfileFilter').val();
                const countFilter = $('#multipleUsersCountFilter').val();
                 
                multipleUsersTable.columns().search('').draw();
                 
                if (customerFilter) multipleUsersTable.column(0).search('^' + customerFilter + '$', true, false);
                if (deviceFilter) multipleUsersTable.column(1).search('^' + deviceFilter + '$', true, false);
                if (userFilter) multipleUsersTable.column(2).search('^' + userFilter + '$', true, false);
                if (profileFilter) multipleUsersTable.column(3).search('^' + profileFilter + '$', true, false);
                if (countFilter) {
                    if (countFilter === '2') multipleUsersTable.column(4).search('^2$', true, false);
                    else if (countFilter === '3') multipleUsersTable.column(4).search('^3$', true, false);
                    else if (countFilter === '4+') multipleUsersTable.column(4).search('[4-9]|[1-9][0-9]+', true, false);
                }
                 
                multipleUsersTable.draw();
            };
 
            window.clearMultipleUsersFilters = function() {
                $('#multipleUsersCustomerFilter, #multipleUsersDeviceFilter, #multipleUsersPrimaryUserFilter, #multipleUsersProfileFilter, #multipleUsersCountFilter').val('');
                multipleUsersTable.search('').columns().search('').draw();
            };
 
            // Not Encrypted filter functions
            window.applyNotEncryptedFilters = function() {
                const customerFilter = $('#notEncryptedCustomerFilter').val();
                const deviceFilter = $('#notEncryptedDeviceFilter').val();
                const userFilter = $('#notEncryptedUserFilter').val();
                const manufacturerFilter = $('#notEncryptedManufacturerFilter').val();
                const modelFilter = $('#notEncryptedModelFilter').val();
                 
                notEncryptedTable.columns().search('').draw();
                 
                if (customerFilter) notEncryptedTable.column(0).search('^' + customerFilter + '$', true, false);
                if (deviceFilter) notEncryptedTable.column(1).search('^' + deviceFilter + '$', true, false);
                if (userFilter) notEncryptedTable.column(2).search('^' + userFilter + '$', true, false);
                if (manufacturerFilter) notEncryptedTable.column(4).search('^' + manufacturerFilter + '$', true, false);
                if (modelFilter) notEncryptedTable.column(5).search('^' + modelFilter + '$', true, false);
                 
                notEncryptedTable.draw();
            };
 
            window.clearNotEncryptedFilters = function() {
                $('#notEncryptedCustomerFilter, #notEncryptedDeviceFilter, #notEncryptedUserFilter, #notEncryptedManufacturerFilter, #notEncryptedModelFilter').val('');
                notEncryptedTable.search('').columns().search('').draw();
            };
 
            // No Autopilot filter functions
            window.applyNoAutopilotFilters = function() {
                const customerFilter = $('#noAutopilotCustomerFilter').val();
                const deviceFilter = $('#noAutopilotDeviceFilter').val();
                const userFilter = $('#noAutopilotUserFilter').val();
                const manufacturerFilter = $('#noAutopilotManufacturerFilter').val();
                const modelFilter = $('#noAutopilotModelFilter').val();
                 
                noAutopilotTable.columns().search('').draw();
                 
                if (customerFilter) noAutopilotTable.column(0).search('^' + customerFilter + '$', true, false);
                if (deviceFilter) noAutopilotTable.column(1).search('^' + deviceFilter + '$', true, false);
                if (userFilter) noAutopilotTable.column(2).search('^' + userFilter + '$', true, false);
                if (manufacturerFilter) noAutopilotTable.column(4).search('^' + manufacturerFilter + '$', true, false);
                if (modelFilter) noAutopilotTable.column(5).search('^' + modelFilter + '$', true, false);
                 
                noAutopilotTable.draw();
            };
 
            window.clearNoAutopilotFilters = function() {
                $('#noAutopilotCustomerFilter, #noAutopilotDeviceFilter, #noAutopilotUserFilter, #noAutopilotManufacturerFilter, #noAutopilotModelFilter').val('');
                noAutopilotTable.search('').columns().search('').draw();
            };
 
            // Inactive Devices filter functions
            window.applyInactiveFilters = function() {
                const customerFilter = $('#inactiveCustomerFilter').val();
                const deviceFilter = $('#inactiveDeviceFilter').val();
                const userFilter = $('#inactiveUserFilter').val();
                const manufacturerFilter = $('#inactiveManufacturerFilter').val();
                const modelFilter = $('#inactiveModelFilter').val();
                const inactivityFilter = $('#inactiveInactivityFilter').val();
                 
                inactiveDevicesTable.columns().search('').draw();
                 
                if (customerFilter) inactiveDevicesTable.column(0).search('^' + customerFilter + '$', true, false);
                if (deviceFilter) inactiveDevicesTable.column(1).search('^' + deviceFilter + '$', true, false);
                if (userFilter) inactiveDevicesTable.column(2).search('^' + userFilter + '$', true, false);
                if (manufacturerFilter) inactiveDevicesTable.column(4).search('^' + manufacturerFilter + '$', true, false);
                if (modelFilter) inactiveDevicesTable.column(5).search('^' + modelFilter + '$', true, false);
                 
                inactiveDevicesTable.draw();
            };
 
            window.clearInactiveFilters = function() {
                $('#inactiveCustomerFilter, #inactiveDeviceFilter, #inactiveUserFilter, #inactiveManufacturerFilter, #inactiveModelFilter, #inactiveInactivityFilter').val('');
                inactiveDevicesTable.search('').columns().search('').draw();
            };
 
            // Noncompliant Devices filter functions
            window.applyNoncompliantFilters = function() {
                const customerFilter = $('#noncompliantCustomerFilter').val();
                const deviceFilter = $('#noncompliantDeviceFilter').val();
                const userFilter = $('#noncompliantUserFilter').val();
                const manufacturerFilter = $('#noncompliantManufacturerFilter').val();
                const modelFilter = $('#noncompliantModelFilter').val();
                const reasonFilter = $('#noncompliantReasonFilter').val();
                 
                noncompliantTable.columns().search('').draw();
                 
                if (customerFilter) noncompliantTable.column(0).search('^' + customerFilter + '$', true, false);
                if (deviceFilter) noncompliantTable.column(1).search('^' + deviceFilter + '$', true, false);
                if (userFilter) noncompliantTable.column(2).search('^' + userFilter + '$', true, false);
                if (manufacturerFilter) noncompliantTable.column(4).search('^' + manufacturerFilter + '$', true, false);
                if (modelFilter) noncompliantTable.column(5).search('^' + modelFilter + '$', true, false);
                if (reasonFilter) noncompliantTable.column(7).search('^' + reasonFilter + '$', true, false);
                 
                noncompliantTable.draw();
            };
 
            // OS Edition Overview filter functions
            window.applyOSEditionFilters = function() {
                const customerFilter = $('#osEditionCustomerFilter').val();
                const deviceFilter = $('#osEditionDeviceFilter').val();
                const userFilter = $('#osEditionUserFilter').val();
                const editionFilter = $('#osEditionEditionFilter').val();
                const friendlyNameFilter = $('#osEditionFriendlyNameFilter').val();
                 
                osEditionTable.columns().search('').draw();
                 
                if (customerFilter) osEditionTable.column(0).search('^' + customerFilter + '$', true, false);
                if (deviceFilter) osEditionTable.column(1).search('^' + deviceFilter + '$', true, false);
                if (userFilter) osEditionTable.column(2).search('^' + userFilter + '$', true, false);
                if (editionFilter) osEditionTable.column(3).search('^' + editionFilter + '$', true, false);
                if (friendlyNameFilter) osEditionTable.column(4).search('^' + friendlyNameFilter + '$', true, false);
                 
                osEditionTable.draw();
            };
 
            // Disabled Primary Users filter functions
            window.applyDisabledUsersFilters = function() {
                const customerFilter = $('#disabledUsersCustomerFilter').val();
                const deviceFilter = $('#disabledUsersDeviceFilter').val();
                const userFilter = $('#disabledUsersUserFilter').val();
                const manufacturerFilter = $('#disabledUsersManufacturerFilter').val();
                const modelFilter = $('#disabledUsersModelFilter').val();
                 
                disabledUsersTable.columns().search('').draw();
                 
                if (customerFilter) disabledUsersTable.column(0).search('^' + customerFilter + '$', true, false);
                if (deviceFilter) disabledUsersTable.column(1).search('^' + deviceFilter + '$', true, false);
                if (userFilter) disabledUsersTable.column(2).search('^' + userFilter + '$', true, false);
                if (manufacturerFilter) disabledUsersTable.column(4).search('^' + manufacturerFilter + '$', true, false);
                if (modelFilter) disabledUsersTable.column(5).search('^' + modelFilter + '$', true, false);
                 
                disabledUsersTable.draw();
            };
 
            window.clearDisabledUsersFilters = function() {
                $('#disabledUsersCustomerFilter, #disabledUsersDeviceFilter, #disabledUsersUserFilter, #disabledUsersManufacturerFilter, #disabledUsersModelFilter').val('');
                disabledUsersTable.search('').columns().search('').draw();
            };
 
            // Auto-apply filters on change - Disabled Primary Users
            $('#disabledUsersCustomerFilter, #disabledUsersDeviceFilter, #disabledUsersUserFilter, #disabledUsersManufacturerFilter, #disabledUsersModelFilter').on('change', function() {
                applyDisabledUsersFilters();
            });
 
            $('#disabledUsersShowAllToggle').on('change', function() {
                disabledUsersTable.page.len($(this).is(':checked') ? -1 : 10).draw();
            });
 
            window.clearOSEditionFilters = function() {
                $('#osEditionCustomerFilter, #osEditionDeviceFilter, #osEditionUserFilter, #osEditionEditionFilter, #osEditionFriendlyNameFilter').val('');
                osEditionTable.search('').columns().search('').draw();
            };
 
            window.clearNoncompliantFilters = function() {
                $('#noncompliantCustomerFilter, #noncompliantDeviceFilter, #noncompliantUserFilter, #noncompliantManufacturerFilter, #noncompliantModelFilter, #noncompliantReasonFilter').val('');
                noncompliantTable.search('').columns().search('').draw();
            };
 
            // Add automatic filter event listeners after the tables are initialized
            // Add this after the existing table initialization code and before the setTimeout function:
 
            // Auto-apply filters on change - Application Failures
            $('#appFailuresCustomerFilter, #appFailuresAppFilter, #appFailuresPlatformFilter, #appFailuresVersionFilter, #appFailuresPercentageFilter').on('change', function() {
                applyAppFailuresFilters();
            });
 
            // Auto-apply filters on change - Multiple Users
            $('#multipleUsersCustomerFilter, #multipleUsersDeviceFilter, #multipleUsersPrimaryUserFilter, #multipleUsersProfileFilter, #multipleUsersCountFilter').on('change', function() {
                applyMultipleUsersFilters();
            });
 
            // Auto-apply filters on change - Not Encrypted
            $('#notEncryptedCustomerFilter, #notEncryptedDeviceFilter, #notEncryptedUserFilter, #notEncryptedManufacturerFilter, #notEncryptedModelFilter').on('change', function() {
                applyNotEncryptedFilters();
            });
 
            // Auto-apply filters on change - No Autopilot
            $('#noAutopilotCustomerFilter, #noAutopilotDeviceFilter, #noAutopilotUserFilter, #noAutopilotManufacturerFilter, #noAutopilotModelFilter').on('change', function() {
                applyNoAutopilotFilters();
            });
 
            // Auto-apply filters on change - Inactive Devices
            $('#inactiveCustomerFilter, #inactiveDeviceFilter, #inactiveUserFilter, #inactiveManufacturerFilter, #inactiveModelFilter, #inactiveInactivityFilter').on('change', function() {
                applyInactiveFilters();
            });
 
            // Auto-apply filters on change - Noncompliant Devices
            $('#noncompliantCustomerFilter, #noncompliantDeviceFilter, #noncompliantUserFilter, #noncompliantManufacturerFilter, #noncompliantModelFilter, #noncompliantReasonFilter').on('change', function() {
                applyNoncompliantFilters();
            });
 
            // Auto-apply filters on change - OS Edition Overview
            $('#osEditionCustomerFilter, #osEditionDeviceFilter, #osEditionUserFilter, #osEditionEditionFilter, #osEditionFriendlyNameFilter').on('change', function() {
                applyOSEditionFilters();
            });
             
            // Show all toggle functions for each table
            $('#appFailuresShowAllToggle').on('change', function() {
                appFailuresTable.page.len($(this).is(':checked') ? -1 : 10).draw();
            });
             
            $('#multipleUsersShowAllToggle').on('change', function() {
                multipleUsersTable.page.len($(this).is(':checked') ? -1 : 10).draw();
            });
             
            $('#notEncryptedShowAllToggle').on('change', function() {
                notEncryptedTable.page.len($(this).is(':checked') ? -1 : 10).draw();
            });
             
            $('#noAutopilotShowAllToggle').on('change', function() {
                noAutopilotTable.page.len($(this).is(':checked') ? -1 : 10).draw();
            });
             
            $('#inactiveDevicesShowAllToggle').on('change', function() {
                inactiveDevicesTable.page.len($(this).is(':checked') ? -1 : 10).draw();
            });
             
            $('#noncompliantShowAllToggle').on('change', function() {
                noncompliantTable.page.len($(this).is(':checked') ? -1 : 10).draw();
            });
 
            $('#osEditionShowAllToggle').on('change', function() {
                osEditionTable.page.len($(this).is(':checked') ? -1 : 10).draw();
            });
             
            setTimeout(function() {
                populateFilters();
                updateTableColors();
            }, 100);
             
            $(window).on('load', function() {
                setTimeout(updateTableColors, 200);
            });
        });
        </script>
</body>
</html>
'@
    
    # Generate table rows for all application failures
    $applicationFailureRows = ""
    foreach ($item in $Report_ApplicationFailureReport) {
        $applicationFailureRows += @"
        <tr>
            <td>$($item.Customer)</td>
            <td>$($item.Application)</td>
            <td>$($item.Platform)</td>
            <td>$($item.Version)</td>
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
            <td>$($item.Customer)</td>
            <td>$($item.DeviceName)</td>
            <td>$($item.PrimaryUser)</td>
            <td>$($item.EnrollmentProfile)</td>
            <td><span class="badge badge-permanent">$($item.usersLoggedOnCount)</span></td>
            <td>$($item.usersLoggedOnIds)</td>
        </tr>
"@
    }

    # Generate table rows for not encrypted devices
    $notEncryptedRows = ""
    foreach ($item in $Report_NotEncryptedDevices) {
        $notEncryptedRows += @"
        <tr>
            <td>$($item.Customer)</td>
            <td>$($item.DeviceName)</td>
            <td>$($item.PrimaryUser)</td>
            <td>$($item.Serialnumber)</td>
            <td>$($item.DeviceManufacturer)</td>
            <td>$($item.DeviceModel)</td>
        </tr>
"@
    }

    # Generate table rows for Non-company owned devices
    $noAutopilotHashRows = ""
    foreach ($item in $Report_DevicesWithoutAutopilotHash) {
        $noAutopilotHashRows += @"
        <tr>
            <td>$($item.Customer)</td>
            <td>$($item.DeviceName)</td>
            <td>$($item.PrimaryUser)</td>
            <td>$($item.Serialnumber)</td>
            <td>$($item.DeviceManufacturer)</td>
            <td>$($item.DeviceModel)</td>
        </tr>
"@
    }

    # Generate table rows for inactive devices
    $inactiveDevicesRows = ""
    foreach ($item in $Report_InactiveDevices) {
        $inactiveDevicesRows += @"
        <tr>
            <td>$($item.Customer)</td>
            <td>$($item.DeviceName)</td>
            <td>$($item.PrimaryUser)</td>
            <td>$($item.Serialnumber)</td>
            <td>$($item.DeviceManufacturer)</td>
            <td>$($item.DeviceModel)</td>
            <td>$($item.LastContact)</td>
        </tr>
"@
    }

    # Generate table rows for noncompliant devices
    $noncompliantDevicesRows = ""
    foreach ($item in $Report_NoncompliantDevices) {
        $statusBadge = '<span class="badge badge-permanent">Noncompliant</span>'
        
        $noncompliantDevicesRows += @"
        <tr>
            <td>$($item.Customer)</td>
            <td>$($item.DeviceName)</td>
            <td>$($item.PrimaryUser)</td>
            <td>$($item.Serialnumber)</td>
            <td>$($item.DeviceManufacturer)</td>
            <td>$($item.DeviceModel)</td>
            <td>$statusBadge</td>
            <td>$($item.NoncompliantBasedOn)</td>
            <td>$($item.NoncompliantAlert)</td>
        </tr>
"@
    }

    # Generate table rows for OS Edition Overview
    $osEditionOverviewRows = ""
    foreach ($item in $Report_OperationSystemEdtionOverview) {
        $osEditionOverviewRows += @"
        <tr>
            <td>$($item.Customer)</td>
            <td>$($item.DeviceName)</td>
            <td>$($item.PrimaryUser)</td>
            <td>$($item.OperatingSystemEdition)</td>
            <td>$($item.OSFriendlyname)</td>
        </tr>
"@
    }

    # Generate table rows for disabled primary users
    $disabledPrimaryUsersRows = ""
    foreach ($item in $Report_DisabledPrimaryUsers) {
        $disabledPrimaryUsersRows += @"
        <tr>
            <td>$($item.Customer)</td>
            <td>$($item.DeviceName)</td>
            <td><span class="badge badge-permanent">$($item.PrimaryUser)</span></td>
            <td>$($item.Serialnumber)</td>
            <td>$($item.DeviceManufacturer)</td>
            <td>$($item.DeviceModel)</td>
        </tr>
"@
    }

    # Replace placeholders in template with actual values
    $htmlContent = $htmlTemplate
    $htmlContent = $htmlContent.Replace('$TenantName', $TenantName)
    $htmlContent = $htmlContent.Replace('$ReportDate', $currentDate)
    $htmlContent = $htmlContent.Replace('{{APPLICATION_FAILURES_DATA}}', $applicationFailureRows)
    $htmlContent = $htmlContent.Replace('{{MULTIPLE_USERS_DATA}}', $multipleUsersRows)
    $htmlContent = $htmlContent.Replace('{{NOT_ENCRYPTED_DATA}}', $notEncryptedRows)
    $htmlContent = $htmlContent.Replace('{{NO_AUTOPILOT_DATA}}', $noAutopilotHashRows)
    $htmlContent = $htmlContent.Replace('{{INACTIVE_DEVICES_DATA}}', $inactiveDevicesRows)
    $htmlContent = $htmlContent.Replace('{{NONCOMPLIANT_DEVICES_DATA}}', $noncompliantDevicesRows)
    $htmlContent = $htmlContent.Replace('{{OS_EDITION_OVERVIEW_DATA}}', $osEditionOverviewRows)
    $htmlContent = $htmlContent.Replace('{{DISABLED_PRIMARY_USERS_DATA}}', $disabledPrimaryUsersRows)
    $htmlContent = $htmlContent.Replace('$applicationFailures', $Report_ApplicationFailureReport_Count)
    $htmlContent = $htmlContent.Replace('$multipleUsers', $Report_DevicesWithMultipleUsers_Count)
    $htmlContent = $htmlContent.Replace('$notEncrypted', $Report_NotEncryptedDevices_Count)
    $htmlContent = $htmlContent.Replace('$noAutopilot', $Report_DevicesWithoutAutopilotHash_Count)
    $htmlContent = $htmlContent.Replace('$inactiveDevices', $Report_InactiveDevices_Count)
    $htmlContent = $htmlContent.Replace('$noncompliantDevices', $Report_NoncompliantDevices_Count)
    $htmlContent = $htmlContent.Replace('$osEditionOverview', $Report_OperationSystemEdtionOverview_Count)
    $htmlContent = $htmlContent.Replace('$disabledPrimaryUsers', $Report_DisabledPrimaryUsers_Count)

    # Export to HTML file
    $htmlContent | Out-File -FilePath $ExportPath -Encoding utf8

    # Set script-scoped variable for email attachment
    $script:ExportPath = $ExportPath

    Write-Host "INFO: All actions completed successfully."
    Write-Host "INFO: Intune Anomalies Report saved to: $ExportPath" -ForegroundColor Cyan

    # Open the HTML file only if we're not sending email
    if (-not $SendEmail) {
        Invoke-Item $ExportPath
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
        'managementAgent' # Indicates the management agent used (e.g., Intune)
    )

    # Get all Windows Devices from Microsoft Intune
    $AllDeviceData = Invoke-graphRequestWithPaging -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'&`$select=$($Properties -join ',')"
    #filter out managed by MDE
    $AllDeviceData = $AllDeviceData | Where-Object { $_.managementAgent -ne "msSense" }
    
    # Get all AutoPilot registered devices under "Enrollment"
    Write-Host "Fetching Autopilot devices..." -ForegroundColor Yellow
    $AutopilotDevices = (Invoke-GraphRequestWithPaging -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities")

    # Loop through all devices for device data
    $results = @()
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
            # Get detailed device properties
            $DeviceProperties = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/manageddevices/$($DeviceData.id)?`$select=$($Properties -join ',')" -ErrorAction Stop
            
            # Process Autopilot information
            $AutopilotInfo = $AutopilotDevices | Where-Object { $_.serialnumber -eq $DeviceData.SerialNumber } 
            $HashUploaded = $DeviceData.Serialnumber -in $AutopilotDevices.serialnumber

            # Initialize compliance rule variables
            $allRules = @()
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
                                            $allRules += $ruleDetail.setting
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
            $ExistingLoggedInUsers = @()

            if ($LoggedInUsers) {
                foreach ($user in $LoggedInUsers) {
                    if ($user -in $AllEntraIDUsers.Id) {
                        try {
                            $userPrincipalName = (Invoke-MgGraphRequest -Uri "/beta/users/$user" -ErrorAction SilentlyContinue).userprincipalname
                            if ($userPrincipalName) {
                                $ExistingLoggedInUsers += $userPrincipalName
                            }
                        } catch {
                            Write-Verbose "Failed to get user details for ID $user $_"
                        }
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

            $results += [PSCustomObject][ordered]@{
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
                OperatingSystemEdition     = if ($hardwareInfo.operatingSystemEdition) { $hardwareInfo.operatingSystemEdition } else { "Unknown" }
                operatingSystemProductType = if ($hardwareInfo.operatingSystemProductType) { Get-OperatingSystemProductType -Customer "$($hardwareInfo.operatingSystemProductType)" } else { "Unknown" }
                BiosVersion                = if ($hardwareInfo.systemManagementBIOSVersion) { $hardwareInfo.systemManagementBIOSVersion } else { "Unknown" }
                ComplianceStatus           = $DeviceProperties.ComplianceState
                # **FIX**: Use unique rules to prevent duplicates
                NoncompliantBasedOn        = if ($uniqueRules) { $uniqueRules -join ', ' } else { "" }
                NoncompliantAlert          = if ($uniqueRules) { ($uniqueRules | Where-Object { $_ -notin $FilteredForAlerting }) -join ', ' } else { "" }
            }
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
    
    # Set temporary directory and file path based on OS
    if ($detectedWindows) {
        $tempDir = "C:\temp"
    } else {
        # macOS and Linux
        $tempDir = "/tmp"
    }
    
    # Use Join-Path for cross-platform compatibility
    $Data = Join-Path $tempDir "Data.log"
    
    # Ensure temp directory exists
    if (-not (Test-Path -Path $tempDir)) {
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    }

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

    #Maak een nieuwe array
    $ReturnObject = New-Object System.Collections.ArrayList

    #voor elke waardeverzameling in values
    foreach ($value in $Response.Values) {
        #maak een nieuw lineobject[hashtable]
        $LineObject = @{ }

        #voor elke prop in het het schema
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

