function New-RKSolutionsReportTemplate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantName,

        [Parameter(Mandatory = $true)]
        [string]$ReportTitle,

        [Parameter(Mandatory = $true)]
        [string]$ReportSlug,

        [Parameter(Mandatory = $true)]
        [string]$Eyebrow,

        [Parameter(Mandatory = $false)]
        [string]$Lede,

        [Parameter(Mandatory = $false)]
        [string]$StatsCardsHtml = '',

        [Parameter(Mandatory = $true)]
        [string]$BodyContentHtml,

        [Parameter(Mandatory = $false)]
        [string]$CustomCss = '',

        [Parameter(Mandatory = $false)]
        [string]$ReportDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'),

        [Parameter(Mandatory = $false)]
        [string[]]$Tags = @()
    )

    # Build tag pills HTML
    $tagPillsHtml = ''
    if ($Tags.Count -gt 0) {
        $tagPillsHtml = '<div class="rk-tags">'
        foreach ($tag in $Tags) {
            $tagPillsHtml += "<span class=`"rk-tag`">$tag</span>"
        }
        $tagPillsHtml += '</div>'
    }

    # Build lede HTML
    $ledeHtml = ''
    if ($Lede) {
        $ledeHtml = "<p class=`"rk-lede`">$Lede</p>"
    }

    # Build stats section
    $statsHtml = ''
    if ($StatsCardsHtml) {
        $statsHtml = @"
        <div class="rk-stats">
$StatsCardsHtml
        </div>
"@
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>$TenantName - $Eyebrow</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600;700&family=Playfair+Display:ital,wght@0,700;0,800;0,900;1,700&family=Source+Serif+4:ital,opsz,wght@0,8..60,400;0,8..60,600;1,8..60,400&display=swap" rel="stylesheet">
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
<style>
/* ========================================
   RKSolutions Report Template
   Brand: rksolutions.nl editorial parchment
   ======================================== */

/* --- Light Theme (default) --- */
:root {
    --bg-base: #fefefe;
    --bg-elevated: #f0ece4;
    --bg-warm: #eae4d8;
    --bg-deep: #f0ece4;
    --border: #ddd6c8;
    --border-dashed: #ddd6c8;
    --text: #1a1510;
    --text-body: #3a3228;
    --text-muted: #807060;
    --text-dim: #a89878;
    --accent: #c06828;
    --accent-hover: #d87830;
    --accent-soft: #fde8d4;
    --success: #28904a;
    --warn: #d88020;
    --error: #d83830;
    --tile-rust: #c06828;
    --tile-olive: #2e9040;
    --tile-steel: #2878b8;
    --tile-rose: #d04848;
    --toggle-bg: #ddd6c8;
    --input-bg: #fefefe;
    --input-border: #ddd6c8;
    --input-color: #3a3228;
    --button-bg: #f0ece4;
    --button-color: #3a3228;
    --button-border: #ddd6c8;
    --button-hover-bg: #eae4d8;
}

/* --- Dark Theme (warm parchment, vivid) --- */
[data-theme="dark"] {
    --bg-base: #1c1916;
    --bg-elevated: #282420;
    --bg-warm: #302a24;
    --bg-deep: #161410;
    --border: #403830;
    --border-dashed: #403830;
    --text: #f0e8d8;
    --text-body: #f0e8d8;
    --text-muted: #b0a088;
    --text-dim: #685840;
    --accent: #f0a850;
    --accent-hover: #f8b860;
    --accent-soft: rgba(240,168,80,0.1);
    --success: #60e068;
    --warn: #f0b840;
    --error: #f87068;
    --tile-rust-bg: #4a2e10;
    --tile-rust-border: #7a5028;
    --tile-rust-text: #f8c870;
    --tile-rust-eyebrow: #e0a850;
    --tile-rust-caption: #b08038;
    --tile-olive-bg: #183818;
    --tile-olive-border: #286828;
    --tile-olive-text: #80f088;
    --tile-olive-eyebrow: #60d060;
    --tile-olive-caption: #408840;
    --tile-steel-bg: #102838;
    --tile-steel-border: #205878;
    --tile-steel-text: #70d0f8;
    --tile-steel-eyebrow: #58b0e0;
    --tile-steel-caption: #3890a8;
    --tile-rose-bg: #381818;
    --tile-rose-border: #683030;
    --tile-rose-text: #f88880;
    --tile-rose-eyebrow: #e06860;
    --tile-rose-caption: #b04840;
    --toggle-bg: #403830;
    --input-bg: #1c1916;
    --input-border: #403830;
    --input-color: #f0e8d8;
    --button-bg: #282420;
    --button-color: #f0e8d8;
    --button-border: #403830;
    --button-hover-bg: #302a24;
}

/* --- Base --- */
*, *::before, *::after { box-sizing: border-box; }

body {
    font-family: 'Source Serif 4', Georgia, serif;
    margin: 0; padding: 0;
    background-color: var(--bg-base);
    color: var(--text-body);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    transition: background-color 0.3s ease, color 0.3s ease;
}

.rk-container {
    max-width: 1600px;
    margin: 0 auto;
    padding: 24px 30px;
    flex: 1;
}

/* --- Theme Toggle --- */
.rk-theme-toggle {
    position: fixed;
    top: 16px;
    right: 20px;
    z-index: 1000;
    display: flex;
    align-items: center;
    gap: 8px;
    background: var(--bg-elevated);
    border: 1px solid var(--border);
    padding: 6px 14px;
    border-radius: 20px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.7rem;
    color: var(--text-muted);
    transition: background-color 0.3s, border-color 0.3s;
}
.rk-theme-toggle-switch {
    position: relative;
    display: inline-block;
    width: 40px; height: 22px;
}
.rk-theme-toggle-switch input { opacity: 0; width: 0; height: 0; }
.rk-theme-toggle-slider {
    position: absolute; cursor: pointer;
    top: 0; left: 0; right: 0; bottom: 0;
    background-color: var(--toggle-bg);
    transition: 0.3s; border-radius: 22px;
}
.rk-theme-toggle-slider:before {
    position: absolute; content: "";
    height: 16px; width: 16px; left: 3px; bottom: 3px;
    background-color: white;
    transition: 0.3s; border-radius: 50%;
}
input:checked + .rk-theme-toggle-slider { background-color: var(--accent); }
input:checked + .rk-theme-toggle-slider:before { transform: translateX(18px); }

/* --- Breadcrumb Pill --- */
.rk-breadcrumb {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: var(--bg-elevated);
    border: 1px solid var(--border);
    border-radius: 20px;
    padding: 6px 16px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.78rem;
    color: var(--text-muted);
    margin-bottom: 24px;
    transition: background-color 0.3s, border-color 0.3s;
}
.rk-breadcrumb .rk-arrow { color: var(--accent); }

/* --- Eyebrow --- */
.rk-eyebrow {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.68rem;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 8px;
}

/* --- Title --- */
.rk-title {
    font-family: 'Playfair Display', serif;
    font-size: 2rem;
    font-weight: 800;
    color: var(--text);
    letter-spacing: -0.02em;
    margin-bottom: 8px;
    line-height: 1.2;
}
.rk-title .rk-accent { color: var(--accent); }

/* --- Lede --- */
.rk-lede {
    font-family: 'Source Serif 4', serif;
    font-style: italic;
    color: var(--text-muted);
    font-size: 1rem;
    margin-bottom: 20px;
    max-width: 720px;
    line-height: 1.5;
}

/* --- Tags --- */
.rk-tags { margin-bottom: 20px; display: flex; flex-wrap: wrap; gap: 6px; }
.rk-tag {
    display: inline-block;
    background: var(--bg-warm);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 3px 12px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.65rem;
    color: var(--accent);
    transition: background-color 0.3s, border-color 0.3s, color 0.3s;
}

/* --- Dashed Divider --- */
.rk-divider {
    border: none;
    border-top: 1px dashed var(--border-dashed);
    margin: 24px 0;
}

/* --- Stat Tiles --- */
.rk-stats {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 16px;
    margin-bottom: 28px;
}
.rk-stat-tile {
    border-radius: 14px;
    padding: 20px 16px;
    text-align: center;
    transition: transform 0.2s, background-color 0.3s, border-color 0.3s;
}
.rk-stat-tile:hover { transform: translateY(-3px); }
.rk-stat-eyebrow {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.62rem;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    margin-bottom: 6px;
}
.rk-stat-number {
    font-family: 'Playfair Display', serif;
    font-size: 2.2rem;
    font-weight: 800;
    line-height: 1.1;
}
.rk-stat-caption {
    font-family: 'Source Serif 4', serif;
    font-size: 0.72rem;
    margin-top: 4px;
}

/* Light mode tiles: gradient bg, white text */
.rk-stat-tile.t-rust { background: linear-gradient(135deg, #d06828, #e08040); color: #fff; }
.rk-stat-tile.t-rust .rk-stat-eyebrow { opacity: 0.92; }
.rk-stat-tile.t-rust .rk-stat-caption { opacity: 0.85; }
.rk-stat-tile.t-olive { background: linear-gradient(135deg, #2e9040, #48b060); color: #fff; }
.rk-stat-tile.t-olive .rk-stat-eyebrow { opacity: 0.92; }
.rk-stat-tile.t-olive .rk-stat-caption { opacity: 0.85; }
.rk-stat-tile.t-steel { background: linear-gradient(135deg, #2878b8, #4098d0); color: #fff; }
.rk-stat-tile.t-steel .rk-stat-eyebrow { opacity: 0.92; }
.rk-stat-tile.t-steel .rk-stat-caption { opacity: 0.85; }
.rk-stat-tile.t-rose { background: linear-gradient(135deg, #d04848, #e06868); color: #fff; }
.rk-stat-tile.t-rose .rk-stat-eyebrow { opacity: 0.92; }
.rk-stat-tile.t-rose .rk-stat-caption { opacity: 0.85; }

/* Dark mode tiles: gradient tinted bg, bright colored text */
[data-theme="dark"] .rk-stat-tile.t-rust {
    background: linear-gradient(135deg, var(--tile-rust-bg), #5a3818); border: 1px solid var(--tile-rust-border); color: var(--tile-rust-text);
}
[data-theme="dark"] .rk-stat-tile.t-rust .rk-stat-eyebrow { color: var(--tile-rust-eyebrow); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-rust .rk-stat-caption { color: var(--tile-rust-caption); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-olive {
    background: linear-gradient(135deg, var(--tile-olive-bg), #204820); border: 1px solid var(--tile-olive-border); color: var(--tile-olive-text);
}
[data-theme="dark"] .rk-stat-tile.t-olive .rk-stat-eyebrow { color: var(--tile-olive-eyebrow); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-olive .rk-stat-caption { color: var(--tile-olive-caption); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-steel {
    background: linear-gradient(135deg, var(--tile-steel-bg), #183848); border: 1px solid var(--tile-steel-border); color: var(--tile-steel-text);
}
[data-theme="dark"] .rk-stat-tile.t-steel .rk-stat-eyebrow { color: var(--tile-steel-eyebrow); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-steel .rk-stat-caption { color: var(--tile-steel-caption); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-rose {
    background: linear-gradient(135deg, var(--tile-rose-bg), #482020); border: 1px solid var(--tile-rose-border); color: var(--tile-rose-text);
}
[data-theme="dark"] .rk-stat-tile.t-rose .rk-stat-eyebrow { color: var(--tile-rose-eyebrow); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-rose .rk-stat-caption { color: var(--tile-rose-caption); opacity: 1; }

/* --- Filter Bar --- */
.rk-filter-bar {
    background: var(--bg-elevated);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 14px 18px;
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    gap: 12px;
    flex-wrap: wrap;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.78rem;
    color: var(--text-muted);
    transition: background-color 0.3s, border-color 0.3s;
}
.rk-filter-chip {
    background: var(--accent-soft);
    color: var(--accent);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 4px 12px;
    font-size: 0.72rem;
    font-family: 'JetBrains Mono', monospace;
    cursor: pointer;
    transition: background-color 0.3s, color 0.3s, border-color 0.3s;
}
.rk-filter-chip:hover { border-color: var(--accent); }
.rk-filter-chip.active { background: var(--accent); color: #fff; border-color: var(--accent); }

/* --- Tables --- */
.rk-card {
    background: var(--bg-elevated);
    border: 1px solid var(--border);
    border-radius: 14px;
    overflow: hidden;
    margin-bottom: 24px;
    transition: background-color 0.3s, border-color 0.3s;
}
.rk-card-header {
    background: var(--bg-warm);
    padding: 14px 20px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--text-body);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    transition: background-color 0.3s, color 0.3s;
}
.rk-card-body {
    padding: 20px;
    background-color: var(--bg-base);
    transition: background-color 0.3s;
}

table.dataTable, .table {
    width: 100% !important;
    border-collapse: collapse !important;
    color: var(--text-body) !important;
    background-color: var(--bg-base) !important;
}
.table thead th {
    background-color: var(--bg-warm) !important;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.68rem;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: var(--text-body) !important;
    font-weight: 600;
    padding: 12px;
    border: 1px solid var(--border) !important;
    transition: background-color 0.3s, color 0.3s, border-color 0.3s;
}
.table tbody td {
    padding: 12px;
    vertical-align: middle;
    border: 1px solid var(--border) !important;
    color: var(--text-body) !important;
    background-color: inherit !important;
    font-family: 'Source Serif 4', serif;
    transition: background-color 0.3s, color 0.3s, border-color 0.3s;
}
.table tbody td.rk-mono {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.82rem;
}
.table tbody tr { background-color: var(--bg-base) !important; transition: background-color 0.2s; }
.table tbody tr:nth-child(even) { background-color: var(--bg-elevated) !important; }
[data-theme="dark"] .table tbody tr { background-color: var(--bg-base) !important; }
[data-theme="dark"] .table tbody tr:nth-child(even) { background-color: var(--bg-deep) !important; }
.table tbody tr:hover { background-color: var(--accent-soft) !important; }
.table tbody tr:hover td { background-color: inherit !important; }

/* Override Bootstrap table backgrounds */
.table-striped > tbody > tr:nth-of-type(odd),
.table-striped > tbody > tr:nth-of-type(even),
.table-striped > tbody > tr > td,
.table > tbody > tr > td,
.table > tbody > tr {
    background-color: inherit !important;
    --bs-table-bg: transparent;
    --bs-table-striped-bg: transparent;
    --bs-table-hover-bg: transparent;
}

/* DataTables wrapper backgrounds */
.dataTables_wrapper { background-color: transparent !important; color: var(--text-body) !important; }
.dataTables_wrapper .row { background-color: transparent !important; }

/* --- Status Badges --- */
.rk-badge {
    display: inline-block;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.65rem;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    padding: 4px 10px;
    border-radius: 6px;
    font-weight: 600;
}
.rk-badge-ok { background: var(--success); color: #fff; }
.rk-badge-warn { background: var(--warn); color: #fff; }
.rk-badge-error { background: var(--error); color: #fff; }
.rk-badge-na { background: var(--bg-warm); color: var(--text-muted); }
.rk-badge-accent { background: var(--accent); color: #fff; }

[data-theme="dark"] .rk-badge-ok { background: rgba(80,210,90,0.2); color: #60e068; border: 1px solid rgba(80,210,90,0.3); }
[data-theme="dark"] .rk-badge-warn { background: rgba(240,170,50,0.2); color: #f0b840; border: 1px solid rgba(240,170,50,0.3); }
[data-theme="dark"] .rk-badge-error { background: rgba(240,80,60,0.2); color: #f87068; border: 1px solid rgba(240,80,60,0.3); }
[data-theme="dark"] .rk-badge-na { background: rgba(160,140,100,0.15); color: #b0a088; border: 1px solid rgba(160,140,100,0.25); }
[data-theme="dark"] .rk-badge-accent { background: rgba(240,168,80,0.2); color: #f0a850; border: 1px solid rgba(240,168,80,0.3); }

/* --- Report Tabs --- */
.rk-tabs {
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
    margin-bottom: 20px;
    border-bottom: 1px dashed var(--border-dashed);
    padding-bottom: 4px;
}
.rk-tab {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.72rem;
    padding: 8px 16px;
    border-radius: 8px 8px 0 0;
    cursor: pointer;
    color: var(--text-muted);
    background: transparent;
    border: 1px solid transparent;
    border-bottom: none;
    transition: all 0.2s;
    letter-spacing: 0.05em;
}
.rk-tab:hover { color: var(--accent); }
.rk-tab.active {
    color: var(--accent);
    background: var(--bg-elevated);
    border-color: var(--border);
    font-weight: 600;
}
.rk-panel { display: none; }
.rk-panel.active { display: block; }

/* --- Show All Toggle --- */
.rk-show-all {
    display: flex;
    align-items: center;
    gap: 10px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.72rem;
    color: var(--text-muted);
}
.rk-toggle-switch {
    position: relative;
    display: inline-block;
    width: 44px; height: 24px;
}
.rk-toggle-switch input { opacity: 0; width: 0; height: 0; }
.rk-toggle-slider {
    position: absolute; cursor: pointer;
    top: 0; left: 0; right: 0; bottom: 0;
    background-color: var(--toggle-bg);
    transition: 0.3s; border-radius: 24px;
}
.rk-toggle-slider:before {
    position: absolute; content: "";
    height: 18px; width: 18px; left: 3px; bottom: 3px;
    background-color: white;
    transition: 0.3s; border-radius: 50%;
}
input:checked + .rk-toggle-slider { background-color: var(--accent); }
input:checked + .rk-toggle-slider:before { transform: translateX(20px); }

/* --- DataTables Overrides --- */
.dataTables_wrapper .dataTables_length,
.dataTables_wrapper .dataTables_filter,
.dataTables_wrapper .dataTables_info,
.dataTables_wrapper .dataTables_paginate { color: var(--text-body) !important; }

.dataTables_wrapper .dataTables_length select,
.dataTables_wrapper .dataTables_filter input {
    border: 1px solid var(--input-border);
    background-color: var(--input-bg);
    color: var(--input-color);
    border-radius: 6px;
    padding: 5px 10px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.78rem;
}

.dataTables_wrapper .dataTables_paginate .paginate_button {
    padding: 0.3em 0.8em;
    border-radius: 6px;
    margin: 0 2px;
    color: var(--button-color) !important;
    border: 1px solid var(--button-border) !important;
    background-color: var(--button-bg) !important;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.75rem;
}
.dataTables_wrapper .dataTables_paginate .paginate_button.current,
.dataTables_wrapper .dataTables_paginate .paginate_button.current:hover {
    background: var(--accent) !important;
    border-color: var(--accent) !important;
    color: white !important;
}
.dataTables_wrapper .dataTables_paginate .paginate_button:hover {
    background: var(--button-hover-bg) !important;
    border-color: var(--border) !important;
    color: var(--button-color) !important;
}
.dataTables_wrapper .dataTables_paginate .paginate_button.disabled,
.dataTables_wrapper .dataTables_paginate .paginate_button.disabled:hover {
    color: var(--text-dim) !important;
    background-color: var(--bg-base) !important;
    border-color: var(--border) !important;
}

div.dt-buttons { background-color: transparent !important; }
div.dt-buttons .dt-button {
    background-color: var(--button-bg) !important;
    color: var(--button-color) !important;
    border: 1px solid var(--button-border) !important;
    border-radius: 6px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.75rem;
}
div.dt-buttons .dt-button:hover {
    background-color: var(--button-hover-bg) !important;
    color: var(--button-color) !important;
}

.form-select, .form-control {
    background-color: var(--input-bg) !important;
    color: var(--input-color) !important;
    border-color: var(--input-border) !important;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.78rem;
}
.form-label { color: var(--text-body); font-family: 'JetBrains Mono', monospace; font-size: 0.75rem; }

.btn-outline-secondary {
    color: var(--text-body);
    border-color: var(--border);
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.75rem;
}
.btn-outline-secondary:hover { background-color: var(--bg-warm); color: var(--text-body); }

/* --- Filter Tags --- */
.rk-filter-tags { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 8px; }
.rk-filter-tag {
    background-color: var(--bg-warm);
    padding: 4px 12px;
    border-radius: 16px;
    font-size: 0.78rem;
    font-family: 'JetBrains Mono', monospace;
    color: var(--text-body);
    display: flex; align-items: center; gap: 6px;
    transition: background-color 0.3s, color 0.3s;
}
.rk-filter-tag i { cursor: pointer; color: var(--text-muted); }
.rk-filter-tag i:hover { color: var(--error); }

/* --- PS Prompt Footer --- */
.rk-footer {
    border-top: 1px solid var(--border);
    padding: 14px 30px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.78rem;
    transition: border-color 0.3s;
}
.rk-footer-prompt { color: var(--text-body); }
.rk-footer-cursor {
    display: inline-block;
    width: 9px; height: 18px;
    background: var(--accent);
    vertical-align: middle;
    margin-left: 4px;
    animation: rk-blink 1.2s step-end infinite;
}
.rk-footer-site { color: var(--text-dim); }
@keyframes rk-blink { 50% { opacity: 0; } }

/* --- Responsive --- */
@media (max-width: 768px) {
    .rk-stats { grid-template-columns: repeat(2, 1fr); }
    .rk-title { font-size: 1.5rem; }
    .rk-container { padding: 16px; }
    .rk-footer { padding: 14px 16px; }
}

/* --- Report-specific CSS injection --- */
$CustomCss
</style>
</head>
<body>

<!-- Theme Toggle -->
<div class="rk-theme-toggle">
    <i class="fas fa-sun" style="font-size: 14px;"></i>
    <label class="rk-theme-toggle-switch">
        <input type="checkbox" id="rkThemeToggle">
        <span class="rk-theme-toggle-slider"></span>
    </label>
    <i class="fas fa-moon" style="font-size: 14px;"></i>
</div>

<div class="rk-container">
    <!-- Breadcrumb Pill -->
    <div class="rk-breadcrumb">
        <span class="rk-arrow">&larr;</span> cd ./reports/$ReportSlug
    </div>

    <!-- Eyebrow -->
    <div class="rk-eyebrow">$Eyebrow &middot; GENERATED $ReportDate</div>

    <!-- Title -->
    <h1 class="rk-title">$TenantName <span class="rk-accent">$ReportTitle</span> Report</h1>

    <!-- Lede -->
    $ledeHtml

    <!-- Tags -->
    $tagPillsHtml

    <!-- Divider -->
    <hr class="rk-divider">

    <!-- Stats -->
    $statsHtml

    <!-- Body Content (tables, tabs, filters -- report-specific) -->
    $BodyContentHtml

</div>

<!-- PS Prompt Footer -->
<div class="rk-footer">
    <span class="rk-footer-prompt">PS C:\Blog\rksolutions&gt;<span class="rk-footer-cursor"></span></span>
    <span class="rk-footer-site">rksolutions.nl</span>
</div>

<script>
`$(document).ready(function() {
    // Theme toggle
    const toggle = document.getElementById('rkThemeToggle');
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)');
    const saved = localStorage.getItem('rk-theme');
    if (saved === 'dark' || (!saved && prefersDark.matches)) {
        document.documentElement.setAttribute('data-theme', 'dark');
        toggle.checked = true;
    }
    toggle.addEventListener('change', function() {
        if (this.checked) {
            document.documentElement.setAttribute('data-theme', 'dark');
            localStorage.setItem('rk-theme', 'dark');
        } else {
            document.documentElement.setAttribute('data-theme', 'light');
            localStorage.setItem('rk-theme', 'light');
        }
    });

    // Tab switching
    `$(document).on('click', '.rk-tab', function() {
        const target = `$(this).data('target');
        `$(this).closest('.rk-tabs').find('.rk-tab').removeClass('active');
        `$(this).addClass('active');
        `$(this).closest('.rk-tabs').siblings('.rk-panel').removeClass('active');
        `$('#' + target).addClass('active');
    });
});

// DataTable helper
function initRKTable(selector, extraOptions) {
    var defaults = {
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
        responsive: true,
        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],
        order: [[0, 'asc']],
        language: {
            search: "<i class='fas fa-search'></i> _INPUT_",
            searchPlaceholder: 'Search records...',
            lengthMenu: 'Show _MENU_ entries',
            info: 'Showing _START_ to _END_ of _TOTAL_ entries',
            paginate: {
                first: "<i class='fas fa-angle-double-left'></i>",
                last: "<i class='fas fa-angle-double-right'></i>",
                next: "<i class='fas fa-angle-right'></i>",
                previous: "<i class='fas fa-angle-left'></i>"
            }
        }
    };
    if (extraOptions) { `$.extend(true, defaults, extraOptions); }
    return `$(selector).DataTable(defaults);
}
</script>

</body>
</html>
"@

    return $html
}
