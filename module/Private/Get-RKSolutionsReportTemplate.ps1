function Get-RKSolutionsReportTemplate {
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
        [string[]]$Tags = @(),

        [Parameter(Mandatory = $false)]
        [string]$StatsClass = ''
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
        $statsClasses = if ($StatsClass) { "rk-stats $StatsClass" } else { "rk-stats" }
        $statsHtml = @"
        <div class="$statsClasses">
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
<link href="https://fonts.googleapis.com/css2?family=Geist:wght@400;500;600;700&family=Geist+Mono:wght@400;500;600&display=swap" rel="stylesheet">
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
   Brand: Carbon Ember
   ======================================== */

/* --- Light Theme (default) --- */
:root {
    --bg-base: #fafafa;
    --bg-elevated: #ffffff;
    --bg-warm: #f5f5f5;
    --bg-deep: #ededed;
    --border: #e5e5e5;
    --border-dashed: #d4d4d4;
    --text: #0a0a0a;
    --text-body: #171717;
    --text-muted: #737373;
    --text-dim: #a3a3a3;
    --accent: #ea580c;
    --accent-hover: #f97316;
    --accent-soft: #fff7ed;
    --success: #16a34a;
    --warn: #ca8a04;
    --error: #dc2626;
    --tile-rust: #ea580c;
    --tile-olive: #16a34a;
    --tile-steel: #0284c7;
    --tile-rose: #dc2626;
    --tile-amber: #ca8a04;
    --tile-violet: #9333ea;
    --tile-teal: #0891b2;
    --tile-slate: #475569;
    --toggle-bg: #d4d4d4;
    --input-bg: #ffffff;
    --input-border: #e5e5e5;
    --input-color: #171717;
    --button-bg: #f5f5f5;
    --button-color: #171717;
    --button-border: #e5e5e5;
    --button-hover-bg: #ededed;
}

/* --- Dark Theme (Carbon) --- */
[data-theme="dark"] {
    --bg-base: #0a0a0a;
    --bg-elevated: #141414;
    --bg-warm: #1c1c1c;
    --bg-deep: #0e0e0e;
    --border: #262626;
    --border-dashed: #333333;
    --text: #fafafa;
    --text-body: #e5e5e5;
    --text-muted: #a3a3a3;
    --text-dim: #525252;
    --accent: #fb923c;
    --accent-hover: #fdba74;
    --accent-soft: rgba(251,146,60,0.06);
    --success: #4ade80;
    --warn: #facc15;
    --error: #f87171;
    --tile-rust-bg: #431407;
    --tile-rust-border: #7c2d1233;
    --tile-rust-text: #fdba74;
    --tile-rust-eyebrow: #fb923c;
    --tile-rust-caption: #c2410c;
    --tile-olive-bg: #14532d;
    --tile-olive-border: #16653433;
    --tile-olive-text: #86efac;
    --tile-olive-eyebrow: #4ade80;
    --tile-olive-caption: #15803d;
    --tile-steel-bg: #0c4a6e;
    --tile-steel-border: #07598533;
    --tile-steel-text: #7dd3fc;
    --tile-steel-eyebrow: #38bdf8;
    --tile-steel-caption: #0369a1;
    --tile-rose-bg: #7f1d1d;
    --tile-rose-border: #991b1b33;
    --tile-rose-text: #fca5a5;
    --tile-rose-eyebrow: #f87171;
    --tile-rose-caption: #b91c1c;
    --tile-amber-bg: #78350f;
    --tile-amber-border: #92400e33;
    --tile-amber-text: #fde68a;
    --tile-amber-eyebrow: #fbbf24;
    --tile-amber-caption: #b45309;
    --tile-violet-bg: #3b0764;
    --tile-violet-border: #581c8733;
    --tile-violet-text: #c4b5fd;
    --tile-violet-eyebrow: #a78bfa;
    --tile-violet-caption: #7c3aed;
    --tile-teal-bg: #164e63;
    --tile-teal-border: #155e7533;
    --tile-teal-text: #67e8f9;
    --tile-teal-eyebrow: #22d3ee;
    --tile-teal-caption: #0891b2;
    --tile-slate-bg: #1e293b;
    --tile-slate-border: #33415533;
    --tile-slate-text: #cbd5e1;
    --tile-slate-eyebrow: #94a3b8;
    --tile-slate-caption: #64748b;
    --toggle-bg: #404040;
    --input-bg: #141414;
    --input-border: #262626;
    --input-color: #e5e5e5;
    --button-bg: #1c1c1c;
    --button-color: #e5e5e5;
    --button-border: #262626;
    --button-hover-bg: #262626;
}

/* --- Base --- */
*, *::before, *::after { box-sizing: border-box; }

body {
    font-family: 'Geist', -apple-system, BlinkMacSystemFont, system-ui, sans-serif;
    margin: 0; padding: 0;
    background-color: var(--bg-base);
    color: var(--text-body);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    transition: background-color 0.3s ease, color 0.3s ease;
    font-size: 14px;
    line-height: 1.6;
    -webkit-font-smoothing: antialiased;
}

.rk-container {
    position: relative;
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
    font-family: 'Geist Mono', ui-monospace, monospace;
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
    font-family: 'Geist Mono', ui-monospace, monospace;
    font-size: 0.78rem;
    color: var(--text-muted);
    margin-bottom: 24px;
    transition: background-color 0.3s, border-color 0.3s;
}
.rk-breadcrumb .rk-arrow { color: var(--accent); }

/* --- Eyebrow --- */
.rk-eyebrow {
    font-family: 'Geist Mono', ui-monospace, monospace;
    font-size: 0.68rem;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 8px;
}

/* --- Title --- */
.rk-title {
    font-family: 'Geist', -apple-system, sans-serif;
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
    font-family: 'Geist', -apple-system, sans-serif;
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
    font-family: 'Geist Mono', ui-monospace, monospace;
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
.rk-stats.rk-stats-5 { grid-template-columns: repeat(5, 1fr); }
.rk-stat-tile {
    border-radius: 14px;
    padding: 20px 16px;
    text-align: center;
    transition: transform 0.2s, background-color 0.3s, border-color 0.3s;
}
.rk-stat-tile:hover { transform: translateY(-3px); }
.rk-stat-eyebrow {
    font-family: 'Geist Mono', ui-monospace, monospace;
    font-size: 0.62rem;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    margin-bottom: 6px;
}
.rk-stat-number {
    font-family: 'Geist', -apple-system, sans-serif;
    font-size: 2.2rem;
    font-weight: 800;
    line-height: 1.1;
}
.rk-stat-caption {
    font-family: 'Geist', -apple-system, sans-serif;
    font-size: 0.72rem;
    margin-top: 4px;
}

/* Light mode tiles: gradient bg, white text */
.rk-stat-tile.t-rust { background: linear-gradient(135deg, #ea580c, #f97316); color: #fff; }
.rk-stat-tile.t-rust .rk-stat-eyebrow { opacity: 0.92; }
.rk-stat-tile.t-rust .rk-stat-caption { opacity: 0.85; }
.rk-stat-tile.t-olive { background: linear-gradient(135deg, #16a34a, #22c55e); color: #fff; }
.rk-stat-tile.t-olive .rk-stat-eyebrow { opacity: 0.92; }
.rk-stat-tile.t-olive .rk-stat-caption { opacity: 0.85; }
.rk-stat-tile.t-steel { background: linear-gradient(135deg, #0284c7, #0ea5e9); color: #fff; }
.rk-stat-tile.t-steel .rk-stat-eyebrow { opacity: 0.92; }
.rk-stat-tile.t-steel .rk-stat-caption { opacity: 0.85; }
.rk-stat-tile.t-rose { background: linear-gradient(135deg, #dc2626, #ef4444); color: #fff; }
.rk-stat-tile.t-rose .rk-stat-eyebrow { opacity: 0.92; }
.rk-stat-tile.t-rose .rk-stat-caption { opacity: 0.85; }
.rk-stat-tile.t-amber { background: linear-gradient(135deg, #ca8a04, #eab308); color: #fff; }
.rk-stat-tile.t-amber .rk-stat-eyebrow { opacity: 0.92; }
.rk-stat-tile.t-amber .rk-stat-caption { opacity: 0.85; }
.rk-stat-tile.t-violet { background: linear-gradient(135deg, #9333ea, #a855f7); color: #fff; }
.rk-stat-tile.t-violet .rk-stat-eyebrow { opacity: 0.92; }
.rk-stat-tile.t-violet .rk-stat-caption { opacity: 0.85; }
.rk-stat-tile.t-teal { background: linear-gradient(135deg, #0891b2, #06b6d4); color: #fff; }
.rk-stat-tile.t-teal .rk-stat-eyebrow { opacity: 0.92; }
.rk-stat-tile.t-teal .rk-stat-caption { opacity: 0.85; }
.rk-stat-tile.t-slate { background: linear-gradient(135deg, #475569, #64748b); color: #fff; }
.rk-stat-tile.t-slate .rk-stat-eyebrow { opacity: 0.92; }
.rk-stat-tile.t-slate .rk-stat-caption { opacity: 0.85; }

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
[data-theme="dark"] .rk-stat-tile.t-amber {
    background: linear-gradient(135deg, var(--tile-amber-bg), #483808); border: 1px solid var(--tile-amber-border); color: var(--tile-amber-text);
}
[data-theme="dark"] .rk-stat-tile.t-amber .rk-stat-eyebrow { color: var(--tile-amber-eyebrow); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-amber .rk-stat-caption { color: var(--tile-amber-caption); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-violet {
    background: linear-gradient(135deg, var(--tile-violet-bg), #382048); border: 1px solid var(--tile-violet-border); color: var(--tile-violet-text);
}
[data-theme="dark"] .rk-stat-tile.t-violet .rk-stat-eyebrow { color: var(--tile-violet-eyebrow); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-violet .rk-stat-caption { color: var(--tile-violet-caption); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-teal {
    background: linear-gradient(135deg, var(--tile-teal-bg), #104048); border: 1px solid var(--tile-teal-border); color: var(--tile-teal-text);
}
[data-theme="dark"] .rk-stat-tile.t-teal .rk-stat-eyebrow { color: var(--tile-teal-eyebrow); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-teal .rk-stat-caption { color: var(--tile-teal-caption); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-slate {
    background: linear-gradient(135deg, var(--tile-slate-bg), #283040); border: 1px solid var(--tile-slate-border); color: var(--tile-slate-text);
}
[data-theme="dark"] .rk-stat-tile.t-slate .rk-stat-eyebrow { color: var(--tile-slate-eyebrow); opacity: 1; }
[data-theme="dark"] .rk-stat-tile.t-slate .rk-stat-caption { color: var(--tile-slate-caption); opacity: 1; }

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
    font-family: 'Geist', -apple-system, sans-serif;
    font-size: 0.8rem;
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
    font-family: 'Geist Mono', ui-monospace, monospace;
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
    font-family: 'Geist Mono', ui-monospace, monospace;
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
.rk-header-buttons { display: flex; gap: 6px; margin-left: auto; align-items: center; }
.rk-header-buttons .dt-button {
    font-family: 'Geist', -apple-system, sans-serif !important;
    font-size: 0.68rem !important;
    font-weight: 500 !important;
    padding: 4px 10px !important;
    border-radius: 6px !important;
    background: transparent !important;
    border: 1px solid var(--border) !important;
    color: var(--text-muted) !important;
    transition: all 0.15s !important;
    letter-spacing: 0.02em;
}
.rk-header-buttons .dt-button:hover {
    background: var(--accent-soft) !important;
    border-color: var(--accent) !important;
    color: var(--accent) !important;
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
    font-family: 'Geist Mono', ui-monospace, monospace;
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
    font-family: 'Geist', -apple-system, sans-serif;
    transition: background-color 0.3s, color 0.3s, border-color 0.3s;
}
.table tbody td.rk-mono {
    font-family: 'Geist Mono', ui-monospace, monospace;
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
    font-family: 'Geist Mono', ui-monospace, monospace;
    font-size: 0.65rem;
    letter-spacing: 0.05em;
    text-transform: uppercase;
    padding: 3px 8px;
    border-radius: 4px;
    font-weight: 500;
}
.rk-badge-ok { background: rgba(22,163,74,0.1); color: var(--success); }
.rk-badge-warn { background: rgba(202,138,4,0.1); color: var(--warn); }
.rk-badge-error { background: rgba(220,38,38,0.1); color: var(--error); }
.rk-badge-na { background: var(--bg-warm); color: var(--text-muted); }
.rk-badge-accent { background: rgba(234,88,12,0.1); color: var(--accent); }

[data-theme="dark"] .rk-badge-ok { background: rgba(74,222,128,0.1); color: #4ade80; }
[data-theme="dark"] .rk-badge-warn { background: rgba(250,204,21,0.1); color: #facc15; }
[data-theme="dark"] .rk-badge-error { background: rgba(248,113,113,0.1); color: #f87171; }
[data-theme="dark"] .rk-badge-na { background: rgba(82,82,82,0.2); color: #a3a3a3; }
[data-theme="dark"] .rk-badge-accent { background: rgba(251,146,60,0.1); color: #fb923c; }

/* --- Report Tabs (pill switcher) --- */
.rk-tabs {
    display: grid;
    grid-auto-columns: 1fr;
    grid-auto-flow: column;
    gap: 2px;
    margin-bottom: 20px;
    background: var(--bg-warm);
    border-radius: 10px;
    padding: 3px;
}
.rk-tab {
    text-align: center;
    font-family: 'Geist', -apple-system, sans-serif;
    font-size: 0.8rem;
    padding: 8px 16px;
    border-radius: 8px;
    cursor: pointer;
    color: var(--text-muted);
    background: transparent;
    border: none;
    transition: all 0.2s;
    letter-spacing: 0;
    font-weight: 500;
}
.rk-tab:hover { color: var(--text-body); }
.rk-tab.active {
    color: var(--text);
    background: var(--bg-elevated);
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    font-weight: 600;
}
[data-theme="dark"] .rk-tab.active { box-shadow: 0 1px 3px rgba(0,0,0,0.3); }
.rk-panel { visibility: hidden; height: 0; overflow: hidden; }
.rk-panel.active { visibility: visible; height: auto; overflow: visible; }
.rk-panel .dataTables_wrapper,
.rk-panel table.dataTable { width: 100% !important; }

/* --- Show All Toggle --- */
.rk-show-all {
    display: flex;
    align-items: center;
    gap: 10px;
    font-family: 'Geist Mono', ui-monospace, monospace;
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
    font-family: 'Geist Mono', ui-monospace, monospace;
    font-size: 0.78rem;
}

.dataTables_wrapper .dataTables_paginate .paginate_button,
.dataTables_wrapper .dataTables_paginate .page-link {
    padding: 0.3em 0.8em;
    border-radius: 6px;
    margin: 0 2px;
    color: var(--button-color) !important;
    border: 1px solid var(--button-border) !important;
    background-color: var(--button-bg) !important;
    font-family: 'Geist Mono', ui-monospace, monospace;
    font-size: 0.75rem;
}
.dataTables_wrapper .dataTables_paginate .paginate_button.current,
.dataTables_wrapper .dataTables_paginate .paginate_button.current:hover,
.dataTables_wrapper .dataTables_paginate .page-item.active .page-link {
    background: var(--accent) !important;
    border-color: var(--accent) !important;
    color: white !important;
}
.dataTables_wrapper .dataTables_paginate .paginate_button:hover,
.dataTables_wrapper .dataTables_paginate .page-link:hover {
    background: var(--button-hover-bg) !important;
    border-color: var(--border) !important;
    color: var(--button-color) !important;
}
.dataTables_wrapper .dataTables_paginate .paginate_button.disabled,
.dataTables_wrapper .dataTables_paginate .paginate_button.disabled:hover,
.dataTables_wrapper .dataTables_paginate .page-item.disabled .page-link {
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
    font-family: 'Geist', -apple-system, sans-serif;
    font-size: 0.78rem;
}
div.dt-buttons .dt-button:hover {
    background-color: var(--button-hover-bg) !important;
    color: var(--button-color) !important;
}

.form-select, .form-control {
    background-color: var(--input-bg) !important;
    color: var(--input-color) !important;
    border-color: var(--input-border) !important;
    font-family: 'Geist', -apple-system, sans-serif;
    font-size: 0.8rem;
}
.form-label { color: var(--text-body); font-family: 'Geist', -apple-system, sans-serif; font-size: 0.8rem; }

.btn-outline-secondary {
    color: var(--text-body);
    border-color: var(--border);
    font-family: 'Geist', -apple-system, sans-serif;
    font-size: 0.8rem;
}
.btn-outline-secondary:hover { background-color: var(--bg-warm); color: var(--text-body); }

/* --- Filter Tags --- */
.rk-filter-tags { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 8px; }
.rk-filter-tag {
    background-color: var(--bg-warm);
    padding: 4px 12px;
    border-radius: 16px;
    font-size: 0.78rem;
    font-family: 'Geist Mono', ui-monospace, monospace;
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
    font-family: 'Geist Mono', ui-monospace, monospace;
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
        `$('#' + target).find('table.dataTable').each(function() {
            `$(this).DataTable().columns.adjust();
        });
    });
});

// DataTable helper
function initRKTable(selector, extraOptions) {
    var defaults = {
        dom: 'frtip',
        buttons: [
            {
                extend: 'collection',
                text: '<i class="fas fa-download"></i> Export',
                buttons: [
                    { extend: 'excel', text: '<i class="fas fa-file-excel"></i> Excel', exportOptions: { columns: ':visible' } },
                    { extend: 'csv', text: '<i class="fas fa-file-csv"></i> CSV', exportOptions: { columns: ':visible' } },
                    { extend: 'print', text: '<i class="fas fa-print"></i> Print', exportOptions: { columns: ':visible' } }
                ]
            },
            { extend: 'colvis', text: '<i class="fas fa-columns"></i> Columns' }
        ],
        paging: true,
        searching: true,
        ordering: true,
        info: true,
        responsive: false,
        autoWidth: true,
        initComplete: function() {
            var table = this.api().table();
            var headerCells = `$(table.header()).find('th');
            headerCells.each(function() {
                `$(this).css('width', `$(this).outerWidth() + 'px');
            });
            `$(table.node()).css('table-layout', 'fixed');
            // Move buttons into the card header (next to show-all toggle)
            var btnContainer = table.buttons().container();
            var cardHeader = `$(table.node()).closest('.rk-card').find('.rk-card-header');
            if (cardHeader.length) {
                btnContainer.addClass('rk-header-buttons');
                cardHeader.find('.rk-show-all').before(btnContainer);
            }
        },
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
