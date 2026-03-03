#Requires -Version 7.0
<#
.SYNOPSIS
    HTML Report Generator for M365 Tenant Discovery
.DESCRIPTION
    Generates professional HTML reports including:
    - IT Detailed Technical Report
    - Executive Summary Report
    Both with AI-powered analysis integration.
.NOTES
    Author: AI Migration Expert
    Version: 1.0.0
    Target: PowerShell 7.x
#>

#region HTML Templates
function Get-HTMLHeader {
    <#
    .SYNOPSIS
        Returns the HTML header with CSS styling
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $false)]
        [ValidateSet("IT", "Executive")]
        [string]$ReportType = "IT"
    )

    # Vibrant accents on pure black/gray background
    $primaryColor = if ($ReportType -eq "Executive") { "#e91e8c" } else { "#00d4ff" }
    $accentColor = if ($ReportType -eq "Executive") { "#ff6eb4" } else { "#67e8f9" }
    $gradientEnd = if ($ReportType -eq "Executive") { "#9333ea" } else { "#0ea5e9" }

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$Title</title>
    <style>
        :root {
            /* Vibrant accent colors */
            --primary: $primaryColor;
            --primary-light: $accentColor;
            --gradient-end: $gradientEnd;

            /* Status colors - vibrant */
            --success: #10b981;
            --success-light: #34d399;
            --warning: #f59e0b;
            --warning-light: #fbbf24;
            --danger: #ef4444;
            --danger-light: #f87171;
            --info: #3b82f6;

            /* Pure black & neutral gray palette */
            --bg-body: #000000;
            --bg-card: #0a0a0a;
            --bg-card-hover: #111111;
            --bg-elevated: #141414;
            --bg-input: #0d0d0d;
            --bg-subtle: #1a1a1a;

            /* Borders - subtle grays */
            --border-subtle: rgba(255, 255, 255, 0.04);
            --border-default: rgba(255, 255, 255, 0.08);
            --border-strong: rgba(255, 255, 255, 0.12);

            /* Text - pure white to grays */
            --text-primary: #ffffff;
            --text-secondary: #a1a1a1;
            --text-muted: #6b6b6b;
            --text-faint: #404040;

            /* Shadows */
            --shadow-sm: 0 2px 4px rgba(0,0,0,0.4);
            --shadow-md: 0 4px 12px rgba(0,0,0,0.5);
            --shadow-lg: 0 8px 24px rgba(0,0,0,0.6);

            /* Card radius */
            --radius-lg: 20px;
            --radius-md: 14px;
            --radius-sm: 10px;
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }

        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            font-size: 14px;
            line-height: 1.5;
            color: var(--text-primary);
            background: var(--bg-body);
            min-height: 100vh;
            -webkit-font-smoothing: antialiased;
        }

        /* Subtle ambient glow - pink and cyan */
        body::before {
            content: '';
            position: fixed;
            inset: 0;
            background:
                radial-gradient(ellipse 50% 35% at 15% 5%, rgba(233, 30, 140, 0.07) 0%, transparent 50%),
                radial-gradient(ellipse 40% 40% at 85% 95%, rgba(0, 212, 255, 0.05) 0%, transparent 50%);
            pointer-events: none;
            z-index: 0;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 24px;
            position: relative;
            z-index: 1;
        }

        /* === HEADER === */
        .report-header {
            background: var(--bg-card);
            border: 1px solid var(--border-default);
            padding: 36px 32px;
            margin-bottom: 24px;
            border-radius: var(--radius-lg);
            position: relative;
        }

        .report-header::before {
            content: '';
            position: absolute;
            top: 0;
            left: 24px;
            right: 24px;
            height: 3px;
            background: linear-gradient(90deg, var(--primary), var(--gradient-end));
            border-radius: 0 0 2px 2px;
        }

        .report-header h1 {
            font-size: 1.75rem;
            font-weight: 700;
            letter-spacing: -0.01em;
            margin-bottom: 4px;
            color: var(--text-primary);
        }

        .report-header .subtitle {
            font-size: 0.95rem;
            color: var(--text-secondary);
        }

        .report-header .meta {
            margin-top: 20px;
            font-size: 0.8rem;
            color: var(--text-muted);
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
        }

        .report-header .meta strong {
            color: var(--text-secondary);
        }

        /* === KPI CARDS === */
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 16px;
            margin-bottom: 24px;
        }

        @media (max-width: 1100px) {
            .summary-grid { grid-template-columns: repeat(2, 1fr); }
        }

        @media (max-width: 600px) {
            .summary-grid { grid-template-columns: 1fr; }
        }

        .summary-card {
            background: var(--bg-card);
            border: 1px solid var(--border-default);
            padding: 24px 20px;
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow-lg);
            text-align: center;
            position: relative;
        }

        /* Top accent bar */
        .summary-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 16px;
            right: 16px;
            height: 3px;
            border-radius: 0 0 2px 2px;
            background: linear-gradient(90deg, var(--primary), var(--gradient-end));
        }

        .summary-card .value {
            font-size: 2.25rem;
            font-weight: 700;
            line-height: 1;
            color: var(--text-primary);
            margin-bottom: 8px;
        }

        .summary-card .label {
            color: var(--text-muted);
            font-size: 0.7rem;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.08em;
        }

        /* Severity card variants */
        .summary-card.critical::before { background: var(--danger); }
        .summary-card.critical .value { color: var(--danger-light); }
        .summary-card.high::before { background: #f97316; }
        .summary-card.high .value { color: #fb923c; }
        .summary-card.medium::before { background: var(--warning); }
        .summary-card.medium .value { color: var(--warning-light); }
        .summary-card.low::before { background: var(--success); }
        .summary-card.low .value { color: var(--success-light); }
        .summary-card.info::before { background: var(--primary); }

        /* === RISK METER === */
        .risk-meter {
            background: var(--bg-card);
            border: 1px solid var(--border-default);
            padding: 28px;
            border-radius: var(--radius-lg);
            margin-bottom: 24px;
        }

        .risk-meter h2 {
            font-size: 1rem;
            font-weight: 600;
            color: var(--text-primary);
            margin-bottom: 16px;
        }

        .meter-container {
            background: var(--bg-elevated);
            border-radius: var(--radius-sm);
            height: 16px;
            overflow: hidden;
            margin-bottom: 14px;
        }

        .meter-fill {
            height: 100%;
            border-radius: var(--radius-sm);
        }

        .meter-fill.low { background: linear-gradient(90deg, #10b981, #34d399); }
        .meter-fill.medium { background: linear-gradient(90deg, #f59e0b, #fbbf24); }
        .meter-fill.high { background: linear-gradient(90deg, #f97316, #fb923c); }
        .meter-fill.critical { background: linear-gradient(90deg, #ef4444, #f87171); }

        .risk-meter p {
            color: var(--text-secondary);
            font-size: 0.85rem;
            margin-top: 6px;
        }

        .risk-meter p strong {
            color: var(--text-primary);
        }

        /* === SECTIONS === */
        .section {
            background: var(--bg-card);
            border: 1px solid var(--border-default);
            padding: 28px;
            border-radius: var(--radius-lg);
            margin-bottom: 20px;
        }

        .section h2 {
            font-size: 1.1rem;
            font-weight: 600;
            color: var(--text-primary);
            margin-bottom: 20px;
            padding-bottom: 14px;
            border-bottom: 1px solid var(--border-subtle);
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .section h2::before {
            content: '';
            width: 4px;
            height: 20px;
            background: linear-gradient(180deg, var(--primary), var(--gradient-end));
            border-radius: 2px;
        }

        .section h3 {
            font-size: 0.95rem;
            font-weight: 600;
            color: var(--primary-light);
            margin: 24px 0 14px 0;
        }

        /* === TABLES === */
        table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            margin: 14px 0;
            font-size: 0.85rem;
        }

        th, td {
            padding: 12px 14px;
            text-align: left;
        }

        th {
            background: var(--bg-elevated);
            font-weight: 600;
            color: var(--text-secondary);
            text-transform: uppercase;
            font-size: 0.65rem;
            letter-spacing: 0.06em;
        }

        th:first-child { border-radius: var(--radius-sm) 0 0 0; }
        th:last-child { border-radius: 0 var(--radius-sm) 0 0; }

        td {
            background: var(--bg-card-hover);
            border-bottom: 1px solid var(--border-subtle);
            color: var(--text-secondary);
        }

        tr:last-child td:first-child { border-radius: 0 0 0 var(--radius-sm); }
        tr:last-child td:last-child { border-radius: 0 0 var(--radius-sm) 0; }
        tr:last-child td { border-bottom: none; }

        /* === BADGES === */
        .badge {
            display: inline-flex;
            align-items: center;
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 0.65rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }

        .badge.critical { background: rgba(239, 68, 68, 0.12); color: #f87171; }
        .badge.high { background: rgba(249, 115, 22, 0.12); color: #fb923c; }
        .badge.medium { background: rgba(245, 158, 11, 0.12); color: #fbbf24; }
        .badge.low { background: rgba(16, 185, 129, 0.12); color: #34d399; }
        .badge.info { background: rgba(59, 130, 246, 0.12); color: #60a5fa; }

        /* === ISSUE CARDS === */
        .gotcha-card {
            background: var(--bg-card-hover);
            border: 1px solid var(--border-subtle);
            border-left: 4px solid var(--border-default);
            padding: 20px;
            margin: 14px 0;
            border-radius: 0 var(--radius-md) var(--radius-md) 0;
        }

        .gotcha-card.critical { border-left-color: var(--danger); background: rgba(239, 68, 68, 0.03); }
        .gotcha-card.high { border-left-color: #f97316; background: rgba(249, 115, 22, 0.03); }
        .gotcha-card.medium { border-left-color: var(--warning); background: rgba(245, 158, 11, 0.03); }
        .gotcha-card.low { border-left-color: var(--success); background: rgba(16, 185, 129, 0.03); }

        .gotcha-card h4 {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 10px;
            font-size: 1rem;
            font-weight: 600;
            color: var(--text-primary);
        }

        .gotcha-card .description {
            margin-bottom: 14px;
            color: var(--text-secondary);
            line-height: 1.6;
            font-size: 0.85rem;
        }

        .gotcha-card .recommendation {
            background: var(--bg-elevated);
            padding: 12px 16px;
            border-radius: var(--radius-sm);
            font-size: 0.8rem;
            color: var(--text-secondary);
        }

        .gotcha-card .recommendation strong {
            color: var(--primary);
        }

        .gotcha-card .remediation-steps {
            margin-top: 14px;
            padding: 16px;
            background: var(--bg-card);
            border-radius: var(--radius-sm);
            border: 1px solid var(--border-subtle);
        }

        .gotcha-card .remediation-steps strong {
            color: var(--primary-light);
            font-size: 0.8rem;
            display: block;
            margin-bottom: 10px;
        }

        .gotcha-card .steps-list {
            margin: 0;
            padding-left: 18px;
            color: var(--text-secondary);
            font-size: 0.8rem;
        }

        .gotcha-card .steps-list li {
            margin: 6px 0;
            line-height: 1.5;
        }

        .gotcha-card .steps-list li::marker {
            color: var(--primary);
        }

        .gotcha-card .steps-list code {
            background: var(--bg-elevated);
            padding: 2px 5px;
            border-radius: 3px;
            font-family: 'SF Mono', 'Consolas', monospace;
            font-size: 0.75rem;
            color: var(--primary-light);
        }

        .gotcha-card .remediation-meta {
            margin-top: 10px;
            display: flex;
            flex-wrap: wrap;
            gap: 14px;
            padding: 10px 12px;
            background: var(--bg-elevated);
            border-radius: 6px;
            font-size: 0.75rem;
        }

        .gotcha-card .meta-item {
            color: var(--text-muted);
        }

        .gotcha-card .meta-item strong {
            color: var(--primary);
            margin-right: 4px;
        }

        .gotcha-card .prerequisites {
            margin-top: 10px;
            padding: 10px 12px;
            background: rgba(245, 158, 11, 0.06);
            border-radius: 6px;
            font-size: 0.75rem;
            border-left: 3px solid var(--warning);
        }

        .gotcha-card .prerequisites strong {
            color: var(--warning);
        }

        /* === AI ANALYSIS === */
        .ai-analysis {
            background: var(--bg-card-hover);
            padding: 24px;
            border-radius: var(--radius-md);
            margin-top: 18px;
            border: 1px solid var(--border-default);
            position: relative;
        }

        .ai-analysis::before {
            content: '';
            position: absolute;
            top: 0;
            left: 20px;
            right: 20px;
            height: 2px;
            background: linear-gradient(90deg, var(--primary), var(--gradient-end));
            border-radius: 0 0 2px 2px;
        }

        .ai-analysis h4 {
            color: var(--text-primary);
            margin-bottom: 14px;
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 0.95rem;
            font-weight: 600;
        }

        .ai-analysis .ai-badge {
            background: linear-gradient(135deg, var(--primary), var(--gradient-end));
            color: white;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 0.6rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        .ai-content {
            line-height: 1.7;
            color: var(--text-secondary);
            font-size: 0.85rem;
        }

        .ai-content h1, .ai-content h2, .ai-content h3 {
            color: var(--text-primary);
            margin-top: 18px;
            margin-bottom: 8px;
        }

        .ai-content ul, .ai-content ol {
            margin-left: 18px;
            margin-bottom: 14px;
        }

        .ai-content li {
            margin-bottom: 6px;
        }

        .ai-content code {
            background: var(--bg-elevated);
            padding: 2px 5px;
            border-radius: 3px;
            font-family: 'SF Mono', 'Consolas', monospace;
            font-size: 0.8em;
            color: var(--primary);
        }

        .ai-content pre {
            background: var(--bg-body);
            color: var(--text-primary);
            padding: 16px;
            border-radius: var(--radius-sm);
            overflow-x: auto;
            margin: 14px 0;
            border: 1px solid var(--border-subtle);
            font-size: 0.8rem;
        }

        .ai-content pre code {
            background: none;
            padding: 0;
            color: inherit;
        }

        /* === CHARTS === */
        .chart-container {
            height: 260px;
            margin: 18px 0;
        }

        .chart-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(340px, 1fr));
            gap: 16px;
            margin: 18px 0;
        }

        .chart-card {
            background: var(--bg-card);
            padding: 20px;
            border-radius: var(--radius-lg);
            border: 1px solid var(--border-default);
        }

        .chart-card h4 {
            color: var(--text-primary);
            font-size: 0.85rem;
            font-weight: 600;
            margin-bottom: 14px;
            text-align: center;
        }

        .chart-wrapper {
            position: relative;
            height: 240px;
        }

        /* === PROGRESS BARS === */
        .progress-bar {
            background: var(--bg-elevated);
            border-radius: 5px;
            height: 14px;
            margin: 8px 0;
            overflow: hidden;
        }

        .progress-fill {
            height: 100%;
            border-radius: 5px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 0.65rem;
            font-weight: 600;
            background: linear-gradient(90deg, var(--primary), var(--gradient-end));
        }

        /* === TIMELINE === */
        .timeline {
            position: relative;
            padding: 16px 0;
        }

        .timeline::before {
            content: '';
            position: absolute;
            left: 16px;
            top: 0;
            bottom: 0;
            width: 2px;
            background: linear-gradient(180deg, var(--primary), var(--gradient-end));
        }

        .timeline-item {
            display: flex;
            margin-bottom: 24px;
            position: relative;
        }

        .timeline-marker {
            width: 34px;
            height: 34px;
            background: linear-gradient(135deg, var(--primary), var(--gradient-end));
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: 600;
            font-size: 0.8rem;
            flex-shrink: 0;
            z-index: 1;
        }

        .timeline-content {
            flex: 1;
            margin-left: 16px;
            padding: 16px;
            background: var(--bg-card-hover);
            border-radius: var(--radius-md);
            border: 1px solid var(--border-subtle);
        }

        .timeline-content h4 {
            color: var(--primary);
            font-size: 0.9rem;
            font-weight: 600;
            margin-bottom: 6px;
        }

        .timeline-content p {
            color: var(--text-secondary);
            font-size: 0.8rem;
        }

        /* === FOOTER === */
        .report-footer {
            text-align: center;
            padding: 28px;
            color: var(--text-muted);
            font-size: 0.75rem;
            border-top: 1px solid var(--border-subtle);
            margin-top: 32px;
        }

        .report-footer p {
            margin: 4px 0;
        }

        .report-footer a {
            color: var(--primary);
            text-decoration: none;
        }

        /* === PRINT === */
        @media print {
            body {
                background: white !important;
                color: #1a1a1a !important;
            }
            body::before { display: none; }
            .container { max-width: 100%; padding: 0; }
            .section, .summary-card, .report-header, .gotcha-card {
                box-shadow: none !important;
                border: 1px solid #ddd !important;
                background: white !important;
            }
            .section h2, .report-header h1 {
                color: #1a1a1a !important;
            }
            .summary-card .value {
                color: #1a1a1a !important;
            }
        }

        /* === COLLAPSIBLE === */
        .collapsible {
            cursor: pointer;
            user-select: none;
        }

        .collapsible:after {
            content: ' +';
            font-size: 0.85em;
            opacity: 0.5;
        }

        .collapsible.active:after {
            content: ' −';
        }

        .collapsible-content {
            display: none;
            overflow: hidden;
        }

        .collapsible-content.show {
            display: block;
        }

        /* === STATUS DOTS === */
        .status-dot {
            display: inline-block;
            width: 8px;
            height: 8px;
            border-radius: 50%;
            margin-right: 6px;
        }

        .status-dot.green { background: #10b981; }
        .status-dot.yellow { background: #f59e0b; }
        .status-dot.orange { background: #f97316; }
        .status-dot.red { background: #ef4444; }

        /* === METRIC ROWS === */
        .metric-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 12px 0;
            border-bottom: 1px solid var(--border-subtle);
        }

        .metric-row:last-child {
            border-bottom: none;
        }

        .metric-row .label {
            color: var(--text-secondary);
            font-size: 0.875rem;
        }

        .metric-row .value {
            color: var(--text-primary);
            font-weight: 600;
            font-size: 0.875rem;
        }

        /* === TABS === */
        .tab-container {
            margin: 20px 0;
        }

        .tab-buttons {
            display: flex;
            gap: 4px;
            margin-bottom: 20px;
            padding: 4px;
            background: var(--bg-elevated);
            border-radius: 10px;
            width: fit-content;
        }

        .tab-button {
            padding: 10px 20px;
            border: none;
            background: transparent;
            cursor: pointer;
            font-size: 0.85rem;
            color: var(--text-muted);
            border-radius: 8px;
            font-family: inherit;
            font-weight: 500;
        }

        .tab-button.active {
            color: white;
            background: linear-gradient(135deg, var(--primary), var(--primary-light));
            font-weight: 600;
        }

        .tab-content {
            display: none;
        }

        .tab-content.active {
            display: block;
        }

        /* === SCROLLBAR === */
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }

        ::-webkit-scrollbar-track {
            background: var(--bg-card);
            border-radius: 4px;
        }

        ::-webkit-scrollbar-thumb {
            background: var(--bg-subtle);
            border-radius: 4px;
        }

        ::-webkit-scrollbar-thumb:hover {
            background: var(--text-muted);
        }

        /* === SELECTION === */
        ::selection {
            background: var(--primary);
            color: white;
        }

        /* === UTILITY === */
        .text-muted { color: var(--text-muted); }
        .text-secondary { color: var(--text-secondary); }
        .text-primary { color: var(--text-primary); }
        .font-mono { font-family: 'SF Mono', 'Consolas', monospace; }
    </style>
    <!-- Chart.js CDN -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
</head>
<body>
<div class="container">
"@
}

function Get-HTMLFooter {
    <#
    .SYNOPSIS
        Returns the HTML footer
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$GeneratedBy = "M365 Tenant Discovery Tool",

        [Parameter(Mandatory = $false)]
        [string]$AIProvider
    )

    return @"
</div>

<footer class="report-footer">
    <p>Generated by $GeneratedBy</p>
    <p>Report generated on $(Get-Date -Format "MMMM dd, yyyy 'at' HH:mm:ss 'UTC'")</p>
    <p>This report contains confidential tenant information. Handle according to your organization's data classification policies.</p>
</footer>

<script>
    // Collapsible sections
    document.querySelectorAll('.collapsible').forEach(function(elem) {
        elem.addEventListener('click', function() {
            this.classList.toggle('active');
            var content = this.nextElementSibling;
            content.classList.toggle('show');
        });
    });
</script>
</body>
</html>
"@
}

function Get-ChartScript {
    <#
    .SYNOPSIS
        Generates Chart.js visualization scripts for the reports
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AnalysisData
    )

    # Extract data for charts
    $criticalCount = $AnalysisData.BySeverity.Critical.Count
    $highCount = $AnalysisData.BySeverity.High.Count
    $mediumCount = $AnalysisData.BySeverity.Medium.Count
    $lowCount = $AnalysisData.BySeverity.Low.Count

    $categoryData = ($AnalysisData.ByCategory.Keys | ForEach-Object {
        "'$_': $($AnalysisData.ByCategory[$_].Count)"
    }) -join ", "

    return @"
<script>
document.addEventListener('DOMContentLoaded', function() {
    // Chart.js global config for dark theme
    Chart.defaults.color = '#9ca3af';
    Chart.defaults.borderColor = 'rgba(255,255,255,0.06)';

    // Severity Distribution Doughnut Chart
    const severityCtx = document.getElementById('severityChart');
    if (severityCtx) {
        new Chart(severityCtx, {
            type: 'doughnut',
            data: {
                labels: ['Critical', 'High', 'Medium', 'Low'],
                datasets: [{
                    data: [$criticalCount, $highCount, $mediumCount, $lowCount],
                    backgroundColor: ['#ef4444', '#f97316', '#eab308', '#22c55e'],
                    borderWidth: 0,
                    hoverOffset: 4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            padding: 16,
                            usePointStyle: true,
                            font: { size: 12 }
                        }
                    },
                    tooltip: {
                        backgroundColor: 'rgba(17, 24, 39, 0.9)',
                        titleColor: '#f9fafb',
                        bodyColor: '#9ca3af',
                        borderColor: 'rgba(255,255,255,0.1)',
                        borderWidth: 1,
                        padding: 12,
                        callbacks: {
                            label: function(context) {
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = Math.round((context.raw / total) * 100);
                                return context.label + ': ' + context.raw + ' (' + percentage + '%)';
                            }
                        }
                    }
                },
                cutout: '65%'
            }
        });
    }

    // Category Distribution Bar Chart
    const categoryCtx = document.getElementById('categoryChart');
    if (categoryCtx) {
        const categoryLabels = Object.keys({$categoryData});
        const categoryValues = Object.values({$categoryData});

        new Chart(categoryCtx, {
            type: 'bar',
            data: {
                labels: categoryLabels,
                datasets: [{
                    label: 'Issues by Category',
                    data: categoryValues,
                    backgroundColor: [
                        'rgba(14, 165, 233, 0.7)',
                        'rgba(168, 85, 247, 0.7)',
                        'rgba(34, 197, 94, 0.7)',
                        'rgba(234, 179, 8, 0.7)',
                        'rgba(59, 130, 246, 0.7)',
                        'rgba(236, 72, 153, 0.7)'
                    ],
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: 'rgba(17, 24, 39, 0.9)',
                        titleColor: '#f9fafb',
                        bodyColor: '#9ca3af',
                        borderColor: 'rgba(255,255,255,0.1)',
                        borderWidth: 1,
                        padding: 12
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: { stepSize: 1 },
                        grid: { color: 'rgba(255,255,255,0.04)' }
                    },
                    x: {
                        ticks: {
                            maxRotation: 45,
                            minRotation: 45
                        },
                        grid: { display: false }
                    }
                }
            }
        });
    }

    // Complexity Score Gauge
    const gaugeCtx = document.getElementById('complexityGauge');
    if (gaugeCtx) {
        const score = parseInt(gaugeCtx.dataset.score) || 0;
        const color = score >= 70 ? '#ef4444' : score >= 50 ? '#f97316' : score >= 30 ? '#eab308' : '#22c55e';

        new Chart(gaugeCtx, {
            type: 'doughnut',
            data: {
                datasets: [{
                    data: [score, 100 - score],
                    backgroundColor: [color, 'rgba(55, 65, 81, 0.5)'],
                    borderWidth: 0,
                    circumference: 180,
                    rotation: 270
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: { enabled: false }
                },
                cutout: '75%'
            }
        });
    }

    // Tab functionality
    document.querySelectorAll('.tab-button').forEach(button => {
        button.addEventListener('click', () => {
            const tabContainer = button.closest('.tab-container');
            tabContainer.querySelectorAll('.tab-button').forEach(b => b.classList.remove('active'));
            tabContainer.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            button.classList.add('active');
            document.getElementById(button.dataset.tab).classList.add('active');
        });
    });
});
</script>
"@
}
#endregion

#region IT Detailed Report
function New-ITDetailedReport {
    <#
    .SYNOPSIS
        Generates the detailed IT technical report
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $true)]
        $AnalysisResults,

        [Parameter(Mandatory = $true)]
        $ComplexityScore,

        [Parameter(Mandatory = $false)]
        $AIAnalysis,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Generating IT Detailed Report..." -Level Info

    $html = Get-HTMLHeader -Title "M365 Tenant Discovery - IT Technical Report" -ReportType "IT"

    # Header Section
    $tenantName = $CollectedData.EntraID.TenantInfo.DisplayName
    $tenantId = $CollectedData.EntraID.TenantInfo.TenantId

    $html += @"
<header class="report-header">
    <h1>M365 Tenant Migration Assessment</h1>
    <p class="subtitle">Technical Discovery Report</p>
    <div class="meta">
        <p><strong>Tenant:</strong> $tenantName</p>
        <p><strong>Tenant ID:</strong> $tenantId</p>
        <p><strong>Assessment Date:</strong> $(Get-Date -Format "MMMM dd, yyyy")</p>
    </div>
</header>
"@

    # Summary Cards
    $html += @"
<div class="summary-grid">
    <div class="summary-card info">
        <div class="value">$($CollectedData.EntraID.Users.Analysis.LicensedUsers)</div>
        <div class="label">Licensed Users</div>
    </div>
    <div class="summary-card info">
        <div class="value">$(($CollectedData.Exchange.Mailboxes.Mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' }).Count)</div>
        <div class="label">Shared Mailboxes</div>
    </div>
    <div class="summary-card info">
        <div class="value">$($CollectedData.SharePoint.Sites.Analysis.SharePointSites)</div>
        <div class="label">SharePoint Sites</div>
    </div>
    <div class="summary-card info">
        <div class="value">$($CollectedData.Teams.Teams.Analysis.TotalTeams)</div>
        <div class="label">Teams</div>
    </div>
    <div class="summary-card critical">
        <div class="value">$($AnalysisResults.BySeverity.Critical.Count)</div>
        <div class="label">Critical Issues</div>
    </div>
    <div class="summary-card high">
        <div class="value">$($AnalysisResults.BySeverity.High.Count)</div>
        <div class="label">High Priority</div>
    </div>
    <div class="summary-card medium">
        <div class="value">$($AnalysisResults.BySeverity.Medium.Count)</div>
        <div class="label">Medium Priority</div>
    </div>
    <div class="summary-card low">
        <div class="value">$($AnalysisResults.BySeverity.Low.Count)</div>
        <div class="label">Low Priority</div>
    </div>
</div>
"@

    # Risk Meter
    $riskClass = $AnalysisResults.RiskLevel.ToLower()
    $html += @"
<div class="risk-meter">
    <h2>Overall Migration Risk Assessment</h2>
    <div class="meter-container">
        <div class="meter-fill $riskClass" style="width: $($AnalysisResults.RiskPercentage)%"></div>
    </div>
    <p><strong>Risk Level:</strong> $($AnalysisResults.RiskLevel) ($($AnalysisResults.RiskPercentage)%)</p>
    <p><strong>Complexity Score:</strong> $($ComplexityScore.TotalScore)/100 - $($ComplexityScore.ComplexityLevel)</p>
</div>
"@

    # Migration Gotchas Section
    $html += @"
<section class="section">
    <h2>Migration Gotchas and Risks</h2>
    <p>Identified $($AnalysisResults.RulesTriggered) potential issues requiring attention.</p>
"@

    # Helper function to build remediation details
    $buildRemediationHtml = {
        param($issue)
        $remediationHtml = ""

        # Remediation Steps
        if ($issue.RemediationSteps -and $issue.RemediationSteps.Count -gt 0) {
            $remediationHtml += "<div class='remediation-steps'><strong>How to Resolve:</strong><ol class='steps-list'>"
            foreach ($step in $issue.RemediationSteps) {
                $escapedStep = ($step -replace '^\d+\.\s*', '')
                $escapedStep = [System.Web.HttpUtility]::HtmlEncode($escapedStep) -replace '`([^`]+)`', '<code>$1</code>'
                $remediationHtml += "<li>$escapedStep</li>"
            }
            $remediationHtml += "</ol></div>"
        }

        # Tools and Effort
        $metaHtml = ""
        if ($issue.Tools -and $issue.Tools.Count -gt 0) {
            $toolsList = ($issue.Tools -join ", ")
            $metaHtml += "<span class='meta-item'><strong>Tools:</strong> $toolsList</span>"
        }
        if ($issue.EstimatedEffort -and $issue.EstimatedEffort -ne "Not estimated") {
            $metaHtml += "<span class='meta-item'><strong>Effort:</strong> $($issue.EstimatedEffort)</span>"
        }
        if ($metaHtml) {
            $remediationHtml += "<div class='remediation-meta'>$metaHtml</div>"
        }

        # Prerequisites
        if ($issue.Prerequisites -and $issue.Prerequisites.Count -gt 0) {
            $prereqList = ($issue.Prerequisites -join ", ")
            $remediationHtml += "<div class='prerequisites'><strong>Prerequisites:</strong> $prereqList</div>"
        }

        return $remediationHtml
    }

    # Critical Issues
    if ($AnalysisResults.BySeverity.Critical.Count -gt 0) {
        $html += "<h3>Critical Issues (Immediate Attention Required)</h3>"
        foreach ($issue in $AnalysisResults.BySeverity.Critical) {
            $remediationDetails = & $buildRemediationHtml $issue
            $html += @"
    <div class="gotcha-card critical">
        <h4><span class="badge critical">CRITICAL</span> $($issue.Name)</h4>
        <p class="description">$($issue.Description)</p>
        <div class="recommendation"><strong>Summary:</strong> $($issue.Recommendation)</div>
        $remediationDetails
    </div>
"@
        }
    }

    # High Priority Issues
    if ($AnalysisResults.BySeverity.High.Count -gt 0) {
        $html += "<h3>High Priority Issues</h3>"
        foreach ($issue in $AnalysisResults.BySeverity.High) {
            $remediationDetails = & $buildRemediationHtml $issue
            $html += @"
    <div class="gotcha-card high">
        <h4><span class="badge high">HIGH</span> $($issue.Name)</h4>
        <p class="description">$($issue.Description)</p>
        <div class="recommendation"><strong>Summary:</strong> $($issue.Recommendation)</div>
        $remediationDetails
    </div>
"@
        }
    }

    # Medium Priority Issues
    if ($AnalysisResults.BySeverity.Medium.Count -gt 0) {
        $html += "<h3 class='collapsible'>Medium Priority Issues ($($AnalysisResults.BySeverity.Medium.Count))</h3>"
        $html += "<div class='collapsible-content'>"
        foreach ($issue in $AnalysisResults.BySeverity.Medium) {
            $remediationDetails = & $buildRemediationHtml $issue
            $html += @"
    <div class="gotcha-card medium">
        <h4><span class="badge medium">MEDIUM</span> $($issue.Name)</h4>
        <p class="description">$($issue.Description)</p>
        <div class="recommendation"><strong>Summary:</strong> $($issue.Recommendation)</div>
        $remediationDetails
    </div>
"@
        }
        $html += "</div>"
    }

    $html += "</section>"

    # Identity Section
    $html += @"
<section class="section">
    <h2>Identity & Access Management</h2>

    <h3>User Analysis</h3>
    <table>
        <tr><th>Metric</th><th>Value</th><th>Migration Impact</th></tr>
        <tr style="background: rgba(6, 182, 212, 0.1);"><td><strong>Licensed Users (Migration Scope)</strong></td><td><strong>$($CollectedData.EntraID.Users.Analysis.LicensedUsers)</strong></td><td>Primary migration population</td></tr>
        <tr><td>Total Directory Users</td><td>$($CollectedData.EntraID.Users.Analysis.TotalUsers)</td><td>Includes unlicensed/service accounts</td></tr>
        <tr><td>Synced Users (Hybrid)</td><td>$($CollectedData.EntraID.Users.Analysis.SyncedUsers)</td><td>Require ImmutableID mapping</td></tr>
        <tr><td>Cloud-Only Users</td><td>$($CollectedData.EntraID.Users.Analysis.CloudOnlyUsers)</td><td>Direct migration possible</td></tr>
        <tr><td>Guest Users</td><td>$($CollectedData.EntraID.Users.Analysis.GuestUsers)</td><td>Not migrated - require reinvitation</td></tr>
        <tr><td>Unlicensed Users</td><td>$($CollectedData.EntraID.Users.Analysis.UnlicensedUsers)</td><td>Review for service accounts</td></tr>
    </table>

    <h3>Group Analysis</h3>
    <table>
        <tr><th>Metric</th><th>Value</th></tr>
        <tr><td>Total Groups</td><td>$($CollectedData.EntraID.Groups.Analysis.TotalGroups)</td></tr>
        <tr><td>M365 Groups</td><td>$($CollectedData.EntraID.Groups.Analysis.M365Groups)</td></tr>
        <tr><td>Security Groups</td><td>$($CollectedData.EntraID.Groups.Analysis.SecurityGroups)</td></tr>
        <tr><td>Distribution Lists</td><td>$($CollectedData.EntraID.Groups.Analysis.DistributionLists)</td></tr>
        <tr><td>Dynamic Groups</td><td>$($CollectedData.EntraID.Groups.Analysis.DynamicGroups)</td></tr>
        <tr><td>Synced Groups</td><td>$($CollectedData.EntraID.Groups.Analysis.SyncedGroups)</td></tr>
    </table>

    <h3>Device Analysis</h3>
    <table>
        <tr><th>Metric</th><th>Value</th></tr>
        <tr><td>Total Devices</td><td>$($CollectedData.EntraID.Devices.Analysis.TotalDevices)</td></tr>
        <tr><td>Hybrid Azure AD Joined</td><td>$($CollectedData.EntraID.Devices.Analysis.HybridJoined)</td></tr>
        <tr><td>Azure AD Joined</td><td>$($CollectedData.EntraID.Devices.Analysis.AzureADJoined)</td></tr>
        <tr><td>Registered</td><td>$($CollectedData.EntraID.Devices.Analysis.Registered)</td></tr>
        <tr><td>Compliant</td><td>$($CollectedData.EntraID.Devices.Analysis.Compliant)</td></tr>
    </table>

    <h3>Conditional Access</h3>
    <table>
        <tr><th>Metric</th><th>Value</th></tr>
        <tr><td>Total Policies</td><td>$($CollectedData.EntraID.ConditionalAccess.Analysis.TotalPolicies)</td></tr>
        <tr><td>Enabled Policies</td><td>$($CollectedData.EntraID.ConditionalAccess.Analysis.EnabledPolicies)</td></tr>
        <tr><td>Report-Only Policies</td><td>$($CollectedData.EntraID.ConditionalAccess.Analysis.ReportOnlyPolicies)</td></tr>
        <tr><td>Named Locations</td><td>$($CollectedData.EntraID.ConditionalAccess.Analysis.NamedLocations)</td></tr>
    </table>
</section>
"@

    # Exchange Section
    $html += @"
<section class="section">
    <h2>Exchange Online</h2>

    <h3>Mailbox Analysis</h3>
    <table>
        <tr><th>Metric</th><th>Value</th></tr>
        <tr><td>Total Mailboxes</td><td>$($CollectedData.Exchange.Mailboxes.Analysis.TotalMailboxes)</td></tr>
        <tr><td>User Mailboxes</td><td>$($CollectedData.Exchange.Mailboxes.Analysis.UserMailboxes)</td></tr>
        <tr><td>Shared Mailboxes</td><td>$($CollectedData.Exchange.Mailboxes.Analysis.SharedMailboxes)</td></tr>
        <tr><td>Room Mailboxes</td><td>$($CollectedData.Exchange.Mailboxes.Analysis.RoomMailboxes)</td></tr>
        <tr><td>Equipment Mailboxes</td><td>$($CollectedData.Exchange.Mailboxes.Analysis.EquipmentMailboxes)</td></tr>
        <tr><td>Archive Enabled</td><td>$($CollectedData.Exchange.Mailboxes.Analysis.ArchiveEnabled)</td></tr>
        <tr><td>Litigation Hold</td><td>$($CollectedData.Exchange.Mailboxes.Analysis.LitigationHold)</td></tr>
    </table>

    <h3>Distribution & Transport</h3>
    <table>
        <tr><th>Metric</th><th>Value</th></tr>
        <tr><td>Distribution Lists</td><td>$($CollectedData.Exchange.DistributionLists.Analysis.TotalDistributionLists)</td></tr>
        <tr><td>Dynamic DLs</td><td>$($CollectedData.Exchange.DistributionLists.Analysis.TotalDynamicDLs)</td></tr>
        <tr><td>Transport Rules</td><td>$($CollectedData.Exchange.TransportConfig.Analysis.TotalTransportRules)</td></tr>
        <tr><td>Inbound Connectors</td><td>$($CollectedData.Exchange.TransportConfig.Analysis.InboundConnectors)</td></tr>
        <tr><td>Outbound Connectors</td><td>$($CollectedData.Exchange.TransportConfig.Analysis.OutboundConnectors)</td></tr>
    </table>
</section>
"@

    # SharePoint Section
    $html += @"
<section class="section">
    <h2>SharePoint & OneDrive</h2>

    <h3>Site Analysis</h3>
    <table>
        <tr><th>Metric</th><th>Value</th></tr>
        <tr><td>SharePoint Sites</td><td>$($CollectedData.SharePoint.Sites.Analysis.SharePointSites)</td></tr>
        <tr><td>OneDrive Sites (All Users)</td><td>$($CollectedData.SharePoint.Sites.Analysis.OneDriveSites)</td></tr>
        <tr><td>OneDrive Sites (Licensed Users)</td><td>$(if ($CollectedData.SharePoint.Sites.Analysis.OneDriveSitesLicensedUsers -ge 0) { $CollectedData.SharePoint.Sites.Analysis.OneDriveSitesLicensedUsers } else { 'N/A' })</td></tr>
        <tr><td>Team Sites</td><td>$($CollectedData.SharePoint.Sites.Analysis.TeamSites)</td></tr>
        <tr><td>Communication Sites</td><td>$($CollectedData.SharePoint.Sites.Analysis.CommunicationSites)</td></tr>
        <tr><td>Classic Sites</td><td>$($CollectedData.SharePoint.Sites.Analysis.ClassicSites)</td></tr>
        <tr><td>Hub Sites</td><td>$($CollectedData.SharePoint.Sites.Analysis.HubSites)</td></tr>
        <tr><td>Total Storage (GB)</td><td>$($CollectedData.SharePoint.Sites.Analysis.TotalStorageUsedGB)</td></tr>
    </table>
</section>
"@

    # Teams Section
    $html += @"
<section class="section">
    <h2>Microsoft Teams</h2>

    <table>
        <tr><th>Metric</th><th>Value</th></tr>
        <tr><td>Total Teams</td><td>$($CollectedData.Teams.Teams.Analysis.TotalTeams)</td></tr>
        <tr><td>Public Teams</td><td>$($CollectedData.Teams.Teams.Analysis.PublicTeams)</td></tr>
        <tr><td>Private Teams</td><td>$($CollectedData.Teams.Teams.Analysis.PrivateTeams)</td></tr>
        <tr><td>Archived Teams</td><td>$($CollectedData.Teams.Teams.Analysis.ArchivedTeams)</td></tr>
        <tr><td>Total Channels</td><td>$($CollectedData.Teams.Teams.Analysis.TotalChannels)</td></tr>
        <tr><td>Private Channels</td><td>$($CollectedData.Teams.Teams.Analysis.PrivateChannels)</td></tr>
        <tr><td>Shared Channels</td><td>$($CollectedData.Teams.Teams.Analysis.SharedChannels)</td></tr>
    </table>
</section>
"@

    # Hybrid Identity Section
    $html += @"
<section class="section">
    <h2>Hybrid Identity Configuration</h2>

    <table>
        <tr><th>Configuration</th><th>Status</th></tr>
        <tr><td>Directory Sync Enabled</td><td>$($CollectedData.HybridIdentity.AADConnect.Configuration.OnPremisesSyncEnabled)</td></tr>
        <tr><td>Federated Domains</td><td>$($CollectedData.HybridIdentity.Federation.Analysis.FederatedDomains)</td></tr>
        <tr><td>Pass-Through Authentication</td><td>$($CollectedData.HybridIdentity.AuthenticationMethods.Analysis.PTAEnabled)</td></tr>
        <tr><td>PTA Agents</td><td>$($CollectedData.HybridIdentity.AuthenticationMethods.Analysis.PTAAgentCount)</td></tr>
        <tr><td>Synced Users</td><td>$($CollectedData.HybridIdentity.SyncObjects.SyncedUserCount)</td></tr>
        <tr><td>Cloud-Only Users</td><td>$($CollectedData.HybridIdentity.SyncObjects.CloudUserCount)</td></tr>
        <tr><td>Hybrid Devices</td><td>$($CollectedData.HybridIdentity.DeviceWriteback.Configuration.HybridJoinedDeviceCount)</td></tr>
    </table>
</section>
"@

    # AI Analysis Section
    if ($AIAnalysis -and $AIAnalysis.Success) {
        # Use proper markdown converter instead of inline regex
        $aiContent = ConvertTo-HTMLFromMarkdown -Markdown $AIAnalysis.Analysis

        $html += '<section class="section"><h2>Deep Analysis</h2><div class="ai-analysis"><h4>Intelligent Migration Assessment</h4><div class="ai-content">' + $aiContent + '</div></div></section>'
    }

    $html += Get-HTMLFooter -GeneratedBy "M365 Tenant Discovery Tool"

    # Write to file
    $html | Out-File -FilePath $OutputPath -Encoding UTF8

    Write-Log -Message "IT Detailed Report generated: $OutputPath" -Level Success

    return $OutputPath
}
#endregion

#region Executive Summary Report
function New-ExecutiveSummaryReport {
    <#
    .SYNOPSIS
        Generates a visually stunning executive KPI dashboard - high-level statistics only
    .DESCRIPTION
        Creates a non-technical, statistics-focused executive summary designed
        for leadership. Contains only high-level metrics and visual indicators.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $true)]
        $AnalysisResults,

        [Parameter(Mandatory = $true)]
        $ComplexityScore,

        [Parameter(Mandatory = $false)]
        $AIExecutiveSummary,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Generating Executive Summary Report..." -Level Info

    # Calculate all stats upfront
    $tenantName = $CollectedData.EntraID.TenantInfo.DisplayName
    $licensedUsers = $CollectedData.EntraID.Users.Analysis.LicensedUsers
    $guestUsers = $CollectedData.EntraID.Users.Analysis.GuestUsers
    $totalTeams = $CollectedData.Teams.Teams.Analysis.TotalTeams
    $sharePointSites = $CollectedData.SharePoint.Sites.Analysis.SharePointSites
    $oneDriveSites = $CollectedData.SharePoint.Sites.Analysis.OneDriveSites
    $oneDriveSitesLicensed = if ($CollectedData.SharePoint.Sites.Analysis.OneDriveSitesLicensedUsers -ge 0) { $CollectedData.SharePoint.Sites.Analysis.OneDriveSitesLicensedUsers } else { $null }
    $totalStorageGB = $CollectedData.SharePoint.Sites.Analysis.TotalStorageUsedGB
    $userMailboxes = ($CollectedData.Exchange.Mailboxes.Mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' }).Count
    $sharedMailboxes = ($CollectedData.Exchange.Mailboxes.Mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' }).Count
    $totalApps = $CollectedData.EntraID.Applications.Analysis.TotalServicePrincipals
    $totalGroups = $CollectedData.EntraID.Groups.Analysis.TotalGroups
    $securityGroups = $CollectedData.EntraID.Groups.Analysis.SecurityGroups
    $m365Groups = $CollectedData.EntraID.Groups.Analysis.Microsoft365Groups
    $totalDevices = if ($CollectedData.EntraID.Devices.Analysis) { $CollectedData.EntraID.Devices.Analysis.TotalDevices } else { 0 }

    # Readiness score (inverse of risk - higher is better for executives)
    $readinessScore = 100 - $AnalysisResults.RiskPercentage
    $readinessLevel = switch ($readinessScore) {
        { $_ -ge 80 } { "Excellent" }
        { $_ -ge 60 } { "Good" }
        { $_ -ge 40 } { "Fair" }
        default { "Needs Attention" }
    }
    $readinessColor = switch ($readinessScore) {
        { $_ -ge 80 } { "#22c55e" }
        { $_ -ge 60 } { "#eab308" }
        { $_ -ge 40 } { "#f97316" }
        default { "#ef4444" }
    }

    # Custom Executive CSS
    $executiveCSS = @"
    <style>
        /* Executive Dashboard Overrides */
        .exec-hero {
            text-align: center;
            padding: 60px 40px;
            margin-bottom: 40px;
            background: linear-gradient(145deg, rgba(168, 85, 247, 0.15) 0%, rgba(14, 165, 233, 0.1) 100%);
            border-radius: 32px;
            border: 1px solid rgba(255,255,255,0.08);
        }

        .exec-hero h1 {
            font-size: 2.75rem;
            font-weight: 700;
            margin-bottom: 8px;
            background: linear-gradient(135deg, #f9fafb 0%, #a855f7 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }

        .exec-hero .org-name {
            font-size: 1.5rem;
            color: #c084fc;
            font-weight: 500;
            margin-bottom: 8px;
        }

        .exec-hero .date {
            color: #6b7280;
            font-size: 0.9rem;
        }

        /* Giant readiness indicator */
        .readiness-hero {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 48px;
            padding: 48px;
            margin-bottom: 40px;
            background: var(--bg-card);
            border-radius: 28px;
            border: 1px solid var(--border-default);
        }

        .readiness-circle {
            position: relative;
            width: 240px;
            height: 240px;
        }

        .readiness-circle svg {
            transform: rotate(-90deg);
            width: 100%;
            height: 100%;
        }

        .readiness-circle .score {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            text-align: center;
        }

        .readiness-circle .score-value {
            font-size: 2.75rem;
            font-weight: 700;
            color: var(--text-primary);
            line-height: 1;
        }

        .readiness-circle .score-label {
            font-size: 0.8rem;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-top: 6px;
        }

        .readiness-details h2 {
            font-size: 2rem;
            font-weight: 600;
            color: var(--text-primary);
            margin-bottom: 8px;
        }

        .readiness-details .status {
            font-size: 1.125rem;
            margin-bottom: 20px;
        }

        .readiness-details .summary-text {
            color: var(--text-secondary);
            font-size: 1rem;
            line-height: 1.7;
            max-width: 400px;
        }

        /* Big KPI Grid */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 24px;
            margin-bottom: 40px;
        }

        @media (max-width: 1200px) {
            .kpi-grid { grid-template-columns: repeat(2, 1fr); }
        }

        .kpi-card {
            background: var(--bg-card);
            border: 1px solid var(--border-default);
            border-radius: 24px;
            padding: 32px 28px;
            text-align: center;
            position: relative;
            overflow: hidden;
        }

        .kpi-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: linear-gradient(90deg, var(--primary), var(--primary-light));
            border-radius: 4px 4px 0 0;
        }

        .kpi-card .kpi-icon {
            width: 48px;
            height: 48px;
            margin: 0 auto 16px;
            background: linear-gradient(135deg, rgba(168, 85, 247, 0.2), rgba(14, 165, 233, 0.2));
            border-radius: 14px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.5rem;
        }

        .kpi-card .kpi-value {
            font-size: 3rem;
            font-weight: 700;
            color: var(--text-primary);
            line-height: 1;
            margin-bottom: 8px;
        }

        .kpi-card .kpi-label {
            font-size: 0.85rem;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.06em;
        }

        /* Secondary stats */
        .stats-panel {
            background: var(--bg-card);
            border: 1px solid var(--border-default);
            border-radius: 24px;
            padding: 32px;
            margin-bottom: 40px;
        }

        .stats-panel h3 {
            font-size: 1.125rem;
            font-weight: 600;
            color: var(--text-primary);
            margin-bottom: 24px;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .stats-panel h3::before {
            content: '';
            width: 4px;
            height: 20px;
            background: linear-gradient(180deg, var(--primary), var(--primary-light));
            border-radius: 2px;
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
            gap: 20px;
        }

        .stat-item {
            text-align: center;
            padding: 20px 16px;
            background: var(--bg-elevated);
            border-radius: 16px;
            border: 1px solid var(--border-subtle);
        }

        .stat-item .stat-value {
            font-size: 2rem;
            font-weight: 700;
            color: var(--text-primary);
            line-height: 1;
            margin-bottom: 6px;
        }

        .stat-item .stat-label {
            font-size: 0.75rem;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        /* Simple timeline for executives */
        .exec-timeline {
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 16px;
            margin-bottom: 40px;
        }

        @media (max-width: 1000px) {
            .exec-timeline { grid-template-columns: 1fr; }
        }

        .phase-card {
            background: var(--bg-card);
            border: 1px solid var(--border-default);
            border-radius: 20px;
            padding: 24px 20px;
            text-align: center;
            position: relative;
        }

        .phase-card.current {
            border-color: var(--primary);
            background: linear-gradient(145deg, rgba(233, 30, 140, 0.08), rgba(0, 212, 255, 0.04));
        }

        .phase-card .phase-num {
            width: 36px;
            height: 36px;
            background: linear-gradient(135deg, var(--primary), var(--primary-light));
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 14px;
            font-weight: 700;
            font-size: 1rem;
            color: white;
        }

        .phase-card h4 {
            font-size: 0.9rem;
            font-weight: 600;
            color: var(--text-primary);
            margin-bottom: 6px;
        }

        .phase-card p {
            font-size: 0.8rem;
            color: var(--text-muted);
            line-height: 1.5;
        }

        .phase-card.current h4 {
            color: var(--primary-light);
        }

        /* Executive footer */
        .exec-footer {
            text-align: center;
            padding: 40px 32px;
            margin-top: 40px;
            border-top: 1px solid var(--border-subtle);
        }

        .exec-footer .confidential {
            font-size: 0.75rem;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.1em;
            margin-bottom: 8px;
        }

        .exec-footer .generated {
            font-size: 0.8rem;
            color: var(--text-muted);
        }
    </style>
"@

    $html = Get-HTMLHeader -Title "Migration Assessment - Executive Summary" -ReportType "Executive"
    $html += $executiveCSS

    # Hero Header
    $html += @"
<div class="exec-hero">
    <h1>Migration Readiness Assessment</h1>
    <div class="org-name">$tenantName</div>
    <div class="date">$(Get-Date -Format "MMMM dd, yyyy")</div>
</div>
"@

    # Readiness Score Hero
    $circumference = 2 * 3.14159 * 100
    $dashOffset = $circumference * (1 - ($readinessScore / 100))

    $html += @"
<div class="readiness-hero">
    <div class="readiness-circle">
        <svg viewBox="0 0 240 240">
            <circle cx="120" cy="120" r="100" fill="none" stroke="rgba(55,65,81,0.5)" stroke-width="10"/>
            <circle cx="120" cy="120" r="100" fill="none" stroke="$readinessColor" stroke-width="10"
                stroke-linecap="round" stroke-dasharray="$circumference" stroke-dashoffset="$dashOffset"/>
        </svg>
        <div class="score">
            <div class="score-value">$readinessScore%</div>
            <div class="score-label">Ready</div>
        </div>
    </div>
    <div class="readiness-details">
        <h2>Migration Readiness</h2>
        <div class="status" style="color: $readinessColor;">$readinessLevel</div>
        <p class="summary-text">
            Your organization has been assessed for Microsoft 365 tenant migration.
            This report provides a high-level overview of your current environment and migration scope.
        </p>
    </div>
</div>
"@

    # Primary KPIs
    $html += @"
<div class="kpi-grid">
    <div class="kpi-card">
        <div class="kpi-icon">👥</div>
        <div class="kpi-value">$licensedUsers</div>
        <div class="kpi-label">Users</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-icon">📧</div>
        <div class="kpi-value">$userMailboxes</div>
        <div class="kpi-label">Mailboxes</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-icon">👥</div>
        <div class="kpi-value">$totalTeams</div>
        <div class="kpi-label">Teams</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-icon">💾</div>
        <div class="kpi-value">${totalStorageGB}<span style="font-size: 1.25rem; font-weight: 400;">GB</span></div>
        <div class="kpi-label">Data Volume</div>
    </div>
</div>
"@

    # Communication & Collaboration Stats
    $html += @"
<div class="stats-panel">
    <h3>Communication & Collaboration</h3>
    <div class="stats-grid">
        <div class="stat-item">
            <div class="stat-value">$userMailboxes</div>
            <div class="stat-label">User Mailboxes</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">$sharedMailboxes</div>
            <div class="stat-label">Shared Mailboxes</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">$totalTeams</div>
            <div class="stat-label">Teams</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">$sharePointSites</div>
            <div class="stat-label">SharePoint Sites</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">$oneDriveSites</div>
            <div class="stat-label">OneDrive Accounts</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">$totalGroups</div>
            <div class="stat-label">Groups</div>
        </div>
    </div>
</div>
"@

    # Identity & Access Stats
    $html += @"
<div class="stats-panel">
    <h3>Identity & Access</h3>
    <div class="stats-grid">
        <div class="stat-item">
            <div class="stat-value">$licensedUsers</div>
            <div class="stat-label">Licensed Users</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">$guestUsers</div>
            <div class="stat-label">Guest Users</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">$securityGroups</div>
            <div class="stat-label">Security Groups</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">$m365Groups</div>
            <div class="stat-label">M365 Groups</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">$totalApps</div>
            <div class="stat-label">Applications</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">$totalDevices</div>
            <div class="stat-label">Devices</div>
        </div>
    </div>
</div>
"@

    # Simple Phase Timeline
    $html += @"
<div class="stats-panel">
    <h3>Migration Journey</h3>
    <div class="exec-timeline">
        <div class="phase-card current">
            <div class="phase-num">1</div>
            <h4>Assessment</h4>
            <p>Current phase complete</p>
        </div>
        <div class="phase-card">
            <div class="phase-num">2</div>
            <h4>Planning</h4>
            <p>Strategy & timeline</p>
        </div>
        <div class="phase-card">
            <div class="phase-num">3</div>
            <h4>Preparation</h4>
            <p>Environment setup</p>
        </div>
        <div class="phase-card">
            <div class="phase-num">4</div>
            <h4>Migration</h4>
            <p>Phased data move</p>
        </div>
        <div class="phase-card">
            <div class="phase-num">5</div>
            <h4>Completion</h4>
            <p>Validation & close</p>
        </div>
    </div>
</div>
"@

    # AI Executive Summary (if available) - simplified
    if ($AIExecutiveSummary -and $AIExecutiveSummary.Success) {
        # Use proper markdown converter instead of inline regex
        $aiContent = ConvertTo-HTMLFromMarkdown -Markdown $AIExecutiveSummary.Summary

        $html += '<div class="stats-panel"><h3>Executive Summary</h3><div style="color: var(--text-secondary); line-height: 1.8; font-size: 0.95rem;">' + $aiContent + '</div></div>'
    }

    # Executive Footer
    $html += @"
</div>
<div class="exec-footer">
    <div class="confidential">Confidential - Executive Summary</div>
    <div class="generated">Generated $(Get-Date -Format "MMMM dd, yyyy") | M365 Migration Assessment Tool</div>
</div>
</body>
</html>
"@

    # Write to file
    $html | Out-File -FilePath $OutputPath -Encoding UTF8

    Write-Log -Message "Executive Summary Report generated: $OutputPath" -Level Success

    return $OutputPath
}
#endregion

#region PDF Generation

function Get-PDFStylesheet {
    <#
    .SYNOPSIS
        Returns professional light-themed CSS for PDF generation
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet("IT", "Executive")]
        [string]$ReportType = "IT"
    )

    $primaryColor = if ($ReportType -eq "Executive") { "#7c3aed" } else { "#0369a1" }
    $accentColor = if ($ReportType -eq "Executive") { "#a855f7" } else { "#0ea5e9" }
    $lightBg = if ($ReportType -eq "Executive") { "#faf5ff" } else { "#f0f9ff" }

    return @"
        @page {
            size: A4;
            margin: 20mm 15mm 25mm 15mm;
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }

        body {
            font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
            font-size: 10pt;
            line-height: 1.6;
            color: #1f2937;
            background: #ffffff;
            counter-reset: page;
        }

        .container {
            max-width: 100%;
            padding: 0;
        }

        /* Cover Page - uses simple solid background, no SVG patterns or emojis for reliable PDF rendering */
        .cover-page {
            background: $primaryColor;
            color: white;
            page-break-after: always;
            padding: 80pt 30pt 60pt 30pt;
            position: relative;
            min-height: 600pt;
            border-radius: 6pt;
        }

        .cover-content {
            text-align: center;
            padding-top: 40pt;
        }

        .cover-logo {
            width: 80pt;
            height: 80pt;
            background: rgba(255,255,255,0.15);
            border: 3pt solid rgba(255,255,255,0.3);
            border-radius: 50%;
            line-height: 80pt;
            margin: 0 auto 30pt auto;
            font-size: 28pt;
            font-weight: 700;
            color: white;
        }

        .cover-title {
            font-size: 28pt;
            font-weight: 700;
            margin-bottom: 12pt;
            letter-spacing: -0.5pt;
        }

        .cover-subtitle {
            font-size: 14pt;
            opacity: 0.9;
            margin-bottom: 40pt;
            font-weight: 300;
        }

        .cover-tenant {
            font-size: 18pt;
            font-weight: 600;
            background: rgba(255,255,255,0.15);
            padding: 16pt 32pt;
            border-radius: 8pt;
            margin-bottom: 30pt;
            display: inline-block;
        }

        .cover-meta {
            font-size: 10pt;
            opacity: 0.8;
            line-height: 1.8;
        }

        .cover-footer {
            position: absolute;
            bottom: 40pt;
            left: 0;
            right: 0;
            text-align: center;
            font-size: 9pt;
            opacity: 0.7;
        }

        .cover-accent-bar {
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 8pt;
            background: $accentColor;
        }

        .cover-accent-bar-bottom {
            position: absolute;
            bottom: 0;
            left: 0;
            right: 0;
            height: 8pt;
            background: $accentColor;
        }

        /* Table of Contents */
        .toc {
            page-break-after: always;
            padding: 40pt 20pt;
        }

        .toc h2 {
            font-size: 18pt;
            color: $primaryColor;
            margin-bottom: 24pt;
            padding-bottom: 10pt;
            border-bottom: 3px solid $accentColor;
        }

        .toc-item {
            display: flex;
            justify-content: space-between;
            align-items: baseline;
            padding: 8pt 0;
            border-bottom: 1px dotted #d1d5db;
        }

        .toc-item.level-1 { font-weight: 600; font-size: 11pt; }
        .toc-item.level-2 { padding-left: 16pt; font-size: 10pt; }
        .toc-item.level-3 { padding-left: 32pt; font-size: 9pt; color: #4b5563; }

        .toc-title { flex: 1; }
        .toc-page {
            color: $primaryColor;
            font-weight: 600;
            min-width: 30pt;
            text-align: right;
        }

        /* Page Header */
        .page-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10pt 0 16pt 0;
            margin-bottom: 16pt;
            border-bottom: 2px solid $accentColor;
        }

        .page-header .doc-title {
            font-size: 9pt;
            color: $primaryColor;
            font-weight: 600;
        }

        .page-header .tenant-name {
            font-size: 9pt;
            color: #6b7280;
        }

        /* Header */
        .report-header {
            background: linear-gradient(135deg, $primaryColor, $accentColor);
            color: white;
            padding: 24pt 20pt;
            margin-bottom: 20pt;
            border-radius: 8pt;
        }

        .report-header h1 {
            font-size: 18pt;
            font-weight: 700;
            margin-bottom: 4pt;
        }

        .report-header .subtitle {
            font-size: 11pt;
            opacity: 0.9;
        }

        .report-header .meta {
            margin-top: 12pt;
            font-size: 9pt;
            opacity: 0.85;
        }

        /* Section Cards - NO page-break-inside:avoid since sections can span multiple pages */
        .section {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            padding: 16pt;
            margin-bottom: 16pt;
            border-radius: 6pt;
        }

        /* Compact sections that should stay on one page */
        .section.keep-together {
            page-break-inside: avoid;
        }

        .section.page-break {
            page-break-before: always;
            margin-top: 0;
        }

        .section h2 {
            font-size: 13pt;
            font-weight: 600;
            color: $primaryColor;
            margin-bottom: 12pt;
            padding-bottom: 6pt;
            border-bottom: 2px solid $accentColor;
        }

        .section h3 {
            font-size: 11pt;
            font-weight: 600;
            color: #374151;
            margin: 12pt 0 8pt 0;
        }

        .section h4 {
            font-size: 10pt;
            font-weight: 600;
            color: #4b5563;
            margin: 10pt 0 6pt 0;
        }

        /* Prevent headings from being orphaned at page bottom */
        h1, h2, h3, h4 {
            page-break-after: avoid;
        }

        /* Keep summary grids, KPI grids, risk matrices together */
        .summary-grid,
        .kpi-grid,
        .risk-matrix,
        .readiness-hero,
        .risk-meter,
        .exec-hero {
            page-break-inside: avoid;
        }

        /* Better text flow across pages */
        p {
            orphans: 3;
            widows: 3;
        }

        .section-intro {
            background: $lightBg;
            border-left: 4px solid $primaryColor;
            padding: 12pt;
            margin-bottom: 16pt;
            font-size: 9pt;
            color: #4b5563;
            border-radius: 0 4pt 4pt 0;
        }

        /* Summary Cards Grid */
        .summary-grid {
            display: flex;
            flex-wrap: wrap;
            gap: 12pt;
            margin-bottom: 16pt;
        }

        .summary-card {
            flex: 1 1 120pt;
            background: #f9fafb;
            border: 1px solid #e5e7eb;
            border-left: 4px solid $primaryColor;
            padding: 12pt;
            border-radius: 4pt;
            text-align: center;
        }

        .summary-card .value {
            font-size: 20pt;
            font-weight: 700;
            color: $primaryColor;
            line-height: 1.2;
        }

        .summary-card .label {
            font-size: 8pt;
            color: #6b7280;
            text-transform: uppercase;
            letter-spacing: 0.5pt;
            margin-top: 4pt;
        }

        .summary-card.critical { border-left-color: #dc2626; }
        .summary-card.critical .value { color: #dc2626; }
        .summary-card.high { border-left-color: #ea580c; }
        .summary-card.high .value { color: #ea580c; }
        .summary-card.medium { border-left-color: #d97706; }
        .summary-card.medium .value { color: #d97706; }
        .summary-card.low { border-left-color: #16a34a; }
        .summary-card.low .value { color: #16a34a; }

        /* Tables */
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 10pt 0;
            font-size: 9pt;
        }

        th {
            background: #f3f4f6;
            color: #374151;
            font-weight: 600;
            text-align: left;
            padding: 8pt 10pt;
            border: 1px solid #e5e7eb;
        }

        td {
            padding: 8pt 10pt;
            border: 1px solid #e5e7eb;
            vertical-align: top;
            word-break: break-word;
            overflow-wrap: break-word;
            max-width: 300pt;
        }

        tr:nth-child(even) {
            background: #f9fafb;
        }

        /* Severity badges */
        .severity-badge {
            display: inline-block;
            padding: 2pt 8pt;
            border-radius: 10pt;
            font-size: 8pt;
            font-weight: 600;
            text-transform: uppercase;
        }

        .severity-critical { background: #fef2f2; color: #dc2626; }
        .severity-high { background: #fff7ed; color: #ea580c; }
        .severity-medium { background: #fffbeb; color: #d97706; }
        .severity-low { background: #f0fdf4; color: #16a34a; }

        /* Gotcha Cards */
        .gotcha-card {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-left: 4px solid #6b7280;
            padding: 12pt;
            margin-bottom: 10pt;
            border-radius: 4pt;
            page-break-inside: avoid;
        }

        .gotcha-card.critical { border-left-color: #dc2626; }
        .gotcha-card.high { border-left-color: #ea580c; }
        .gotcha-card.medium { border-left-color: #d97706; }
        .gotcha-card.low { border-left-color: #16a34a; }

        .gotcha-card h4 {
            font-size: 10pt;
            font-weight: 600;
            color: #1f2937;
            margin-bottom: 6pt;
        }

        .gotcha-card p {
            font-size: 9pt;
            color: #4b5563;
            margin-bottom: 8pt;
        }

        .gotcha-card .recommendation {
            background: #eff6ff;
            border-left: 3px solid #3b82f6;
            padding: 8pt 10pt;
            font-size: 9pt;
            color: #1e40af;
            border-radius: 0 4pt 4pt 0;
        }

        /* AI Analysis */
        .ai-analysis {
            background: #faf5ff;
            border: 1px solid #e9d5ff;
            border-left: 4px solid $primaryColor;
            padding: 14pt;
            margin: 12pt 0;
            border-radius: 4pt;
        }

        .ai-analysis h4 {
            color: $primaryColor;
            font-size: 10pt;
            margin-bottom: 8pt;
        }

        .ai-content {
            font-size: 9pt;
            color: #374151;
            line-height: 1.6;
        }

        .ai-content h1, .ai-content h2, .ai-content h3 {
            color: $primaryColor;
            margin: 12pt 0 6pt 0;
        }

        .ai-content h1 { font-size: 12pt; }
        .ai-content h2 { font-size: 11pt; }
        .ai-content h3 { font-size: 10pt; }

        .ai-content ul, .ai-content ol {
            margin-left: 16pt;
            margin-bottom: 8pt;
        }

        .ai-content li {
            margin-bottom: 4pt;
        }

        .ai-content code {
            background: #f3f4f6;
            padding: 1pt 4pt;
            border-radius: 2pt;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 8pt;
            color: #1f2937;
        }

        .ai-content pre {
            background: #f1f5f9;
            color: #1e293b;
            padding: 12pt;
            border-radius: 6pt;
            border-left: 4px solid $primaryColor;
            font-size: 8pt;
            line-height: 1.4;
            margin: 10pt 0;
            page-break-inside: avoid;
            white-space: pre-wrap;
            word-break: break-word;
            overflow-wrap: break-word;
        }

        .ai-content pre code {
            background: none;
            color: inherit;
            padding: 0;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            white-space: pre-wrap;
            word-break: break-word;
        }

        .ai-content table {
            font-size: 8pt;
            margin: 10pt 0;
        }

        .ai-content p {
            margin-bottom: 8pt;
        }

        .ai-content blockquote {
            border-left: 4px solid $accentColor;
            background: $lightBg;
            padding: 10pt 12pt;
            margin: 10pt 0;
            font-style: italic;
            color: #4b5563;
        }

        /* Risk Meter */
        .risk-meter {
            background: #f9fafb;
            border: 1px solid #e5e7eb;
            padding: 16pt;
            border-radius: 6pt;
            margin-bottom: 16pt;
        }

        .risk-meter h2 {
            font-size: 12pt;
            color: #374151;
            margin-bottom: 10pt;
        }

        .risk-bar-container {
            background: #e5e7eb;
            border-radius: 8pt;
            height: 16pt;
            overflow: hidden;
            margin-bottom: 8pt;
        }

        .risk-bar {
            height: 100%;
            border-radius: 8pt;
            background: linear-gradient(90deg, #16a34a, #d97706, #dc2626);
        }

        /* Footer */
        .report-footer {
            margin-top: 20pt;
            padding-top: 12pt;
            border-top: 1px solid #e5e7eb;
            text-align: center;
            font-size: 8pt;
            color: #6b7280;
        }

        /* Executive-specific styles */
        .exec-hero {
            background: linear-gradient(135deg, $primaryColor, $accentColor);
            color: white;
            padding: 30pt;
            margin-bottom: 20pt;
            border-radius: 8pt;
            text-align: center;
        }

        .exec-hero h1 {
            font-size: 22pt;
            font-weight: 700;
            margin-bottom: 6pt;
        }

        .exec-hero .org-name {
            font-size: 14pt;
            opacity: 0.9;
        }

        .exec-hero .date {
            font-size: 10pt;
            opacity: 0.8;
            margin-top: 8pt;
        }

        .readiness-hero {
            display: flex;
            align-items: center;
            gap: 30pt;
            padding: 24pt;
            background: #f9fafb;
            border: 1px solid #e5e7eb;
            border-radius: 8pt;
            margin-bottom: 20pt;
        }

        .readiness-circle {
            width: 140pt;
            height: 140pt;
            flex-shrink: 0;
        }

        .readiness-circle svg {
            width: 100%;
            height: 100%;
        }

        /* SVG text styling */
        .readiness-circle svg text {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif;
        }

        .readiness-circle svg .score-value {
            font-size: 18pt;
            font-weight: 700;
            fill: $primaryColor;
        }

        .readiness-circle svg .score-label {
            font-size: 7pt;
            fill: #6b7280;
            text-transform: uppercase;
        }

        .kpi-grid {
            display: flex;
            flex-wrap: wrap;
            gap: 12pt;
            margin-bottom: 20pt;
        }

        .kpi-card {
            flex: 1 1 100pt;
            background: #f9fafb;
            border: 1px solid #e5e7eb;
            border-top: 4px solid $primaryColor;
            padding: 14pt;
            border-radius: 6pt;
            text-align: center;
        }

        .kpi-card .kpi-value {
            font-size: 18pt;
            font-weight: 700;
            color: $primaryColor;
        }

        .kpi-card .kpi-label {
            font-size: 8pt;
            color: #6b7280;
            text-transform: uppercase;
            margin-top: 4pt;
        }

        /* Timeline/Roadmap */
        .timeline {
            position: relative;
            padding-left: 30pt;
            margin: 16pt 0;
        }

        .timeline::before {
            content: '';
            position: absolute;
            left: 10pt;
            top: 0;
            bottom: 0;
            width: 2pt;
            background: $accentColor;
        }

        .timeline-item {
            position: relative;
            margin-bottom: 16pt;
            padding-bottom: 16pt;
            page-break-inside: avoid;
        }

        .timeline-item::before {
            content: '';
            position: absolute;
            left: -24pt;
            top: 2pt;
            width: 12pt;
            height: 12pt;
            border-radius: 50%;
            background: $primaryColor;
            border: 3pt solid white;
            box-shadow: 0 0 0 1pt $primaryColor;
        }

        .timeline-item h4 {
            font-size: 10pt;
            color: $primaryColor;
            margin-bottom: 4pt;
        }

        .timeline-item .timeline-meta {
            font-size: 8pt;
            color: #6b7280;
            margin-bottom: 6pt;
        }

        .timeline-item p {
            font-size: 9pt;
            color: #4b5563;
            line-height: 1.5;
        }

        /* Risk Matrix */
        .risk-matrix {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 10pt;
            margin: 16pt 0;
        }

        .risk-matrix-cell {
            padding: 12pt;
            border-radius: 6pt;
            text-align: center;
            border: 1px solid #e5e7eb;
        }

        .risk-matrix-cell.risk-critical {
            background: #fef2f2;
            border-color: #dc2626;
        }

        .risk-matrix-cell.risk-high {
            background: #fff7ed;
            border-color: #ea580c;
        }

        .risk-matrix-cell.risk-medium {
            background: #fffbeb;
            border-color: #d97706;
        }

        .risk-matrix-cell.risk-low {
            background: #f0fdf4;
            border-color: #16a34a;
        }

        .risk-matrix-cell .risk-label {
            font-size: 9pt;
            font-weight: 600;
            margin-bottom: 6pt;
            text-transform: uppercase;
        }

        .risk-matrix-cell .risk-count {
            font-size: 18pt;
            font-weight: 700;
        }

        .risk-matrix-cell.risk-critical .risk-count { color: #dc2626; }
        .risk-matrix-cell.risk-high .risk-count { color: #ea580c; }
        .risk-matrix-cell.risk-medium .risk-count { color: #d97706; }
        .risk-matrix-cell.risk-low .risk-count { color: #16a34a; }

        /* Data Table Enhancement */
        .data-table-wrapper {
            overflow-x: auto;
            margin: 10pt 0;
        }

        .data-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 8pt;
        }

        .data-table thead {
            background: $lightBg;
        }

        .data-table th {
            background: $primaryColor;
            color: white;
            font-weight: 600;
            text-align: left;
            padding: 8pt 10pt;
            border: 1px solid darken($primaryColor, 10%);
        }

        .data-table td {
            padding: 6pt 10pt;
            border: 1px solid #e5e7eb;
            vertical-align: top;
            word-break: break-word;
            overflow-wrap: break-word;
        }

        .data-table tbody tr:nth-child(even) {
            background: #fafafa;
        }

        .data-table tbody tr:hover {
            background: $lightBg;
        }

        /* Print-specific */
        @media print {
            body {
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            .cover-page {
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            .gotcha-card { page-break-inside: avoid; }
            .timeline-item { page-break-inside: avoid; }
            .summary-grid { page-break-inside: avoid; }
            .kpi-grid { page-break-inside: avoid; }
            .risk-matrix { page-break-inside: avoid; }
            .readiness-hero { page-break-inside: avoid; }
            .exec-hero { page-break-inside: avoid; }
            pre { page-break-inside: avoid; }
            table { page-break-inside: avoid; }
            h1, h2, h3, h4 { page-break-after: avoid; }
            p { orphans: 3; widows: 3; }
        }
"@
}

function ConvertTo-HTMLFromMarkdown {
    <#
    .SYNOPSIS
        Converts markdown text to well-formatted HTML
    .DESCRIPTION
        Enhanced markdown to HTML conversion with support for tables, code blocks,
        lists, blockquotes, and inline formatting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Markdown
    )

    if ([string]::IsNullOrWhiteSpace($Markdown)) {
        return ""
    }

    # Pre-process: restore newlines if content appears collapsed (all on one line)
    $lineCount = ($Markdown -split '\r?\n').Count
    if ($Markdown.Length -gt 500 -and $lineCount -lt 10) {
        $Markdown = $Markdown -replace '(?<=\S)\s*(#{1,4}\s+)', "`n`n`$1"
        $Markdown = $Markdown -replace '(?<=\S)\s*(---+)', "`n`n`$1`n"
        $Markdown = $Markdown -replace '(?<=\S)\s*(```)', "`n`$1"
        $Markdown = $Markdown -replace '(?<=[\.\!\?\:\"]\s*)(-\s+\S)', "`n`$1"
        $Markdown = $Markdown -replace '(?<=[\.\!\?\:\"\s])(\d+\.\s+)', "`n`$1"
        $Markdown = $Markdown -replace '(?<=\S)\s*(\|[^\|]+\|)', "`n`$1"
        $Markdown = $Markdown -replace '(?<=\S)\s*(>\s+)', "`n`$1"
        $Markdown = $Markdown -replace '\n{3,}', "`n`n"
    }

    $html = $Markdown

    # Escape HTML entities first
    $html = $html -replace '&', '&amp;'
    $html = $html -replace '<(?!/?(?:strong|em|code|pre|h[1-6]|ul|ol|li|p|br|blockquote|table|thead|tbody|tr|th|td)>)', '&lt;'
    $html = $html -replace '(?<!</?(?:strong|em|code|pre|h[1-6]|ul|ol|li|p|br|blockquote|table|thead|tbody|tr|th|td))>', '&gt;'

    # Code blocks (``` ... ```) - must be done before inline code
    $html = $html -replace '(?ms)```(\w*)\r?\n(.*?)```', '<pre><code class="language-$1">$2</code></pre>'

    # Tables (GitHub-flavored markdown)
    $html = [regex]::Replace($html, '(?m)^\|(.+)\|\r?$\n\|[\s:|-]+\|\r?$\n((?:^\|.+\|\r?$\n?)+)', {
        param($match)
        $headerRow = $match.Groups[1].Value
        $bodyRows = $match.Groups[2].Value

        $headers = $headerRow -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        $headerHtml = ($headers | ForEach-Object { "<th>$_</th>" }) -join ''

        $rows = $bodyRows -split '\r?\n' | Where-Object { $_ -match '^\|' }
        $rowsHtml = ($rows | ForEach-Object {
            $cells = $_ -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            $cellsHtml = ($cells | ForEach-Object { "<td>$_</td>" }) -join ''
            "<tr>$cellsHtml</tr>"
        }) -join ''

        "<table><thead><tr>$headerHtml</tr></thead><tbody>$rowsHtml</tbody></table>"
    })

    # Headers
    $html = $html -replace '(?m)^#### (.+)$', '<h4>$1</h4>'
    $html = $html -replace '(?m)^### (.+)$', '<h3>$1</h3>'
    $html = $html -replace '(?m)^## (.+)$', '<h2>$1</h2>'
    $html = $html -replace '(?m)^# (.+)$', '<h1>$1</h1>'

    # Blockquotes
    $html = $html -replace '(?m)^> (.+)$', '<blockquote>$1</blockquote>'

    # Horizontal rules
    $html = $html -replace '(?m)^---+$', '<hr/>'
    $html = $html -replace '(?m)^\*\*\*+$', '<hr/>'

    # Bold and italic
    $html = $html -replace '\*\*\*(.+?)\*\*\*', '<strong><em>$1</em></strong>'
    $html = $html -replace '\*\*(.+?)\*\*', '<strong>$1</strong>'
    $html = $html -replace '\*(.+?)\*', '<em>$1</em>'
    $html = $html -replace '___(.+?)___', '<strong><em>$1</em></strong>'
    $html = $html -replace '__(.+?)__', '<strong>$1</strong>'
    $html = $html -replace '_(.+?)_', '<em>$1</em>'

    # Inline code
    $html = $html -replace '`([^`]+)`', '<code>$1</code>'

    # Ordered lists
    $html = [regex]::Replace($html, '(?m)((?:^\d+\. .+$\r?\n?)+)', {
        param($match)
        $items = $match.Groups[0].Value -split '\r?\n' | Where-Object { $_ -match '^\d+\. ' }
        $listItems = ($items | ForEach-Object {
            $_ -replace '^\d+\. (.+)$', '<li>$1</li>'
        }) -join "`n"
        "<ol>`n$listItems`n</ol>"
    })

    # Unordered lists
    $html = [regex]::Replace($html, '(?m)((?:^[-*+] .+$\r?\n?)+)', {
        param($match)
        $items = $match.Groups[0].Value -split '\r?\n' | Where-Object { $_ -match '^[-*+] ' }
        $listItems = ($items | ForEach-Object {
            $_ -replace '^[-*+] (.+)$', '<li>$1</li>'
        }) -join "`n"
        "<ul>`n$listItems`n</ul>"
    })

    # Links
    $html = $html -replace '\[([^\]]+)\]\(([^)]+)\)', '<a href="$2">$1</a>'

    # Line breaks and paragraphs
    $html = $html -replace '\r?\n\r?\n', '</p><p>'
    $html = "<p>$html</p>"

    # Clean up empty paragraphs
    $html = $html -replace '<p>\s*</p>', ''
    $html = $html -replace '<p>(<(?:h[1-6]|ul|ol|table|pre|blockquote|hr)>)', '$1'
    $html = $html -replace '(</(?:h[1-6]|ul|ol|table|pre|blockquote|hr)>)</p>', '$1'

    return $html
}

function New-PDFCoverPage {
    <#
    .SYNOPSIS
        Creates a professional cover page for PDF reports
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Subtitle,

        [Parameter(Mandatory = $true)]
        [string]$TenantName,

        [Parameter(Mandatory = $false)]
        [string]$TenantId,

        [Parameter(Mandatory = $false)]
        [ValidateSet("IT", "Executive")]
        [string]$ReportType = "IT"
    )

    $reportTypeDisplay = if ($ReportType -eq "Executive") { "Executive Summary" } else { "Technical Assessment" }
    $icon = if ($ReportType -eq "Executive") { "M365" } else { "M365" }

    return @"
<div class="cover-page">
    <div class="cover-accent-bar"></div>
    <div class="cover-content">
        <div class="cover-logo">$icon</div>
        <div class="cover-title">$Title</div>
        <div class="cover-subtitle">$Subtitle</div>
        <div class="cover-tenant">$TenantName</div>
        <div class="cover-meta">
            <div><strong>Report Type:</strong> $reportTypeDisplay</div>
            $(if ($TenantId) { "<div><strong>Tenant ID:</strong> $TenantId</div>" })
            <div><strong>Generated:</strong> $(Get-Date -Format "MMMM dd, yyyy")</div>
        </div>
    </div>
    <div class="cover-footer">
        <p>Confidential - Internal Use Only</p>
        <p>M365 Tenant Discovery &amp; Migration Assessment Tool</p>
    </div>
    <div class="cover-accent-bar-bottom"></div>
</div>
"@
}

function New-PDFTOC {
    <#
    .SYNOPSIS
        Creates a table of contents for PDF reports
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable[]]$Sections
    )

    $tocItems = ""
    foreach ($section in $Sections) {
        $level = if ($section.Level) { $section.Level } else { 1 }
        $title = $section.Title
        $page = if ($section.Page) { $section.Page } else { "" }

        $tocItems += @"
        <div class="toc-item level-$level">
            <span class="toc-title">$title</span>
            <span class="toc-page">$page</span>
        </div>
"@
    }

    return @"
<div class="toc">
    <h2>Table of Contents</h2>
    $tocItems
</div>
"@
}

function New-PDFTimeline {
    <#
    .SYNOPSIS
        Creates a timeline/roadmap visualization for PDF reports
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable[]]$Milestones
    )

    $timelineItems = ""
    foreach ($milestone in $Milestones) {
        $title = $milestone.Title
        $duration = if ($milestone.Duration) { $milestone.Duration } else { "" }
        $description = if ($milestone.Description) { $milestone.Description } else { "" }

        $timelineItems += @"
        <div class="timeline-item">
            <h4>$title</h4>
            $(if ($duration) { "<div class='timeline-meta'>$duration</div>" })
            $(if ($description) { "<p>$description</p>" })
        </div>
"@
    }

    return @"
<div class="timeline">
    $timelineItems
</div>
"@
}

function New-PDFRiskMatrix {
    <#
    .SYNOPSIS
        Creates a risk matrix visualization for PDF reports
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$Critical = 0,

        [Parameter(Mandatory = $false)]
        [int]$High = 0,

        [Parameter(Mandatory = $false)]
        [int]$Medium = 0,

        [Parameter(Mandatory = $false)]
        [int]$Low = 0
    )

    return @"
<div class="risk-matrix">
    <div class="risk-matrix-cell risk-critical">
        <div class="risk-label">Critical</div>
        <div class="risk-count">$Critical</div>
    </div>
    <div class="risk-matrix-cell risk-high">
        <div class="risk-label">High</div>
        <div class="risk-count">$High</div>
    </div>
    <div class="risk-matrix-cell risk-medium">
        <div class="risk-label">Medium</div>
        <div class="risk-count">$Medium</div>
    </div>
    <div class="risk-matrix-cell risk-low">
        <div class="risk-label">Low</div>
        <div class="risk-count">$Low</div>
    </div>
</div>
"@
}

function New-PDFReport {
    <#
    .SYNOPSIS
        Generates a PDF report from HTML content
    .DESCRIPTION
        Creates a light-themed professional PDF report suitable for printing and sharing.
        Uses browser-based PDF generation or wkhtmltopdf if available.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HTMLContent,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $false)]
        [ValidateSet("IT", "Executive")]
        [string]$ReportType = "IT"
    )

    # Create a temporary HTML file with light theme
    $tempHtmlPath = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.html'

    try {
        # Write HTML to temp file
        $HTMLContent | Out-File -FilePath $tempHtmlPath -Encoding UTF8

        # Try different PDF generation methods
        $pdfGenerated = $false

        # Method 1: Try wkhtmltopdf (most reliable for server environments)
        $wkhtmltopdf = Get-Command wkhtmltopdf -ErrorAction SilentlyContinue
        if ($wkhtmltopdf) {
            Write-Log -Message "Using wkhtmltopdf for PDF generation" -Level Info
            $wkArgs = @(
                "--enable-local-file-access",
                "--page-size", "A4",
                "--margin-top", "20mm",
                "--margin-bottom", "25mm",
                "--margin-left", "15mm",
                "--margin-right", "15mm",
                "--print-media-type",
                "--no-stop-slow-scripts",
                "--enable-smart-shrinking",
                $tempHtmlPath,
                $OutputPath
            )
            & wkhtmltopdf @wkArgs 2>&1 | Out-Null
            $pdfGenerated = Test-Path $OutputPath
        }

        # Method 2: Try Microsoft Edge/Chrome with headless mode
        if (-not $pdfGenerated) {
            $browsers = @(
                "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                "C:\Program Files\Google\Chrome\Application\chrome.exe",
                "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                "/usr/bin/google-chrome",
                "/usr/bin/chromium-browser",
                "/usr/bin/microsoft-edge"
            )

            foreach ($browser in $browsers) {
                if (Test-Path $browser) {
                    Write-Log -Message "Using browser for PDF generation: $browser" -Level Info
                    $browserArgs = @(
                        "--headless",
                        "--disable-gpu",
                        "--no-sandbox",
                        "--run-all-compositor-stages-before-draw",
                        "--print-to-pdf=$OutputPath",
                        "--print-to-pdf-no-header",
                        "file:///$($tempHtmlPath -replace '\\', '/')"
                    )
                    & $browser @browserArgs 2>&1 | Out-Null
                    Start-Sleep -Seconds 2
                    $pdfGenerated = Test-Path $OutputPath
                    if ($pdfGenerated) { break }
                }
            }
        }

        # Method 3: PowerShell with iTextSharp or similar (fallback - just copy HTML)
        if (-not $pdfGenerated) {
            Write-Log -Message "PDF generation tools not found. Saving as HTML for manual conversion." -Level Warning
            Copy-Item $tempHtmlPath $($OutputPath -replace '\.pdf$', '.html')
            return $($OutputPath -replace '\.pdf$', '.html')
        }

        Write-Log -Message "PDF report generated: $OutputPath" -Level Success
        return $OutputPath
    }
    finally {
        # Cleanup temp file
        if (Test-Path $tempHtmlPath) {
            Remove-Item $tempHtmlPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function New-ITDetailedReportPDF {
    <#
    .SYNOPSIS
        Generates a professional light-themed PDF version of the IT Detailed Report
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $true)]
        $AnalysisResults,

        [Parameter(Mandatory = $true)]
        $ComplexityScore,

        [Parameter(Mandatory = $false)]
        $AIGotchaAnalysis,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Generating IT Detailed Report (PDF)..." -Level Info

    $css = Get-PDFStylesheet -ReportType "IT"

    # Get tenant info with fallbacks for different data structures
    $tenantName = if ($CollectedData.TenantInfo.DisplayName) {
        $CollectedData.TenantInfo.DisplayName
    } elseif ($CollectedData.EntraID.TenantInfo.DisplayName) {
        $CollectedData.EntraID.TenantInfo.DisplayName
    } else {
        "Unknown Tenant"
    }

    $tenantId = if ($CollectedData.TenantInfo.TenantId) {
        $CollectedData.TenantInfo.TenantId
    } elseif ($CollectedData.EntraID.TenantInfo.TenantId) {
        $CollectedData.EntraID.TenantInfo.TenantId
    } else {
        ""
    }

    # Create cover page
    $coverPage = New-PDFCoverPage `
        -Title "M365 Migration Technical Assessment" `
        -Subtitle "Comprehensive IT Analysis & Migration Readiness Report" `
        -TenantName $tenantName `
        -TenantId $tenantId `
        -ReportType "IT"

    # Create table of contents
    $tocSections = @(
        @{ Title = "Environment Overview"; Level = 1 }
        @{ Title = "Risk Assessment Summary"; Level = 1 }
        @{ Title = "Critical Priority Issues"; Level = 1 }
        @{ Title = "High Priority Issues"; Level = 1 }
        @{ Title = "Medium Priority Issues"; Level = 1 }
        @{ Title = "Low Priority Issues"; Level = 1 }
        @{ Title = "Deep Analysis"; Level = 1 }
        @{ Title = "Migration Recommendations"; Level = 2 }
        @{ Title = "Technical Considerations"; Level = 2 }
    )
    $toc = New-PDFTOC -Sections $tocSections

    # Build HTML with light theme
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>IT Technical Report - $tenantName</title>
    <style>$css</style>
</head>
<body>
$coverPage
$toc

<div class="container">

<div class="page-header">
    <div class="doc-title">M365 Migration Technical Assessment</div>
    <div class="tenant-name">$tenantName</div>
</div>

"@

    # Summary Statistics
    $totalUsers = $CollectedData.EntraID.Users.Analysis.TotalUsers
    $licensedUsers = $CollectedData.EntraID.Users.Analysis.LicensedUsers
    $totalMailboxes = $CollectedData.Exchange.Mailboxes.Analysis.TotalMailboxes
    $totalSites = $CollectedData.SharePoint.Sites.Analysis.SharePointSites
    $totalTeams = $CollectedData.Teams.Teams.Analysis.TotalTeams
    $criticalCount = $AnalysisResults.BySeverity.Critical.Count
    $highCount = $AnalysisResults.BySeverity.High.Count
    $mediumCount = $AnalysisResults.BySeverity.Medium.Count
    $lowCount = $AnalysisResults.BySeverity.Low.Count
    $totalIssues = $criticalCount + $highCount + $mediumCount + $lowCount

    # Calculate top issues by category for key findings
    $issuesByCategory = @{}
    foreach ($severity in @("Critical", "High", "Medium", "Low")) {
        $issues = $AnalysisResults.BySeverity[$severity]
        if ($issues) {
            foreach ($issue in $issues) {
                if (-not $issuesByCategory.ContainsKey($issue.Category)) {
                    $issuesByCategory[$issue.Category] = 0
                }
                $issuesByCategory[$issue.Category]++
            }
        }
    }
    $topCategories = $issuesByCategory.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 3

    # Determine migration readiness status
    $readinessStatus = if ($criticalCount -eq 0 -and $highCount -le 2) {
        "Ready to proceed with planning"
    } elseif ($criticalCount -eq 0) {
        "Address high priority items before migration"
    } else {
        "Critical blockers must be resolved"
    }

    $html += @"
<div class="section keep-together">
    <h2>Executive Summary</h2>
    <div class="section-intro">
        This assessment analyzed your M365 environment to identify potential migration risks, complexities, and blockers.
        Below is a high-level overview of key findings.
    </div>

    <div style="background: #f0f9ff; border-left: 4px solid #0ea5e9; padding: 14pt; border-radius: 4pt; margin-bottom: 16pt;">
        <h3 style="margin-top: 0; color: #0369a1; font-size: 11pt;">Migration Readiness: $readinessStatus</h3>
        <p style="margin: 8pt 0; color: #1f2937;">
            <strong>Complexity Score:</strong> $($ComplexityScore.TotalScore)/100 ($($ComplexityScore.ComplexityLevel)) |
            <strong>Risk Level:</strong> $($AnalysisResults.RiskLevel) |
            <strong>Total Issues:</strong> $totalIssues
        </p>
    </div>

    <h3 style="font-size: 10pt; margin-bottom: 8pt; color: #1f2937;">Environment Scope</h3>
    <div class="summary-grid">
        <div class="summary-card">
            <div class="value">$licensedUsers</div>
            <div class="label">Licensed Users</div>
        </div>
        <div class="summary-card">
            <div class="value">$totalMailboxes</div>
            <div class="label">Mailboxes</div>
        </div>
        <div class="summary-card">
            <div class="value">$totalSites</div>
            <div class="label">SharePoint Sites</div>
        </div>
        <div class="summary-card">
            <div class="value">$totalTeams</div>
            <div class="label">Teams</div>
        </div>
    </div>
</div>

"@

    $riskMatrix = New-PDFRiskMatrix -Critical $criticalCount -High $highCount -Medium $mediumCount -Low $lowCount

    $html += @"
<div class="section page-break keep-together">
    <h2>Assessment At a Glance</h2>

    <h3 style="font-size: 10pt; margin-bottom: 10pt; color: #1f2937;">Issues by Severity</h3>
    $riskMatrix

    <h3 style="font-size: 10pt; margin: 16pt 0 10pt 0; color: #1f2937;">Top Areas of Concern</h3>
"@

    if ($topCategories.Count -gt 0) {
        $html += "<div style='background: #f9fafb; padding: 12pt; border-radius: 4pt;'>`n"
        $html += "<ul style='margin: 0; padding-left: 20pt; color: #4b5563;'>`n"
        foreach ($cat in $topCategories) {
            $html += "<li style='margin-bottom: 6pt;'><strong>$($cat.Key):</strong> $($cat.Value) issue(s)</li>`n"
        }
        $html += "</ul>`n</div>`n"
    }

    $html += @"

    <h3 style="font-size: 10pt; margin: 16pt 0 10pt 0; color: #1f2937;">Immediate Action Required</h3>
    <div style="background: #fef2f2; border-left: 4px solid #dc2626; padding: 12pt; border-radius: 0 4pt 4pt 0; margin-bottom: 16pt;">
"@

    if ($criticalCount -gt 0) {
        $html += "<p style='margin: 0 0 8pt 0; color: #991b1b; font-weight: 600;'>⚠ $criticalCount Critical Issue(s) Detected</p>`n"
        $html += "<p style='margin: 0; color: #4b5563; font-size: 9pt;'>These must be resolved before proceeding with migration. See Critical Priority Issues section for details.</p>`n"
    } elseif ($highCount -gt 0) {
        $html += "<p style='margin: 0 0 8pt 0; color: #c2410c; font-weight: 600;'>⚠ $highCount High Priority Issue(s) Detected</p>`n"
        $html += "<p style='margin: 0; color: #4b5563; font-size: 9pt;'>Address these during the preparation phase to ensure a smooth migration.</p>`n"
    } else {
        $html += "<p style='margin: 0; color: #16a34a; font-weight: 600;'>✓ No Critical or High Priority Blockers</p>`n"
        $html += "<p style='margin: 0; color: #4b5563; font-size: 9pt;'>Your environment is well-positioned for migration. Review medium and low priority items for optimization opportunities.</p>`n"
    }

    $html += @"
    </div>
</div>

"@

    # Migration Gotchas by Severity
    $severityDescriptions = @{
        "Critical" = "These issues must be resolved before migration as they will block or severely impact the process."
        "High"     = "These issues should be addressed during the preparation phase to ensure a smooth migration."
        "Medium"   = "These issues should be reviewed and planned for, though they may not be migration blockers."
        "Low"      = "These are minor considerations that can typically be addressed after migration."
    }

    $severityColors = @{
        "Critical" = "#dc2626"
        "High"     = "#ea580c"
        "Medium"   = "#d97706"
        "Low"      = "#16a34a"
    }

    foreach ($severity in @("Critical", "High", "Medium", "Low")) {
        $gotchas = $AnalysisResults.BySeverity[$severity]
        if ($gotchas -and $gotchas.Count -gt 0) {
            # All severity sections start on a new page for clean separation
            $pageBreakClass = "page-break"

            # Group issues by category
            $issuesByCategory = @{}
            foreach ($issue in $gotchas) {
                if (-not $issuesByCategory.ContainsKey($issue.Category)) {
                    $issuesByCategory[$issue.Category] = @()
                }
                $issuesByCategory[$issue.Category] += $issue
            }

            $sevColor = $severityColors[$severity]
            $html += @"
<div class="section $pageBreakClass">
    <h2><span style="color: $sevColor; font-size: 14pt;">&#9679;</span> $severity Priority Issues ($($gotchas.Count))</h2>
    <div class="section-intro">$($severityDescriptions[$severity])</div>

    <div style="background: #f9fafb; padding: 10pt; border-radius: 4pt; margin-bottom: 12pt;">
        <p style="margin: 0; font-size: 9pt; color: #6b7280;"><strong>Categories affected:</strong> $($issuesByCategory.Keys -join ', ')</p>
    </div>

"@
            # Display issues grouped by category
            foreach ($category in ($issuesByCategory.Keys | Sort-Object)) {
                $categoryIssues = $issuesByCategory[$category]

                $html += @"
    <h3 style="font-size: 10pt; color: #374151; margin: 14pt 0 8pt 0; border-bottom: 1px solid #e5e7eb; padding-bottom: 4pt;">
        $category ($($categoryIssues.Count))
    </h3>

"@
                foreach ($gotcha in $categoryIssues) {
                    $html += @"
    <div class="gotcha-card $($severity.ToLower())">
        <h4>$($gotcha.Name)</h4>
        <p style="color: #6b7280; font-size: 9pt; margin-bottom: 8pt;"><strong>Affected:</strong> $($gotcha.AffectedCount) items</p>
        <p style="margin-bottom: 8pt;">$($gotcha.Description)</p>
        <div class="recommendation"><strong>Recommendation:</strong> $($gotcha.Recommendation)</div>
    </div>

"@
                }
            }
            $html += "</div>`n"
        }
    }

    # AI Analysis Section - uses page-break-before but NOT page-break-inside:avoid
    # since AI content can span many pages
    if ($AIGotchaAnalysis -and $AIGotchaAnalysis.Success -and $AIGotchaAnalysis.Analysis) {
        # Convert markdown to HTML using improved converter
        $aiHtml = ConvertTo-HTMLFromMarkdown -Markdown $AIGotchaAnalysis.Analysis

        # Use string concatenation to avoid PowerShell interpolation of $ in AI content
        $html += '<div class="section page-break"><h2>Deep Migration Analysis</h2><div class="section-intro">Comprehensive insights, recommendations, and actionable guidance for your migration journey.</div><div class="ai-content">' + $aiHtml + '</div></div>'
    }

    # Footer
    $html += @"
<div class="report-footer">
    <p>Generated by M365 Tenant Discovery Tool | $(Get-Date -Format "MMMM dd, yyyy HH:mm:ss")</p>
    <p>This report contains confidential tenant information. Handle according to your organization's data classification policies.</p>
</div>

</div>
</body>
</html>
"@

    # Generate PDF
    $pdfPath = New-PDFReport -HTMLContent $html -OutputPath $OutputPath -ReportType "IT"

    return $pdfPath
}

function New-ExecutiveSummaryReportPDF {
    <#
    .SYNOPSIS
        Generates a professional light-themed PDF version of the Executive Summary Report
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $true)]
        $AnalysisResults,

        [Parameter(Mandatory = $true)]
        $ComplexityScore,

        [Parameter(Mandatory = $false)]
        $AIExecutiveSummary,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Generating Executive Summary Report (PDF)..." -Level Info

    $css = Get-PDFStylesheet -ReportType "Executive"

    # Get tenant info with fallbacks for different data structures
    $tenantName = if ($CollectedData.TenantInfo.DisplayName) {
        $CollectedData.TenantInfo.DisplayName
    } elseif ($CollectedData.EntraID.TenantInfo.DisplayName) {
        $CollectedData.EntraID.TenantInfo.DisplayName
    } else {
        "Unknown Tenant"
    }

    $tenantId = if ($CollectedData.TenantInfo.TenantId) {
        $CollectedData.TenantInfo.TenantId
    } elseif ($CollectedData.EntraID.TenantInfo.TenantId) {
        $CollectedData.EntraID.TenantInfo.TenantId
    } else {
        ""
    }

    # Create cover page
    $coverPage = New-PDFCoverPage `
        -Title "M365 Migration Readiness Assessment" `
        -Subtitle "Executive Summary & Strategic Overview" `
        -TenantName $tenantName `
        -TenantId $tenantId `
        -ReportType "Executive"

    # Calculate metrics with null safety
    $licensedUsers = if ($CollectedData.EntraID.Users.Analysis) { $CollectedData.EntraID.Users.Analysis.LicensedUsers } else { 0 }
    $totalMailboxes = if ($CollectedData.Exchange.Mailboxes.Analysis) { $CollectedData.Exchange.Mailboxes.Analysis.TotalMailboxes } else { 0 }
    $totalSites = if ($CollectedData.SharePoint.Sites.Analysis) { $CollectedData.SharePoint.Sites.Analysis.SharePointSites } else { 0 }
    $totalTeams = if ($CollectedData.Teams.Teams.Analysis) { $CollectedData.Teams.Teams.Analysis.TotalTeams } else { 0 }

    # Calculate readiness score (inverse of complexity)
    $readinessScore = [math]::Max(0, 100 - $ComplexityScore.TotalScore)
    $readinessColor = switch ($readinessScore) {
        { $_ -ge 80 } { "#16a34a" }
        { $_ -ge 60 } { "#d97706" }
        { $_ -ge 40 } { "#ea580c" }
        default { "#dc2626" }
    }
    $readinessLevel = switch ($readinessScore) {
        { $_ -ge 80 } { "Ready for Migration" }
        { $_ -ge 60 } { "Minor Preparation Needed" }
        { $_ -ge 40 } { "Moderate Preparation Required" }
        default { "Significant Work Required" }
    }

    $circumference = 2 * 3.14159 * 52
    $dashOffset = $circumference * (1 - ($readinessScore / 100))

    # Calculate timeline based on complexity
    $timelineMilestones = @(
        @{
            Title = "Phase 1: Discovery & Planning"
            Duration = "2-4 weeks"
            Description = "Complete tenant discovery, stakeholder alignment, and detailed migration planning."
        },
        @{
            Title = "Phase 2: Preparation & Remediation"
            Duration = switch ($ComplexityScore.TotalScore) {
                { $_ -lt 30 } { "2-3 weeks" }
                { $_ -lt 60 } { "4-6 weeks" }
                default { "6-8 weeks" }
            }
            Description = "Address critical blockers, prepare infrastructure, configure target tenant, and conduct pilot testing."
        },
        @{
            Title = "Phase 3: Migration Execution"
            Duration = "Varies by workload"
            Description = "Execute phased migration of users, data, and workloads with continuous monitoring and support."
        },
        @{
            Title = "Phase 4: Post-Migration & Optimization"
            Duration = "2-4 weeks"
            Description = "Validate migration success, optimize configurations, and provide hypercare support."
        }
    )
    $timeline = New-PDFTimeline -Milestones $timelineMilestones

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Executive Summary - $tenantName</title>
    <style>$css</style>
</head>
<body>
$coverPage

<div class="container">

<div class="page-header">
    <div class="doc-title">M365 Migration Readiness Assessment</div>
    <div class="tenant-name">$tenantName</div>
</div>

<div class="exec-hero">
    <h1>Migration Readiness Assessment</h1>
    <div class="org-name">$tenantName</div>
    <div class="date">$(Get-Date -Format "MMMM dd, yyyy")</div>
</div>

<div class="readiness-hero">
    <div class="readiness-circle">
        <svg viewBox="0 0 140 140" xmlns="http://www.w3.org/2000/svg">
            <!-- Background circle -->
            <circle cx="70" cy="70" r="52" fill="none" stroke="#e5e7eb" stroke-width="10"/>
            <!-- Progress circle -->
            <circle cx="70" cy="70" r="52" fill="none" stroke="$readinessColor" stroke-width="10"
                stroke-linecap="round" stroke-dasharray="$circumference" stroke-dashoffset="$dashOffset"
                transform="rotate(-90 70 70)"/>
            <!-- Score text -->
            <text x="70" y="65" text-anchor="middle" class="score-value">$readinessScore%</text>
            <text x="70" y="80" text-anchor="middle" class="score-label">READY</text>
        </svg>
    </div>
    <div style="flex: 1;">
        <h2 style="font-size: 16pt; color: #1f2937; margin-bottom: 6pt;">Migration Readiness</h2>
        <p style="font-size: 12pt; color: $readinessColor; font-weight: 600; margin-bottom: 12pt;">$readinessLevel</p>
        <p style="font-size: 10pt; color: #6b7280; line-height: 1.6;">
            Your organization has been assessed for Microsoft 365 tenant migration.
            This report provides a high-level overview of your current environment and migration scope.
        </p>
    </div>
</div>

<div class="kpi-grid">
    <div class="kpi-card">
        <div class="kpi-value">$licensedUsers</div>
        <div class="kpi-label">Users</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-value">$totalMailboxes</div>
        <div class="kpi-label">Mailboxes</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-value">$totalSites</div>
        <div class="kpi-label">SharePoint Sites</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-value">$totalTeams</div>
        <div class="kpi-label">Teams</div>
    </div>
</div>

"@

    # AI Executive Summary - page-break-before but content flows naturally across pages
    if ($AIExecutiveSummary -and $AIExecutiveSummary.Success -and $AIExecutiveSummary.Summary) {
        # Convert markdown to HTML using improved converter
        $aiHtml = ConvertTo-HTMLFromMarkdown -Markdown $AIExecutiveSummary.Summary

        # Use string concatenation to avoid PowerShell interpolation of $ in AI content
        $html += '<div class="section page-break"><h2>Executive Briefing</h2><div class="section-intro">Strategic insights and executive-level recommendations for your migration initiative.</div><div class="ai-content">' + $aiHtml + '</div></div>'
    }

    # Migration Timeline
    $html += @"
<div class="section page-break keep-together">
    <h2>Recommended Migration Timeline</h2>
    <div class="section-intro">
        Based on the complexity assessment and identified issues, the following timeline is recommended for your migration project.
    </div>
    $timeline
</div>

"@

    # Assessment Summary with Risk Matrix
    $criticalCount = if ($AnalysisResults.BySeverity.Critical) { $AnalysisResults.BySeverity.Critical.Count } else { 0 }
    $highCount = if ($AnalysisResults.BySeverity.High) { $AnalysisResults.BySeverity.High.Count } else { 0 }
    $mediumCount = if ($AnalysisResults.BySeverity.Medium) { $AnalysisResults.BySeverity.Medium.Count } else { 0 }
    $lowCount = if ($AnalysisResults.BySeverity.Low) { $AnalysisResults.BySeverity.Low.Count } else { 0 }

    $execRiskMatrix = New-PDFRiskMatrix -Critical $criticalCount -High $highCount -Medium $mediumCount -Low $lowCount

    $html += @"
<div class="section page-break keep-together">
    <h2>Risk &amp; Complexity Assessment</h2>
    <div class="section-intro">
        Our comprehensive analysis has identified and categorized potential risks and complexities that may impact your migration.
    </div>
    $execRiskMatrix
    <table style="margin-top: 16pt;">
        <tr>
            <th>Category</th>
            <th>Status</th>
            <th>Action Required</th>
        </tr>
        <tr>
            <td>Critical Blockers</td>
            <td><span class="severity-badge severity-$(if($criticalCount -gt 0){'critical'}else{'low'})">$criticalCount Issues</span></td>
            <td>$(if($criticalCount -gt 0){"Must resolve before migration"}else{"No action needed"})</td>
        </tr>
        <tr>
            <td>High Priority Items</td>
            <td><span class="severity-badge severity-$(if($highCount -gt 0){'high'}else{'low'})">$highCount Issues</span></td>
            <td>$(if($highCount -gt 0){"Address during preparation phase"}else{"No action needed"})</td>
        </tr>
        <tr>
            <td>Overall Complexity</td>
            <td>$($ComplexityScore.ComplexityLevel)</td>
            <td>Score: $($ComplexityScore.TotalScore)/100</td>
        </tr>
        <tr>
            <td>Migration Readiness</td>
            <td><span style="color: $readinessColor; font-weight: 600;">$readinessScore%</span></td>
            <td>$readinessLevel</td>
        </tr>
    </table>
</div>

<div class="report-footer">
    <p><strong>CONFIDENTIAL - Executive Summary</strong></p>
    <p>Generated $(Get-Date -Format "MMMM dd, yyyy") | M365 Migration Assessment Tool</p>
</div>

</div>
</body>
</html>
"@

    # Generate PDF
    $pdfPath = New-PDFReport -HTMLContent $html -OutputPath $OutputPath -ReportType "Executive"

    return $pdfPath
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Get-HTMLHeader',
    'Get-HTMLFooter',
    'New-ITDetailedReport',
    'New-ExecutiveSummaryReport',
    'Get-PDFStylesheet',
    'ConvertTo-HTMLFromMarkdown',
    'New-PDFCoverPage',
    'New-PDFTOC',
    'New-PDFTimeline',
    'New-PDFRiskMatrix',
    'New-PDFReport',
    'New-ITDetailedReportPDF',
    'New-ExecutiveSummaryReportPDF'
)
