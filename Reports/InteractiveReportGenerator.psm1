#Requires -Version 7.0
<#
.SYNOPSIS
    Interactive HTML Report Generator for M365 Tenant Discovery
.DESCRIPTION
    Generates dashboard-style interactive HTML reports with:
    - Fixed sidebar navigation with multi-page layout
    - Glassmorphism KPI cards
    - Interactive charts (Chart.js)
    - Sortable tables, search, severity filters
    - Dark/light theme toggle
    - Responsive design with mobile sidebar overlay
.NOTES
    Author: AI Migration Expert
    Version: 3.0.0
    Target: PowerShell 7.x
#>

#region Private Helper Functions

function Get-SidebarHTML {
    param([string]$TenantName, [string]$ReportType = "IT", [array]$NavItems)
    $brandTitle = if ($ReportType -eq "Executive") { "Executive Summary" } else { "IT Technical Report" }
    $navHtml = ""
    $isFirst = $true
    foreach ($item in $NavItems) {
        $activeClass = if ($isFirst) { " active" } else { "" }
        $navHtml += "        <a class=`"nav-item$activeClass`" data-page=`"$($item.Id)`" onclick=`"navigateTo('$($item.Id)')`">$($item.Icon)<span>$($item.Label)</span></a>`n"
        $isFirst = $false
    }
    return @"
<aside class="sidebar" id="sidebar">
    <div class="sidebar-brand">
        <div class="brand-icon">M365</div>
        <div class="brand-text">
            <div class="brand-title">$brandTitle</div>
            <div class="brand-subtitle" title="$TenantName">$TenantName</div>
        </div>
    </div>
    <nav class="sidebar-nav">
$navHtml    </nav>
    <div class="sidebar-footer">
        <div>M365 Tenant Discovery</div>
        <div style="margin-top:4px;">$(Get-Date -Format 'MMM dd, yyyy')</div>
    </div>
</aside>
<div class="sidebar-overlay" id="sidebarOverlay" onclick="toggleSidebar()"></div>
"@
}

function Get-TopBarHTML {
    return @"
    <div class="top-bar">
        <button class="sidebar-toggle" onclick="toggleSidebar()" title="Toggle sidebar">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>
        </button>
        <div class="search-box">
            <input type="text" id="searchInput" placeholder="Search across all pages..." onkeyup="searchContent()">
        </div>
        <div class="filter-buttons">
            <button class="filter-btn active" onclick="filterBySeverity('all')">All</button>
            <button class="filter-btn" onclick="filterBySeverity('critical')">Critical</button>
            <button class="filter-btn" onclick="filterBySeverity('high')">High</button>
            <button class="filter-btn" onclick="filterBySeverity('medium')">Medium</button>
            <button class="filter-btn" onclick="filterBySeverity('low')">Low</button>
        </div>
        <button class="theme-toggle" onclick="toggleTheme()" title="Toggle theme">
            <span id="themeIcon">☀️</span>
            <span id="themeText">Light</span>
        </button>
    </div>
"@
}

function Get-KPICardHTML {
    param(
        [string]$Value, [string]$Label, [string]$IconSvg = "",
        [string]$Color = "#2563eb", [string]$Subtitle = ""
    )
    $subtitleHtml = if ($Subtitle) { "<div class='kpi-subtitle'>$Subtitle</div>" } else { "" }
    return @"
        <div class="kpi-card" style="--kpi-accent:$Color;">
            <div class="kpi-icon" style="background:$($Color)22;color:$Color;">$IconSvg</div>
            <div class="kpi-value">$Value</div>
            <div class="kpi-label">$Label</div>$subtitleHtml
        </div>
"@
}

function Get-ChartCardHTML {
    param([string]$Title, [string]$CanvasId, [string]$Height = "300px")
    return @"
        <div class="chart-card">
            <h3>$Title</h3>
            <div class="chart-container" style="height:$Height;"><canvas id="$CanvasId"></canvas></div>
        </div>
"@
}

function Convert-AIMarkdownToHTML {
    <#
    .SYNOPSIS
        Converts AI analysis markdown to well-structured HTML with proper list grouping
    #>
    param([string]$Markdown)

    if ([string]::IsNullOrWhiteSpace($Markdown)) { return "" }

    # Pre-process: restore newlines if content appears collapsed (all on one line)
    # This handles cases where newlines are lost during JSON serialization, batch combination, or API transport
    $lineCount = ($Markdown -split '\r?\n').Count
    if ($Markdown.Length -gt 500 -and $lineCount -lt 10) {
        # Content is long but has very few lines - likely collapsed markdown
        # Insert newlines before markdown structural patterns
        $restored = $Markdown

        # Restore newlines before headers (# ## ### ####) - but not inside code or URLs
        $restored = $restored -replace '(?<=\S)\s*(#{1,4}\s+)', "`n`n`$1"

        # Restore newlines before horizontal rules (--- or ***)
        $restored = $restored -replace '(?<=\S)\s*(---+)', "`n`n`$1`n"

        # Restore newlines before fenced code blocks (```)
        $restored = $restored -replace '(?<=\S)\s*(```)', "`n`$1"

        # Restore newlines before unordered list items (- item, * item) when preceded by non-list content
        # Be careful not to match hyphens in normal text - require space after hyphen
        $restored = $restored -replace '(?<=[\.\!\?\:\"]\s*)(-\s+\S)', "`n`$1"

        # Restore newlines before ordered list items (1. item, 2. item)
        $restored = $restored -replace '(?<=[\.\!\?\:\"\s])(\d+\.\s+)', "`n`$1"

        # Restore newlines before table rows
        $restored = $restored -replace '(?<=\S)\s*(\|[^\|]+\|)', "`n`$1"

        # Restore newlines before blockquotes
        $restored = $restored -replace '(?<=\S)\s*(>\s+)', "`n`$1"

        # Clean up any triple+ newlines into double
        $restored = $restored -replace '\n{3,}', "`n`n"

        $Markdown = $restored
    }

    $lines = $Markdown -split '\r?\n'
    $html = [System.Text.StringBuilder]::new()
    $inCodeBlock = $false
    $codeLang = ""
    $codeBuffer = [System.Text.StringBuilder]::new()
    $inUL = $false
    $inOL = $false
    $inTable = $false
    $tableStarted = $false

    foreach ($line in $lines) {
        # Handle fenced code blocks
        if ($line -match '^```(\w*)') {
            if ($inCodeBlock) {
                # Close code block
                [void]$html.Append("<pre><code>$([System.Net.WebUtility]::HtmlEncode($codeBuffer.ToString().TrimEnd()))</code></pre>`n")
                $codeBuffer.Clear()
                $inCodeBlock = $false
            } else {
                # Close any open lists first
                if ($inUL) { [void]$html.Append("</ul>`n"); $inUL = $false }
                if ($inOL) { [void]$html.Append("</ol>`n"); $inOL = $false }
                $inCodeBlock = $true
                $codeLang = $matches[1]
            }
            continue
        }
        if ($inCodeBlock) {
            [void]$codeBuffer.AppendLine($line)
            continue
        }

        # Handle markdown tables
        if ($line -match '^\|(.+)\|$') {
            # Close any open lists
            if ($inUL) { [void]$html.Append("</ul>`n"); $inUL = $false }
            if ($inOL) { [void]$html.Append("</ol>`n"); $inOL = $false }

            # Skip separator rows
            if ($line -match '^\|[\s:|-]+\|$') { continue }

            $cells = ($line -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
            if (-not $inTable) {
                [void]$html.Append("<table><thead><tr>")
                foreach ($cell in $cells) { [void]$html.Append("<th>$cell</th>") }
                [void]$html.Append("</tr></thead><tbody>`n")
                $inTable = $true
                $tableStarted = $true
            } else {
                [void]$html.Append("<tr>")
                foreach ($cell in $cells) { [void]$html.Append("<td>$cell</td>") }
                [void]$html.Append("</tr>`n")
            }
            continue
        } elseif ($inTable) {
            [void]$html.Append("</tbody></table>`n")
            $inTable = $false
            $tableStarted = $false
        }

        # Apply inline formatting
        $processed = $line
        $processed = $processed -replace '\*\*\*(.+?)\*\*\*', '<strong><em>$1</em></strong>'
        $processed = $processed -replace '\*\*(.+?)\*\*', '<strong>$1</strong>'
        $processed = $processed -replace '\*(.+?)\*', '<em>$1</em>'
        $processed = $processed -replace '`([^`]+)`', '<code>$1</code>'

        # Headers
        if ($processed -match '^#### (.+)$') {
            if ($inUL) { [void]$html.Append("</ul>`n"); $inUL = $false }
            if ($inOL) { [void]$html.Append("</ol>`n"); $inOL = $false }
            [void]$html.Append("<h4>$($matches[1])</h4>`n")
        }
        elseif ($processed -match '^### (.+)$') {
            if ($inUL) { [void]$html.Append("</ul>`n"); $inUL = $false }
            if ($inOL) { [void]$html.Append("</ol>`n"); $inOL = $false }
            [void]$html.Append("<h4>$($matches[1])</h4>`n")
        }
        elseif ($processed -match '^## (.+)$') {
            if ($inUL) { [void]$html.Append("</ul>`n"); $inUL = $false }
            if ($inOL) { [void]$html.Append("</ol>`n"); $inOL = $false }
            [void]$html.Append("<h3>$($matches[1])</h3>`n")
        }
        elseif ($processed -match '^# (.+)$') {
            if ($inUL) { [void]$html.Append("</ul>`n"); $inUL = $false }
            if ($inOL) { [void]$html.Append("</ol>`n"); $inOL = $false }
            [void]$html.Append("<h2>$($matches[1])</h2>`n")
        }
        # Blockquote
        elseif ($processed -match '^>\s?(.*)$') {
            if ($inUL) { [void]$html.Append("</ul>`n"); $inUL = $false }
            if ($inOL) { [void]$html.Append("</ol>`n"); $inOL = $false }
            [void]$html.Append("<blockquote>$($matches[1])</blockquote>`n")
        }
        # Horizontal rule
        elseif ($processed -match '^---+$' -or $processed -match '^\*\*\*+$') {
            if ($inUL) { [void]$html.Append("</ul>`n"); $inUL = $false }
            if ($inOL) { [void]$html.Append("</ol>`n"); $inOL = $false }
            [void]$html.Append("<hr>`n")
        }
        # Unordered list item
        elseif ($processed -match '^[-*+]\s+(.+)$') {
            if ($inOL) { [void]$html.Append("</ol>`n"); $inOL = $false }
            if (-not $inUL) { [void]$html.Append("<ul>`n"); $inUL = $true }
            [void]$html.Append("<li>$($matches[1])</li>`n")
        }
        # Ordered list item
        elseif ($processed -match '^\d+[\.\)]\s+(.+)$') {
            if ($inUL) { [void]$html.Append("</ul>`n"); $inUL = $false }
            if (-not $inOL) { [void]$html.Append("<ol>`n"); $inOL = $true }
            [void]$html.Append("<li>$($matches[1])</li>`n")
        }
        # Empty line
        elseif ($processed.Trim() -eq '') {
            if ($inUL) { [void]$html.Append("</ul>`n"); $inUL = $false }
            if ($inOL) { [void]$html.Append("</ol>`n"); $inOL = $false }
        }
        # Regular paragraph text
        else {
            if ($inUL) { [void]$html.Append("</ul>`n"); $inUL = $false }
            if ($inOL) { [void]$html.Append("</ol>`n"); $inOL = $false }
            [void]$html.Append("<p>$processed</p>`n")
        }
    }

    # Close any remaining open elements
    if ($inUL) { [void]$html.Append("</ul>`n") }
    if ($inOL) { [void]$html.Append("</ol>`n") }
    if ($inTable) { [void]$html.Append("</tbody></table>`n") }
    if ($inCodeBlock) {
        [void]$html.Append("<pre><code>$([System.Net.WebUtility]::HtmlEncode($codeBuffer.ToString().TrimEnd()))</code></pre>`n")
    }

    return $html.ToString()
}

#endregion

#region Interactive HTML Components

function Get-InteractiveHTMLHeader {
    <#
    .SYNOPSIS
        Returns enhanced HTML header with dashboard layout CSS
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [Parameter(Mandatory = $false)]
        [ValidateSet("IT", "Executive")]
        [string]$ReportType = "IT"
    )

    # Enterprise dark theme - electric blue accent on solid dark panels
    $primaryColor = "#3b82f6"
    $accentColor = "#60a5fa"
    $gradientEnd = "#1d4ed8"
    $primaryRGB = "59,130,246"

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$Title</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <style>
        :root {
            --primary: $primaryColor;
            --primary-light: $accentColor;
            --gradient-end: $gradientEnd;
            --success: #10b981; --success-light: #34d399;
            --warning: #f59e0b; --warning-light: #fbbf24;
            --danger: #ef4444; --danger-light: #f87171;
            --info: #06b6d4;
            --sidebar-width: 220px;
            --sidebar-bg: rgba(8,14,28,0.92);
            --sidebar-border: rgba(255,255,255,0.08);
            --glass-bg: rgba(16,28,52,0.65);
            --glass-border: rgba(255,255,255,0.1);
            --glass-blur: 14px;
            --topbar-height: 60px;
            --bg-body: #080e1c;
            --bg-card: rgba(16,28,52,0.65);
            --bg-card-hover: rgba(22,38,68,0.75);
            --bg-elevated: rgba(28,44,76,0.7);
            --bg-input: rgba(6,12,24,0.8);
            --bg-subtle: rgba(14,24,44,0.6);
            --border-subtle: rgba(255,255,255,0.05);
            --border-default: rgba(255,255,255,0.1);
            --border-strong: rgba(255,255,255,0.18);
            --text-primary: #e2e8f0;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --text-faint: #334466;
            --shadow-sm: 0 2px 8px rgba(0,0,0,0.5);
            --shadow-md: 0 4px 20px rgba(0,0,0,0.6);
            --shadow-lg: 0 8px 36px rgba(0,0,0,0.75);
            --radius-lg: 14px;
            --radius-md: 10px;
            --radius-sm: 8px;
            --transition-fast: 150ms ease;
            --transition-normal: 250ms ease;
        }
        [data-theme="light"] {
            --sidebar-bg: #1e293b;
            --sidebar-border: rgba(255,255,255,0.06);
            --glass-bg: #ffffff;
            --glass-border: rgba(0,0,0,0.08);
            --bg-body: #f1f5f9;
            --bg-card: #ffffff;
            --bg-card-hover: #f8fafc;
            --bg-elevated: #f1f5f9;
            --bg-input: #ffffff;
            --bg-subtle: #e2e8f0;
            --border-subtle: rgba(0,0,0,0.06);
            --border-default: rgba(0,0,0,0.1);
            --border-strong: rgba(0,0,0,0.16);
            --text-primary: #0f172a;
            --text-secondary: #334155;
            --text-muted: #64748b;
            --text-faint: #94a3b8;
            --shadow-sm: 0 1px 4px rgba(0,0,0,0.06);
            --shadow-md: 0 4px 12px rgba(0,0,0,0.08);
            --shadow-lg: 0 8px 24px rgba(0,0,0,0.1);
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            font-size: 14px; line-height: 1.6; color: var(--text-primary);
            background: var(--bg-body); min-height: 100vh;
            -webkit-font-smoothing: antialiased;
            transition: background-color var(--transition-normal), color var(--transition-normal);
        }
        body::before {
            content: ''; position: fixed; inset: 0; pointer-events: none; z-index: 0;
            background:
                radial-gradient(ellipse 60% 20% at 20% 0%, rgba($primaryRGB, 0.07) 0%, transparent 55%),
                radial-gradient(ellipse 40% 15% at 80% 0%, rgba(139,92,246,0.05) 0%, transparent 50%);
            transition: opacity var(--transition-normal);
        }

        /* Sidebar */
        .sidebar {
            position: fixed; top: 0; left: 0; bottom: 0;
            width: var(--sidebar-width); background: var(--sidebar-bg);
            backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px);
            border-right: 1px solid var(--sidebar-border);
            display: flex; flex-direction: column; z-index: 200;
            transition: width 0.3s ease, transform 0.3s ease;
            overflow: hidden;
        }
        .sidebar.collapsed { width: 72px; }
        .sidebar.collapsed .brand-text,
        .sidebar.collapsed .nav-item span,
        .sidebar.collapsed .sidebar-footer { display: none; }
        .sidebar.collapsed .nav-item { justify-content: center; padding: 10px; }
        .sidebar.collapsed .sidebar-brand { justify-content: center; }
        .sidebar-brand {
            padding: 20px; display: flex; align-items: center; gap: 12px;
            border-bottom: 1px solid var(--sidebar-border); flex-shrink: 0;
        }
        .brand-icon {
            width: 40px; height: 40px; border-radius: 10px;
            background: linear-gradient(135deg, var(--primary), var(--gradient-end));
            display: flex; align-items: center; justify-content: center;
            color: white; font-weight: 800; font-size: 0.7rem; flex-shrink: 0;
        }
        .brand-title { font-size: 0.85rem; font-weight: 600; color: var(--text-primary); white-space: nowrap; }
        .brand-subtitle {
            font-size: 0.72rem; color: var(--text-muted); white-space: nowrap;
            overflow: hidden; text-overflow: ellipsis; max-width: 160px;
        }
        .sidebar-nav { flex: 1; padding: 12px 8px; overflow-y: auto; }
        .nav-item {
            display: flex; align-items: center; gap: 12px; padding: 10px 16px;
            border-radius: 10px; color: var(--text-secondary); cursor: pointer;
            transition: all 0.15s ease; margin-bottom: 2px; text-decoration: none;
            font-size: 0.88rem; white-space: nowrap;
        }
        .nav-item:hover { background: rgba(255,255,255,0.05); color: var(--text-primary); }
        .nav-item.active {
            background: rgba($primaryRGB, 0.18); color: #fff;
            font-weight: 600; border: 1px solid rgba($primaryRGB, 0.3);
        }
        .nav-item.active svg { color: var(--primary-light); }
        .nav-item svg { flex-shrink: 0; }
        .sidebar-footer {
            padding: 16px 20px; border-top: 1px solid var(--sidebar-border);
            font-size: 0.72rem; color: var(--text-muted); flex-shrink: 0;
        }
        .sidebar-overlay {
            display: none; position: fixed; inset: 0;
            background: rgba(0,0,0,0.5); z-index: 199;
        }
        .sidebar-overlay.active { display: block; }
        [data-theme="light"] .nav-item { color: rgba(255,255,255,0.6); }
        [data-theme="light"] .nav-item:hover { background: rgba(255,255,255,0.08); color: #fff; }
        [data-theme="light"] .nav-item.active { background: rgba($primaryRGB, 0.25); color: #fff; border-color: rgba($primaryRGB, 0.4); }

        /* Main Content */
        .main-content {
            margin-left: var(--sidebar-width); min-height: 100vh;
            transition: margin-left 0.3s ease; position: relative; z-index: 1;
        }
        .main-content.sidebar-collapsed { margin-left: 72px; }

        /* Top Bar */
        .top-bar {
            position: sticky; top: 0; z-index: 100; height: var(--topbar-height);
            background: rgba(8,14,28,0.75);
            backdrop-filter: blur(var(--glass-blur)); -webkit-backdrop-filter: blur(var(--glass-blur));
            border-bottom: 1px solid var(--border-default);
            display: flex; align-items: center; gap: 12px; padding: 0 24px;
        }
        .sidebar-toggle {
            background: none; border: 1px solid var(--border-default);
            border-radius: var(--radius-sm); padding: 8px; cursor: pointer;
            color: var(--text-secondary); display: flex; align-items: center;
            transition: all var(--transition-fast);
        }
        .sidebar-toggle:hover { background: var(--bg-elevated); color: var(--text-primary); }
        .search-box { flex: 1; min-width: 200px; position: relative; }
        .search-box input {
            width: 100%; padding: 8px 14px 8px 36px; background: var(--bg-input);
            border: 1px solid var(--border-default); border-radius: var(--radius-sm);
            color: var(--text-primary); font-size: 13px; transition: all var(--transition-fast);
        }
        .search-box input:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba($primaryRGB, 0.1); }
        .search-box::before {
            content: '🔍'; position: absolute; left: 10px; top: 50%;
            transform: translateY(-50%); opacity: 0.5; font-size: 13px;
        }
        .filter-buttons { display: flex; gap: 6px; flex-wrap: wrap; }
        .filter-btn {
            padding: 6px 12px; background: var(--bg-elevated); border: 1px solid var(--border-default);
            border-radius: var(--radius-sm); color: var(--text-secondary); cursor: pointer;
            font-size: 12px; transition: all var(--transition-fast);
        }
        .filter-btn:hover { background: var(--bg-card-hover); color: var(--text-primary); }
        .filter-btn.active { background: var(--primary); color: white; border-color: var(--primary); }
        .theme-toggle {
            padding: 8px 14px; background: var(--bg-elevated); border: 1px solid var(--border-default);
            border-radius: var(--radius-sm); color: var(--text-primary); cursor: pointer;
            font-size: 13px; display: flex; align-items: center; gap: 6px;
            transition: all var(--transition-fast); white-space: nowrap;
        }
        .theme-toggle:hover { background: var(--bg-card-hover); border-color: var(--primary); }

        /* Page System */
        .page-content { padding: 24px; }
        .page { display: none; }
        .page.active { display: block; }
        .page-header { margin-bottom: 24px; }
        .page-header h1 { font-size: 1.5rem; font-weight: 700; color: var(--text-primary); }
        .page-header p { color: var(--text-muted); font-size: 0.88rem; margin-top: 4px; }

        /* KPI Cards */
        .kpi-grid {
            display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px; margin-bottom: 24px;
        }
        .kpi-card {
            background: var(--glass-bg);
            backdrop-filter: blur(var(--glass-blur)); -webkit-backdrop-filter: blur(var(--glass-blur));
            border: 1px solid var(--glass-border); border-radius: var(--radius-lg);
            padding: 20px 22px; transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
            position: relative; overflow: hidden;
        }
        .kpi-card::before {
            content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px;
            background: var(--kpi-accent, var(--primary)); opacity: 0.7; border-radius: var(--radius-lg) var(--radius-lg) 0 0;
        }
        .kpi-card:hover { transform: translateY(-3px); box-shadow: var(--shadow-md); border-color: var(--border-strong); }
        .kpi-icon {
            width: 40px; height: 40px; border-radius: 10px;
            display: flex; align-items: center; justify-content: center; margin-bottom: 14px;
        }
        .kpi-icon svg { width: 20px; height: 20px; }
        .kpi-value { font-size: 2rem; font-weight: 700; line-height: 1.1; color: var(--text-primary); }
        .kpi-label {
            font-size: 0.7rem; font-weight: 600; text-transform: uppercase;
            letter-spacing: 0.1em; color: var(--text-muted); margin-top: 6px;
        }
        .kpi-subtitle { font-size: 0.76rem; color: var(--text-faint); margin-top: 3px; }

        /* Chart Cards */
        .chart-grid {
            display: grid; grid-template-columns: repeat(auto-fit, minmax(380px, 1fr));
            gap: 20px; margin-bottom: 24px;
        }
        .chart-card {
            background: var(--glass-bg);
            backdrop-filter: blur(var(--glass-blur)); -webkit-backdrop-filter: blur(var(--glass-blur));
            border: 1px solid var(--glass-border); border-radius: var(--radius-lg); padding: 24px;
            transition: box-shadow var(--transition-normal), border-color var(--transition-normal);
        }
        .chart-card:hover {
            box-shadow: 0 0 0 1px rgba($primaryRGB, 0.2), var(--shadow-md);
            border-color: rgba($primaryRGB, 0.25);
        }
        .chart-card h3 {
            font-size: 0.88rem; font-weight: 600; color: var(--text-secondary);
            text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 16px;
        }
        .chart-container { position: relative; }

        /* Tables */
        table { width: 100%; border-collapse: collapse; margin: 16px 0; }
        th {
            background: var(--bg-elevated); color: var(--text-muted); font-weight: 600;
            font-size: 0.7rem; text-transform: uppercase; letter-spacing: 0.08em;
            text-align: left; padding: 10px 16px; border-bottom: 1px solid var(--border-strong);
            position: relative; cursor: pointer; user-select: none; transition: all var(--transition-fast);
        }
        th:hover { background: var(--bg-card-hover); color: var(--text-primary); }
        th.sortable::after { content: '⇅'; position: absolute; right: 8px; opacity: 0.3; font-size: 11px; }
        th.sorted-asc::after { content: '▲'; opacity: 1; color: var(--primary); }
        th.sorted-desc::after { content: '▼'; opacity: 1; color: var(--primary); }
        td { padding: 13px 16px; border-bottom: 1px solid var(--border-subtle); color: var(--text-secondary); }
        tr:last-child td { border-bottom: none; }
        tr:hover td { background: var(--bg-card-hover); color: var(--text-primary); }
        .data-section {
            background: var(--glass-bg);
            backdrop-filter: blur(var(--glass-blur)); -webkit-backdrop-filter: blur(var(--glass-blur));
            border: 1px solid var(--glass-border); border-radius: var(--radius-lg); padding: 24px; margin-bottom: 20px;
        }
        .data-section h3 {
            font-size: 0.88rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em;
            color: var(--text-muted); margin-bottom: 16px;
        }

        /* Severity Badges */
        .severity-badge {
            display: inline-flex; align-items: center; gap: 6px;
            padding: 4px 10px; border-radius: 6px;
            font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em;
        }
        .severity-badge::before {
            content: ''; width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0;
        }
        .severity-critical { background: rgba(239,68,68,0.12); color: #f87171; border: 1px solid rgba(239,68,68,0.25); }
        .severity-critical::before { background: #ef4444; box-shadow: 0 0 6px rgba(239,68,68,0.7); }
        .severity-high { background: rgba(249,115,22,0.12); color: #fb923c; border: 1px solid rgba(249,115,22,0.25); }
        .severity-high::before { background: #f97316; box-shadow: 0 0 6px rgba(249,115,22,0.7); }
        .severity-medium { background: rgba(245,158,11,0.12); color: #fbbf24; border: 1px solid rgba(245,158,11,0.25); }
        .severity-medium::before { background: #f59e0b; box-shadow: 0 0 6px rgba(245,158,11,0.7); }
        .severity-low { background: rgba(16,185,129,0.12); color: #34d399; border: 1px solid rgba(16,185,129,0.25); }
        .severity-low::before { background: #10b981; box-shadow: 0 0 6px rgba(16,185,129,0.7); }
        .severity-informational { background: rgba(6,182,212,0.12); color: #22d3ee; border: 1px solid rgba(6,182,212,0.25); }
        .severity-informational::before { background: #06b6d4; box-shadow: 0 0 6px rgba(6,182,212,0.7); }

        /* Issue Cards & Checklist */
        .severity-group { margin-bottom: 32px; }
        .severity-group-header {
            padding: 16px 20px; border-radius: var(--radius-md);
            margin-bottom: 16px; display: flex; justify-content: space-between; align-items: center;
        }
        .issue-card {
            background: var(--glass-bg);
            backdrop-filter: blur(var(--glass-blur)); -webkit-backdrop-filter: blur(var(--glass-blur));
            border: 1px solid var(--glass-border);
            border-radius: var(--radius-md); padding: 16px; margin-bottom: 10px; transition: all 0.2s ease;
        }
        .issue-card:hover { border-color: rgba($primaryRGB, 0.4); box-shadow: 0 0 0 1px rgba($primaryRGB, 0.15), var(--shadow-sm); }
        .issue-card.completed { opacity: 0.5; }
        .checkbox-wrapper { flex-shrink: 0; cursor: pointer; }
        .checkbox-wrapper input[type="checkbox"] { display: none; }
        .checkmark {
            width: 24px; height: 24px; border: 2px solid var(--text-muted);
            border-radius: 6px; display: flex; align-items: center; justify-content: center;
            transition: all 0.2s ease;
        }
        .checkmark svg { width: 14px; height: 14px; opacity: 0; transition: opacity 0.2s ease; }
        .checkbox-wrapper input:checked + .checkmark {
            background: var(--success); border-color: var(--success);
        }
        .checkbox-wrapper input:checked + .checkmark svg { opacity: 1; }
        .progress-bar-track {
            height: 6px; background: var(--bg-elevated); border-radius: 3px; overflow: hidden;
        }
        .progress-bar-fill { height: 100%; transition: width 0.3s ease; border-radius: 3px; }

        /* AI Content */
        .ai-card {
            background: var(--glass-bg);
            backdrop-filter: blur(var(--glass-blur)); -webkit-backdrop-filter: blur(var(--glass-blur));
            border: 1px solid var(--glass-border);
            border-radius: var(--radius-lg); padding: 32px;
            overflow-wrap: break-word; word-wrap: break-word; word-break: break-word;
        }
        .ai-card h2 {
            color: var(--primary); font-size: 1.3rem; font-weight: 700;
            margin: 32px 0 12px; padding-bottom: 8px;
            border-bottom: 1px solid var(--border-default);
        }
        .ai-card h3 {
            color: var(--text-primary); font-size: 1.1rem; font-weight: 600;
            margin: 24px 0 10px;
        }
        .ai-card h4 {
            color: var(--text-secondary); font-size: 1rem; font-weight: 600;
            margin: 20px 0 8px;
        }
        .ai-card p {
            color: var(--text-secondary); line-height: 1.8; margin-bottom: 12px;
        }
        .ai-card ul {
            list-style: none; padding: 0; margin: 12px 0 16px 0;
        }
        .ai-card ul li {
            color: var(--text-secondary); line-height: 1.7; padding: 6px 0 6px 24px;
            position: relative; border-bottom: 1px solid var(--border-subtle);
        }
        .ai-card ul li:last-child { border-bottom: none; }
        .ai-card ul li::before {
            content: ''; position: absolute; left: 8px; top: 14px;
            width: 6px; height: 6px; border-radius: 50%;
            background: var(--primary); opacity: 0.6;
        }
        .ai-card ol {
            list-style: none; padding: 0; margin: 12px 0 16px 0;
            counter-reset: ai-step;
        }
        .ai-card ol li {
            color: var(--text-secondary); line-height: 1.7;
            padding: 10px 12px 10px 16px; margin-bottom: 4px;
            background: var(--glass-bg); border: 1px solid var(--border-subtle);
            border-radius: 8px; display: flex; align-items: flex-start; gap: 12px;
            counter-increment: ai-step; cursor: pointer;
            transition: all 0.15s ease;
        }
        .ai-card ol li:hover {
            border-color: var(--primary); background: rgba($primaryRGB, 0.04);
        }
        .ai-card ol li::before {
            content: counter(ai-step); flex-shrink: 0;
            width: 26px; height: 26px; border-radius: 50%;
            background: rgba($primaryRGB, 0.12); color: var(--primary);
            font-size: 0.78rem; font-weight: 700;
            display: flex; align-items: center; justify-content: center;
            margin-top: 1px;
        }
        .ai-card ol li.checked {
            opacity: 0.5; text-decoration: line-through;
            text-decoration-color: var(--text-muted);
        }
        .ai-card ol li.checked::before {
            background: var(--success); color: white; content: '\2713';
        }
        .ai-card pre {
            background: rgba(0,0,0,0.3); border: 1px solid var(--border-default);
            border-radius: 10px; padding: 16px 20px; margin: 12px 0 16px 0;
            overflow-x: auto; font-size: 0.85rem; line-height: 1.5;
        }
        .ai-card pre code {
            background: none; padding: 0; color: #c9d1d9;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
        }
        .ai-card code {
            background: var(--bg-elevated); padding: 2px 6px; border-radius: 4px;
            font-size: 0.85rem; color: var(--primary);
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
        }
        .ai-card table {
            width: 100%; border-collapse: collapse; margin: 12px 0 16px 0;
            font-size: 0.88rem;
        }
        .ai-card table th {
            background: rgba($primaryRGB, 0.08); color: var(--primary);
            font-weight: 600; padding: 10px 14px; text-align: left;
            border-bottom: 2px solid rgba($primaryRGB, 0.2);
        }
        .ai-card table td {
            padding: 10px 14px; border-bottom: 1px solid var(--border-subtle);
            color: var(--text-secondary);
        }
        .ai-card table tr:hover td { background: rgba($primaryRGB, 0.03); }
        .ai-card blockquote {
            border-left: 3px solid var(--primary); background: rgba($primaryRGB, 0.05);
            padding: 12px 16px; margin: 12px 0; border-radius: 0 8px 8px 0;
            color: var(--text-secondary); font-style: italic;
        }
        .ai-card hr {
            border: none; border-top: 1px solid var(--border-default);
            margin: 24px 0;
        }
        /* Gauge */
        .gauge-container { text-align: center; padding: 20px; }
        .gauge-value { font-size: 2.5rem; font-weight: 700; }
        .gauge-label { font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; }
        .gauge-status {
            margin-top: 16px; padding: 10px 20px; background: var(--bg-elevated);
            border-radius: 10px; display: inline-block; font-size: 0.9rem; font-weight: 600;
        }

        /* Animations */
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }

        /* Print */
        @media print {
            .sidebar, .sidebar-overlay, .top-bar, .theme-toggle, .filter-buttons, .sidebar-toggle { display: none !important; }
            .main-content { margin-left: 0 !important; }
            .page { display: block !important; page-break-before: auto; }
            .page-content { padding: 0; }
            body::before { display: none; }
            .kpi-card { box-shadow: none; backdrop-filter: none; border: 1px solid #ddd; }
            .kpi-card:hover { transform: none; }
        }

        /* Responsive */
        @media (max-width: 1024px) {
            .sidebar { transform: translateX(-100%); }
            .sidebar.mobile-open { transform: translateX(0); }
            .main-content { margin-left: 0 !important; }
            .top-bar { padding: 0 16px; }
            .page-content { padding: 16px; }
            .kpi-grid { grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); }
            .chart-grid { grid-template-columns: 1fr; }
            .filter-buttons { display: none; }
        }
    </style>
<script>
// Chart helpers - defined early so inline scripts can call them
var chartRetries = 0;
var maxChartRetries = 10;

function initChartColors() {
    return {
        critical: 'rgb(107,114,128)', high: 'rgb(120,113,108)',
        medium: 'rgb(156,163,175)', low: 'rgb(209,213,219)',
        info: 'rgb(229,231,235)',
        primary: getComputedStyle(document.documentElement).getPropertyValue('--primary').trim()
    };
}

function getThemeColor(prop) {
    return getComputedStyle(document.documentElement).getPropertyValue(prop).trim();
}

function isDarkMode() {
    return document.documentElement.getAttribute('data-theme') !== 'light';
}

function colorToRGBA(c, alpha) {
    if (c.startsWith('rgba')) return c.replace(/,\s*[\d.]+\)/, ',' + alpha + ')');
    if (c.startsWith('rgb(')) return c.replace('rgb(', 'rgba(').replace(')', ',' + alpha + ')');
    if (c.startsWith('#')) {
        var r = parseInt(c.slice(1,3),16), g = parseInt(c.slice(3,5),16), b = parseInt(c.slice(5,7),16);
        return 'rgba(' + r + ',' + g + ',' + b + ',' + alpha + ')';
    }
    return c;
}

function glowPalette(colors) {
    if (!isDarkMode()) return { bg: colors, border: colors, hover: colors };
    return {
        bg: colors.map(function(c) { return colorToRGBA(c, 0.5); }),
        border: colors.map(function(c) { return colorToRGBA(c, 0.85); }),
        hover: colors.map(function(c) { return colorToRGBA(c, 0.75); })
    };
}

var softGlowPlugin = {
    id: 'softGlow',
    beforeDatasetsDraw: function(chart) {
        if (!isDarkMode()) return;
        chart.ctx.save();
        chart.ctx.shadowBlur = 28;
        chart.ctx.shadowColor = 'rgba(56, 189, 248, 0.35)';
    },
    afterDatasetsDraw: function(chart) {
        if (!isDarkMode()) return;
        chart.ctx.restore();
    }
};

function createDoughnutChart(canvasId, labels, data, colors) {
    if (typeof Chart === 'undefined') {
        if (chartRetries < maxChartRetries) { chartRetries++; setTimeout(function() { createDoughnutChart(canvasId, labels, data, colors); }, 200); }
        else { showChartFallback(canvasId, labels, data); }
        return;
    }
    var canvas = document.getElementById(canvasId);
    if (!canvas) return;
    if (!data.some(function(v) { return v > 0; })) {
        canvas.parentElement.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:var(--text-muted);">No data available</div>';
        return;
    }
    var dark = isDarkMode();
    var glow = glowPalette(colors);
    try {
        new Chart(canvas, {
            type: 'doughnut',
            data: { labels: labels, datasets: [{ data: data, backgroundColor: glow.bg, borderColor: dark ? glow.border : (getThemeColor('--bg-card') || '#fff'), borderWidth: 2, hoverBackgroundColor: glow.hover }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: getThemeColor('--text-primary') || '#fff', padding: 15, font: { size: 12 } } } } },
            plugins: dark ? [softGlowPlugin] : []
        });
    } catch (e) { showChartFallback(canvasId, labels, data); }
}

function createBarChart(canvasId, labels, data, label, color) {
    if (typeof Chart === 'undefined') { if (chartRetries < maxChartRetries) { chartRetries++; setTimeout(function() { createBarChart(canvasId, labels, data, label, color); }, 200); } return; }
    var canvas = document.getElementById(canvasId);
    if (!canvas) return;
    var dark = isDarkMode();
    var bgColor = dark ? colorToRGBA(color, 0.45) : color;
    var bdColor = dark ? colorToRGBA(color, 0.8) : color;
    try {
        new Chart(canvas, {
            type: 'bar',
            data: { labels: labels, datasets: [{ label: label, data: data, backgroundColor: bgColor, borderColor: bdColor, borderWidth: dark ? 1 : 0, borderRadius: 6, hoverBackgroundColor: dark ? colorToRGBA(color, 0.7) : color }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } },
                scales: { y: { beginAtZero: true, ticks: { color: getThemeColor('--text-secondary') }, grid: { color: getThemeColor('--border-subtle') } },
                          x: { ticks: { color: getThemeColor('--text-secondary') }, grid: { color: getThemeColor('--border-subtle') } } } },
            plugins: dark ? [softGlowPlugin] : []
        });
    } catch (e) { console.error('Bar chart failed:', e); }
}

function createHorizontalBarChart(canvasId, labels, data, colors) {
    if (typeof Chart === 'undefined') { if (chartRetries < maxChartRetries) { chartRetries++; setTimeout(function() { createHorizontalBarChart(canvasId, labels, data, colors); }, 200); } return; }
    var canvas = document.getElementById(canvasId);
    if (!canvas) return;
    var dark = isDarkMode();
    var glow = glowPalette(colors);
    try {
        new Chart(canvas, {
            type: 'bar',
            data: { labels: labels, datasets: [{ data: data, backgroundColor: glow.bg, borderColor: dark ? glow.border : 'transparent', borderWidth: dark ? 1 : 0, borderRadius: 6, hoverBackgroundColor: glow.hover }] },
            options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y',
                plugins: { legend: { display: false } },
                scales: { x: { beginAtZero: true, grid: { color: getThemeColor('--border-subtle') }, ticks: { color: getThemeColor('--text-secondary') } },
                          y: { grid: { display: false }, ticks: { color: getThemeColor('--text-secondary') } } } },
            plugins: dark ? [softGlowPlugin] : []
        });
    } catch (e) { console.error('Horizontal bar chart failed:', e); }
}

function createRadarChart(canvasId, labels, data, label, color) {
    if (typeof Chart === 'undefined') { if (chartRetries < maxChartRetries) { chartRetries++; setTimeout(function() { createRadarChart(canvasId, labels, data, label, color); }, 200); } return; }
    var canvas = document.getElementById(canvasId);
    if (!canvas) return;
    var dark = isDarkMode();
    var fillColor = colorToRGBA(color, dark ? 0.12 : 0.2);
    var edgeColor = dark ? colorToRGBA(color, 0.85) : color;
    var pointColor = dark ? colorToRGBA(color, 0.9) : color;
    try {
        new Chart(canvas, {
            type: 'radar',
            data: { labels: labels, datasets: [{ label: label, data: data, backgroundColor: fillColor,
                borderColor: edgeColor, borderWidth: 2, pointBackgroundColor: pointColor, pointBorderColor: dark ? 'rgba(255,255,255,0.5)' : '#fff', pointRadius: 4, pointHoverRadius: 6 }] },
            options: { responsive: true, maintainAspectRatio: false,
                scales: { r: { beginAtZero: true, ticks: { color: getThemeColor('--text-muted'), backdropColor: 'transparent' },
                    grid: { color: getThemeColor('--border-default') }, pointLabels: { color: getThemeColor('--text-secondary'), font: { size: 11 } } } },
                plugins: { legend: { display: false } } },
            plugins: dark ? [softGlowPlugin] : []
        });
    } catch (e) { showChartFallback(canvasId, labels, data); }
}

function showChartFallback(canvasId, labels, data) {
    var canvas = document.getElementById(canvasId);
    if (!canvas) return;
    var html = '<div style="padding:20px;font-size:0.85rem;">';
    html += '<div style="color:var(--text-muted);margin-bottom:12px;">Chart unavailable - Data summary:</div>';
    html += '<ul style="list-style:none;padding:0;">';
    for (var i = 0; i < labels.length; i++) {
        if (data[i] > 0) {
            html += '<li style="padding:6px 0;border-bottom:1px solid var(--border-subtle);display:flex;justify-content:space-between;">';
            html += '<span>' + labels[i] + '</span><strong>' + data[i] + '</strong></li>';
        }
    }
    html += '</ul></div>';
    canvas.parentElement.innerHTML = html;
}
</script>
</head>
<body data-theme="dark">
"@
}

function Get-InteractiveHTMLFooter {
    <#
    .SYNOPSIS
        Returns interactive HTML footer with all JavaScript
    #>
    return @"
</main>
</div>

<script>
// ===== Navigation =====
function navigateTo(pageId) {
    document.querySelectorAll('.page').forEach(function(p) { p.classList.remove('active'); });
    document.querySelectorAll('.nav-item').forEach(function(n) { n.classList.remove('active'); });
    var page = document.getElementById(pageId);
    if (page) page.classList.add('active');
    var navItem = document.querySelector('[data-page="' + pageId + '"]');
    if (navItem) navItem.classList.add('active');
    var sidebar = document.getElementById('sidebar');
    var overlay = document.getElementById('sidebarOverlay');
    if (sidebar) sidebar.classList.remove('mobile-open');
    if (overlay) overlay.classList.remove('active');
    window.scrollTo(0, 0);
}

// ===== Sidebar Toggle =====
function toggleSidebar() {
    var sidebar = document.getElementById('sidebar');
    var overlay = document.getElementById('sidebarOverlay');
    if (window.innerWidth <= 1024) {
        sidebar.classList.toggle('mobile-open');
        overlay.classList.toggle('active');
    } else {
        sidebar.classList.toggle('collapsed');
        document.querySelector('.main-content').classList.toggle('sidebar-collapsed');
    }
}

// ===== Theme Toggle =====
function toggleTheme() {
    var body = document.body;
    var newTheme = body.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
    body.setAttribute('data-theme', newTheme);
    localStorage.setItem('theme', newTheme);
    document.getElementById('themeIcon').textContent = newTheme === 'light' ? '🌙' : '☀️';
    document.getElementById('themeText').textContent = newTheme === 'light' ? 'Dark' : 'Light';
}

// ===== Search =====
function searchContent() {
    var filter = document.getElementById('searchInput').value.toLowerCase();
    var pages = document.querySelectorAll('.page');
    if (!filter) {
        pages.forEach(function(p) { p.querySelectorAll('[data-searchable]').forEach(function(el) { el.style.display = ''; }); });
        return;
    }
    pages.forEach(function(page) {
        page.querySelectorAll('[data-searchable]').forEach(function(el) {
            var text = (el.textContent || el.innerText).toLowerCase();
            el.style.display = text.indexOf(filter) > -1 ? '' : 'none';
        });
    });
}

// ===== Severity Filter =====
function filterBySeverity(severity) {
    var buttons = document.querySelectorAll('.filter-btn');
    buttons.forEach(function(btn) { btn.classList.remove('active'); });
    event.target.classList.add('active');
    var cards = document.querySelectorAll('.issue-card');
    cards.forEach(function(card) {
        if (severity === 'all') { card.style.display = ''; }
        else {
            var hasSeverity = card.classList.contains(severity) || card.getAttribute('data-severity') === severity;
            card.style.display = hasSeverity ? '' : 'none';
        }
    });
    var groups = document.querySelectorAll('.severity-group');
    groups.forEach(function(group) {
        if (severity === 'all') { group.style.display = ''; }
        else { group.style.display = group.getAttribute('data-severity') === severity ? '' : 'none'; }
    });
}

// ===== Sortable Tables =====
function sortTable(table, column, asc) {
    var tbody = table.querySelector('tbody');
    if (!tbody) return;
    var rows = Array.from(tbody.querySelectorAll('tr'));
    rows.sort(function(a, b) {
        var aVal = a.cells[column] ? a.cells[column].textContent.trim() : '';
        var bVal = b.cells[column] ? b.cells[column].textContent.trim() : '';
        var aNum = parseFloat(aVal.replace(/[^0-9.\-]/g, ''));
        var bNum = parseFloat(bVal.replace(/[^0-9.\-]/g, ''));
        if (!isNaN(aNum) && !isNaN(bNum)) return asc ? aNum - bNum : bNum - aNum;
        return asc ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
    });
    rows.forEach(function(row) { tbody.appendChild(row); });
}

// ===== Checklist Progress =====
function updateProgress(severity, total) {
    var checkboxes = document.querySelectorAll('.issue-checkbox[data-severity="' + severity + '"]');
    var checked = 0;
    checkboxes.forEach(function(cb) { if (cb.checked) checked++; });
    var el = document.getElementById('progress-' + severity);
    if (el) el.textContent = checked;
    var bar = document.getElementById('progressbar-' + severity);
    if (bar) bar.style.width = (checked / total * 100) + '%';
}

function toggleCardComplete(checkbox, cardId) {
    var card = document.getElementById(cardId);
    if (!card) return;
    var checkmark = checkbox.nextElementSibling;
    if (checkbox.checked) {
        card.classList.add('completed');
        if (checkmark) { checkmark.style.background = 'var(--success)'; checkmark.style.borderColor = 'var(--success)'; }
        if (checkmark && checkmark.querySelector('svg')) checkmark.querySelector('svg').style.opacity = '1';
    } else {
        card.classList.remove('completed');
        if (checkmark) { checkmark.style.background = 'transparent'; checkmark.style.borderColor = ''; }
        if (checkmark && checkmark.querySelector('svg')) checkmark.querySelector('svg').style.opacity = '0';
    }
}

// ===== Issue Detail Toggle =====
function toggleIssueDetail(detailId) {
    var detail = document.getElementById(detailId);
    var arrow = document.getElementById('arrow-' + detailId);
    if (!detail) return;
    if (detail.style.display === 'none') {
        detail.style.display = 'block';
        if (arrow) arrow.style.transform = 'rotate(90deg)';
    } else {
        detail.style.display = 'none';
        if (arrow) arrow.style.transform = 'rotate(0deg)';
    }
}

// ===== Init on DOM Ready =====
document.addEventListener('DOMContentLoaded', function() {
    var savedTheme = localStorage.getItem('theme') || 'dark';
    if (savedTheme === 'light') toggleTheme();

    document.querySelectorAll('table').forEach(function(table) {
        var headers = table.querySelectorAll('th');
        headers.forEach(function(header, index) {
            header.classList.add('sortable');
            var asc = true;
            header.addEventListener('click', function() {
                headers.forEach(function(h) { h.classList.remove('sorted-asc', 'sorted-desc'); });
                header.classList.add(asc ? 'sorted-asc' : 'sorted-desc');
                sortTable(table, index, asc);
                asc = !asc;
            });
        });
    });

    // Checkbox toggling on ordered list items in AI cards
    document.querySelectorAll('.ai-card ol li').forEach(function(li) {
        li.addEventListener('click', function(e) {
            if (e.target.tagName === 'A') return; // Don't toggle when clicking links
            this.classList.toggle('checked');
        });
    });
});
</script>
</body>
</html>
"@
}

#endregion

#region Complete Report Functions

function New-InteractiveITReport {
    <#
    .SYNOPSIS
        Generates a dashboard-style interactive IT Detailed Report
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

    Write-Host "Generating Interactive IT Report..." -ForegroundColor Cyan

    # SVG Icons
    $svgDash = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>'
    $svgIdentity = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>'
    $svgExchange = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/></svg>'
    $svgSP = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/><polyline points="14 2 14 8 20 8"/></svg>'
    $svgTeams = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>'
    $svgSecurity = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>'
    $svgIssues = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>'
    $svgAI = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="4"/><path d="M12 2v4m0 12v4M4.93 4.93l2.83 2.83m8.48 8.48 2.83 2.83M2 12h4m12 0h4M4.93 19.07l2.83-2.83m8.48-8.48 2.83-2.83"/></svg>'
    $svgUsers = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/></svg>'
    $svgMail = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/></svg>'
    $svgFolder = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>'
    $svgCloud = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 10h-1.26A8 8 0 1 0 9 20h9a5 5 0 0 0 0-10z"/></svg>'
    $svgChat = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>'
    $svgGroup = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>'
    $svgShield = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>'
    $svgDb = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/></svg>'
    $svgLock = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>'
    $svgApp = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="2" width="8" height="8" rx="1"/><rect x="14" y="2" width="8" height="8" rx="1"/><rect x="2" y="14" width="8" height="8" rx="1"/><rect x="14" y="14" width="8" height="8" rx="1"/></svg>'
    $svgDynamics = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2L2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/></svg>'
    $svgPowerBI = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="12" width="4" height="9" rx="1"/><rect x="10" y="7" width="4" height="14" rx="1"/><rect x="17" y="3" width="4" height="18" rx="1"/></svg>'
    $svgPhone = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92z"/></svg>'

    $tenantName = $CollectedData.TenantInfo.DisplayName
    $tenantId = $CollectedData.TenantInfo.TenantId

    # === Extract data with null-safety ===
    $totalUsers = if ($CollectedData.EntraID.Users.Analysis) { $CollectedData.EntraID.Users.Analysis.TotalUsers } else { 0 }
    $licensedUsers = if ($CollectedData.EntraID.Users.Analysis) { $CollectedData.EntraID.Users.Analysis.LicensedUsers } else { 0 }
    $guestUsers = if ($CollectedData.EntraID.Users.Analysis) { $CollectedData.EntraID.Users.Analysis.GuestUsers } else { 0 }
    $syncedUsers = if ($CollectedData.EntraID.Users.Analysis) { $CollectedData.EntraID.Users.Analysis.SyncedUsers } else { 0 }
    $disabledUsers = if ($CollectedData.EntraID.Users.Analysis) { $CollectedData.EntraID.Users.Analysis.DisabledUsers } else { 0 }
    $totalGroups = if ($CollectedData.EntraID.Groups.Analysis) { $CollectedData.EntraID.Groups.Analysis.TotalGroups } else { 0 }
    $m365Groups = if ($CollectedData.EntraID.Groups.Analysis) { $CollectedData.EntraID.Groups.Analysis.M365Groups } else { 0 }
    $securityGroups = if ($CollectedData.EntraID.Groups.Analysis) { $CollectedData.EntraID.Groups.Analysis.SecurityGroups } else { 0 }
    $distLists = if ($CollectedData.EntraID.Groups.Analysis) { $CollectedData.EntraID.Groups.Analysis.DistributionLists } else { 0 }
    $dynamicGroups = if ($CollectedData.EntraID.Groups.Analysis) { $CollectedData.EntraID.Groups.Analysis.DynamicGroups } else { 0 }
    $totalDevices = if ($CollectedData.EntraID.Devices.Analysis) { $CollectedData.EntraID.Devices.Analysis.TotalDevices } else { 0 }
    $windowsDevices = if ($CollectedData.EntraID.Devices.Analysis) { $CollectedData.EntraID.Devices.Analysis.Windows } else { 0 }
    $iosDevices = if ($CollectedData.EntraID.Devices.Analysis) { $CollectedData.EntraID.Devices.Analysis.iOS } else { 0 }
    $androidDevices = if ($CollectedData.EntraID.Devices.Analysis) { $CollectedData.EntraID.Devices.Analysis.Android } else { 0 }
    $macDevices = if ($CollectedData.EntraID.Devices.Analysis) { $CollectedData.EntraID.Devices.Analysis.Mac } else { 0 }
    $totalMailboxes = if ($CollectedData.Exchange.Mailboxes.Analysis) { $CollectedData.Exchange.Mailboxes.Analysis.TotalMailboxes } else { 0 }
    $userMailboxes = if ($CollectedData.Exchange.Mailboxes.Analysis) { $CollectedData.Exchange.Mailboxes.Analysis.UserMailboxes } else { 0 }
    $sharedMailboxes = if ($CollectedData.Exchange.Mailboxes.Analysis) { $CollectedData.Exchange.Mailboxes.Analysis.SharedMailboxes } else { 0 }
    $roomMailboxes = if ($CollectedData.Exchange.Mailboxes.Analysis) { $CollectedData.Exchange.Mailboxes.Analysis.RoomMailboxes } else { 0 }
    $totalSPSites = if ($CollectedData.SharePoint.Sites.Analysis) { $CollectedData.SharePoint.Sites.Analysis.SharePointSites } else { 0 }
    $oneDriveSites = if ($CollectedData.SharePoint.Sites.Analysis) { $CollectedData.SharePoint.Sites.Analysis.OneDriveSites } else { 0 }
    $storageGB = if ($CollectedData.SharePoint.Sites.Analysis.TotalStorageUsedGB) { [math]::Round($CollectedData.SharePoint.Sites.Analysis.TotalStorageUsedGB, 1) } else { 0 }
    $hubSites = if ($CollectedData.SharePoint.Sites.Analysis) { $CollectedData.SharePoint.Sites.Analysis.HubSites } else { 0 }
    $teamSites = if ($CollectedData.SharePoint.Sites.Analysis) { $CollectedData.SharePoint.Sites.Analysis.TeamSites } else { 0 }
    $commSites = if ($CollectedData.SharePoint.Sites.Analysis) { $CollectedData.SharePoint.Sites.Analysis.CommunicationSites } else { 0 }
    $classicSites = if ($CollectedData.SharePoint.Sites.Analysis) { $CollectedData.SharePoint.Sites.Analysis.ClassicSites } else { 0 }
    $totalTeams = if ($CollectedData.Teams.Teams.Analysis) { $CollectedData.Teams.Teams.Analysis.TotalTeams } else { 0 }
    $totalChannels = if ($CollectedData.Teams.Teams.Analysis) { $CollectedData.Teams.Teams.Analysis.TotalChannels } else { 0 }
    $publicTeams = if ($CollectedData.Teams.Teams.Analysis) { $CollectedData.Teams.Teams.Analysis.PublicTeams } else { 0 }
    $privateTeams = if ($CollectedData.Teams.Teams.Analysis) { $CollectedData.Teams.Teams.Analysis.PrivateTeams } else { 0 }
    $customApps = if ($CollectedData.Teams.Apps.Analysis) { $CollectedData.Teams.Apps.Analysis.CustomApps } else { 0 }
    $phoneEnabled = if ($CollectedData.Teams.PhoneSystem.Analysis) { $CollectedData.Teams.PhoneSystem.Analysis.PhoneSystemEnabled } else { $false }
    $dlpPolicies = if ($CollectedData.Security.DLPPolicies.Analysis) { $CollectedData.Security.DLPPolicies.Analysis.TotalPolicies } else { 0 }
    $retentionPolicies = if ($CollectedData.Security.RetentionPolicies.Analysis) { $CollectedData.Security.RetentionPolicies.Analysis.TotalPolicies } else { 0 }
    $sensitivityLabels = if ($CollectedData.Security.SensitivityLabels.Analysis) { $CollectedData.Security.SensitivityLabels.Analysis.TotalLabels } else { 0 }
    $eDiscoveryCases = if ($CollectedData.Security.eDiscovery.Analysis) { $CollectedData.Security.eDiscovery.Analysis.TotalCases } else { 0 }
    # Dynamics 365 / Power Platform
    $d365Environments  = if ($CollectedData.Dynamics365.Environments.Analysis) { $CollectedData.Dynamics365.Environments.Analysis.TotalEnvironments } else { 0 }
    $d365ProdEnvs      = if ($CollectedData.Dynamics365.Environments.Analysis) { $CollectedData.Dynamics365.Environments.Analysis.ProductionEnvs } else { 0 }
    $d365ApiAccessible = if ($CollectedData.Dynamics365.Environments.Analysis.ContainsKey('ApiAccessible')) { $CollectedData.Dynamics365.Environments.Analysis.ApiAccessible } else { $true }
    $d365Apps          = if ($CollectedData.Dynamics365.PowerApps.Analysis) { $CollectedData.Dynamics365.PowerApps.Analysis.TotalApps } else { 0 }
    $d365CanvasApps    = if ($CollectedData.Dynamics365.PowerApps.Analysis) { $CollectedData.Dynamics365.PowerApps.Analysis.CanvasApps } else { 0 }
    $d365ModelApps     = if ($CollectedData.Dynamics365.PowerApps.Analysis) { $CollectedData.Dynamics365.PowerApps.Analysis.ModelDrivenApps } else { 0 }
    $d365Flows         = if ($CollectedData.Dynamics365.PowerAutomate.Analysis) { $CollectedData.Dynamics365.PowerAutomate.Analysis.TotalFlows } else { 0 }
    $d365ActiveFlows   = if ($CollectedData.Dynamics365.PowerAutomate.Analysis) { $CollectedData.Dynamics365.PowerAutomate.Analysis.ActiveFlows } else { 0 }
    $d365Connectors    = if ($CollectedData.Dynamics365.Connectors.Analysis) { $CollectedData.Dynamics365.Connectors.Analysis.CustomConnectors } else { 0 }
    $d365DLPPolicies   = if ($CollectedData.Dynamics365.DLPPolicies.Analysis) { $CollectedData.Dynamics365.DLPPolicies.Analysis.TotalDLPPolicies } else { 0 }
    $d365Users         = if ($CollectedData.Dynamics365.Users.Analysis) { $CollectedData.Dynamics365.Users.Analysis.TotalDynamicsUsers } else { 0 }
    $d365Solutions     = if ($CollectedData.Dynamics365.Solutions.Analysis) { $CollectedData.Dynamics365.Solutions.Analysis.TotalSolutions } else { 0 }
    $d365CustomSols    = if ($CollectedData.Dynamics365.Solutions.Analysis) { $CollectedData.Dynamics365.Solutions.Analysis.CustomSolutions } else { 0 }
    # Power BI
    $pbiWorkspaces     = if ($CollectedData.PowerBI.Workspaces.Analysis) { $CollectedData.PowerBI.Workspaces.Analysis.TotalWorkspaces } else { 0 }
    $pbiPremiumWS      = if ($CollectedData.PowerBI.Workspaces.Analysis) { $CollectedData.PowerBI.Workspaces.Analysis.PremiumWorkspaces } else { 0 }
    $pbiGateways       = if ($CollectedData.PowerBI.Gateways.Analysis) { $CollectedData.PowerBI.Gateways.Analysis.TotalGateways } else { 0 }
    $pbiOnPremGW       = if ($CollectedData.PowerBI.Gateways.Analysis) { $CollectedData.PowerBI.Gateways.Analysis.OnPremisesGateways } else { 0 }
    $pbiPersonalGW     = if ($CollectedData.PowerBI.Gateways.Analysis) { $CollectedData.PowerBI.Gateways.Analysis.PersonalGateways } else { 0 }
    $pbiCapacities     = if ($CollectedData.PowerBI.Capacities.Analysis) { $CollectedData.PowerBI.Capacities.Analysis.TotalCapacities } else { 0 }
    $pbiPremiumCap     = if ($CollectedData.PowerBI.Capacities.Analysis) { $CollectedData.PowerBI.Capacities.Analysis.PremiumCapacities } else { 0 }

    # Severity counts
    $criticalCount = if ($AnalysisResults.BySeverity.Critical) { $AnalysisResults.BySeverity.Critical.Count } else { 0 }
    $highCount = if ($AnalysisResults.BySeverity.High) { $AnalysisResults.BySeverity.High.Count } else { 0 }
    $mediumCount = if ($AnalysisResults.BySeverity.Medium) { $AnalysisResults.BySeverity.Medium.Count } else { 0 }
    $lowCount = if ($AnalysisResults.BySeverity.Low) { $AnalysisResults.BySeverity.Low.Count } else { 0 }
    $totalIssues = $criticalCount + $highCount + $mediumCount + $lowCount

    # Issues by category
    $categoryData = @{}
    foreach ($sev in @("Critical", "High", "Medium", "Low")) {
        $issues = $AnalysisResults.BySeverity[$sev]
        if ($issues) { foreach ($i in $issues) { if (-not $categoryData.ContainsKey($i.Category)) { $categoryData[$i.Category] = 0 }; $categoryData[$i.Category]++ } }
    }
    $catLabels = ($categoryData.Keys | Sort-Object | ForEach-Object { "'$_'" }) -join ","
    $catValues = ($categoryData.Keys | Sort-Object | ForEach-Object { $categoryData[$_] }) -join ","
    $catColors = @("'rgba(59,130,246,0.8)'","'rgba(239,68,68,0.8)'","'rgba(249,115,22,0.8)'","'rgba(34,197,94,0.8)'","'rgba(139,92,246,0.8)'","'rgba(236,72,153,0.8)'","'rgba(14,165,233,0.8)'","'rgba(245,158,11,0.8)'")
    $catColorsJS = ($catColors[0..([math]::Min($categoryData.Count - 1, 7))]) -join ","

    # Complexity breakdown for radar
    $radarLabels = @(); $radarData = @()
    if ($ComplexityScore.Breakdown) {
        foreach ($key in ($ComplexityScore.Breakdown.Keys | Sort-Object)) {
            $radarLabels += "'$key'"
            $bv = $ComplexityScore.Breakdown[$key]
            $radarData += if ($bv -is [hashtable] -and $bv.WeightedScore) { [math]::Round($bv.WeightedScore, 1) } elseif ($bv -is [hashtable] -and $bv.Score) { [math]::Round($bv.Score, 1) } else { try { [math]::Round([double]$bv, 1) } catch { 0 } }
        }
    }
    $radarLabelsJS = $radarLabels -join ","
    $radarDataJS = $radarData -join ","

    # Readiness
    $readinessPercent = [math]::Max(0, 100 - $ComplexityScore.TotalScore)
    $readinessColor = if ($readinessPercent -ge 80) { "#22c55e" } elseif ($readinessPercent -ge 60) { "#eab308" } elseif ($readinessPercent -ge 40) { "#f97316" } else { "#ef4444" }
    $readinessStatus = if ($criticalCount -eq 0 -and $highCount -le 2) { "Ready to proceed" } elseif ($criticalCount -eq 0) { "Address high priority items" } else { "Critical blockers must be resolved" }

    # === Build HTML ===
    $html = Get-InteractiveHTMLHeader -Title "IT Technical Report - $tenantName" -ReportType "IT"

    # Sidebar
    $navItems = @(
        @{ Id = 'page-dashboard'; Label = 'Dashboard'; Icon = $svgDash }
        @{ Id = 'page-identity'; Label = 'Identity'; Icon = $svgIdentity }
        @{ Id = 'page-exchange'; Label = 'Exchange'; Icon = $svgExchange }
        @{ Id = 'page-sharepoint'; Label = 'SharePoint'; Icon = $svgSP }
        @{ Id = 'page-teams'; Label = 'Teams'; Icon = $svgTeams }
        @{ Id = 'page-security'; Label = 'Security'; Icon = $svgSecurity }
        @{ Id = 'page-dynamics'; Label = 'Dynamics 365'; Icon = $svgDynamics }
        @{ Id = 'page-powerbi'; Label = 'Power BI'; Icon = $svgPowerBI }
        @{ Id = 'page-issues'; Label = 'Issues'; Icon = $svgIssues }
        @{ Id = 'page-ai-analysis'; Label = 'Deep Analysis'; Icon = $svgAI }
    )
    $html += Get-SidebarHTML -TenantName $tenantName -ReportType "IT" -NavItems $navItems
    $html += '<div class="main-content">'
    $html += Get-TopBarHTML
    $html += '<main class="page-content">'

    # ==================== PAGE: DASHBOARD ====================
    $html += '<div class="page active" id="page-dashboard">'
    $html += '<div class="page-header"><h1>Migration Dashboard</h1><p>Tenant: ' + $tenantName + ' | ID: ' + $tenantId + ' | ' + (Get-Date -Format "MMMM dd, yyyy") + '</p></div>'

    # KPI row
    $html += '<div class="kpi-grid">'
    $html += Get-KPICardHTML -Value $licensedUsers -Label "Licensed Users" -IconSvg $svgUsers -Color "#3b82f6" -Subtitle "$totalUsers total"
    $html += Get-KPICardHTML -Value $totalMailboxes -Label "Mailboxes" -IconSvg $svgMail -Color "#0891b2" -Subtitle "Exchange Online"
    $html += Get-KPICardHTML -Value $totalSPSites -Label "SharePoint Sites" -IconSvg $svgFolder -Color "#f59e0b" -Subtitle "$($storageGB) GB storage"
    $html += Get-KPICardHTML -Value $oneDriveSites -Label "OneDrive" -IconSvg $svgCloud -Color "#06b6d4"
    $html += Get-KPICardHTML -Value $totalTeams -Label "Teams" -IconSvg $svgChat -Color "#10b981" -Subtitle "$totalChannels channels"
    $html += Get-KPICardHTML -Value $totalGroups -Label "Groups" -IconSvg $svgGroup -Color "#8b5cf6"
    $html += Get-KPICardHTML -Value $d365Environments -Label "D365 Environments" -IconSvg $svgDynamics -Color "#7c3aed" -Subtitle "$d365Apps apps / $d365Flows flows"
    $html += Get-KPICardHTML -Value $pbiWorkspaces -Label "Power BI Workspaces" -IconSvg $svgPowerBI -Color "#f59e0b" -Subtitle "$pbiGateways gateway(s)"
    $html += '</div>'

    # Readiness gauge + severity bars row
    $html += @"
    <div style="display:grid;grid-template-columns:280px 1fr;gap:20px;margin-bottom:24px;">
        <div class="chart-card" style="text-align:center;">
            <h3>Migration Readiness</h3>
            <div style="position:relative;width:180px;height:180px;margin:16px auto;">
                <svg viewBox="0 0 180 180" style="transform:rotate(-90deg);">
                    <circle cx="90" cy="90" r="80" fill="none" stroke="var(--bg-elevated)" stroke-width="12"/>
                    <circle cx="90" cy="90" r="80" fill="none" stroke="$readinessColor" stroke-width="12" stroke-linecap="round" stroke-dasharray="502.65" stroke-dashoffset="$([math]::Round(502.65 * (1 - $readinessPercent / 100), 2))"/>
                </svg>
                <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center;">
                    <div class="gauge-value" style="color:$readinessColor;">$readinessPercent%</div>
                    <div class="gauge-label">Ready</div>
                </div>
            </div>
            <div class="gauge-status" style="color:$readinessColor;">$readinessStatus</div>
        </div>
        <div class="chart-card">
            <h3>Issue Summary</h3>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:12px;">
                <div>
                    <div style="display:flex;justify-content:space-between;margin-bottom:6px;"><span style="color:#ef4444;font-weight:600;">Critical</span><span style="font-weight:700;color:#ef4444;">$criticalCount</span></div>
                    <div class="progress-bar-track"><div class="progress-bar-fill" style="width:$(if($totalIssues -gt 0){[math]::Round($criticalCount/$totalIssues*100)}else{0})%;background:#ef4444;"></div></div>
                </div>
                <div>
                    <div style="display:flex;justify-content:space-between;margin-bottom:6px;"><span style="color:#f97316;font-weight:600;">High</span><span style="font-weight:700;color:#f97316;">$highCount</span></div>
                    <div class="progress-bar-track"><div class="progress-bar-fill" style="width:$(if($totalIssues -gt 0){[math]::Round($highCount/$totalIssues*100)}else{0})%;background:#f97316;"></div></div>
                </div>
                <div>
                    <div style="display:flex;justify-content:space-between;margin-bottom:6px;"><span style="color:#eab308;font-weight:600;">Medium</span><span style="font-weight:700;color:#eab308;">$mediumCount</span></div>
                    <div class="progress-bar-track"><div class="progress-bar-fill" style="width:$(if($totalIssues -gt 0){[math]::Round($mediumCount/$totalIssues*100)}else{0})%;background:#eab308;"></div></div>
                </div>
                <div>
                    <div style="display:flex;justify-content:space-between;margin-bottom:6px;"><span style="color:#22c55e;font-weight:600;">Low</span><span style="font-weight:700;color:#22c55e;">$lowCount</span></div>
                    <div class="progress-bar-track"><div class="progress-bar-fill" style="width:$(if($totalIssues -gt 0){[math]::Round($lowCount/$totalIssues*100)}else{0})%;background:#22c55e;"></div></div>
                </div>
            </div>
            <div style="margin-top:16px;padding-top:16px;border-top:1px solid var(--border-default);display:flex;justify-content:space-between;align-items:center;">
                <span style="color:var(--text-muted);">Total: <strong style="color:var(--text-primary);">$totalIssues</strong> issues</span>
                <span class="severity-badge severity-$(if($criticalCount -gt 0){'critical'}elseif($highCount -gt 0){'high'}else{'low'})">$($ComplexityScore.ComplexityLevel)</span>
            </div>
        </div>
    </div>
"@

    # Charts row: severity doughnut + category horizontal bar + complexity radar
    $html += '<div class="chart-grid" style="grid-template-columns:repeat(3,1fr);">'
    $html += Get-ChartCardHTML -Title "Severity Distribution" -CanvasId "chartSeverity" -Height "280px"
    $html += Get-ChartCardHTML -Title "Issues by Category" -CanvasId "chartCategory" -Height "280px"
    $html += Get-ChartCardHTML -Title "Complexity Breakdown" -CanvasId "chartComplexity" -Height "280px"
    $html += '</div>'
    $html += @"
<script>
createDoughnutChart('chartSeverity',['Critical','High','Medium','Low'],[$criticalCount,$highCount,$mediumCount,$lowCount],['#ef4444','#f97316','#eab308','#22c55e']);
createHorizontalBarChart('chartCategory',[$catLabels],[$catValues],[$catColorsJS]);
createRadarChart('chartComplexity',[$radarLabelsJS],[$radarDataJS],'Risk Score','rgb(59,130,246)');
</script>
"@
    $html += '</div>' # end page-dashboard

    # ==================== PAGE: IDENTITY ====================
    $html += '<div class="page" id="page-identity">'
    $html += '<div class="page-header"><h1>Identity &amp; Access</h1><p>Entra ID users, groups, devices, and conditional access</p></div>'
    $html += '<div class="kpi-grid">'
    $html += Get-KPICardHTML -Value $totalUsers -Label "Total Users" -IconSvg $svgUsers -Color "#3b82f6"
    $html += Get-KPICardHTML -Value $licensedUsers -Label "Licensed Users" -IconSvg $svgUsers -Color "#10b981"
    $html += Get-KPICardHTML -Value $syncedUsers -Label "Synced Users" -IconSvg $svgUsers -Color "#f59e0b" -Subtitle "From on-premises AD"
    $html += Get-KPICardHTML -Value $guestUsers -Label "Guest Users" -IconSvg $svgUsers -Color "#8b5cf6"
    $html += '</div>'
    $html += '<div class="chart-grid">'
    $html += Get-ChartCardHTML -Title "User Composition" -CanvasId "chartUserComp"
    $html += Get-ChartCardHTML -Title "Device Distribution" -CanvasId "chartDevices"
    $html += Get-ChartCardHTML -Title "Group Types" -CanvasId "chartGroups"
    $html += '</div>'
    $html += @"
<script>
createDoughnutChart('chartUserComp',['Licensed','Guest','Disabled'],[$licensedUsers,$guestUsers,$disabledUsers],['#3b82f6','#8b5cf6','#6b7280']);
createDoughnutChart('chartDevices',['Windows','iOS','Android','macOS'],[$windowsDevices,$iosDevices,$androidDevices,$macDevices],['#3b82f6','#f97316','#22c55e','#8b5cf6']);
createDoughnutChart('chartGroups',['M365','Security','Distribution','Dynamic'],[$m365Groups,$securityGroups,$distLists,$dynamicGroups],['#3b82f6','#ef4444','#f97316','#8b5cf6']);
</script>
"@
    # Identity tables
    $caEnabled = if ($CollectedData.EntraID.ConditionalAccess.Analysis) { $CollectedData.EntraID.ConditionalAccess.Analysis.EnabledPolicies } else { 0 }
    $caMFA = if ($CollectedData.EntraID.ConditionalAccess.Analysis) { $CollectedData.EntraID.ConditionalAccess.Analysis.MFAPolicies } else { 0 }
    $html += @"
    <div class="data-section" data-searchable>
        <h3>Identity Summary</h3>
        <table><thead><tr><th>Metric</th><th>Value</th></tr></thead><tbody>
        <tr><td>Total Users</td><td>$totalUsers</td></tr>
        <tr><td>Licensed Users</td><td>$licensedUsers</td></tr>
        <tr><td>Guest Users</td><td>$guestUsers</td></tr>
        <tr><td>Synced from AD</td><td>$syncedUsers</td></tr>
        <tr><td>Disabled Users</td><td>$disabledUsers</td></tr>
        <tr><td>Total Groups</td><td>$totalGroups</td></tr>
        <tr><td>Total Devices</td><td>$totalDevices</td></tr>
        <tr><td>Conditional Access Policies (Enabled)</td><td>$caEnabled</td></tr>
        <tr><td>MFA Policies</td><td>$caMFA</td></tr>
        </tbody></table>
    </div>
"@
    $html += '</div>' # end page-identity

    # ==================== PAGE: EXCHANGE ====================
    $html += '<div class="page" id="page-exchange">'
    $html += '<div class="page-header"><h1>Exchange Online</h1><p>Mailboxes, distribution lists, and mail flow</p></div>'
    $html += '<div class="kpi-grid">'
    $html += Get-KPICardHTML -Value $totalMailboxes -Label "Total Mailboxes" -IconSvg $svgMail -Color "#0891b2"
    $html += Get-KPICardHTML -Value $userMailboxes -Label "User Mailboxes" -IconSvg $svgMail -Color "#3b82f6"
    $html += Get-KPICardHTML -Value $sharedMailboxes -Label "Shared Mailboxes" -IconSvg $svgMail -Color "#f59e0b"
    $html += Get-KPICardHTML -Value $roomMailboxes -Label "Room Mailboxes" -IconSvg $svgMail -Color "#10b981"
    $html += '</div>'
    $html += '<div class="chart-grid" style="grid-template-columns:1fr;">'
    $html += Get-ChartCardHTML -Title "Mailbox Types" -CanvasId "chartMailboxTypes"
    $html += '</div>'
    $equipmentMbx = if ($CollectedData.Exchange.Mailboxes.Analysis) { $CollectedData.Exchange.Mailboxes.Analysis.EquipmentMailboxes } else { 0 }
    $archiveEnabled = if ($CollectedData.Exchange.Mailboxes.Analysis) { $CollectedData.Exchange.Mailboxes.Analysis.ArchiveEnabled } else { 0 }
    $litigationHold = if ($CollectedData.Exchange.Mailboxes.Analysis) { $CollectedData.Exchange.Mailboxes.Analysis.LitigationHold } else { 0 }
    $totalDLs = if ($CollectedData.Exchange.DistributionLists.Analysis) { $CollectedData.Exchange.DistributionLists.Analysis.TotalDistributionLists } else { 0 }
    $html += @"
<script>createDoughnutChart('chartMailboxTypes',['User','Shared','Room','Equipment'],[$userMailboxes,$sharedMailboxes,$roomMailboxes,$equipmentMbx],['#3b82f6','#22c55e','#f97316','#8b5cf6']);</script>
    <div class="data-section" data-searchable>
        <h3>Exchange Details</h3>
        <table><thead><tr><th>Metric</th><th>Value</th></tr></thead><tbody>
        <tr><td>Total Mailboxes</td><td>$totalMailboxes</td></tr>
        <tr><td>Archive Enabled</td><td>$archiveEnabled</td></tr>
        <tr><td>Litigation Hold</td><td>$litigationHold</td></tr>
        <tr><td>Distribution Lists</td><td>$totalDLs</td></tr>
        <tr><td>Forwarding Configured</td><td>$(if ($CollectedData.Exchange.Mailboxes.Analysis) { $CollectedData.Exchange.Mailboxes.Analysis.ForwardingConfigured } else { 0 })</td></tr>
        </tbody></table>
    </div>
"@
    $html += '</div>' # end page-exchange

    # ==================== PAGE: SHAREPOINT ====================
    $html += '<div class="page" id="page-sharepoint">'
    $html += '<div class="page-header"><h1>SharePoint &amp; OneDrive</h1><p>Sites, storage, and content</p></div>'
    $html += '<div class="kpi-grid">'
    $html += Get-KPICardHTML -Value $totalSPSites -Label "SharePoint Sites" -IconSvg $svgFolder -Color "#f59e0b"
    $html += Get-KPICardHTML -Value $oneDriveSites -Label "OneDrive Sites" -IconSvg $svgCloud -Color "#06b6d4"
    $html += Get-KPICardHTML -Value "$($storageGB) GB" -Label "Storage Used" -IconSvg $svgDb -Color "#10b981"
    $html += Get-KPICardHTML -Value $hubSites -Label "Hub Sites" -IconSvg $svgFolder -Color "#8b5cf6"
    $html += '</div>'
    $html += '<div class="chart-grid" style="grid-template-columns:1fr;">'
    $html += Get-ChartCardHTML -Title "SharePoint Site Types" -CanvasId "chartSPTypes"
    $html += '</div>'
    $groupConnected = if ($CollectedData.SharePoint.Sites.Analysis) { $CollectedData.SharePoint.Sites.Analysis.GroupConnectedSites } else { 0 }
    $html += @"
<script>createDoughnutChart('chartSPTypes',['Team Sites','Communication','Classic','Group-Connected'],[$teamSites,$commSites,$classicSites,$groupConnected],['#3b82f6','#22c55e','#f97316','#8b5cf6']);</script>
    <div class="data-section" data-searchable>
        <h3>SharePoint Details</h3>
        <table><thead><tr><th>Metric</th><th>Value</th></tr></thead><tbody>
        <tr><td>SharePoint Sites</td><td>$totalSPSites</td></tr>
        <tr><td>OneDrive Sites</td><td>$oneDriveSites</td></tr>
        <tr><td>Team Sites</td><td>$teamSites</td></tr>
        <tr><td>Communication Sites</td><td>$commSites</td></tr>
        <tr><td>Classic Sites</td><td>$classicSites</td></tr>
        <tr><td>Hub Sites</td><td>$hubSites</td></tr>
        <tr><td>Group-Connected</td><td>$groupConnected</td></tr>
        <tr><td>Total Storage Used</td><td>$storageGB GB</td></tr>
        </tbody></table>
    </div>
"@
    $html += '</div>' # end page-sharepoint

    # ==================== PAGE: TEAMS ====================
    $html += '<div class="page" id="page-teams">'
    $html += '<div class="page-header"><h1>Microsoft Teams</h1><p>Teams, channels, apps, and phone system</p></div>'
    $html += '<div class="kpi-grid">'
    $html += Get-KPICardHTML -Value $totalTeams -Label "Teams" -IconSvg $svgChat -Color "#10b981"
    $html += Get-KPICardHTML -Value $totalChannels -Label "Channels" -IconSvg $svgChat -Color "#3b82f6"
    $html += Get-KPICardHTML -Value $customApps -Label "Custom Apps" -IconSvg $svgApp -Color "#f59e0b"
    $phoneVal = if ($phoneEnabled) { "Enabled" } else { "Disabled" }
    $html += Get-KPICardHTML -Value $phoneVal -Label "Phone System" -IconSvg $svgPhone -Color "#8b5cf6"
    $html += '</div>'
    $html += '<div class="chart-grid" style="grid-template-columns:1fr;">'
    $html += Get-ChartCardHTML -Title "Team Visibility" -CanvasId "chartTeamVis"
    $html += '</div>'
    $archivedTeams = if ($CollectedData.Teams.Teams.Analysis) { $CollectedData.Teams.Teams.Analysis.ArchivedTeams } else { 0 }
    $guestTeams = if ($CollectedData.Teams.Teams.Analysis) { $CollectedData.Teams.Teams.Analysis.TeamsWithGuests } else { 0 }
    $privateChannels = if ($CollectedData.Teams.Teams.Analysis) { $CollectedData.Teams.Teams.Analysis.PrivateChannels } else { 0 }
    $sharedChannels = if ($CollectedData.Teams.Teams.Analysis) { $CollectedData.Teams.Teams.Analysis.SharedChannels } else { 0 }
    $html += @"
<script>createDoughnutChart('chartTeamVis',['Public','Private'],[$publicTeams,$privateTeams],['#10b981','#3b82f6']);</script>
    <div class="data-section" data-searchable>
        <h3>Teams Details</h3>
        <table><thead><tr><th>Metric</th><th>Value</th></tr></thead><tbody>
        <tr><td>Total Teams</td><td>$totalTeams</td></tr>
        <tr><td>Public Teams</td><td>$publicTeams</td></tr>
        <tr><td>Private Teams</td><td>$privateTeams</td></tr>
        <tr><td>Archived Teams</td><td>$archivedTeams</td></tr>
        <tr><td>Teams with Guests</td><td>$guestTeams</td></tr>
        <tr><td>Total Channels</td><td>$totalChannels</td></tr>
        <tr><td>Private Channels</td><td>$privateChannels</td></tr>
        <tr><td>Shared Channels</td><td>$sharedChannels</td></tr>
        <tr><td>Custom Apps</td><td>$customApps</td></tr>
        </tbody></table>
    </div>
"@
    $html += '</div>' # end page-teams

    # ==================== PAGE: SECURITY ====================
    $html += '<div class="page" id="page-security">'
    $html += '<div class="page-header"><h1>Security &amp; Compliance</h1><p>DLP, retention, sensitivity labels, and eDiscovery</p></div>'
    $html += '<div class="kpi-grid">'
    $html += Get-KPICardHTML -Value $dlpPolicies -Label "DLP Policies" -IconSvg $svgShield -Color "#ef4444"
    $html += Get-KPICardHTML -Value $retentionPolicies -Label "Retention Policies" -IconSvg $svgLock -Color "#3b82f6"
    $html += Get-KPICardHTML -Value $sensitivityLabels -Label "Sensitivity Labels" -IconSvg $svgShield -Color "#f59e0b"
    $html += Get-KPICardHTML -Value $eDiscoveryCases -Label "eDiscovery Cases" -IconSvg $svgShield -Color "#8b5cf6"
    $html += '</div>'
    $dlpEnabled = if ($CollectedData.Security.DLPPolicies.Analysis) { $CollectedData.Security.DLPPolicies.Analysis.EnabledPolicies } else { 0 }
    $dlpEnforced = if ($CollectedData.Security.DLPPolicies.Analysis) { $CollectedData.Security.DLPPolicies.Analysis.EnforcedPolicies } else { 0 }
    $retEnabled = if ($CollectedData.Security.RetentionPolicies.Analysis) { $CollectedData.Security.RetentionPolicies.Analysis.EnabledPolicies } else { 0 }
    $activeLabels = if ($CollectedData.Security.SensitivityLabels.Analysis) { $CollectedData.Security.SensitivityLabels.Analysis.ActiveLabels } else { 0 }
    $activeCases = if ($CollectedData.Security.eDiscovery.Analysis) { $CollectedData.Security.eDiscovery.Analysis.ActiveCases } else { 0 }
    $auditEnabled = if ($CollectedData.Security.AuditConfig.Analysis) { $CollectedData.Security.AuditConfig.Analysis.AuditEnabled } else { "Unknown" }
    $alertPolicies = if ($CollectedData.Security.AlertPolicies.Analysis) { $CollectedData.Security.AlertPolicies.Analysis.TotalPolicies } else { 0 }
    $html += @"
    <div class="data-section" data-searchable>
        <h3>Security Overview</h3>
        <table><thead><tr><th>Area</th><th>Total</th><th>Active/Enabled</th><th>Details</th></tr></thead><tbody>
        <tr><td>DLP Policies</td><td>$dlpPolicies</td><td>$dlpEnabled</td><td>$dlpEnforced enforced</td></tr>
        <tr><td>Retention Policies</td><td>$retentionPolicies</td><td>$retEnabled</td><td>-</td></tr>
        <tr><td>Sensitivity Labels</td><td>$sensitivityLabels</td><td>$activeLabels</td><td>-</td></tr>
        <tr><td>eDiscovery Cases</td><td>$eDiscoveryCases</td><td>$activeCases active</td><td>-</td></tr>
        <tr><td>Audit Logging</td><td>-</td><td>$auditEnabled</td><td>-</td></tr>
        <tr><td>Alert Policies</td><td>$alertPolicies</td><td>-</td><td>-</td></tr>
        </tbody></table>
    </div>
"@
    $html += '</div>' # end page-security

    # ==================== PAGE: DYNAMICS 365 ====================
    $html += '<div class="page" id="page-dynamics">'
    $html += '<div class="page-header"><h1>Dynamics 365 &amp; Power Platform</h1><p>Environments, Power Apps, Power Automate, and governance</p></div>'
    if (-not $d365ApiAccessible) {
        $html += '<div style="background:rgba(245,158,11,0.12);border:1px solid rgba(245,158,11,0.45);border-radius:10px;padding:14px 18px;margin-bottom:18px;color:#fbbf24;font-size:0.9rem;line-height:1.5;">'
        $html += '<strong style="display:block;margin-bottom:4px;">Power Platform Admin API Not Accessible</strong>'
        $html += 'Environment counts could not be retrieved. The discovery account may lack the <strong>Power Platform Administrator</strong> role, or a Power Platform-scoped token could not be obtained. '
        $html += 'Install the <code style="background:rgba(0,0,0,0.3);padding:1px 5px;border-radius:4px;">Az.Accounts</code> module and ensure the account has Power Platform Admin permissions, then re-run discovery for accurate environment data.'
        $html += '</div>'
    }
    $html += '<div class="kpi-grid">'
    $html += Get-KPICardHTML -Value $d365Environments -Label "Environments" -IconSvg $svgDynamics -Color "#7c3aed" -Subtitle "$d365ProdEnvs production"
    $html += Get-KPICardHTML -Value $d365Apps -Label "Power Apps" -IconSvg $svgApp -Color "#2563eb" -Subtitle "$d365CanvasApps canvas / $d365ModelApps model-driven"
    $html += Get-KPICardHTML -Value $d365Flows -Label "Power Automate Flows" -IconSvg $svgCloud -Color "#0891b2" -Subtitle "$d365ActiveFlows active"
    $html += Get-KPICardHTML -Value $d365Users -Label "D365 Licensed Users" -IconSvg $svgIdentity -Color "#dc2626"
    $html += '</div>'
    $html += '<div class="chart-grid" style="grid-template-columns:1fr 1fr;">'
    $html += Get-ChartCardHTML -Title "App Types" -CanvasId "chartD365Apps"
    $html += Get-ChartCardHTML -Title "Flow Status" -CanvasId "chartD365Flows"
    $html += '</div>'
    $d365StoppedFlows = if ($CollectedData.Dynamics365.PowerAutomate.Analysis) { $CollectedData.Dynamics365.PowerAutomate.Analysis.StoppedFlows } else { 0 }
    $html += @"
<script>
createDoughnutChart('chartD365Apps',['Canvas Apps','Model-Driven'],[$d365CanvasApps,$d365ModelApps],['#2563eb','#7c3aed']);
createDoughnutChart('chartD365Flows',['Active','Stopped'],[$d365ActiveFlows,$d365StoppedFlows],['#10b981','#ef4444']);
</script>
    <div class="data-section" data-searchable>
        <h3>Power Platform Inventory</h3>
        <table><thead><tr><th>Area</th><th>Total</th><th>Details</th></tr></thead><tbody>
        <tr><td>Environments</td><td>$(if ($d365ApiAccessible) { $d365Environments } else { '<span style="color:#fbbf24;">N/A</span>' })</td><td>$(if ($d365ApiAccessible) { "$d365ProdEnvs production, $(if ($CollectedData.Dynamics365.Environments.Analysis) { $CollectedData.Dynamics365.Environments.Analysis.SandboxEnvs } else { 0 }) sandbox" } else { '<span style="color:#fbbf24;">Power Platform Admin API not accessible — re-run with Power Platform Admin role</span>' })</td></tr>
        <tr><td>Power Apps</td><td>$d365Apps</td><td>$d365CanvasApps canvas, $d365ModelApps model-driven</td></tr>
        <tr><td>Power Automate Flows</td><td>$d365Flows</td><td>$d365ActiveFlows active, $d365StoppedFlows stopped</td></tr>
        <tr><td>Custom Connectors</td><td>$d365Connectors</td><td>Must be recreated in target tenant</td></tr>
        <tr><td>DLP Policies</td><td>$d365DLPPolicies</td><td>$(if ($d365DLPPolicies -eq 0 -and $d365Environments -gt 0) { '<span style="color:#ef4444;font-weight:600;">None configured — governance gap</span>' } else { 'Must be recreated in target' })</td></tr>
        <tr><td>Dataverse Solutions</td><td>$d365Solutions</td><td>$d365CustomSols custom (unmanaged) solutions</td></tr>
        <tr><td>D365 Licensed Users</td><td>$d365Users</td><td>Security roles and business units must be recreated</td></tr>
        </tbody></table>
    </div>
"@
    $html += '</div>' # end page-dynamics

    # ==================== PAGE: POWER BI ====================
    $html += '<div class="page" id="page-powerbi">'
    $html += '<div class="page-header"><h1>Power BI</h1><p>Workspaces, gateways, and capacity</p></div>'
    $html += '<div class="kpi-grid">'
    $html += Get-KPICardHTML -Value $pbiWorkspaces -Label "Workspaces" -IconSvg $svgPowerBI -Color "#f59e0b" -Subtitle "$pbiPremiumWS on Premium capacity"
    $html += Get-KPICardHTML -Value $pbiGateways -Label "Gateways" -IconSvg $svgCloud -Color "#ef4444" -Subtitle "$pbiOnPremGW on-premises"
    $html += Get-KPICardHTML -Value $pbiCapacities -Label "Capacities" -IconSvg $svgShield -Color "#7c3aed" -Subtitle "$pbiPremiumCap Premium"
    $pbiSharedWS = $pbiWorkspaces - $pbiPremiumWS
    $html += Get-KPICardHTML -Value $pbiSharedWS -Label "Shared Capacity WS" -IconSvg $svgGroup -Color "#3b82f6"
    $html += '</div>'
    $html += '<div class="chart-grid" style="grid-template-columns:1fr 1fr;">'
    $html += Get-ChartCardHTML -Title "Workspace Capacity" -CanvasId "chartPBIWorkspaces"
    $html += Get-ChartCardHTML -Title "Gateway Types" -CanvasId "chartPBIGateways"
    $html += '</div>'
    $html += @"
<script>
createDoughnutChart('chartPBIWorkspaces',['Shared Capacity','Premium Capacity'],[$pbiSharedWS,$pbiPremiumWS],['#3b82f6','#f59e0b']);
createDoughnutChart('chartPBIGateways',['On-Premises','Personal'],[$pbiOnPremGW,$pbiPersonalGW],['#ef4444','#f97316']);
</script>
    <div class="data-section" data-searchable>
        <h3>Power BI Inventory</h3>
        <table><thead><tr><th>Area</th><th>Total</th><th>Migration Impact</th></tr></thead><tbody>
        <tr><td>Workspaces</td><td>$pbiWorkspaces</td><td>All reports and datasets must be republished to target tenant</td></tr>
        <tr><td>Premium Workspaces</td><td>$pbiPremiumWS</td><td>$(if ($pbiPremiumWS -gt 0) { 'Premium capacity required in target tenant' } else { 'No premium capacity required' })</td></tr>
        <tr><td>On-Premises Gateways</td><td>$pbiOnPremGW</td><td>$(if ($pbiOnPremGW -gt 0) { '<span style="color:#ef4444;font-weight:600;">Must be reinstalled and reconfigured for target tenant</span>' } else { 'None' })</td></tr>
        <tr><td>Personal Gateways</td><td>$pbiPersonalGW</td><td>$(if ($pbiPersonalGW -gt 0) { 'Users must reinstall personal gateways after migration' } else { 'None' })</td></tr>
        <tr><td>Premium Capacities</td><td>$pbiPremiumCap</td><td>$(if ($pbiPremiumCap -gt 0) { 'Provision equivalent capacity in target tenant before migration' } else { 'None configured' })</td></tr>
        </tbody></table>
    </div>
"@
    $html += '</div>' # end page-powerbi

    # ==================== PAGE: ISSUES (Checklist) ====================
    $html += '<div class="page" id="page-issues">'
    $html += '<div class="page-header"><h1>Migration Issues Checklist</h1><p>' + $totalIssues + ' issues found across ' + $categoryData.Count + ' categories</p></div>'

    $severityMeta = @{
        "Critical" = @{ Color = "#ef4444"; Icon = "🔴"; Desc = "Must resolve before migration" }
        "High"     = @{ Color = "#f97316"; Icon = "🟠"; Desc = "Address during preparation phase" }
        "Medium"   = @{ Color = "#eab308"; Icon = "🟡"; Desc = "Plan for these - not blockers" }
        "Low"      = @{ Color = "#22c55e"; Icon = "🟢"; Desc = "Can be addressed post-migration" }
    }
    $issueIndex = 0

    foreach ($severity in @("Critical", "High", "Medium", "Low")) {
        $gotchas = $AnalysisResults.BySeverity[$severity]
        if ($gotchas -and $gotchas.Count -gt 0) {
            $sLower = $severity.ToLower()
            $sColor = $severityMeta[$severity].Color
            $sIcon = $severityMeta[$severity].Icon
            $sDesc = $severityMeta[$severity].Desc

            $html += @"
    <div class="severity-group" data-severity="$sLower" data-searchable>
        <div class="severity-group-header" style="background:linear-gradient(135deg,rgba($($sColor -replace '#',''),0.12),transparent);border-left:4px solid $sColor;">
            <div>
                <strong style="color:$sColor;font-size:1.1rem;">$sIcon $severity ($($gotchas.Count))</strong>
                <div style="font-size:0.85rem;color:var(--text-muted);margin-top:4px;">$sDesc</div>
            </div>
            <div style="text-align:right;">
                <span id="progress-$sLower" style="font-size:1.5rem;font-weight:700;color:$sColor;">0</span>
                <span style="color:var(--text-muted);">/ $($gotchas.Count)</span>
                <div class="progress-bar-track" style="width:120px;margin-top:6px;">
                    <div id="progressbar-$sLower" class="progress-bar-fill" style="width:0%;background:$sColor;"></div>
                </div>
            </div>
        </div>
"@
            # Group by category
            $byCategory = @{}
            foreach ($g in $gotchas) {
                if (-not $byCategory.ContainsKey($g.Category)) { $byCategory[$g.Category] = @() }
                $byCategory[$g.Category] += $g
            }
            foreach ($cat in ($byCategory.Keys | Sort-Object)) {
                $html += "        <h4 style='font-size:0.85rem;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.05em;margin:16px 0 8px;padding-bottom:6px;border-bottom:1px solid var(--border-default);'>$cat ($($byCategory[$cat].Count))</h4>`n"
                foreach ($gotcha in $byCategory[$cat]) {
                    $issueIndex++
                    $html += @"
        <div class="issue-card $sLower" id="card-$issueIndex" data-severity="$sLower">
            <div style="display:flex;gap:12px;align-items:flex-start;">
                <label class="checkbox-wrapper">
                    <input type="checkbox" class="issue-checkbox" data-severity="$sLower" onchange="updateProgress('$sLower',$($gotchas.Count));toggleCardComplete(this,'card-$issueIndex')">
                    <span class="checkmark" style="border-color:$sColor;">
                        <svg fill="none" stroke="white" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="3" d="M5 13l4 4L19 7"></path></svg>
                    </span>
                </label>
                <div style="flex:1;min-width:0;cursor:pointer;" onclick="toggleIssueDetail('detail-$issueIndex')">
                    <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px;">
                        <h5 style="margin:0;font-size:0.95rem;font-weight:600;color:var(--text-primary);">$($gotcha.Name)</h5>
                        <div style="display:flex;align-items:center;gap:8px;flex-shrink:0;">
                            <span class="severity-badge severity-$sLower">$($gotcha.AffectedCount) affected</span>
                            <span class="expand-arrow" id="arrow-detail-$issueIndex" style="color:var(--text-muted);font-size:0.75rem;transition:transform 0.2s;">&#9654;</span>
                        </div>
                    </div>
                    <div id="detail-$issueIndex" style="display:none;margin-top:10px;">
                        <p style="margin:0 0 10px;color:var(--text-secondary);font-size:0.88rem;line-height:1.5;">$($gotcha.Description)</p>
                        <div style="background:var(--bg-elevated);border-radius:8px;padding:10px;">
                            <div style="font-size:0.72rem;color:var(--primary);font-weight:600;text-transform:uppercase;margin-bottom:4px;">Action Required</div>
                            <p style="margin:0;color:var(--text-primary);font-size:0.85rem;">$($gotcha.Recommendation)</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
"@
                }
            }
            $html += "    </div>`n" # end severity-group
        }
    }
    $html += '</div>' # end page-issues

    # ==================== PAGE: AI ANALYSIS ====================
    $html += '<div class="page" id="page-ai-analysis">'
    $html += '<div class="page-header"><h1>Deep Analysis</h1><p>Comprehensive migration analysis and recommendations</p></div>'

    if ($AIAnalysis -and $AIAnalysis.Success -and $AIAnalysis.Analysis) {
        $aiContent = Convert-AIMarkdownToHTML -Markdown $AIAnalysis.Analysis

        # Use string concatenation instead of here-string to avoid PowerShell interpolation
        # of $ characters in AI-generated content (e.g., $UserPrincipalName, $true)
        $html += '    <div class="ai-card"><div>' + $aiContent + '</div></div>'
    } else {
        $html += '<div class="ai-card"><p style="color:var(--text-muted);text-align:center;padding:40px;">Deep analysis was not performed. Re-run with analysis enabled for enhanced insights.</p></div>'
    }
    $html += '</div>' # end page-ai-analysis

    # Close layout
    $html += '</main></div>'
    $html += Get-InteractiveHTMLFooter

    $html | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Host "  Interactive IT Report generated: $OutputPath" -ForegroundColor Green
    return $OutputPath
}

function New-InteractiveExecutiveSummary {
    <#
    .SYNOPSIS
        Generates a dashboard-style interactive Executive Summary Report
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

    Write-Host "Generating Interactive Executive Summary..." -ForegroundColor Cyan

    # SVG Icons
    $svgOverview = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>'
    $svgEnv = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>'
    $svgAssess = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>'
    $svgBrief = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>'
    $svgUsers = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/></svg>'
    $svgMail = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/></svg>'
    $svgChat = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>'
    $svgDb = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/></svg>'

    $tenantName = $CollectedData.TenantInfo.DisplayName

    # Data extraction
    $licensedUsers = if ($CollectedData.EntraID.Users.Analysis) { $CollectedData.EntraID.Users.Analysis.LicensedUsers } else { 0 }
    $totalMailboxes = if ($CollectedData.Exchange.Mailboxes.Analysis) { $CollectedData.Exchange.Mailboxes.Analysis.TotalMailboxes } else { 0 }
    $totalSites = if ($CollectedData.SharePoint.Sites.Analysis) { $CollectedData.SharePoint.Sites.Analysis.SharePointSites } else { 0 }
    $totalTeams = if ($CollectedData.Teams.Teams.Analysis) { $CollectedData.Teams.Teams.Analysis.TotalTeams } else { 0 }
    $storageGB = if ($CollectedData.SharePoint.Sites.Analysis.TotalStorageUsedGB) { [math]::Round($CollectedData.SharePoint.Sites.Analysis.TotalStorageUsedGB, 1) } else { 0 }
    $criticalCount = if ($AnalysisResults.BySeverity.Critical) { $AnalysisResults.BySeverity.Critical.Count } else { 0 }
    $highCount = if ($AnalysisResults.BySeverity.High) { $AnalysisResults.BySeverity.High.Count } else { 0 }
    $mediumCount = if ($AnalysisResults.BySeverity.Medium) { $AnalysisResults.BySeverity.Medium.Count } else { 0 }
    $lowCount = if ($AnalysisResults.BySeverity.Low) { $AnalysisResults.BySeverity.Low.Count } else { 0 }
    $totalIssues = $criticalCount + $highCount + $mediumCount + $lowCount
    $readinessScore = [math]::Max(0, 100 - $ComplexityScore.TotalScore)
    $readinessColor = if ($readinessScore -ge 80) { "#22c55e" } elseif ($readinessScore -ge 60) { "#eab308" } elseif ($readinessScore -ge 40) { "#f97316" } else { "#ef4444" }
    $readinessLevel = switch ($readinessScore) {
        { $_ -ge 80 } { "Ready for Migration" }
        { $_ -ge 60 } { "Minor Preparation Needed" }
        { $_ -ge 40 } { "Moderate Preparation Required" }
        default { "Significant Work Required" }
    }

    # Build HTML
    $html = Get-InteractiveHTMLHeader -Title "Executive Summary - $tenantName" -ReportType "Executive"

    # Sidebar
    $navItems = @(
        @{ Id = 'page-dashboard'; Label = 'Overview'; Icon = $svgOverview }
        @{ Id = 'page-environment'; Label = 'Environment'; Icon = $svgEnv }
        @{ Id = 'page-assessment'; Label = 'Assessment'; Icon = $svgAssess }
        @{ Id = 'page-ai-analysis'; Label = 'Expert Briefing'; Icon = $svgBrief }
    )
    $html += Get-SidebarHTML -TenantName $tenantName -ReportType "Executive" -NavItems $navItems
    $html += '<div class="main-content">'
    $html += Get-TopBarHTML
    $html += '<main class="page-content">'

    # ==================== PAGE: OVERVIEW ====================
    $html += '<div class="page active" id="page-dashboard">'
    $html += '<div class="page-header"><h1>Migration Readiness Overview</h1><p>' + $tenantName + ' | ' + (Get-Date -Format "MMMM dd, yyyy") + '</p></div>'

    # Giant readiness gauge
    $html += @"
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:24px;margin-bottom:24px;">
        <div class="chart-card" style="text-align:center;padding:40px;">
            <h3 style="margin-bottom:24px;">Migration Readiness Score</h3>
            <div style="position:relative;width:220px;height:220px;margin:0 auto;">
                <svg viewBox="0 0 220 220" style="transform:rotate(-90deg);">
                    <circle cx="110" cy="110" r="95" fill="none" stroke="var(--bg-elevated)" stroke-width="14"/>
                    <circle cx="110" cy="110" r="95" fill="none" stroke="$readinessColor" stroke-width="14" stroke-linecap="round" stroke-dasharray="596.9" stroke-dashoffset="$([math]::Round(596.9 * (1 - $readinessScore / 100), 2))"/>
                </svg>
                <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center;">
                    <div style="font-size:3.5rem;font-weight:700;color:$readinessColor;">$readinessScore%</div>
                    <div style="font-size:0.85rem;color:var(--text-muted);text-transform:uppercase;">Readiness</div>
                </div>
            </div>
            <div class="gauge-status" style="color:$readinessColor;margin-top:24px;">$readinessLevel</div>
        </div>
        <div>
            <div class="kpi-grid" style="grid-template-columns:1fr 1fr;">
$(Get-KPICardHTML -Value $licensedUsers -Label "Users" -IconSvg $svgUsers -Color "#7c3aed")
$(Get-KPICardHTML -Value $totalMailboxes -Label "Mailboxes" -IconSvg $svgMail -Color "#6d28d9")
$(Get-KPICardHTML -Value $totalTeams -Label "Teams" -IconSvg $svgChat -Color "#7c3aed")
$(Get-KPICardHTML -Value "$($storageGB) GB" -Label "Data Volume" -IconSvg $svgDb -Color "#6d28d9")
            </div>
            <div class="data-section" style="margin-top:16px;">
                <h3>Issue Summary</h3>
                <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:12px;">
                    <div style="text-align:center;padding:12px;border-radius:10px;background:rgba(239,68,68,0.08);">
                        <div style="font-size:1.5rem;font-weight:700;color:#ef4444;">$criticalCount</div>
                        <div style="font-size:0.72rem;color:var(--text-muted);text-transform:uppercase;">Critical</div>
                    </div>
                    <div style="text-align:center;padding:12px;border-radius:10px;background:rgba(249,115,22,0.08);">
                        <div style="font-size:1.5rem;font-weight:700;color:#f97316;">$highCount</div>
                        <div style="font-size:0.72rem;color:var(--text-muted);text-transform:uppercase;">High</div>
                    </div>
                    <div style="text-align:center;padding:12px;border-radius:10px;background:rgba(234,179,8,0.08);">
                        <div style="font-size:1.5rem;font-weight:700;color:#eab308;">$mediumCount</div>
                        <div style="font-size:0.72rem;color:var(--text-muted);text-transform:uppercase;">Medium</div>
                    </div>
                    <div style="text-align:center;padding:12px;border-radius:10px;background:rgba(34,197,94,0.08);">
                        <div style="font-size:1.5rem;font-weight:700;color:#22c55e;">$lowCount</div>
                        <div style="font-size:0.72rem;color:var(--text-muted);text-transform:uppercase;">Low</div>
                    </div>
                </div>
            </div>
        </div>
    </div>
"@
    $html += '</div>' # end page-dashboard

    # ==================== PAGE: ENVIRONMENT ====================
    $html += '<div class="page" id="page-environment">'
    $html += '<div class="page-header"><h1>Environment Scope</h1><p>Overview of the M365 environment being assessed</p></div>'
    $html += '<div class="kpi-grid">'
    $html += Get-KPICardHTML -Value $licensedUsers -Label "Licensed Users" -IconSvg $svgUsers -Color "#7c3aed"
    $html += Get-KPICardHTML -Value $totalMailboxes -Label "Mailboxes" -IconSvg $svgMail -Color "#6d28d9"
    $html += Get-KPICardHTML -Value $totalSites -Label "SharePoint Sites" -IconSvg $svgDb -Color "#7c3aed"
    $html += Get-KPICardHTML -Value $totalTeams -Label "Teams" -IconSvg $svgChat -Color "#6d28d9"
    $html += '</div>'
    $html += '<div class="chart-grid" style="grid-template-columns:1fr;">'
    $html += Get-ChartCardHTML -Title "Environment Distribution" -CanvasId "chartEnvDist"
    $html += '</div>'
    $html += @"
<script>createDoughnutChart('chartEnvDist',['Users','Mailboxes','SharePoint','Teams'],[$licensedUsers,$totalMailboxes,$totalSites,$totalTeams],['#3b82f6','#22c55e','#f97316','#8b5cf6']);</script>
"@
    $html += '</div>' # end page-environment

    # ==================== PAGE: ASSESSMENT ====================
    $html += '<div class="page" id="page-assessment">'
    $html += '<div class="page-header"><h1>Risk Assessment</h1><p>Summary of migration risks and complexity</p></div>'
    $html += @"
    <div class="data-section" data-searchable>
        <h3>Assessment Summary</h3>
        <table><thead><tr><th>Category</th><th>Status</th><th>Action Required</th></tr></thead><tbody>
        <tr><td>Critical Blockers</td><td><span class="severity-badge severity-$(if($criticalCount -gt 0){'critical'}else{'low'})">$criticalCount Issues</span></td><td>$(if($criticalCount -gt 0){"Must resolve before migration"}else{"No action needed"})</td></tr>
        <tr><td>High Priority</td><td><span class="severity-badge severity-$(if($highCount -gt 0){'high'}else{'low'})">$highCount Issues</span></td><td>$(if($highCount -gt 0){"Address during preparation"}else{"No action needed"})</td></tr>
        <tr><td>Medium Priority</td><td><span class="severity-badge severity-medium">$mediumCount Issues</span></td><td>Plan during migration</td></tr>
        <tr><td>Low Priority</td><td><span class="severity-badge severity-low">$lowCount Issues</span></td><td>Address post-migration</td></tr>
        <tr><td>Overall Complexity</td><td><strong>$($ComplexityScore.ComplexityLevel)</strong></td><td>Score: $($ComplexityScore.TotalScore)/100</td></tr>
        </tbody></table>
    </div>
"@
    if ($ComplexityScore.TopFactors) {
        $html += '<div class="data-section"><h3>Top Risk Factors</h3><table><thead><tr><th>Factor</th><th>Impact Score</th></tr></thead><tbody>'
        foreach ($factor in $ComplexityScore.TopFactors) {
            $factorName = if ($factor.Key) { $factor.Key } elseif ($factor.Name) { $factor.Name } else { "$factor" }
            $factorScore = if ($factor.Value.WeightedScore) { [math]::Round($factor.Value.WeightedScore, 1) } elseif ($factor.Value) { $factor.Value } else { "N/A" }
            $html += "<tr><td>$factorName</td><td>$factorScore</td></tr>"
        }
        $html += '</tbody></table></div>'
    }
    $html += '</div>' # end page-assessment

    # ==================== PAGE: AI BRIEFING ====================
    $html += '<div class="page" id="page-ai-analysis">'
    $html += '<div class="page-header"><h1>Expert Briefing</h1><p>Executive summary and strategic recommendations</p></div>'

    if ($AIExecutiveSummary -and $AIExecutiveSummary.Success -and $AIExecutiveSummary.Summary) {
        $aiContent = Convert-AIMarkdownToHTML -Markdown $AIExecutiveSummary.Summary

        # Use string concatenation instead of here-string to avoid PowerShell interpolation
        # of $ characters in AI-generated content (e.g., $UserPrincipalName, $true)
        $html += '    <div class="ai-card"><div>' + $aiContent + '</div></div>'
    } else {
        $html += '<div class="ai-card"><p style="color:var(--text-muted);text-align:center;padding:40px;">Expert briefing was not performed. Re-run with analysis enabled for enhanced insights.</p></div>'
    }
    $html += '</div>' # end page-ai-analysis

    # Close layout
    $html += '</main></div>'
    $html += Get-InteractiveHTMLFooter

    $html | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Host "  Interactive Executive Summary generated: $OutputPath" -ForegroundColor Green
    return $OutputPath
}

#endregion

# Export functions
Export-ModuleMember -Function @(
    'Get-InteractiveHTMLHeader',
    'Get-InteractiveHTMLFooter',
    'New-InteractiveITReport',
    'New-InteractiveExecutiveSummary'
)
