# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

M365 Tenant Discovery & Migration Assessment Tool — a PowerShell 7+ toolset that connects to Microsoft 365 tenants, collects configuration data across all major workloads, runs 96 migration risk detection rules, optionally sends data to an AI provider (GPT-5.2, Claude Opus 4.6, or Gemini 3) for deeper analysis, and generates interactive HTML/PDF/Excel/CSV reports for IT teams and executives.

## Development Setup

```powershell
# Install required PowerShell modules
Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
Install-Module -Name MicrosoftTeams -Scope CurrentUser -Force
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force

# Clone and navigate to repo
cd tenantdiscovery-claude

# Test module imports
pwsh -Command "Import-Module .\Modules\Core\TenantDiscovery.Core.psm1 -Force"
```

## Running the Tool

### Main Discovery Workflow

```powershell
# Basic interactive run
.\Start-TenantDiscovery.ps1 -SharePointAdminUrl "https://contoso-admin.sharepoint.com"

# With AI analysis
.\Start-TenantDiscovery.ps1 -SharePointAdminUrl "https://contoso-admin.sharepoint.com" -AIProvider "Opus4.6" -AIApiKey $env:ANTHROPIC_API_KEY

# With app registration auth
.\Start-TenantDiscovery.ps1 -SharePointAdminUrl "https://contoso-admin.sharepoint.com" -TenantId "xxx" -ClientId "yyy" -ClientSecret "zzz"

# Skip specific workloads
.\Start-TenantDiscovery.ps1 -SharePointAdminUrl "https://contoso-admin.sharepoint.com" -SkipPowerBI -SkipDynamics

# Generate detailed CSV export in addition to reports
.\Start-TenantDiscovery.ps1 -SharePointAdminUrl "https://contoso-admin.sharepoint.com" -ExportDetailedCSV

# With configuration file
.\Start-TenantDiscovery.ps1 -ConfigPath ".\Config\discovery-config.json"
```

### Utility Scripts (Invoke-*Only)

Five scripts for running specific phases independently:

```powershell
# Re-run full AI analysis on existing discovery data (all reports + Excel + playbook scripts)
.\Invoke-AIAnalysisOnly.ps1 -DiscoveryDataPath "./Output/Discovery_YYYYMMDD_HHMMSS/Data/TenantDiscovery_Full.json" -AIProvider "Opus4.6" -AIApiKey $env:ANTHROPIC_API_KEY

# Generate only the IT Technical Analysis report (HTML + ai-gotcha-analysis.txt)
# Requires -AIProvider and -AIApiKey; runs gotcha analysis + AI gotcha analysis internally
.\Invoke-ITReportOnly.ps1 -DiscoveryDataPath "./Output/Discovery_YYYYMMDD_HHMMSS/Data/TenantDiscovery_Full.json" -AIProvider "Opus4.6" -AIApiKey $env:ANTHROPIC_API_KEY

# Generate only the Executive Summary report (HTML + ai-executive-summary.txt)
# Requires -AIProvider and -AIApiKey; runs gotcha analysis + AI executive summary internally
.\Invoke-ExecutiveReportOnly.ps1 -DiscoveryDataPath "./Output/Discovery_YYYYMMDD_HHMMSS/Data/TenantDiscovery_Full.json" -AIProvider "Opus4.6" -AIApiKey $env:ANTHROPIC_API_KEY

# Export detailed CSVs from collected data (mailbox stats, retention holds, email addresses, etc.)
.\Invoke-CSVExportOnly.ps1 -DiscoveryDataPath "./Output/Discovery_YYYYMMDD_HHMMSS/Data/TenantDiscovery_Full.json"

# Generate communication plan (email templates, SharePoint posts, migration timeline)
.\Invoke-CommunicationsOnly.ps1 -DiscoveryDataPath "./Output/Discovery_YYYYMMDD_HHMMSS/Data/TenantDiscovery_Full.json"

# Run AI-generated migration scripts in isolated PowerShell processes
.\Reports\Run-MigrationScripts.ps1 -ScriptsPath "./Output/Discovery_YYYYMMDD_HHMMSS/Scripts"
```

Use `Invoke-ITReportOnly.ps1` or `Invoke-ExecutiveReportOnly.ps1` when you need to regenerate a single report type without running the full pipeline. Use `Invoke-AIAnalysisOnly.ps1` when you need all outputs (both reports, Excel workbook, and extracted playbook scripts).

## Architecture

### Execution Flow (Start-TenantDiscovery.ps1)
1. **Module Loading** — imports all `.psm1` modules from `Modules/`, `Analysis/`, `Reports/`
2. **Prerequisite Check** — verifies Microsoft.Graph, ExchangeOnlineManagement, MicrosoftTeams, PnP.PowerShell are installed
3. **Service Connection** — `Connect-M365Services` handles auth (interactive or ServicePrincipal) for Graph, Exchange, SharePoint (PnP), Teams, Security & Compliance
4. **Data Collection** — 8 sequential collection phases, each via `Invoke-*Collection` functions; collected data stored in `$script:CollectedData` hashtable in Core module
5. **Gotcha Analysis** — `Invoke-GotchaAnalysis` runs 96 rule-based checks; `Get-ComplexityScore` and `Get-MigrationPriorities` compute risk metrics
6. **AI Analysis** (optional) — 3-batch approach: (1) Executive Summary + Remediation, (2) Planning + Resources, (3) Full Playbook with scripts; scripts extracted to individual `.ps1` files
7. **Report Generation** — Static HTML, PDF, Interactive HTML (Chart.js), Excel workbook, optional detailed CSVs
8. **Communication Plan** (optional) — email templates, SharePoint posts, migration timeline

### Module Dependency Chain
- **Core** (`Modules/Core/TenantDiscovery.Core.psm1`) — all other modules depend on this; provides `Write-Log`, `Add-CollectedData`, `Get-CollectedData`, `Add-MigrationGotcha`, `Connect-M365Services`
- **Workload Modules** (`Modules/{EntraID,Exchange,SharePoint,Teams,PowerBI,Dynamics365,Security,HybridIdentity}/`) — each exports an `Invoke-*Collection` function; stores results via `Add-CollectedData`
- **GotchaAnalysisEngine** (`Analysis/GotchaAnalysisEngine.psm1`) — rule-based risk detection; rules defined as scriptblock `Condition` properties; exports `Invoke-GotchaAnalysis`, `Get-ComplexityScore`, `Get-MigrationPriorities`, `Get-MigrationRoadmap`
- **AIIntegration** (`Analysis/AIIntegration.psm1`) — multi-provider AI abstraction; `Invoke-AIRequest` normalizes request/response across OpenAI, Anthropic, and Google APIs
- **Report Modules** (`Reports/`):
  - `ReportGenerator` — Static HTML/PDF reports with professional styling
  - `InteractiveReportGenerator` — Dashboard-style multi-page HTML with fixed sidebar navigation, liquid glass cards (royal blue accent on deep dark background, `backdrop-filter` blur), 10+ Chart.js interactive visualizations, dark/light theme toggle, responsive mobile design, search/filter across pages; AI analysis sections rendered via `Convert-AIMarkdownToHTML` (internal helper, with collapsed-newline restoration preprocessing) with interactive checkboxes on ordered list items
  - `ExcelWorkbookExporter` — Multi-sheet workbook with executive summary, gotchas, metrics, raw data
  - `CSVExporter` — Quick exports for each data category
  - `DetailedCSVExporter` — Comprehensive CSV exports with mailbox stats, retention holds, email addresses
- **Communications** (`Modules/Communications/`) — optional; generates migration communication templates, email content, timelines

### Key Data Structures
- **`$script:CollectedData`** (Core) — central hashtable with keys: `Metadata`, `EntraID`, `Exchange`, `SharePoint`, `Teams`, `PowerBI`, `Dynamics365`, `Security`, `HybridIdentity`, `Licensing`; each workload stores nested `Analysis` summary + raw data
- **`$script:DiscoveredGotchas`** (Core) — ArrayList of `[PSCustomObject]` with `Id`, `Category`, `Title`, `Description`, `Severity`, `Recommendation`, `AffectedObjects`, `AffectedCount`, `MigrationPhase`
- **Analysis Rules** (GotchaAnalysisEngine) — array of hashtables with `Id`, `Category`, `Condition` (scriptblock), `Severity`, `Description`, `Recommendation`, `RemediationSteps`, `Tools`, `EstimatedEffort`

### AI Provider Integration
Three providers configured in `$script:AIProviders` within AIIntegration.psm1, with provider-specific auth and response handling:
- **GPT-5.2** — OpenAI endpoint, Bearer token auth (`Authorization: Bearer`), model `gpt-5.2`
- **Opus4.6** — Anthropic endpoint, `x-api-key` header auth, model `claude-opus-4-6` (Recommended - best for enterprise analysis)
- **Gemini-3-Pro** — Google endpoint, URL parameter auth (`key=`), model `gemini-3.1-pro-preview`
- **Gemini-3-Flash** — Google endpoint, URL parameter auth (`key=`), model `gemini-3-flash-preview`

The `Invoke-AIRequest` function in AIIntegration.psm1 normalizes provider differences:
- Standardizes request format (system prompt, user message, tokens, temperature)
- Handles provider-specific response structures (content extraction, token counting)
- Implements retry logic and fallback responses
- Enforces token limits per batch (8000 for exec summary, 6000 for planning, 10000 for playbook)

## Conventions

- All modules require PowerShell 7.0+ (`#Requires -Version 7.0`)
- Module functions use `[CmdletBinding()]` and PowerShell comment-based help
- Gotchas are added via `Add-MigrationGotcha` with severity levels: `Critical`, `High`, `Medium`, `Low`, `Informational`
- Risk categories have weights: Compliance (1.8) > Identity (1.5) > Infrastructure (1.4) > Data (1.3) > Integration (1.2) > Operations (1.0)
- Functions are exported explicitly via `Export-ModuleMember -Function @(...)` at bottom of each `.psm1`
- Config file is `Config/discovery-config.json` (copy from `discovery-config.sample.json`); never commit actual config with secrets
- Output goes to `./Output/Discovery_YYYYMMDD_HHMMSS/` with `Data/`, `Reports/`, `Logs/`, optionally `Scripts/` and `Communications/` subdirectories
- Interactive report features (dark/light theme, search, filters) use inline JavaScript; Chart.js loaded from CDN, gracefully degrades if unavailable
- Interactive report theme uses CSS custom properties (variables) defined in `:root`; primary color is royal blue (`#4169E1`, RGB `65,105,225`); cards and panels use transparent `rgba()` backgrounds with `backdrop-filter: blur()` for a liquid glass effect — changing `$primaryColor`/`$primaryRGB` in `Get-InteractiveHTMLHeader` propagates the accent color throughout

## Configuration

Discovery behavior is controlled via `Config/discovery-config.json`:

```json
{
  "Collection": {
    "EntraID": { "Enabled": true, "MaxUsersToProcess": 50000 },
    "Exchange": { "Enabled": true, "IncludeHybridConfig": true },
    "SharePoint": { "AdminUrl": "https://contoso-admin.sharepoint.com" },
    "Teams": { "Enabled": true },
    "PowerBI": { "Enabled": true },
    "Dynamics365": { "Enabled": true },
    "Security": { "Enabled": true },
    "HybridIdentity": { "Enabled": true }
  },
  "AI": {
    "Enabled": true,
    "Provider": "Opus4.6",
    "Options": { "MaxTokens": 12000 }
  },
  "Reporting": {
    "GenerateITReport": true,
    "GenerateExecutiveReport": true,
    "IncludeCharts": true
  },
  "Performance": {
    "TimeoutMinutes": 180,
    "RetryCount": 5
  }
}
```

Key options:
- `MaxUsersToProcess`: Limit EntraID user enumeration for large tenants
- `IncludeHybridConfig`: Flag for Exchange and HybridIdentity modules
- `TimeoutMinutes`: Connection and API timeout threshold
- `IncludeDLPPolicies`: Set to false to skip slow DLP enumeration in Dynamics365

## Development & Debugging

### Running Individual Modules
```powershell
# Load a single module and test a function
Import-Module .\Modules\Core\TenantDiscovery.Core.psm1 -Force
Import-Module .\Analysis\GotchaAnalysisEngine.psm1 -Force

# Get list of analysis rules
Get-AnalysisRules | Select-Object Id, Category, Severity
```

### Common Issues

**Module Import Fails:**
- Verify PowerShell version: `$PSVersionTable.PSVersion` should be 7.x
- Check for syntax errors: `Test-ModuleManifest` doesn't work for .psm1 files; use `[scriptblock]::Create((Get-Content file.psm1)) | Out-Null`
- Reload module with `-Force` flag: `Import-Module path.psm1 -Force`

**M365 Service Connections Fail:**
- Clear cached tokens: `Disconnect-MgGraph; Disconnect-ExchangeOnline -Confirm:$false; Disconnect-SPOService`
- Verify app permissions if using ServicePrincipal auth (Graph and SharePoint permissions must be granted separately)
- Check certificate path and thumbprint if using cert-based auth for SharePoint

**AI Analysis Errors:**
- Set environment variables before script execution: `$env:ANTHROPIC_API_KEY = "..."`
- Verify JSON structure of discovery data: `Get-Content data.json | ConvertFrom-Json` should not throw errors
- Check token limits — if prompts are too large, trim input data via config before AI request

**AI Analysis Content Renders as a Wall of Unformatted Text:**
Three known root causes, all addressed in the current codebase:
1. **Collapsed newlines** — AI response newlines are stripped during JSON serialization or API transport; both `Convert-AIMarkdownToHTML` and `ConvertTo-HTMLFromMarkdown` detect this (length >500 chars, <10 lines) and restore newlines before structural markdown patterns (headers, lists, code fences, tables, blockquotes)
2. **PowerShell variable interpolation** — embedding `$aiContent` inside a double-quoted here-string (`@"...$aiContent..."@`) causes PowerShell to expand `$` characters in AI text as variables; always use string concatenation: `'<div>' + $aiContent + '</div>'`
3. **Weak inline regex conversion** — `ReportGenerator.psm1` static reports previously used ad-hoc regex for markdown conversion; they now call `ConvertTo-HTMLFromMarkdown`; do not reintroduce inline regex substitution for markdown-to-HTML conversion

**Report Generation Hangs:**
- Chart.js CDN may be slow; InteractiveReportGenerator has timeout logic
- Check that JSON data is properly closed (no truncation)
- Verify output path is writable: `Test-Path (Split-Path $OutputPath -Parent) -PathType Container`

**Charts Not Rendering in Executive Summary:**
- Chart helper functions must be defined before they're called. Interactive reports load chart functions in the `<head>` `<script>` block to ensure they're available when inline `<script>` tags execute.
- If charts appear blank, check browser console for JavaScript errors (`F12` > Console tab)

**SharePoint Site Count Inflated:**
- OneDrive sites may be included in the "SharePoint Sites" count if not properly filtered
- The collection uses both template filtering (`SPSPERS#10`) and URL pattern filtering (`*-my.sharepoint.com/personal/*`)
- Separate counts: `SharePointSites`, `OneDriveSites`, `TeamSites`, `CommunicationSites`, `ClassicSites`

**Top Risk Factors Not Showing in Executive Summary:**
- TopFactors from ComplexityScore are dictionary entries with `.Key`/`.Value.WeightedScore` properties
- Ensure the AI analysis completes successfully (requires -AIProvider and -AIApiKey)

**PDF Cover Page Has Visual Artifacts or Rendering Issues:**
- Do not use SVG patterns or `::before`/`::after` pseudo-element overlays on the cover page — they render as scratchy noise in headless Chrome and wkhtmltopdf
- Do not use emoji in cover page content — use plain text or HTML entities (e.g., `&bull;`) instead
- If the cover background disappears, add `-webkit-print-color-adjust: exact; print-color-adjust: exact;` to both the `body` and `.cover-page` rules in the `@media print` block

**PDF Page Breaks Cutting Content Mid-Section:**
- Remove `page-break-inside: avoid` from `.section` — this property is ignored when the section is taller than one page and causes the browser to cut content unpredictably
- Use the `.keep-together` class only on sections known to be compact (a few KPI cards or a risk matrix); never apply it to sections containing long lists or AI-generated text
- Long AI analysis sections should use only `page-break` (force start on new page) and let content flow naturally across subsequent pages

### Adding New Gotcha Rules

In `GotchaAnalysisEngine.psm1`, add to `Get-AnalysisRules`:

```powershell
@{
    Id = "NEW-001"
    Category = "YourCategory"  # Identity, Data, Compliance, Integration, Infrastructure, Operations
    Title = "Descriptive title"
    Description = "What the issue is"
    Condition = { $collectedData.SomeField.Count -gt 100 }  # ScriptBlock that returns $true/$false
    Severity = "High"  # Critical, High, Medium, Low, Informational
    Recommendation = "How to fix it"
    RemediationSteps = @("Step 1", "Step 2")
    Tools = @("Cmdlet-Name")
    EstimatedEffort = "1-2 hours"
    AffectedArea = "YourWorkload"
}
```

The `Condition` scriptblock has access to `$collectedData` (the full collected data hashtable).

**Rule ID prefixes by category:** `ID-` Identity, `DT-` Data, `CP-` Compliance, `IN-` Integration, `IF-` Infrastructure, `OP-` Operations, `MF-` Mail Flow, `EX-` Exchange, `TM-` Teams, `SP-` SharePoint, `DV-` Dataverse, `AP-` Application Proxy, `LI-` Licensing, `CO-` Connectors/Communications, `PP-` Power Platform

## Report Branding Notes

- All references to "AI" have been replaced with neutral terms: "Deep Analysis", "Expert Briefing", etc.
- The tool integrates with optional AI providers but does not require them for core functionality
- Reports work standalone without AI; AI integration enhances analysis when `-AIProvider` and `-AIApiKey` are provided
- Interactive reports use clickable issue cards that expand on click to show details

## When Modifying Code

- When adding a new gotcha rule, follow the structure above and add to `Get-AnalysisRules` in `GotchaAnalysisEngine.psm1`
- When adding a new workload collector, create `Modules/{WorkloadName}/TenantDiscovery.{WorkloadName}.psm1`, export an `Invoke-{WorkloadName}Collection` function, add it to the module loading list in `Start-TenantDiscovery.ps1`, and store data via `Add-CollectedData`
- Reports embed all CSS/JS inline (Chart.js via CDN) — no external file dependencies for generated reports
- Interactive reports: chart functions must be defined in `<head>` before inline `<script>` calls; chart initialization is wrapped in retry logic for Chart.js CDN delays
- The AI analysis uses a 3-batch pattern to avoid token limits; if modifying prompts, keep each batch within its token allocation (8000/6000/10000)
- InteractiveReportGenerator creates a single multi-page HTML file; test theme toggle and sidebar navigation on both desktop and mobile viewport sizes; the liquid glass aesthetic uses `rgba()` transparent backgrounds with `backdrop-filter: blur()` — elements that look blank may be missing a solid background on the `<body>` behind them
- AI analysis content is converted from markdown to HTML by two functions: `Convert-AIMarkdownToHTML` (in `InteractiveReportGenerator.psm1`, used by `New-InteractiveITReport` and `New-InteractiveExecutiveSummary`) and `ConvertTo-HTMLFromMarkdown` (in `ReportGenerator.psm1`, used by `New-ITDetailedReport`, `New-ExecutiveSummaryReport`, and the PDF gotcha section); both functions include collapsed-newline restoration preprocessing — if AI content arrives as a single long line (e.g., due to JSON transport), they detect this (length >500 chars and <10 lines) and re-insert newlines before headers, lists, code fences, tables, and blockquotes before converting
- When embedding AI-converted HTML content into a larger HTML string, always use string concatenation (`'<div>' + $aiContent + '</div>'`), never double-quoted here-strings (`@"...$aiContent..."@`) — PowerShell will expand `$` characters inside AI content (e.g., `$UserPrincipalName`, `$true`) as variable references, silently corrupting the output
- When combining AI batch responses in `AIIntegration.psm1`, always cast each batch result to `[string]` before concatenation — `Invoke-AIRequest` can return complex objects and the cast ensures clean string joining
- Ordered list items (`<ol><li>`) inside `.ai-card` are interactive checkboxes — clicking toggles the `.checked` CSS class (strikethrough + green check via `::before`); do not use native `<input type="checkbox">` inside these items
- Migration sequencing (not dependency graphs) is preferred in AI prompts — numbered sequences are more readable than ASCII diagrams

### PDF Report Rendering Rules

PDF output is generated via headless Chrome (`--print-to-pdf`) or wkhtmltopdf. Both renderers have constraints that differ from regular browser rendering:

**Cover page (`New-PDFCoverPage` in `ReportGenerator.psm1`):**
- Use a solid `background` color, not CSS gradients with SVG overlays — SVG `::before` pattern overlays render as scratchy artifacts in headless Chrome
- Use HTML text content or HTML entities for icons, not emoji characters — emoji rendering is unreliable in headless/wkhtmltopdf environments
- Structural decoration is done with `.cover-accent-bar` and `.cover-accent-bar-bottom` divs (solid color strips), not pseudo-element patterns
- Use standard `position: absolute` for footer placement, not flexbox vertical centering — flexbox min-height behavior differs in print contexts

**Page break strategy (CSS in `Get-PDFStylesheet`):**
- `.section` does NOT have `page-break-inside: avoid` — sections can span multiple pages; applying it caused mid-content cuts on long sections
- Use `.section.keep-together` for compact sections that must fit on one page (executive summary header, risk matrix, migration timeline)
- Use `.section.page-break` to force a section to start on a new page
- Both classes can be combined: `<div class="section page-break keep-together">` starts on a new page and stays together
- All severity group sections (Critical, High, Medium, Low) use `page-break` to start on a fresh page
- AI analysis sections use `page-break` only, not `keep-together` — AI content spans many pages and must flow naturally
- Compact elements with `page-break-inside: avoid` in the `@media print` block: `.gotcha-card`, `.timeline-item`, `.summary-grid`, `.kpi-grid`, `.risk-matrix`, `.readiness-hero`, `.exec-hero`, `pre`, `table`
- Headings use `page-break-after: avoid` to prevent orphaned headers at the bottom of a page
- Paragraphs use `orphans: 3; widows: 3` to prevent single-line stranding at page edges
