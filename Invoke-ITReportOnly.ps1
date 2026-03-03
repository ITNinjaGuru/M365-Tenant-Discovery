#Requires -Version 7.0
<#
.SYNOPSIS
    Generates only the IT Technical Analysis report from previously collected discovery data
.DESCRIPTION
    This script loads existing discovery data, runs AI gotcha analysis, and generates
    only the IT Technical Analysis report (HTML). Use this when you need just
    the technical report without the executive summary, Excel workbook, or playbook scripts.

    For the full pipeline (all reports + scripts), use Invoke-AIAnalysisOnly.ps1 instead.
.PARAMETER DiscoveryDataPath
    Path to the discovery data JSON file (e.g., ./Output/Discovery_YYYYMMDD_HHMMSS/Data/TenantDiscovery_Full.json)
.PARAMETER AIProvider
    AI provider to use (GPT-5.2, Opus4.6, or Gemini-3-Pro)
.PARAMETER AIApiKey
    API key for the selected AI provider
.PARAMETER OutputPath
    Path where reports will be saved (defaults to parent folder of discovery data)
.EXAMPLE
    .\Invoke-ITReportOnly.ps1 -DiscoveryDataPath "./Output/Discovery_2024-01-29/Data/TenantDiscovery_Full.json" -AIProvider "Opus4.6" -AIApiKey $env:ANTHROPIC_API_KEY
.NOTES
    Author: M365 Migration Tool
    Version: 1.0.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$DiscoveryDataPath,

    [Parameter(Mandatory = $true)]
    [ValidateSet("GPT-5.2", "Opus4.6", "Gemini-3-Pro", "Gemini-3-Flash")]
    [string]$AIProvider,

    [Parameter(Mandatory = $true)]
    [string]$AIApiKey,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

# Import required modules
$scriptRoot = $PSScriptRoot
Import-Module (Join-Path $scriptRoot "Modules" "Core" "TenantDiscovery.Core.psm1") -Force
Import-Module (Join-Path $scriptRoot "Analysis" "GotchaAnalysisEngine.psm1") -Force
Import-Module (Join-Path $scriptRoot "Analysis" "AIIntegration.psm1") -Force
Import-Module (Join-Path $scriptRoot "Reports" "ReportGenerator.psm1") -Force
Import-Module (Join-Path $scriptRoot "Reports" "InteractiveReportGenerator.psm1") -Force

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  IT Report Only - M365 Migration Tool" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Load discovery data
Write-Host "Loading discovery data from: $DiscoveryDataPath" -ForegroundColor Yellow
try {
    $collectedData = Get-Content $DiscoveryDataPath -Raw | ConvertFrom-Json -AsHashtable -Depth 100
    Write-Host "  Discovery data loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "  Failed to load discovery data: $_" -ForegroundColor Red
    exit 1
}

# Determine output path
if (-not $OutputPath) {
    $dataFolder = Split-Path $DiscoveryDataPath -Parent
    $OutputPath = Split-Path $dataFolder -Parent
}
$reportsPath = Join-Path $OutputPath "Reports"
if (-not (Test-Path $reportsPath)) {
    New-Item -Path $reportsPath -ItemType Directory -Force | Out-Null
}

Write-Host ""
Write-Host "Configuring AI provider: $AIProvider..." -ForegroundColor Yellow
Set-AIProvider -Provider $AIProvider -APIKey $AIApiKey

# Run gotcha analysis (non-AI rule-based checks)
Write-Host ""
Write-Host "Running gotcha analysis..." -ForegroundColor Yellow
$analysisResults = Invoke-GotchaAnalysis -CollectedData $collectedData
Write-Host "  Found $($analysisResults.RulesTriggered) gotchas" -ForegroundColor Green

# Calculate complexity score
Write-Host ""
Write-Host "Calculating complexity score..." -ForegroundColor Yellow
$complexityScore = Get-ComplexityScore -CollectedData $collectedData -AnalysisResults $analysisResults
Write-Host "  Complexity Score: $($complexityScore.TotalScore) ($($complexityScore.ComplexityLevel))" -ForegroundColor Green

# Build tenant context for AI - use LICENSED users (not total directory users)
$sharedMailboxCount = @($collectedData.Exchange.Mailboxes.Mailboxes | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" }).Count
$tenantContext = @{
    UserCount           = $collectedData.EntraID.Users.Analysis.LicensedUsers
    TotalDirectoryUsers = $collectedData.EntraID.Users.Analysis.TotalUsers
    SharedMailboxCount  = $sharedMailboxCount
    HybridEnabled       = $collectedData.HybridIdentity.AADConnect.Configuration.OnPremisesSyncEnabled
    MailboxCount        = $collectedData.Exchange.Mailboxes.Analysis.TotalMailboxes
    SiteCount           = $collectedData.SharePoint.Sites.Analysis.SharePointSites
    TeamCount           = $collectedData.Teams.Teams.Analysis.TotalTeams
    GuestCount          = $collectedData.EntraID.Users.Analysis.GuestUsers
    SyncedUserCount     = $collectedData.EntraID.Users.Analysis.SyncedUsers
    HybridDeviceCount   = if ($collectedData.EntraID.Devices.Analysis.HybridJoined) { $collectedData.EntraID.Devices.Analysis.HybridJoined } else { 0 }
    PowerBIWorkspaces   = if ($collectedData.PowerBI.Workspaces.Analysis.TotalWorkspaces) { $collectedData.PowerBI.Workspaces.Analysis.TotalWorkspaces } else { 0 }
    PowerBIGateways     = if ($collectedData.PowerBI.Gateways.Analysis.TotalGateways) { $collectedData.PowerBI.Gateways.Analysis.TotalGateways } else { 0 }
    D365Environments    = if ($collectedData.Dynamics365.Environments.Analysis.TotalEnvironments) { $collectedData.Dynamics365.Environments.Analysis.TotalEnvironments } else { 0 }
    D365Apps            = if ($collectedData.Dynamics365.PowerApps.Analysis.TotalApps) { $collectedData.Dynamics365.PowerApps.Analysis.TotalApps } else { 0 }
    D365Flows           = if ($collectedData.Dynamics365.PowerAutomate.Analysis.TotalFlows) { $collectedData.Dynamics365.PowerAutomate.Analysis.TotalFlows } else { 0 }
}

# Run AI gotcha analysis
Write-Host ""
Write-Host "Running AI gotcha analysis (this may take 1-2 minutes)..." -ForegroundColor Yellow
$aiGotchaAnalysis = $null
try {
    $allGotchas = @()
    foreach ($category in $analysisResults.ByCategory.Keys) {
        $allGotchas += $analysisResults.ByCategory[$category]
    }

    $aiGotchaAnalysis = Get-AIGotchaAnalysis -Gotchas $allGotchas -TenantContext $tenantContext

    if ($aiGotchaAnalysis.Success) {
        Write-Host "  AI analysis complete" -ForegroundColor Green
    }
    else {
        Write-Host "  AI gotcha analysis failed: $($aiGotchaAnalysis.Error)" -ForegroundColor Red
    }
}
catch {
    Write-Host "  AI gotcha analysis error: $_" -ForegroundColor Red
}

# Generate interactive IT report (HTML)
Write-Host ""
Write-Host "Generating IT Analysis Report (HTML)..." -ForegroundColor Yellow
try {
    $itReportPath = Join-Path $reportsPath "IT-Analysis-Report.html"

    New-InteractiveITReport -CollectedData $collectedData `
        -AnalysisResults $analysisResults `
        -ComplexityScore $complexityScore `
        -AIAnalysis $aiGotchaAnalysis `
        -OutputPath $itReportPath

    Write-Host "  IT report (HTML) generated" -ForegroundColor Green
    Write-Host "  Location: $itReportPath" -ForegroundColor Gray
}
catch {
    Write-Host "  Failed to generate IT report (HTML): $_" -ForegroundColor Red
}

# Summary
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  IT Report Generation Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Reports saved to: $reportsPath" -ForegroundColor Green
Write-Host ""
Write-Host "Files generated:" -ForegroundColor White
Write-Host "  - IT-Analysis-Report.html  (Interactive technical report with AI insights)" -ForegroundColor Gray
Write-Host ""
Write-Host "Open IT-Analysis-Report.html in your browser to view the detailed technical analysis." -ForegroundColor Cyan
Write-Host ""
