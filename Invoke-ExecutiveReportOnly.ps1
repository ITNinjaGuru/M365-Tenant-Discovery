#Requires -Version 7.0
<#
.SYNOPSIS
    Generates only the Executive Summary report from previously collected discovery data
.DESCRIPTION
    This script loads existing discovery data, runs AI executive summary analysis, and generates
    only the Executive Summary report (HTML). Use this when you need just the executive
    overview without the IT technical report, Excel workbook, or playbook scripts.

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
    .\Invoke-ExecutiveReportOnly.ps1 -DiscoveryDataPath "./Output/Discovery_2024-01-29/Data/TenantDiscovery_Full.json" -AIProvider "Opus4.6" -AIApiKey $env:ANTHROPIC_API_KEY
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

Write-Host "================================================" -ForegroundColor Magenta
Write-Host "  Executive Report Only - M365 Migration Tool" -ForegroundColor Magenta
Write-Host "================================================" -ForegroundColor Magenta
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

# Run gotcha analysis (non-AI rule-based checks) — needed for severity counts and risk level
Write-Host ""
Write-Host "Running gotcha analysis..." -ForegroundColor Yellow
$analysisResults = Invoke-GotchaAnalysis -CollectedData $collectedData
Write-Host "  Found $($analysisResults.RulesTriggered) gotchas" -ForegroundColor Green

# Calculate complexity score — needed for readiness score
Write-Host ""
Write-Host "Calculating complexity score..." -ForegroundColor Yellow
$complexityScore = Get-ComplexityScore -CollectedData $collectedData -AnalysisResults $analysisResults
Write-Host "  Complexity Score: $($complexityScore.TotalScore) ($($complexityScore.ComplexityLevel))" -ForegroundColor Green

# Run AI executive summary
Write-Host ""
Write-Host "Running AI executive summary (this may take 1-2 minutes)..." -ForegroundColor Yellow
$aiExecutiveSummary = $null
try {
    $aiExecutiveSummary = Get-AIExecutiveSummary -CollectedData $collectedData -AnalysisResults $analysisResults -ComplexityScore $complexityScore

    if ($aiExecutiveSummary.Success) {
        Write-Host "  AI executive summary complete" -ForegroundColor Green
    }
    else {
        Write-Host "  AI executive summary failed: $($aiExecutiveSummary.Error)" -ForegroundColor Red
    }
}
catch {
    Write-Host "  AI executive summary error: $_" -ForegroundColor Red
}

# Generate interactive Executive Summary report (HTML)
Write-Host ""
Write-Host "Generating Executive Summary Report (HTML)..." -ForegroundColor Yellow
try {
    $execReportPath = Join-Path $reportsPath "Executive-Summary.html"

    New-InteractiveExecutiveSummary -CollectedData $collectedData `
        -AnalysisResults $analysisResults `
        -ComplexityScore $complexityScore `
        -AIExecutiveSummary $aiExecutiveSummary `
        -OutputPath $execReportPath

    Write-Host "  Executive summary (HTML) generated" -ForegroundColor Green
    Write-Host "  Location: $execReportPath" -ForegroundColor Gray
}
catch {
    Write-Host "  Failed to generate executive summary (HTML): $_" -ForegroundColor Red
}

# Summary
Write-Host ""
Write-Host "================================================" -ForegroundColor Magenta
Write-Host "  Executive Report Generation Complete!" -ForegroundColor Magenta
Write-Host "================================================" -ForegroundColor Magenta
Write-Host ""
Write-Host "Reports saved to: $reportsPath" -ForegroundColor Green
Write-Host ""
Write-Host "Files generated:" -ForegroundColor White
Write-Host "  - Executive-Summary.html  (Interactive executive overview with AI insights)" -ForegroundColor Gray
Write-Host ""
Write-Host "Open Executive-Summary.html in your browser to view the executive overview." -ForegroundColor Magenta
Write-Host ""
