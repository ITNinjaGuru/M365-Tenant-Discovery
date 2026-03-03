#Requires -Version 7.0
<#
.SYNOPSIS
    Exports discovery data to CSV files from previously collected data
.DESCRIPTION
    This script loads existing discovery data and exports it to CSV files
    without re-running data collection. Useful for:
    - Generating CSV exports after initial discovery
    - Creating data extracts for migration tools
    - Sharing specific data with stakeholders
    - Analysis in Excel or other tools
.PARAMETER DiscoveryDataPath
    Path to the discovery data JSON file (e.g., ./Output/Discovery_2024-01-29/Data/TenantDiscovery_Full.json)
.PARAMETER OutputPath
    Path where CSV files will be saved (defaults to CSV folder next to discovery data)
.PARAMETER IncludeGotchaAnalysis
    If specified, also runs gotcha analysis and exports findings to CSV
.EXAMPLE
    .\Invoke-CSVExportOnly.ps1 -DiscoveryDataPath ".\Output\Discovery_2024-01-29\Data\TenantDiscovery_Full.json"
.EXAMPLE
    .\Invoke-CSVExportOnly.ps1 -DiscoveryDataPath ".\Output\Discovery_2024-01-29\Data\TenantDiscovery_Full.json" -IncludeGotchaAnalysis
.NOTES
    Author: M365 Migration Tool
    Version: 1.0.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$DiscoveryDataPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeGotchaAnalysis
)

# Import required modules
$scriptRoot = $PSScriptRoot
Import-Module (Join-Path $scriptRoot "Modules" "Core" "TenantDiscovery.Core.psm1") -Force
Import-Module (Join-Path $scriptRoot "Reports" "CSVExporter.psm1") -Force

if ($IncludeGotchaAnalysis) {
    Import-Module (Join-Path $scriptRoot "Analysis" "GotchaAnalysisEngine.psm1") -Force
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  CSV Data Export Tool" -ForegroundColor Cyan
Write-Host "  M365 Migration Discovery" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Load discovery data
Write-Host "Loading discovery data from: $DiscoveryDataPath" -ForegroundColor Yellow
try {
    $collectedData = Get-Content $DiscoveryDataPath -Raw | ConvertFrom-Json -AsHashtable -Depth 100
    Write-Host "OK Discovery data loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "X Failed to load discovery data: $_" -ForegroundColor Red
    exit 1
}

# Determine output path
if (-not $OutputPath) {
    $dataFolder = Split-Path $DiscoveryDataPath -Parent
    $OutputPath = Split-Path $dataFolder -Parent
}

Write-Host ""
Write-Host "Output location: $OutputPath" -ForegroundColor Gray
Write-Host ""

# Run gotcha analysis if requested
$analysisResults = $null
if ($IncludeGotchaAnalysis) {
    Write-Host "Running gotcha analysis..." -ForegroundColor Yellow
    $analysisResults = Invoke-GotchaAnalysis -CollectedData $collectedData
    Write-Host "OK Analysis complete - Found $($analysisResults.RulesTriggered) gotchas" -ForegroundColor Green
    Write-Host ""
}

# Export to CSV
Write-Host "Exporting data to CSV files..." -ForegroundColor Yellow
Write-Host ""

try {
    $exportResult = Export-DiscoveryDataToCSV `
        -CollectedData $collectedData `
        -AnalysisResults $analysisResults `
        -OutputPath $OutputPath

    Write-Host ""
    Write-Host "OK CSV export complete!" -ForegroundColor Green
}
catch {
    Write-Host "X Failed to export CSV files: $_" -ForegroundColor Red
    exit 1
}

# Summary
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Export Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "CSV files saved to: $($exportResult.OutputPath)" -ForegroundColor Green
Write-Host ""
Write-Host "Exported files ($($exportResult.Count)):" -ForegroundColor White

foreach ($file in $exportResult.Files) {
    $fileName = Split-Path $file -Leaf
    $fileSize = (Get-Item $file).Length
    $sizeDisplay = if ($fileSize -gt 1MB) { "{0:N2} MB" -f ($fileSize / 1MB) }
                   elseif ($fileSize -gt 1KB) { "{0:N0} KB" -f ($fileSize / 1KB) }
                   else { "{0} bytes" -f $fileSize }
    Write-Host ("  {0,-35} {1,12}" -f $fileName, $sizeDisplay) -ForegroundColor Gray
}

Write-Host ""
Write-Host "CSV files can be opened in Excel or imported into migration tools." -ForegroundColor Cyan
Write-Host ""
