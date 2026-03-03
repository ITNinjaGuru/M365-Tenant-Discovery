#Requires -Version 7.0
<#
.SYNOPSIS
    Generates communication plan from previously collected discovery data
.DESCRIPTION
    This script loads existing discovery data and generates the communication plan
    without re-running data collection or AI analysis. Useful for:
    - Generating communications after initial discovery
    - Re-generating communications with different branding
    - Updating target migration date
.PARAMETER DiscoveryDataPath
    Path to the discovery data JSON file (e.g., ./Output/Discovery_2024-01-29/Data/discovery-data.json)
.PARAMETER TargetMigrationDate
    Target migration date (format: YYYY-MM-DD). Defaults to 60 days from today.
.PARAMETER CompanyName
    Company name for branding. Defaults to "Your Organization".
.PARAMETER PrimaryColor
    Primary brand color (hex). Defaults to "#0078d4" (Microsoft Blue).
.PARAMETER SecondaryColor
    Secondary brand color (hex). Defaults to "#106ebe".
.PARAMETER AIProvider
    AI provider to use for content generation (GPT-5.2, Opus4.6, or Gemini-3-Pro). Optional.
.PARAMETER AIApiKey
    API key for AI provider. Required if AIProvider is specified.
.PARAMETER OutputPath
    Path where communication files will be saved (defaults to same folder as discovery data)
.EXAMPLE
    .\Invoke-CommunicationsOnly.ps1 -DiscoveryDataPath "./Output/Discovery_2024-01-29/Data/discovery-data.json"
.EXAMPLE
    .\Invoke-CommunicationsOnly.ps1 -DiscoveryDataPath "./Output/Discovery_2024-01-29/Data/discovery-data.json" -TargetMigrationDate "2024-03-15" -CompanyName "Contoso Inc"
.EXAMPLE
    .\Invoke-CommunicationsOnly.ps1 -DiscoveryDataPath "./Output/Discovery_2024-01-29/Data/discovery-data.json" -AIProvider "GPT-5.2" -AIApiKey "sk-..."
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
    [string]$TargetMigrationDate,

    [Parameter(Mandatory = $false)]
    [string]$CompanyName = "Your Organization",

    [Parameter(Mandatory = $false)]
    [string]$PrimaryColor = "#0078d4",

    [Parameter(Mandatory = $false)]
    [string]$SecondaryColor = "#106ebe",

    [Parameter(Mandatory = $false)]
    [ValidateSet("GPT-5.2", "Opus4.6", "Gemini-3-Pro", "Gemini-3-Flash")]
    [string]$AIProvider,

    [Parameter(Mandatory = $false)]
    [string]$AIApiKey,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

# Import required modules
$scriptRoot = $PSScriptRoot
Import-Module (Join-Path $scriptRoot "Modules" "Core" "TenantDiscovery.Core.psm1") -Force
Import-Module (Join-Path $scriptRoot "Analysis" "GotchaAnalysisEngine.psm1") -Force
Import-Module (Join-Path $scriptRoot "Modules" "Communications" "TenantDiscovery.Communications.psm1") -Force

# Import AI Integration if AI provider specified
if ($AIProvider) {
    Import-Module (Join-Path $scriptRoot "Analysis" "AIIntegration.psm1") -Force
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Communication Plan Generator" -ForegroundColor Cyan
Write-Host "  M365 Migration Tool" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Load discovery data
Write-Host "Loading discovery data from: $DiscoveryDataPath" -ForegroundColor Yellow
try {
    $collectedData = Get-Content $DiscoveryDataPath -Raw | ConvertFrom-Json -AsHashtable -Depth 100
    Write-Host "✓ Discovery data loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "✗ Failed to load discovery data: $_" -ForegroundColor Red
    exit 1
}

# Determine output path
if (-not $OutputPath) {
    $dataFolder = Split-Path $DiscoveryDataPath -Parent
    $OutputPath = Split-Path $dataFolder -Parent
}
$communicationsPath = Join-Path $OutputPath "Communications"
if (-not (Test-Path $communicationsPath)) {
    New-Item -Path $communicationsPath -ItemType Directory -Force | Out-Null
}

# Parse target migration date
$migrationDate = if ($TargetMigrationDate) {
    try {
        [datetime]::ParseExact($TargetMigrationDate, "yyyy-MM-dd", $null)
    }
    catch {
        Write-Host "✗ Invalid date format. Please use YYYY-MM-DD format." -ForegroundColor Red
        exit 1
    }
}
else {
    (Get-Date).AddDays(60)
}

Write-Host ""
Write-Host "Configuration:" -ForegroundColor White
Write-Host "  Company Name:    $CompanyName" -ForegroundColor Gray
Write-Host "  Target Date:     $($migrationDate.ToString('yyyy-MM-dd'))" -ForegroundColor Gray
Write-Host "  Output Path:     $communicationsPath" -ForegroundColor Gray
if ($AIProvider) {
    Write-Host "  AI Provider:     $AIProvider" -ForegroundColor Gray
}
Write-Host ""

# Configure AI provider if specified
if ($AIProvider -and $AIApiKey) {
    Write-Host "Configuring AI provider: $AIProvider..." -ForegroundColor Yellow
    Set-AIProvider -Provider $AIProvider -APIKey $AIApiKey
    Write-Host "✓ AI provider configured" -ForegroundColor Green
    Write-Host ""
}

# Run gotcha analysis (needed for communication plan context)
Write-Host "Running gotcha analysis..." -ForegroundColor Yellow
$analysisResults = Invoke-GotchaAnalysis -CollectedData $collectedData
Write-Host "✓ Analysis complete - Found $($analysisResults.RulesTriggered) gotchas" -ForegroundColor Green
Write-Host ""

# Build branding
$branding = @{
    CompanyName    = $CompanyName
    LogoUrl        = $null
    PrimaryColor   = $PrimaryColor
    SecondaryColor = $SecondaryColor
}

# Generate communication plan
Write-Host "Generating communication plan..." -ForegroundColor Yellow
Write-Host "  This includes:" -ForegroundColor Gray
Write-Host "    - 7 phased email templates" -ForegroundColor Gray
Write-Host "    - 7 SharePoint news posts" -ForegroundColor Gray
Write-Host "    - Visual migration timeline" -ForegroundColor Gray
Write-Host "    - Support resources page" -ForegroundColor Gray
if ($AIProvider) {
    Write-Host "    - AI-generated content" -ForegroundColor Gray
}
Write-Host ""

try {
    $communicationPlan = New-CommunicationPlan `
        -AnalysisResults $analysisResults `
        -CollectedData $collectedData `
        -TargetMigrationDate $migrationDate `
        -Branding $branding `
        -OutputPath $communicationsPath

    Write-Host "✓ Communication plan generated successfully" -ForegroundColor Green
}
catch {
    Write-Host "✗ Failed to generate communication plan: $_" -ForegroundColor Red
    exit 1
}

# Summary
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Communication Plan Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Files saved to: $communicationsPath" -ForegroundColor Green
Write-Host ""
Write-Host "Generated content:" -ForegroundColor White
Write-Host "  Emails/              - 7 email templates (HTML & plain text)" -ForegroundColor Gray
Write-Host "  SharePointPosts/     - 7 SharePoint news articles" -ForegroundColor Gray
Write-Host "  Timeline/            - Interactive migration timeline" -ForegroundColor Gray
Write-Host "  Resources/           - User support resources page" -ForegroundColor Gray
Write-Host "  communication-plan-summary.html - Index page with all links" -ForegroundColor Gray
Write-Host ""
Write-Host "Communication phases:" -ForegroundColor White
$phases = @(
    @{ Name = "Pre-Announcement"; Timing = -42 }
    @{ Name = "Detailed Notice"; Timing = -21 }
    @{ Name = "One Week Reminder"; Timing = -7 }
    @{ Name = "Day Before"; Timing = -1 }
    @{ Name = "Migration Day"; Timing = 0 }
    @{ Name = "Completion"; Timing = 1 }
    @{ Name = "Post-Migration"; Timing = 7 }
)
foreach ($phase in $phases) {
    $phaseDate = $migrationDate.AddDays($phase.Timing)
    Write-Host ("  {0,-20} - {1}" -f $phase.Name, $phaseDate.ToString("MMM d, yyyy")) -ForegroundColor Gray
}
Write-Host ""
Write-Host "Open communication-plan-summary.html in your browser to view all communications." -ForegroundColor Cyan
Write-Host ""
