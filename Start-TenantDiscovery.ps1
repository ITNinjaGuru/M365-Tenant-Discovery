#Requires -Version 7.0
<#
.SYNOPSIS
    M365 Tenant Discovery and Migration Assessment Tool
.DESCRIPTION
    Comprehensive tool for discovering and analyzing Microsoft 365 tenant configuration
    to identify migration gotchas and risks for tenant-to-tenant migrations.

    Supports AI-powered analysis using GPT-5.2, Claude Opus 4.6, or Google Gemini 3 Pro.

    Generates detailed IT technical reports and executive summaries.

.PARAMETER ConfigPath
    Path to JSON configuration file. If not provided, uses default configuration.

.PARAMETER OutputPath
    Directory for output files. Defaults to ./Output

.PARAMETER SharePointAdminUrl
    SharePoint admin center URL (e.g., https://contoso-admin.sharepoint.com)

.PARAMETER AIProvider
    AI provider for enhanced analysis. Options: GPT-5.2, Opus4.6, Gemini-3-Pro

.PARAMETER AIApiKey
    API key for the selected AI provider

.PARAMETER SkipAI
    Skip AI-powered analysis

.PARAMETER SkipExchange
    Skip Exchange Online collection

.PARAMETER SkipSharePoint
    Skip SharePoint Online collection

.PARAMETER SkipTeams
    Skip Microsoft Teams collection

.PARAMETER SkipPowerBI
    Skip Power BI collection

.PARAMETER SkipDynamics
    Skip Dynamics 365 / Power Platform collection

.PARAMETER SkipSecurity
    Skip Security & Compliance collection

.PARAMETER SkipCommunications
    Skip communication plan generation (email templates, SharePoint posts, timelines)

.PARAMETER TargetMigrationDate
    Target migration date for communication plan scheduling. Defaults to 60 days from discovery.

.PARAMETER Interactive
    Use interactive authentication (default)

.PARAMETER ExportDetailedCSV
    Export detailed CSV files with comprehensive mailbox information including retention holds,
    archive details, mailbox statistics, and all email addresses. Optional - not run by default.

.PARAMETER TenantId
    Azure AD Tenant ID for app-based authentication

.PARAMETER ClientId
    Azure AD Application (Client) ID for app-based authentication

.PARAMETER ClientSecret
    Azure AD Application Client Secret for app-based authentication

.EXAMPLE
    .\Start-TenantDiscovery.ps1 -SharePointAdminUrl "https://contoso-admin.sharepoint.com"

.EXAMPLE
    .\Start-TenantDiscovery.ps1 -SharePointAdminUrl "https://contoso-admin.sharepoint.com" -TenantId "xxx" -ClientId "yyy" -ClientSecret "zzz"

.EXAMPLE
    .\Start-TenantDiscovery.ps1 -SharePointAdminUrl "https://contoso-admin.sharepoint.com" -AIProvider "Opus4.6" -AIApiKey "your-api-key"

.EXAMPLE
    .\Start-TenantDiscovery.ps1 -ConfigPath ".\Config\discovery-config.json" -OutputPath "C:\Reports"

.EXAMPLE
    .\Start-TenantDiscovery.ps1 -SharePointAdminUrl "https://contoso-admin.sharepoint.com" -ExportDetailedCSV

.NOTES
    Author: AI Migration Expert
    Version: 1.0.0
    Requires: PowerShell 7.x
    Modules: Microsoft.Graph, ExchangeOnlineManagement, MicrosoftTeams, PnP.PowerShell
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\Output",

    [Parameter(Mandatory = $false)]
    [string]$SharePointAdminUrl,

    [Parameter(Mandatory = $false)]
    [ValidateSet("GPT-5.2", "Opus4.6", "Gemini-3-Pro", "Gemini-3-Flash")]
    [string]$AIProvider,

    [Parameter(Mandatory = $false)]
    [string]$AIApiKey,

    [Parameter(Mandatory = $false)]
    [switch]$SkipAI,

    [Parameter(Mandatory = $false)]
    [switch]$SkipExchange,

    [Parameter(Mandatory = $false)]
    [switch]$SkipSharePoint,

    [Parameter(Mandatory = $false)]
    [switch]$SkipTeams,

    [Parameter(Mandatory = $false)]
    [switch]$SkipPowerBI,

    [Parameter(Mandatory = $false)]
    [switch]$SkipDynamics,

    [Parameter(Mandatory = $false)]
    [switch]$SkipSecurity,

    [Parameter(Mandatory = $false)]
    [switch]$SkipCommunications,

    [Parameter(Mandatory = $false)]
    [datetime]$TargetMigrationDate,

    [Parameter(Mandatory = $false)]
    [switch]$Interactive,

    [Parameter(Mandatory = $false)]
    [switch]$ExportDetailedCSV,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret
)

#region Initialization
$ErrorActionPreference = "Stop"
$script:StartTime = Get-Date

# Banner
Write-Host @"

╔═══════════════════════════════════════════════════════════════════════════════╗
║                                                                               ║
║   ███╗   ███╗██████╗  ██████╗ ███████╗    ██████╗ ██╗███████╗ ██████╗        ║
║   ████╗ ████║╚════██╗██╔════╝ ██╔════╝    ██╔══██╗██║██╔════╝██╔════╝        ║
║   ██╔████╔██║ █████╔╝███████╗ ███████╗    ██║  ██║██║███████╗██║             ║
║   ██║╚██╔╝██║ ╚═══██╗██╔═══██╗╚════██║    ██║  ██║██║╚════██║██║             ║
║   ██║ ╚═╝ ██║██████╔╝╚██████╔╝███████║    ██████╔╝██║███████║╚██████╗        ║
║   ╚═╝     ╚═╝╚═════╝  ╚═════╝ ╚══════╝    ╚═════╝ ╚═╝╚══════╝ ╚═════╝        ║
║                                                                               ║
║              Tenant Discovery & Migration Assessment Tool                     ║
║                           Version 1.0.0                                       ║
║                                                                               ║
╚═══════════════════════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Cyan

Write-Host "Starting M365 Tenant Discovery..." -ForegroundColor Green
Write-Host "Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host ""

# Get script root
$ScriptRoot = $PSScriptRoot
if (-not $ScriptRoot) {
    $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $ScriptRoot) {
    $ScriptRoot = Get-Location
}

# Import modules
Write-Host "Loading modules..." -ForegroundColor Yellow

$modules = @(
    @{ Path = "Modules\Core\TenantDiscovery.Core.psm1"; Required = $true }
    @{ Path = "Modules\EntraID\TenantDiscovery.EntraID.psm1"; Required = $true }
    @{ Path = "Modules\Exchange\TenantDiscovery.Exchange.psm1"; Required = $true }
    @{ Path = "Modules\SharePoint\TenantDiscovery.SharePoint.psm1"; Required = $true }
    @{ Path = "Modules\Teams\TenantDiscovery.Teams.psm1"; Required = $true }
    @{ Path = "Modules\PowerBI\TenantDiscovery.PowerBI.psm1"; Required = $true }
    @{ Path = "Modules\Dynamics365\TenantDiscovery.Dynamics365.psm1"; Required = $true }
    @{ Path = "Modules\Security\TenantDiscovery.Security.psm1"; Required = $true }
    @{ Path = "Modules\HybridIdentity\TenantDiscovery.HybridIdentity.psm1"; Required = $true }
    @{ Path = "Analysis\GotchaAnalysisEngine.psm1"; Required = $true }
    @{ Path = "Analysis\AIIntegration.psm1"; Required = $true }
    @{ Path = "Reports\ReportGenerator.psm1"; Required = $true }
    @{ Path = "Reports\InteractiveReportGenerator.psm1"; Required = $true }
    @{ Path = "Reports\ExcelWorkbookExporter.psm1"; Required = $true }
    @{ Path = "Modules\Communications\TenantDiscovery.Communications.psm1"; Required = $false }
)

foreach ($module in $modules) {
    $modulePath = Join-Path $ScriptRoot $module.Path
    if (Test-Path $modulePath) {
        try {
            Import-Module $modulePath -Force -Global -ErrorAction Stop
            Write-Host "  [OK] Loaded: $($module.Path)" -ForegroundColor Green
        }
        catch {
            if ($module.Required) {
                Write-Host "  [FAIL] Failed to load required module: $($module.Path)" -ForegroundColor Red
                Write-Host "  Error: $_" -ForegroundColor Red
                exit 1
            }
            else {
                Write-Host "  [WARN] Failed to load optional module: $($module.Path)" -ForegroundColor Yellow
            }
        }
    }
    else {
        if ($module.Required) {
            Write-Host "  [FAIL] Required module not found: $($module.Path)" -ForegroundColor Red
            exit 1
        }
    }
}

Write-Host ""
#endregion

#region Check Prerequisites
Write-Host "Checking prerequisites..." -ForegroundColor Yellow

# Fallback function in case module export fails
if (-not (Get-Command Test-TDModuleAvailability -ErrorAction SilentlyContinue)) {
    function Test-TDModuleAvailability {
        param([string[]]$ModuleNames)
        $results = @{}
        foreach ($module in $ModuleNames) {
            $available = Get-Module -ListAvailable -Name $module
            $results[$module] = @{
                Available = $null -ne $available
                Version   = if ($available) { $available[0].Version.ToString() } else { $null }
            }
        }
        return $results
    }
}

$requiredModules = @(
    "Microsoft.Graph"
    "ExchangeOnlineManagement"
    "MicrosoftTeams"
    "PnP.PowerShell"
)

$moduleStatus = Test-TDModuleAvailability -ModuleNames $requiredModules

$allAvailable = $true
foreach ($module in $requiredModules) {
    if ($moduleStatus[$module].Available) {
        Write-Host "  [OK] $module v$($moduleStatus[$module].Version)" -ForegroundColor Green
    }
    else {
        Write-Host "  [MISSING] $module" -ForegroundColor Red
        $allAvailable = $false
    }
}

if (-not $allAvailable) {
    Write-Host ""
    Write-Host "Some required modules are missing. Install them using:" -ForegroundColor Yellow
    Write-Host "  Install-Module -Name Microsoft.Graph -Scope CurrentUser" -ForegroundColor Cyan
    Write-Host "  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Cyan
    Write-Host "  Install-Module -Name MicrosoftTeams -Scope CurrentUser" -ForegroundColor Cyan
    Write-Host "  Install-Module -Name PnP.PowerShell -Scope CurrentUser" -ForegroundColor Cyan
    Write-Host ""

    $install = Read-Host "Would you like to install missing modules now? (Y/N)"
    if ($install -eq 'Y' -or $install -eq 'y') {
        $installResult = Install-RequiredModules
        if ($installResult.Failed.Count -gt 0) {
            Write-Host "Failed to install some modules. Please install manually and retry." -ForegroundColor Red
            exit 1
        }
    }
    else {
        Write-Host "Please install required modules and run again." -ForegroundColor Yellow
        exit 1
    }
}

Write-Host ""
#endregion

#region Initialize
Write-Host "Initializing discovery environment..." -ForegroundColor Yellow

$initResult = Initialize-TenantDiscovery -OutputPath $OutputPath

if (-not $initResult.Success) {
    Write-Host "Failed to initialize: $($initResult.Error)" -ForegroundColor Red
    exit 1
}

Write-Host "  Output Directory: $($initResult.OutputPath)" -ForegroundColor Gray
Write-Host "  Log File: $($initResult.LogPath)" -ForegroundColor Gray
Write-Host ""
#endregion

#region Load Configuration
$script:Config = $null
if ($ConfigPath -and (Test-Path $ConfigPath)) {
    Write-Host "Loading configuration from: $ConfigPath" -ForegroundColor Yellow
    try {
        $script:Config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        Write-Host "  [OK] Configuration loaded" -ForegroundColor Green
    }
    catch {
        Write-Host "  [WARN] Failed to parse config file: $_" -ForegroundColor Yellow
        Write-Host "  Continuing with default settings..." -ForegroundColor Gray
    }
}
else {
    Write-Host "No configuration file specified. Using interactive authentication." -ForegroundColor Gray
}
Write-Host ""
#endregion

#region Connect to Services
$authMethod = "Interactive"
$authConfig = $null

# Check for command-line credentials first (takes precedence over config)
if ($TenantId -and $ClientId -and $ClientSecret) {
    Write-Host "  Using credentials from command-line parameters" -ForegroundColor Gray
    $authMethod = "ServicePrincipal"
    $authConfig = @{
        Method       = "ServicePrincipal"
        TenantId     = $TenantId
        ClientId     = $ClientId
        ClientSecret = $ClientSecret
    }
}
elseif ($script:Config -and $script:Config.Authentication) {
    Write-Host "  Using credentials from config file" -ForegroundColor Gray
    $authConfig = $script:Config.Authentication
    $authMethod = $authConfig.Method
}

Write-Host "  Auth Method: $authMethod" -ForegroundColor Gray

if ($authMethod -eq "ServicePrincipal") {
    Write-Host "Connecting to Microsoft 365 services using app registration..." -ForegroundColor Yellow
    if (-not $authConfig.TenantId -or -not $authConfig.ClientId -or -not $authConfig.ClientSecret) {
        Write-Host "  [FAIL] ServicePrincipal auth requires TenantId, ClientId, and ClientSecret in config" -ForegroundColor Red
        Write-Host "  TenantId present: $([bool]$authConfig.TenantId)" -ForegroundColor Red
        Write-Host "  ClientId present: $([bool]$authConfig.ClientId)" -ForegroundColor Red
        Write-Host "  ClientSecret present: $([bool]$authConfig.ClientSecret)" -ForegroundColor Red
        exit 1
    }
    Write-Host "  TenantId: $($authConfig.TenantId)" -ForegroundColor Gray
    Write-Host "  ClientId: $($authConfig.ClientId)" -ForegroundColor Gray
    Write-Host "  ClientSecret: [PRESENT]" -ForegroundColor Gray
}
else {
    Write-Host "Connecting to Microsoft 365 services..." -ForegroundColor Yellow
    Write-Host "You will be prompted to authenticate." -ForegroundColor Gray
    Write-Host "  (To use app registration, set Method to 'ServicePrincipal' in config)" -ForegroundColor DarkGray
}
Write-Host ""

$servicesToConnect = @("Graph")

if (-not $SkipExchange) { $servicesToConnect += "ExchangeOnline" }
if (-not $SkipTeams) { $servicesToConnect += "Teams" }
if (-not $SkipSecurity) { $servicesToConnect += "Security" }
if (-not $SkipSharePoint) { $servicesToConnect += "SharePoint" }

# Build connection parameters
$connectParams = @{
    Services = $servicesToConnect
}

if ($authMethod -eq "ServicePrincipal" -and $authConfig) {
    $connectParams.AuthConfig = $authConfig
}
else {
    $connectParams.Interactive = $true
}

# Get SharePoint admin URL from config or parameter
if (-not $SkipSharePoint) {
    if ($SharePointAdminUrl) {
        $connectParams.SharePointAdminUrl = $SharePointAdminUrl
    }
    elseif ($authConfig -and $authConfig.SharePoint -and $authConfig.SharePoint.AdminUrl) {
        $connectParams.SharePointAdminUrl = $authConfig.SharePoint.AdminUrl
    }
    elseif ($script:Config -and $script:Config.Collection -and $script:Config.Collection.SharePoint -and $script:Config.Collection.SharePoint.AdminUrl) {
        $connectParams.SharePointAdminUrl = $script:Config.Collection.SharePoint.AdminUrl
    }
}

$connections = Connect-M365Services @connectParams

$allConnected = $true
$failedServices = @()
foreach ($service in $servicesToConnect) {
    if ($connections[$service].Connected) {
        Write-Host "  [OK] Connected to $service" -ForegroundColor Green
    }
    else {
        Write-Host "  [FAIL] Failed to connect to $service" -ForegroundColor Red
        if ($connections[$service].Error) {
            Write-Host "        Error: $($connections[$service].Error)" -ForegroundColor Red
        }
        $failedServices += $service
        $allConnected = $false
    }
}

# Graph is required - exit if it failed
if ($failedServices -contains "Graph") {
    Write-Host ""
    Write-Host "Microsoft Graph connection is required. Please check credentials and try again." -ForegroundColor Red
    exit 1
}

# Warn about other failed services but continue
if (-not $allConnected) {
    Write-Host ""
    Write-Host "Warning: Some services failed to connect. Data collection will skip: $($failedServices -join ', ')" -ForegroundColor Yellow
    Write-Host "Continuing with available services..." -ForegroundColor Yellow
}

Write-Host ""
#endregion

#region Data Collection
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "                    DATA COLLECTION                            " -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$collectionErrors = @()

# Entra ID Collection
Write-Host "[1/8] Collecting Entra ID data..." -ForegroundColor Yellow
try {
    $entraResults = Invoke-EntraIDCollection
    Write-Host "      Completed in $($entraResults.Duration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Gray
    if ($entraResults.Errors.Count -gt 0) {
        $collectionErrors += $entraResults.Errors
    }
}
catch {
    Write-Host "      Error: $_" -ForegroundColor Red
    $collectionErrors += @{ Collection = "EntraID"; Error = $_.Exception.Message }
}

# Exchange Collection
if (-not $SkipExchange -and $failedServices -notcontains "ExchangeOnline") {
    Write-Host "[2/8] Collecting Exchange Online data..." -ForegroundColor Yellow
    try {
        $exchangeResults = Invoke-ExchangeCollection
        Write-Host "      Completed in $($exchangeResults.Duration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Gray
        if ($exchangeResults.Errors.Count -gt 0) {
            $collectionErrors += $exchangeResults.Errors
        }
    }
    catch {
        Write-Host "      Error: $_" -ForegroundColor Red
        $collectionErrors += @{ Collection = "Exchange"; Error = $_.Exception.Message }
    }
}
elseif ($failedServices -contains "ExchangeOnline") {
    Write-Host "[2/8] Skipping Exchange Online (connection failed)" -ForegroundColor Yellow
}
else {
    Write-Host "[2/8] Skipping Exchange Online (--SkipExchange)" -ForegroundColor Gray
}

# SharePoint Collection
if (-not $SkipSharePoint -and $failedServices -notcontains "SharePoint") {
    if ($SharePointAdminUrl) {
        Write-Host "[3/8] Collecting SharePoint Online data..." -ForegroundColor Yellow
        try {
            $spResults = Invoke-SharePointCollection -AdminUrl $SharePointAdminUrl
            Write-Host "      Completed in $($spResults.Duration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Gray
            if ($spResults.Errors.Count -gt 0) {
                $collectionErrors += $spResults.Errors
            }
        }
        catch {
            Write-Host "      Error: $_" -ForegroundColor Red
            $collectionErrors += @{ Collection = "SharePoint"; Error = $_.Exception.Message }
        }
    }
    else {
        Write-Host "[3/8] Skipping SharePoint (no admin URL provided)" -ForegroundColor Gray
        Write-Host "      Use -SharePointAdminUrl to include SharePoint discovery" -ForegroundColor Gray
    }
}
elseif ($failedServices -contains "SharePoint") {
    Write-Host "[3/8] Skipping SharePoint Online (connection failed)" -ForegroundColor Yellow
}
else {
    Write-Host "[3/8] Skipping SharePoint Online (--SkipSharePoint)" -ForegroundColor Gray
}

# Teams Collection
if (-not $SkipTeams -and $failedServices -notcontains "Teams") {
    Write-Host "[4/8] Collecting Microsoft Teams data..." -ForegroundColor Yellow
    try {
        $teamsResults = Invoke-TeamsCollection
        Write-Host "      Completed in $($teamsResults.Duration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Gray
        if ($teamsResults.Errors.Count -gt 0) {
            $collectionErrors += $teamsResults.Errors
        }
    }
    catch {
        Write-Host "      Error: $_" -ForegroundColor Red
        $collectionErrors += @{ Collection = "Teams"; Error = $_.Exception.Message }
    }
}
elseif ($failedServices -contains "Teams") {
    Write-Host "[4/8] Skipping Microsoft Teams (connection failed)" -ForegroundColor Yellow
}
else {
    Write-Host "[4/8] Skipping Microsoft Teams (--SkipTeams)" -ForegroundColor Gray
}

# Power BI Collection
if (-not $SkipPowerBI) {
    Write-Host "[5/8] Collecting Power BI data..." -ForegroundColor Yellow
    try {
        $pbiResults = Invoke-PowerBICollection
        Write-Host "      Completed in $($pbiResults.Duration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Gray
        if ($pbiResults.Errors.Count -gt 0) {
            $collectionErrors += $pbiResults.Errors
        }
    }
    catch {
        Write-Host "      Error: $_" -ForegroundColor Red
        $collectionErrors += @{ Collection = "PowerBI"; Error = $_.Exception.Message }
    }
}
else {
    Write-Host "[5/8] Skipping Power BI (--SkipPowerBI)" -ForegroundColor Gray
}

# Dynamics 365 Collection
if (-not $SkipDynamics) {
    Write-Host "[6/8] Collecting Dynamics 365 / Power Platform data..." -ForegroundColor Yellow
    try {
        $d365Results = Invoke-Dynamics365Collection
        Write-Host "      Completed in $($d365Results.Duration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Gray
        if ($d365Results.Errors.Count -gt 0) {
            $collectionErrors += $d365Results.Errors
        }
    }
    catch {
        Write-Host "      Error: $_" -ForegroundColor Red
        $collectionErrors += @{ Collection = "Dynamics365"; Error = $_.Exception.Message }
    }
}
else {
    Write-Host "[6/8] Skipping Dynamics 365 (--SkipDynamics)" -ForegroundColor Gray
}

# Security & Compliance Collection
if (-not $SkipSecurity -and $failedServices -notcontains "Security") {
    Write-Host "[7/8] Collecting Security & Compliance data..." -ForegroundColor Yellow
    try {
        $secResults = Invoke-SecurityCollection
        Write-Host "      Completed in $($secResults.Duration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Gray
        if ($secResults.Errors.Count -gt 0) {
            $collectionErrors += $secResults.Errors
        }
    }
    catch {
        Write-Host "      Error: $_" -ForegroundColor Red
        $collectionErrors += @{ Collection = "Security"; Error = $_.Exception.Message }
    }
}
elseif ($failedServices -contains "Security") {
    Write-Host "[7/8] Skipping Security & Compliance (connection failed)" -ForegroundColor Yellow
}
else {
    Write-Host "[7/8] Skipping Security & Compliance (--SkipSecurity)" -ForegroundColor Gray
}

# Hybrid Identity Collection
Write-Host "[8/8] Collecting Hybrid Identity data..." -ForegroundColor Yellow
try {
    $hybridResults = Invoke-HybridIdentityCollection
    Write-Host "      Completed in $($hybridResults.Duration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Gray
    if ($hybridResults.Errors.Count -gt 0) {
        $collectionErrors += $hybridResults.Errors
    }
}
catch {
    Write-Host "      Error: $_" -ForegroundColor Red
    $collectionErrors += @{ Collection = "HybridIdentity"; Error = $_.Exception.Message }
}

Write-Host ""
Write-Host "Data collection completed." -ForegroundColor Green
if ($collectionErrors.Count -gt 0) {
    Write-Host "  Warning: $($collectionErrors.Count) collection error(s) occurred" -ForegroundColor Yellow
}
Write-Host ""
#endregion

#region Export Data
Write-Host "Exporting collected data..." -ForegroundColor Yellow
$exportPath = Export-CollectedData
Write-Host "  Data exported to: $exportPath" -ForegroundColor Gray
Write-Host ""
#endregion

#region Analysis
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "                    ANALYSIS                                   " -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

Write-Host "Running gotcha analysis..." -ForegroundColor Yellow

$collectedData = Get-CollectedData
$analysisResults = Invoke-GotchaAnalysis -CollectedData $collectedData
$complexityScore = Get-ComplexityScore -CollectedData $collectedData -AnalysisResults $analysisResults
$priorities = Get-MigrationPriorities -AnalysisResults $analysisResults
$roadmap = Get-MigrationRoadmap -AnalysisResults $analysisResults -CollectedData $collectedData

Write-Host ""
Write-Host "Analysis Summary:" -ForegroundColor Cyan
Write-Host "  Risk Level: $($analysisResults.RiskLevel)" -ForegroundColor $(
    switch ($analysisResults.RiskLevel) {
        "Critical" { "Red" }
        "High" { "Yellow" }
        "Medium" { "Yellow" }
        default { "Green" }
    }
)
Write-Host "  Risk Score: $($analysisResults.RiskPercentage)%" -ForegroundColor Gray
Write-Host "  Complexity: $($complexityScore.ComplexityLevel) ($($complexityScore.TotalScore)/100)" -ForegroundColor Gray
Write-Host ""
Write-Host "  Issues Found:" -ForegroundColor Cyan
Write-Host "    Critical: $($analysisResults.BySeverity.Critical.Count)" -ForegroundColor Red
Write-Host "    High:     $($analysisResults.BySeverity.High.Count)" -ForegroundColor Yellow
Write-Host "    Medium:   $($analysisResults.BySeverity.Medium.Count)" -ForegroundColor Yellow
Write-Host "    Low:      $($analysisResults.BySeverity.Low.Count)" -ForegroundColor Green
Write-Host ""
#endregion

#region AI Analysis
$aiGotchaAnalysis = $null
$aiExecutiveSummary = $null

if (-not $SkipAI -and $AIProvider -and $AIApiKey) {
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "                    AI ANALYSIS                                " -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Configuring AI provider: $AIProvider..." -ForegroundColor Yellow
    Set-AIProvider -Provider $AIProvider -APIKey $AIApiKey

    Write-Host "Testing AI connection..." -ForegroundColor Yellow
    $testResult = Test-AIConnection
    if ($testResult.Success) {
        Write-Host "  AI connection successful" -ForegroundColor Green
        Write-Host ""

        # Get all gotchas for AI analysis
        $allGotchas = Get-MigrationGotchas -SortBySeverity

        # Create tenant context - use LICENSED users (not total directory users)
        $sharedMailboxCount = @($collectedData.Exchange.Mailboxes.Mailboxes | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" }).Count
        $tenantContext = @{
            UserCount           = $collectedData.EntraID.Users.Analysis.LicensedUsers  # Licensed users only
            TotalDirectoryUsers = $collectedData.EntraID.Users.Analysis.TotalUsers     # For reference
            SharedMailboxCount  = $sharedMailboxCount
            MailboxCount        = $collectedData.Exchange.Mailboxes.Analysis.TotalMailboxes
            SiteCount           = $collectedData.SharePoint.Sites.Analysis.SharePointSites
            TeamCount           = $collectedData.Teams.Teams.Analysis.TotalTeams
            HybridEnabled       = $collectedData.HybridIdentity.AADConnect.Configuration.OnPremisesSyncEnabled
            GuestCount          = $collectedData.EntraID.Users.Analysis.GuestUsers
            SyncedUserCount     = $collectedData.EntraID.Users.Analysis.SyncedUsers
            HybridDeviceCount   = if ($collectedData.EntraID.Devices.Analysis.HybridJoined) { $collectedData.EntraID.Devices.Analysis.HybridJoined } else { 0 }
            PowerBIWorkspaces   = if ($collectedData.PowerBI.Workspaces.Analysis.TotalWorkspaces) { $collectedData.PowerBI.Workspaces.Analysis.TotalWorkspaces } else { 0 }
            PowerBIGateways     = if ($collectedData.PowerBI.Gateways.Analysis.TotalGateways) { $collectedData.PowerBI.Gateways.Analysis.TotalGateways } else { 0 }
            D365Environments    = if ($collectedData.Dynamics365.Environments.Analysis.TotalEnvironments) { $collectedData.Dynamics365.Environments.Analysis.TotalEnvironments } else { 0 }
            D365Apps            = if ($collectedData.Dynamics365.PowerApps.Analysis.TotalApps) { $collectedData.Dynamics365.PowerApps.Analysis.TotalApps } else { 0 }
            D365Flows           = if ($collectedData.Dynamics365.PowerAutomate.Analysis.TotalFlows) { $collectedData.Dynamics365.PowerAutomate.Analysis.TotalFlows } else { 0 }
        }

        # AI Gotcha Analysis
        Write-Host "Generating AI analysis of migration gotchas..." -ForegroundColor Yellow
        $aiGotchaAnalysis = Get-AIGotchaAnalysis -Gotchas $allGotchas -TenantContext $tenantContext
        if ($aiGotchaAnalysis.Success) {
            Write-Host "  AI gotcha analysis complete" -ForegroundColor Green

            # Save playbook scripts to separate files
            if ($aiGotchaAnalysis.Playbook) {
                Write-Host "  Extracting migration scripts from playbook..." -ForegroundColor Yellow
                $playbookResult = Save-MigrationPlaybook -PlaybookContent $aiGotchaAnalysis.Playbook -OutputPath $outputPath -CollectedData $collectedData
                if ($playbookResult.Success) {
                    Write-Host "  Saved $($playbookResult.SavedScripts.Count) migration scripts to: $($playbookResult.ScriptsPath)" -ForegroundColor Green
                }
            }
        }
        else {
            Write-Host "  AI gotcha analysis failed: $($aiGotchaAnalysis.Error)" -ForegroundColor Yellow
        }

        # AI Executive Summary
        Write-Host "Generating AI executive summary..." -ForegroundColor Yellow
        $aiExecutiveSummary = Get-AIExecutiveSummary -CollectedData $collectedData -AnalysisResults $analysisResults -ComplexityScore $complexityScore
        if ($aiExecutiveSummary.Success) {
            Write-Host "  AI executive summary complete" -ForegroundColor Green
        }
        else {
            Write-Host "  AI executive summary failed: $($aiExecutiveSummary.Error)" -ForegroundColor Yellow
        }

        Write-Host ""
    }
    else {
        Write-Host "  AI connection failed: $($testResult.Error)" -ForegroundColor Yellow
        Write-Host "  Continuing without AI analysis..." -ForegroundColor Yellow
        Write-Host ""
    }
}
elseif (-not $SkipAI -and (-not $AIProvider -or -not $AIApiKey)) {
    Write-Host "AI analysis skipped (no provider/key specified)" -ForegroundColor Gray
    Write-Host "  Use -AIProvider and -AIApiKey to enable AI analysis" -ForegroundColor Gray
    Write-Host ""
}
#endregion

#region Report Generation
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "                    REPORT GENERATION                          " -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$reportsPath = Join-Path $initResult.OutputPath "Reports"

# IT Detailed Report (interactive version is the primary report)
Write-Host "Generating IT Detailed Report..." -ForegroundColor Yellow
$itReportPath = Join-Path $reportsPath "IT_Technical_Report.html"
try {
    New-InteractiveITReport -CollectedData $collectedData -AnalysisResults $analysisResults -ComplexityScore $complexityScore -AIAnalysis $aiGotchaAnalysis -OutputPath $itReportPath
    Write-Host "  IT Report: $itReportPath" -ForegroundColor Gray
} catch {
    Write-Host "  Failed to generate IT Report: $_" -ForegroundColor Red
}

# Executive Summary Report (interactive version is the primary report)
Write-Host "Generating Executive Summary Report..." -ForegroundColor Yellow
$execReportPath = Join-Path $reportsPath "Executive_Summary.html"
try {
    New-InteractiveExecutiveSummary -CollectedData $collectedData -AnalysisResults $analysisResults -ComplexityScore $complexityScore -AIExecutiveSummary $aiExecutiveSummary -OutputPath $execReportPath
    Write-Host "  Executive Report: $execReportPath" -ForegroundColor Gray
} catch {
    Write-Host "  Failed to generate Executive Report: $_" -ForegroundColor Red
}

Write-Host ""

# Excel Workbook Export (single file with all data)
Write-Host "Generating Excel Workbook (comprehensive data export)..." -ForegroundColor Yellow

$excelPath = Join-Path $reportsPath "Tenant_Discovery_Data.xlsx"
try {
    New-TenantDiscoveryWorkbook -CollectedData $collectedData -AnalysisResults $analysisResults -ComplexityScore $complexityScore -OutputPath $excelPath
    if (Test-Path $excelPath) {
        Write-Host "  Excel Workbook: $excelPath" -ForegroundColor Gray
    }
} catch {
    Write-Host "  Excel Workbook: Could not generate - $_" -ForegroundColor Yellow
    Write-Host "  Tip: For best results, install ImportExcel module: Install-Module -Name ImportExcel" -ForegroundColor Yellow
}

Write-Host ""

# Detailed CSV Exports (optional - not run by default)
if ($ExportDetailedCSV) {
    Write-Host "Generating Detailed CSV Files (comprehensive data for record-keeping)..." -ForegroundColor Yellow

    try {
        # Import the DetailedCSVExporter module
        $csvExporterPath = Join-Path $PSScriptRoot "Reports\DetailedCSVExporter.psm1"
        if (Test-Path $csvExporterPath) {
            Import-Module $csvExporterPath -Force

            # Export all detailed CSVs
            $csvFiles = Export-AllDetailedCSVs -CollectedData $collectedData -OutputPath $reportsPath

            if ($csvFiles -and $csvFiles.Count -gt 0) {
                foreach ($csvFile in $csvFiles) {
                    $fileName = Split-Path -Path $csvFile -Leaf
                    Write-Host "  $fileName" -ForegroundColor Gray
                }
            }
        } else {
            Write-Host "  DetailedCSVExporter module not found. Skipping CSV exports." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  Detailed CSV Export: Could not generate - $_" -ForegroundColor Yellow
    }

    Write-Host ""
}

#endregion

#region Communication Plan Generation
$communicationsEnabled = $false
if (-not $SkipCommunications) {
    if ($Config -and $Config.Communications -and $Config.Communications.Enabled) {
        $communicationsEnabled = $true
    }
}

if ($communicationsEnabled -and (Get-Command New-CommunicationPlan -ErrorAction SilentlyContinue)) {
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "                COMMUNICATION PLAN GENERATION                  " -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Generating communication plan..." -ForegroundColor Yellow

    # Determine target migration date - command line parameter takes precedence
    $commTargetDate = (Get-Date).AddDays(60)  # Default to 60 days from now
    if ($TargetMigrationDate) {
        $commTargetDate = $TargetMigrationDate
        Write-Host "  Using target migration date from command line: $($commTargetDate.ToString('yyyy-MM-dd'))" -ForegroundColor Gray
    }
    elseif ($Config.Communications.TargetMigrationDate) {
        try {
            $commTargetDate = [datetime]::Parse($Config.Communications.TargetMigrationDate)
            Write-Host "  Using target migration date from config: $($commTargetDate.ToString('yyyy-MM-dd'))" -ForegroundColor Gray
        }
        catch {
            Write-Host "  Invalid target migration date in config, using default (60 days from now)" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "  Using default target migration date: $($commTargetDate.ToString('yyyy-MM-dd')) (60 days from now)" -ForegroundColor Gray
    }

    # Get branding from config
    $branding = @{
        CompanyName    = "Your Organization"
        LogoUrl        = $null
        PrimaryColor   = "#0078d4"
        SecondaryColor = "#106ebe"
    }
    if ($Config.Reporting.Branding) {
        $branding.CompanyName = $Config.Reporting.Branding.CompanyName
        $branding.LogoUrl = $Config.Reporting.Branding.LogoUrl
        $branding.PrimaryColor = $Config.Reporting.Branding.PrimaryColor
        $branding.SecondaryColor = $Config.Reporting.Branding.SecondaryColor
    }

    # Set communications output path
    $commOutputPath = Join-Path $initResult.OutputPath "Communications"

    try {
        $communicationPlan = New-CommunicationPlan `
            -AnalysisResults $analysisResults `
            -CollectedData $collectedData `
            -TargetMigrationDate $commTargetDate `
            -Branding $branding `
            -OutputPath $commOutputPath

        Write-Host ""
        Write-Host "  Communication plan generated successfully:" -ForegroundColor Green
        Write-Host "    - 7 Email templates (HTML + Plain Text)" -ForegroundColor Gray
        Write-Host "    - 7 SharePoint news posts" -ForegroundColor Gray
        Write-Host "    - Visual migration timeline (Chart.js)" -ForegroundColor Gray
        Write-Host "    - Support resources page" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  Output location: $commOutputPath" -ForegroundColor Gray
        Write-Host "  Open: $commOutputPath\communication-plan-summary.html" -ForegroundColor Cyan
    }
    catch {
        Write-Host "  Failed to generate communication plan: $_" -ForegroundColor Yellow
        Write-Host "  This is optional - continuing with discovery completion." -ForegroundColor Gray
    }

    Write-Host ""
}
elseif ($communicationsEnabled) {
    Write-Host "Communication plan generation skipped (module not loaded)" -ForegroundColor Gray
    Write-Host ""
}
#endregion

#region Cleanup
Write-Host "Disconnecting from services..." -ForegroundColor Yellow
Disconnect-M365Services
Write-Host ""
#endregion

#region Summary
$endTime = Get-Date
$totalDuration = $endTime - $script:StartTime

Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "                    DISCOVERY COMPLETE                         " -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  Total Duration: $($totalDuration.ToString('hh\:mm\:ss'))" -ForegroundColor Gray
Write-Host "  Output Directory: $($initResult.OutputPath)" -ForegroundColor Gray
Write-Host ""
Write-Host "Reports Generated:" -ForegroundColor Cyan
Write-Host "  IT Technical Report: $itReportPath" -ForegroundColor Gray
Write-Host "  Executive Summary: $execReportPath" -ForegroundColor Gray
if ($ExportDetailedCSV -and $csvFiles -and $csvFiles.Count -gt 0) {
    Write-Host "  Detailed CSV Files: $($csvFiles.Count) files in $(Join-Path $reportsPath 'DetailedCSV')" -ForegroundColor Gray
}
Write-Host ""
Write-Host "Risk Assessment:" -ForegroundColor Cyan
Write-Host "  Overall Risk: $($analysisResults.RiskLevel)" -ForegroundColor $(
    switch ($analysisResults.RiskLevel) {
        "Critical" { "Red" }
        "High" { "Yellow" }
        "Medium" { "Yellow" }
        default { "Green" }
    }
)
Write-Host "  Critical Issues: $($analysisResults.BySeverity.Critical.Count)" -ForegroundColor $(if ($analysisResults.BySeverity.Critical.Count -gt 0) { "Red" } else { "Green" })
Write-Host "  Total Findings: $($analysisResults.RulesTriggered)" -ForegroundColor Gray
Write-Host ""

if ($analysisResults.BySeverity.Critical.Count -gt 0) {
    Write-Host "ATTENTION: Critical issues found that require immediate attention!" -ForegroundColor Red
    Write-Host "Review the IT Technical Report for details and remediation steps." -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "Thank you for using M365 Tenant Discovery Tool!" -ForegroundColor Cyan
Write-Host ""
#endregion
