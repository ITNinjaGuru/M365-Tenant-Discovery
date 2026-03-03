#Requires -Version 7.0
<#
.SYNOPSIS
    Core module for M365 Tenant Discovery and Migration Assessment Tool
.DESCRIPTION
    Provides core functionality including configuration management, logging,
    module initialization, and common utilities for tenant discovery operations.
.NOTES
    Author: AI Migration Expert
    Version: 1.0.0
    Target: PowerShell 7.x
#>

#region Module Variables
$script:ModuleVersion = "1.0.0"
$script:LogPath = $null
$script:Config = $null
$script:CollectedData = @{}
$script:DiscoveredGotchas = [System.Collections.ArrayList]::new()
#endregion

#region Configuration Management
function Initialize-TenantDiscovery {
    <#
    .SYNOPSIS
        Initializes the Tenant Discovery environment
    .PARAMETER ConfigPath
        Path to the configuration JSON file
    .PARAMETER OutputPath
        Path for output files and reports
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ConfigPath,

        [Parameter(Mandatory = $false)]
        [string]$OutputPath = ".\Output"
    )

    try {
        # Create output directory structure
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $script:OutputRoot = Join-Path $OutputPath "Discovery_$timestamp"

        $directories = @(
            $script:OutputRoot,
            (Join-Path $script:OutputRoot "Data"),
            (Join-Path $script:OutputRoot "Reports"),
            (Join-Path $script:OutputRoot "Logs")
        )

        foreach ($dir in $directories) {
            if (-not (Test-Path $dir)) {
                New-Item -Path $dir -ItemType Directory -Force | Out-Null
            }
        }

        # Initialize logging
        $script:LogPath = Join-Path $script:OutputRoot "Logs" "TenantDiscovery_$timestamp.log"

        # Load configuration
        if ($ConfigPath -and (Test-Path $ConfigPath)) {
            $script:Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json -AsHashtable
            Write-Log -Message "Configuration loaded from: $ConfigPath" -Level Info
        }
        else {
            $script:Config = Get-DefaultConfiguration
            Write-Log -Message "Using default configuration" -Level Info
        }

        # Initialize collected data structure
        $script:CollectedData = @{
            Metadata = @{
                CollectionStartTime = Get-Date
                CollectionEndTime   = $null
                TenantId            = $null
                TenantName          = $null
                CollectorVersion    = $script:ModuleVersion
            }
            EntraID       = @{}
            Exchange      = @{}
            SharePoint    = @{}
            Teams         = @{}
            PowerBI       = @{}
            Dynamics365   = @{}
            Security      = @{}
            HybridIdentity = @{}
            Licensing     = @{}
        }

        Write-Log -Message "Tenant Discovery initialized successfully" -Level Info
        Write-Log -Message "Output directory: $script:OutputRoot" -Level Info

        return @{
            Success    = $true
            OutputPath = $script:OutputRoot
            LogPath    = $script:LogPath
        }
    }
    catch {
        Write-Error "Failed to initialize Tenant Discovery: $_"
        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}

function Get-DefaultConfiguration {
    <#
    .SYNOPSIS
        Returns the default configuration for tenant discovery
    #>
    return @{
        Collection = @{
            EntraID = @{
                Enabled             = $true
                IncludeUsers        = $true
                IncludeGroups       = $true
                IncludeDevices      = $true
                IncludeApps         = $true
                IncludeServicePrincipals = $true
                IncludeConditionalAccess = $true
                IncludeRoles        = $true
                MaxUsersToProcess   = 50000
            }
            Exchange = @{
                Enabled               = $true
                IncludeMailboxes      = $true
                IncludeDistributionLists = $true
                IncludePublicFolders  = $true
                IncludeTransportRules = $true
                IncludeConnectors     = $true
            }
            SharePoint = @{
                Enabled           = $true
                IncludeSites      = $true
                IncludeOneDrive   = $true
                IncludeSharingSettings = $true
            }
            Teams = @{
                Enabled          = $true
                IncludeTeams     = $true
                IncludePolicies  = $true
                IncludeTemplates = $true
            }
            PowerBI = @{
                Enabled            = $true
                IncludeWorkspaces  = $true
                IncludeDatasets    = $true
                IncludeGateways    = $true
            }
            Dynamics365 = @{
                Enabled             = $true
                IncludeEnvironments = $true
                IncludeUsers        = $true
                IncludeSolutions    = $true
            }
            Security = @{
                Enabled             = $true
                IncludeDLP          = $true
                IncludeRetention    = $true
                IncludeSensitivity  = $true
                IncludeeDiscovery   = $true
            }
            HybridIdentity = @{
                Enabled              = $true
                IncludeAADConnect    = $true
                IncludeFederation    = $true
                IncludePassthrough   = $true
            }
        }
        AI = @{
            Enabled  = $true
            Provider = "Opus4.6"  # Options: GPT-5.2, Opus4.6, Gemini-3-Pro
            ApiKey   = $null
            Endpoint = $null
        }
        Reporting = @{
            GenerateITReport        = $true
            GenerateExecutiveReport = $true
            IncludeRawData          = $true
            CompressOutput          = $true
        }
    }
}

function Get-TenantDiscoveryConfig {
    <#
    .SYNOPSIS
        Returns the current configuration
    #>
    return $script:Config
}

function Set-TenantDiscoveryConfig {
    <#
    .SYNOPSIS
        Updates the configuration
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )

    $script:Config = $Config
    Write-Log -Message "Configuration updated" -Level Info
}
#endregion

#region Logging Functions
function Write-Log {
    <#
    .SYNOPSIS
        Writes a log entry to file and optionally to console
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error", "Debug", "Success")]
        [string]$Level = "Info",

        [Parameter(Mandatory = $false)]
        [switch]$NoConsole
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    # Write to log file if path is set
    if ($script:LogPath) {
        Add-Content -Path $script:LogPath -Value $logEntry -ErrorAction SilentlyContinue
    }

    # Write to console with color coding
    if (-not $NoConsole) {
        $color = switch ($Level) {
            "Info"    { "Cyan" }
            "Warning" { "Yellow" }
            "Error"   { "Red" }
            "Debug"   { "Gray" }
            "Success" { "Green" }
            default   { "White" }
        }
        Write-Host $logEntry -ForegroundColor $color
    }
}

function Write-Progress-Status {
    <#
    .SYNOPSIS
        Displays a progress status with spinner
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Activity,

        [Parameter(Mandatory = $true)]
        [string]$Status,

        [Parameter(Mandatory = $false)]
        [int]$PercentComplete = -1
    )

    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
}
#endregion

#region Data Collection Helpers
function Add-CollectedData {
    <#
    .SYNOPSIS
        Adds collected data to the central data store
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Category,

        [Parameter(Mandatory = $true)]
        [string]$SubCategory,

        [Parameter(Mandatory = $true)]
        $Data
    )

    if (-not $script:CollectedData.ContainsKey($Category)) {
        $script:CollectedData[$Category] = @{}
    }

    $script:CollectedData[$Category][$SubCategory] = $Data
    Write-Log -Message "Data collected: $Category/$SubCategory" -Level Debug
}

function Get-CollectedData {
    <#
    .SYNOPSIS
        Retrieves collected data
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Category,

        [Parameter(Mandatory = $false)]
        [string]$SubCategory
    )

    if ($Category -and $SubCategory) {
        return $script:CollectedData[$Category][$SubCategory]
    }
    elseif ($Category) {
        return $script:CollectedData[$Category]
    }
    else {
        return $script:CollectedData
    }
}

function Export-CollectedData {
    <#
    .SYNOPSIS
        Exports collected data to JSON files
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$OutputPath
    )

    $path = if ($OutputPath) { $OutputPath } else { Join-Path $script:OutputRoot "Data" }

    try {
        # Update metadata
        $script:CollectedData.Metadata.CollectionEndTime = Get-Date

        # Export full dataset
        $fullPath = Join-Path $path "TenantDiscovery_Full.json"
        $script:CollectedData | ConvertTo-Json -Depth 20 | Out-File $fullPath -Encoding UTF8

        # Export individual categories
        foreach ($category in $script:CollectedData.Keys) {
            if ($category -ne "Metadata") {
                $categoryPath = Join-Path $path "$category.json"
                $script:CollectedData[$category] | ConvertTo-Json -Depth 15 | Out-File $categoryPath -Encoding UTF8
            }
        }

        Write-Log -Message "Data exported to: $path" -Level Success
        return $fullPath
    }
    catch {
        Write-Log -Message "Failed to export data: $_" -Level Error
        throw
    }
}
#endregion

#region Gotcha Management
function Add-MigrationGotcha {
    <#
    .SYNOPSIS
        Adds a discovered migration gotcha/risk to the collection
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Category,

        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Description,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Critical", "High", "Medium", "Low", "Informational")]
        [string]$Severity,

        [Parameter(Mandatory = $false)]
        [string]$Recommendation,

        [Parameter(Mandatory = $false)]
        [array]$AffectedObjects,

        [Parameter(Mandatory = $false)]
        [int]$AffectedCount = 0,

        [Parameter(Mandatory = $false)]
        [string]$MigrationPhase,

        [Parameter(Mandatory = $false)]
        [hashtable]$AdditionalData
    )

    $gotcha = [PSCustomObject]@{
        Id              = [Guid]::NewGuid().ToString()
        Category        = $Category
        Title           = $Title
        Description     = $Description
        Severity        = $Severity
        Recommendation  = $Recommendation
        AffectedObjects = $AffectedObjects
        AffectedCount   = if ($AffectedObjects) { $AffectedObjects.Count } else { $AffectedCount }
        MigrationPhase  = $MigrationPhase
        AdditionalData  = $AdditionalData
        DiscoveredAt    = Get-Date
    }

    [void]$script:DiscoveredGotchas.Add($gotcha)

    $logLevel = switch ($Severity) {
        "Critical" { "Error" }
        "High"     { "Warning" }
        default    { "Info" }
    }

    Write-Log -Message "GOTCHA [$Severity]: $Title - $Description" -Level $logLevel

    return $gotcha
}

function Get-MigrationGotchas {
    <#
    .SYNOPSIS
        Retrieves all discovered gotchas, optionally filtered
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Category,

        [Parameter(Mandatory = $false)]
        [string]$Severity,

        [Parameter(Mandatory = $false)]
        [switch]$SortBySeverity
    )

    $results = $script:DiscoveredGotchas

    if ($Category) {
        $results = $results | Where-Object { $_.Category -eq $Category }
    }

    if ($Severity) {
        $results = $results | Where-Object { $_.Severity -eq $Severity }
    }

    if ($SortBySeverity) {
        $severityOrder = @{
            "Critical"      = 1
            "High"          = 2
            "Medium"        = 3
            "Low"           = 4
            "Informational" = 5
        }
        $results = $results | Sort-Object { $severityOrder[$_.Severity] }
    }

    return $results
}

function Get-GotchaSummary {
    <#
    .SYNOPSIS
        Returns a summary of discovered gotchas by severity
    #>
    return @{
        Total         = $script:DiscoveredGotchas.Count
        Critical      = ($script:DiscoveredGotchas | Where-Object { $_.Severity -eq "Critical" }).Count
        High          = ($script:DiscoveredGotchas | Where-Object { $_.Severity -eq "High" }).Count
        Medium        = ($script:DiscoveredGotchas | Where-Object { $_.Severity -eq "Medium" }).Count
        Low           = ($script:DiscoveredGotchas | Where-Object { $_.Severity -eq "Low" }).Count
        Informational = ($script:DiscoveredGotchas | Where-Object { $_.Severity -eq "Informational" }).Count
        ByCategory    = $script:DiscoveredGotchas | Group-Object Category | ForEach-Object {
            @{
                Category = $_.Name
                Count    = $_.Count
            }
        }
    }
}
#endregion

#region Module Connection Helpers
function Test-TDModuleAvailability {
    <#
    .SYNOPSIS
        Tests if required PowerShell modules are available
    .NOTES
        Prefixed with TD to avoid conflicts with other modules
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ModuleNames
    )

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

function Install-RequiredModules {
    <#
    .SYNOPSIS
        Installs required PowerShell modules if not present
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    $requiredModules = @(
        @{ Name = "Microsoft.Graph"; MinVersion = "2.0.0" }
        @{ Name = "ExchangeOnlineManagement"; MinVersion = "3.0.0" }
        @{ Name = "PnP.PowerShell"; MinVersion = "2.0.0" }
        @{ Name = "MicrosoftTeams"; MinVersion = "5.0.0" }
        @{ Name = "Az.Accounts"; MinVersion = "2.0.0" }
        @{ Name = "Az.Resources"; MinVersion = "6.0.0" }
    )

    $installed = @()
    $failed = @()

    foreach ($module in $requiredModules) {
        try {
            $existing = Get-Module -ListAvailable -Name $module.Name |
                Where-Object { $_.Version -ge [Version]$module.MinVersion }

            if (-not $existing -or $Force) {
                Write-Log -Message "Installing module: $($module.Name)" -Level Info
                Install-Module -Name $module.Name -MinimumVersion $module.MinVersion -Force -AllowClobber -Scope CurrentUser
                $installed += $module.Name
            }
            else {
                Write-Log -Message "Module already installed: $($module.Name) v$($existing[0].Version)" -Level Debug
            }
        }
        catch {
            Write-Log -Message "Failed to install module $($module.Name): $_" -Level Error
            $failed += $module.Name
        }
    }

    return @{
        Installed = $installed
        Failed    = $failed
    }
}

function Connect-M365Services {
    <#
    .SYNOPSIS
        Connects to required M365 services using interactive or app registration authentication
    .PARAMETER AuthConfig
        Authentication configuration object containing TenantId, ClientId, ClientSecret for service principal auth
    .PARAMETER Interactive
        Use interactive authentication (browser sign-in)
    .PARAMETER Services
        Array of services to connect to
    .PARAMETER SharePointAdminUrl
        SharePoint admin URL for PnP connection
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [PSCredential]$Credential,

        [Parameter(Mandatory = $false)]
        [switch]$Interactive,

        [Parameter(Mandatory = $false)]
        [PSObject]$AuthConfig,

        [Parameter(Mandatory = $false)]
        [string]$SharePointAdminUrl,

        [Parameter(Mandatory = $false)]
        [string[]]$Services = @("Graph", "ExchangeOnline", "SharePoint", "Teams")
    )

    $connections = @{}
    $useServicePrincipal = $AuthConfig -and $AuthConfig.Method -eq "ServicePrincipal"

    # Extract auth details if using service principal
    $tenantId = $null
    $clientId = $null
    $clientSecret = $null
    $certThumbprint = $null
    $certPath = $null

    if ($useServicePrincipal) {
        $tenantId = $AuthConfig.TenantId
        $clientId = $AuthConfig.ClientId
        $clientSecret = $AuthConfig.ClientSecret
        $certThumbprint = $AuthConfig.CertificateThumbprint
        $certPath = $AuthConfig.CertificatePath

        if (-not $tenantId -or -not $clientId) {
            throw "ServicePrincipal authentication requires TenantId and ClientId"
        }

        # Need either client secret OR certificate
        if (-not $clientSecret -and -not $certThumbprint -and -not $certPath) {
            throw "ServicePrincipal authentication requires ClientSecret, CertificateThumbprint, or CertificatePath"
        }

        $authMethod = if ($certThumbprint) { "Certificate (Thumbprint)" } elseif ($certPath) { "Certificate (File)" } else { "ClientSecret" }
        Write-Log -Message "Using app registration authentication (ClientId: $clientId, Method: $authMethod)" -Level Info
    }

    foreach ($service in $Services) {
        try {
            switch ($service) {
                "Graph" {
                    Write-Log -Message "Connecting to Microsoft Graph..." -Level Info

                    if ($useServicePrincipal) {
                        # Create credential for service principal
                        $secureSecret = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
                        $clientCredential = New-Object System.Management.Automation.PSCredential($clientId, $secureSecret)

                        $null = Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $clientCredential -NoWelcome
                    }
                    else {
                        $scopes = @(
                            "User.Read.All",
                            "Group.Read.All",
                            "Directory.Read.All",
                            "Device.Read.All",
                            "Application.Read.All",
                            "Policy.Read.All",
                            "RoleManagement.Read.All",
                            "Organization.Read.All"
                        )
                        $null = Connect-MgGraph -Scopes $scopes -NoWelcome
                    }

                    $context = Get-MgContext
                    $connections["Graph"] = @{
                        Connected = $true
                        TenantId  = $context.TenantId
                        Account   = $context.Account
                        AuthType  = if ($useServicePrincipal) { "ServicePrincipal" } else { "Interactive" }
                    }

                    # Store tenant info
                    $script:CollectedData.Metadata.TenantId = $context.TenantId

                    # Store tenant ID for other services if using interactive
                    if (-not $useServicePrincipal) {
                        $tenantId = $context.TenantId
                    }
                }

                "ExchangeOnline" {
                    Write-Log -Message "Connecting to Exchange Online..." -Level Info

                    if ($useServicePrincipal) {
                        # Exchange Online app auth requires certificate, not client secret
                        # Fall back to interactive for now with a warning
                        Write-Log -Message "Exchange Online requires certificate auth for app-only. Using interactive." -Level Warning
                        $null = Connect-ExchangeOnline -ShowBanner:$false
                    }
                    else {
                        $null = Connect-ExchangeOnline -ShowBanner:$false
                    }
                    $connections["ExchangeOnline"] = @{ Connected = $true }
                }

                "SharePoint" {
                    Write-Log -Message "Connecting to SharePoint Online via PnP.PowerShell..." -Level Info

                    if (-not $SharePointAdminUrl) {
                        Write-Log -Message "SharePoint admin URL not provided. Skipping SharePoint connection." -Level Warning
                        $connections["SharePoint"] = @{ Connected = $false; Note = "Requires admin URL" }
                    }
                    else {
                        if ($useServicePrincipal) {
                            Write-Log -Message "Connecting to SharePoint with ClientId: $clientId" -Level Debug

                            if ($certThumbprint) {
                                # Azure AD app with certificate thumbprint (cert must be in cert store)
                                Write-Log -Message "Using certificate thumbprint authentication" -Level Info
                                $null = Connect-PnPOnline -Url $SharePointAdminUrl -ClientId $clientId -Tenant $tenantId -Thumbprint $certThumbprint
                            }
                            elseif ($certPath) {
                                # Azure AD app with certificate file
                                Write-Log -Message "Using certificate file authentication" -Level Info
                                $null = Connect-PnPOnline -Url $SharePointAdminUrl -ClientId $clientId -Tenant $tenantId -CertificatePath $certPath
                            }
                            else {
                                # Client secret (for ACS/SharePoint Add-in apps only)
                                Write-Log -Message "Using client secret authentication (ACS mode)" -Level Info
                                $null = Connect-PnPOnline -Url $SharePointAdminUrl -ClientId $clientId -ClientSecret $clientSecret
                            }
                        }
                        else {
                            # Use interactive browser login
                            Write-Log -Message "Opening browser for SharePoint login..." -Level Info
                            $null = Connect-PnPOnline -Url $SharePointAdminUrl -Interactive
                        }
                        $connections["SharePoint"] = @{
                            Connected = $true
                            AdminUrl  = $SharePointAdminUrl
                            AuthType  = if ($useServicePrincipal) { "ServicePrincipal" } else { "Interactive" }
                        }
                    }
                }

                "Teams" {
                    Write-Log -Message "Connecting to Microsoft Teams..." -Level Info
                    # Teams Admin cmdlets (Get-CsTeams*) require a user with Teams Admin role
                    # App registrations cannot be assigned this role, so always use interactive auth
                    Write-Log -Message "Teams Admin cmdlets require interactive sign-in (using your admin credentials)" -Level Info
                    $null = Connect-MicrosoftTeams
                    $connections["Teams"] = @{ Connected = $true }
                }

                "Security" {
                    Write-Log -Message "Connecting to Security & Compliance..." -Level Info
                    # Security & Compliance Center doesn't support client secret auth
                    $null = Connect-IPPSSession -ShowBanner:$false
                    $connections["Security"] = @{ Connected = $true }
                }
            }

            Write-Log -Message "Connected to $service successfully" -Level Success
        }
        catch {
            Write-Log -Message "Failed to connect to $service : $_" -Level Error
            $connections[$service] = @{
                Connected = $false
                Error     = $_.Exception.Message
            }
        }
    }

    return $connections
}

function Disconnect-M365Services {
    <#
    .SYNOPSIS
        Disconnects from all M365 services
    #>
    [CmdletBinding()]
    param()

    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue

        Write-Log -Message "Disconnected from all services" -Level Info
    }
    catch {
        Write-Log -Message "Error during disconnect: $_" -Level Warning
    }
}
#endregion

#region Utility Functions
function ConvertTo-FriendlySize {
    <#
    .SYNOPSIS
        Converts bytes to human-readable size
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [long]$Bytes
    )

    switch ($Bytes) {
        { $_ -ge 1PB } { return "{0:N2} PB" -f ($_ / 1PB) }
        { $_ -ge 1TB } { return "{0:N2} TB" -f ($_ / 1TB) }
        { $_ -ge 1GB } { return "{0:N2} GB" -f ($_ / 1GB) }
        { $_ -ge 1MB } { return "{0:N2} MB" -f ($_ / 1MB) }
        { $_ -ge 1KB } { return "{0:N2} KB" -f ($_ / 1KB) }
        default { return "{0} Bytes" -f $_ }
    }
}

function Get-ObjectHash {
    <#
    .SYNOPSIS
        Generates a hash for an object for comparison purposes
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $InputObject
    )

    $json = $InputObject | ConvertTo-Json -Depth 10 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    return [Convert]::ToBase64String($hash)
}

function Test-ValidEmail {
    <#
    .SYNOPSIS
        Validates an email address format
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email
    )

    return $Email -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
}

function Get-UPNFromEmail {
    <#
    .SYNOPSIS
        Extracts UPN prefix from email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email
    )

    if ($Email -match '^([^@]+)@') {
        return $Matches[1]
    }
    return $null
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    # Initialization
    'Initialize-TenantDiscovery',
    'Get-TenantDiscoveryConfig',
    'Set-TenantDiscoveryConfig',
    'Get-DefaultConfiguration',

    # Logging
    'Write-Log',
    'Write-Progress-Status',

    # Data Collection
    'Add-CollectedData',
    'Get-CollectedData',
    'Export-CollectedData',

    # Gotcha Management
    'Add-MigrationGotcha',
    'Get-MigrationGotchas',
    'Get-GotchaSummary',

    # Connection Management
    'Test-TDModuleAvailability',
    'Install-RequiredModules',
    'Connect-M365Services',
    'Disconnect-M365Services',

    # Utilities
    'ConvertTo-FriendlySize',
    'Get-ObjectHash',
    'Test-ValidEmail',
    'Get-UPNFromEmail'
)
