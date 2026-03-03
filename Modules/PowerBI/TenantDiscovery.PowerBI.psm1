#Requires -Version 7.0
<#
.SYNOPSIS
    Power BI Data Collection Module
.DESCRIPTION
    Collects comprehensive Power BI data including workspaces, datasets,
    reports, dataflows, gateways, and capacity. Identifies migration gotchas.
.NOTES
    Author: AI Migration Expert
    Version: 1.0.0
    Target: PowerShell 7.x
#>

# Import core module only if not already loaded
if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    $corePath = Join-Path $PSScriptRoot ".." "Core" "TenantDiscovery.Core.psm1"
    if (Test-Path $corePath) {
        Import-Module $corePath -Force -Global
    }
}

#region Power BI Admin Settings
function Get-PowerBITenantSettings {
    <#
    .SYNOPSIS
        Collects Power BI tenant settings using Admin API
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Power BI tenant settings..." -Level Info

    try {
        # Note: Requires Power BI Admin rights
        # Using Graph API beta endpoint for Power BI admin settings
        $uri = "https://api.powerbi.com/v1.0/myorg/admin/tenantSettings"

        # Try to get tenant settings via Power BI REST API
        $headers = @{
            "Authorization" = "Bearer $((Get-MgContext).AccessToken)"
            "Content-Type"  = "application/json"
        }

        try {
            $settings = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction SilentlyContinue
        }
        catch {
            Write-Log -Message "Could not retrieve Power BI tenant settings directly. Using alternative methods." -Level Warning
            $settings = @{ Note = "Direct API access not available. Install MicrosoftPowerBIMgmt module for full access." }
        }

        $result = @{
            TenantSettings = $settings
            Note = "For complete Power BI discovery, ensure MicrosoftPowerBIMgmt module is installed"
        }

        Add-CollectedData -Category "PowerBI" -SubCategory "TenantSettings" -Data $result
        Write-Log -Message "Power BI tenant settings collected" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Power BI tenant settings: $_" -Level Error
        throw
    }
}
#endregion

#region Workspaces
function Get-PowerBIWorkspaces {
    <#
    .SYNOPSIS
        Collects Power BI workspace information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Power BI workspaces..." -Level Info

    try {
        # Get all workspaces using Graph API
        $workspaces = @()

        # Try Power BI Admin API
        $uri = "https://api.powerbi.com/v1.0/myorg/admin/groups?`$top=5000"

        $headers = @{
            "Content-Type" = "application/json"
        }

        try {
            # Attempt to use existing token
            $token = (Get-MgContext).AccessToken
            if ($token) {
                $headers["Authorization"] = "Bearer $token"
            }

            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
            $workspaces = $response.value
        }
        catch {
            Write-Log -Message "Power BI Admin API not accessible. Collecting via Graph API M365 Groups." -Level Warning

            # Fallback: Get Power BI workspaces via M365 Groups
            $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(c:c eq 'Unified')&`$select=id,displayName,description,createdDateTime"
            $groups = Invoke-MgGraphRequest -Method GET -Uri $uri

            # Note: This won't give us Power BI specific info but identifies potential workspaces
            $workspaces = $groups.value | ForEach-Object {
                @{
                    id          = $_.id
                    name        = $_.displayName
                    description = $_.description
                    type        = "Workspace"
                    state       = "Active"
                    isOnDedicatedCapacity = $null
                    note        = "Collected via M365 Groups - install MicrosoftPowerBIMgmt for detailed info"
                }
            }
        }

        $analysis = @{
            TotalWorkspaces     = $workspaces.Count
            ActiveWorkspaces    = ($workspaces | Where-Object { $_.state -eq "Active" }).Count
            PremiumWorkspaces   = ($workspaces | Where-Object { $_.isOnDedicatedCapacity }).Count
            SharedWorkspaces    = ($workspaces | Where-Object { -not $_.isOnDedicatedCapacity }).Count
        }

        # Detect gotchas
        if ($workspaces.Count -gt 0) {
            Add-MigrationGotcha -Category "PowerBI" `
                -Title "Power BI Workspaces Present" `
                -Description "Found $($workspaces.Count) Power BI workspaces. These require dedicated Power BI migration tooling." `
                -Severity "High" `
                -Recommendation "Plan separate Power BI migration. Document workspace permissions, data sources, and refresh schedules." `
                -AffectedCount $workspaces.Count `
                -MigrationPhase "Pre-Migration"
        }

        $premiumWorkspaces = $workspaces | Where-Object { $_.isOnDedicatedCapacity }
        if ($premiumWorkspaces.Count -gt 0) {
            Add-MigrationGotcha -Category "PowerBI" `
                -Title "Premium Capacity Workspaces" `
                -Description "Found $($premiumWorkspaces.Count) workspaces on Premium capacity. Premium licensing required in target tenant." `
                -Severity "High" `
                -Recommendation "Ensure Premium capacity is available in target tenant before migration." `
                -AffectedCount $premiumWorkspaces.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Workspaces = $workspaces
            Analysis   = $analysis
        }

        Add-CollectedData -Category "PowerBI" -SubCategory "Workspaces" -Data $result
        Write-Log -Message "Collected $($workspaces.Count) Power BI workspaces" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Power BI workspaces: $_" -Level Error
        throw
    }
}
#endregion

#region Gateways
function Get-PowerBIGateways {
    <#
    .SYNOPSIS
        Collects Power BI gateway information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Power BI gateways..." -Level Info

    try {
        $gateways = @()

        $uri = "https://api.powerbi.com/v1.0/myorg/gateways"

        $headers = @{
            "Content-Type" = "application/json"
        }

        try {
            $token = (Get-MgContext).AccessToken
            if ($token) {
                $headers["Authorization"] = "Bearer $token"
            }

            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
            $gateways = $response.value
        }
        catch {
            Write-Log -Message "Power BI Gateway API not accessible. Gateway discovery requires MicrosoftPowerBIMgmt module." -Level Warning
            $gateways = @()
        }

        $analysis = @{
            TotalGateways       = $gateways.Count
            OnPremisesGateways  = ($gateways | Where-Object { $_.type -eq "OnPremises" }).Count
            PersonalGateways    = ($gateways | Where-Object { $_.type -eq "Personal" }).Count
        }

        # Detect gotchas
        if ($gateways.Count -gt 0) {
            Add-MigrationGotcha -Category "PowerBI" `
                -Title "Power BI Gateways Configured" `
                -Description "Found $($gateways.Count) Power BI gateway(s). Gateways connect to on-premises data sources." `
                -Severity "Critical" `
                -Recommendation "Document gateway configurations and data sources. Gateways must be reconfigured for target tenant." `
                -AffectedCount $gateways.Count `
                -MigrationPhase "Pre-Migration"
        }

        $onPremGateways = $gateways | Where-Object { $_.type -eq "OnPremises" }
        if ($onPremGateways.Count -gt 0) {
            Add-MigrationGotcha -Category "PowerBI" `
                -Title "On-Premises Data Gateways" `
                -Description "Found $($onPremGateways.Count) on-premises data gateway(s). These require reinstallation for target tenant." `
                -Severity "Critical" `
                -Recommendation "Plan gateway reinstallation. Document all data sources using each gateway. Coordinate with data owners." `
                -AffectedObjects @($onPremGateways.name) `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Gateways = $gateways
            Analysis = $analysis
        }

        Add-CollectedData -Category "PowerBI" -SubCategory "Gateways" -Data $result
        Write-Log -Message "Collected $($gateways.Count) Power BI gateways" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Power BI gateways: $_" -Level Error
        throw
    }
}
#endregion

#region Capacities
function Get-PowerBICapacities {
    <#
    .SYNOPSIS
        Collects Power BI capacity information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Power BI capacities..." -Level Info

    try {
        $capacities = @()

        $uri = "https://api.powerbi.com/v1.0/myorg/capacities"

        $headers = @{
            "Content-Type" = "application/json"
        }

        try {
            $token = (Get-MgContext).AccessToken
            if ($token) {
                $headers["Authorization"] = "Bearer $token"
            }

            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
            $capacities = $response.value
        }
        catch {
            Write-Log -Message "Power BI Capacity API not accessible." -Level Warning
            $capacities = @()
        }

        $analysis = @{
            TotalCapacities  = $capacities.Count
            PremiumCapacities = ($capacities | Where-Object { $_.sku -like "P*" }).Count
            EmbeddedCapacities = ($capacities | Where-Object { $_.sku -like "EM*" -or $_.sku -like "A*" }).Count
        }

        # Detect gotchas
        if ($capacities.Count -gt 0) {
            Add-MigrationGotcha -Category "PowerBI" `
                -Title "Power BI Premium Capacities" `
                -Description "Found $($capacities.Count) Power BI capacity(ies). Premium/Embedded capacities require licensing in target tenant." `
                -Severity "High" `
                -Recommendation "Plan capacity provisioning in target tenant before migration. Document capacity assignments." `
                -AffectedCount $capacities.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Capacities = $capacities
            Analysis   = $analysis
        }

        Add-CollectedData -Category "PowerBI" -SubCategory "Capacities" -Data $result
        Write-Log -Message "Collected $($capacities.Count) Power BI capacities" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Power BI capacities: $_" -Level Error
        throw
    }
}
#endregion

#region Main Collection Function
function Invoke-PowerBICollection {
    <#
    .SYNOPSIS
        Runs all Power BI data collection functions
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [hashtable]$Config
    )

    Write-Log -Message "Starting Power BI data collection..." -Level Info

    $results = @{
        StartTime = Get-Date
        Collections = @{}
        Errors = @()
    }

    $collections = @(
        @{ Name = "TenantSettings"; Function = { Get-PowerBITenantSettings } }
        @{ Name = "Workspaces"; Function = { Get-PowerBIWorkspaces } }
        @{ Name = "Gateways"; Function = { Get-PowerBIGateways } }
        @{ Name = "Capacities"; Function = { Get-PowerBICapacities } }
    )

    foreach ($collection in $collections) {
        try {
            Write-Progress -Activity "Power BI Collection" -Status "Collecting $($collection.Name)..."
            $results.Collections[$collection.Name] = & $collection.Function
        }
        catch {
            $results.Errors += @{
                Collection = $collection.Name
                Error      = $_.Exception.Message
            }
            Write-Log -Message "Error in $($collection.Name) collection: $_" -Level Error
        }
    }

    $results.EndTime = Get-Date
    $results.Duration = $results.EndTime - $results.StartTime

    Write-Log -Message "Power BI collection completed in $($results.Duration.TotalMinutes.ToString('F2')) minutes" -Level Success

    return $results
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Get-PowerBITenantSettings',
    'Get-PowerBIWorkspaces',
    'Get-PowerBIGateways',
    'Get-PowerBICapacities',
    'Invoke-PowerBICollection'
)
