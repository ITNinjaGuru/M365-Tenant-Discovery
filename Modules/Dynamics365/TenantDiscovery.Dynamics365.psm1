#Requires -Version 7.0
<#
.SYNOPSIS
    Dynamics 365 / Power Platform Data Collection Module
.DESCRIPTION
    Collects comprehensive Dynamics 365 and Power Platform data including
    environments, solutions, users, and integrations. Identifies migration gotchas.
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

#region Power Platform Environments
function Get-PowerPlatformEnvironments {
    <#
    .SYNOPSIS
        Collects Power Platform environment information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Power Platform environments..." -Level Info

    try {
        # Use Power Platform Admin API
        $uri = "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2020-10-01"

        $headers = @{
            "Content-Type" = "application/json"
        }

        $environments = @()
        $bapApiAccessible = $false

        # Try to acquire a Power Platform-scoped token.
        # BAP API requires audience 'https://service.powerapps.com/', not Graph.
        # Prefer Az.Accounts if available since it can issue tokens for any Azure resource.
        $ppToken = $null
        if (Get-Command Get-AzAccessToken -ErrorAction SilentlyContinue) {
            try {
                $azToken = Get-AzAccessToken -ResourceUrl "https://service.powerapps.com/" -ErrorAction Stop
                $ppToken = $azToken.Token
                Write-Log -Message "Acquired Power Platform token via Az.Accounts" -Level Info
            }
            catch {
                Write-Log -Message "Az.Accounts available but could not get Power Platform token: $_" -Level Warning
            }
        }

        # Fall back to Graph token (works only if tenant has BAP API delegated access)
        if (-not $ppToken) {
            $ppToken = (Get-MgContext).AccessToken
        }

        if ($ppToken) {
            $headers["Authorization"] = "Bearer $ppToken"
        }

        try {
            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
            $environments = $response.value

            # Handle pagination — BAP API may return nextLink
            while ($response.'nextLink') {
                $response = Invoke-RestMethod -Uri $response.'nextLink' -Headers $headers -Method Get -ErrorAction Stop
                $environments += $response.value
            }

            $bapApiAccessible = $true
            Write-Log -Message "Power Platform Admin API returned $($environments.Count) environment(s)" -Level Info
        }
        catch {
            Write-Log -Message "Power Platform Admin API not accessible (HTTP $($_.Exception.Response.StatusCode)). Environments cannot be enumerated; falling back to license detection." -Level Warning

            # Check licenses so we can add an informative gotcha, but do NOT fabricate
            # a fake environment entry — it causes incorrect counts in the report.
            $skuUri = "https://graph.microsoft.com/v1.0/subscribedSkus"
            try {
                $skus = Invoke-MgGraphRequest -Method GET -Uri $skuUri
                $dynamicsSkus = $skus.value | Where-Object {
                    $_.skuPartNumber -like "DYN365*" -or
                    $_.skuPartNumber -like "*DYNAMICS*" -or
                    $_.skuPartNumber -like "*POWERAPPS*" -or
                    $_.skuPartNumber -like "*FLOW*" -or
                    $_.skuPartNumber -like "*CRM*"
                }
                if ($dynamicsSkus) {
                    $detectedSkuNames = ($dynamicsSkus | Select-Object -ExpandProperty skuPartNumber) -join ", "
                    Write-Log -Message "Dynamics/Power Platform licenses detected but environment count unavailable: $detectedSkuNames" -Level Warning
                    Add-MigrationGotcha -Category "Dynamics365" `
                        -Title "Power Platform Environments Could Not Be Enumerated" `
                        -Description "Dynamics 365 / Power Platform licenses are assigned ($detectedSkuNames) but the Power Platform Admin API was not accessible. Environment count is unknown. To get accurate environment data, ensure the discovery account has Power Platform Admin role and run with the Az.Accounts module installed." `
                        -Severity "High" `
                        -Recommendation "Grant the discovery account Power Platform Administrator role and re-run discovery, or install the Az.Accounts module so a proper Power Platform token can be acquired." `
                        -MigrationPhase "Pre-Migration"
                }
            }
            catch {
                Write-Log -Message "Could not check license SKUs: $_" -Level Warning
            }

            # Return early with empty results — no fake data
            $result = @{
                Environments = @()
                Analysis     = @{
                    TotalEnvironments = 0
                    ProductionEnvs    = 0
                    SandboxEnvs       = 0
                    DeveloperEnvs     = 0
                    TrialEnvs         = 0
                    DefaultEnv        = 0
                    DynamicsEnvs      = 0
                    ApiAccessible     = $false
                }
            }
            Add-CollectedData -Category "Dynamics365" -SubCategory "Environments" -Data $result
            return $result
        }

        $envDetails = foreach ($env in $environments) {
            @{
                Name              = $env.name
                DisplayName       = $env.properties.displayName
                Location          = $env.location
                Type              = $env.properties.environmentSku
                State             = $env.properties.states.management.id
                CreatedTime       = $env.properties.createdTime
                LinkedAppType     = $env.properties.linkedEnvironmentMetadata.type
                LinkedAppId       = $env.properties.linkedEnvironmentMetadata.resourceId
                SecurityGroupId   = $env.properties.linkedEnvironmentMetadata.securityGroupId
                IsDefault         = $env.properties.isDefault
                EnvironmentUrl    = $env.properties.linkedEnvironmentMetadata.instanceUrl
            }
        }

        $analysis = @{
            TotalEnvironments    = $environments.Count
            # "Default" is the default production environment — count it alongside explicit Production
            ProductionEnvs       = ($envDetails | Where-Object { $_.Type -eq "Production" -or $_.Type -eq "Default" }).Count
            SandboxEnvs          = ($envDetails | Where-Object { $_.Type -eq "Sandbox" }).Count
            DeveloperEnvs        = ($envDetails | Where-Object { $_.Type -eq "Developer" }).Count
            TrialEnvs            = ($envDetails | Where-Object { $_.Type -eq "Trial" }).Count
            DefaultEnv           = ($envDetails | Where-Object { $_.IsDefault }).Count
            DynamicsEnvs         = ($envDetails | Where-Object { $_.LinkedAppType -eq "Dynamics365" }).Count
            ApiAccessible        = $bapApiAccessible
        }

        # Detect gotchas
        if ($environments.Count -gt 0) {
            Add-MigrationGotcha -Category "Dynamics365" `
                -Title "Power Platform Environments Present" `
                -Description "Found $($environments.Count) Power Platform environment(s). These require specialized migration planning." `
                -Severity "Critical" `
                -Recommendation "Dynamics 365 and Power Platform require dedicated migration approach. Document all environments, solutions, and customizations." `
                -AffectedCount $environments.Count `
                -MigrationPhase "Pre-Migration"
        }

        $productionEnvs = $envDetails | Where-Object { $_.Type -eq "Production" -or $_.Type -eq "Default" }
        if ($productionEnvs.Count -gt 0) {
            Add-MigrationGotcha -Category "Dynamics365" `
                -Title "Production Dynamics Environments" `
                -Description "Found $($productionEnvs.Count) production environment(s). Production data migration requires careful planning." `
                -Severity "Critical" `
                -Recommendation "Plan comprehensive testing in sandbox environments first. Document all integrations and data flows." `
                -AffectedObjects @($productionEnvs.DisplayName) `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Environments = $envDetails
            Analysis     = $analysis
        }

        Add-CollectedData -Category "Dynamics365" -SubCategory "Environments" -Data $result
        Write-Log -Message "Collected $($environments.Count) Power Platform environments" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Power Platform environments: $_" -Level Error
        throw
    }
}
#endregion

#region Power Apps
function Get-PowerAppsInventory {
    <#
    .SYNOPSIS
        Collects Power Apps information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Power Apps..." -Level Info

    try {
        # Use Power Apps Admin API
        $uri = "https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/apps?api-version=2020-08-01"

        $headers = @{
            "Content-Type" = "application/json"
        }

        $apps = @()

        try {
            $token = (Get-MgContext).AccessToken
            if ($token) {
                $headers["Authorization"] = "Bearer $token"
            }

            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
            $apps = $response.value
        }
        catch {
            Write-Log -Message "Power Apps Admin API not accessible." -Level Warning
            $apps = @()
        }

        $appDetails = foreach ($app in $apps) {
            @{
                Name          = $app.name
                DisplayName   = $app.properties.displayName
                AppType       = $app.properties.appType
                Owner         = $app.properties.owner.displayName
                CreatedTime   = $app.properties.createdTime
                LastModified  = $app.properties.lastModifiedTime
                Environment   = $app.properties.environment.name
                SharedWith    = $app.properties.sharedGroupsCount
                Status        = $app.properties.appPackageDetails.status
            }
        }

        # Categorize apps
        $canvasApps = $appDetails | Where-Object { $_.AppType -eq "CanvasApp" }
        $modelDrivenApps = $appDetails | Where-Object { $_.AppType -eq "ModelDrivenApp" }

        $analysis = @{
            TotalApps        = $apps.Count
            CanvasApps       = $canvasApps.Count
            ModelDrivenApps  = $modelDrivenApps.Count
        }

        # Detect gotchas
        if ($apps.Count -gt 0) {
            Add-MigrationGotcha -Category "Dynamics365" `
                -Title "Power Apps Present" `
                -Description "Found $($apps.Count) Power Apps. Apps require export/import migration with connection reconfiguration." `
                -Severity "High" `
                -Recommendation "Export app packages. Document all connectors and data sources. Plan for connection reconfiguration in target tenant." `
                -AffectedCount $apps.Count `
                -MigrationPhase "Pre-Migration"
        }

        if ($modelDrivenApps.Count -gt 0) {
            Add-MigrationGotcha -Category "Dynamics365" `
                -Title "Model-Driven Apps Present" `
                -Description "Found $($modelDrivenApps.Count) model-driven apps. These are tied to Dataverse and require solution export." `
                -Severity "High" `
                -Recommendation "Model-driven apps must be migrated as part of Dataverse solution migration." `
                -AffectedCount $modelDrivenApps.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Apps     = $appDetails
            Analysis = $analysis
        }

        Add-CollectedData -Category "Dynamics365" -SubCategory "PowerApps" -Data $result
        Write-Log -Message "Collected $($apps.Count) Power Apps" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Power Apps: $_" -Level Error
        throw
    }
}
#endregion

#region Power Automate Flows
function Get-PowerAutomateFlows {
    <#
    .SYNOPSIS
        Collects Power Automate flows information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Power Automate flows..." -Level Info

    try {
        $uri = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/flows?api-version=2016-11-01"

        $headers = @{
            "Content-Type" = "application/json"
        }

        $flows = @()

        try {
            $token = (Get-MgContext).AccessToken
            if ($token) {
                $headers["Authorization"] = "Bearer $token"
            }

            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
            $flows = $response.value
        }
        catch {
            Write-Log -Message "Power Automate Admin API not accessible." -Level Warning
            $flows = @()
        }

        $flowDetails = foreach ($flow in $flows) {
            @{
                Name         = $flow.name
                DisplayName  = $flow.properties.displayName
                State        = $flow.properties.state
                FlowType     = $flow.properties.definitionSummary.flowType
                CreatedTime  = $flow.properties.createdTime
                LastModified = $flow.properties.lastModifiedTime
                Owner        = $flow.properties.creator.userId
                Environment  = $flow.properties.environment.name
                Triggers     = $flow.properties.definitionSummary.triggers
                Actions      = $flow.properties.definitionSummary.actions
            }
        }

        $analysis = @{
            TotalFlows    = $flows.Count
            ActiveFlows   = ($flowDetails | Where-Object { $_.State -eq "Started" }).Count
            StoppedFlows  = ($flowDetails | Where-Object { $_.State -eq "Stopped" }).Count
        }

        # Detect gotchas
        if ($flows.Count -gt 0) {
            Add-MigrationGotcha -Category "Dynamics365" `
                -Title "Power Automate Flows Present" `
                -Description "Found $($flows.Count) Power Automate flows. Flows require export with connection reconfiguration." `
                -Severity "High" `
                -Recommendation "Export flow packages. Document all connections. Plan for connection reconfiguration and testing in target tenant." `
                -AffectedCount $flows.Count `
                -MigrationPhase "Pre-Migration"
        }

        $activeFlows = $flowDetails | Where-Object { $_.State -eq "Started" }
        if ($activeFlows.Count -gt 0) {
            Add-MigrationGotcha -Category "Dynamics365" `
                -Title "Active Power Automate Flows" `
                -Description "Found $($activeFlows.Count) active flows. Active flows may be processing data during migration." `
                -Severity "Medium" `
                -Recommendation "Plan flow cutover timing. Consider pausing flows during critical migration phases." `
                -AffectedCount $activeFlows.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Flows    = $flowDetails
            Analysis = $analysis
        }

        Add-CollectedData -Category "Dynamics365" -SubCategory "PowerAutomate" -Data $result
        Write-Log -Message "Collected $($flows.Count) Power Automate flows" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Power Automate flows: $_" -Level Error
        throw
    }
}
#endregion

#region Connectors
function Get-PowerPlatformConnectors {
    <#
    .SYNOPSIS
        Collects Power Platform custom connectors
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Power Platform connectors..." -Level Info

    try {
        $uri = "https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/connectors?api-version=2020-08-01"

        $headers = @{
            "Content-Type" = "application/json"
        }

        $connectors = @()

        try {
            $token = (Get-MgContext).AccessToken
            if ($token) {
                $headers["Authorization"] = "Bearer $token"
            }

            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
            $connectors = $response.value
        }
        catch {
            Write-Log -Message "Power Platform Connectors API not accessible." -Level Warning
            $connectors = @()
        }

        $customConnectors = $connectors | Where-Object { $_.properties.isCustomApi }

        $analysis = @{
            TotalConnectors  = $connectors.Count
            CustomConnectors = $customConnectors.Count
            StandardConnectors = ($connectors | Where-Object { -not $_.properties.isCustomApi }).Count
        }

        # Detect gotchas
        if ($customConnectors.Count -gt 0) {
            Add-MigrationGotcha -Category "Dynamics365" `
                -Title "Custom Power Platform Connectors" `
                -Description "Found $($customConnectors.Count) custom connector(s). These must be recreated in target tenant." `
                -Severity "High" `
                -Recommendation "Export custom connector definitions. Document API endpoints and authentication. Recreate in target tenant." `
                -AffectedCount $customConnectors.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Connectors       = $connectors
            CustomConnectors = $customConnectors
            Analysis         = $analysis
        }

        Add-CollectedData -Category "Dynamics365" -SubCategory "Connectors" -Data $result
        Write-Log -Message "Collected $($connectors.Count) Power Platform connectors" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Power Platform connectors: $_" -Level Error
        throw
    }
}
#endregion

#region Dynamics 365 Specific
function Get-Dynamics365Users {
    <#
    .SYNOPSIS
        Collects Dynamics 365 licensed users
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Dynamics 365 users..." -Level Info

    try {
        # Get users with Dynamics licenses
        $uri = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,assignedLicenses&`$top=999"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri

        $allUsers = $response.value
        while ($response.'@odata.nextLink') {
            $response = Invoke-MgGraphRequest -Method GET -Uri $response.'@odata.nextLink'
            $allUsers += $response.value
        }

        # Get license SKU info
        $skuUri = "https://graph.microsoft.com/v1.0/subscribedSkus"
        $skus = (Invoke-MgGraphRequest -Method GET -Uri $skuUri).value

        # Filter for Dynamics 365 / Power Platform SKUs
        # Modern D365 licenses use DYN365_ prefix (e.g. DYN365_ENTERPRISE_CUSTOMER_SERVICE,
        # DYN365_ENTERPRISE_SALES, DYN365_FINANCE, DYN365_SUPPLY_CHAIN_MANAGEMENT).
        # Older/alternate SKUs may use DYNAMICS, CRM, or D365 substrings.
        # POWERAPPS_PER_APP and FLOW_PER_FLOW are also Power Platform licenses.
        $dynamicsSkuIds = @(
            $skus | Where-Object {
                $_.skuPartNumber -like "DYN365*" -or
                $_.skuPartNumber -like "*DYNAMICS*" -or
                $_.skuPartNumber -like "*CRM*" -or
                $_.skuPartNumber -like "*D365*" -or
                $_.skuPartNumber -like "POWERAPPS_PER*" -or
                $_.skuPartNumber -like "FLOW_PER*"
            } | Select-Object -ExpandProperty skuId
        )

        $dynamicsUsers = $allUsers | Where-Object {
            $userLicenses = $_.assignedLicenses.skuId
            $dynamicsSkuIds | Where-Object { $_ -in $userLicenses }
        }

        $analysis = @{
            TotalDynamicsUsers = $dynamicsUsers.Count
        }

        # Detect gotchas
        if ($dynamicsUsers.Count -gt 0) {
            Add-MigrationGotcha -Category "Dynamics365" `
                -Title "Dynamics 365 Licensed Users" `
                -Description "Found $($dynamicsUsers.Count) users with Dynamics 365 licenses. User security roles and data access need migration." `
                -Severity "High" `
                -Recommendation "Document user security roles and business units. Plan for security role recreation in target tenant." `
                -AffectedCount $dynamicsUsers.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Users    = @($dynamicsUsers | Select-Object id, displayName, userPrincipalName)
            Analysis = $analysis
        }

        Add-CollectedData -Category "Dynamics365" -SubCategory "Users" -Data $result
        Write-Log -Message "Collected $($dynamicsUsers.Count) Dynamics 365 users" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Dynamics 365 users: $_" -Level Error
        throw
    }
}
#endregion

#region Data Loss Prevention Policies
function Get-PowerPlatformDLPPolicies {
    <#
    .SYNOPSIS
        Collects Power Platform DLP policies
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Power Platform DLP policies..." -Level Info

    try {
        $uri = "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/apiPolicies?api-version=2016-11-01"

        $headers = @{
            "Content-Type" = "application/json"
        }

        $policies = @()

        try {
            $token = (Get-MgContext).AccessToken
            if ($token) {
                $headers["Authorization"] = "Bearer $token"
            }

            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
            $policies = $response.value
        }
        catch {
            Write-Log -Message "Power Platform DLP Policy API not accessible." -Level Warning
            $policies = @()
        }

        $policyDetails = foreach ($policy in $policies) {
            @{
                Name                = $policy.name
                DisplayName         = $policy.properties.displayName
                CreatedTime         = $policy.properties.createdTime
                LastModifiedTime    = $policy.properties.lastModifiedTime
                EnvironmentType     = $policy.properties.environmentType
                DefaultConnectorsClassification = $policy.properties.definition.defaultConnectorsClassification
                BusinessDataGroup   = $policy.properties.definition.businessDataGroup
                NonBusinessDataGroup = $policy.properties.definition.nonBusinessDataGroup
            }
        }

        $analysis = @{
            TotalDLPPolicies     = $policies.Count
            TenantLevelPolicies  = ($policyDetails | Where-Object { $_.EnvironmentType -eq "AllEnvironments" }).Count
            EnvironmentPolicies  = ($policyDetails | Where-Object { $_.EnvironmentType -ne "AllEnvironments" }).Count
        }

        # Detect gotchas
        if ($policies.Count -gt 0) {
            Add-MigrationGotcha -Category "Dynamics365" `
                -Title "Power Platform DLP Policies" `
                -Description "Found $($policies.Count) DLP policies. These must be recreated in target tenant." `
                -Severity "High" `
                -Recommendation "Export DLP policy configurations. Document connector classifications. Recreate policies in target before enabling Power Platform." `
                -AffectedCount $policies.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Policies = $policyDetails
            Analysis = $analysis
        }

        Add-CollectedData -Category "Dynamics365" -SubCategory "DLPPolicies" -Data $result
        Write-Log -Message "Collected $($policies.Count) Power Platform DLP policies" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Power Platform DLP policies: $_" -Level Error
        throw
    }
}
#endregion

#region Dataverse Solutions
function Get-DataverseSolutions {
    <#
    .SYNOPSIS
        Collects Dataverse/CDS solutions information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Dataverse solutions..." -Level Info

    try {
        # This requires Dataverse environment URL - we'll collect from environment info
        $envData = Get-CollectedData -Category "Dynamics365" -SubCategory "Environments"

        $allSolutions = @()

        if ($envData -and $envData.Environments) {
            foreach ($env in $envData.Environments) {
                if ($env.EnvironmentUrl) {
                    try {
                        $token = (Get-MgContext).AccessToken
                        $headers = @{
                            "Authorization" = "Bearer $token"
                            "Content-Type" = "application/json"
                            "OData-MaxVersion" = "4.0"
                            "OData-Version" = "4.0"
                        }

                        $solutionUri = "$($env.EnvironmentUrl)/api/data/v9.2/solutions?`$select=solutionid,uniquename,friendlyname,version,ismanaged,publisherid,createdby,modifiedon&`$expand=publisherid(`$select=friendlyname)"

                        $response = Invoke-RestMethod -Uri $solutionUri -Headers $headers -Method Get -ErrorAction SilentlyContinue

                        if ($response.value) {
                            $allSolutions += $response.value | ForEach-Object {
                                @{
                                    Environment    = $env.DisplayName
                                    SolutionId     = $_.solutionid
                                    UniqueName     = $_.uniquename
                                    FriendlyName   = $_.friendlyname
                                    Version        = $_.version
                                    IsManaged      = $_.ismanaged
                                    Publisher      = $_.publisherid.friendlyname
                                    ModifiedOn     = $_.modifiedon
                                }
                            }
                        }
                    }
                    catch {
                        Write-Log -Message "Could not retrieve solutions from $($env.DisplayName): $_" -Level Warning
                    }
                }
            }
        }

        $customSolutions = $allSolutions | Where-Object { -not $_.IsManaged -and $_.UniqueName -notlike "Default*" -and $_.UniqueName -ne "Active" }
        $managedSolutions = $allSolutions | Where-Object { $_.IsManaged }

        $analysis = @{
            TotalSolutions    = $allSolutions.Count
            CustomSolutions   = $customSolutions.Count
            ManagedSolutions  = $managedSolutions.Count
        }

        # Detect gotchas
        if ($customSolutions.Count -gt 0) {
            Add-MigrationGotcha -Category "Dynamics365" `
                -Title "Custom Dataverse Solutions" `
                -Description "Found $($customSolutions.Count) custom (unmanaged) solution(s). These contain customizations that need export and migration." `
                -Severity "Critical" `
                -Recommendation "Export all custom solutions as managed. Document solution dependencies. Plan for solution import order in target." `
                -AffectedCount $customSolutions.Count `
                -AffectedObjects @($customSolutions.FriendlyName | Select-Object -Unique) `
                -MigrationPhase "Pre-Migration"
        }

        if ($managedSolutions.Count -gt 10) {
            Add-MigrationGotcha -Category "Dynamics365" `
                -Title "Multiple Managed Solutions Deployed" `
                -Description "Found $($managedSolutions.Count) managed solutions. ISV solutions may require vendor coordination for target tenant." `
                -Severity "Medium" `
                -Recommendation "Identify ISV vs Microsoft solutions. Contact vendors for migration guidance. Plan solution import sequence." `
                -AffectedCount $managedSolutions.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Solutions        = $allSolutions
            CustomSolutions  = $customSolutions
            ManagedSolutions = $managedSolutions
            Analysis         = $analysis
        }

        Add-CollectedData -Category "Dynamics365" -SubCategory "Solutions" -Data $result
        Write-Log -Message "Collected $($allSolutions.Count) Dataverse solutions" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Dataverse solutions: $_" -Level Error
        throw
    }
}
#endregion

#region Power Platform Governance
function Get-PowerPlatformGovernance {
    <#
    .SYNOPSIS
        Collects Power Platform governance settings
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Power Platform governance settings..." -Level Info

    try {
        $uri = "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environmentSettings?api-version=2021-04-01"

        $headers = @{
            "Content-Type" = "application/json"
        }

        $settings = $null

        try {
            $token = (Get-MgContext).AccessToken
            if ($token) {
                $headers["Authorization"] = "Bearer $token"
            }

            $settings = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
        }
        catch {
            Write-Log -Message "Power Platform Settings API not accessible." -Level Warning
        }

        $governance = @{
            EnvironmentCreationRestricted = $settings.disableEnvironmentCreationByNonAdminUsers
            TrialEnvironmentCreationRestricted = $settings.disableTrialEnvironmentCreationByNonAdminUsers
            CapacityAllocationEnabled = $settings.enableCapacityAllocation
            AIBuilderEnabled = $settings.enableAIBuilder
            ShareWithEveryoneEnabled = $settings.shareWithEveryoneEnabled
            EnvironmentRoutingEnabled = $settings.environmentRoutingEnabled
        }

        $analysis = @{
            GovernanceConfigured = ($settings -ne $null)
            RestrictionsEnabled = $governance.EnvironmentCreationRestricted
        }

        # Detect gotchas
        if ($settings) {
            Add-MigrationGotcha -Category "Dynamics365" `
                -Title "Power Platform Governance Settings" `
                -Description "Power Platform governance is configured. Settings must be replicated in target tenant." `
                -Severity "Medium" `
                -Recommendation "Document all governance settings. Configure in target tenant before enabling Power Platform for users." `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Settings   = $governance
            Analysis   = $analysis
        }

        Add-CollectedData -Category "Dynamics365" -SubCategory "Governance" -Data $result
        Write-Log -Message "Collected Power Platform governance settings" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Power Platform governance: $_" -Level Error
        throw
    }
}
#endregion

#region Power Platform Summary
function Get-PowerPlatformSummary {
    <#
    .SYNOPSIS
        Generates a summary of Power Platform footprint
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Generating Power Platform summary..." -Level Info

    try {
        $collectedData = Get-CollectedData -Category "Dynamics365"

        $summary = @{
            OverallComplexity = "Unknown"
            KeyFindings = @()
            CriticalDependencies = @()
            MigrationApproach = "Unknown"
        }

        # Calculate complexity
        $complexityScore = 0

        if ($collectedData.Environments.Analysis.TotalEnvironments -gt 0) {
            $complexityScore += 20 * [Math]::Min($collectedData.Environments.Analysis.TotalEnvironments, 5)
            $summary.KeyFindings += "Found $($collectedData.Environments.Analysis.TotalEnvironments) Power Platform environment(s)"
        }

        if ($collectedData.PowerApps.Analysis.TotalApps -gt 0) {
            $complexityScore += 10 * [Math]::Min([Math]::Ceiling($collectedData.PowerApps.Analysis.TotalApps / 10), 5)
            $summary.KeyFindings += "Found $($collectedData.PowerApps.Analysis.TotalApps) Power Apps"
        }

        if ($collectedData.PowerAutomate.Analysis.TotalFlows -gt 0) {
            $complexityScore += 5 * [Math]::Min([Math]::Ceiling($collectedData.PowerAutomate.Analysis.TotalFlows / 20), 5)
            $summary.KeyFindings += "Found $($collectedData.PowerAutomate.Analysis.TotalFlows) Power Automate flows"
        }

        if ($collectedData.Connectors.Analysis.CustomConnectors -gt 0) {
            $complexityScore += 15 * [Math]::Min($collectedData.Connectors.Analysis.CustomConnectors, 3)
            $summary.CriticalDependencies += "Custom connectors require recreation"
        }

        if ($collectedData.Solutions.Analysis.CustomSolutions -gt 0) {
            $complexityScore += 20 * [Math]::Min($collectedData.Solutions.Analysis.CustomSolutions, 5)
            $summary.CriticalDependencies += "Custom Dataverse solutions require migration"
        }

        # Set overall complexity
        $summary.OverallComplexity = switch ($complexityScore) {
            { $_ -ge 80 } { "Very High - Dedicated Power Platform migration project required" }
            { $_ -ge 50 } { "High - Significant Power Platform footprint" }
            { $_ -ge 25 } { "Medium - Moderate Power Platform usage" }
            { $_ -gt 0 }  { "Low - Limited Power Platform adoption" }
            default { "None - No Power Platform detected" }
        }

        # Recommend migration approach
        $summary.MigrationApproach = if ($complexityScore -ge 50) {
            "Phased migration with dedicated Power Platform workstream and solution-by-solution migration"
        } elseif ($complexityScore -ge 25) {
            "Coordinated migration with Power Platform included in main migration waves"
        } elseif ($complexityScore -gt 0) {
            "Lightweight migration with manual recreation of Power Platform assets"
        } else {
            "No Power Platform migration required"
        }

        $summary.ComplexityScore = $complexityScore

        return $summary
    }
    catch {
        Write-Log -Message "Failed to generate Power Platform summary: $_" -Level Error
        throw
    }
}
#endregion

#region Main Collection Function
function Invoke-Dynamics365Collection {
    <#
    .SYNOPSIS
        Runs all Dynamics 365 / Power Platform data collection functions
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [hashtable]$Config
    )

    Write-Log -Message "Starting Dynamics 365 / Power Platform data collection..." -Level Info

    $results = @{
        StartTime = Get-Date
        Collections = @{}
        Errors = @()
    }

    $collections = @(
        @{ Name = "Environments"; Function = { Get-PowerPlatformEnvironments } }
        @{ Name = "PowerApps"; Function = { Get-PowerAppsInventory } }
        @{ Name = "PowerAutomate"; Function = { Get-PowerAutomateFlows } }
        @{ Name = "Connectors"; Function = { Get-PowerPlatformConnectors } }
        @{ Name = "Users"; Function = { Get-Dynamics365Users } }
        @{ Name = "DLPPolicies"; Function = { Get-PowerPlatformDLPPolicies } }
        @{ Name = "Solutions"; Function = { Get-DataverseSolutions } }
        @{ Name = "Governance"; Function = { Get-PowerPlatformGovernance } }
    )

    foreach ($collection in $collections) {
        try {
            Write-Progress -Activity "Dynamics 365 Collection" -Status "Collecting $($collection.Name)..."
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

    # Generate summary
    try {
        $results.Summary = Get-PowerPlatformSummary
    }
    catch {
        Write-Log -Message "Error generating Power Platform summary: $_" -Level Warning
    }

    $results.EndTime = Get-Date
    $results.Duration = $results.EndTime - $results.StartTime

    Write-Log -Message "Dynamics 365 collection completed in $($results.Duration.TotalMinutes.ToString('F2')) minutes" -Level Success

    return $results
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Get-PowerPlatformEnvironments',
    'Get-PowerAppsInventory',
    'Get-PowerAutomateFlows',
    'Get-PowerPlatformConnectors',
    'Get-Dynamics365Users',
    'Get-PowerPlatformDLPPolicies',
    'Get-DataverseSolutions',
    'Get-PowerPlatformGovernance',
    'Get-PowerPlatformSummary',
    'Invoke-Dynamics365Collection'
)
