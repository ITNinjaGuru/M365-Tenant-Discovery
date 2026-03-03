#Requires -Version 7.0
<#
.SYNOPSIS
    Hybrid Identity Assessment Module
.DESCRIPTION
    Assesses hybrid identity configuration including Azure AD Connect,
    ADFS, Pass-through Authentication, and federation settings.
    Identifies migration gotchas related to hybrid identity.
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

#region Azure AD Connect Configuration
function Get-AADConnectConfiguration {
    <#
    .SYNOPSIS
        Collects Azure AD Connect configuration from Graph API
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Azure AD Connect configuration..." -Level Info

    try {
        # Get organization info for sync status
        $org = Get-MgOrganization

        # Get directory sync status
        $uri = "https://graph.microsoft.com/beta/directory/onPremisesSynchronization"
        $syncConfig = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue

        # Get service principals for AAD Connect
        $aadConnectApps = @(
            "cb1056e2-e479-49de-ae31-7812af012ed8"  # Azure AD Connect
            "6eb59a73-39b2-4c23-a70f-e2e3ce8965b1"  # Directory Sync
        )

        $aadConnectSP = foreach ($appId in $aadConnectApps) {
            $uri = "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$appId'"
            $result = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue
            $result.value
        }

        $config = @{
            OnPremisesSyncEnabled    = $org.OnPremisesSyncEnabled
            DirSyncEnabled           = $org.OnPremisesSyncEnabled
            LastDirSyncTime          = $org.OnPremisesLastSyncDateTime
            LastPasswordSyncTime     = $org.OnPremisesLastPasswordSyncDateTime
            SyncConfiguration        = if ($syncConfig) {
                @{
                    Configuration = $syncConfig.configuration
                    Features     = $syncConfig.features
                }
            } else { $null }
            AADConnectInstalled      = ($aadConnectSP | Measure-Object).Count -gt 0
        }

        $analysis = @{
            HybridIdentityEnabled = $org.OnPremisesSyncEnabled
            SyncHealthy           = if ($org.OnPremisesLastSyncDateTime) {
                ([datetime]$org.OnPremisesLastSyncDateTime) -gt (Get-Date).AddHours(-3)
            } else { $false }
            LastSyncAgeHours      = if ($org.OnPremisesLastSyncDateTime) {
                [math]::Round(((Get-Date) - [datetime]$org.OnPremisesLastSyncDateTime).TotalHours, 2)
            } else { $null }
        }

        # Detect gotchas
        if ($org.OnPremisesSyncEnabled) {
            Add-MigrationGotcha -Category "HybridIdentity" `
                -Title "Azure AD Connect Synchronization Active" `
                -Description "Directory sync is enabled from on-premises AD. This is a critical consideration for migration." `
                -Severity "Critical" `
                -Recommendation "Plan AAD Connect migration approach: new install in target, or staged cutover. Document all sync settings and filtering rules." `
                -MigrationPhase "Pre-Migration"
        }

        # Check for sync issues
        if ($org.OnPremisesLastSyncDateTime) {
            $syncAge = (Get-Date) - [datetime]$org.OnPremisesLastSyncDateTime
            if ($syncAge.TotalHours -gt 3) {
                Add-MigrationGotcha -Category "HybridIdentity" `
                    -Title "Directory Sync Potentially Stale" `
                    -Description "Last sync was $([math]::Round($syncAge.TotalHours, 2)) hours ago. This may indicate sync health issues." `
                    -Severity "High" `
                    -Recommendation "Investigate sync health before migration. Ensure sync is healthy for clean baseline." `
                    -MigrationPhase "Pre-Migration"
            }
        }

        $result = @{
            Configuration = $config
            Analysis      = $analysis
        }

        Add-CollectedData -Category "HybridIdentity" -SubCategory "AADConnect" -Data $result
        Write-Log -Message "Azure AD Connect configuration collected" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect AAD Connect config: $_" -Level Error
        throw
    }
}
#endregion

#region Federation Configuration
function Get-FederationConfiguration {
    <#
    .SYNOPSIS
        Collects federation (ADFS) configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting federation configuration..." -Level Info

    try {
        # Get domain federation status
        $domains = Get-MgDomain

        $federatedDomains = $domains | Where-Object { $_.AuthenticationType -eq "Federated" }

        $federationDetails = foreach ($domain in $federatedDomains) {
            # Get federation configuration for domain
            $uri = "https://graph.microsoft.com/v1.0/domains/$($domain.Id)/federationConfiguration"
            $fedConfig = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue

            @{
                DomainName              = $domain.Id
                IsVerified              = $domain.IsVerified
                IsDefault               = $domain.IsDefault
                AuthenticationType      = $domain.AuthenticationType
                FederationConfiguration = if ($fedConfig.value) {
                    @{
                        DisplayName           = $fedConfig.value[0].displayName
                        IssuerUri             = $fedConfig.value[0].issuerUri
                        MetadataExchangeUri   = $fedConfig.value[0].metadataExchangeUri
                        SigningCertificate    = $fedConfig.value[0].signingCertificate
                        PassiveSignInUri      = $fedConfig.value[0].passiveSignInUri
                        SignOutUri            = $fedConfig.value[0].signOutUri
                        PreferredAuthenticationProtocol = $fedConfig.value[0].preferredAuthenticationProtocol
                    }
                } else { $null }
            }
        }

        $managedDomains = $domains | Where-Object { $_.AuthenticationType -eq "Managed" }

        $analysis = @{
            TotalDomains        = $domains.Count
            FederatedDomains    = $federatedDomains.Count
            ManagedDomains      = $managedDomains.Count
            FederationInUse     = $federatedDomains.Count -gt 0
        }

        # Detect gotchas
        if ($federatedDomains.Count -gt 0) {
            Add-MigrationGotcha -Category "HybridIdentity" `
                -Title "Federated Domains Present" `
                -Description "Found $($federatedDomains.Count) federated domain(s). Federation requires careful cutover planning." `
                -Severity "Critical" `
                -Recommendation "Document ADFS configuration. Plan for federation cutover or staged migration. Consider converting to managed authentication." `
                -AffectedObjects @($federatedDomains.Id) `
                -MigrationPhase "Pre-Migration"

            # Check federation certificate expiry
            foreach ($fed in $federationDetails) {
                if ($fed.FederationConfiguration.SigningCertificate) {
                    try {
                        $certBytes = [Convert]::FromBase64String($fed.FederationConfiguration.SigningCertificate)
                        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($certBytes)

                        if ($cert.NotAfter -lt (Get-Date).AddDays(90)) {
                            Add-MigrationGotcha -Category "HybridIdentity" `
                                -Title "Federation Certificate Expiring" `
                                -Description "Federation certificate for $($fed.DomainName) expires on $($cert.NotAfter.ToString('yyyy-MM-dd'))." `
                                -Severity "High" `
                                -Recommendation "Plan certificate renewal or migration before expiry." `
                                -MigrationPhase "Pre-Migration"
                        }
                    }
                    catch { }
                }
            }
        }

        $result = @{
            FederatedDomains = $federationDetails
            ManagedDomains   = @($managedDomains | Select-Object Id, IsDefault, IsVerified)
            Analysis         = $analysis
        }

        Add-CollectedData -Category "HybridIdentity" -SubCategory "Federation" -Data $result
        Write-Log -Message "Federation configuration collected" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect federation config: $_" -Level Error
        throw
    }
}
#endregion

#region Password Hash Sync & PTA
function Get-AuthenticationMethodConfiguration {
    <#
    .SYNOPSIS
        Collects password hash sync and pass-through authentication configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting authentication method configuration..." -Level Info

    try {
        # Get authentication methods policies
        $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy"
        $authMethods = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue

        # Check for Application Proxy connector groups (may not exist if App Proxy not configured)
        $connectorGroups = $null
        $ptaAgents = @()
        try {
            $uri = "https://graph.microsoft.com/beta/onPremisesPublishingProfiles/applicationProxy/connectorGroups"
            $connectorGroups = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop

            if ($connectorGroups.value) {
                foreach ($group in $connectorGroups.value) {
                    $uri = "https://graph.microsoft.com/beta/onPremisesPublishingProfiles/applicationProxy/connectorGroups/$($group.id)/members"
                    $members = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue
                    if ($members.value) {
                        $ptaAgents += $members.value
                    }
                }
            }
        }
        catch {
            # Application Proxy may not be configured - this is expected
            Write-Log -Message "Application Proxy not configured or not accessible" -Level Debug
        }

        $config = @{
            PasswordHashSyncEnabled  = $null  # This would need AAD Connect direct query
            PassThroughAuthEnabled   = ($ptaAgents | Measure-Object).Count -gt 0
            PTAAgents                = @($ptaAgents | ForEach-Object {
                @{
                    Id              = $_.id
                    MachineName     = $_.machineName
                    Status          = $_.status
                    ExternalIp      = $_.externalIp
                }
            })
            ConnectorGroups          = @($connectorGroups.value | ForEach-Object {
                @{
                    Id   = $_.id
                    Name = $_.name
                }
            })
        }

        $analysis = @{
            PTAEnabled     = ($ptaAgents | Measure-Object).Count -gt 0
            PTAAgentCount  = ($ptaAgents | Measure-Object).Count
            ActiveAgents   = ($ptaAgents | Where-Object { $_.status -eq "active" }).Count
        }

        # Detect gotchas
        if ($analysis.PTAEnabled) {
            Add-MigrationGotcha -Category "HybridIdentity" `
                -Title "Pass-Through Authentication Enabled" `
                -Description "Found $($analysis.PTAAgentCount) PTA agent(s). PTA requires agents to be installed and configured for target tenant." `
                -Severity "High" `
                -Recommendation "Plan PTA agent deployment for target tenant. Consider staged rollout approach." `
                -AffectedCount $analysis.PTAAgentCount `
                -MigrationPhase "Pre-Migration"

            # Check for unhealthy agents
            $unhealthyAgents = $ptaAgents | Where-Object { $_.status -ne "active" }
            if ($unhealthyAgents.Count -gt 0) {
                Add-MigrationGotcha -Category "HybridIdentity" `
                    -Title "Unhealthy PTA Agents" `
                    -Description "Found $($unhealthyAgents.Count) PTA agent(s) not in active status." `
                    -Severity "Medium" `
                    -Recommendation "Investigate and remediate unhealthy agents before migration." `
                    -AffectedObjects @($unhealthyAgents.machineName) `
                    -MigrationPhase "Pre-Migration"
            }
        }

        $result = @{
            Configuration = $config
            Analysis      = $analysis
        }

        Add-CollectedData -Category "HybridIdentity" -SubCategory "AuthenticationMethods" -Data $result
        Write-Log -Message "Authentication method configuration collected" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect authentication config: $_" -Level Error
        throw
    }
}
#endregion

#region Sync Object Analysis
function Get-SyncObjectAnalysis {
    <#
    .SYNOPSIS
        Analyzes synced objects for migration considerations
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Analyzing synced objects..." -Level Info

    try {
        # Get synced users count
        $uri = "https://graph.microsoft.com/v1.0/users?`$filter=onPremisesSyncEnabled eq true&`$count=true&`$top=1"
        $syncedUsers = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers @{"ConsistencyLevel"="eventual"}
        $syncedUserCount = $syncedUsers.'@odata.count'

        # Get cloud-only users count
        $uri = "https://graph.microsoft.com/v1.0/users?`$filter=onPremisesSyncEnabled ne true&`$count=true&`$top=1"
        $cloudUsers = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers @{"ConsistencyLevel"="eventual"}
        $cloudUserCount = $cloudUsers.'@odata.count'

        # Get synced groups count
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=onPremisesSyncEnabled eq true&`$count=true&`$top=1"
        $syncedGroups = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers @{"ConsistencyLevel"="eventual"}
        $syncedGroupCount = $syncedGroups.'@odata.count'

        # Sample synced users for ImmutableID analysis
        $uri = "https://graph.microsoft.com/v1.0/users?`$filter=onPremisesSyncEnabled eq true&`$select=id,displayName,userPrincipalName,onPremisesImmutableId,onPremisesDistinguishedName&`$top=100"
        $sampleUsers = Invoke-MgGraphRequest -Method GET -Uri $uri

        # Analyze ImmutableIDs
        $usersWithImmutableId = $sampleUsers.value | Where-Object { $_.onPremisesImmutableId }
        $usersWithDN = $sampleUsers.value | Where-Object { $_.onPremisesDistinguishedName }

        $analysis = @{
            TotalSyncedUsers    = $syncedUserCount
            TotalCloudUsers     = $cloudUserCount
            TotalSyncedGroups   = $syncedGroupCount
            SyncedToCloudRatio  = if ($cloudUserCount -gt 0) {
                [math]::Round($syncedUserCount / $cloudUserCount, 2)
            } else { $null }
            ImmutableIdCoverage = if ($sampleUsers.value.Count -gt 0) {
                [math]::Round(($usersWithImmutableId.Count / $sampleUsers.value.Count) * 100, 2)
            } else { 0 }
        }

        # Detect gotchas
        if ($syncedUserCount -gt 0) {
            Add-MigrationGotcha -Category "HybridIdentity" `
                -Title "Synced User Objects" `
                -Description "Found $syncedUserCount synced user(s). These have on-premises source of authority." `
                -Severity "High" `
                -Recommendation "Plan for ImmutableID preservation or remapping. Document sync scope and OU filtering." `
                -AffectedCount $syncedUserCount `
                -MigrationPhase "Pre-Migration"
        }

        if ($syncedGroupCount -gt 0) {
            Add-MigrationGotcha -Category "HybridIdentity" `
                -Title "Synced Group Objects" `
                -Description "Found $syncedGroupCount synced group(s). These cannot be modified in cloud." `
                -Severity "Medium" `
                -Recommendation "Consider if groups should remain synced or converted to cloud-managed in target." `
                -AffectedCount $syncedGroupCount `
                -MigrationPhase "Pre-Migration"
        }

        if ($cloudUserCount -gt 0 -and $syncedUserCount -gt 0) {
            Add-MigrationGotcha -Category "HybridIdentity" `
                -Title "Mixed Identity Sources" `
                -Description "Environment has both synced ($syncedUserCount) and cloud-only ($cloudUserCount) users. This requires dual migration approach." `
                -Severity "Medium" `
                -Recommendation "Document which users are cloud-only vs synced. Plan appropriate migration path for each." `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            SyncedUserCount  = $syncedUserCount
            CloudUserCount   = $cloudUserCount
            SyncedGroupCount = $syncedGroupCount
            SampleUsers      = @($sampleUsers.value | Select-Object id, displayName, userPrincipalName, onPremisesImmutableId)
            Analysis         = $analysis
        }

        Add-CollectedData -Category "HybridIdentity" -SubCategory "SyncObjects" -Data $result
        Write-Log -Message "Sync object analysis completed" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to analyze sync objects: $_" -Level Error
        throw
    }
}
#endregion

#region Seamless SSO
function Get-SeamlessSSOConfiguration {
    <#
    .SYNOPSIS
        Collects Seamless SSO configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Seamless SSO configuration..." -Level Info

    try {
        # Seamless SSO is detected through AZUREADSSOACC computer account presence
        # We can check for related service principals

        $uri = "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=displayName eq 'Microsoft Azure AD SSO'"
        $ssoSP = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue

        $ssoEnabled = ($ssoSP.value | Measure-Object).Count -gt 0

        $config = @{
            SeamlessSSOEnabled = $ssoEnabled
            ServicePrincipal   = $ssoSP.value
        }

        # Detect gotchas
        if ($ssoEnabled) {
            Add-MigrationGotcha -Category "HybridIdentity" `
                -Title "Seamless SSO Enabled" `
                -Description "Seamless SSO is configured. AZUREADSSOACC computer account exists in on-premises AD." `
                -Severity "Medium" `
                -Recommendation "Plan Seamless SSO setup in target tenant. Requires new Kerberos key rollout." `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Configuration = $config
        }

        Add-CollectedData -Category "HybridIdentity" -SubCategory "SeamlessSSO" -Data $result
        Write-Log -Message "Seamless SSO configuration collected" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Seamless SSO config: $_" -Level Error
        throw
    }
}
#endregion

#region Device Writeback
function Get-DeviceWritebackConfiguration {
    <#
    .SYNOPSIS
        Assesses device writeback and hybrid join configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Assessing device writeback configuration..." -Level Info

    try {
        # Get hybrid joined devices
        $uri = "https://graph.microsoft.com/v1.0/devices?`$filter=trustType eq 'ServerAd'&`$count=true&`$top=1"
        $hybridDevices = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers @{"ConsistencyLevel"="eventual"}
        $hybridDeviceCount = $hybridDevices.'@odata.count'

        # Get Azure AD joined devices
        $uri = "https://graph.microsoft.com/v1.0/devices?`$filter=trustType eq 'AzureAd'&`$count=true&`$top=1"
        $aadDevices = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers @{"ConsistencyLevel"="eventual"}
        $aadDeviceCount = $aadDevices.'@odata.count'

        $config = @{
            HybridJoinedDeviceCount = $hybridDeviceCount
            AzureADJoinedDeviceCount = $aadDeviceCount
            DeviceWritebackLikely   = $hybridDeviceCount -gt 0
        }

        $analysis = @{
            TotalManagedDevices = $hybridDeviceCount + $aadDeviceCount
            HybridDeviceRatio   = if (($hybridDeviceCount + $aadDeviceCount) -gt 0) {
                [math]::Round($hybridDeviceCount / ($hybridDeviceCount + $aadDeviceCount) * 100, 2)
            } else { 0 }
        }

        # Detect gotchas
        if ($hybridDeviceCount -gt 0) {
            Add-MigrationGotcha -Category "HybridIdentity" `
                -Title "Hybrid Azure AD Joined Devices" `
                -Description "Found $hybridDeviceCount hybrid joined device(s). These require careful migration handling." `
                -Severity "Critical" `
                -Recommendation "Devices will need to be unjoined and rejoined to target tenant. Plan for Conditional Access impact. Consider phased device migration." `
                -AffectedCount $hybridDeviceCount `
                -MigrationPhase "Post-Migration"
        }

        $result = @{
            Configuration = $config
            Analysis      = $analysis
        }

        Add-CollectedData -Category "HybridIdentity" -SubCategory "DeviceWriteback" -Data $result
        Write-Log -Message "Device writeback assessment completed" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to assess device writeback: $_" -Level Error
        throw
    }
}
#endregion

#region Application Proxy Configuration
function Get-ApplicationProxyConfiguration {
    <#
    .SYNOPSIS
        Collects Azure AD Application Proxy configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Application Proxy configuration..." -Level Info

    try {
        # Get Application Proxy connector groups
        $uri = "https://graph.microsoft.com/beta/onPremisesPublishingProfiles/applicationProxy/connectorGroups"
        $connectorGroups = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue

        $connectors = @()
        $publishedApps = @()

        if ($connectorGroups.value) {
            foreach ($group in $connectorGroups.value) {
                # Get connectors in group
                $uri = "https://graph.microsoft.com/beta/onPremisesPublishingProfiles/applicationProxy/connectorGroups/$($group.id)/members"
                $members = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue
                if ($members.value) {
                    $connectors += $members.value | ForEach-Object {
                        @{
                            Id              = $_.id
                            MachineName     = $_.machineName
                            ExternalIp      = $_.externalIp
                            Status          = $_.status
                            ConnectorGroupId = $group.id
                            ConnectorGroupName = $group.name
                        }
                    }
                }

                # Get applications using this connector group
                $uri = "https://graph.microsoft.com/beta/onPremisesPublishingProfiles/applicationProxy/connectorGroups/$($group.id)/applications"
                $apps = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue
                if ($apps.value) {
                    $publishedApps += $apps.value | ForEach-Object {
                        @{
                            Id                    = $_.id
                            DisplayName           = $_.displayName
                            ExternalUrl           = $_.onPremisesPublishing.externalUrl
                            InternalUrl           = $_.onPremisesPublishing.internalUrl
                            ExternalAuthenticationType = $_.onPremisesPublishing.externalAuthenticationType
                            PreAuthentication     = $_.onPremisesPublishing.preAuthentication
                            ConnectorGroupId      = $group.id
                        }
                    }
                }
            }
        }

        $config = @{
            ConnectorGroups    = @($connectorGroups.value)
            Connectors         = $connectors
            PublishedApps      = $publishedApps
        }

        $analysis = @{
            AppProxyEnabled       = $connectors.Count -gt 0
            TotalConnectors       = $connectors.Count
            ActiveConnectors      = ($connectors | Where-Object { $_.Status -eq "active" }).Count
            InactiveConnectors    = ($connectors | Where-Object { $_.Status -ne "active" }).Count
            TotalPublishedApps    = $publishedApps.Count
            ConnectorGroupCount   = ($connectorGroups.value | Measure-Object).Count
        }

        # Detect gotchas
        if ($analysis.AppProxyEnabled) {
            Add-MigrationGotcha -Category "HybridIdentity" `
                -Title "Azure AD Application Proxy Configured" `
                -Description "Found $($analysis.TotalConnectors) App Proxy connector(s) publishing $($analysis.TotalPublishedApps) application(s). These require recreation in target tenant." `
                -Severity "Critical" `
                -Recommendation "Document all published applications and connector configurations. Plan for new connector deployment and application republishing in target tenant." `
                -AffectedCount $analysis.TotalPublishedApps `
                -MigrationPhase "Pre-Migration"

            if ($analysis.InactiveConnectors -gt 0) {
                Add-MigrationGotcha -Category "HybridIdentity" `
                    -Title "Inactive Application Proxy Connectors" `
                    -Description "Found $($analysis.InactiveConnectors) inactive connector(s). These may indicate infrastructure issues." `
                    -Severity "Medium" `
                    -Recommendation "Investigate inactive connectors. Ensure healthy connector infrastructure before migration." `
                    -AffectedCount $analysis.InactiveConnectors `
                    -MigrationPhase "Pre-Migration"
            }
        }

        $result = @{
            Configuration = $config
            Analysis      = $analysis
        }

        Add-CollectedData -Category "HybridIdentity" -SubCategory "ApplicationProxy" -Data $result
        Write-Log -Message "Application Proxy configuration collected" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Application Proxy config: $_" -Level Error
        throw
    }
}
#endregion

#region Password Protection & SSPR
function Get-PasswordProtectionConfiguration {
    <#
    .SYNOPSIS
        Collects password protection and SSPR writeback configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting password protection configuration..." -Level Info

    try {
        # Get authorization policy (includes lockout settings)
        $uri = "https://graph.microsoft.com/beta/policies/authorizationPolicy"
        $authzPolicy = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue

        # Get authentication methods policy (SSPR info is here)
        $sspr = $null
        try {
            $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy"
            $sspr = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        }
        catch {
            Write-Log -Message "Could not retrieve authentication methods policy" -Level Debug
        }

        # Get directory settings (includes password protection)
        $settings = $null
        $passwordSettings = $null
        try {
            $uri = "https://graph.microsoft.com/v1.0/groupSettings"
            $settings = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            $passwordSettings = $settings.value | Where-Object { $_.displayName -like "*Password*" }
        }
        catch {
            Write-Log -Message "Could not retrieve directory settings" -Level Debug
        }

        $config = @{
            SSPREnabled                    = $sspr -ne $null -and $sspr.registrationEnforcement.authenticationMethodsRegistrationCampaign.state -eq "enabled"
            SSPRWritebackEnabled           = $null  # Requires AAD Connect direct query
            SSPRRequiredAuthMethods        = $sspr.registrationEnforcement.authenticationMethodsRegistrationCampaign.snoozeDurationInDays
            PasswordProtectionEnabled      = $null
            BannedPasswordsEnabled         = $passwordSettings -ne $null
            LockoutThreshold               = $authzPolicy.lockoutThreshold
            LockoutDurationInSeconds       = $authzPolicy.lockoutDurationInSeconds
        }

        $analysis = @{
            SSPRConfigured          = $config.SSPREnabled
            CustomBannedPasswords   = $passwordSettings -ne $null
        }

        # Detect gotchas
        if ($config.SSPREnabled) {
            Add-MigrationGotcha -Category "HybridIdentity" `
                -Title "Self-Service Password Reset Enabled" `
                -Description "SSPR is configured. If writeback is enabled, it requires AAD Connect configuration for target tenant." `
                -Severity "Medium" `
                -Recommendation "Document SSPR settings and registration data. Plan SSPR reconfiguration in target. Users may need to re-register authentication methods." `
                -MigrationPhase "Post-Migration"
        }

        if ($passwordSettings) {
            Add-MigrationGotcha -Category "HybridIdentity" `
                -Title "Custom Banned Password List" `
                -Description "Custom banned password list is configured. This must be recreated in target tenant." `
                -Severity "Low" `
                -Recommendation "Export banned password list. Configure in target tenant before user migration." `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Configuration = $config
            Analysis      = $analysis
        }

        Add-CollectedData -Category "HybridIdentity" -SubCategory "PasswordProtection" -Data $result
        Write-Log -Message "Password protection configuration collected" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect password protection config: $_" -Level Error
        throw
    }
}
#endregion

#region On-Premises Attribute Analysis
function Get-OnPremisesAttributeAnalysis {
    <#
    .SYNOPSIS
        Analyzes on-premises extended attributes and directory extensions
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Analyzing on-premises attributes and extensions..." -Level Info

    try {
        # Get directory extension properties
        $uri = "https://graph.microsoft.com/v1.0/directoryObjects/getAvailableExtensionProperties"
        $extensions = Invoke-MgGraphRequest -Method POST -Uri $uri -Body '{"isSyncedFromOnPremises": true}' -ContentType "application/json" -ErrorAction SilentlyContinue

        # Sample users to check extension attributes usage
        $uri = "https://graph.microsoft.com/v1.0/users?`$filter=onPremisesSyncEnabled eq true&`$select=id,displayName,onPremisesExtensionAttributes&`$top=100"
        $sampleUsers = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue

        # Analyze extension attribute usage
        $extensionUsage = @{}
        foreach ($attr in 1..15) {
            $attrName = "extensionAttribute$attr"
            $usageCount = ($sampleUsers.value | Where-Object {
                $_.onPremisesExtensionAttributes.$attrName -ne $null -and
                $_.onPremisesExtensionAttributes.$attrName -ne ""
            }).Count
            if ($usageCount -gt 0) {
                $extensionUsage[$attrName] = @{
                    UsageCount = $usageCount
                    SamplePercentage = [math]::Round(($usageCount / $sampleUsers.value.Count) * 100, 2)
                }
            }
        }

        $config = @{
            DirectoryExtensions     = $extensions.value
            ExtensionAttributeUsage = $extensionUsage
            SampleSize              = $sampleUsers.value.Count
        }

        $analysis = @{
            TotalDirectoryExtensions    = ($extensions.value | Measure-Object).Count
            ExtensionAttributesInUse    = $extensionUsage.Count
            CustomExtensionsPresent     = ($extensions.value | Where-Object { $_.name -match "extension_" }).Count -gt 0
        }

        # Detect gotchas
        if ($extensionUsage.Count -gt 0) {
            $usedAttrs = $extensionUsage.Keys -join ", "
            Add-MigrationGotcha -Category "HybridIdentity" `
                -Title "Extension Attributes In Use" `
                -Description "Found $($extensionUsage.Count) extension attribute(s) in use ($usedAttrs). These may be used by applications or CA policies." `
                -Severity "High" `
                -Recommendation "Document extension attribute usage and dependencies. Ensure sync configuration preserves these in target." `
                -AffectedCount $extensionUsage.Count `
                -MigrationPhase "Pre-Migration"
        }

        if ($analysis.CustomExtensionsPresent) {
            Add-MigrationGotcha -Category "HybridIdentity" `
                -Title "Custom Directory Extensions" `
                -Description "Custom directory extensions (via app registrations) are in use. These need to be recreated in target tenant." `
                -Severity "High" `
                -Recommendation "Document all custom extensions and their source applications. Plan for extension recreation in target." `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Configuration = $config
            Analysis      = $analysis
        }

        Add-CollectedData -Category "HybridIdentity" -SubCategory "ExtensionAttributes" -Data $result
        Write-Log -Message "On-premises attribute analysis completed" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to analyze on-premises attributes: $_" -Level Error
        throw
    }
}
#endregion

#region Hybrid Identity Summary
function Get-HybridIdentitySummary {
    <#
    .SYNOPSIS
        Generates a summary of hybrid identity configuration for migration planning
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Generating hybrid identity summary..." -Level Info

    try {
        $collectedData = Get-CollectedData -Category "HybridIdentity"

        $summary = @{
            OverallHybridComplexity = "Unknown"
            KeyFindings = @()
            CriticalDependencies = @()
            MigrationApproach = "Unknown"
        }

        # Determine complexity
        $complexityScore = 0

        if ($collectedData.AADConnect.Configuration.OnPremisesSyncEnabled) {
            $complexityScore += 30
            $summary.KeyFindings += "Directory synchronization is active"
        }

        if ($collectedData.Federation.Analysis.FederatedDomains -gt 0) {
            $complexityScore += 25
            $summary.CriticalDependencies += "ADFS federation requires cutover planning"
        }

        if ($collectedData.AuthenticationMethods.Analysis.PTAEnabled) {
            $complexityScore += 15
            $summary.KeyFindings += "Pass-through authentication in use"
        }

        if ($collectedData.DeviceWriteback.Configuration.HybridJoinedDeviceCount -gt 0) {
            $complexityScore += 20
            $summary.CriticalDependencies += "Hybrid devices require re-registration"
        }

        if ($collectedData.ApplicationProxy.Analysis.AppProxyEnabled) {
            $complexityScore += 10
            $summary.KeyFindings += "Application Proxy publishing applications"
        }

        # Set overall complexity
        $summary.OverallHybridComplexity = switch ($complexityScore) {
            { $_ -ge 60 } { "Very High - Requires extensive planning" }
            { $_ -ge 40 } { "High - Significant hybrid dependencies" }
            { $_ -ge 20 } { "Medium - Some hybrid considerations" }
            default { "Low - Minimal hybrid complexity" }
        }

        # Recommend migration approach
        $summary.MigrationApproach = if ($complexityScore -ge 40) {
            "Staged migration with dedicated hybrid infrastructure workstream"
        } elseif ($complexityScore -ge 20) {
            "Coordinated migration with hybrid cutover planning"
        } else {
            "Standard migration with minimal hybrid considerations"
        }

        $summary.ComplexityScore = $complexityScore

        return $summary
    }
    catch {
        Write-Log -Message "Failed to generate hybrid identity summary: $_" -Level Error
        throw
    }
}
#endregion

#region Main Collection Function
function Invoke-HybridIdentityCollection {
    <#
    .SYNOPSIS
        Runs all Hybrid Identity assessment functions
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [hashtable]$Config
    )

    Write-Log -Message "Starting Hybrid Identity assessment..." -Level Info

    $results = @{
        StartTime = Get-Date
        Collections = @{}
        Errors = @()
    }

    $collections = @(
        @{ Name = "AADConnect"; Function = { Get-AADConnectConfiguration } }
        @{ Name = "Federation"; Function = { Get-FederationConfiguration } }
        @{ Name = "AuthenticationMethods"; Function = { Get-AuthenticationMethodConfiguration } }
        @{ Name = "SyncObjects"; Function = { Get-SyncObjectAnalysis } }
        @{ Name = "SeamlessSSO"; Function = { Get-SeamlessSSOConfiguration } }
        @{ Name = "DeviceWriteback"; Function = { Get-DeviceWritebackConfiguration } }
        @{ Name = "ApplicationProxy"; Function = { Get-ApplicationProxyConfiguration } }
        @{ Name = "PasswordProtection"; Function = { Get-PasswordProtectionConfiguration } }
        @{ Name = "ExtensionAttributes"; Function = { Get-OnPremisesAttributeAnalysis } }
    )

    foreach ($collection in $collections) {
        try {
            Write-Progress -Activity "Hybrid Identity Assessment" -Status "Collecting $($collection.Name)..."
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
        $results.Summary = Get-HybridIdentitySummary
    }
    catch {
        Write-Log -Message "Error generating summary: $_" -Level Warning
    }

    $results.EndTime = Get-Date
    $results.Duration = $results.EndTime - $results.StartTime

    Write-Log -Message "Hybrid Identity assessment completed in $($results.Duration.TotalMinutes.ToString('F2')) minutes" -Level Success

    return $results
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Get-AADConnectConfiguration',
    'Get-FederationConfiguration',
    'Get-AuthenticationMethodConfiguration',
    'Get-SyncObjectAnalysis',
    'Get-SeamlessSSOConfiguration',
    'Get-DeviceWritebackConfiguration',
    'Get-ApplicationProxyConfiguration',
    'Get-PasswordProtectionConfiguration',
    'Get-OnPremisesAttributeAnalysis',
    'Get-HybridIdentitySummary',
    'Invoke-HybridIdentityCollection'
)
