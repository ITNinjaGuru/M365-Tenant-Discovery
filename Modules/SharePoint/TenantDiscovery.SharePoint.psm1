#Requires -Version 7.0
<#
.SYNOPSIS
    SharePoint Online & OneDrive Data Collection Module
.DESCRIPTION
    Collects comprehensive SharePoint and OneDrive data including sites,
    permissions, sharing settings, and storage configurations.
    Identifies migration gotchas related to SharePoint Online.
    Uses PnP.PowerShell for app registration support with client secret.
.NOTES
    Author: AI Migration Expert
    Version: 1.1.0
    Target: PowerShell 7.x
    Requires: PnP.PowerShell module
#>

# Import core module only if not already loaded
if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    $corePath = Join-Path $PSScriptRoot ".." "Core" "TenantDiscovery.Core.psm1"
    if (Test-Path $corePath) {
        Import-Module $corePath -Force -Global
    }
}

#region Tenant Configuration
function Get-SharePointTenantConfig {
    <#
    .SYNOPSIS
        Collects SharePoint Online tenant configuration using PnP.PowerShell
    .DESCRIPTION
        Uses PnP.PowerShell for better app registration support with client secret
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$AdminUrl
    )

    Write-Log -Message "Collecting SharePoint tenant configuration..." -Level Info

    try {
        # Use existing PnP connection (established in Connect-M365Services)
        # If no connection exists, this will fail with a clear message
        $tenantConfig = Get-PnPTenant

        $config = @{
            StorageQuota                     = $tenantConfig.StorageQuota
            StorageQuotaAllocated            = $tenantConfig.StorageQuotaAllocated
            ResourceQuota                    = $tenantConfig.ResourceQuota
            ResourceQuotaAllocated           = $tenantConfig.ResourceQuotaAllocated
            SharingCapability                = $tenantConfig.SharingCapability.ToString()
            ShowEveryoneClaim               = $tenantConfig.ShowEveryoneClaim
            ShowEveryoneExceptExternalUsersClaim = $tenantConfig.ShowEveryoneExceptExternalUsersClaim
            ExternalServicesEnabled         = $tenantConfig.ExternalServicesEnabled
            NoAccessRedirectUrl             = $tenantConfig.NoAccessRedirectUrl
            SharingAllowedDomainList        = $tenantConfig.SharingAllowedDomainList
            SharingBlockedDomainList        = $tenantConfig.SharingBlockedDomainList
            SharingDomainRestrictionMode    = $tenantConfig.SharingDomainRestrictionMode.ToString()
            OneDriveSharingCapability       = $tenantConfig.OneDriveSharingCapability.ToString()
            DefaultSharingLinkType          = $tenantConfig.DefaultSharingLinkType.ToString()
            DefaultLinkPermission           = $tenantConfig.DefaultLinkPermission.ToString()
            RequireAcceptingAccountMatchInvitedAccount = $tenantConfig.RequireAcceptingAccountMatchInvitedAccount
            RequireAnonymousLinksExpireInDays = $tenantConfig.RequireAnonymousLinksExpireInDays
            FileAnonymousLinkType           = $tenantConfig.FileAnonymousLinkType.ToString()
            FolderAnonymousLinkType         = $tenantConfig.FolderAnonymousLinkType.ToString()
            NotifyOwnersWhenItemsReshared   = $tenantConfig.NotifyOwnersWhenItemsReshared
            NotifyOwnersWhenInvitationsAccepted = $tenantConfig.NotifyOwnersWhenInvitationsAccepted
            NotificationsInOneDriveForBusinessEnabled = $tenantConfig.NotificationsInOneDriveForBusinessEnabled
            NotificationsInSharePointEnabled = $tenantConfig.NotificationsInSharePointEnabled
            ConditionalAccessPolicy         = $tenantConfig.ConditionalAccessPolicy.ToString()
            DisallowInfectedFileDownload    = $tenantConfig.DisallowInfectedFileDownload
            AllowDownloadingNonWebViewableFiles = $tenantConfig.AllowDownloadingNonWebViewableFiles
            CommentsOnSitePagesDisabled     = $tenantConfig.CommentsOnSitePagesDisabled
            SocialBarOnSitePagesDisabled    = $tenantConfig.SocialBarOnSitePagesDisabled
            OrphanedPersonalSitesRetentionPeriod = $tenantConfig.OrphanedPersonalSitesRetentionPeriod
            DisabledWebPartIds              = $tenantConfig.DisabledWebPartIds
            LegacyAuthProtocolsEnabled      = $tenantConfig.LegacyAuthProtocolsEnabled
            EnableGuestSignInAcceleration   = $tenantConfig.EnableGuestSignInAcceleration
            BccExternalSharingInvitations   = $tenantConfig.BccExternalSharingInvitations
            BccExternalSharingInvitationsList = $tenantConfig.BccExternalSharingInvitationsList
            UsePersistentCookiesForExplorerView = $tenantConfig.UsePersistentCookiesForExplorerView
            UserVoiceForFeedbackEnabled     = $tenantConfig.UserVoiceForFeedbackEnabled
            HideSyncButtonOnTeamSite        = $tenantConfig.HideSyncButtonOnTeamSite
            PermissiveBrowserFileHandlingOverride = $tenantConfig.PermissiveBrowserFileHandlingOverride
            DisabledModernListTemplateIds   = $tenantConfig.DisabledModernListTemplateIds
            IsWBFluidEnabled                = $tenantConfig.IsWBFluidEnabled
            IsCollabMeetingNotesFluidEnabled = $tenantConfig.IsCollabMeetingNotesFluidEnabled
            SpecialCharactersStateInFileFolderNames = $tenantConfig.SpecialCharactersStateInFileFolderNames.ToString()
        }

        # Check for gotchas
        if ($tenantConfig.SharingCapability -eq "ExternalUserAndGuestSharing") {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "External Sharing Fully Enabled" `
                -Description "SharePoint is configured to allow sharing with anyone (anonymous links). This may pose security risks in target tenant." `
                -Severity "Medium" `
                -Recommendation "Review sharing policy. Consider tightening external sharing settings. Document current settings for comparison." `
                -MigrationPhase "Pre-Migration"
        }

        if ($tenantConfig.LegacyAuthProtocolsEnabled) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Legacy Authentication Enabled" `
                -Description "Legacy authentication protocols are enabled for SharePoint. This is a security concern and should be disabled." `
                -Severity "High" `
                -Recommendation "Plan to disable legacy authentication. Identify apps/scripts using legacy auth before disabling." `
                -MigrationPhase "Pre-Migration"
        }

        if ($tenantConfig.SharingBlockedDomainList -or $tenantConfig.SharingAllowedDomainList) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Domain-Restricted Sharing Configured" `
                -Description "Sharing domain restrictions are configured. These policies need recreation in target tenant." `
                -Severity "Low" `
                -Recommendation "Document allowed/blocked domain lists. Recreate in target tenant." `
                -MigrationPhase "Pre-Migration"
        }

        Add-CollectedData -Category "SharePoint" -SubCategory "TenantConfig" -Data $config
        Write-Log -Message "SharePoint tenant configuration collected" -Level Success

        return $config
    }
    catch {
        Write-Log -Message "Failed to collect SharePoint tenant config: $_" -Level Error
        throw
    }
}
#endregion

#region Site Collections
function Get-SharePointSites {
    <#
    .SYNOPSIS
        Collects SharePoint site collection information
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$IncludeOneDrive
    )

    Write-Log -Message "Collecting SharePoint sites..." -Level Info

    try {
        # Get all site collections using PnP.PowerShell
        $sites = Get-PnPTenantSite -IncludeOneDriveSites:$IncludeOneDrive

        # Separate by type (only include Active sites)
        $teamSites = $sites | Where-Object { ($_.Template -like "GROUP*" -or $_.Template -eq "STS#3") -and $_.Status -eq "Active" }
        $commSites = $sites | Where-Object { $_.Template -eq "SITEPAGEPUBLISHING#0" -and $_.Status -eq "Active" }
        $classicSites = $sites | Where-Object { $_.Template -eq "STS#0" -and $_.Status -eq "Active" }
        $oneDriveSites = $sites | Where-Object { $_.Template -eq "SPSPERS#10" -and $_.Status -eq "Active" }
        $sharePointOnlySites = $sites | Where-Object { $_.Template -ne "SPSPERS#10" -and $_.Url -notlike "*-my.sharepoint.com/personal/*" -and $_.Template -notlike "GROUP*" -and $_.Status -eq "Active" }
        $hubSites = $sites | Where-Object { $_.IsHubSite -and $_.Status -eq "Active" }
        $groupConnected = $sites | Where-Object { $_.GroupId -ne [Guid]::Empty -and $_.Status -eq "Active" }

        # Calculate storage
        $totalStorageUsed = ($sites | Measure-Object -Property StorageUsageCurrent -Sum).Sum
        $totalStorageAllocated = ($sites | Measure-Object -Property StorageQuota -Sum).Sum

        $siteDetails = foreach ($site in $sites) {
            @{
                Url                     = $site.Url
                Title                   = $site.Title
                Template                = $site.Template
                StorageUsageMB          = $site.StorageUsageCurrent
                StorageQuotaMB          = $site.StorageQuota
                StorageQuotaWarningMB   = $site.StorageQuotaWarningLevel
                ResourceUsage           = $site.ResourceUsageCurrent
                ResourceQuota           = $site.ResourceQuota
                Owner                   = $site.Owner
                SharingCapability       = $site.SharingCapability.ToString()
                Status                  = $site.Status
                LockState               = $site.LockState
                IsHubSite               = $site.IsHubSite
                HubSiteId               = $site.HubSiteId
                GroupId                 = $site.GroupId
                RelatedGroupId          = $site.RelatedGroupId
                SensitivityLabel        = $site.SensitivityLabel
                ConditionalAccessPolicy = $site.ConditionalAccessPolicy.ToString()
                LastContentModifiedDate = $site.LastContentModifiedDate
                LocaleId                = $site.LocaleId
                DenyAddAndCustomizePages = $site.DenyAddAndCustomizePages.ToString()
                PWAEnabled              = $site.PWAEnabled
                WebsCount               = $site.WebsCount
            }
        }

        # Try to get licensed users from EntraID data for OneDrive licensed count
        $licensedUserPrincipalNames = @()
        try {
            $collectedData = Get-CollectedData -ErrorAction SilentlyContinue
            if ($collectedData -and $collectedData.EntraID -and $collectedData.EntraID.Users -and $collectedData.EntraID.Users.Data) {
                $licensedUserPrincipalNames = @($collectedData.EntraID.Users.Data |
                    Where-Object { $_.AssignedLicenses.Count -gt 0 } |
                    Select-Object -ExpandProperty UserPrincipalName)
            }
        }
        catch {
            Write-Log -Message "Could not retrieve EntraID user data for licensed OneDrive count. Will show all OneDrive sites." -Level Warning
        }

        # Calculate OneDrive sites for licensed users if we have user data
        $oneDriveSitesLicensed = 0
        if ($licensedUserPrincipalNames.Count -gt 0) {
            $oneDriveSitesLicensed = ($oneDriveSites | Where-Object {
                $siteOwner = $_.Owner
                $licensedUserPrincipalNames -contains $siteOwner
            }).Count
        }

        $analysis = @{
            SharePointSites            = $sharePointOnlySites.Count
            OneDriveSites              = $oneDriveSites.Count
            OneDriveSitesLicensedUsers = $oneDriveSitesLicensed
            TeamSites                  = $teamSites.Count
            CommunicationSites         = $commSites.Count
            ClassicSites               = $classicSites.Count
            HubSites                   = $hubSites.Count
            GroupConnectedSites        = $groupConnected.Count
            TotalStorageUsedGB         = [math]::Round(($sites | Where-Object { $_.Status -eq "Active" } | Measure-Object -Property StorageUsageCurrent -Sum).Sum / 1024, 2)
            TotalStorageAllocatedGB    = [math]::Round(($sites | Where-Object { $_.Status -eq "Active" } | Measure-Object -Property StorageQuota -Sum).Sum / 1024, 2)
            LockedSites                = ($sites | Where-Object { $_.LockState -ne "Unlock" -and $_.Status -eq "Active" }).Count
        }

        # Detect gotchas

        # Large sites
        $largeSites = $sites | Where-Object { $_.StorageUsageCurrent -gt 100000 } # > 100GB
        if ($largeSites.Count -gt 0) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Large SharePoint Sites" `
                -Description "Found $($largeSites.Count) sites larger than 100GB. These require extended migration windows and may need incremental migration." `
                -Severity "High" `
                -Recommendation "Plan for incremental migration. Consider pre-stage approach. May need to migrate during off-peak hours." `
                -AffectedObjects @($largeSites.Url | Select-Object -First 10) `
                -AffectedCount $largeSites.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Hub sites
        if ($hubSites.Count -gt 0) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Hub Sites Configured" `
                -Description "Found $($hubSites.Count) hub sites. Hub site associations will need recreation in target tenant." `
                -Severity "Medium" `
                -Recommendation "Document hub site hierarchy. Plan for hub recreation before associating sites. Test navigation and search after migration." `
                -AffectedObjects @($hubSites.Url) `
                -MigrationPhase "Post-Migration"
        }

        # Sites with custom sharing settings
        $customSharing = $sites | Where-Object {
            $_.SharingCapability -ne "ExternalUserAndGuestSharing"
        }
        if ($customSharing.Count -gt 0) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Sites with Custom Sharing Settings" `
                -Description "Found $($customSharing.Count) sites with sharing settings different from tenant default." `
                -Severity "Low" `
                -Recommendation "Document per-site sharing configurations. Apply same settings in target tenant post-migration." `
                -AffectedCount $customSharing.Count `
                -MigrationPhase "Post-Migration"
        }

        # Classic sites
        if ($classicSites.Count -gt 0) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Classic SharePoint Sites" `
                -Description "Found $($classicSites.Count) classic SharePoint sites. Consider modernization before or during migration." `
                -Severity "Low" `
                -Recommendation "Evaluate modernization opportunities. Classic sites may have customizations requiring attention." `
                -AffectedObjects @($classicSites.Url | Select-Object -First 10) `
                -AffectedCount $classicSites.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Locked sites
        $lockedSites = $sites | Where-Object { $_.LockState -ne "Unlock" }
        if ($lockedSites.Count -gt 0) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Locked SharePoint Sites" `
                -Description "Found $($lockedSites.Count) locked sites. Locked sites cannot be accessed and may indicate issues." `
                -Severity "Medium" `
                -Recommendation "Review locked sites. Unlock or exclude from migration as appropriate." `
                -AffectedObjects @($lockedSites.Url) `
                -MigrationPhase "Pre-Migration"
        }

        # Sites with sensitivity labels
        $labeledSites = $sites | Where-Object { $_.SensitivityLabel }
        if ($labeledSites.Count -gt 0) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Sites with Sensitivity Labels" `
                -Description "Found $($labeledSites.Count) sites with sensitivity labels applied. Labels need to exist in target tenant." `
                -Severity "High" `
                -Recommendation "Ensure sensitivity labels are created in target tenant before migration. Map label GUIDs between tenants." `
                -AffectedCount $labeledSites.Count `
                -MigrationPhase "Pre-Migration"
        }

        # PWA enabled sites (Project Online)
        $pwaSites = $sites | Where-Object { $_.PWAEnabled }
        if ($pwaSites.Count -gt 0) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Project Web App Sites" `
                -Description "Found $($pwaSites.Count) Project Web App (PWA) enabled sites. Project Online requires special migration handling." `
                -Severity "Critical" `
                -Recommendation "Plan separate Project Online migration. Document project data, timelines, and resources." `
                -AffectedObjects @($pwaSites.Url) `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Sites    = $siteDetails
            Analysis = $analysis
        }

        Add-CollectedData -Category "SharePoint" -SubCategory "Sites" -Data $result
        Write-Log -Message "Collected $($sites.Count) SharePoint sites" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect SharePoint sites: $_" -Level Error
        throw
    }
}
#endregion

#region Hub Sites
function Get-SharePointHubSites {
    <#
    .SYNOPSIS
        Collects hub site configuration details
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting hub site details..." -Level Info

    try {
        $hubSites = Get-PnPHubSite

        $hubDetails = foreach ($hub in $hubSites) {
            # Get associated sites using PnP
            $associatedSites = Get-PnPTenantSite | Where-Object { $_.HubSiteId -eq $hub.SiteId }

            @{
                SiteId            = $hub.SiteId
                SiteUrl           = $hub.SiteUrl
                Title             = $hub.Title
                Description       = $hub.Description
                LogoUrl           = $hub.LogoUrl
                Permissions       = $hub.Permissions
                SiteDesignId      = $hub.SiteDesignId
                RequiresJoinApproval = $hub.RequiresJoinApproval
                AssociatedSiteCount = $associatedSites.Count
                AssociatedSites   = @($associatedSites.Url)
            }
        }

        $result = @{
            HubSites = $hubDetails
            TotalHubs = $hubSites.Count
        }

        Add-CollectedData -Category "SharePoint" -SubCategory "HubSites" -Data $result
        Write-Log -Message "Collected $($hubSites.Count) hub sites" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect hub sites: $_" -Level Error
        throw
    }
}
#endregion

#region Site Designs and Scripts
function Get-SharePointSiteDesigns {
    <#
    .SYNOPSIS
        Collects site design and site script configurations
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting site designs and scripts..." -Level Info

    try {
        $siteDesigns = Get-PnPSiteDesign
        $siteScripts = Get-PnPSiteScript

        $designDetails = foreach ($design in $siteDesigns) {
            @{
                Id              = $design.Id
                Title           = $design.Title
                Description     = $design.Description
                WebTemplate     = $design.WebTemplate
                SiteScriptIds   = $design.SiteScriptIds
                IsDefault       = $design.IsDefault
                PreviewImageUrl = $design.PreviewImageUrl
                Version         = $design.Version
            }
        }

        $scriptDetails = foreach ($script in $siteScripts) {
            @{
                Id          = $script.Id
                Title       = $script.Title
                Description = $script.Description
                Content     = $script.Content
                Version     = $script.Version
            }
        }

        # Detect gotchas
        if ($siteDesigns.Count -gt 0 -or $siteScripts.Count -gt 0) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Custom Site Designs and Scripts" `
                -Description "Found $($siteDesigns.Count) site designs and $($siteScripts.Count) site scripts. These need recreation in target tenant." `
                -Severity "Medium" `
                -Recommendation "Export site design and script definitions. Recreate in target tenant. Scripts may reference IDs that need updating." `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            SiteDesigns = $designDetails
            SiteScripts = $scriptDetails
        }

        Add-CollectedData -Category "SharePoint" -SubCategory "SiteDesigns" -Data $result
        Write-Log -Message "Collected $($siteDesigns.Count) site designs and $($siteScripts.Count) site scripts" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect site designs: $_" -Level Error
        throw
    }
}
#endregion

#region Term Store
function Get-SharePointTermStore {
    <#
    .SYNOPSIS
        Collects managed metadata (term store) information using Graph API
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting term store information..." -Level Info

    try {
        # Use Graph API for term store
        $uri = "https://graph.microsoft.com/v1.0/sites/root/termStore"
        $termStore = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue

        if (-not $termStore) {
            Write-Log -Message "Term store not accessible or not configured" -Level Warning
            return @{ Configured = $false }
        }

        # Get term groups
        $uri = "https://graph.microsoft.com/v1.0/sites/root/termStore/groups"
        $termGroups = Invoke-MgGraphRequest -Method GET -Uri $uri

        $groupDetails = foreach ($group in $termGroups.value) {
            # Get term sets in each group
            $setsUri = "https://graph.microsoft.com/v1.0/sites/root/termStore/groups/$($group.id)/sets"
            $termSets = Invoke-MgGraphRequest -Method GET -Uri $setsUri -ErrorAction SilentlyContinue

            @{
                Id          = $group.id
                DisplayName = $group.displayName
                Description = $group.description
                CreatedDateTime = $group.createdDateTime
                TermSetCount = ($termSets.value | Measure-Object).Count
                TermSets    = @($termSets.value | ForEach-Object {
                    @{
                        Id          = $_.id
                        LocalizedNames = $_.localizedNames
                        Description = $_.description
                        CreatedDateTime = $_.createdDateTime
                    }
                })
            }
        }

        $analysis = @{
            TermStoreId     = $termStore.id
            DefaultLanguage = $termStore.defaultLanguageTag
            Languages       = $termStore.languageTags
            GroupCount      = $termGroups.value.Count
            TotalTermSets   = ($groupDetails | ForEach-Object { $_.TermSetCount } | Measure-Object -Sum).Sum
        }

        # Detect gotchas
        if ($termGroups.value.Count -gt 0) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Managed Metadata Term Store Configured" `
                -Description "Found $($termGroups.value.Count) term groups with term sets. Managed metadata requires careful migration." `
                -Severity "High" `
                -Recommendation "Export full term store structure. Plan for term store migration before content migration. Term IDs will change." `
                -AffectedCount ($groupDetails | ForEach-Object { $_.TermSetCount } | Measure-Object -Sum).Sum `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Configured = $true
            TermStore  = $termStore
            TermGroups = $groupDetails
            Analysis   = $analysis
        }

        Add-CollectedData -Category "SharePoint" -SubCategory "TermStore" -Data $result
        Write-Log -Message "Collected term store with $($termGroups.value.Count) groups" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect term store: $_" -Level Error
        return @{ Configured = $false; Error = $_.Exception.Message }
    }
}
#endregion

#region Content Types
function Get-SharePointContentTypes {
    <#
    .SYNOPSIS
        Collects content type information from the content type hub
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ContentTypeHubUrl
    )

    Write-Log -Message "Collecting content type information..." -Level Info

    try {
        # Use Graph API to get site content types
        $uri = "https://graph.microsoft.com/v1.0/sites/root/contentTypes"
        $contentTypes = Invoke-MgGraphRequest -Method GET -Uri $uri

        $ctDetails = foreach ($ct in $contentTypes.value) {
            @{
                Id           = $ct.id
                Name         = $ct.name
                Description  = $ct.description
                Group        = $ct.group
                Hidden       = $ct.hidden
                ReadOnly     = $ct.readOnly
                Sealed       = $ct.sealed
                IsBuiltIn    = $ct.isBuiltIn
                ParentId     = $ct.parentId
            }
        }

        $customContentTypes = $ctDetails | Where-Object { -not $_.IsBuiltIn }

        $analysis = @{
            TotalContentTypes  = $contentTypes.value.Count
            CustomContentTypes = $customContentTypes.Count
            BuiltInContentTypes = ($ctDetails | Where-Object { $_.IsBuiltIn }).Count
        }

        # Detect gotchas
        if ($customContentTypes.Count -gt 0) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Custom Content Types" `
                -Description "Found $($customContentTypes.Count) custom content types. These need recreation in target tenant." `
                -Severity "Medium" `
                -Recommendation "Document custom content type definitions. Consider using PnP provisioning for content type migration." `
                -AffectedCount $customContentTypes.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            ContentTypes = $ctDetails
            CustomTypes  = $customContentTypes
            Analysis     = $analysis
        }

        Add-CollectedData -Category "SharePoint" -SubCategory "ContentTypes" -Data $result
        Write-Log -Message "Collected $($contentTypes.value.Count) content types" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect content types: $_" -Level Error
        throw
    }
}
#endregion

#region OneDrive
function Get-OneDriveUsage {
    <#
    .SYNOPSIS
        Collects OneDrive usage information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting OneDrive usage information..." -Level Info

    try {
        # Get OneDrive sites using PnP.PowerShell
        $oneDriveSites = Get-PnPTenantSite -IncludeOneDriveSites | Where-Object { $_.Url -like "*-my.sharepoint.com/personal/*" }

        $odDetails = foreach ($od in $oneDriveSites) {
            @{
                Url                = $od.Url
                Owner              = $od.Owner
                StorageUsageMB     = $od.StorageUsageCurrent
                StorageQuotaMB     = $od.StorageQuota
                LastContentModified = $od.LastContentModifiedDate
                LockState          = $od.LockState
                Status             = $od.Status
                SharingCapability  = $od.SharingCapability.ToString()
            }
        }

        # Calculate statistics
        $totalStorage = ($oneDriveSites | Measure-Object -Property StorageUsageCurrent -Sum).Sum
        $avgStorage = ($oneDriveSites | Measure-Object -Property StorageUsageCurrent -Average).Average

        # Identify stale OneDrives (no activity in 90 days)
        $staleOneDrives = $oneDriveSites | Where-Object {
            $_.LastContentModifiedDate -and
            ([datetime]$_.LastContentModifiedDate) -lt (Get-Date).AddDays(-90)
        }

        # Large OneDrives
        $largeOneDrives = $oneDriveSites | Where-Object { $_.StorageUsageCurrent -gt 50000 } # > 50GB

        $analysis = @{
            TotalOneDrives     = $oneDriveSites.Count
            TotalStorageGB     = [math]::Round($totalStorage / 1024, 2)
            AverageStorageMB   = [math]::Round($avgStorage, 2)
            LargeOneDrives     = $largeOneDrives.Count
            StaleOneDrives     = $staleOneDrives.Count
        }

        # Detect gotchas
        if ($largeOneDrives.Count -gt 0) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Large OneDrive Accounts" `
                -Description "Found $($largeOneDrives.Count) OneDrive accounts larger than 50GB. These require extended migration time." `
                -Severity "Medium" `
                -Recommendation "Plan for incremental OneDrive migration. Consider pre-staging approach for large accounts." `
                -AffectedCount $largeOneDrives.Count `
                -MigrationPhase "Pre-Migration"
        }

        if ($staleOneDrives.Count -gt 0) {
            Add-MigrationGotcha -Category "SharePoint" `
                -Title "Stale OneDrive Accounts" `
                -Description "Found $($staleOneDrives.Count) OneDrive accounts with no activity in 90+ days. May belong to departed users." `
                -Severity "Low" `
                -Recommendation "Review stale accounts. Consider if migration is needed or if data can be archived." `
                -AffectedCount $staleOneDrives.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            OneDrives = $odDetails
            Analysis  = $analysis
        }

        Add-CollectedData -Category "SharePoint" -SubCategory "OneDrive" -Data $result
        Write-Log -Message "Collected $($oneDriveSites.Count) OneDrive accounts" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect OneDrive usage: $_" -Level Error
        throw
    }
}
#endregion

#region Main Collection Function
function Invoke-SharePointCollection {
    <#
    .SYNOPSIS
        Runs all SharePoint Online data collection functions
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AdminUrl,

        [Parameter(Mandatory = $false)]
        [hashtable]$Config
    )

    Write-Log -Message "Starting SharePoint Online data collection..." -Level Info

    $results = @{
        StartTime = Get-Date
        Collections = @{}
        Errors = @()
    }

    $collections = @(
        @{ Name = "TenantConfig"; Function = { Get-SharePointTenantConfig -AdminUrl $AdminUrl } }
        @{ Name = "Sites"; Function = { Get-SharePointSites -IncludeOneDrive } }
        @{ Name = "HubSites"; Function = { Get-SharePointHubSites } }
        @{ Name = "SiteDesigns"; Function = { Get-SharePointSiteDesigns } }
        @{ Name = "TermStore"; Function = { Get-SharePointTermStore } }
        @{ Name = "ContentTypes"; Function = { Get-SharePointContentTypes } }
        @{ Name = "OneDrive"; Function = { Get-OneDriveUsage } }
    )

    foreach ($collection in $collections) {
        try {
            Write-Progress -Activity "SharePoint Collection" -Status "Collecting $($collection.Name)..."
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

    Write-Log -Message "SharePoint collection completed in $($results.Duration.TotalMinutes.ToString('F2')) minutes" -Level Success

    return $results
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Get-SharePointTenantConfig',
    'Get-SharePointSites',
    'Get-SharePointHubSites',
    'Get-SharePointSiteDesigns',
    'Get-SharePointTermStore',
    'Get-SharePointContentTypes',
    'Get-OneDriveUsage',
    'Invoke-SharePointCollection'
)
