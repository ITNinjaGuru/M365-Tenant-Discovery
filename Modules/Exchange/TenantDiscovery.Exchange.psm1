#Requires -Version 7.0
<#
.SYNOPSIS
    Exchange Online Data Collection Module
.DESCRIPTION
    Collects comprehensive Exchange Online data including mailboxes, distribution
    lists, public folders, transport rules, connectors, and mail flow configuration.
    Identifies migration gotchas related to Exchange Online.
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

#region Organization Configuration
function Get-ExchangeOrganizationConfig {
    <#
    .SYNOPSIS
        Collects Exchange Online organization configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Exchange organization configuration..." -Level Info

    try {
        $orgConfig = Get-OrganizationConfig
        $acceptedDomains = Get-AcceptedDomain
        $remoteDomains = Get-RemoteDomain

        $config = @{
            OrganizationName        = $orgConfig.Name
            Guid                    = $orgConfig.Guid
            DefaultDomain           = ($acceptedDomains | Where-Object { $_.Default }).DomainName
            AcceptedDomains         = @($acceptedDomains | ForEach-Object {
                @{
                    DomainName       = $_.DomainName
                    DomainType       = $_.DomainType
                    Default          = $_.Default
                    AuthenticationType = $_.AuthenticationType
                }
            })
            RemoteDomains           = @($remoteDomains | ForEach-Object {
                @{
                    Name             = $_.Name
                    DomainName       = $_.DomainName
                    AutoForwardEnabled = $_.AutoForwardEnabled
                    TNEFEnabled      = $_.TNEFEnabled
                    AllowedOOFType   = $_.AllowedOOFType
                }
            })
            HybridConfiguration     = $orgConfig.IsExchangeOnlineOrganization
            OAuth2ClientProfileEnabled = $orgConfig.OAuth2ClientProfileEnabled
            LeanPopoutEnabled       = $orgConfig.LeanPopoutEnabled
            LinkPreviewEnabled      = $orgConfig.LinkPreviewEnabled
            PublicFoldersEnabled    = $orgConfig.PublicFoldersEnabled
            MailTipsAllTipsEnabled  = $orgConfig.MailTipsAllTipsEnabled
            FocusedInboxOn          = $orgConfig.FocusedInboxOn
            AuditDisabled           = $orgConfig.AuditDisabled
        }

        # Check for gotchas
        $federatedDomains = $acceptedDomains | Where-Object { $_.DomainType -eq "InternalRelay" }
        if ($federatedDomains) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Internal Relay Domains Configured" `
                -Description "Found $($federatedDomains.Count) internal relay domain(s). Mail flow may be routing to on-premises Exchange." `
                -Severity "High" `
                -Recommendation "Document mail flow paths. Plan for mail routing changes during migration cutover." `
                -AffectedObjects @($federatedDomains.DomainName) `
                -MigrationPhase "Pre-Migration"
        }

        $autoForwardDomains = $remoteDomains | Where-Object { $_.AutoForwardEnabled }
        if ($autoForwardDomains.Count -gt 1) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Auto-Forward Enabled Remote Domains" `
                -Description "Auto-forwarding is enabled for remote domains. This may allow data exfiltration and needs review." `
                -Severity "Medium" `
                -Recommendation "Review auto-forward settings. Consider security implications for target tenant." `
                -AffectedObjects @($autoForwardDomains.DomainName) `
                -MigrationPhase "Pre-Migration"
        }

        Add-CollectedData -Category "Exchange" -SubCategory "OrganizationConfig" -Data $config
        Write-Log -Message "Exchange organization configuration collected" -Level Success

        return $config
    }
    catch {
        Write-Log -Message "Failed to collect Exchange org config: $_" -Level Error
        throw
    }
}
#endregion

#region Mailbox Collection
function Get-ExchangeMailboxes {
    <#
    .SYNOPSIS
        Collects mailbox information from Exchange Online
    .DESCRIPTION
        Optimized to only collect UserMailbox and SharedMailbox types.
        Statistics are collected only for migration-relevant mailboxes.
        Permissions are collected only for shared mailboxes (where they matter most).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$MaxMailboxes = 50000,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeStatistics = $true,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeResourceMailboxes = $false
    )

    Write-Log -Message "Collecting Exchange mailboxes (optimized for migration scope)..." -Level Info

    try {
        # OPTIMIZATION: Only get UserMailbox and SharedMailbox types - these are what we migrate
        # This dramatically reduces the number of mailboxes and API calls
        Write-Log -Message "  Fetching user mailboxes..." -Level Info
        $userMailboxes = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties `
            DisplayName,UserPrincipalName,PrimarySmtpAddress,EmailAddresses,RecipientTypeDetails,`
            ArchiveStatus,LitigationHoldEnabled,InPlaceHolds,ForwardingAddress,ForwardingSmtpAddress,`
            DeliverToMailboxAndForward,HiddenFromAddressListsEnabled,WhenCreated,ExchangeGuid

        Write-Log -Message "  Found $($userMailboxes.Count) user mailboxes" -Level Info

        Write-Log -Message "  Fetching shared mailboxes..." -Level Info
        $sharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -Properties `
            DisplayName,UserPrincipalName,PrimarySmtpAddress,EmailAddresses,RecipientTypeDetails,`
            ArchiveStatus,LitigationHoldEnabled,InPlaceHolds,ForwardingAddress,ForwardingSmtpAddress,`
            DeliverToMailboxAndForward,HiddenFromAddressListsEnabled,GrantSendOnBehalfTo,WhenCreated,ExchangeGuid

        Write-Log -Message "  Found $($sharedMailboxes.Count) shared mailboxes" -Level Info

        # Combine for analysis
        $mailboxes = @($userMailboxes) + @($sharedMailboxes)

        # Optionally get resource mailboxes (rooms/equipment) - usually not needed for migration planning
        $roomMailboxes = @()
        $equipmentMailboxes = @()
        if ($IncludeResourceMailboxes) {
            Write-Log -Message "  Fetching resource mailboxes..." -Level Info
            $roomMailboxes = @(Get-EXOMailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited -Properties DisplayName,UserPrincipalName,PrimarySmtpAddress)
            $equipmentMailboxes = @(Get-EXOMailbox -RecipientTypeDetails EquipmentMailbox -ResultSize Unlimited -Properties DisplayName,UserPrincipalName,PrimarySmtpAddress)
            Write-Log -Message "  Found $($roomMailboxes.Count) room and $($equipmentMailboxes.Count) equipment mailboxes" -Level Info
        }

        # Get mailbox statistics ONLY for shared mailboxes and a SAMPLE of user mailboxes
        # This is the slow part - optimize by limiting scope
        $mailboxStats = @{}

        if ($IncludeStatistics) {
            # Always get stats for shared mailboxes (important for migration planning)
            Write-Log -Message "  Collecting statistics for shared mailboxes..." -Level Info
            $sharedCount = 0
            foreach ($mbx in $sharedMailboxes) {
                try {
                    $stats = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -ErrorAction SilentlyContinue
                    if ($stats) {
                        $mailboxStats[$mbx.UserPrincipalName] = @{
                            ItemCount          = $stats.ItemCount
                            TotalItemSize      = $stats.TotalItemSize.ToString()
                            DeletedItemCount   = $stats.DeletedItemCount
                            LastLogonTime      = $stats.LastLogonTime
                            LastUserActionTime = $stats.LastUserActionTime
                        }
                    }
                    $sharedCount++
                    if ($sharedCount % 25 -eq 0) {
                        Write-Progress -Activity "Collecting Shared Mailbox Statistics" -Status "$sharedCount of $($sharedMailboxes.Count)" -PercentComplete (($sharedCount / $sharedMailboxes.Count) * 100)
                    }
                }
                catch { }
            }
            Write-Progress -Activity "Collecting Shared Mailbox Statistics" -Completed

            # Sample user mailbox statistics (get 100 or 10%, whichever is smaller)
            $sampleSize = [math]::Min(100, [math]::Ceiling($userMailboxes.Count * 0.1))
            if ($sampleSize -gt 0 -and $userMailboxes.Count -gt 0) {
                Write-Log -Message "  Collecting statistics for $sampleSize sample user mailboxes..." -Level Info
                $sampleMailboxes = $userMailboxes | Get-Random -Count $sampleSize
                $sampleCount = 0
                foreach ($mbx in $sampleMailboxes) {
                    try {
                        $stats = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -ErrorAction SilentlyContinue
                        if ($stats) {
                            $mailboxStats[$mbx.UserPrincipalName] = @{
                                ItemCount          = $stats.ItemCount
                                TotalItemSize      = $stats.TotalItemSize.ToString()
                                DeletedItemCount   = $stats.DeletedItemCount
                                LastLogonTime      = $stats.LastLogonTime
                                LastUserActionTime = $stats.LastUserActionTime
                            }
                        }
                        $sampleCount++
                        if ($sampleCount % 25 -eq 0) {
                            Write-Progress -Activity "Collecting Sample User Statistics" -Status "$sampleCount of $sampleSize" -PercentComplete (($sampleCount / $sampleSize) * 100)
                        }
                    }
                    catch { }
                }
                Write-Progress -Activity "Collecting Sample User Statistics" -Completed
            }
        }

        # Get archive mailboxes
        $archiveEnabled = $mailboxes | Where-Object { $_.ArchiveStatus -eq "Active" }

        # Get litigation hold
        $litigationHold = $mailboxes | Where-Object { $_.LitigationHoldEnabled }

        # Get in-place hold
        $inPlaceHold = $mailboxes | Where-Object { $_.InPlaceHolds.Count -gt 0 }

        # Mailboxes with forwarding
        $forwardingMailboxes = $mailboxes | Where-Object {
            $_.ForwardingAddress -or $_.ForwardingSmtpAddress
        }

        # Collect permissions for ALL mailboxes (user + shared)
        # Important for migration: Full Access delegates, Send-As (executive assistants, team delegates)
        $mailboxesWithFullAccess = @()
        $mailboxesWithSendAs = @()

        # Combine user and shared mailboxes for permission collection
        $allMailboxesForPerms = @($userMailboxes) + @($sharedMailboxes)
        $totalForPerms = $allMailboxesForPerms.Count

        Write-Log -Message "  Collecting Full Access permissions for $totalForPerms mailboxes (user + shared)..." -Level Info
        $permCount = 0
        foreach ($mbx in $allMailboxesForPerms) {
            try {
                $fullAccess = Get-EXOMailboxPermission -Identity $mbx.UserPrincipalName -ErrorAction SilentlyContinue |
                    Where-Object { $_.User -ne "NT AUTHORITY\SELF" -and $_.AccessRights -contains "FullAccess" -and -not $_.Deny }
                if ($fullAccess) {
                    $mailboxesWithFullAccess += @{
                        Mailbox      = $mbx.UserPrincipalName
                        MailboxType  = $mbx.RecipientTypeDetails
                        Delegates    = @($fullAccess.User)
                    }
                }
                $permCount++
                if ($permCount % 50 -eq 0) {
                    Write-Progress -Activity "Collecting Mailbox Permissions (Full Access)" -Status "$permCount of $totalForPerms" -PercentComplete (($permCount / [math]::Max(1, $totalForPerms)) * 100)
                }
            }
            catch { }
        }
        Write-Progress -Activity "Collecting Mailbox Permissions (Full Access)" -Completed

        # Get Send-As permissions for all mailboxes
        Write-Log -Message "  Collecting Send-As permissions for $totalForPerms mailboxes..." -Level Info
        $permCount = 0
        foreach ($mbx in $allMailboxesForPerms) {
            try {
                $sendAs = Get-EXORecipientPermission -Identity $mbx.UserPrincipalName -ErrorAction SilentlyContinue |
                    Where-Object { $_.Trustee -ne "NT AUTHORITY\SELF" -and $_.AccessRights -contains "SendAs" }
                if ($sendAs) {
                    $mailboxesWithSendAs += @{
                        Mailbox      = $mbx.UserPrincipalName
                        MailboxType  = $mbx.RecipientTypeDetails
                        Trustees     = @($sendAs.Trustee)
                    }
                }
                $permCount++
                if ($permCount % 50 -eq 0) {
                    Write-Progress -Activity "Collecting Mailbox Permissions (Send-As)" -Status "$permCount of $totalForPerms" -PercentComplete (($permCount / [math]::Max(1, $totalForPerms)) * 100)
                }
            }
            catch { }
        }
        Write-Progress -Activity "Collecting Mailbox Permissions (Send-As)" -Completed

        $analysis = @{
            TotalMailboxes       = $mailboxes.Count
            UserMailboxes        = $userMailboxes.Count
            SharedMailboxes      = $sharedMailboxes.Count
            RoomMailboxes        = $roomMailboxes.Count
            EquipmentMailboxes   = $equipmentMailboxes.Count
            SchedulingMailboxes  = 0  # Not collected in optimized mode
            ArchiveEnabled       = $archiveEnabled.Count
            LitigationHold       = $litigationHold.Count
            InPlaceHold          = $inPlaceHold.Count
            ForwardingConfigured = $forwardingMailboxes.Count
            WithFullAccess       = $mailboxesWithFullAccess.Count
            WithSendAs           = $mailboxesWithSendAs.Count
            StatisticsSampled    = $mailboxStats.Count
        }

        # Detect gotchas

        # Large mailboxes
        $largeMailboxes = foreach ($mbx in $mailboxes) {
            $stats = $mailboxStats[$mbx.UserPrincipalName]
            if ($stats -and $stats.TotalItemSize) {
                try {
                    $sizeStr = $stats.TotalItemSize -replace '[^\d]', ''
                    $sizeBytes = [long]$sizeStr
                    if ($sizeBytes -gt 50GB) {
                        @{
                            Mailbox = $mbx.UserPrincipalName
                            Size    = $stats.TotalItemSize
                        }
                    }
                }
                catch { }
            }
        }

        if ($largeMailboxes.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Large Mailboxes Detected" `
                -Description "Found $($largeMailboxes.Count) mailboxes larger than 50GB. These will require longer migration windows." `
                -Severity "Medium" `
                -Recommendation "Plan for extended migration time. Consider archive mailbox strategy. May need incremental sync approach." `
                -AffectedObjects @($largeMailboxes.Mailbox | Select-Object -First 20) `
                -AffectedCount $largeMailboxes.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Shared mailboxes with licenses
        $licensedShared = $sharedMailboxes | Where-Object {
            # Check if shared mailbox has any user license assigned
            $_.ArchiveStatus -eq "Active" -or $_.LitigationHoldEnabled
        }

        if ($licensedShared.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Shared Mailboxes Requiring Licenses" `
                -Description "Found $($licensedShared.Count) shared mailboxes with features requiring licenses (archive or litigation hold)." `
                -Severity "Medium" `
                -Recommendation "Plan for license assignment in target tenant. Consider if features are still needed." `
                -AffectedCount $licensedShared.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Mailboxes with litigation hold
        if ($litigationHold.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Mailboxes with Litigation Hold" `
                -Description "Found $($litigationHold.Count) mailboxes with litigation hold enabled. Legal hold must be maintained during migration." `
                -Severity "Critical" `
                -Recommendation "Coordinate with legal. Ensure hold is applied in target before source is removed. Document all hold configurations." `
                -AffectedObjects @($litigationHold.UserPrincipalName | Select-Object -First 20) `
                -AffectedCount $litigationHold.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Complex forwarding
        $externalForwarding = $forwardingMailboxes | Where-Object {
            $_.ForwardingSmtpAddress -and $_.ForwardingSmtpAddress -notmatch $orgConfig.DefaultDomain
        }

        if ($externalForwarding.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "External Mail Forwarding Configured" `
                -Description "Found $($externalForwarding.Count) mailboxes forwarding to external addresses. This may indicate shadow IT or data exfiltration." `
                -Severity "High" `
                -Recommendation "Review external forwarding. Validate business need. May need policy adjustment in target tenant." `
                -AffectedCount $externalForwarding.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Archive mailboxes
        if ($archiveEnabled.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Archive Mailboxes Enabled" `
                -Description "Found $($archiveEnabled.Count) mailboxes with archive enabled. Archive data must be migrated separately." `
                -Severity "Medium" `
                -Recommendation "Plan for archive migration. May require additional migration passes. Verify archive licensing in target." `
                -AffectedCount $archiveEnabled.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Collect detailed mailbox data
        $mailboxDetails = foreach ($mbx in $mailboxes) {
            @{
                DisplayName              = $mbx.DisplayName
                UserPrincipalName        = $mbx.UserPrincipalName
                PrimarySmtpAddress       = $mbx.PrimarySmtpAddress
                EmailAddresses           = $mbx.EmailAddresses
                RecipientTypeDetails     = $mbx.RecipientTypeDetails
                ArchiveStatus            = $mbx.ArchiveStatus
                ArchiveGuid              = $mbx.ArchiveGuid
                LitigationHoldEnabled    = $mbx.LitigationHoldEnabled
                LitigationHoldDate       = $mbx.LitigationHoldDate
                LitigationHoldOwner      = $mbx.LitigationHoldOwner
                InPlaceHolds             = $mbx.InPlaceHolds
                RetentionHoldEnabled     = $mbx.RetentionHoldEnabled
                ForwardingAddress        = $mbx.ForwardingAddress
                ForwardingSmtpAddress    = $mbx.ForwardingSmtpAddress
                DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
                HiddenFromAddressListsEnabled = $mbx.HiddenFromAddressListsEnabled
                IsDirSynced              = $mbx.IsDirSynced
                ExchangeGuid             = $mbx.ExchangeGuid
                MailboxMoveStatus        = $mbx.MailboxMoveStatus
                WhenCreated              = $mbx.WhenCreated
                WhenMailboxCreated       = $mbx.WhenMailboxCreated
                Statistics               = $mailboxStats[$mbx.UserPrincipalName]
            }
        }

        $result = @{
            Mailboxes               = $mailboxDetails
            MailboxesWithFullAccess = $mailboxesWithFullAccess
            MailboxesWithSendAs     = $mailboxesWithSendAs
            Analysis                = $analysis
        }

        Add-CollectedData -Category "Exchange" -SubCategory "Mailboxes" -Data $result

        # Summary log
        Write-Log -Message "Mailbox collection complete:" -Level Success
        Write-Log -Message "  - User mailboxes: $($userMailboxes.Count)" -Level Info
        Write-Log -Message "  - Shared mailboxes: $($sharedMailboxes.Count) (permissions collected)" -Level Info
        Write-Log -Message "  - Statistics sampled: $($mailboxStats.Count) mailboxes" -Level Info
        Write-Log -Message "  - Full Access permissions found: $($mailboxesWithFullAccess.Count)" -Level Info
        Write-Log -Message "  - Send-As permissions found: $($mailboxesWithSendAs.Count)" -Level Info

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect mailboxes: $_" -Level Error
        throw
    }
}
#endregion

#region Distribution Lists and Groups
function Get-ExchangeDistributionLists {
    <#
    .SYNOPSIS
        Collects distribution list information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting distribution lists..." -Level Info

    try {
        $distributionLists = Get-DistributionGroup -ResultSize Unlimited
        $dynamicDLs = Get-DynamicDistributionGroup -ResultSize Unlimited
        $unifiedGroups = Get-UnifiedGroup -ResultSize Unlimited

        # Collect membership for distribution lists
        $dlDetails = foreach ($dl in $distributionLists) {
            $members = Get-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -ResultSize Unlimited

            # Check for external members
            $externalMembers = $members | Where-Object { $_.RecipientType -eq "MailContact" -or $_.RecipientType -eq "MailUser" }

            @{
                DisplayName            = $dl.DisplayName
                PrimarySmtpAddress     = $dl.PrimarySmtpAddress
                EmailAddresses         = $dl.EmailAddresses
                ManagedBy              = $dl.ManagedBy
                MemberCount            = $members.Count
                ExternalMemberCount    = $externalMembers.Count
                MemberJoinRestriction  = $dl.MemberJoinRestriction
                MemberDepartRestriction = $dl.MemberDepartRestriction
                RequireSenderAuthenticationEnabled = $dl.RequireSenderAuthenticationEnabled
                HiddenFromAddressListsEnabled = $dl.HiddenFromAddressListsEnabled
                IsDirSynced            = $dl.IsDirSynced
                ModeratedBy            = $dl.ModeratedBy
                ModerationEnabled      = $dl.ModerationEnabled
                SendModerationNotifications = $dl.SendModerationNotifications
                BypassModerationFromSendersOrMembers = $dl.BypassModerationFromSendersOrMembers
                WhenCreated            = $dl.WhenCreated
                HasExternalMembers     = $externalMembers.Count -gt 0
            }
        }

        # Dynamic distribution lists
        $dynamicDLDetails = foreach ($ddl in $dynamicDLs) {
            @{
                DisplayName            = $ddl.DisplayName
                PrimarySmtpAddress     = $ddl.PrimarySmtpAddress
                RecipientFilter        = $ddl.RecipientFilter
                RecipientFilterType    = $ddl.RecipientFilterType
                IncludedRecipients     = $ddl.IncludedRecipients
                ManagedBy              = $ddl.ManagedBy
                WhenCreated            = $ddl.WhenCreated
            }
        }

        $analysis = @{
            TotalDistributionLists    = $distributionLists.Count
            TotalDynamicDLs           = $dynamicDLs.Count
            TotalUnifiedGroups        = $unifiedGroups.Count
            SyncedDistributionLists   = ($distributionLists | Where-Object { $_.IsDirSynced }).Count
            ModeratedLists            = ($distributionLists | Where-Object { $_.ModerationEnabled }).Count
            ListsWithExternalMembers  = ($dlDetails | Where-Object { $_.HasExternalMembers }).Count
        }

        # Detect gotchas

        # DLs with external members
        $externalMemberDLs = $dlDetails | Where-Object { $_.HasExternalMembers }
        if ($externalMemberDLs.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Distribution Lists with External Members" `
                -Description "Found $($externalMemberDLs.Count) distribution lists containing external members (contacts). External contacts may need recreation." `
                -Severity "Medium" `
                -Recommendation "Document external members. Plan for mail contact creation in target. Verify external addresses are still valid." `
                -AffectedObjects @($externalMemberDLs.DisplayName | Select-Object -First 20) `
                -AffectedCount $externalMemberDLs.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Synced distribution lists
        $syncedDLs = $distributionLists | Where-Object { $_.IsDirSynced }
        if ($syncedDLs.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "On-Premises Synced Distribution Lists" `
                -Description "Found $($syncedDLs.Count) distribution lists synced from on-premises AD. These are managed on-prem and require source of authority consideration." `
                -Severity "High" `
                -Recommendation "Plan for DL migration approach. May need to convert to cloud-managed or migrate on-prem AD groups." `
                -AffectedCount $syncedDLs.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Dynamic DLs with complex filters
        $complexDynamicDLs = $dynamicDLs | Where-Object {
            $_.RecipientFilter -and $_.RecipientFilter.Length -gt 100
        }
        if ($complexDynamicDLs.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Dynamic Distribution Lists with Complex Filters" `
                -Description "Found $($complexDynamicDLs.Count) dynamic DLs with complex recipient filters. Filters may reference attributes that need mapping." `
                -Severity "Medium" `
                -Recommendation "Review and document dynamic DL filters. Test filters in target tenant. Update attribute references as needed." `
                -AffectedObjects @($complexDynamicDLs.DisplayName) `
                -MigrationPhase "Post-Migration"
        }

        # Large distribution lists
        $largeDLs = $dlDetails | Where-Object { $_.MemberCount -gt 1000 }
        if ($largeDLs.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Large Distribution Lists" `
                -Description "Found $($largeDLs.Count) distribution lists with more than 1000 members. Consider M365 Groups for better scalability." `
                -Severity "Low" `
                -Recommendation "Evaluate migration to M365 Groups for large lists. Consider dynamic membership where appropriate." `
                -AffectedObjects @($largeDLs.DisplayName) `
                -AffectedCount $largeDLs.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            DistributionLists    = $dlDetails
            DynamicDLs           = $dynamicDLDetails
            UnifiedGroups        = @($unifiedGroups | Select-Object DisplayName, PrimarySmtpAddress, ManagedBy, GroupMemberCount, WhenCreated)
            Analysis             = $analysis
        }

        Add-CollectedData -Category "Exchange" -SubCategory "DistributionLists" -Data $result
        Write-Log -Message "Collected $($distributionLists.Count) distribution lists" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect distribution lists: $_" -Level Error
        throw
    }
}
#endregion

#region Public Folders
function Get-ExchangePublicFolders {
    <#
    .SYNOPSIS
        Collects public folder information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting public folders..." -Level Info

    try {
        # Check if public folders exist
        $pfMailboxes = Get-Mailbox -PublicFolder -ResultSize Unlimited -ErrorAction SilentlyContinue

        if (-not $pfMailboxes) {
            Write-Log -Message "No public folder mailboxes found" -Level Info
            Add-CollectedData -Category "Exchange" -SubCategory "PublicFolders" -Data @{
                Enabled = $false
                Message = "No public folders configured"
            }
            return @{ Enabled = $false }
        }

        # Get public folders
        $publicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction SilentlyContinue
        $mailEnabledPFs = Get-MailPublicFolder -ResultSize Unlimited -ErrorAction SilentlyContinue

        # Get public folder statistics
        $pfStats = foreach ($pf in $publicFolders | Select-Object -First 500) {
            try {
                $stats = Get-PublicFolderStatistics -Identity $pf.EntryId -ErrorAction SilentlyContinue
                if ($stats) {
                    @{
                        FolderPath  = $pf.FolderPath
                        ItemCount   = $stats.ItemCount
                        TotalSize   = $stats.TotalItemSize.ToString()
                        FolderType  = $pf.FolderType
                    }
                }
            }
            catch { }
        }

        $analysis = @{
            PublicFolderMailboxCount = $pfMailboxes.Count
            TotalPublicFolders       = $publicFolders.Count
            MailEnabledFolders       = $mailEnabledPFs.Count
            TotalItems               = ($pfStats | Measure-Object -Property ItemCount -Sum).Sum
        }

        # Detect gotchas

        if ($publicFolders.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Public Folders Present" `
                -Description "Found $($publicFolders.Count) public folders in $($pfMailboxes.Count) public folder mailbox(es). Public folder migration requires special handling." `
                -Severity "High" `
                -Recommendation "Plan dedicated public folder migration. Consider modernization to M365 Groups/SharePoint. Document folder permissions." `
                -AffectedCount $publicFolders.Count `
                -MigrationPhase "Pre-Migration" `
                -AdditionalData @{
                    MailEnabledCount = $mailEnabledPFs.Count
                    MailboxCount     = $pfMailboxes.Count
                }
        }

        # Mail-enabled public folders
        if ($mailEnabledPFs.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Mail-Enabled Public Folders" `
                -Description "Found $($mailEnabledPFs.Count) mail-enabled public folders. Email addresses need migration and may conflict with other objects." `
                -Severity "Medium" `
                -Recommendation "Document email addresses. Check for conflicts. Plan for address migration." `
                -AffectedObjects @($mailEnabledPFs.PrimarySmtpAddress | Select-Object -First 20) `
                -AffectedCount $mailEnabledPFs.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Large public folder hierarchy
        if ($publicFolders.Count -gt 10000) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Large Public Folder Hierarchy" `
                -Description "Public folder hierarchy contains $($publicFolders.Count) folders. Large hierarchies have performance implications and longer migration times." `
                -Severity "High" `
                -Recommendation "Consider public folder hierarchy restructuring. Plan for batch migration. May need hierarchy freeze during migration." `
                -AffectedCount $publicFolders.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Enabled              = $true
            PublicFolderMailboxes = @($pfMailboxes | Select-Object Name, Alias, PrimarySmtpAddress, TotalItemSize)
            PublicFolders        = $pfStats
            MailEnabledFolders   = @($mailEnabledPFs | Select-Object Name, PrimarySmtpAddress, EmailAddresses)
            Analysis             = $analysis
        }

        Add-CollectedData -Category "Exchange" -SubCategory "PublicFolders" -Data $result
        Write-Log -Message "Collected $($publicFolders.Count) public folders" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect public folders: $_" -Level Error
        throw
    }
}
#endregion

#region Transport Rules and Connectors
function Get-ExchangeTransportConfig {
    <#
    .SYNOPSIS
        Collects transport rules, connectors, and mail flow configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting transport configuration..." -Level Info

    try {
        # Get transport rules
        $transportRules = Get-TransportRule -ResultSize Unlimited

        # Get connectors
        $inboundConnectors = Get-InboundConnector
        $outboundConnectors = Get-OutboundConnector

        # Get mail flow rules details
        $ruleDetails = foreach ($rule in $transportRules) {
            @{
                Name                = $rule.Name
                State               = $rule.State
                Priority            = $rule.Priority
                Mode                = $rule.Mode
                Conditions          = @{
                    From              = $rule.From
                    SentTo            = $rule.SentTo
                    FromMemberOf      = $rule.FromMemberOf
                    SentToMemberOf    = $rule.SentToMemberOf
                    SubjectContainsWords = $rule.SubjectContainsWords
                    SubjectOrBodyContainsWords = $rule.SubjectOrBodyContainsWords
                    FromAddressContainsWords = $rule.FromAddressContainsWords
                    FromScope         = $rule.FromScope
                    SentToScope       = $rule.SentToScope
                    AttachmentSizeOver = $rule.AttachmentSizeOver
                    HasAttachment     = $rule.HasAttachment
                }
                Actions             = @{
                    AddToRecipients   = $rule.AddToRecipients
                    CopyTo            = $rule.CopyTo
                    BlindCopyTo       = $rule.BlindCopyTo
                    RedirectMessageTo = $rule.RedirectMessageTo
                    ModerateMessageByUser = $rule.ModerateMessageByUser
                    RejectMessageReasonText = $rule.RejectMessageReasonText
                    DeleteMessage     = $rule.DeleteMessage
                    Quarantine        = $rule.Quarantine
                    PrependSubject    = $rule.PrependSubject
                    SetHeaderName     = $rule.SetHeaderName
                    SetHeaderValue    = $rule.SetHeaderValue
                    ApplyHtmlDisclaimerText = $rule.ApplyHtmlDisclaimerText
                }
                WhenChanged         = $rule.WhenChanged
            }
        }

        # Connector details
        $inboundDetails = foreach ($conn in $inboundConnectors) {
            @{
                Name                = $conn.Name
                Enabled             = $conn.Enabled
                ConnectorType       = $conn.ConnectorType
                SenderDomains       = $conn.SenderDomains
                SenderIPAddresses   = $conn.SenderIPAddresses
                RequireTls          = $conn.RequireTls
                TreatMessagesAsInternal = $conn.TreatMessagesAsInternal
                CloudServicesMailEnabled = $conn.CloudServicesMailEnabled
            }
        }

        $outboundDetails = foreach ($conn in $outboundConnectors) {
            @{
                Name                = $conn.Name
                Enabled             = $conn.Enabled
                ConnectorType       = $conn.ConnectorType
                RecipientDomains    = $conn.RecipientDomains
                SmartHosts          = $conn.SmartHosts
                UseMXRecord         = $conn.UseMXRecord
                TlsSettings         = $conn.TlsSettings
                CloudServicesMailEnabled = $conn.CloudServicesMailEnabled
            }
        }

        $analysis = @{
            TotalTransportRules   = $transportRules.Count
            EnabledRules          = ($transportRules | Where-Object { $_.State -eq "Enabled" }).Count
            DisabledRules         = ($transportRules | Where-Object { $_.State -eq "Disabled" }).Count
            InboundConnectors     = $inboundConnectors.Count
            OutboundConnectors    = $outboundConnectors.Count
        }

        # Detect gotchas

        # Rules with group references
        $rulesWithGroups = $transportRules | Where-Object {
            $_.FromMemberOf -or $_.SentToMemberOf -or
            $_.ExceptIfFromMemberOf -or $_.ExceptIfSentToMemberOf
        }

        if ($rulesWithGroups.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Transport Rules with Group References" `
                -Description "Found $($rulesWithGroups.Count) transport rules referencing groups. Group identities will change in target tenant." `
                -Severity "High" `
                -Recommendation "Document all group references in transport rules. Plan for rule recreation with updated group references." `
                -AffectedObjects @($rulesWithGroups.Name) `
                -MigrationPhase "Post-Migration"
        }

        # Rules with specific recipients
        $rulesWithRecipients = $transportRules | Where-Object {
            $_.From -or $_.SentTo -or $_.CopyTo -or $_.BlindCopyTo -or $_.RedirectMessageTo
        }

        if ($rulesWithRecipients.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Transport Rules with Specific Recipients" `
                -Description "Found $($rulesWithRecipients.Count) transport rules with specific recipient addresses. Addresses may need updating during migration." `
                -Severity "Medium" `
                -Recommendation "Document recipient addresses in rules. Verify addresses will be valid in target tenant." `
                -AffectedObjects @($rulesWithRecipients.Name) `
                -MigrationPhase "Post-Migration"
        }

        # Connectors
        if ($inboundConnectors.Count -gt 0 -or $outboundConnectors.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Mail Flow Connectors Configured" `
                -Description "Found $($inboundConnectors.Count) inbound and $($outboundConnectors.Count) outbound connectors. These require recreation in target tenant." `
                -Severity "High" `
                -Recommendation "Document connector configurations including certificates and IP addresses. Plan for connector recreation. Consider mail flow impact during cutover." `
                -MigrationPhase "Pre-Migration"
        }

        # Hybrid connectors
        $hybridConnectors = $inboundConnectors + $outboundConnectors | Where-Object {
            $_.Name -like "*Hybrid*" -or $_.CloudServicesMailEnabled
        }

        if ($hybridConnectors.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Hybrid Mail Flow Connectors" `
                -Description "Found hybrid mail flow connectors. These indicate on-premises Exchange integration." `
                -Severity "High" `
                -Recommendation "Plan for hybrid configuration changes. May need connector updates during migration coexistence." `
                -AffectedObjects @($hybridConnectors.Name) `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            TransportRules    = $ruleDetails
            InboundConnectors = $inboundDetails
            OutboundConnectors = $outboundDetails
            Analysis          = $analysis
        }

        Add-CollectedData -Category "Exchange" -SubCategory "TransportConfig" -Data $result
        Write-Log -Message "Collected $($transportRules.Count) transport rules and $($inboundConnectors.Count + $outboundConnectors.Count) connectors" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect transport config: $_" -Level Error
        throw
    }
}
#endregion

#region Email Address Policies
function Get-ExchangeAddressPolicies {
    <#
    .SYNOPSIS
        Collects email address policies and recipient configurations
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting email address policies..." -Level Info

    try {
        # Get mail contacts
        $mailContacts = Get-MailContact -ResultSize Unlimited -ErrorAction SilentlyContinue

        # Get mail users
        $mailUsers = Get-MailUser -ResultSize Unlimited -ErrorAction SilentlyContinue

        $analysis = @{
            MailContacts = $mailContacts.Count
            MailUsers    = $mailUsers.Count
        }

        # Detect gotchas

        # External mail contacts
        if ($mailContacts.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Mail Contacts Present" `
                -Description "Found $($mailContacts.Count) mail contacts. These represent external recipients and need recreation in target tenant." `
                -Severity "Low" `
                -Recommendation "Export mail contact details. Verify external addresses are current. Plan for contact recreation." `
                -AffectedCount $mailContacts.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Mail users (external users with mailbox elsewhere)
        if ($mailUsers.Count -gt 0) {
            # Check for synced mail users
            $syncedMailUsers = $mailUsers | Where-Object { $_.IsDirSynced }

            Add-MigrationGotcha -Category "Exchange" `
                -Title "Mail Users Present" `
                -Description "Found $($mailUsers.Count) mail users ($($syncedMailUsers.Count) synced). These represent users with external mailboxes." `
                -Severity "Medium" `
                -Recommendation "Review mail user purpose. May be remote mailboxes or external users. Plan for migration or recreation." `
                -AffectedCount $mailUsers.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            MailContacts = @($mailContacts | Select-Object DisplayName, PrimarySmtpAddress, ExternalEmailAddress, IsDirSynced)
            MailUsers    = @($mailUsers | Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress, ExternalEmailAddress, IsDirSynced)
            Analysis     = $analysis
        }

        Add-CollectedData -Category "Exchange" -SubCategory "AddressPolicies" -Data $result
        Write-Log -Message "Collected $($mailContacts.Count) mail contacts and $($mailUsers.Count) mail users" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect address policies: $_" -Level Error
        throw
    }
}
#endregion

#region Hybrid Configuration
function Get-ExchangeHybridConfiguration {
    <#
    .SYNOPSIS
        Collects Exchange hybrid configuration details
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Exchange hybrid configuration..." -Level Info

    try {
        $hybridConfig = $null
        $migrationEndpoints = @()
        $orgRelationships = @()
        $intraOrgConnectors = @()

        try {
            $hybridConfig = Get-HybridConfiguration -ErrorAction SilentlyContinue
        }
        catch {
            Write-Log -Message "No hybrid configuration found or not accessible" -Level Warning
        }

        try {
            $migrationEndpoints = Get-MigrationEndpoint -ErrorAction SilentlyContinue
        }
        catch { }

        try {
            $orgRelationships = Get-OrganizationRelationship -ErrorAction SilentlyContinue
        }
        catch { }

        try {
            $intraOrgConnectors = Get-IntraOrganizationConnector -ErrorAction SilentlyContinue
        }
        catch { }

        $config = @{
            HybridEnabled = $hybridConfig -ne $null
            HybridConfiguration = if ($hybridConfig) {
                @{
                    Domains                = $hybridConfig.Domains
                    Features               = $hybridConfig.Features
                    ReceivingTransportServers = $hybridConfig.ReceivingTransportServers
                    SendingTransportServers = $hybridConfig.SendingTransportServers
                    EdgeTransportServers   = $hybridConfig.EdgeTransportServers
                    TlsCertificateName     = $hybridConfig.TlsCertificateName
                    ServiceInstance        = $hybridConfig.ServiceInstance
                }
            } else { $null }
            MigrationEndpoints = @($migrationEndpoints | ForEach-Object {
                @{
                    Identity         = $_.Identity
                    EndpointType     = $_.EndpointType
                    RemoteServer     = $_.RemoteServer
                    MaxConcurrentMigrations = $_.MaxConcurrentMigrations
                    ExchangeServer   = $_.ExchangeServer
                }
            })
            OrganizationRelationships = @($orgRelationships | ForEach-Object {
                @{
                    Name                    = $_.Name
                    DomainNames             = $_.DomainNames
                    FreeBusyAccessEnabled   = $_.FreeBusyAccessEnabled
                    FreeBusyAccessLevel     = $_.FreeBusyAccessLevel
                    MailTipsAccessEnabled   = $_.MailTipsAccessEnabled
                    PhotosEnabled           = $_.PhotosEnabled
                    TargetApplicationUri    = $_.TargetApplicationUri
                    TargetOwaURL            = $_.TargetOwaURL
                    Enabled                 = $_.Enabled
                }
            })
            IntraOrgConnectors = @($intraOrgConnectors | ForEach-Object {
                @{
                    Name                    = $_.Name
                    TargetAddressDomains    = $_.TargetAddressDomains
                    DiscoveryEndpoint       = $_.DiscoveryEndpoint
                    Enabled                 = $_.Enabled
                }
            })
        }

        $analysis = @{
            HybridEnabled            = $config.HybridEnabled
            MigrationEndpointCount   = $migrationEndpoints.Count
            OrgRelationshipCount     = $orgRelationships.Count
            IntraOrgConnectorCount   = $intraOrgConnectors.Count
        }

        # Detect gotchas
        if ($config.HybridEnabled) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Exchange Hybrid Configuration Active" `
                -Description "Exchange hybrid is configured. This requires careful decommissioning planning during migration." `
                -Severity "Critical" `
                -Recommendation "Document all hybrid features in use. Plan for hybrid removal or reconfiguration to target tenant. Coordinate with on-premises Exchange team." `
                -MigrationPhase "Pre-Migration"
        }

        if ($migrationEndpoints.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Migration Endpoints Configured" `
                -Description "Found $($migrationEndpoints.Count) migration endpoint(s). These may be in use for ongoing migrations." `
                -Severity "Medium" `
                -Recommendation "Document migration endpoints. May need to wait for in-progress migrations to complete." `
                -AffectedCount $migrationEndpoints.Count `
                -MigrationPhase "Pre-Migration"
        }

        if ($orgRelationships.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Organization Relationships Exist" `
                -Description "Found $($orgRelationships.Count) organization relationship(s). Free/busy sharing with other organizations." `
                -Severity "Medium" `
                -Recommendation "Document org relationships. Partner organizations may need to update their configurations." `
                -AffectedCount $orgRelationships.Count `
                -MigrationPhase "Post-Migration"
        }

        $result = @{
            Configuration = $config
            Analysis      = $analysis
        }

        Add-CollectedData -Category "Exchange" -SubCategory "HybridConfig" -Data $result
        Write-Log -Message "Exchange hybrid configuration collected" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect hybrid configuration: $_" -Level Error
        throw
    }
}
#endregion

#region Journaling and Audit
function Get-ExchangeJournalingConfig {
    <#
    .SYNOPSIS
        Collects Exchange journaling and compliance configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Exchange journaling configuration..." -Level Info

    try {
        $journalRules = @()
        $auditConfig = $null

        try {
            $journalRules = Get-JournalRule -ErrorAction SilentlyContinue
        }
        catch { }

        try {
            $auditConfig = Get-AdminAuditLogConfig -ErrorAction SilentlyContinue
        }
        catch { }

        $config = @{
            JournalRules = @($journalRules | ForEach-Object {
                @{
                    Name          = $_.Name
                    Recipient     = $_.Recipient
                    JournalEmailAddress = $_.JournalEmailAddress
                    Scope         = $_.Scope
                    Enabled       = $_.Enabled
                }
            })
            AdminAuditLogConfig = if ($auditConfig) {
                @{
                    UnifiedAuditLogIngestionEnabled = $auditConfig.UnifiedAuditLogIngestionEnabled
                    AdminAuditLogEnabled    = $auditConfig.AdminAuditLogEnabled
                    AdminAuditLogCmdlets    = $auditConfig.AdminAuditLogCmdlets
                    AdminAuditLogParameters = $auditConfig.AdminAuditLogParameters
                    AdminAuditLogAgeLimit   = $auditConfig.AdminAuditLogAgeLimit
                }
            } else { $null }
        }

        $analysis = @{
            JournalRulesCount    = $journalRules.Count
            ActiveJournalRules   = ($journalRules | Where-Object { $_.Enabled }).Count
            AuditingEnabled      = $auditConfig.UnifiedAuditLogIngestionEnabled
        }

        # Detect gotchas
        if ($journalRules.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Journal Rules Configured" `
                -Description "Found $($journalRules.Count) journal rule(s). Journaling must be recreated in target tenant." `
                -Severity "Critical" `
                -Recommendation "Document all journal rules. Verify journal recipients exist in target. May have compliance/legal implications." `
                -AffectedCount $journalRules.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Configuration = $config
            Analysis      = $analysis
        }

        Add-CollectedData -Category "Exchange" -SubCategory "Journaling" -Data $result
        Write-Log -Message "Exchange journaling configuration collected" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect journaling configuration: $_" -Level Error
        throw
    }
}
#endregion

#region Calendar and Resource Configuration
function Get-ExchangeResourceConfig {
    <#
    .SYNOPSIS
        Collects calendar and resource booking configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Exchange resource configuration..." -Level Info

    try {
        # Get resource mailboxes with booking policies
        $roomMailboxes = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue
        $equipmentMailboxes = Get-Mailbox -RecipientTypeDetails EquipmentMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue

        $resourceDetails = @()

        foreach ($resource in ($roomMailboxes + $equipmentMailboxes)) {
            try {
                $calConfig = Get-CalendarProcessing -Identity $resource.PrimarySmtpAddress -ErrorAction SilentlyContinue

                $resourceDetails += @{
                    DisplayName                = $resource.DisplayName
                    PrimarySmtpAddress         = $resource.PrimarySmtpAddress
                    RecipientTypeDetails       = $resource.RecipientTypeDetails
                    ResourceType               = $resource.ResourceType
                    ResourceCapacity           = $resource.ResourceCapacity
                    AutomateProcessing         = $calConfig.AutomateProcessing
                    AllowConflicts             = $calConfig.AllowConflicts
                    AllowRecurringMeetings     = $calConfig.AllowRecurringMeetings
                    BookingWindowInDays        = $calConfig.BookingWindowInDays
                    MaximumDurationInMinutes   = $calConfig.MaximumDurationInMinutes
                    ResourceDelegates          = $calConfig.ResourceDelegates
                    BookInPolicy               = $calConfig.BookInPolicy
                    RequestInPolicy            = $calConfig.RequestInPolicy
                    RequestOutOfPolicy         = $calConfig.RequestOutOfPolicy
                }
            }
            catch {
                $resourceDetails += @{
                    DisplayName          = $resource.DisplayName
                    PrimarySmtpAddress   = $resource.PrimarySmtpAddress
                    RecipientTypeDetails = $resource.RecipientTypeDetails
                    Error                = "Could not retrieve calendar processing"
                }
            }
        }

        $analysis = @{
            TotalRoomMailboxes       = $roomMailboxes.Count
            TotalEquipmentMailboxes  = $equipmentMailboxes.Count
            AutoAcceptEnabled        = ($resourceDetails | Where-Object { $_.AutomateProcessing -eq "AutoAccept" }).Count
            WithDelegates            = ($resourceDetails | Where-Object { $_.ResourceDelegates.Count -gt 0 }).Count
            WithBookingPolicies      = ($resourceDetails | Where-Object { $_.BookInPolicy.Count -gt 0 }).Count
        }

        # Detect gotchas
        if ($resourceDetails.Count -gt 20) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Significant Resource Mailbox Estate" `
                -Description "Found $($resourceDetails.Count) resource mailboxes (rooms/equipment). Booking policies must be recreated." `
                -Severity "Medium" `
                -Recommendation "Export resource booking configurations. Plan for policy recreation in target. Consider impact on room booking during migration." `
                -AffectedCount $resourceDetails.Count `
                -MigrationPhase "Pre-Migration"
        }

        $complexResources = $resourceDetails | Where-Object {
            $_.BookInPolicy.Count -gt 0 -or $_.RequestInPolicy.Count -gt 0 -or $_.ResourceDelegates.Count -gt 0
        }

        if ($complexResources.Count -gt 0) {
            Add-MigrationGotcha -Category "Exchange" `
                -Title "Resources with Complex Booking Policies" `
                -Description "Found $($complexResources.Count) resources with custom booking policies, delegates, or permissions." `
                -Severity "Medium" `
                -Recommendation "Document all booking configurations. Group/user IDs in policies will need updating in target." `
                -AffectedCount $complexResources.Count `
                -MigrationPhase "Post-Migration"
        }

        $result = @{
            Resources = $resourceDetails
            Analysis  = $analysis
        }

        Add-CollectedData -Category "Exchange" -SubCategory "Resources" -Data $result
        Write-Log -Message "Collected $($resourceDetails.Count) resource mailboxes" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect resource configuration: $_" -Level Error
        throw
    }
}
#endregion

#region Main Collection Function
function Invoke-ExchangeCollection {
    <#
    .SYNOPSIS
        Runs all Exchange Online data collection functions
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [hashtable]$Config
    )

    Write-Log -Message "Starting Exchange Online data collection..." -Level Info

    $results = @{
        StartTime = Get-Date
        Collections = @{}
        Errors = @()
    }

    $collections = @(
        @{ Name = "OrganizationConfig"; Function = { Get-ExchangeOrganizationConfig } }
        @{ Name = "Mailboxes"; Function = { Get-ExchangeMailboxes } }
        @{ Name = "DistributionLists"; Function = { Get-ExchangeDistributionLists } }
        @{ Name = "PublicFolders"; Function = { Get-ExchangePublicFolders } }
        @{ Name = "TransportConfig"; Function = { Get-ExchangeTransportConfig } }
        @{ Name = "AddressPolicies"; Function = { Get-ExchangeAddressPolicies } }
        @{ Name = "HybridConfig"; Function = { Get-ExchangeHybridConfiguration } }
        @{ Name = "Journaling"; Function = { Get-ExchangeJournalingConfig } }
        @{ Name = "Resources"; Function = { Get-ExchangeResourceConfig } }
    )

    foreach ($collection in $collections) {
        try {
            Write-Progress -Activity "Exchange Collection" -Status "Collecting $($collection.Name)..."
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

    Write-Log -Message "Exchange collection completed in $($results.Duration.TotalMinutes.ToString('F2')) minutes" -Level Success

    return $results
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Get-ExchangeOrganizationConfig',
    'Get-ExchangeMailboxes',
    'Get-ExchangeDistributionLists',
    'Get-ExchangePublicFolders',
    'Get-ExchangeTransportConfig',
    'Get-ExchangeAddressPolicies',
    'Get-ExchangeHybridConfiguration',
    'Get-ExchangeJournalingConfig',
    'Get-ExchangeResourceConfig',
    'Invoke-ExchangeCollection'
)
