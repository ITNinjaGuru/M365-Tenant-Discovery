#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft Teams Data Collection Module
.DESCRIPTION
    Collects comprehensive Teams data including teams, channels, policies,
    apps, and governance settings. Identifies migration gotchas.
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

#region Teams Configuration
function Get-TeamsConfiguration {
    <#
    .SYNOPSIS
        Collects Teams tenant configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Teams configuration..." -Level Info

    try {
        $teamsConfig = Get-CsTeamsClientConfiguration -ErrorAction SilentlyContinue
        $meetingConfig = Get-CsTeamsMeetingConfiguration -ErrorAction SilentlyContinue
        $messagingConfig = Get-CsTeamsMessagingPolicy -Identity Global -ErrorAction SilentlyContinue
        $guestConfig = Get-CsTeamsGuestMeetingConfiguration -ErrorAction SilentlyContinue
        $callingConfig = Get-CsTeamsCallingPolicy -Identity Global -ErrorAction SilentlyContinue

        $config = @{
            ClientConfiguration = @{
                AllowDropBox            = $teamsConfig.AllowDropBox
                AllowBox                = $teamsConfig.AllowBox
                AllowGoogleDrive        = $teamsConfig.AllowGoogleDrive
                AllowShareFile          = $teamsConfig.AllowShareFile
                AllowEgnyte             = $teamsConfig.AllowEgnyte
                AllowEmailIntoChannel   = $teamsConfig.AllowEmailIntoChannel
                AllowOrganizationTab    = $teamsConfig.AllowOrganizationTab
                AllowSkypeBusinessInterop = $teamsConfig.AllowSkypeBusinessInterop
                AllowGuestUser          = $teamsConfig.AllowGuestUser
                ContentPin              = $teamsConfig.ContentPin
                ResourceAccountContentAccess = $teamsConfig.ResourceAccountContentAccess
            }
            MeetingConfiguration = @{
                LogoURL                 = $meetingConfig.LogoURL
                LegalURL                = $meetingConfig.LegalURL
                HelpURL                 = $meetingConfig.HelpURL
                CustomFooterText        = $meetingConfig.CustomFooterText
                DisableAnonymousJoin    = $meetingConfig.DisableAnonymousJoin
                EnableQoS               = $meetingConfig.EnableQoS
                ClientAudioPort         = $meetingConfig.ClientAudioPort
                ClientAudioPortRange    = $meetingConfig.ClientAudioPortRange
                ClientVideoPort         = $meetingConfig.ClientVideoPort
                ClientVideoPortRange    = $meetingConfig.ClientVideoPortRange
            }
            GuestConfiguration = @{
                AllowIPVideo            = $guestConfig.AllowIPVideo
                ScreenSharingMode       = $guestConfig.ScreenSharingMode
                AllowMeetNow            = $guestConfig.AllowMeetNow
            }
        }

        # Detect gotchas
        if ($teamsConfig.AllowGuestUser) {
            Add-MigrationGotcha -Category "Teams" `
                -Title "Guest Access Enabled" `
                -Description "Guest access is enabled for Teams. Guest users will need reinvitation in target tenant." `
                -Severity "Medium" `
                -Recommendation "Document guest users and their team memberships. Plan for guest reinvitation post-migration." `
                -MigrationPhase "Post-Migration"
        }

        # Third-party cloud storage
        $thirdPartyStorage = @()
        if ($teamsConfig.AllowDropBox) { $thirdPartyStorage += "DropBox" }
        if ($teamsConfig.AllowBox) { $thirdPartyStorage += "Box" }
        if ($teamsConfig.AllowGoogleDrive) { $thirdPartyStorage += "Google Drive" }

        if ($thirdPartyStorage.Count -gt 0) {
            Add-MigrationGotcha -Category "Teams" `
                -Title "Third-Party Cloud Storage Enabled" `
                -Description "Teams allows third-party cloud storage: $($thirdPartyStorage -join ', '). Users may have linked external storage." `
                -Severity "Low" `
                -Recommendation "Review third-party integrations. Users will need to reconfigure integrations post-migration." `
                -MigrationPhase "Post-Migration"
        }

        Add-CollectedData -Category "Teams" -SubCategory "Configuration" -Data $config
        Write-Log -Message "Teams configuration collected" -Level Success

        return $config
    }
    catch {
        Write-Log -Message "Failed to collect Teams config: $_" -Level Error
        throw
    }
}
#endregion

#region Teams Collection
function Get-TeamsInventory {
    <#
    .SYNOPSIS
        Collects Teams inventory using Graph API
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Teams inventory..." -Level Info

    try {
        # Get all teams using Graph API
        $teams = @()
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&`$select=id,displayName,description,visibility,createdDateTime,mailNickname,mail"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri

        $teams += $response.value

        while ($response.'@odata.nextLink') {
            $response = Invoke-MgGraphRequest -Method GET -Uri $response.'@odata.nextLink'
            $teams += $response.value
        }

        # Get team details
        $teamDetails = foreach ($team in $teams) {
            try {
                $teamInfo = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$($team.id)" -ErrorAction SilentlyContinue

                # Get channels
                $channels = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$($team.id)/channels"

                # Get members count
                $members = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($team.id)/members?`$count=true" -Headers @{"ConsistencyLevel"="eventual"}

                # Get owners
                $owners = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($team.id)/owners"

                @{
                    Id                = $team.id
                    DisplayName       = $team.displayName
                    Description       = $team.description
                    Visibility        = $team.visibility
                    CreatedDateTime   = $team.createdDateTime
                    Mail              = $team.mail
                    MailNickname      = $team.mailNickname
                    IsArchived        = $teamInfo.isArchived
                    MemberSettings    = $teamInfo.memberSettings
                    GuestSettings     = $teamInfo.guestSettings
                    MessagingSettings = $teamInfo.messagingSettings
                    FunSettings       = $teamInfo.funSettings
                    ChannelCount      = $channels.value.Count
                    MemberCount       = $members.'@odata.count'
                    OwnerCount        = $owners.value.Count
                    Owners            = @($owners.value | Select-Object -ExpandProperty userPrincipalName -ErrorAction SilentlyContinue)
                    Channels          = @($channels.value | ForEach-Object {
                        @{
                            Id              = $_.id
                            DisplayName     = $_.displayName
                            MembershipType  = $_.membershipType
                            Description     = $_.description
                        }
                    })
                }
            }
            catch {
                @{
                    Id          = $team.id
                    DisplayName = $team.displayName
                    Error       = $_.Exception.Message
                }
            }
        }

        # Analyze teams
        $archivedTeams = $teamDetails | Where-Object { $_.IsArchived }
        $privateChannels = $teamDetails | ForEach-Object {
            $_.Channels | Where-Object { $_.MembershipType -eq "private" }
        }
        $sharedChannels = $teamDetails | ForEach-Object {
            $_.Channels | Where-Object { $_.MembershipType -eq "shared" }
        }

        $analysis = @{
            TotalTeams          = $teams.Count
            ArchivedTeams       = $archivedTeams.Count
            PublicTeams         = ($teamDetails | Where-Object { $_.Visibility -eq "Public" }).Count
            PrivateTeams        = ($teamDetails | Where-Object { $_.Visibility -eq "Private" }).Count
            TotalChannels       = ($teamDetails | ForEach-Object { $_.ChannelCount } | Measure-Object -Sum).Sum
            PrivateChannels     = ($privateChannels | Measure-Object).Count
            SharedChannels      = ($sharedChannels | Measure-Object).Count
            TeamsWithGuests     = ($teamDetails | Where-Object { $_.GuestSettings.allowCreateUpdateChannels -or $_.GuestSettings.allowDeleteChannels }).Count
        }

        # Detect gotchas

        # Private channels
        if (($privateChannels | Measure-Object).Count -gt 0) {
            Add-MigrationGotcha -Category "Teams" `
                -Title "Private Channels Present" `
                -Description "Found $(($privateChannels | Measure-Object).Count) private channels across teams. Private channels have separate SharePoint sites." `
                -Severity "High" `
                -Recommendation "Private channels require special migration handling. Each private channel has its own SharePoint site collection." `
                -AffectedCount ($privateChannels | Measure-Object).Count `
                -MigrationPhase "Pre-Migration"
        }

        # Shared channels (cross-tenant)
        if (($sharedChannels | Measure-Object).Count -gt 0) {
            Add-MigrationGotcha -Category "Teams" `
                -Title "Shared Channels Present" `
                -Description "Found $(($sharedChannels | Measure-Object).Count) shared channels. These may include cross-tenant sharing." `
                -Severity "Critical" `
                -Recommendation "Shared channels with external tenants will break during migration. Document all external participants." `
                -AffectedCount ($sharedChannels | Measure-Object).Count `
                -MigrationPhase "Pre-Migration"
        }

        # Large teams
        $largeTeams = $teamDetails | Where-Object { $_.MemberCount -gt 1000 }
        if ($largeTeams.Count -gt 0) {
            Add-MigrationGotcha -Category "Teams" `
                -Title "Large Teams Detected" `
                -Description "Found $($largeTeams.Count) teams with more than 1000 members. These may require extended migration time." `
                -Severity "Medium" `
                -Recommendation "Plan for incremental membership migration for large teams." `
                -AffectedObjects @($largeTeams.DisplayName) `
                -MigrationPhase "Pre-Migration"
        }

        # Teams with no owners
        $noOwners = $teamDetails | Where-Object { $_.OwnerCount -eq 0 }
        if ($noOwners.Count -gt 0) {
            Add-MigrationGotcha -Category "Teams" `
                -Title "Teams Without Owners" `
                -Description "Found $($noOwners.Count) teams with no owners. Orphaned teams need owner assignment." `
                -Severity "Medium" `
                -Recommendation "Assign owners to orphaned teams before migration." `
                -AffectedObjects @($noOwners.DisplayName) `
                -MigrationPhase "Pre-Migration"
        }

        # Archived teams
        if ($archivedTeams.Count -gt 0) {
            Add-MigrationGotcha -Category "Teams" `
                -Title "Archived Teams Present" `
                -Description "Found $($archivedTeams.Count) archived teams. Archive status needs preservation during migration." `
                -Severity "Low" `
                -Recommendation "Document archived teams. Ensure archive status is set after migration." `
                -AffectedCount $archivedTeams.Count `
                -MigrationPhase "Post-Migration"
        }

        $result = @{
            Teams    = $teamDetails
            Analysis = $analysis
        }

        Add-CollectedData -Category "Teams" -SubCategory "Teams" -Data $result
        Write-Log -Message "Collected $($teams.Count) teams" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Teams inventory: $_" -Level Error
        throw
    }
}
#endregion

#region Teams Policies
function Get-TeamsPolicies {
    <#
    .SYNOPSIS
        Collects Teams policies
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Teams policies..." -Level Info

    try {
        $policies = @{
            MeetingPolicies     = Get-CsTeamsMeetingPolicy | Select-Object Identity, Description, AllowChannelMeetingScheduling, AllowMeetNow, AllowPrivateMeetNow, AllowAnonymousUsersToJoinMeeting, AllowCloudRecording, AllowTranscription
            MessagingPolicies   = Get-CsTeamsMessagingPolicy | Select-Object Identity, Description, AllowUrlPreviews, AllowOwnerDeleteMessage, AllowUserEditMessage, AllowUserDeleteMessage, AllowUserChat, AllowGiphy, AllowMemes, AllowStickers
            CallingPolicies     = Get-CsTeamsCallingPolicy | Select-Object Identity, Description, AllowPrivateCalling, AllowVoicemail, AllowCallGroups, AllowDelegation, AllowCallForwardingToUser, AllowCallForwardingToPhone
            AppPermissionPolicies = Get-CsTeamsAppPermissionPolicy | Select-Object Identity, Description, DefaultCatalogApps, GlobalCatalogApps, PrivateCatalogApps
            AppSetupPolicies    = Get-CsTeamsAppSetupPolicy | Select-Object Identity, Description, AllowUserPinning, AllowSideloading, PinnedAppBarApps
            LiveEventPolicies   = Get-CsTeamsMeetingBroadcastPolicy | Select-Object Identity, Description, AllowBroadcastScheduling, AllowBroadcastTranscription, BroadcastAttendeeVisibilityMode, BroadcastRecordingMode
        }

        $customPolicies = @{
            CustomMeetingPolicies    = ($policies.MeetingPolicies | Where-Object { $_.Identity -ne "Global" -and $_.Identity -notlike "Tag:*" }).Count
            CustomMessagingPolicies  = ($policies.MessagingPolicies | Where-Object { $_.Identity -ne "Global" }).Count
            CustomCallingPolicies    = ($policies.CallingPolicies | Where-Object { $_.Identity -ne "Global" }).Count
            CustomAppPolicies        = ($policies.AppPermissionPolicies | Where-Object { $_.Identity -ne "Global" }).Count
        }

        # Detect gotchas
        $totalCustomPolicies = ($customPolicies.Values | Measure-Object -Sum).Sum
        if ($totalCustomPolicies -gt 0) {
            Add-MigrationGotcha -Category "Teams" `
                -Title "Custom Teams Policies" `
                -Description "Found $totalCustomPolicies custom Teams policies. These need recreation in target tenant." `
                -Severity "Medium" `
                -Recommendation "Export policy configurations. Recreate policies in target tenant. Update policy assignments for users/groups." `
                -MigrationPhase "Pre-Migration"
        }

        # Check for Teams apps
        $customApps = Get-CsTeamsAppPermissionPolicy | Where-Object { $_.PrivateCatalogApps -and $_.PrivateCatalogApps.Count -gt 0 }
        if ($customApps) {
            Add-MigrationGotcha -Category "Teams" `
                -Title "Custom Teams Apps in Use" `
                -Description "Custom or LOB Teams apps are configured in app policies. These need publishing in target tenant." `
                -Severity "High" `
                -Recommendation "Identify all custom Teams apps. Plan for app package migration and republishing." `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Policies      = $policies
            CustomCounts  = $customPolicies
        }

        Add-CollectedData -Category "Teams" -SubCategory "Policies" -Data $result
        Write-Log -Message "Teams policies collected" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Teams policies: $_" -Level Error
        throw
    }
}
#endregion

#region Teams Apps
function Get-TeamsApps {
    <#
    .SYNOPSIS
        Collects Teams app catalog and usage information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Teams apps..." -Level Info

    try {
        # Get app catalog
        $uri = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps"
        $apps = Invoke-MgGraphRequest -Method GET -Uri $uri

        $appDetails = foreach ($app in $apps.value) {
            @{
                Id               = $app.id
                ExternalId       = $app.externalId
                DisplayName      = $app.displayName
                DistributionMethod = $app.distributionMethod
            }
        }

        # Custom apps (sideloaded or org-specific)
        $customApps = $appDetails | Where-Object { $_.DistributionMethod -eq "organization" -or $_.DistributionMethod -eq "sideloaded" }

        $analysis = @{
            TotalApps  = $apps.value.Count
            CustomApps = $customApps.Count
            StoreApps  = ($appDetails | Where-Object { $_.DistributionMethod -eq "store" }).Count
        }

        # Detect gotchas
        if ($customApps.Count -gt 0) {
            Add-MigrationGotcha -Category "Teams" `
                -Title "Custom Teams Apps" `
                -Description "Found $($customApps.Count) custom/sideloaded Teams apps. These require app package migration." `
                -Severity "High" `
                -Recommendation "Export app packages. Update app manifests for target tenant. Republish to target app catalog." `
                -AffectedObjects @($customApps.DisplayName) `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Apps       = $appDetails
            CustomApps = $customApps
            Analysis   = $analysis
        }

        Add-CollectedData -Category "Teams" -SubCategory "Apps" -Data $result
        Write-Log -Message "Collected $($apps.value.Count) Teams apps" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Teams apps: $_" -Level Error
        throw
    }
}
#endregion

#region Phone System
function Get-TeamsPhoneSystem {
    <#
    .SYNOPSIS
        Collects Teams Phone System configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Teams Phone System configuration..." -Level Info

    try {
        $phoneConfig = @{
            VoiceRoutes        = Get-CsOnlineVoiceRoute -ErrorAction SilentlyContinue
            PSTNUsages         = Get-CsOnlinePstnUsage -ErrorAction SilentlyContinue
            VoiceRoutingPolicies = Get-CsOnlineVoiceRoutingPolicy -ErrorAction SilentlyContinue
            DialPlans          = Get-CsTenantDialPlan -ErrorAction SilentlyContinue
            EmergencyPolicies  = Get-CsTeamsEmergencyCallingPolicy -ErrorAction SilentlyContinue
            CallQueues         = Get-CsCallQueue -ErrorAction SilentlyContinue
            AutoAttendants     = Get-CsAutoAttendant -ErrorAction SilentlyContinue
            ResourceAccounts   = Get-CsOnlineApplicationInstance -ErrorAction SilentlyContinue
        }

        # Check if Phone System is in use
        $hasPhoneSystem = $phoneConfig.VoiceRoutes -or $phoneConfig.CallQueues -or $phoneConfig.AutoAttendants

        $analysis = @{
            PhoneSystemEnabled   = $hasPhoneSystem
            VoiceRouteCount      = ($phoneConfig.VoiceRoutes | Measure-Object).Count
            DialPlanCount        = ($phoneConfig.DialPlans | Measure-Object).Count
            CallQueueCount       = ($phoneConfig.CallQueues | Measure-Object).Count
            AutoAttendantCount   = ($phoneConfig.AutoAttendants | Measure-Object).Count
            ResourceAccountCount = ($phoneConfig.ResourceAccounts | Measure-Object).Count
        }

        # Detect gotchas
        if ($hasPhoneSystem) {
            Add-MigrationGotcha -Category "Teams" `
                -Title "Teams Phone System Configured" `
                -Description "Teams Phone System is configured with voice routes, call queues, and/or auto attendants. This requires specialized migration." `
                -Severity "Critical" `
                -Recommendation "Plan separate Phone System migration. Document all configurations, numbers, and resource accounts." `
                -MigrationPhase "Pre-Migration" `
                -AdditionalData $analysis
        }

        if (($phoneConfig.CallQueues | Measure-Object).Count -gt 0) {
            Add-MigrationGotcha -Category "Teams" `
                -Title "Call Queues Configured" `
                -Description "Found $(($phoneConfig.CallQueues | Measure-Object).Count) call queues. Call queues need recreation in target tenant." `
                -Severity "High" `
                -Recommendation "Document call queue configurations including agents, overflow settings, and holiday schedules." `
                -AffectedCount ($phoneConfig.CallQueues | Measure-Object).Count `
                -MigrationPhase "Pre-Migration"
        }

        if (($phoneConfig.AutoAttendants | Measure-Object).Count -gt 0) {
            Add-MigrationGotcha -Category "Teams" `
                -Title "Auto Attendants Configured" `
                -Description "Found $(($phoneConfig.AutoAttendants | Measure-Object).Count) auto attendants. Auto attendants need recreation in target tenant." `
                -Severity "High" `
                -Recommendation "Document auto attendant configurations including greetings, menus, and business hours." `
                -AffectedCount ($phoneConfig.AutoAttendants | Measure-Object).Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Configuration = $phoneConfig
            Analysis      = $analysis
        }

        Add-CollectedData -Category "Teams" -SubCategory "PhoneSystem" -Data $result
        Write-Log -Message "Teams Phone System configuration collected" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Phone System config: $_" -Level Error
        throw
    }
}
#endregion

#region Main Collection Function
function Invoke-TeamsCollection {
    <#
    .SYNOPSIS
        Runs all Teams data collection functions
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [hashtable]$Config
    )

    Write-Log -Message "Starting Teams data collection..." -Level Info

    $results = @{
        StartTime = Get-Date
        Collections = @{}
        Errors = @()
    }

    $collections = @(
        @{ Name = "Configuration"; Function = { Get-TeamsConfiguration } }
        @{ Name = "Teams"; Function = { Get-TeamsInventory } }
        @{ Name = "Policies"; Function = { Get-TeamsPolicies } }
        @{ Name = "Apps"; Function = { Get-TeamsApps } }
        @{ Name = "PhoneSystem"; Function = { Get-TeamsPhoneSystem } }
    )

    foreach ($collection in $collections) {
        try {
            Write-Progress -Activity "Teams Collection" -Status "Collecting $($collection.Name)..."
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

    Write-Log -Message "Teams collection completed in $($results.Duration.TotalMinutes.ToString('F2')) minutes" -Level Success

    return $results
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Get-TeamsConfiguration',
    'Get-TeamsInventory',
    'Get-TeamsPolicies',
    'Get-TeamsApps',
    'Get-TeamsPhoneSystem',
    'Invoke-TeamsCollection'
)
