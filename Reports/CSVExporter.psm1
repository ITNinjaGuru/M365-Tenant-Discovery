#Requires -Version 7.0
<#
.SYNOPSIS
    CSV Data Exporter for M365 Tenant Discovery
.DESCRIPTION
    Exports discovery data to CSV files for analysis, migration planning, and reporting.
    Generates the following CSV files:
    - Users.csv - All Entra ID users with key attributes
    - Mailboxes.csv - Exchange mailboxes with size and configuration
    - MailboxPermissions.csv - Full access, send-as, send-on-behalf permissions
    - SharePointSites.csv - SharePoint site collections
    - OneDriveSites.csv - OneDrive for Business sites
    - Teams.csv - Microsoft Teams with membership info
    - Groups.csv - M365 and security groups
    - GotchaFindings.csv - Migration risks and remediation steps
    - LicenseSummary.csv - License allocation summary
    - ConditionalAccessPolicies.csv - CA policies
    - Devices.csv - Entra ID registered/joined devices
.NOTES
    Author: M365 Migration Tool
    Version: 1.0.0
#>

# Import core module for logging
if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    $corePath = Join-Path $PSScriptRoot ".." "Modules" "Core" "TenantDiscovery.Core.psm1"
    if (Test-Path $corePath) {
        Import-Module $corePath -Force -Global
    }
}

#region Main Export Function
function Export-DiscoveryDataToCSV {
    <#
    .SYNOPSIS
        Exports all discovery data to CSV files
    .PARAMETER CollectedData
        The collected discovery data hashtable
    .PARAMETER AnalysisResults
        The gotcha analysis results (optional)
    .PARAMETER OutputPath
        Directory where CSV files will be saved
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $false)]
        [hashtable]$AnalysisResults,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Starting CSV export..." -Level Info

    # Create CSV output directory
    $csvPath = Join-Path $OutputPath "CSV"
    if (-not (Test-Path $csvPath)) {
        New-Item -Path $csvPath -ItemType Directory -Force | Out-Null
    }

    $exportedFiles = @()

    # Export Users
    if ($CollectedData.EntraID.Users.Users) {
        $file = Export-UsersToCSV -Users $CollectedData.EntraID.Users.Users -OutputPath $csvPath
        if ($file) { $exportedFiles += $file }
    }

    # Export Mailboxes
    if ($CollectedData.Exchange.Mailboxes.Mailboxes) {
        $file = Export-MailboxesToCSV -Mailboxes $CollectedData.Exchange.Mailboxes.Mailboxes -OutputPath $csvPath
        if ($file) { $exportedFiles += $file }
    }

    # Export Mailbox Permissions
    if ($CollectedData.Exchange.Permissions) {
        $file = Export-MailboxPermissionsToCSV -Permissions $CollectedData.Exchange.Permissions -OutputPath $csvPath
        if ($file) { $exportedFiles += $file }
    }

    # Export SharePoint Sites
    if ($CollectedData.SharePoint.Sites.Sites) {
        $file = Export-SharePointSitesToCSV -Sites $CollectedData.SharePoint.Sites.Sites -OutputPath $csvPath
        if ($file) { $exportedFiles += $file }
    }

    # Export OneDrive Sites
    if ($CollectedData.SharePoint.OneDrive.Sites) {
        $file = Export-OneDriveSitesToCSV -Sites $CollectedData.SharePoint.OneDrive.Sites -OutputPath $csvPath
        if ($file) { $exportedFiles += $file }
    }

    # Export Teams
    if ($CollectedData.Teams.Teams.Teams) {
        $file = Export-TeamsToCSV -Teams $CollectedData.Teams.Teams.Teams -OutputPath $csvPath
        if ($file) { $exportedFiles += $file }
    }

    # Export Groups
    if ($CollectedData.EntraID.Groups.Groups) {
        $file = Export-GroupsToCSV -Groups $CollectedData.EntraID.Groups.Groups -OutputPath $csvPath
        if ($file) { $exportedFiles += $file }
    }

    # Export Licenses
    if ($CollectedData.EntraID.Licenses) {
        $file = Export-LicensesToCSV -Licenses $CollectedData.EntraID.Licenses -OutputPath $csvPath
        if ($file) { $exportedFiles += $file }
    }

    # Export Conditional Access Policies
    if ($CollectedData.EntraID.ConditionalAccess.Policies) {
        $file = Export-ConditionalAccessToCSV -Policies $CollectedData.EntraID.ConditionalAccess.Policies -OutputPath $csvPath
        if ($file) { $exportedFiles += $file }
    }

    # Export Devices
    if ($CollectedData.EntraID.Devices.Devices) {
        $file = Export-DevicesToCSV -Devices $CollectedData.EntraID.Devices.Devices -OutputPath $csvPath
        if ($file) { $exportedFiles += $file }
    }

    # Export Gotcha Findings
    if ($AnalysisResults -and $AnalysisResults.AllIssues) {
        $file = Export-GotchaFindingsToCSV -Issues $AnalysisResults.AllIssues -OutputPath $csvPath
        if ($file) { $exportedFiles += $file }
    }

    # Export Distribution Lists
    if ($CollectedData.Exchange.DistributionGroups.Groups) {
        $file = Export-DistributionGroupsToCSV -Groups $CollectedData.Exchange.DistributionGroups.Groups -OutputPath $csvPath
        if ($file) { $exportedFiles += $file }
    }

    # Export Shared Mailboxes separately
    if ($CollectedData.Exchange.Mailboxes.Mailboxes) {
        $sharedMailboxes = $CollectedData.Exchange.Mailboxes.Mailboxes | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" }
        if ($sharedMailboxes) {
            $file = Export-SharedMailboxesToCSV -Mailboxes $sharedMailboxes -OutputPath $csvPath
            if ($file) { $exportedFiles += $file }
        }
    }

    Write-Log -Message "CSV export complete. $($exportedFiles.Count) files created in $csvPath" -Level Success

    return @{
        OutputPath = $csvPath
        Files      = $exportedFiles
        Count      = $exportedFiles.Count
    }
}
#endregion

#region Individual Export Functions
function Export-UsersToCSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Users,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Exporting Users to CSV..." -Level Info

    $exportData = foreach ($user in $Users) {
        [PSCustomObject]@{
            DisplayName           = $user.DisplayName
            UserPrincipalName     = $user.UserPrincipalName
            Mail                  = $user.Mail
            JobTitle              = $user.JobTitle
            Department            = $user.Department
            Office                = $user.OfficeLocation
            City                  = $user.City
            Country               = $user.Country
            AccountEnabled        = $user.AccountEnabled
            UserType              = $user.UserType
            CreatedDateTime       = $user.CreatedDateTime
            LastSignInDateTime    = $user.SignInActivity.LastSignInDateTime
            OnPremisesSyncEnabled = $user.OnPremisesSyncEnabled
            OnPremisesImmutableId = $user.OnPremisesImmutableId
            AssignedLicenses      = if ($user.AssignedLicenses) { ($user.AssignedLicenses.SkuId -join "; ") } else { "" }
            ProxyAddresses        = if ($user.ProxyAddresses) { ($user.ProxyAddresses -join "; ") } else { "" }
            Id                    = $user.Id
        }
    }

    $filePath = Join-Path $OutputPath "Users.csv"
    $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "  Exported $($exportData.Count) users" -Level Info

    return $filePath
}

function Export-MailboxesToCSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Mailboxes,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Exporting Mailboxes to CSV..." -Level Info

    $exportData = foreach ($mbx in $Mailboxes) {
        [PSCustomObject]@{
            DisplayName              = $mbx.DisplayName
            UserPrincipalName        = $mbx.UserPrincipalName
            PrimarySmtpAddress       = $mbx.PrimarySmtpAddress
            RecipientTypeDetails     = $mbx.RecipientTypeDetails
            MailboxSizeGB            = if ($mbx.TotalItemSize) { [math]::Round($mbx.TotalItemSize / 1GB, 2) } else { $null }
            ItemCount                = $mbx.ItemCount
            ArchiveEnabled           = $mbx.ArchiveStatus -eq "Active"
            ArchiveSizeGB            = if ($mbx.ArchiveTotalItemSize) { [math]::Round($mbx.ArchiveTotalItemSize / 1GB, 2) } else { $null }
            LitigationHoldEnabled    = $mbx.LitigationHoldEnabled
            InPlaceHolds             = if ($mbx.InPlaceHolds) { ($mbx.InPlaceHolds -join "; ") } else { "" }
            RetentionPolicy          = $mbx.RetentionPolicy
            ForwardingAddress        = $mbx.ForwardingAddress
            ForwardingSmtpAddress    = $mbx.ForwardingSmtpAddress
            DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
            HiddenFromAddressLists   = $mbx.HiddenFromAddressListsEnabled
            EmailAddresses           = if ($mbx.EmailAddresses) { ($mbx.EmailAddresses -join "; ") } else { "" }
            WhenCreated              = $mbx.WhenCreated
            ExchangeGuid             = $mbx.ExchangeGuid
        }
    }

    $filePath = Join-Path $OutputPath "Mailboxes.csv"
    $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "  Exported $($exportData.Count) mailboxes" -Level Info

    return $filePath
}

function Export-MailboxPermissionsToCSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Permissions,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Exporting Mailbox Permissions to CSV..." -Level Info

    $exportData = [System.Collections.ArrayList]@()

    # Full Access permissions
    if ($Permissions.FullAccess) {
        foreach ($perm in $Permissions.FullAccess) {
            $null = $exportData.Add([PSCustomObject]@{
                MailboxIdentity    = $perm.Identity
                MailboxDisplayName = $perm.MailboxDisplayName
                PermissionType     = "FullAccess"
                Trustee            = $perm.User
                TrusteeDisplayName = $perm.UserDisplayName
                AccessRights       = "FullAccess"
                IsInherited        = $perm.IsInherited
                Deny               = $perm.Deny
            })
        }
    }

    # Send-As permissions
    if ($Permissions.SendAs) {
        foreach ($perm in $Permissions.SendAs) {
            $null = $exportData.Add([PSCustomObject]@{
                MailboxIdentity    = $perm.Identity
                MailboxDisplayName = $perm.MailboxDisplayName
                PermissionType     = "SendAs"
                Trustee            = $perm.Trustee
                TrusteeDisplayName = $perm.TrusteeDisplayName
                AccessRights       = "SendAs"
                IsInherited        = $perm.IsInherited
                Deny               = $false
            })
        }
    }

    # Send-On-Behalf permissions
    if ($Permissions.SendOnBehalf) {
        foreach ($perm in $Permissions.SendOnBehalf) {
            $null = $exportData.Add([PSCustomObject]@{
                MailboxIdentity    = $perm.Identity
                MailboxDisplayName = $perm.MailboxDisplayName
                PermissionType     = "SendOnBehalf"
                Trustee            = $perm.GrantedTo
                TrusteeDisplayName = $perm.GrantedToDisplayName
                AccessRights       = "SendOnBehalf"
                IsInherited        = $false
                Deny               = $false
            })
        }
    }

    # Calendar Delegates
    if ($Permissions.CalendarDelegates) {
        foreach ($perm in $Permissions.CalendarDelegates) {
            $null = $exportData.Add([PSCustomObject]@{
                MailboxIdentity    = $perm.Identity
                MailboxDisplayName = $perm.MailboxDisplayName
                PermissionType     = "CalendarDelegate"
                Trustee            = $perm.Delegate
                TrusteeDisplayName = $perm.DelegateDisplayName
                AccessRights       = $perm.AccessRights
                IsInherited        = $false
                Deny               = $false
            })
        }
    }

    $filePath = Join-Path $OutputPath "MailboxPermissions.csv"
    $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "  Exported $($exportData.Count) permission entries" -Level Info

    return $filePath
}

function Export-SharePointSitesToCSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Sites,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Exporting SharePoint Sites to CSV..." -Level Info

    # Filter out OneDrive sites
    $spSites = $Sites | Where-Object { $_.Template -ne "SPSPERS#10" -and $_.Url -notlike "*-my.sharepoint.com/personal/*" }

    $exportData = foreach ($site in $spSites) {
        [PSCustomObject]@{
            Title                = $site.Title
            Url                  = $site.Url
            Template             = $site.Template
            StorageUsedGB        = if ($site.StorageUsageCurrent) { [math]::Round($site.StorageUsageCurrent / 1024, 2) } else { $null }
            StorageQuotaGB       = if ($site.StorageQuota) { [math]::Round($site.StorageQuota / 1024, 2) } else { $null }
            StorageWarningLevelGB = if ($site.StorageQuotaWarningLevel) { [math]::Round($site.StorageQuotaWarningLevel / 1024, 2) } else { $null }
            Owner                = $site.Owner
            SharingCapability    = $site.SharingCapability
            LockState            = $site.LockState
            ConditionalAccessPolicy = $site.ConditionalAccessPolicy
            SensitivityLabel     = $site.SensitivityLabel
            HubSiteId            = $site.HubSiteId
            IsHubSite            = $site.IsHubSite
            LastContentModifiedDate = $site.LastContentModifiedDate
            GroupId              = $site.GroupId
            RelatedGroupId       = $site.RelatedGroupId
            SiteId               = $site.Id
        }
    }

    $filePath = Join-Path $OutputPath "SharePointSites.csv"
    $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "  Exported $($exportData.Count) SharePoint sites" -Level Info

    return $filePath
}

function Export-OneDriveSitesToCSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Sites,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Exporting OneDrive Sites to CSV..." -Level Info

    $exportData = foreach ($site in $Sites) {
        [PSCustomObject]@{
            Owner                = $site.Owner
            OwnerDisplayName     = $site.OwnerDisplayName
            Url                  = $site.Url
            StorageUsedGB        = if ($site.StorageUsageCurrent) { [math]::Round($site.StorageUsageCurrent / 1024, 2) } else { $null }
            StorageQuotaGB       = if ($site.StorageQuota) { [math]::Round($site.StorageQuota / 1024, 2) } else { $null }
            LastContentModifiedDate = $site.LastContentModifiedDate
            SharingCapability    = $site.SharingCapability
            LockState            = $site.LockState
            Status               = $site.Status
            SiteId               = $site.Id
        }
    }

    $filePath = Join-Path $OutputPath "OneDriveSites.csv"
    $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "  Exported $($exportData.Count) OneDrive sites" -Level Info

    return $filePath
}

function Export-TeamsToCSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Teams,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Exporting Teams to CSV..." -Level Info

    $exportData = foreach ($team in $Teams) {
        [PSCustomObject]@{
            DisplayName          = $team.DisplayName
            Description          = $team.Description
            Visibility           = $team.Visibility
            MailNickname         = $team.MailNickname
            MemberCount          = $team.MemberCount
            OwnerCount           = $team.OwnerCount
            GuestCount           = $team.GuestCount
            ChannelCount         = $team.ChannelCount
            PrivateChannelCount  = $team.PrivateChannelCount
            SharedChannelCount   = $team.SharedChannelCount
            IsArchived           = $team.IsArchived
            Classification       = $team.Classification
            SensitivityLabel     = $team.SensitivityLabel
            CreatedDateTime      = $team.CreatedDateTime
            RenewalDateTime      = $team.RenewalDateTime
            ExpirationDateTime   = $team.ExpirationDateTime
            GroupId              = $team.Id
            Owners               = if ($team.Owners) { ($team.Owners -join "; ") } else { "" }
        }
    }

    $filePath = Join-Path $OutputPath "Teams.csv"
    $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "  Exported $($exportData.Count) teams" -Level Info

    return $filePath
}

function Export-GroupsToCSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Groups,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Exporting Groups to CSV..." -Level Info

    $exportData = foreach ($group in $Groups) {
        $groupTypes = $group.GroupTypes -join ", "
        $groupType = if ($group.GroupTypes -contains "Unified") { "Microsoft 365" }
                     elseif ($group.SecurityEnabled -and $group.MailEnabled) { "Mail-Enabled Security" }
                     elseif ($group.SecurityEnabled) { "Security" }
                     elseif ($group.MailEnabled) { "Distribution" }
                     else { "Other" }

        [PSCustomObject]@{
            DisplayName          = $group.DisplayName
            Mail                 = $group.Mail
            MailNickname         = $group.MailNickname
            GroupType            = $groupType
            GroupTypes           = $groupTypes
            SecurityEnabled      = $group.SecurityEnabled
            MailEnabled          = $group.MailEnabled
            MembershipRule       = $group.MembershipRule
            MembershipRuleProcessingState = $group.MembershipRuleProcessingState
            IsDynamic            = $group.MembershipRuleProcessingState -eq "On"
            MemberCount          = $group.MemberCount
            OwnerCount           = $group.OwnerCount
            Visibility           = $group.Visibility
            CreatedDateTime      = $group.CreatedDateTime
            RenewalDateTime      = $group.RenewalDateTime
            ExpirationDateTime   = $group.ExpirationDateTime
            OnPremisesSyncEnabled = $group.OnPremisesSyncEnabled
            ProxyAddresses       = if ($group.ProxyAddresses) { ($group.ProxyAddresses -join "; ") } else { "" }
            GroupId              = $group.Id
        }
    }

    $filePath = Join-Path $OutputPath "Groups.csv"
    $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "  Exported $($exportData.Count) groups" -Level Info

    return $filePath
}

function Export-LicensesToCSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Licenses,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Exporting Licenses to CSV..." -Level Info

    $exportData = foreach ($license in $Licenses.Subscriptions) {
        [PSCustomObject]@{
            SkuPartNumber        = $license.SkuPartNumber
            SkuId                = $license.SkuId
            FriendlyName         = $license.FriendlyName
            ConsumedUnits        = $license.ConsumedUnits
            PrepaidUnitsEnabled  = $license.PrepaidUnits.Enabled
            PrepaidUnitsWarning  = $license.PrepaidUnits.Warning
            PrepaidUnitsSuspended = $license.PrepaidUnits.Suspended
            AvailableUnits       = $license.PrepaidUnits.Enabled - $license.ConsumedUnits
            CapabilityStatus     = $license.CapabilityStatus
            AppliesTo            = $license.AppliesTo
        }
    }

    $filePath = Join-Path $OutputPath "LicenseSummary.csv"
    $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "  Exported $($exportData.Count) license SKUs" -Level Info

    return $filePath
}

function Export-ConditionalAccessToCSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Policies,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Exporting Conditional Access Policies to CSV..." -Level Info

    $exportData = foreach ($policy in $Policies) {
        [PSCustomObject]@{
            DisplayName          = $policy.DisplayName
            State                = $policy.State
            CreatedDateTime      = $policy.CreatedDateTime
            ModifiedDateTime     = $policy.ModifiedDateTime
            IncludeUsers         = if ($policy.Conditions.Users.IncludeUsers) { ($policy.Conditions.Users.IncludeUsers -join "; ") } else { "" }
            ExcludeUsers         = if ($policy.Conditions.Users.ExcludeUsers) { ($policy.Conditions.Users.ExcludeUsers -join "; ") } else { "" }
            IncludeGroups        = if ($policy.Conditions.Users.IncludeGroups) { ($policy.Conditions.Users.IncludeGroups -join "; ") } else { "" }
            ExcludeGroups        = if ($policy.Conditions.Users.ExcludeGroups) { ($policy.Conditions.Users.ExcludeGroups -join "; ") } else { "" }
            IncludeApplications  = if ($policy.Conditions.Applications.IncludeApplications) { ($policy.Conditions.Applications.IncludeApplications -join "; ") } else { "" }
            ExcludeApplications  = if ($policy.Conditions.Applications.ExcludeApplications) { ($policy.Conditions.Applications.ExcludeApplications -join "; ") } else { "" }
            IncludePlatforms     = if ($policy.Conditions.Platforms.IncludePlatforms) { ($policy.Conditions.Platforms.IncludePlatforms -join "; ") } else { "" }
            IncludeLocations     = if ($policy.Conditions.Locations.IncludeLocations) { ($policy.Conditions.Locations.IncludeLocations -join "; ") } else { "" }
            GrantControls        = if ($policy.GrantControls.BuiltInControls) { ($policy.GrantControls.BuiltInControls -join "; ") } else { "" }
            SessionControls      = if ($policy.SessionControls) { "Yes" } else { "No" }
            PolicyId             = $policy.Id
        }
    }

    $filePath = Join-Path $OutputPath "ConditionalAccessPolicies.csv"
    $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "  Exported $($exportData.Count) CA policies" -Level Info

    return $filePath
}

function Export-DevicesToCSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Devices,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Exporting Devices to CSV..." -Level Info

    $exportData = foreach ($device in $Devices) {
        [PSCustomObject]@{
            DisplayName             = $device.DisplayName
            DeviceId                = $device.DeviceId
            OperatingSystem         = $device.OperatingSystem
            OperatingSystemVersion  = $device.OperatingSystemVersion
            TrustType               = $device.TrustType
            IsCompliant             = $device.IsCompliant
            IsManaged               = $device.IsManaged
            RegisteredOwner         = if ($device.RegisteredOwners -and $device.RegisteredOwners.Count -gt 0) { $device.RegisteredOwners[0].UserPrincipalName } else { "" }
            RegisteredUsers         = if ($device.RegisteredUsers) { ($device.RegisteredUsers.UserPrincipalName -join "; ") } else { "" }
            ApproximateLastSignInDateTime = $device.ApproximateLastSignInDateTime
            CreatedDateTime         = $device.CreatedDateTime
            AccountEnabled          = $device.AccountEnabled
            ProfileType             = $device.ProfileType
            ManagementType          = $device.ManagementType
            EnrollmentType          = $device.EnrollmentType
            DeviceObjectId          = $device.Id
        }
    }

    $filePath = Join-Path $OutputPath "Devices.csv"
    $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "  Exported $($exportData.Count) devices" -Level Info

    return $filePath
}

function Export-GotchaFindingsToCSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Issues,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Exporting Gotcha Findings to CSV..." -Level Info

    $exportData = foreach ($issue in $Issues) {
        [PSCustomObject]@{
            Id                   = $issue.Id
            Name                 = $issue.Name
            Category             = $issue.Category
            Severity             = $issue.Severity
            Description          = $issue.Description
            Impact               = $issue.Impact
            Recommendation       = $issue.Recommendation
            EstimatedEffort      = $issue.EstimatedEffort
            AffectedCount        = $issue.AffectedCount
            AffectedItems        = if ($issue.AffectedItems) { ($issue.AffectedItems | Select-Object -First 10) -join "; " } else { "" }
            RemediationSteps     = if ($issue.RemediationSteps) { ($issue.RemediationSteps -join " | ") } else { "" }
            Tools                = if ($issue.Tools) { ($issue.Tools -join "; ") } else { "" }
            Prerequisites        = if ($issue.Prerequisites) { ($issue.Prerequisites -join "; ") } else { "" }
        }
    }

    $filePath = Join-Path $OutputPath "GotchaFindings.csv"
    $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "  Exported $($exportData.Count) gotcha findings" -Level Info

    return $filePath
}

function Export-DistributionGroupsToCSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Groups,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Exporting Distribution Groups to CSV..." -Level Info

    $exportData = foreach ($group in $Groups) {
        [PSCustomObject]@{
            DisplayName          = $group.DisplayName
            PrimarySmtpAddress   = $group.PrimarySmtpAddress
            Alias                = $group.Alias
            RecipientTypeDetails = $group.RecipientTypeDetails
            MemberCount          = $group.MemberCount
            ManagedBy            = if ($group.ManagedBy) { ($group.ManagedBy -join "; ") } else { "" }
            MemberJoinRestriction = $group.MemberJoinRestriction
            MemberDepartRestriction = $group.MemberDepartRestriction
            RequireSenderAuthenticationEnabled = $group.RequireSenderAuthenticationEnabled
            HiddenFromAddressListsEnabled = $group.HiddenFromAddressListsEnabled
            EmailAddresses       = if ($group.EmailAddresses) { ($group.EmailAddresses -join "; ") } else { "" }
            WhenCreated          = $group.WhenCreated
            ExchangeGuid         = $group.ExchangeGuid
        }
    }

    $filePath = Join-Path $OutputPath "DistributionGroups.csv"
    $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "  Exported $($exportData.Count) distribution groups" -Level Info

    return $filePath
}

function Export-SharedMailboxesToCSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Mailboxes,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Exporting Shared Mailboxes to CSV..." -Level Info

    $exportData = foreach ($mbx in $Mailboxes) {
        [PSCustomObject]@{
            DisplayName              = $mbx.DisplayName
            PrimarySmtpAddress       = $mbx.PrimarySmtpAddress
            Alias                    = $mbx.Alias
            MailboxSizeGB            = if ($mbx.TotalItemSize) { [math]::Round($mbx.TotalItemSize / 1GB, 2) } else { $null }
            ItemCount                = $mbx.ItemCount
            ArchiveEnabled           = $mbx.ArchiveStatus -eq "Active"
            LitigationHoldEnabled    = $mbx.LitigationHoldEnabled
            HiddenFromAddressLists   = $mbx.HiddenFromAddressListsEnabled
            ForwardingAddress        = $mbx.ForwardingAddress
            GrantSendOnBehalfTo      = if ($mbx.GrantSendOnBehalfTo) { ($mbx.GrantSendOnBehalfTo -join "; ") } else { "" }
            EmailAddresses           = if ($mbx.EmailAddresses) { ($mbx.EmailAddresses -join "; ") } else { "" }
            WhenCreated              = $mbx.WhenCreated
            ExchangeGuid             = $mbx.ExchangeGuid
        }
    }

    $filePath = Join-Path $OutputPath "SharedMailboxes.csv"
    $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "  Exported $($exportData.Count) shared mailboxes" -Level Info

    return $filePath
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Export-DiscoveryDataToCSV',
    'Export-UsersToCSV',
    'Export-MailboxesToCSV',
    'Export-MailboxPermissionsToCSV',
    'Export-SharePointSitesToCSV',
    'Export-OneDriveSitesToCSV',
    'Export-TeamsToCSV',
    'Export-GroupsToCSV',
    'Export-LicensesToCSV',
    'Export-ConditionalAccessToCSV',
    'Export-DevicesToCSV',
    'Export-GotchaFindingsToCSV',
    'Export-DistributionGroupsToCSV',
    'Export-SharedMailboxesToCSV'
)
