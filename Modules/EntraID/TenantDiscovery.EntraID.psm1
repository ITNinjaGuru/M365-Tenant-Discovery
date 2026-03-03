#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft Entra ID (Azure AD) Data Collection Module
.DESCRIPTION
    Collects comprehensive Entra ID data including users, groups, devices,
    applications, service principals, conditional access policies, and roles.
    Identifies migration gotchas related to identity management.
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

#region Organization and Tenant Info
function Get-EntraTenantInfo {
    <#
    .SYNOPSIS
        Collects tenant and organization information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Entra ID tenant information..." -Level Info

    try {
        $org = Get-MgOrganization
        $domains = Get-MgDomain
        $context = Get-MgContext

        $tenantInfo = @{
            TenantId              = $org.Id
            DisplayName           = $org.DisplayName
            TenantType            = $org.TenantType
            VerifiedDomains       = @($domains | ForEach-Object {
                @{
                    Name         = $_.Id
                    IsDefault    = $_.IsDefault
                    IsInitial    = $_.IsInitial
                    IsVerified   = $_.IsVerified
                    IsRoot       = $_.IsRoot
                    AuthType     = $_.AuthenticationType
                    Capabilities = $_.SupportedServices
                }
            })
            OnPremisesSyncEnabled = $org.OnPremisesSyncEnabled
            CreatedDateTime       = $org.CreatedDateTime
            Country               = $org.CountryLetterCode
            PreferredLanguage     = $org.PreferredLanguage
            TechnicalContacts     = $org.TechnicalNotificationMails
            PrivacyProfile        = $org.PrivacyProfile
            AssignedPlans         = $org.AssignedPlans | ForEach-Object {
                @{
                    Service   = $_.Service
                    Plan      = $_.ServicePlanId
                    Status    = $_.CapabilityStatus
                }
            }
        }

        # Check for gotchas
        $federatedDomains = $domains | Where-Object { $_.AuthenticationType -eq "Federated" }
        if ($federatedDomains) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Federated Domains Detected" `
                -Description "Found $($federatedDomains.Count) federated domain(s). These require special handling during migration as they are tied to ADFS or third-party IdP." `
                -Severity "High" `
                -Recommendation "Document federation configuration. Plan for federation cutover or conversion to managed domains. Consider staged migration approach." `
                -AffectedObjects @($federatedDomains.Id) `
                -MigrationPhase "Pre-Migration"
        }

        if ($org.OnPremisesSyncEnabled) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Directory Synchronization Enabled" `
                -Description "On-premises directory sync is enabled. Objects synced from AD cannot be modified in cloud and have special migration considerations." `
                -Severity "High" `
                -Recommendation "Identify all synced vs cloud-only objects. Plan for source of authority changes. Consider AAD Connect migration approach." `
                -MigrationPhase "Pre-Migration"
        }

        Add-CollectedData -Category "EntraID" -SubCategory "TenantInfo" -Data $tenantInfo
        Write-Log -Message "Tenant information collected successfully" -Level Success

        return $tenantInfo
    }
    catch {
        Write-Log -Message "Failed to collect tenant info: $_" -Level Error
        throw
    }
}
#endregion

#region User Collection
function Get-EntraUsers {
    <#
    .SYNOPSIS
        Collects comprehensive user information from Entra ID
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$MaxUsers = 50000
    )

    Write-Log -Message "Collecting Entra ID users..." -Level Info

    try {
        $users = @()
        $properties = @(
            "Id", "DisplayName", "UserPrincipalName", "Mail", "MailNickname",
            "OnPremisesSyncEnabled", "OnPremisesImmutableId", "OnPremisesSamAccountName",
            "OnPremisesDistinguishedName", "OnPremisesSecurityIdentifier", "OnPremisesUserPrincipalName",
            "OnPremisesLastSyncDateTime", "OnPremisesDomainName",
            "UserType", "AccountEnabled", "CreatedDateTime", "LastPasswordChangeDateTime",
            "PasswordPolicies", "ProxyAddresses", "OtherMails",
            "UsageLocation", "PreferredLanguage", "JobTitle", "Department", "CompanyName",
            "OfficeLocation", "City", "State", "Country", "PostalCode",
            "MobilePhone", "BusinessPhones", "FaxNumber",
            "AssignedLicenses", "AssignedPlans", "ProvisionedPlans",
            "Manager", "DirectReports",
            "MemberOf", "TransitiveMemberOf",
            "LicenseAssignmentStates", "SignInSessionsValidFromDateTime",
            "ExternalUserState", "ExternalUserStateChangeDateTime",
            "Identities", "AuthenticationMethods"
        )

        $selectString = $properties -join ","

        # Get all users with pagination
        $uri = "https://graph.microsoft.com/v1.0/users?`$select=$selectString&`$top=999"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri

        $users += $response.value

        while ($response.'@odata.nextLink' -and $users.Count -lt $MaxUsers) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $response.'@odata.nextLink'
            $users += $response.value
            Write-Progress -Activity "Collecting Users" -Status "Retrieved $($users.Count) users..."
        }

        Write-Progress -Activity "Collecting Users" -Completed

        # Analyze users
        $analysis = @{
            TotalUsers       = $users.Count
            EnabledUsers     = ($users | Where-Object { $_.AccountEnabled }).Count
            DisabledUsers    = ($users | Where-Object { -not $_.AccountEnabled }).Count
            CloudOnlyUsers   = ($users | Where-Object { -not $_.OnPremisesSyncEnabled }).Count
            SyncedUsers      = ($users | Where-Object { $_.OnPremisesSyncEnabled }).Count
            GuestUsers       = ($users | Where-Object { $_.UserType -eq "Guest" }).Count
            MemberUsers      = ($users | Where-Object { $_.UserType -eq "Member" }).Count
            LicensedUsers    = ($users | Where-Object { $_.AssignedLicenses.Count -gt 0 }).Count
            UnlicensedUsers  = ($users | Where-Object { $_.AssignedLicenses.Count -eq 0 }).Count
        }

        # Collect detailed user data with gotcha detection
        $processedUsers = foreach ($user in $users) {
            $userData = @{
                Id                        = $user.Id
                DisplayName               = $user.DisplayName
                UserPrincipalName         = $user.UserPrincipalName
                Mail                      = $user.Mail
                MailNickname              = $user.MailNickname
                OnPremisesSyncEnabled     = $user.OnPremisesSyncEnabled
                OnPremisesImmutableId     = $user.OnPremisesImmutableId
                OnPremisesSamAccountName  = $user.OnPremisesSamAccountName
                OnPremisesDN              = $user.OnPremisesDistinguishedName
                OnPremisesSID             = $user.OnPremisesSecurityIdentifier
                OnPremisesUPN             = $user.OnPremisesUserPrincipalName
                OnPremisesLastSync        = $user.OnPremisesLastSyncDateTime
                OnPremisesDomain          = $user.OnPremisesDomainName
                UserType                  = $user.UserType
                AccountEnabled            = $user.AccountEnabled
                CreatedDateTime           = $user.CreatedDateTime
                LastPasswordChange        = $user.LastPasswordChangeDateTime
                PasswordPolicies          = $user.PasswordPolicies
                ProxyAddresses            = $user.ProxyAddresses
                OtherMails                = $user.OtherMails
                UsageLocation             = $user.UsageLocation
                JobTitle                  = $user.JobTitle
                Department                = $user.Department
                CompanyName               = $user.CompanyName
                AssignedLicenses          = $user.AssignedLicenses
                ExternalUserState         = $user.ExternalUserState
                Identities                = $user.Identities
            }
            $userData
        }

        # Detect gotchas

        # Users without usage location
        $noUsageLocation = $users | Where-Object {
            $_.AssignedLicenses.Count -gt 0 -and [string]::IsNullOrEmpty($_.UsageLocation)
        }
        if ($noUsageLocation.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Licensed Users Without Usage Location" `
                -Description "Found $($noUsageLocation.Count) licensed users without a usage location set. This may cause licensing issues in target tenant." `
                -Severity "Medium" `
                -Recommendation "Set usage location for all users before migration. Consider automation to set default location." `
                -AffectedObjects @($noUsageLocation.UserPrincipalName) `
                -MigrationPhase "Pre-Migration"
        }

        # Users with complex proxy addresses
        $multiProxyUsers = $users | Where-Object { $_.ProxyAddresses.Count -gt 5 }
        if ($multiProxyUsers.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Users with Multiple Proxy Addresses" `
                -Description "Found $($multiProxyUsers.Count) users with more than 5 proxy addresses. These need careful mapping during migration." `
                -Severity "Medium" `
                -Recommendation "Document all proxy addresses. Plan for address migration and conflict resolution." `
                -AffectedCount $multiProxyUsers.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Guest users
        $guestUsers = $users | Where-Object { $_.UserType -eq "Guest" }
        if ($guestUsers.Count -gt 0) {
            # Check for stale guests
            $staleGuests = $guestUsers | Where-Object {
                $_.ExternalUserState -eq "PendingAcceptance"
            }

            Add-MigrationGotcha -Category "EntraID" `
                -Title "Guest Users Present" `
                -Description "Found $($guestUsers.Count) guest users. $($staleGuests.Count) are pending acceptance. Guest users need reinvitation in target tenant." `
                -Severity "Medium" `
                -Recommendation "Review guest access. Plan for guest reinvitation process. Consider cleaning up stale invitations first." `
                -AffectedCount $guestUsers.Count `
                -MigrationPhase "Post-Migration"
        }

        # Disabled synced users
        $disabledSynced = $users | Where-Object {
            -not $_.AccountEnabled -and $_.OnPremisesSyncEnabled
        }
        if ($disabledSynced.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Disabled Synced Users" `
                -Description "Found $($disabledSynced.Count) disabled users that are synced from on-premises. These may represent terminated employees or service accounts." `
                -Severity "Low" `
                -Recommendation "Review if disabled accounts should be migrated. Consider cleanup before migration." `
                -AffectedCount $disabledSynced.Count `
                -MigrationPhase "Pre-Migration"
        }

        # UPN suffix analysis
        $upnSuffixes = $users | ForEach-Object {
            if ($_.UserPrincipalName -match '@(.+)$') { $Matches[1] }
        } | Sort-Object -Unique

        $result = @{
            Users    = $processedUsers
            Analysis = $analysis
            UPNSuffixes = $upnSuffixes
        }

        Add-CollectedData -Category "EntraID" -SubCategory "Users" -Data $result
        Write-Log -Message "Collected $($users.Count) users" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect users: $_" -Level Error
        throw
    }
}
#endregion

#region Group Collection
function Get-EntraGroups {
    <#
    .SYNOPSIS
        Collects comprehensive group information from Entra ID
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Entra ID groups..." -Level Info

    try {
        $groups = @()
        $properties = @(
            "Id", "DisplayName", "Description", "Mail", "MailEnabled", "MailNickname",
            "SecurityEnabled", "GroupTypes", "MembershipRule", "MembershipRuleProcessingState",
            "OnPremisesSyncEnabled", "OnPremisesSecurityIdentifier", "OnPremisesSamAccountName",
            "OnPremisesLastSyncDateTime", "OnPremisesDomainName", "OnPremisesNetBiosName",
            "ProxyAddresses", "IsAssignableToRole", "Visibility",
            "CreatedDateTime", "RenewedDateTime", "ExpirationDateTime",
            "Classification", "Theme", "ResourceProvisioningOptions"
        )

        $selectString = $properties -join ","
        $uri = "https://graph.microsoft.com/v1.0/groups?`$select=$selectString&`$top=999"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri

        $groups += $response.value

        while ($response.'@odata.nextLink') {
            $response = Invoke-MgGraphRequest -Method GET -Uri $response.'@odata.nextLink'
            $groups += $response.value
            Write-Progress -Activity "Collecting Groups" -Status "Retrieved $($groups.Count) groups..."
        }

        Write-Progress -Activity "Collecting Groups" -Completed

        # Analyze groups
        $m365Groups = $groups | Where-Object { $_.GroupTypes -contains "Unified" }
        $securityGroups = $groups | Where-Object { $_.SecurityEnabled -and -not ($_.GroupTypes -contains "Unified") }
        $distributionLists = $groups | Where-Object { $_.MailEnabled -and -not $_.SecurityEnabled }
        $mailEnabledSecurity = $groups | Where-Object { $_.MailEnabled -and $_.SecurityEnabled -and -not ($_.GroupTypes -contains "Unified") }
        $dynamicGroups = $groups | Where-Object { $_.GroupTypes -contains "DynamicMembership" }
        $syncedGroups = $groups | Where-Object { $_.OnPremisesSyncEnabled }
        $roleAssignableGroups = $groups | Where-Object { $_.IsAssignableToRole }

        $analysis = @{
            TotalGroups          = $groups.Count
            M365Groups           = $m365Groups.Count
            SecurityGroups       = $securityGroups.Count
            DistributionLists    = $distributionLists.Count
            MailEnabledSecurity  = $mailEnabledSecurity.Count
            DynamicGroups        = $dynamicGroups.Count
            SyncedGroups         = $syncedGroups.Count
            CloudOnlyGroups      = ($groups | Where-Object { -not $_.OnPremisesSyncEnabled }).Count
            RoleAssignableGroups = $roleAssignableGroups.Count
        }

        # Detect gotchas

        # Mail-enabled security groups
        if ($mailEnabledSecurity.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Mail-Enabled Security Groups" `
                -Description "Found $($mailEnabledSecurity.Count) mail-enabled security groups. These have complex migration requirements as they exist in both AD and Exchange." `
                -Severity "High" `
                -Recommendation "Document group membership. Plan for recreation in target. Consider if these should remain mail-enabled or split into separate objects." `
                -AffectedObjects @($mailEnabledSecurity.DisplayName) `
                -MigrationPhase "Pre-Migration"
        }

        # Dynamic groups with complex rules
        $complexDynamic = $dynamicGroups | Where-Object {
            $_.MembershipRule -and $_.MembershipRule.Length -gt 200
        }
        if ($complexDynamic.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Complex Dynamic Group Rules" `
                -Description "Found $($complexDynamic.Count) dynamic groups with complex membership rules. These rules may reference attributes that differ in target tenant." `
                -Severity "High" `
                -Recommendation "Review and document all dynamic membership rules. Test rules in target tenant. Consider if extension attributes are mapped correctly." `
                -AffectedObjects @($complexDynamic.DisplayName) `
                -MigrationPhase "Post-Migration"
        }

        # Synced groups
        if ($syncedGroups.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "On-Premises Synced Groups" `
                -Description "Found $($syncedGroups.Count) groups synced from on-premises AD. These cannot be modified in cloud and require source of authority consideration." `
                -Severity "Medium" `
                -Recommendation "Plan for group migration approach. Consider if sync scope needs adjustment. Document nested group structures." `
                -AffectedCount $syncedGroups.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Role-assignable groups
        if ($roleAssignableGroups.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Role-Assignable Groups" `
                -Description "Found $($roleAssignableGroups.Count) groups that are assignable to Entra ID roles. These provide privileged access and need security review." `
                -Severity "High" `
                -Recommendation "Document role assignments. Review if these should exist in target tenant. Plan for recreation with proper governance." `
                -AffectedObjects @($roleAssignableGroups.DisplayName) `
                -MigrationPhase "Post-Migration"
        }

        # Get group membership counts
        $groupDetails = foreach ($group in $groups) {
            @{
                Id                      = $group.Id
                DisplayName             = $group.DisplayName
                Description             = $group.Description
                Mail                    = $group.Mail
                MailEnabled             = $group.MailEnabled
                SecurityEnabled         = $group.SecurityEnabled
                GroupTypes              = $group.GroupTypes
                MembershipRule          = $group.MembershipRule
                OnPremisesSyncEnabled   = $group.OnPremisesSyncEnabled
                OnPremisesSID           = $group.OnPremisesSecurityIdentifier
                ProxyAddresses          = $group.ProxyAddresses
                IsAssignableToRole      = $group.IsAssignableToRole
                Visibility              = $group.Visibility
                CreatedDateTime         = $group.CreatedDateTime
                Classification          = $group.Classification
            }
        }

        $result = @{
            Groups   = $groupDetails
            Analysis = $analysis
        }

        Add-CollectedData -Category "EntraID" -SubCategory "Groups" -Data $result
        Write-Log -Message "Collected $($groups.Count) groups" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect groups: $_" -Level Error
        throw
    }
}
#endregion

#region Device Collection
function Get-EntraDevices {
    <#
    .SYNOPSIS
        Collects device information from Entra ID
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Entra ID devices..." -Level Info

    try {
        $devices = @()
        $properties = @(
            "Id", "DisplayName", "DeviceId", "OperatingSystem", "OperatingSystemVersion",
            "TrustType", "IsCompliant", "IsManaged", "IsRooted",
            "ManagementType", "Manufacturer", "Model",
            "ProfileType", "RegistrationDateTime", "ApproximateLastSignInDateTime",
            "OnPremisesSyncEnabled", "OnPremisesSecurityIdentifier", "OnPremisesLastSyncDateTime",
            "AccountEnabled", "AlternativeSecurityIds", "DeviceOwnership"
        )

        $selectString = $properties -join ","
        $uri = "https://graph.microsoft.com/v1.0/devices?`$select=$selectString&`$top=999"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri

        $devices += $response.value

        while ($response.'@odata.nextLink') {
            $response = Invoke-MgGraphRequest -Method GET -Uri $response.'@odata.nextLink'
            $devices += $response.value
            Write-Progress -Activity "Collecting Devices" -Status "Retrieved $($devices.Count) devices..."
        }

        Write-Progress -Activity "Collecting Devices" -Completed

        # Analyze devices
        $hybridJoined = $devices | Where-Object { $_.TrustType -eq "ServerAd" }
        $azureADJoined = $devices | Where-Object { $_.TrustType -eq "AzureAd" }
        $registered = $devices | Where-Object { $_.TrustType -eq "Workplace" }

        $windowsDevices = $devices | Where-Object { $_.OperatingSystem -like "*Windows*" }
        $macDevices = $devices | Where-Object { $_.OperatingSystem -like "*Mac*" }
        $iosDevices = $devices | Where-Object { $_.OperatingSystem -like "*iOS*" }
        $androidDevices = $devices | Where-Object { $_.OperatingSystem -like "*Android*" }
        $linuxDevices = $devices | Where-Object { $_.OperatingSystem -like "*Linux*" }

        $compliantDevices = $devices | Where-Object { $_.IsCompliant }
        $managedDevices = $devices | Where-Object { $_.IsManaged }

        $staleDevices = $devices | Where-Object {
            $_.ApproximateLastSignInDateTime -and
            ([datetime]$_.ApproximateLastSignInDateTime) -lt (Get-Date).AddDays(-90)
        }

        $analysis = @{
            TotalDevices     = $devices.Count
            HybridJoined     = $hybridJoined.Count
            AzureADJoined    = $azureADJoined.Count
            Registered       = $registered.Count
            Windows          = $windowsDevices.Count
            Mac              = $macDevices.Count
            iOS              = $iosDevices.Count
            Android          = $androidDevices.Count
            Linux            = $linuxDevices.Count
            Compliant        = $compliantDevices.Count
            Managed          = $managedDevices.Count
            Stale90Days      = $staleDevices.Count
            SyncedDevices    = ($devices | Where-Object { $_.OnPremisesSyncEnabled }).Count
        }

        # Detect gotchas

        # Hybrid Azure AD Joined devices
        if ($hybridJoined.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Hybrid Azure AD Joined Devices" `
                -Description "Found $($hybridJoined.Count) Hybrid Azure AD Joined devices. These are domain-joined and require careful migration planning." `
                -Severity "Critical" `
                -Recommendation "Plan for device registration cutover. Consider impact on Conditional Access. Devices may need to be unjoined and rejoined to target tenant." `
                -AffectedCount $hybridJoined.Count `
                -MigrationPhase "Post-Migration" `
                -AdditionalData @{
                    Note = "Hybrid join requires AAD Connect device sync. Users may need to re-register devices."
                }
        }

        # Stale devices
        if ($staleDevices.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Stale Device Objects" `
                -Description "Found $($staleDevices.Count) devices with no sign-in activity in 90+ days. These may be obsolete." `
                -Severity "Low" `
                -Recommendation "Review and cleanup stale devices before migration. This reduces migration scope and improves security posture." `
                -AffectedCount $staleDevices.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Non-compliant devices
        $nonCompliant = $devices | Where-Object { $_.IsCompliant -eq $false }
        if ($nonCompliant.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Non-Compliant Devices" `
                -Description "Found $($nonCompliant.Count) non-compliant devices. Compliance policies will need recreation in target tenant." `
                -Severity "Medium" `
                -Recommendation "Document compliance policies. Ensure policies are recreated before device migration." `
                -AffectedCount $nonCompliant.Count `
                -MigrationPhase "Pre-Migration"
        }

        $deviceDetails = foreach ($device in $devices) {
            @{
                Id                        = $device.Id
                DisplayName               = $device.DisplayName
                DeviceId                  = $device.DeviceId
                OperatingSystem           = $device.OperatingSystem
                OperatingSystemVersion    = $device.OperatingSystemVersion
                TrustType                 = $device.TrustType
                IsCompliant               = $device.IsCompliant
                IsManaged                 = $device.IsManaged
                ManagementType            = $device.ManagementType
                Manufacturer              = $device.Manufacturer
                Model                     = $device.Model
                RegistrationDateTime      = $device.RegistrationDateTime
                LastSignIn                = $device.ApproximateLastSignInDateTime
                OnPremisesSyncEnabled     = $device.OnPremisesSyncEnabled
                AccountEnabled            = $device.AccountEnabled
                DeviceOwnership           = $device.DeviceOwnership
            }
        }

        $result = @{
            Devices  = $deviceDetails
            Analysis = $analysis
        }

        Add-CollectedData -Category "EntraID" -SubCategory "Devices" -Data $result
        Write-Log -Message "Collected $($devices.Count) devices" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect devices: $_" -Level Error
        throw
    }
}
#endregion

#region Application Collection
function Get-EntraApplications {
    <#
    .SYNOPSIS
        Collects application registrations and enterprise applications
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Entra ID applications..." -Level Info

    try {
        # Get App Registrations
        $appRegistrations = @()
        $uri = "https://graph.microsoft.com/v1.0/applications?`$top=999"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri

        $appRegistrations += $response.value

        while ($response.'@odata.nextLink') {
            $response = Invoke-MgGraphRequest -Method GET -Uri $response.'@odata.nextLink'
            $appRegistrations += $response.value
        }

        # Get Service Principals (Enterprise Apps)
        $servicePrincipals = @()
        $uri = "https://graph.microsoft.com/v1.0/servicePrincipals?`$top=999"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri

        $servicePrincipals += $response.value

        while ($response.'@odata.nextLink') {
            $response = Invoke-MgGraphRequest -Method GET -Uri $response.'@odata.nextLink'
            $servicePrincipals += $response.value
        }

        # Analyze applications
        $customApps = $appRegistrations | Where-Object { $_.PublisherDomain }
        $multiTenantApps = $appRegistrations | Where-Object { $_.SignInAudience -ne "AzureADMyOrg" }
        $appsWithSecrets = $appRegistrations | Where-Object { $_.PasswordCredentials.Count -gt 0 }
        $appsWithCerts = $appRegistrations | Where-Object { $_.KeyCredentials.Count -gt 0 }

        # Check for expiring credentials
        $now = Get-Date
        $expiringSecrets = foreach ($app in $appRegistrations) {
            foreach ($secret in $app.PasswordCredentials) {
                if ($secret.EndDateTime -and ([datetime]$secret.EndDateTime) -lt $now.AddDays(90)) {
                    @{
                        AppName    = $app.DisplayName
                        AppId      = $app.AppId
                        SecretId   = $secret.KeyId
                        ExpiryDate = $secret.EndDateTime
                        DaysLeft   = (([datetime]$secret.EndDateTime) - $now).Days
                    }
                }
            }
        }

        $expiringCerts = foreach ($app in $appRegistrations) {
            foreach ($cert in $app.KeyCredentials) {
                if ($cert.EndDateTime -and ([datetime]$cert.EndDateTime) -lt $now.AddDays(90)) {
                    @{
                        AppName    = $app.DisplayName
                        AppId      = $app.AppId
                        CertId     = $cert.KeyId
                        ExpiryDate = $cert.EndDateTime
                        DaysLeft   = (([datetime]$cert.EndDateTime) - $now).Days
                    }
                }
            }
        }

        $analysis = @{
            TotalAppRegistrations   = $appRegistrations.Count
            TotalServicePrincipals  = $servicePrincipals.Count
            CustomApps              = $customApps.Count
            MultiTenantApps         = $multiTenantApps.Count
            AppsWithSecrets         = $appsWithSecrets.Count
            AppsWithCertificates    = $appsWithCerts.Count
            ExpiringSecrets90Days   = ($expiringSecrets | Measure-Object).Count
            ExpiringCerts90Days     = ($expiringCerts | Measure-Object).Count
            MicrosoftApps           = ($servicePrincipals | Where-Object { $_.AppOwnerOrganizationId -eq "f8cdef31-a31e-4b4a-93e4-5f571e91255a" }).Count
            ThirdPartyApps          = ($servicePrincipals | Where-Object { $_.AppOwnerOrganizationId -and $_.AppOwnerOrganizationId -ne "f8cdef31-a31e-4b4a-93e4-5f571e91255a" }).Count
        }

        # Detect gotchas

        # Expiring credentials
        $expiredOrExpiring = $expiringSecrets + $expiringCerts
        if ($expiredOrExpiring.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Applications with Expiring Credentials" `
                -Description "Found $($expiredOrExpiring.Count) application credentials expiring within 90 days. These need renewal before or during migration." `
                -Severity "High" `
                -Recommendation "Renew credentials before migration. Document all app credentials and notify app owners." `
                -AffectedObjects @($expiredOrExpiring | ForEach-Object { $_.AppName } | Sort-Object -Unique) `
                -MigrationPhase "Pre-Migration"
        }

        # Apps with high-privilege permissions
        $highPrivApps = foreach ($sp in $servicePrincipals) {
            try {
                $permissions = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)/appRoleAssignments" -ErrorAction SilentlyContinue
                $dangerousRoles = @(
                    "RoleManagement.ReadWrite.Directory",
                    "Directory.ReadWrite.All",
                    "Application.ReadWrite.All",
                    "Mail.ReadWrite",
                    "Mail.Send"
                )
                if ($permissions.value | Where-Object { $dangerousRoles -contains $_.appRoleId }) {
                    @{
                        AppName = $sp.DisplayName
                        AppId   = $sp.AppId
                    }
                }
            }
            catch { }
        }

        if ($highPrivApps.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "High-Privilege Applications" `
                -Description "Found applications with high-privilege API permissions. These require security review before recreation in target." `
                -Severity "High" `
                -Recommendation "Document all API permissions. Review if permissions are still required. Plan for consent in target tenant." `
                -AffectedCount $highPrivApps.Count `
                -MigrationPhase "Pre-Migration"
        }

        # SAML/SSO configured apps
        $ssoApps = $servicePrincipals | Where-Object {
            $_.PreferredSingleSignOnMode -in @("saml", "password", "oidc")
        }

        if ($ssoApps.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "SSO Configured Applications" `
                -Description "Found $($ssoApps.Count) applications with SSO configuration (SAML, OIDC, or Password). These need reconfiguration in target tenant." `
                -Severity "High" `
                -Recommendation "Document SSO configurations including certificates, claim mappings, and reply URLs. Plan for SSO reconfiguration and testing." `
                -AffectedObjects @($ssoApps.DisplayName) `
                -MigrationPhase "Post-Migration"
        }

        $result = @{
            AppRegistrations   = $appRegistrations
            ServicePrincipals  = $servicePrincipals
            ExpiringSecrets    = $expiringSecrets
            ExpiringCerts      = $expiringCerts
            Analysis           = $analysis
        }

        Add-CollectedData -Category "EntraID" -SubCategory "Applications" -Data $result
        Write-Log -Message "Collected $($appRegistrations.Count) app registrations and $($servicePrincipals.Count) service principals" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect applications: $_" -Level Error
        throw
    }
}
#endregion

#region Conditional Access
function Get-EntraConditionalAccess {
    <#
    .SYNOPSIS
        Collects Conditional Access policies
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Conditional Access policies..." -Level Info

    try {
        $policies = @()
        $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri

        $policies = $response.value

        # Get named locations
        $namedLocations = @()
        $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $namedLocations = $response.value

        # Analyze policies
        $enabledPolicies = $policies | Where-Object { $_.State -eq "enabled" }
        $reportOnlyPolicies = $policies | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }
        $disabledPolicies = $policies | Where-Object { $_.State -eq "disabled" }

        $mfaPolicies = $policies | Where-Object {
            $_.GrantControls.BuiltInControls -contains "mfa"
        }

        $blockPolicies = $policies | Where-Object {
            $_.GrantControls.BuiltInControls -contains "block"
        }

        $devicePolicies = $policies | Where-Object {
            $_.GrantControls.BuiltInControls -contains "compliantDevice" -or
            $_.GrantControls.BuiltInControls -contains "domainJoinedDevice"
        }

        $analysis = @{
            TotalPolicies         = $policies.Count
            EnabledPolicies       = $enabledPolicies.Count
            ReportOnlyPolicies    = $reportOnlyPolicies.Count
            DisabledPolicies      = $disabledPolicies.Count
            MFAPolicies           = $mfaPolicies.Count
            BlockPolicies         = $blockPolicies.Count
            DeviceCompliancePolicies = $devicePolicies.Count
            NamedLocations        = $namedLocations.Count
        }

        # Detect gotchas

        # Policies targeting specific groups
        $groupTargetedPolicies = $policies | Where-Object {
            $_.Conditions.Users.IncludeGroups.Count -gt 0 -or
            $_.Conditions.Users.ExcludeGroups.Count -gt 0
        }

        if ($groupTargetedPolicies.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Conditional Access Policies with Group Targeting" `
                -Description "Found $($groupTargetedPolicies.Count) CA policies targeting specific groups. Group IDs will differ in target tenant." `
                -Severity "High" `
                -Recommendation "Document all group references in CA policies. Plan for group ID mapping in target tenant. Use group names for reconstruction." `
                -AffectedObjects @($groupTargetedPolicies.DisplayName) `
                -MigrationPhase "Post-Migration"
        }

        # Policies with named locations
        $locationPolicies = $policies | Where-Object {
            $_.Conditions.Locations.IncludeLocations.Count -gt 0 -or
            $_.Conditions.Locations.ExcludeLocations.Count -gt 0
        }

        if ($locationPolicies.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Conditional Access Policies with Location Conditions" `
                -Description "Found $($locationPolicies.Count) CA policies using named locations. Named locations must be recreated in target tenant first." `
                -Severity "Medium" `
                -Recommendation "Export named location definitions. Recreate named locations in target before CA policies." `
                -AffectedObjects @($locationPolicies.DisplayName) `
                -MigrationPhase "Pre-Migration"
        }

        # Policies targeting applications
        $appPolicies = $policies | Where-Object {
            $_.Conditions.Applications.IncludeApplications -and
            $_.Conditions.Applications.IncludeApplications[0] -ne "All"
        }

        if ($appPolicies.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Conditional Access Policies Targeting Specific Apps" `
                -Description "Found $($appPolicies.Count) CA policies targeting specific applications. App IDs will differ for custom apps in target tenant." `
                -Severity "Medium" `
                -Recommendation "Document application targets. Map app IDs between tenants. Microsoft apps use same IDs across tenants." `
                -AffectedObjects @($appPolicies.DisplayName) `
                -MigrationPhase "Post-Migration"
        }

        # Complex grant controls
        $complexGrants = $policies | Where-Object {
            ($_.GrantControls.BuiltInControls.Count -gt 2) -or
            $_.GrantControls.AuthenticationStrength -or
            $_.GrantControls.CustomAuthenticationFactors
        }

        if ($complexGrants.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Complex Conditional Access Grant Controls" `
                -Description "Found $($complexGrants.Count) policies with complex grant controls. These may include authentication strengths or custom factors." `
                -Severity "Medium" `
                -Recommendation "Document all grant control configurations. Ensure authentication strengths exist in target tenant." `
                -AffectedObjects @($complexGrants.DisplayName) `
                -MigrationPhase "Post-Migration"
        }

        $result = @{
            Policies       = $policies
            NamedLocations = $namedLocations
            Analysis       = $analysis
        }

        Add-CollectedData -Category "EntraID" -SubCategory "ConditionalAccess" -Data $result
        Write-Log -Message "Collected $($policies.Count) Conditional Access policies" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Conditional Access: $_" -Level Error
        throw
    }
}
#endregion

#region Role Management
function Get-EntraRoles {
    <#
    .SYNOPSIS
        Collects Entra ID role assignments and definitions
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Entra ID roles..." -Level Info

    try {
        # Get role definitions
        $roleDefinitions = @()
        $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $roleDefinitions = $response.value

        # Get role assignments
        $roleAssignments = @()
        $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$expand=principal"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $roleAssignments = $response.value

        while ($response.'@odata.nextLink') {
            $response = Invoke-MgGraphRequest -Method GET -Uri $response.'@odata.nextLink'
            $roleAssignments += $response.value
        }

        # Get custom roles
        $customRoles = $roleDefinitions | Where-Object { $_.IsBuiltIn -eq $false }

        # Get privileged role assignments (Global Admin, etc.)
        $privilegedRoles = @(
            "62e90394-69f5-4237-9190-012177145e10", # Global Administrator
            "e8611ab8-c189-46e8-94e1-60213ab1f814", # Privileged Role Administrator
            "9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3", # Application Administrator
            "158c047a-c907-4556-b7ef-446551a6b5f7", # Cloud Application Administrator
            "b1be1c3e-b65d-4f19-8427-f6fa0d97feb9", # Conditional Access Administrator
            "29232cdf-9323-42fd-ade2-1d097af3e4de", # Exchange Administrator
            "f28a1f50-f6e7-4571-818b-6a12f2af6b6c", # SharePoint Administrator
            "194ae4cb-b126-40b2-bd5b-6091b380977d", # Security Administrator
            "7be44c8a-adaf-4e2a-84d6-ab2649e08a13"  # Privileged Authentication Administrator
        )

        $privilegedAssignments = $roleAssignments | Where-Object {
            $_.RoleDefinitionId -in $privilegedRoles
        }

        # Get users with multiple privileged roles
        $multiRoleUsers = $roleAssignments |
            Group-Object -Property { $_.Principal.Id } |
            Where-Object { $_.Count -gt 1 } |
            ForEach-Object {
                @{
                    PrincipalId   = $_.Name
                    PrincipalName = $_.Group[0].Principal.DisplayName
                    RoleCount     = $_.Count
                    Roles         = $_.Group.RoleDefinitionId
                }
            }

        $analysis = @{
            TotalRoleDefinitions      = $roleDefinitions.Count
            CustomRoleDefinitions     = $customRoles.Count
            TotalRoleAssignments      = $roleAssignments.Count
            PrivilegedRoleAssignments = $privilegedAssignments.Count
            UsersWithMultipleRoles    = ($multiRoleUsers | Measure-Object).Count
        }

        # Detect gotchas

        # Custom roles
        if ($customRoles.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Custom Role Definitions" `
                -Description "Found $($customRoles.Count) custom role definitions. These must be recreated in target tenant." `
                -Severity "Medium" `
                -Recommendation "Export custom role definitions. Recreate in target tenant. Update role assignments after recreation." `
                -AffectedObjects @($customRoles.DisplayName) `
                -MigrationPhase "Pre-Migration"
        }

        # Global Admins
        $globalAdminRole = "62e90394-69f5-4237-9190-012177145e10"
        $globalAdmins = $roleAssignments | Where-Object { $_.RoleDefinitionId -eq $globalAdminRole }

        if ($globalAdmins.Count -gt 5) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Excessive Global Administrators" `
                -Description "Found $($globalAdmins.Count) Global Administrators. Best practice recommends 2-4 emergency access accounts only." `
                -Severity "Medium" `
                -Recommendation "Review Global Admin assignments. Consider least-privilege roles. Plan for role cleanup during migration." `
                -AffectedCount $globalAdmins.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Users with multiple privileged roles
        $multiPrivUsers = $multiRoleUsers | Where-Object {
            $_.Roles | Where-Object { $_ -in $privilegedRoles }
        }

        if ($multiPrivUsers.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Users with Multiple Privileged Roles" `
                -Description "Found $($multiPrivUsers.Count) users with multiple privileged role assignments. This may violate least-privilege principle." `
                -Severity "Low" `
                -Recommendation "Review role assignments. Consider consolidation or separation of duties." `
                -AffectedCount $multiPrivUsers.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            RoleDefinitions  = $roleDefinitions
            RoleAssignments  = $roleAssignments
            CustomRoles      = $customRoles
            MultiRoleUsers   = $multiRoleUsers
            Analysis         = $analysis
        }

        Add-CollectedData -Category "EntraID" -SubCategory "Roles" -Data $result
        Write-Log -Message "Collected $($roleDefinitions.Count) role definitions and $($roleAssignments.Count) assignments" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect roles: $_" -Level Error
        throw
    }
}
#endregion

#region Authentication Methods
function Get-EntraAuthenticationMethods {
    <#
    .SYNOPSIS
        Collects authentication method configurations
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting authentication method configurations..." -Level Info

    try {
        # Get authentication method policies
        $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy"
        $authMethodsPolicy = Invoke-MgGraphRequest -Method GET -Uri $uri

        # Get legacy MFA settings
        $mfaSettings = $null
        try {
            # This requires Azure AD Premium
            $uri = "https://graph.microsoft.com/beta/policies/authenticationMethodsPolicy/authenticationMethodConfigurations"
            $mfaSettings = Invoke-MgGraphRequest -Method GET -Uri $uri
        }
        catch {
            Write-Log -Message "Could not retrieve detailed MFA settings (may require premium license)" -Level Warning
        }

        # Analyze enabled methods
        $enabledMethods = $authMethodsPolicy.AuthenticationMethodConfigurations |
            Where-Object { $_.State -eq "enabled" }

        $analysis = @{
            EnabledMethods = $enabledMethods | ForEach-Object { $_.Id }
            PolicyMigrationState = $authMethodsPolicy.PolicyMigrationState
            RegistrationEnforcement = $authMethodsPolicy.RegistrationEnforcement
        }

        # Detect gotchas

        # Legacy MFA portal usage
        if ($authMethodsPolicy.PolicyMigrationState -ne "migrationComplete") {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Legacy MFA Portal Still in Use" `
                -Description "Authentication methods policy migration is not complete. Some MFA settings may still be in legacy portal." `
                -Severity "Medium" `
                -Recommendation "Complete migration to Authentication Methods policies before tenant migration. Document legacy MFA settings." `
                -MigrationPhase "Pre-Migration"
        }

        # Check for third-party authenticator apps
        $fido2Method = $authMethodsPolicy.AuthenticationMethodConfigurations |
            Where-Object { $_.Id -eq "fido2" -and $_.State -eq "enabled" }

        if ($fido2Method) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "FIDO2 Security Keys Enabled" `
                -Description "FIDO2 security keys are enabled. Users with registered keys will need to re-register in target tenant." `
                -Severity "Medium" `
                -Recommendation "Identify users with FIDO2 keys. Plan for key re-registration post-migration." `
                -MigrationPhase "Post-Migration"
        }

        $result = @{
            AuthenticationMethodsPolicy = $authMethodsPolicy
            Analysis = $analysis
        }

        Add-CollectedData -Category "EntraID" -SubCategory "AuthenticationMethods" -Data $result
        Write-Log -Message "Collected authentication method configurations" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect authentication methods: $_" -Level Error
        throw
    }
}
#endregion

#region Licensing
function Get-EntraLicensing {
    <#
    .SYNOPSIS
        Collects license information and subscription details
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting licensing information..." -Level Info

    try {
        # Get subscribed SKUs
        $uri = "https://graph.microsoft.com/v1.0/subscribedSkus"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $subscribedSkus = $response.value

        # Calculate license usage
        $licenseUsage = foreach ($sku in $subscribedSkus) {
            @{
                SkuId           = $sku.SkuId
                SkuPartNumber   = $sku.SkuPartNumber
                AppliesTo       = $sku.AppliesTo
                CapabilityStatus = $sku.CapabilityStatus
                TotalLicenses   = $sku.PrepaidUnits.Enabled
                ConsumedLicenses = $sku.ConsumedUnits
                AvailableLicenses = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
                SuspendedLicenses = $sku.PrepaidUnits.Suspended
                WarningLicenses = $sku.PrepaidUnits.Warning
                ServicePlans    = $sku.ServicePlans | ForEach-Object {
                    @{
                        ServicePlanId   = $_.ServicePlanId
                        ServicePlanName = $_.ServicePlanName
                        AppliesTo       = $_.AppliesTo
                        ProvisioningStatus = $_.ProvisioningStatus
                    }
                }
            }
        }

        $analysis = @{
            TotalSkus                = $subscribedSkus.Count
            TotalLicensesPurchased   = ($licenseUsage | Measure-Object -Property TotalLicenses -Sum).Sum
            TotalLicensesConsumed    = ($licenseUsage | Measure-Object -Property ConsumedLicenses -Sum).Sum
            TotalLicensesAvailable   = ($licenseUsage | Measure-Object -Property AvailableLicenses -Sum).Sum
            OverallocatedSkus        = ($licenseUsage | Where-Object { $_.AvailableLicenses -lt 0 }).Count
        }

        # Detect gotchas

        # Over-allocated licenses
        $overAllocated = $licenseUsage | Where-Object { $_.AvailableLicenses -lt 0 }
        if ($overAllocated.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Over-Allocated Licenses" `
                -Description "Found $($overAllocated.Count) SKUs with more assigned licenses than purchased. This needs resolution." `
                -Severity "High" `
                -Recommendation "Resolve license over-allocation before migration. Review and remove unnecessary assignments." `
                -AffectedObjects @($overAllocated.SkuPartNumber) `
                -MigrationPhase "Pre-Migration"
        }

        # High license utilization
        $highUtilization = $licenseUsage | Where-Object {
            $_.TotalLicenses -gt 0 -and (($_.ConsumedLicenses / $_.TotalLicenses) -gt 0.95)
        }
        if ($highUtilization.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "High License Utilization" `
                -Description "Found $($highUtilization.Count) SKUs with >95% license utilization. May need additional licenses for target tenant." `
                -Severity "Medium" `
                -Recommendation "Plan license procurement for target tenant. Consider temporary license needs during coexistence." `
                -AffectedObjects @($highUtilization.SkuPartNumber) `
                -MigrationPhase "Pre-Migration"
        }

        # Disabled service plans
        $disabledPlans = foreach ($sku in $subscribedSkus) {
            $disabled = $sku.ServicePlans | Where-Object { $_.ProvisioningStatus -eq "Disabled" }
            if ($disabled.Count -gt 0) {
                @{
                    SkuPartNumber    = $sku.SkuPartNumber
                    DisabledPlans    = $disabled.ServicePlanName
                }
            }
        }

        if ($disabledPlans.Count -gt 0) {
            Add-MigrationGotcha -Category "EntraID" `
                -Title "Disabled Service Plans" `
                -Description "Found SKUs with disabled service plans. This configuration needs recreation in target tenant." `
                -Severity "Low" `
                -Recommendation "Document disabled service plans per SKU. Plan to replicate configuration in target." `
                -AffectedCount $disabledPlans.Count `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            SubscribedSkus = $subscribedSkus
            LicenseUsage   = $licenseUsage
            DisabledPlans  = $disabledPlans
            Analysis       = $analysis
        }

        Add-CollectedData -Category "Licensing" -SubCategory "Licenses" -Data $result
        Write-Log -Message "Collected $($subscribedSkus.Count) subscribed SKUs" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect licensing: $_" -Level Error
        throw
    }
}
#endregion

#region Main Collection Function
function Invoke-EntraIDCollection {
    <#
    .SYNOPSIS
        Runs all Entra ID data collection functions
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [hashtable]$Config
    )

    Write-Log -Message "Starting Entra ID data collection..." -Level Info

    $results = @{
        StartTime = Get-Date
        Collections = @{}
        Errors = @()
    }

    $collections = @(
        @{ Name = "TenantInfo"; Function = { Get-EntraTenantInfo } }
        @{ Name = "Users"; Function = { Get-EntraUsers } }
        @{ Name = "Groups"; Function = { Get-EntraGroups } }
        @{ Name = "Devices"; Function = { Get-EntraDevices } }
        @{ Name = "Applications"; Function = { Get-EntraApplications } }
        @{ Name = "ConditionalAccess"; Function = { Get-EntraConditionalAccess } }
        @{ Name = "Roles"; Function = { Get-EntraRoles } }
        @{ Name = "AuthenticationMethods"; Function = { Get-EntraAuthenticationMethods } }
        @{ Name = "Licensing"; Function = { Get-EntraLicensing } }
    )

    foreach ($collection in $collections) {
        try {
            Write-Progress -Activity "Entra ID Collection" -Status "Collecting $($collection.Name)..."
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

    Write-Log -Message "Entra ID collection completed in $($results.Duration.TotalMinutes.ToString('F2')) minutes" -Level Success

    return $results
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Get-EntraTenantInfo',
    'Get-EntraUsers',
    'Get-EntraGroups',
    'Get-EntraDevices',
    'Get-EntraApplications',
    'Get-EntraConditionalAccess',
    'Get-EntraRoles',
    'Get-EntraAuthenticationMethods',
    'Get-EntraLicensing',
    'Invoke-EntraIDCollection'
)
