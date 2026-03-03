#Requires -Version 7.0
<#
.SYNOPSIS
    Security & Compliance Data Collection Module
.DESCRIPTION
    Collects comprehensive security and compliance data including DLP policies,
    retention policies, sensitivity labels, eDiscovery cases, and audit configuration.
    Identifies migration gotchas related to security and compliance.
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

#region Sensitivity Labels
function Get-SecuritySensitivityLabels {
    <#
    .SYNOPSIS
        Collects sensitivity label configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting sensitivity labels..." -Level Info

    try {
        # Get sensitivity labels using Graph API
        $uri = "https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels"
        $labels = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue

        if (-not $labels) {
            Write-Log -Message "Sensitivity labels not accessible or not configured" -Level Warning
            return @{ Configured = $false }
        }

        $labelDetails = foreach ($label in $labels.value) {
            @{
                Id              = $label.id
                Name            = $label.name
                DisplayName     = $label.displayName
                Description     = $label.description
                IsActive        = $label.isActive
                Tooltip         = $label.tooltip
                Color           = $label.color
                Priority        = $label.priority
                ContentFormats  = $label.contentFormats
                HasProtection   = $label.hasProtection
                IsDefault       = $label.isDefault
                Parent          = $label.parent
            }
        }

        $analysis = @{
            TotalLabels      = $labels.value.Count
            ActiveLabels     = ($labelDetails | Where-Object { $_.IsActive }).Count
            ProtectedLabels  = ($labelDetails | Where-Object { $_.HasProtection }).Count
            DefaultLabels    = ($labelDetails | Where-Object { $_.IsDefault }).Count
        }

        # Detect gotchas
        if ($labels.value.Count -gt 0) {
            Add-MigrationGotcha -Category "Security" `
                -Title "Sensitivity Labels Configured" `
                -Description "Found $($labels.value.Count) sensitivity labels. Labels must be recreated in target tenant with matching GUIDs or remapped." `
                -Severity "Critical" `
                -Recommendation "Export label configurations. Consider label GUID preservation. Plan for label policy recreation and publishing." `
                -AffectedCount $labels.value.Count `
                -MigrationPhase "Pre-Migration"
        }

        $protectedLabels = $labelDetails | Where-Object { $_.HasProtection }
        if ($protectedLabels.Count -gt 0) {
            Add-MigrationGotcha -Category "Security" `
                -Title "Labels with Protection Settings" `
                -Description "Found $($protectedLabels.Count) labels with encryption/protection. Protected content needs special handling during migration." `
                -Severity "Critical" `
                -Recommendation "Document protection settings. Plan for Azure RMS migration. Consider super user access for protected content." `
                -AffectedObjects @($protectedLabels.DisplayName) `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Configured = $true
            Labels     = $labelDetails
            Analysis   = $analysis
        }

        Add-CollectedData -Category "Security" -SubCategory "SensitivityLabels" -Data $result
        Write-Log -Message "Collected $($labels.value.Count) sensitivity labels" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect sensitivity labels: $_" -Level Error
        return @{ Configured = $false; Error = $_.Exception.Message }
    }
}
#endregion

#region DLP Policies
function Get-SecurityDLPPolicies {
    <#
    .SYNOPSIS
        Collects Data Loss Prevention policy information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting DLP policies..." -Level Info

    try {
        # Get DLP policies via Security & Compliance PowerShell
        $dlpPolicies = Get-DlpCompliancePolicy -ErrorAction SilentlyContinue

        if (-not $dlpPolicies) {
            Write-Log -Message "DLP policies not accessible or not configured" -Level Warning
            return @{ Configured = $false }
        }

        $policyDetails = foreach ($policy in $dlpPolicies) {
            # Get rules for each policy
            $rules = Get-DlpComplianceRule -Policy $policy.Name -ErrorAction SilentlyContinue

            @{
                Name            = $policy.Name
                Mode            = $policy.Mode
                Enabled         = $policy.Enabled
                Type            = $policy.Type
                Workload        = $policy.Workload
                Priority        = $policy.Priority
                CreatedBy       = $policy.CreatedBy
                WhenCreated     = $policy.WhenCreated
                WhenChanged     = $policy.WhenChanged
                Comment         = $policy.Comment
                ExchangeLocation = $policy.ExchangeLocation
                SharePointLocation = $policy.SharePointLocation
                OneDriveLocation = $policy.OneDriveLocation
                TeamsLocation   = $policy.TeamsLocation
                RuleCount       = ($rules | Measure-Object).Count
                Rules           = @($rules | ForEach-Object {
                    @{
                        Name            = $_.Name
                        Disabled        = $_.Disabled
                        Priority        = $_.Priority
                        ContentContainsSensitiveInformation = $_.ContentContainsSensitiveInformation
                        BlockAccess     = $_.BlockAccess
                        NotifyUser      = $_.NotifyUser
                    }
                })
            }
        }

        $analysis = @{
            TotalPolicies     = $dlpPolicies.Count
            EnabledPolicies   = ($policyDetails | Where-Object { $_.Enabled }).Count
            TestModePolicies  = ($policyDetails | Where-Object { $_.Mode -eq "TestWithNotifications" -or $_.Mode -eq "TestWithoutNotifications" }).Count
            EnforcedPolicies  = ($policyDetails | Where-Object { $_.Mode -eq "Enable" }).Count
            TotalRules        = ($policyDetails | ForEach-Object { $_.RuleCount } | Measure-Object -Sum).Sum
        }

        # Detect gotchas
        if ($dlpPolicies.Count -gt 0) {
            Add-MigrationGotcha -Category "Security" `
                -Title "DLP Policies Configured" `
                -Description "Found $($dlpPolicies.Count) DLP policies with $($analysis.TotalRules) rules. DLP policies must be recreated in target tenant." `
                -Severity "High" `
                -Recommendation "Export DLP policy configurations. Document sensitive information types used. Plan for policy recreation and testing." `
                -AffectedCount $dlpPolicies.Count `
                -MigrationPhase "Pre-Migration"
        }

        $enforcedPolicies = $policyDetails | Where-Object { $_.Mode -eq "Enable" }
        if ($enforcedPolicies.Count -gt 0) {
            Add-MigrationGotcha -Category "Security" `
                -Title "Enforced DLP Policies" `
                -Description "Found $($enforcedPolicies.Count) DLP policies in enforcement mode. Consider implications during migration." `
                -Severity "Medium" `
                -Recommendation "Plan DLP policy deployment timeline in target. May need to run in test mode initially." `
                -AffectedObjects @($enforcedPolicies.Name) `
                -MigrationPhase "Post-Migration"
        }

        $result = @{
            Configured = $true
            Policies   = $policyDetails
            Analysis   = $analysis
        }

        Add-CollectedData -Category "Security" -SubCategory "DLPPolicies" -Data $result
        Write-Log -Message "Collected $($dlpPolicies.Count) DLP policies" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect DLP policies: $_" -Level Error
        return @{ Configured = $false; Error = $_.Exception.Message }
    }
}
#endregion

#region Retention Policies
function Get-SecurityRetentionPolicies {
    <#
    .SYNOPSIS
        Collects retention policy information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting retention policies..." -Level Info

    try {
        $retentionPolicies = Get-RetentionCompliancePolicy -ErrorAction SilentlyContinue

        if (-not $retentionPolicies) {
            Write-Log -Message "Retention policies not accessible or not configured" -Level Warning
            return @{ Configured = $false }
        }

        $policyDetails = foreach ($policy in $retentionPolicies) {
            # Get rules for each policy
            $rules = Get-RetentionComplianceRule -Policy $policy.Name -ErrorAction SilentlyContinue

            @{
                Name            = $policy.Name
                Enabled         = $policy.Enabled
                Mode            = $policy.Mode
                Type            = $policy.Type
                Workload        = $policy.Workload
                WhenCreated     = $policy.WhenCreated
                WhenChanged     = $policy.WhenChanged
                Comment         = $policy.Comment
                ExchangeLocation = $policy.ExchangeLocation
                SharePointLocation = $policy.SharePointLocation
                OneDriveLocation = $policy.OneDriveLocation
                ModernGroupLocation = $policy.ModernGroupLocation
                TeamsChannelLocation = $policy.TeamsChannelLocation
                TeamsChatLocation = $policy.TeamsChatLocation
                Rules           = @($rules | ForEach-Object {
                    @{
                        Name                = $_.Name
                        RetentionDuration   = $_.RetentionDuration
                        RetentionDurationDisplayHint = $_.RetentionDurationDisplayHint
                        RetentionComplianceAction = $_.RetentionComplianceAction
                        ExpirationDateOption = $_.ExpirationDateOption
                    }
                })
            }
        }

        $analysis = @{
            TotalPolicies    = $retentionPolicies.Count
            EnabledPolicies  = ($policyDetails | Where-Object { $_.Enabled }).Count
            ExchangePolicies = ($policyDetails | Where-Object { $_.ExchangeLocation }).Count
            SharePointPolicies = ($policyDetails | Where-Object { $_.SharePointLocation }).Count
            TeamsPolicies    = ($policyDetails | Where-Object { $_.TeamsChannelLocation -or $_.TeamsChatLocation }).Count
        }

        # Detect gotchas
        if ($retentionPolicies.Count -gt 0) {
            Add-MigrationGotcha -Category "Security" `
                -Title "Retention Policies Configured" `
                -Description "Found $($retentionPolicies.Count) retention policies. These define legal and compliance data retention requirements." `
                -Severity "Critical" `
                -Recommendation "Document all retention policies and their scope. Coordinate with legal/compliance. Recreate policies in target tenant." `
                -AffectedCount $retentionPolicies.Count `
                -MigrationPhase "Pre-Migration"
        }

        # Check for Teams retention
        $teamsPolicies = $policyDetails | Where-Object { $_.TeamsChannelLocation -or $_.TeamsChatLocation }
        if ($teamsPolicies.Count -gt 0) {
            Add-MigrationGotcha -Category "Security" `
                -Title "Teams Retention Policies" `
                -Description "Found $($teamsPolicies.Count) retention policies for Teams. Teams chat history is subject to retention." `
                -Severity "High" `
                -Recommendation "Plan Teams data retention continuity. Ensure policies are created before Teams migration." `
                -AffectedObjects @($teamsPolicies.Name) `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Configured = $true
            Policies   = $policyDetails
            Analysis   = $analysis
        }

        Add-CollectedData -Category "Security" -SubCategory "RetentionPolicies" -Data $result
        Write-Log -Message "Collected $($retentionPolicies.Count) retention policies" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect retention policies: $_" -Level Error
        return @{ Configured = $false; Error = $_.Exception.Message }
    }
}
#endregion

#region eDiscovery Cases
function Get-SecurityEDiscoveryCases {
    <#
    .SYNOPSIS
        Collects eDiscovery case information
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting eDiscovery cases..." -Level Info

    try {
        # Get Core eDiscovery cases
        $coreCases = Get-ComplianceCase -ErrorAction SilentlyContinue

        # Get Advanced eDiscovery cases if available
        $advancedCases = Get-ComplianceCase -CaseType AdvancedEdiscovery -ErrorAction SilentlyContinue

        $allCases = @()
        if ($coreCases) { $allCases += $coreCases }
        if ($advancedCases) { $allCases += $advancedCases }

        if ($allCases.Count -eq 0) {
            Write-Log -Message "No eDiscovery cases found" -Level Info
            return @{ Configured = $false }
        }

        $caseDetails = foreach ($case in $allCases) {
            # Get holds for each case
            $holds = Get-CaseHoldPolicy -Case $case.Identity -ErrorAction SilentlyContinue

            @{
                Name            = $case.Name
                Identity        = $case.Identity
                Status          = $case.Status
                CaseType        = $case.CaseType
                CreatedDateTime = $case.CreatedDateTime
                LastModifiedDateTime = $case.LastModifiedDateTime
                ClosedDateTime  = $case.ClosedDateTime
                ClosedBy        = $case.ClosedBy
                HoldCount       = ($holds | Measure-Object).Count
                Holds           = @($holds | ForEach-Object {
                    @{
                        Name    = $_.Name
                        Enabled = $_.Enabled
                    }
                })
            }
        }

        $analysis = @{
            TotalCases       = $allCases.Count
            ActiveCases      = ($caseDetails | Where-Object { $_.Status -eq "Active" }).Count
            ClosedCases      = ($caseDetails | Where-Object { $_.Status -eq "Closed" }).Count
            CoreCases        = ($coreCases | Measure-Object).Count
            AdvancedCases    = ($advancedCases | Measure-Object).Count
            TotalHolds       = ($caseDetails | ForEach-Object { $_.HoldCount } | Measure-Object -Sum).Sum
        }

        # Detect gotchas
        $activeCases = $caseDetails | Where-Object { $_.Status -eq "Active" }
        if ($activeCases.Count -gt 0) {
            Add-MigrationGotcha -Category "Security" `
                -Title "Active eDiscovery Cases" `
                -Description "Found $($activeCases.Count) active eDiscovery cases with $($analysis.TotalHolds) holds. Active cases require legal coordination." `
                -Severity "Critical" `
                -Recommendation "Coordinate with legal before migration. Active holds must be maintained. Plan for case data export and recreation." `
                -AffectedCount $activeCases.Count `
                -MigrationPhase "Pre-Migration"
        }

        if ($analysis.TotalHolds -gt 0) {
            Add-MigrationGotcha -Category "Security" `
                -Title "eDiscovery Holds Active" `
                -Description "Found $($analysis.TotalHolds) eDiscovery hold(s). Content under hold must be preserved during migration." `
                -Severity "Critical" `
                -Recommendation "Maintain chain of custody. Ensure holds are replicated in target before releasing source holds." `
                -AffectedCount $analysis.TotalHolds `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Configured = $true
            Cases      = $caseDetails
            Analysis   = $analysis
        }

        Add-CollectedData -Category "Security" -SubCategory "eDiscovery" -Data $result
        Write-Log -Message "Collected $($allCases.Count) eDiscovery cases" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect eDiscovery cases: $_" -Level Error
        return @{ Configured = $false; Error = $_.Exception.Message }
    }
}
#endregion

#region Audit Configuration
function Get-SecurityAuditConfig {
    <#
    .SYNOPSIS
        Collects audit logging configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting audit configuration..." -Level Info

    try {
        # Get audit configuration
        $auditConfig = Get-AdminAuditLogConfig -ErrorAction SilentlyContinue

        # Get unified audit log configuration
        $unifiedAuditEnabled = $true  # Default in M365

        $config = @{
            AdminAuditLogEnabled     = $auditConfig.AdminAuditLogEnabled
            UnifiedAuditLogEnabled   = $unifiedAuditEnabled
            AdminAuditLogAgeLimit    = $auditConfig.AdminAuditLogAgeLimit
            AdminAuditLogCmdlets     = $auditConfig.AdminAuditLogCmdlets
            AdminAuditLogParameters  = $auditConfig.AdminAuditLogParameters
        }

        $analysis = @{
            AuditEnabled = $auditConfig.AdminAuditLogEnabled
        }

        # Detect gotchas
        if ($config.AdminAuditLogEnabled) {
            Add-MigrationGotcha -Category "Security" `
                -Title "Audit Logging Enabled" `
                -Description "Audit logging is enabled. Historical audit logs cannot be migrated to target tenant." `
                -Severity "High" `
                -Recommendation "Export historical audit logs before migration if needed for compliance. Configure audit logging in target tenant." `
                -MigrationPhase "Pre-Migration"
        }

        $result = @{
            Configuration = $config
            Analysis      = $analysis
        }

        Add-CollectedData -Category "Security" -SubCategory "AuditConfig" -Data $result
        Write-Log -Message "Audit configuration collected" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect audit config: $_" -Level Error
        throw
    }
}
#endregion

#region Alert Policies
function Get-SecurityAlertPolicies {
    <#
    .SYNOPSIS
        Collects security alert policies
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting alert policies..." -Level Info

    try {
        $alertPolicies = Get-ProtectionAlert -ErrorAction SilentlyContinue

        if (-not $alertPolicies) {
            Write-Log -Message "Alert policies not accessible" -Level Warning
            return @{ Configured = $false }
        }

        $policyDetails = foreach ($policy in $alertPolicies) {
            @{
                Name            = $policy.Name
                Comment         = $policy.Comment
                Severity        = $policy.Severity
                Category        = $policy.Category
                NotifyUser      = $policy.NotifyUser
                NotifyUserOnFilterMatch = $policy.NotifyUserOnFilterMatch
                Disabled        = $policy.Disabled
                IsSystemRule    = $policy.IsSystemRule
                WhenCreated     = $policy.WhenCreated
            }
        }

        $customPolicies = $policyDetails | Where-Object { -not $_.IsSystemRule }

        $analysis = @{
            TotalPolicies   = $alertPolicies.Count
            CustomPolicies  = $customPolicies.Count
            SystemPolicies  = ($policyDetails | Where-Object { $_.IsSystemRule }).Count
            EnabledPolicies = ($policyDetails | Where-Object { -not $_.Disabled }).Count
        }

        # Detect gotchas
        if ($customPolicies.Count -gt 0) {
            Add-MigrationGotcha -Category "Security" `
                -Title "Custom Alert Policies" `
                -Description "Found $($customPolicies.Count) custom alert policies. These need recreation in target tenant." `
                -Severity "Medium" `
                -Recommendation "Export custom alert policy configurations. Recreate in target tenant." `
                -AffectedCount $customPolicies.Count `
                -MigrationPhase "Post-Migration"
        }

        $result = @{
            Configured = $true
            Policies   = $policyDetails
            Analysis   = $analysis
        }

        Add-CollectedData -Category "Security" -SubCategory "AlertPolicies" -Data $result
        Write-Log -Message "Collected $($alertPolicies.Count) alert policies" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect alert policies: $_" -Level Error
        return @{ Configured = $false; Error = $_.Exception.Message }
    }
}
#endregion

#region Insider Risk Management
function Get-SecurityInsiderRiskConfig {
    <#
    .SYNOPSIS
        Collects Insider Risk Management configuration
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message "Collecting Insider Risk Management configuration..." -Level Info

    try {
        # Insider Risk policies via Graph API
        $uri = "https://graph.microsoft.com/beta/security/triggerTypes/retentionEventTypes"

        try {
            $config = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            $insiderRiskConfigured = $true
        }
        catch {
            $insiderRiskConfigured = $false
        }

        if ($insiderRiskConfigured) {
            Add-MigrationGotcha -Category "Security" `
                -Title "Insider Risk Management Configured" `
                -Description "Insider Risk Management appears to be configured. IRM policies and data do not migrate automatically." `
                -Severity "High" `
                -Recommendation "Document IRM policies. Recreate policies in target tenant. Historical alerts will not migrate." `
                -MigrationPhase "Post-Migration"
        }

        $result = @{
            Configured = $insiderRiskConfigured
        }

        Add-CollectedData -Category "Security" -SubCategory "InsiderRisk" -Data $result
        Write-Log -Message "Insider Risk Management configuration collected" -Level Success

        return $result
    }
    catch {
        Write-Log -Message "Failed to collect Insider Risk config: $_" -Level Error
        return @{ Configured = $false }
    }
}
#endregion

#region Main Collection Function
function Invoke-SecurityCollection {
    <#
    .SYNOPSIS
        Runs all Security & Compliance data collection functions
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [hashtable]$Config
    )

    Write-Log -Message "Starting Security & Compliance data collection..." -Level Info

    $results = @{
        StartTime = Get-Date
        Collections = @{}
        Errors = @()
    }

    $collections = @(
        @{ Name = "SensitivityLabels"; Function = { Get-SecuritySensitivityLabels } }
        @{ Name = "DLPPolicies"; Function = { Get-SecurityDLPPolicies } }
        @{ Name = "RetentionPolicies"; Function = { Get-SecurityRetentionPolicies } }
        @{ Name = "eDiscovery"; Function = { Get-SecurityEDiscoveryCases } }
        @{ Name = "AuditConfig"; Function = { Get-SecurityAuditConfig } }
        @{ Name = "AlertPolicies"; Function = { Get-SecurityAlertPolicies } }
        @{ Name = "InsiderRisk"; Function = { Get-SecurityInsiderRiskConfig } }
    )

    foreach ($collection in $collections) {
        try {
            Write-Progress -Activity "Security Collection" -Status "Collecting $($collection.Name)..."
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

    Write-Log -Message "Security collection completed in $($results.Duration.TotalMinutes.ToString('F2')) minutes" -Level Success

    return $results
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Get-SecuritySensitivityLabels',
    'Get-SecurityDLPPolicies',
    'Get-SecurityRetentionPolicies',
    'Get-SecurityEDiscoveryCases',
    'Get-SecurityAuditConfig',
    'Get-SecurityAlertPolicies',
    'Get-SecurityInsiderRiskConfig',
    'Invoke-SecurityCollection'
)
