#Requires -Version 7.0
<#
.SYNOPSIS
    Gotcha Analysis Engine for M365 Migration
.DESCRIPTION
    Comprehensive analysis engine that processes collected tenant data
    and identifies migration risks, gotchas, and recommendations.
    Includes severity scoring, categorization, and prioritization.
.NOTES
    Author: AI Migration Expert
    Version: 1.0.0
    Target: PowerShell 7.x
#>

# Import core module only if not already loaded
if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    $corePath = Join-Path $PSScriptRoot ".." "Modules" "Core" "TenantDiscovery.Core.psm1"
    if (Test-Path $corePath) {
        Import-Module $corePath -Force -Global
    }
}

#region Risk Categories
$script:RiskCategories = @{
    Identity = @{
        Name        = "Identity & Access Management"
        Description = "Risks related to users, groups, authentication, and access control"
        Weight      = 1.5
    }
    Data = @{
        Name        = "Data & Content"
        Description = "Risks related to mailboxes, files, SharePoint, and data migration"
        Weight      = 1.3
    }
    Compliance = @{
        Name        = "Compliance & Security"
        Description = "Risks related to security policies, compliance requirements, and legal holds"
        Weight      = 1.8
    }
    Integration = @{
        Name        = "Applications & Integrations"
        Description = "Risks related to apps, connectors, and third-party integrations"
        Weight      = 1.2
    }
    Infrastructure = @{
        Name        = "Hybrid Infrastructure"
        Description = "Risks related to on-premises connectivity and hybrid configurations"
        Weight      = 1.4
    }
    Operations = @{
        Name        = "Operational Readiness"
        Description = "Risks related to timing, dependencies, and operational impact"
        Weight      = 1.0
    }
}

$script:SeverityWeights = @{
    Critical      = 100
    High          = 75
    Medium        = 50
    Low           = 25
    Informational = 10
}
#endregion

#region Analysis Rules
function Get-AnalysisRules {
    <#
    .SYNOPSIS
        Returns the complete set of analysis rules
    #>
    return @(
        # Identity Rules
        @{
            Id          = "ID-001"
            Category    = "Identity"
            Name        = "Synced Users Assessment"
            Condition   = { param($data) $data.EntraID.Users.Analysis.SyncedUsers -gt 0 }
            Severity    = "High"
            Description = "On-premises synced users require special handling for ImmutableID preservation"
            Recommendation = "Document sync configuration and plan ImmutableID strategy"
            RemediationSteps = @(
                "1. EXPORT SOURCE DATA: Run 'Get-MgUser -All | Select DisplayName,UserPrincipalName,OnPremisesImmutableId,OnPremisesSyncEnabled | Export-Csv SyncedUsers.csv'"
                "2. DOCUMENT AAD CONNECT CONFIG: On AAD Connect server, export config via 'Get-ADSyncServerConfiguration -Path C:\AADConnectExport'"
                "3. CHOOSE STRATEGY: Option A (Hard Match) - Preserve ImmutableID by pre-creating users in target with same ImmutableID. Option B (Soft Match) - Let AAD Connect match by UPN/ProxyAddresses"
                "4. FOR HARD MATCH: Create users in target tenant with: New-MgUser -OnPremisesImmutableId '<base64-guid>' before connecting AAD Connect"
                "5. CONFIGURE TARGET AAD CONNECT: Install new AAD Connect pointing to target tenant. Use same OU filtering and sync rules"
                "6. VALIDATE: Run 'Get-MgUser -Filter `"onPremisesSyncEnabled eq true`"' in target to verify sync"
                "7. CUTOVER: Disable sync in source (Set-MsolDirSyncEnabled -EnableDirSync `$false), wait 72hrs, enable in target"
            )
            Tools = @("Azure AD Connect", "Microsoft Graph PowerShell", "ADSyncTools")
            EstimatedEffort = "2-5 days depending on user count"
            Prerequisites = @("Target tenant AAD Connect server prepared", "Service account created in target")
        }
        @{
            Id          = "ID-002"
            Category    = "Identity"
            Name        = "Guest User Presence"
            Condition   = { param($data) $data.EntraID.Users.Analysis.GuestUsers -gt 100 }
            Severity    = "Medium"
            Description = "Large number of guest users will require reinvitation in target tenant"
            Recommendation = "Automate guest reinvitation process using PowerShell scripts"
            RemediationSteps = @(
                "1. EXPORT GUEST LIST: Get-MgUser -Filter `"userType eq 'Guest'`" | Select DisplayName,Mail,UserPrincipalName | Export-Csv Guests.csv"
                "2. DOCUMENT PERMISSIONS: For each guest, export group memberships and app assignments using Get-MgUserMemberOf"
                "3. PREPARE INVITATION SCRIPT: Create bulk invitation script using New-MgInvitation cmdlet"
                "4. SAMPLE INVITATION CODE: `$guests | ForEach-Object { New-MgInvitation -InvitedUserEmailAddress `$_.Mail -InviteRedirectUrl 'https://myapps.microsoft.com' -SendInvitationMessage:`$true }"
                "5. RESTORE PERMISSIONS: After guests accept, re-add to groups using Add-MgGroupMember"
                "6. NOTIFY GUESTS: Send communication explaining they'll receive new invitation and need to accept"
                "7. VALIDATE: Compare guest count and permissions between source and target tenants"
            )
            Tools = @("Microsoft Graph PowerShell", "Excel for tracking")
            EstimatedEffort = "1-2 days + guest acceptance time"
            Prerequisites = @("Guest email addresses verified", "Target groups created")
        }
        @{
            Id          = "ID-003"
            Category    = "Identity"
            Name        = "Hybrid Devices"
            Condition   = { param($data) $data.EntraID.Devices.Analysis.HybridJoined -gt 0 }
            Severity    = "Critical"
            Description = "Hybrid Azure AD joined devices must be unjoined and rejoined to target"
            Recommendation = "Plan phased device migration with user communication strategy"
            RemediationSteps = @(
                "1. INVENTORY DEVICES: Get-MgDevice -Filter `"trustType eq 'ServerAd'`" | Export-Csv HybridDevices.csv"
                "2. CREATE SCP FOR TARGET: In AD, create new Service Connection Point for target tenant using: Initialize-ADSyncDomainJoinedComputerSync -AdConnectorAccount <account> -AzureADCredentials <creds>"
                "3. UNJOIN FROM SOURCE: On each device run 'dsregcmd /leave' (requires local admin)"
                "4. AUTOMATED UNJOIN SCRIPT: Deploy via Intune/SCCM: `$dsregcmd = dsregcmd /status; if (`$dsregcmd -match 'AzureAdJoined : YES') { dsregcmd /leave }"
                "5. TRIGGER REJOIN: After SCP updated, devices auto-rejoin on next login. Force with: dsregcmd /join"
                "6. VERIFY JOIN: Run 'dsregcmd /status' - confirm AzureADJoined=YES and correct TenantId"
                "7. VALIDATE IN PORTAL: Check Entra ID > Devices for device registration in target tenant"
                "8. UPDATE CONDITIONAL ACCESS: Ensure CA policies in target include device compliance requirements"
            )
            Tools = @("dsregcmd.exe", "Group Policy", "Intune/SCCM for deployment")
            EstimatedEffort = "3-7 days for phased rollout"
            Prerequisites = @("Target tenant SCP configured", "User communication sent", "CA policies prepared")
        }
        @{
            Id          = "ID-004"
            Category    = "Identity"
            Name        = "Conditional Access Complexity"
            Condition   = { param($data) $data.EntraID.ConditionalAccess.Analysis.TotalPolicies -gt 20 }
            Severity    = "High"
            Description = "Complex Conditional Access policy set requires careful recreation"
            Recommendation = "Export and document all CA policies. Plan phased deployment in target."
            RemediationSteps = @(
                "1. EXPORT ALL POLICIES: Get-MgIdentityConditionalAccessPolicy -All | ConvertTo-Json -Depth 10 | Out-File CAPolicies.json"
                "2. EXPORT NAMED LOCATIONS: Get-MgIdentityConditionalAccessNamedLocation | ConvertTo-Json -Depth 10 | Out-File NamedLocations.json"
                "3. DOCUMENT DEPENDENCIES: List all groups, apps, and users referenced in policies. These must exist in target first"
                "4. CREATE NAMED LOCATIONS FIRST: In target, recreate named locations using New-MgIdentityConditionalAccessNamedLocation"
                "5. CREATE POLICIES IN REPORT-ONLY: Import policies to target in Report-Only mode: `$policy.State = 'enabledForReportingButNotEnforced'"
                "6. MAP OBJECT IDS: Update all GUIDs (groups, apps, users) to target tenant equivalents before import"
                "7. TEST WITH PILOT GROUP: Enable policies for pilot users first, monitor Sign-in logs for unexpected blocks"
                "8. ENABLE INCREMENTALLY: Move policies to 'enabled' state in batches, monitoring for 24-48hrs between batches"
                "9. VALIDATE: Compare policy count and settings between tenants using Get-MgIdentityConditionalAccessPolicy"
            )
            Tools = @("Microsoft Graph PowerShell", "Conditional Access Documentation Workbook", "Azure AD Sign-in Logs")
            EstimatedEffort = "3-5 days for policy recreation and testing"
            Prerequisites = @("All referenced groups/apps created in target", "Named locations created", "Test users identified")
        }
        @{
            Id          = "ID-005"
            Category    = "Identity"
            Name        = "Custom Roles"
            Condition   = { param($data) $data.EntraID.Roles.Analysis.CustomRoleDefinitions -gt 0 }
            Severity    = "Medium"
            Description = "Custom role definitions must be recreated in target tenant"
            Recommendation = "Export custom role definitions and permissions. Test in target before migration."
            RemediationSteps = @(
                "1. EXPORT CUSTOM ROLES: Get-MgRoleManagementDirectoryRoleDefinition -Filter `"isBuiltIn eq false`" | ConvertTo-Json -Depth 10 | Out-File CustomRoles.json"
                "2. DOCUMENT ROLE ASSIGNMENTS: Get-MgRoleManagementDirectoryRoleAssignment | Where {`$_.RoleDefinitionId -in `$customRoleIds} | Export-Csv RoleAssignments.csv"
                "3. CREATE ROLES IN TARGET: Use New-MgRoleManagementDirectoryRoleDefinition with exported permissions"
                "4. SAMPLE CREATION: `$roleParams = @{DisplayName='Custom Role';RolePermissions=@(@{AllowedResourceActions=@('microsoft.directory/users/read')})}; New-MgRoleManagementDirectoryRoleDefinition -BodyParameter `$roleParams"
                "5. VERIFY PERMISSIONS: Compare rolePermissions.allowedResourceActions between source and target"
                "6. REASSIGN ROLES: After user migration, recreate assignments using New-MgRoleManagementDirectoryRoleAssignment"
                "7. TEST ACCESS: Have role holders verify they can perform expected actions in target tenant"
            )
            Tools = @("Microsoft Graph PowerShell", "Entra ID Portal")
            EstimatedEffort = "1-2 days"
            Prerequisites = @("P1/P2 license in target for custom roles", "Role holders identified")
        }

        # Data Rules
        @{
            Id          = "DT-001"
            Category    = "Data"
            Name        = "Large Mailbox Detection"
            Condition   = { param($data)
                $data.Exchange.Mailboxes.Mailboxes | Where-Object {
                    $_.Statistics.TotalItemSize -gt "50 GB"
                } | Measure-Object | Select-Object -ExpandProperty Count
            }
            Severity    = "Medium"
            Description = "Large mailboxes require extended migration windows"
            Recommendation = "Plan incremental mailbox migration. Consider archive strategy."
            RemediationSteps = @(
                "1. IDENTIFY LARGE MAILBOXES: Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics | Where {[int64](`$_.TotalItemSize.Value.ToString().Split('(')[1].Split(' ')[0].Replace(',','')) -gt 53687091200} | Select DisplayName,TotalItemSize"
                "2. CALCULATE MIGRATION TIME: Estimate ~1GB/hour for cross-tenant moves. 50GB = ~50 hours per mailbox"
                "3. ENABLE ARCHIVE FIRST: For mailboxes >50GB, enable archive and move older items: Enable-Mailbox -Archive -Identity user@domain.com"
                "4. CREATE RETENTION POLICY: New-RetentionPolicy 'Archive Policy' with tag to move items >1 year to archive"
                "5. USE MIGRATION BATCHES: Create separate batch for large mailboxes: New-MigrationBatch -Name 'LargeMailboxes' -SourceEndpoint `$endpoint -TargetDeliveryDomain target.mail.onmicrosoft.com"
                "6. SCHEDULE OFF-HOURS: Start large mailbox batches Friday evening: Start-MigrationBatch -Identity 'LargeMailboxes'"
                "7. MONITOR PROGRESS: Get-MigrationUserStatistics -Identity user@domain.com | Select StatusDetail,BytesTransferred,PercentComplete"
                "8. PLAN INCREMENTAL SYNC: Large mailboxes sync incrementally - initial sync, then delta syncs until cutover"
            )
            Tools = @("Exchange Online PowerShell", "Migration Batch cmdlets", "BitTitan MigrationWiz (optional)")
            EstimatedEffort = "Extended migration window - 2-4 days per batch of large mailboxes"
            Prerequisites = @("Archive licenses available", "Migration endpoint configured", "Off-hours window identified")
        }
        @{
            Id          = "DT-002"
            Category    = "Data"
            Name        = "Public Folders Present"
            Condition   = { param($data) $data.Exchange.PublicFolders.Enabled -eq $true }
            Severity    = "High"
            Description = "Public folders require specialized migration approach"
            Recommendation = "Plan batch migration. Consider modernization to M365 Groups."
            RemediationSteps = @(
                "1. INVENTORY PUBLIC FOLDERS: Get-PublicFolder -Recurse | Select Identity,Name,FolderSize,ItemCount | Export-Csv PublicFolders.csv"
                "2. CHECK SIZE LIMITS: Get-PublicFolderStatistics | Measure-Object ItemCount,TotalItemSize -Sum (Max 1M items, 50GB per folder)"
                "3. DECIDE: MIGRATE OR MODERNIZE? Consider converting mail-enabled PFs to M365 Groups/Shared Mailboxes"
                "4. FOR MIGRATION - CREATE MAPPING: New-PublicFolderMigrationRequest requires CSV mapping source to target"
                "5. LOCK SOURCE FOLDERS: Set-OrganizationConfig -PublicFoldersLockedForMigration `$true"
                "6. CREATE MIGRATION BATCH: New-MigrationBatch -Name PFMigration -SourcePublicFolderDatabase (source) -CSVData (mapping)"
                "7. COMPLETE MIGRATION: Complete-MigrationBatch -Identity PFMigration"
                "8. FOR MODERNIZATION: Use Public Folder to M365 Groups migration tool in EAC, or manually recreate as Groups"
                "9. UPDATE PERMISSIONS: Recreate public folder permissions in target using Add-PublicFolderClientPermission"
            )
            Tools = @("Exchange Online PowerShell", "EAC Migration wizard", "PublicFolderToGroupsMigration script")
            EstimatedEffort = "3-7 days depending on size and complexity"
            Prerequisites = @("Public folder hierarchy documented", "Permissions exported", "Decision on modernization made")
        }
        @{
            Id          = "DT-003"
            Category    = "Data"
            Name        = "Large SharePoint Sites"
            Condition   = { param($data)
                ($data.SharePoint.Sites.Sites | Where-Object { $_.StorageUsageMB -gt 100000 }).Count -gt 0
            }
            Severity    = "High"
            Description = "Large SharePoint sites require incremental migration"
            Recommendation = "Use pre-staging approach. Plan for extended migration windows."
            RemediationSteps = @(
                "1. IDENTIFY LARGE SITES: Get-PnPTenantSite | Where {`$_.StorageUsageCurrent -gt 100000} | Select Url,StorageUsageCurrent,LastContentModifiedDate"
                "2. ANALYZE CONTENT: Run SharePoint Migration Assessment Tool (SMAT) to identify migration blockers"
                "3. PRE-STAGE APPROACH: Start migration weeks before cutover to sync bulk data. Use ShareGate or SPMT"
                "4. SPMT COMMAND: Start-SPMTMigration -MigrationType SharePoint -SourceUri 'https://source.sharepoint.com/sites/large' -TargetUri 'https://target.sharepoint.com/sites/large'"
                "5. ENABLE INCREMENTAL: Configure tool to run incremental syncs daily until cutover"
                "6. HANDLE SPECIAL CONTENT: Check for workflows, custom solutions, InfoPath forms - these need separate handling"
                "7. PERMISSION MAPPING: Create user mapping file matching source to target user UPNs"
                "8. CUTOVER STEPS: Final incremental sync -> Make source read-only -> Verify target -> Update DNS/links"
                "9. VALIDATE: Compare item counts, folder structure, permissions between source and target"
            )
            Tools = @("SharePoint Migration Tool (SPMT)", "ShareGate", "SMAT", "PnP PowerShell")
            EstimatedEffort = "1-2 weeks for pre-staging plus cutover window"
            Prerequisites = @("Migration tool licensed", "Target site created", "User mapping prepared")
        }
        @{
            Id          = "DT-004"
            Category    = "Data"
            Name        = "Archive Mailboxes"
            Condition   = { param($data) $data.Exchange.Mailboxes.Analysis.ArchiveEnabled -gt 0 }
            Severity    = "Medium"
            Description = "Archive mailboxes require separate migration consideration"
            Recommendation = "Plan archive migration. Verify licensing in target tenant."
            RemediationSteps = @(
                "1. INVENTORY ARCHIVES: Get-Mailbox -Archive -ResultSize Unlimited | Get-MailboxStatistics -Archive | Select DisplayName,TotalItemSize,ItemCount"
                "2. VERIFY TARGET LICENSING: Archives require Exchange Online Plan 2 or E3/E5. Confirm licenses available"
                "3. ENABLE ARCHIVES IN TARGET: After mailbox creation in target, enable archive: Enable-Mailbox -Identity user@target.com -Archive"
                "4. MIGRATE PRIMARY FIRST: Always migrate primary mailbox before archive"
                "5. MIGRATE ARCHIVE: Use New-MigrationBatch with -ArchiveOnly switch, or include PrimaryOnly:`$false in batch"
                "6. SAMPLE COMMAND: New-MigrationBatch -Name 'ArchiveMigration' -UserIds user1,user2 -ArchiveOnly"
                "7. AUTO-EXPANDING ARCHIVES: If source uses auto-expanding, ensure target tenant has it enabled: Set-OrganizationConfig -AutoExpandingArchive"
                "8. VERIFY: Get-MailboxStatistics -Identity user@target.com -Archive | Select TotalItemSize,ItemCount"
            )
            Tools = @("Exchange Online PowerShell", "Migration Batch cmdlets")
            EstimatedEffort = "Additional 50% time on top of primary mailbox migration"
            Prerequisites = @("Archive licenses assigned in target", "Primary mailbox migrated first")
        }

        # Compliance Rules
        @{
            Id          = "CP-001"
            Category    = "Compliance"
            Name        = "Litigation Hold Active"
            Condition   = { param($data) $data.Exchange.Mailboxes.Analysis.LitigationHold -gt 0 }
            Severity    = "Critical"
            Description = "Mailboxes under litigation hold must maintain legal compliance"
            Recommendation = "Coordinate with legal. Maintain chain of custody documentation."
            RemediationSteps = @(
                "1. ENGAGE LEGAL IMMEDIATELY: Get written approval from legal counsel before any migration of held mailboxes"
                "2. DOCUMENT HOLDS: Get-Mailbox -ResultSize Unlimited | Where {`$_.LitigationHoldEnabled} | Select UserPrincipalName,LitigationHoldDate,LitigationHoldOwner,LitigationHoldDuration"
                "3. EXPORT HOLD DETAILS: Get-Mailbox user@domain.com | FL *Litigation*,*Hold* > LitigationHold_user.txt"
                "4. CREATE CHAIN OF CUSTODY: Document migration timeline, who performed it, verification steps, with timestamps"
                "5. ENABLE HOLD IN TARGET FIRST: Before migration, enable hold in target: Set-Mailbox -Identity user@target.com -LitigationHoldEnabled `$true -LitigationHoldDuration <days>"
                "6. MIGRATE WITH VERIFICATION: After migration, verify item counts match between source and target"
                "7. VERIFY HOLD ACTIVE: Get-Mailbox user@target.com | Select LitigationHoldEnabled,LitigationHoldDate,LitigationHoldDuration"
                "8. RETAIN SOURCE: Keep source mailbox in soft-delete or as inactive mailbox for legal protection period"
                "9. DOCUMENT COMPLETION: Provide legal with signed attestation that hold was maintained throughout migration"
            )
            Tools = @("Exchange Online PowerShell", "eDiscovery tools", "Documentation templates")
            EstimatedEffort = "Requires legal coordination - add 1-2 weeks for approvals"
            Prerequisites = @("Legal approval obtained", "Chain of custody documentation prepared", "Target holds configured")
        }
        @{
            Id          = "CP-002"
            Category    = "Compliance"
            Name        = "Active eDiscovery Cases"
            Condition   = { param($data)
                $data.Security.eDiscovery.Analysis.ActiveCases -gt 0
            }
            Severity    = "Critical"
            Description = "Active eDiscovery cases require legal coordination before migration"
            Recommendation = "Engage legal team. Document all case data and holds."
            RemediationSteps = @(
                "1. LIST ALL CASES: Connect to Compliance PowerShell; Get-ComplianceCase | Select Name,Status,CreatedDateTime,ClosedDateTime"
                "2. EXPORT CASE DETAILS: Get-ComplianceCase -Name 'CaseName' | Get-CaseHoldPolicy | Get-CaseHoldRule"
                "3. DOCUMENT CUSTODIANS: For each case, list all custodians and data sources under hold"
                "4. COORDINATE WITH LEGAL: Present migration plan to legal; get approval for each active case"
                "5. OPTIONS FOR ACTIVE CASES: A) Close case if litigation complete, B) Export all data before migration, C) Delay migration of affected custodians"
                "6. EXPORT SEARCH RESULTS: If needed, export search results before migration using New-ComplianceSearchAction -Export"
                "7. RECREATE IN TARGET: Create new eDiscovery case in target; add same custodians after migration"
                "8. HISTORICAL NOTE: eDiscovery cases do NOT migrate - only the content. Cases must be recreated."
                "9. VALIDATE: Run comparison searches in target to verify all custodian data migrated successfully"
            )
            Tools = @("Security & Compliance PowerShell", "Microsoft Purview", "eDiscovery Export Tool")
            EstimatedEffort = "Variable - depends on legal requirements. Plan 2-4 weeks minimum"
            Prerequisites = @("Legal team engaged", "Case inventory complete", "Export strategy defined")
        }
        @{
            Id          = "CP-003"
            Category    = "Compliance"
            Name        = "Sensitivity Labels with Protection"
            Condition   = { param($data)
                $data.Security.SensitivityLabels.Analysis.ProtectedLabels -gt 0
            }
            Severity    = "Critical"
            Description = "Protected content requires Azure RMS migration planning"
            Recommendation = "Plan for label GUID preservation or content re-protection."
            RemediationSteps = @(
                "1. INVENTORY LABELS: Get-Label | Select DisplayName,Guid,ContentType,Settings | Export-Csv SensitivityLabels.csv"
                "2. IDENTIFY PROTECTED CONTENT: Use Content Search to find labeled documents: 'InformationProtectionLabelId:<GUID>'"
                "3. EXPORT LABEL CONFIGURATION: Get-Label -Identity 'Label Name' | ConvertTo-Json -Depth 10 > LabelConfig.json"
                "4. CHOOSE MIGRATION APPROACH: Option A) Re-protect content with new labels (requires bulk relabeling) Option B) Maintain dual-key access during transition"
                "5. RECREATE LABELS IN TARGET: Create matching labels in target tenant with same protection settings"
                "6. FOR OPTION A - BULK RELABEL: Use Set-AIPFileLabel or Microsoft 365 Auto-labeling to apply new labels post-migration"
                "7. FOR ENCRYPTED CONTENT: Users may need to open and re-save documents after migration for new protection"
                "8. AIP SCANNER: Deploy AIP scanner to identify and report on all protected content: Install-AIPScanner; Set-AIPScannerConfiguration"
                "9. VALIDATE: Spot-check protected documents can be opened by target tenant users"
            )
            Tools = @("Azure Information Protection PowerShell", "AIP Scanner", "Microsoft Purview", "Content Search")
            EstimatedEffort = "2-4 weeks for planning and re-protection if needed"
            Prerequisites = @("Label inventory complete", "Protection templates documented", "Target labels created")
        }
        @{
            Id          = "CP-004"
            Category    = "Compliance"
            Name        = "Retention Policies Active"
            Condition   = { param($data)
                $data.Security.RetentionPolicies.Analysis.EnabledPolicies -gt 0
            }
            Severity    = "High"
            Description = "Retention policies must be recreated to maintain compliance"
            Recommendation = "Document all policies. Recreate before content migration."
            RemediationSteps = @(
                "1. EXPORT ALL POLICIES: Get-RetentionCompliancePolicy | ConvertTo-Json -Depth 10 | Out-File RetentionPolicies.json"
                "2. EXPORT RETENTION RULES: Get-RetentionComplianceRule | ConvertTo-Json -Depth 10 | Out-File RetentionRules.json"
                "3. DOCUMENT SCOPE: For each policy, document which locations/users/groups are included"
                "4. CREATE IN TARGET FIRST: Policies must exist in target before content migration to maintain compliance"
                "5. CREATE POLICY: New-RetentionCompliancePolicy -Name 'PolicyName' -ExchangeLocation All -SharePointLocation All"
                "6. CREATE RULE: New-RetentionComplianceRule -Policy 'PolicyName' -RetentionDuration 2555 -RetentionComplianceAction Keep"
                "7. VERIFY POLICY STATUS: Get-RetentionCompliancePolicy 'PolicyName' | Select DistributionStatus (must be 'Success')"
                "8. WAIT FOR PROPAGATION: Policies take 24-48 hours to fully apply. Do not migrate until status is Success"
                "9. POST-MIGRATION: Verify retention labels on migrated content using Get-ComplianceTag"
            )
            Tools = @("Security & Compliance PowerShell", "Microsoft Purview Portal")
            EstimatedEffort = "2-3 days for policy recreation plus 24-48hr propagation"
            Prerequisites = @("Compliance admin access to target", "Policy documentation complete")
        }
        @{
            Id          = "CP-005"
            Category    = "Compliance"
            Name        = "DLP Policies Enforced"
            Condition   = { param($data)
                $data.Security.DLPPolicies.Analysis.EnforcedPolicies -gt 0
            }
            Severity    = "High"
            Description = "DLP policies in enforcement mode need recreation"
            Recommendation = "Export DLP configurations. Deploy in test mode initially."
            RemediationSteps = @(
                "1. EXPORT DLP POLICIES: Get-DlpCompliancePolicy | ConvertTo-Json -Depth 10 | Out-File DLPPolicies.json"
                "2. EXPORT DLP RULES: Get-DlpComplianceRule | ConvertTo-Json -Depth 10 | Out-File DLPRules.json"
                "3. DOCUMENT SENSITIVE INFO TYPES: Get-DlpSensitiveInformationType | Where {`$_.Publisher -ne 'Microsoft'} | Export-Csv CustomSITs.csv"
                "4. RECREATE CUSTOM SITS FIRST: Custom sensitive info types must exist before policies that use them"
                "5. CREATE POLICIES IN TEST MODE: New-DlpCompliancePolicy -Name 'Policy' -Mode TestWithNotifications -ExchangeLocation All"
                "6. CREATE RULES: New-DlpComplianceRule -Policy 'Policy' -ContentContainsSensitiveInformation @{Name='Credit Card Number';minCount=1}"
                "7. MONITOR TEST RESULTS: Review DLP reports in Purview for 1-2 weeks to verify expected behavior"
                "8. ENABLE ENFORCEMENT: Set-DlpCompliancePolicy -Identity 'Policy' -Mode Enable"
                "9. COMMUNICATE: Notify users that DLP is active - they may see different policy tips than before"
            )
            Tools = @("Security & Compliance PowerShell", "Microsoft Purview", "DLP Activity Explorer")
            EstimatedEffort = "1-2 weeks including test mode monitoring"
            Prerequisites = @("Custom sensitive info types recreated", "DLP admin access to target")
        }

        # Integration Rules
        @{
            Id          = "IN-001"
            Category    = "Integration"
            Name        = "Enterprise Applications"
            Condition   = { param($data)
                $data.EntraID.Applications.Analysis.TotalServicePrincipals -gt 50
            }
            Severity    = "High"
            Description = "Large number of enterprise apps require SSO reconfiguration"
            Recommendation = "Document SSO configurations. Plan phased app migration."
            RemediationSteps = @(
                "1. INVENTORY APPS: Get-MgServicePrincipal -All | Select DisplayName,AppId,SignInAudience,PreferredSingleSignOnMode | Export-Csv EnterpriseApps.csv"
                "2. CATEGORIZE BY SSO TYPE: Group apps by SAML, OIDC, Password-based, Linked, or Disabled SSO"
                "3. EXPORT SAML CONFIG: For SAML apps, export metadata: Get-MgServicePrincipal -ServicePrincipalId <id> | Select-Object -ExpandProperty Saml*"
                "4. DOCUMENT CLAIM MAPPINGS: Get-MgServicePrincipalClaimMappingPolicy for custom claims"
                "5. PRIORITIZE BY USAGE: Check Sign-in logs to identify most-used apps for priority migration"
                "6. RECREATE IN TARGET: Register new enterprise app in target: New-MgServicePrincipal -AppId <gallery-app-id>"
                "7. CONFIGURE SSO: For SAML apps, update target with: Update-MgServicePrincipal with SAML metadata URLs"
                "8. UPDATE IDP SETTINGS: In each SaaS app admin console, add target tenant as additional IdP or replace source"
                "9. TEST WITH PILOT: Assign test users and verify SSO works before full rollout"
                "10. MIGRATE IN WAVES: Move user assignments in batches, monitoring for SSO failures"
            )
            Tools = @("Microsoft Graph PowerShell", "App admin consoles", "Azure AD Sign-in logs")
            EstimatedEffort = "1-2 days per complex app, less for simple apps"
            Prerequisites = @("App owner contacts identified", "SSO configurations documented", "Test users ready")
        }
        @{
            Id          = "IN-002"
            Category    = "Integration"
            Name        = "Custom Teams Apps"
            Condition   = { param($data) $data.Teams.Apps.Analysis.CustomApps -gt 0 }
            Severity    = "High"
            Description = "Custom Teams apps must be republished in target tenant"
            Recommendation = "Export app packages. Update manifests for target tenant."
            RemediationSteps = @(
                "1. LIST CUSTOM APPS: Get-TeamsApp | Where {`$_.DistributionMethod -eq 'organization'} | Select DisplayName,Id,ExternalId"
                "2. EXPORT APP PACKAGES: Download .zip packages from Teams Admin Center > Manage apps > Select app > Download"
                "3. EXTRACT MANIFEST: Unzip package, open manifest.json to review app configuration"
                "4. UPDATE MANIFEST: Change 'id' to new GUID, update any tenant-specific URLs or IDs"
                "5. UPDATE BOT REGISTRATION: If app has bot, create new Bot registration in target tenant Azure"
                "6. UPDATE TAB URLS: If app has configurable tabs, update contentUrl and websiteUrl to target"
                "7. REPACKAGE: Zip updated files (manifest.json, icons) into new package"
                "8. PUBLISH TO TARGET: Teams Admin Center > Manage apps > Upload > Select repackaged .zip"
                "9. SET PERMISSIONS: Configure app permission policies: New-TeamsAppPermissionPolicy"
                "10. DEPLOY TO USERS: Add to setup policy or install directly: Install-TeamsApp -AppId <newAppId> -UserId <userId>"
            )
            Tools = @("Teams Admin Center", "Teams PowerShell", "Text editor for manifest", "Azure Bot Service")
            EstimatedEffort = "2-4 hours per custom app"
            Prerequisites = @("App source code/packages available", "Bot registrations recreated if needed")
        }
        @{
            Id          = "IN-003"
            Category    = "Integration"
            Name        = "Power Platform Environments"
            Condition   = { param($data)
                $data.Dynamics365.Environments.Analysis.TotalEnvironments -gt 0
            }
            Severity    = "Critical"
            Description = "Dynamics 365/Power Platform requires specialized migration"
            Recommendation = "Plan dedicated Power Platform migration project."
            RemediationSteps = @(
                "1. INVENTORY ENVIRONMENTS: Use Power Platform Admin Center or: Get-AdminPowerAppEnvironment | Select DisplayName,EnvironmentType,IsDefault"
                "2. ASSESS DATAVERSE: For each environment, document tables, rows, storage used"
                "3. EXPORT SOLUTIONS: Export all managed/unmanaged solutions: Export-CrmSolution -SolutionName 'MySolution' -Managed"
                "4. DOCUMENT CONNECTIONS: List all connections and connection references used by apps/flows"
                "5. CREATE TARGET ENVIRONMENTS: In target tenant, create matching environments with same security roles"
                "6. IMPORT SOLUTIONS: Import-CrmSolution -SolutionFilePath solution.zip (import unmanaged for customization)"
                "7. RECREATE CONNECTIONS: Create new connections in target using target tenant credentials"
                "8. DATA MIGRATION: Use Dataverse Data Export Service or third-party tools for large data volumes"
                "9. UPDATE ENVIRONMENT VARIABLES: Update any environment-specific variables to target values"
                "10. VALIDATE: Test all apps and flows in target environment thoroughly before cutover"
            )
            Tools = @("Power Platform Admin Center", "Power Platform CLI (pac)", "XrmToolBox", "Dataverse Web API")
            EstimatedEffort = "2-6 weeks depending on complexity - this is often a separate project"
            Prerequisites = @("Power Platform licenses in target", "Environment strategy defined", "Data migration tools selected")
        }
        @{
            Id          = "IN-004"
            Category    = "Integration"
            Name        = "Power BI Gateways"
            Condition   = { param($data) $data.PowerBI.Gateways.Analysis.TotalGateways -gt 0 }
            Severity    = "Critical"
            Description = "On-premises gateways must be reinstalled for target tenant"
            Recommendation = "Document gateway data sources. Plan reinstallation."
            RemediationSteps = @(
                "1. INVENTORY GATEWAYS: In Power BI Admin Portal, list all gateways and their data sources"
                "2. DOCUMENT DATA SOURCES: For each gateway, list: data source type, connection string, credentials used"
                "3. EXPORT REPORTS: Download .pbix files for all reports using the gateway"
                "4. INSTALL NEW GATEWAY: On gateway server, install new gateway registered to target tenant"
                "5. IMPORTANT: You can have 2 gateways on same server (one per tenant) during transition"
                "6. CONFIGURE DATA SOURCES: In target tenant Power BI, add same data sources to new gateway"
                "7. UPDATE CREDENTIALS: Re-enter all data source credentials (they don't migrate)"
                "8. REPUBLISH REPORTS: Upload .pbix files to target tenant workspaces"
                "9. REBIND DATASETS: In dataset settings, bind to new gateway and data sources"
                "10. TEST REFRESH: Trigger manual refresh for each dataset to verify gateway connectivity"
                "11. REMOVE OLD GATEWAY: After validation, uninstall source tenant gateway from server"
            )
            Tools = @("On-premises data gateway installer", "Power BI Service", "Power BI Admin Portal")
            EstimatedEffort = "1-2 days per gateway"
            Prerequisites = @("Gateway server access", "Data source credentials available", "Reports backed up")
        }
        @{
            Id          = "IN-005"
            Category    = "Integration"
            Name        = "Expiring App Credentials"
            Condition   = { param($data)
                $data.EntraID.Applications.Analysis.ExpiringSecrets90Days -gt 0 -or
                $data.EntraID.Applications.Analysis.ExpiringCerts90Days -gt 0
            }
            Severity    = "High"
            Description = "Application credentials expiring within 90 days"
            Recommendation = "Renew credentials before migration."
            RemediationSteps = @(
                "1. IDENTIFY EXPIRING CREDENTIALS: Get-MgApplication -All | ForEach { `$app = `$_; `$app.PasswordCredentials | Where {`$_.EndDateTime -lt (Get-Date).AddDays(90)} | Select @{N='AppName';E={`$app.DisplayName}},KeyId,EndDateTime}"
                "2. ALSO CHECK CERTIFICATES: Same query with .KeyCredentials instead of .PasswordCredentials"
                "3. CONTACT APP OWNERS: Notify application owners of pending expiration"
                "4. FOR SECRETS - ADD NEW: Add-MgApplicationPassword -ApplicationId <appId> -PasswordCredential @{DisplayName='New Secret';EndDateTime=(Get-Date).AddYears(2)}"
                "5. FOR CERTIFICATES - ADD NEW: Add-MgApplicationKey -ApplicationId <appId> -KeyCredential @{Type='AsymmetricX509Cert';Usage='Verify';Key=[Convert]::ToBase64String((Get-Content cert.cer -Encoding Byte))}"
                "6. UPDATE APPLICATIONS: Provide new credential to application owner to update their code/config"
                "7. TEST AUTHENTICATION: Verify application can still authenticate with new credential"
                "8. REMOVE OLD CREDENTIAL: After verification, remove expiring credential: Remove-MgApplicationPassword -ApplicationId <appId> -KeyId <keyId>"
                "9. POST-MIGRATION: In target tenant, apps will need NEW credentials anyway - coordinate with owners"
            )
            Tools = @("Microsoft Graph PowerShell", "Azure Portal", "Certificate management tools")
            EstimatedEffort = "2-4 hours per application"
            Prerequisites = @("App owner contact information", "New certificates generated if needed")
        }

        # Infrastructure Rules
        @{
            Id          = "IF-001"
            Category    = "Infrastructure"
            Name        = "Federation Active"
            Condition   = { param($data)
                $data.HybridIdentity.Federation.Analysis.FederationInUse
            }
            Severity    = "Critical"
            Description = "Federated domains require careful cutover planning"
            Recommendation = "Document ADFS config. Plan federation cutover strategy."
            RemediationSteps = @(
                "1. DOCUMENT ADFS CONFIG: Export ADFS configuration using: Export-FederationConfiguration -Path C:\ADFSConfig"
                "2. LIST FEDERATED DOMAINS: Get-MgDomain | Where {`$_.AuthenticationType -eq 'Federated'} | Select Id,AuthenticationType"
                "3. DOCUMENT RELYING PARTY TRUSTS: Get-AdfsRelyingPartyTrust | Export-Clixml RelyingPartyTrusts.xml"
                "4. CHOOSE STRATEGY: Option A) Convert to managed authentication before migration, Option B) Set up new ADFS for target"
                "5. FOR MANAGED CONVERSION: Convert domain to managed first: Set-MgDomain -DomainId 'domain.com' -AuthenticationType Managed"
                "6. STAGED ROLLOUT: Use staged rollout to test managed auth with pilot users before full conversion"
                "7. FOR NEW ADFS: Install new ADFS farm or configure existing for target tenant"
                "8. UPDATE DNS: Change federation endpoint DNS to point to target tenant ADFS"
                "9. CUTOVER: Convert-MgDomainFederatedToManaged or configure new federation trust"
                "10. VALIDATE: Test sign-in for federated users in both internal and external networks"
            )
            Tools = @("ADFS PowerShell", "Microsoft Graph PowerShell", "DNS management")
            EstimatedEffort = "1-2 weeks for planning and execution"
            Prerequisites = @("ADFS admin access", "DNS change approval", "User communication prepared")
        }
        @{
            Id          = "IF-002"
            Category    = "Infrastructure"
            Name        = "Pass-Through Authentication"
            Condition   = { param($data)
                $data.HybridIdentity.AuthenticationMethods.Analysis.PTAEnabled
            }
            Severity    = "High"
            Description = "PTA agents must be deployed for target tenant"
            Recommendation = "Plan PTA agent deployment. Consider staged rollout."
            RemediationSteps = @(
                "1. INVENTORY PTA AGENTS: Get-MgDirectoryOnPremisesPublishingProfileAgentGroup | Get-MgDirectoryOnPremisesPublishingProfileAgentGroupAgent"
                "2. DOCUMENT AGENT SERVERS: List all servers running PTA agents with their network locations"
                "3. DOWNLOAD TARGET AGENT: From Entra ID portal > Hybrid management > PTA, download agent installer for target tenant"
                "4. INSTALL ON SAME SERVERS: You CAN run PTA agents for both tenants on same server during transition"
                "5. REGISTER AGENT: During install, authenticate with target tenant Global Admin"
                "6. VERIFY AGENT STATUS: In target Entra portal, verify agent shows as Active"
                "7. CONFIGURE STAGED ROLLOUT: Enable staged rollout to test PTA with pilot users"
                "8. CUTOVER: After validation, update AAD Connect or Entra Cloud Sync to point to target"
                "9. REMOVE OLD AGENTS: Once complete, uninstall source tenant PTA agents"
                "10. VALIDATE: Test authentication from various network locations"
            )
            Tools = @("PTA Agent installer", "Entra ID Portal", "Microsoft Graph PowerShell")
            EstimatedEffort = "1-2 days for deployment and testing"
            Prerequisites = @("Server access for agent deployment", "Target tenant Global Admin", "Network connectivity verified")
        }
        @{
            Id          = "IF-003"
            Category    = "Infrastructure"
            Name        = "Directory Sync Active"
            Condition   = { param($data)
                $data.HybridIdentity.AADConnect.Configuration.OnPremisesSyncEnabled
            }
            Severity    = "Critical"
            Description = "Azure AD Connect configuration must be addressed"
            Recommendation = "Plan AAD Connect migration: new install or staged approach."
            RemediationSteps = @(
                "1. EXPORT CURRENT CONFIG: On AAD Connect server: Get-ADSyncServerConfiguration -Path C:\AADConnectExport"
                "2. DOCUMENT SYNC RULES: Get-ADSyncRule | Export-Clixml SyncRules.xml"
                "3. DOCUMENT FILTERING: Get-ADSyncConnectorRunStatus; note OU filtering and attribute filtering"
                "4. CHOOSE APPROACH: Option A) New AAD Connect server for target, Option B) Swing migration on existing server"
                "5. FOR NEW SERVER (RECOMMENDED): Install AAD Connect on new server pointing to target tenant"
                "6. CONFIGURE SAME FILTERING: Apply same OU/attribute filtering as source"
                "7. APPLY CUSTOM SYNC RULES: Recreate custom sync rules using Set-ADSyncRule"
                "8. STAGING MODE: Initially enable staging mode to verify sync without writing to target"
                "9. VERIFY IN STAGING: Compare user/group counts and attributes between source and staging output"
                "10. CUTOVER SEQUENCE: Stop source sync -> Disable staging in target -> Verify users sync correctly"
                "11. POST-MIGRATION: Remove AAD Connect from source tenant after 72-hour waiting period"
            )
            Tools = @("Azure AD Connect", "ADSyncTools module", "Microsoft Graph PowerShell")
            EstimatedEffort = "3-5 days including staging validation"
            Prerequisites = @("New AAD Connect server provisioned OR existing server available", "Target tenant GA credentials")
        }
        @{
            Id          = "IF-004"
            Category    = "Infrastructure"
            Name        = "Mail Flow Connectors"
            Condition   = { param($data)
                ($data.Exchange.TransportConfig.InboundConnectors | Measure-Object).Count -gt 0 -or
                ($data.Exchange.TransportConfig.OutboundConnectors | Measure-Object).Count -gt 0
            }
            Severity    = "High"
            Description = "Mail flow connectors require recreation in target"
            Recommendation = "Document connector configurations. Plan mail flow cutover."
            RemediationSteps = @(
                "1. EXPORT INBOUND CONNECTORS: Get-InboundConnector | ConvertTo-Json -Depth 5 | Out-File InboundConnectors.json"
                "2. EXPORT OUTBOUND CONNECTORS: Get-OutboundConnector | ConvertTo-Json -Depth 5 | Out-File OutboundConnectors.json"
                "3. DOCUMENT CERTIFICATES: For TLS connectors, note certificate requirements and thumbprints"
                "4. CREATE INBOUND IN TARGET: New-InboundConnector -Name 'Partner Inbound' -SenderDomains 'partner.com' -ConnectorType Partner"
                "5. CREATE OUTBOUND IN TARGET: New-OutboundConnector -Name 'Partner Outbound' -RecipientDomains 'partner.com' -SmartHosts 'mail.partner.com'"
                "6. CONFIGURE TLS: Set-InboundConnector -Identity 'Connector' -RequireTls `$true -TlsSenderCertificateName 'certificate'"
                "7. UPDATE SPF/DKIM: Ensure target tenant's IPs are in partner SPF records if required"
                "8. TEST MAIL FLOW: Send test messages through each connector and verify delivery"
                "9. COORDINATE WITH PARTNERS: Notify partners of new IP ranges/certificates for their connector configs"
                "10. CUTOVER: Update MX records -> Wait for propagation -> Verify inbound mail flow"
            )
            Tools = @("Exchange Online PowerShell", "Message Trace", "MX Toolbox for verification")
            EstimatedEffort = "1-3 days depending on number of connectors"
            Prerequisites = @("Partner contacts for coordination", "Certificates installed", "IP allowlists updated")
        }

        # Operations Rules
        @{
            Id          = "OP-001"
            Category    = "Operations"
            Name        = "Teams Phone System"
            Condition   = { param($data)
                $data.Teams.PhoneSystem.Analysis.PhoneSystemEnabled
            }
            Severity    = "Critical"
            Description = "Teams Phone System requires dedicated migration project"
            Recommendation = "Plan separate telephony migration. Document all configurations."
            RemediationSteps = @(
                "1. INVENTORY PHONE NUMBERS: Get-CsPhoneNumberAssignment | Export-Csv PhoneNumbers.csv"
                "2. DOCUMENT VOICE POLICIES: Get-CsTeamsCallingPolicy, Get-CsTeamsCallParkPolicy, Get-CsTeamsCallHoldPolicy | Export-Clixml VoicePolicies.xml"
                "3. EXPORT CALL QUEUES: Get-CsCallQueue | ConvertTo-Json -Depth 5 | Out-File CallQueues.json"
                "4. EXPORT AUTO ATTENDANTS: Get-CsAutoAttendant | ConvertTo-Json -Depth 5 | Out-File AutoAttendants.json"
                "5. DOCUMENT DIAL PLANS: Get-CsTenantDialPlan | Export-Clixml DialPlans.xml"
                "6. PORT NUMBERS TO TARGET: Work with Microsoft or carrier to port phone numbers to target tenant"
                "7. RECREATE RESOURCE ACCOUNTS: New-CsOnlineApplicationInstance for each call queue/auto attendant"
                "8. RECREATE CALL QUEUES: New-CsCallQueue with same configuration as source"
                "9. RECREATE AUTO ATTENDANTS: New-CsAutoAttendant with same greetings, menus, and routing"
                "10. ASSIGN NUMBERS: Set-CsPhoneNumberAssignment -Identity user@target.com -PhoneNumber +1555... -PhoneNumberType DirectRouting"
                "11. COORDINATE CUTOVER: Schedule brief telephony downtime during number porting"
            )
            Tools = @("Teams PowerShell", "Teams Admin Center", "Microsoft Calling Plans or Direct Routing SBC")
            EstimatedEffort = "2-4 weeks - treat as separate project"
            Prerequisites = @("Phone numbers portable to target", "Voice licenses in target", "SBC configured if Direct Routing")
        }
        @{
            Id          = "OP-002"
            Category    = "Operations"
            Name        = "Private Channels"
            Condition   = { param($data)
                $data.Teams.Teams.Analysis.PrivateChannels -gt 0
            }
            Severity    = "High"
            Description = "Private channels have separate SharePoint sites"
            Recommendation = "Plan private channel migration separately. Document memberships."
            RemediationSteps = @(
                "1. INVENTORY PRIVATE CHANNELS: Get-Team | Get-TeamChannel | Where {`$_.MembershipType -eq 'Private'} | Export-Csv PrivateChannels.csv"
                "2. DOCUMENT MEMBERSHIPS: For each private channel, export members: Get-TeamChannelUser -GroupId <id> -DisplayName 'Channel'"
                "3. LOCATE SHAREPOINT SITES: Private channels have sites at tenant.sharepoint.com/sites/team-channelname"
                "4. MIGRATE TEAM FIRST: Migrate parent Team before private channels"
                "5. RECREATE PRIVATE CHANNELS: In target Team: New-TeamChannel -GroupId <id> -DisplayName 'Channel' -MembershipType Private"
                "6. ADD MEMBERS: Add-TeamChannelUser -GroupId <id> -DisplayName 'Channel' -User user@target.com"
                "7. MIGRATE SHAREPOINT CONTENT: Use SPMT or ShareGate to migrate private channel SharePoint site separately"
                "8. SET PERMISSIONS: Verify private channel permissions match source"
                "9. SHARED CHANNELS NOTE: Shared channels (B2B) require separate handling with external org coordination"
                "10. VALIDATE: Verify content and membership in target private channels"
            )
            Tools = @("Teams PowerShell", "SharePoint Migration Tool", "ShareGate")
            EstimatedEffort = "Additional 2-4 hours per private channel"
            Prerequisites = @("Parent Team migrated", "Private channel members migrated to target")
        }
        @{
            Id          = "OP-003"
            Category    = "Operations"
            Name        = "Hub Sites Configured"
            Condition   = { param($data)
                $data.SharePoint.HubSites.TotalHubs -gt 0
            }
            Severity    = "Medium"
            Description = "Hub site associations must be recreated"
            Recommendation = "Document hub site hierarchy. Recreate before site migration."
            RemediationSteps = @(
                "1. INVENTORY HUB SITES: Get-PnPHubSite | Select Title,SiteUrl,HubSiteId | Export-Csv HubSites.csv"
                "2. DOCUMENT ASSOCIATIONS: Get-PnPTenantSite | Where {`$_.HubSiteId -ne '00000000-0000-0000-0000-000000000000'} | Select Url,HubSiteId"
                "3. EXPORT HUB NAVIGATION: Get-PnPNavigationNode -Location TopNavigationBar from each hub"
                "4. CREATE HUB SITES IN TARGET FIRST: Register-PnPHubSite -Site 'https://target.sharepoint.com/sites/hub'"
                "5. SET HUB PROPERTIES: Set-PnPHubSite -Identity <url> -Title 'Hub Title' -LogoUrl <logo>"
                "6. MIGRATE CONTENT: Migrate hub site content using SPMT or ShareGate"
                "7. MIGRATE ASSOCIATED SITES: Migrate all sites associated with this hub"
                "8. ASSOCIATE SITES: Add-PnPHubSiteAssociation -Site 'https://target.sharepoint.com/sites/spoke' -HubSite 'https://target.sharepoint.com/sites/hub'"
                "9. RECREATE NAVIGATION: Add-PnPNavigationNode to recreate hub navigation"
                "10. VALIDATE: Verify hub associations and navigation work correctly in target"
            )
            Tools = @("PnP PowerShell", "SharePoint Admin Center")
            EstimatedEffort = "1-2 hours per hub site plus associated site migrations"
            Prerequisites = @("Target hub site created before associated sites migrated")
        }

        # Additional Identity Rules
        @{
            Id          = "ID-006"
            Category    = "Identity"
            Name        = "Cloud-Only Users with ImmutableId"
            Condition   = { param($data)
                $data.EntraID.Users.Users | Where-Object {
                    -not $_.OnPremisesSyncEnabled -and $_.OnPremisesImmutableId
                } | Measure-Object | Select-Object -ExpandProperty Count
            }
            Severity    = "High"
            Description = "Cloud-only users with ImmutableId set - may indicate previous sync or manual configuration"
            Recommendation = "Review and document these users. ImmutableId preservation requires special handling during migration."
        }
        @{
            Id          = "ID-007"
            Category    = "Identity"
            Name        = "Users with External Identity Providers"
            Condition   = { param($data)
                $data.EntraID.Users.Users | Where-Object {
                    $_.Identities | Where-Object { $_.SignInType -eq "federated" }
                } | Measure-Object | Select-Object -ExpandProperty Count
            }
            Severity    = "High"
            Description = "Users with external identity provider federation require special migration planning"
            Recommendation = "Document external IdP configurations. Plan for IdP trust recreation in target tenant."
        }
        @{
            Id          = "ID-008"
            Category    = "Identity"
            Name        = "Privileged Identity Management (PIM) Active"
            Condition   = { param($data)
                $data.EntraID.Roles.Analysis.PrivilegedRoleAssignments -gt 0
            }
            Severity    = "High"
            Description = "PIM role assignments detected - these need careful recreation in target"
            Recommendation = "Export PIM configurations. Plan for PIM setup before role assignment migration."
        }
        @{
            Id          = "ID-009"
            Category    = "Identity"
            Name        = "Service Accounts Without Password Expiry"
            Condition   = { param($data)
                $data.EntraID.Users.Users | Where-Object {
                    $_.PasswordPolicies -match "DisablePasswordExpiration"
                } | Measure-Object | Select-Object -ExpandProperty Count
            }
            Severity    = "Medium"
            Description = "Service accounts with disabled password expiry need security review"
            Recommendation = "Document service accounts. Review if managed identities can replace them in target."
        }
        @{
            Id          = "ID-010"
            Category    = "Identity"
            Name        = "Multiple UPN Suffixes"
            Condition   = { param($data)
                ($data.EntraID.Users.UPNSuffixes | Measure-Object).Count -gt 3
            }
            Severity    = "Medium"
            Description = "Multiple UPN suffixes in use - all domains must be verified in target tenant"
            Recommendation = "Document all UPN suffixes. Plan domain verification sequence in target."
        }
        @{
            Id          = "ID-011"
            Category    = "Identity"
            Name        = "Azure AD B2B Direct Federation"
            Condition   = { param($data)
                $data.EntraID.TenantInfo.VerifiedDomains | Where-Object { $_.AuthType -eq "Federated" }
            }
            Severity    = "Critical"
            Description = "B2B direct federation requires partner coordination during migration"
            Recommendation = "Document all B2B federation partners. Coordinate cutover timing with partners."
        }

        # Additional Data Rules
        @{
            Id          = "DT-005"
            Category    = "Data"
            Name        = "In-Place Archive Mailboxes"
            Condition   = { param($data)
                $data.Exchange.Mailboxes.Analysis.ArchiveEnabled -gt 100
            }
            Severity    = "High"
            Description = "Large number of archive mailboxes requires extended migration windows"
            Recommendation = "Plan incremental archive migration. Consider archive-first approach for large mailboxes."
        }
        @{
            Id          = "DT-006"
            Category    = "Data"
            Name        = "OneDrive Large Storage Users"
            Condition   = { param($data)
                ($data.SharePoint.OneDrive.Users | Where-Object { $_.StorageUsedMB -gt 50000 } | Measure-Object).Count -gt 0
            }
            Severity    = "Medium"
            Description = "Users with OneDrive storage exceeding 50GB require extended migration"
            Recommendation = "Identify large OneDrive users. Plan pre-staging approach for data migration."
        }
        @{
            Id          = "DT-007"
            Category    = "Data"
            Name        = "SharePoint Custom Solutions"
            Condition   = { param($data)
                $data.SharePoint.Sites.Analysis.CustomSolutions -gt 0
            }
            Severity    = "Critical"
            Description = "Custom SharePoint solutions (SPFx, Apps) require redevelopment or migration"
            Recommendation = "Inventory all custom solutions. Plan for solution migration or rebuild in target."
        }
        @{
            Id          = "DT-008"
            Category    = "Data"
            Name        = "SharePoint Workflows Active"
            Condition   = { param($data)
                $data.SharePoint.Sites.Analysis.WorkflowsActive -gt 0
            }
            Severity    = "High"
            Description = "Active SharePoint workflows (2010/2013/Power Automate) need migration planning"
            Recommendation = "Document all workflows. Plan for Power Automate rebuild of legacy workflows."
        }
        @{
            Id          = "DT-009"
            Category    = "Data"
            Name        = "External Sharing Enabled"
            Condition   = { param($data)
                $data.SharePoint.SharingSettings.ExternalSharingEnabled
            }
            Severity    = "Medium"
            Description = "External sharing configuration must be replicated in target tenant"
            Recommendation = "Document sharing policies. External shares will need re-establishing post-migration."
        }
        @{
            Id          = "DT-010"
            Category    = "Data"
            Name        = "Teams with External Members"
            Condition   = { param($data)
                $data.Teams.Teams.Analysis.TeamsWithGuests -gt 0
            }
            Severity    = "Medium"
            Description = "Teams with external members require guest reinvitation in target"
            Recommendation = "Document Teams guest membership. Plan guest re-invitation process."
        }

        # Additional Compliance Rules
        @{
            Id          = "CP-006"
            Category    = "Compliance"
            Name        = "Information Barriers Configured"
            Condition   = { param($data)
                $data.Security.InformationBarriers.Analysis.PoliciesEnabled -gt 0
            }
            Severity    = "Critical"
            Description = "Information barriers require recreation before user migration"
            Recommendation = "Export IB policies and segments. Recreate in target before migrating affected users."
        }
        @{
            Id          = "CP-007"
            Category    = "Compliance"
            Name        = "Communication Compliance Policies"
            Condition   = { param($data)
                $data.Security.CommunicationCompliance.Analysis.ActivePolicies -gt 0
            }
            Severity    = "High"
            Description = "Communication compliance policies must be recreated in target"
            Recommendation = "Document all CC policies including custom classifiers. Plan recreation in target."
        }
        @{
            Id          = "CP-008"
            Category    = "Compliance"
            Name        = "Records Management Labels"
            Condition   = { param($data)
                $data.Security.RecordsManagement.Analysis.RecordLabels -gt 0
            }
            Severity    = "Critical"
            Description = "Records management labels with regulatory requirements need careful migration"
            Recommendation = "Engage legal/compliance team. Document retention schedules and disposition."
        }
        @{
            Id          = "CP-009"
            Category    = "Compliance"
            Name        = "Insider Risk Management"
            Condition   = { param($data)
                $data.Security.InsiderRisk.Analysis.PoliciesActive -gt 0
            }
            Severity    = "High"
            Description = "Insider risk policies and alerts need recreation in target"
            Recommendation = "Document IRM policies. Historical alerts cannot be migrated. Plan fresh baseline."
        }
        @{
            Id          = "CP-010"
            Category    = "Compliance"
            Name        = "Audit Log Retention Policies"
            Condition   = { param($data)
                $data.Security.AuditLog.Analysis.CustomRetention
            }
            Severity    = "Medium"
            Description = "Custom audit log retention policies exist"
            Recommendation = "Document retention periods. Export audit logs before migration if required for compliance."
        }

        # Additional Integration Rules
        @{
            Id          = "IN-006"
            Category    = "Integration"
            Name        = "Power Automate Flows"
            Condition   = { param($data)
                $data.Dynamics365.PowerAutomate.Analysis.TotalFlows -gt 50
            }
            Severity    = "High"
            Description = "Large number of Power Automate flows require migration planning"
            Recommendation = "Inventory all flows. Plan for flow export/import or recreation. Check connection references."
        }
        @{
            Id          = "IN-007"
            Category    = "Integration"
            Name        = "Power Apps Applications"
            Condition   = { param($data)
                $data.Dynamics365.PowerApps.Analysis.TotalApps -gt 20
            }
            Severity    = "High"
            Description = "Power Apps require migration with data sources and connections"
            Recommendation = "Document all Power Apps and their data sources. Plan connection recreation in target."
        }
        @{
            Id          = "IN-008"
            Category    = "Integration"
            Name        = "Dataverse Environments"
            Condition   = { param($data)
                $data.Dynamics365.Environments.Analysis.DataverseEnvironments -gt 0
            }
            Severity    = "Critical"
            Description = "Dataverse environments with business data require dedicated migration"
            Recommendation = "Plan dedicated Dataverse migration project. Consider data volume and relationships."
        }
        @{
            Id          = "IN-009"
            Category    = "Integration"
            Name        = "Custom Connectors"
            Condition   = { param($data)
                $data.Dynamics365.Connectors.Analysis.CustomConnectors -gt 0
            }
            Severity    = "High"
            Description = "Custom connectors need recreation in target tenant"
            Recommendation = "Export custom connector definitions. Recreate and test before flow migration."
        }
        @{
            Id          = "IN-010"
            Category    = "Integration"
            Name        = "Azure Logic Apps Integration"
            Condition   = { param($data)
                $data.Dynamics365.LogicApps.Analysis.IntegratedApps -gt 0
            }
            Severity    = "Medium"
            Description = "Logic Apps with M365 connections need connection updates"
            Recommendation = "Document Logic App connections. Plan for connection recreation post-migration."
        }
        @{
            Id          = "IN-011"
            Category    = "Integration"
            Name        = "Third-Party MDM Integration"
            Condition   = { param($data)
                $data.EntraID.Devices.Analysis.ThirdPartyMDM
            }
            Severity    = "High"
            Description = "Third-party MDM integration requires reconfiguration"
            Recommendation = "Document MDM integration settings. Coordinate with MDM vendor for migration."
        }

        # Additional Infrastructure Rules
        @{
            Id          = "IF-005"
            Category    = "Infrastructure"
            Name        = "Password Hash Sync Enabled"
            Condition   = { param($data)
                $data.HybridIdentity.AuthenticationMethods.Analysis.PHSEnabled
            }
            Severity    = "Medium"
            Description = "Password Hash Sync configuration needs recreation for target tenant"
            Recommendation = "Document PHS configuration. Plan staged rollout in target tenant."
        }
        @{
            Id          = "IF-006"
            Category    = "Infrastructure"
            Name        = "Seamless SSO Configured"
            Condition   = { param($data)
                $data.HybridIdentity.AuthenticationMethods.Analysis.SeamlessSSOEnabled
            }
            Severity    = "High"
            Description = "Seamless SSO requires computer account recreation for target tenant"
            Recommendation = "Plan Seamless SSO cutover. New computer accounts needed in AD for target tenant."
        }
        @{
            Id          = "IF-007"
            Category    = "Infrastructure"
            Name        = "Multiple AAD Connect Servers"
            Condition   = { param($data)
                $data.HybridIdentity.AADConnect.Analysis.StagingServerCount -gt 0
            }
            Severity    = "Medium"
            Description = "Multiple AAD Connect servers (staging mode) exist"
            Recommendation = "Document all AAD Connect servers. Plan migration approach for sync infrastructure."
        }
        @{
            Id          = "IF-008"
            Category    = "Infrastructure"
            Name        = "Group Writeback Enabled"
            Condition   = { param($data)
                $data.HybridIdentity.AADConnect.Analysis.GroupWritebackEnabled
            }
            Severity    = "High"
            Description = "Group writeback to on-premises AD is configured"
            Recommendation = "Document writeback configuration. Plan for writeback setup in target tenant."
        }
        @{
            Id          = "IF-009"
            Category    = "Infrastructure"
            Name        = "Device Writeback Enabled"
            Condition   = { param($data)
                $data.HybridIdentity.DeviceWriteback.Configuration.Enabled
            }
            Severity    = "High"
            Description = "Device writeback to on-premises AD is configured"
            Recommendation = "Document device writeback. Plan for Windows Hello for Business migration."
        }
        @{
            Id          = "IF-010"
            Category    = "Infrastructure"
            Name        = "Exchange Hybrid Configuration"
            Condition   = { param($data)
                $data.Exchange.HybridConfig.Analysis.HybridEnabled
            }
            Severity    = "Critical"
            Description = "Exchange hybrid configuration requires careful decommissioning"
            Recommendation = "Document hybrid config. Plan for hybrid removal or reconfiguration to target."
        }

        # Additional Operations Rules
        @{
            Id          = "OP-004"
            Category    = "Operations"
            Name        = "Shared Channels with External Participants"
            Condition   = { param($data)
                $data.Teams.Teams.Analysis.SharedChannelsWithExternal -gt 0
            }
            Severity    = "Critical"
            Description = "Shared channels with external organizations require B2B direct connect"
            Recommendation = "Document all shared channel relationships. Coordinate with external organizations."
        }
        @{
            Id          = "OP-005"
            Category    = "Operations"
            Name        = "Teams Templates in Use"
            Condition   = { param($data)
                $data.Teams.Templates.Analysis.CustomTemplates -gt 0
            }
            Severity    = "Medium"
            Description = "Custom Teams templates must be recreated in target tenant"
            Recommendation = "Export template definitions. Recreate templates before team provisioning."
        }
        @{
            Id          = "OP-006"
            Category    = "Operations"
            Name        = "Viva Insights Configured"
            Condition   = { param($data)
                $data.Teams.VivaInsights.Analysis.Enabled
            }
            Severity    = "Medium"
            Description = "Viva Insights historical data cannot be migrated"
            Recommendation = "Document Insights configuration. Historical analytics will reset in target."
        }
        @{
            Id          = "OP-007"
            Category    = "Operations"
            Name        = "Planner Plans Active"
            Condition   = { param($data)
                $data.Teams.Planner.Analysis.TotalPlans -gt 100
            }
            Severity    = "High"
            Description = "Large number of Planner plans require migration with associated groups"
            Recommendation = "Inventory Planner usage. Plans migrate with M365 Groups - verify group migration."
        }
        @{
            Id          = "OP-008"
            Category    = "Operations"
            Name        = "Bookings Configured"
            Condition   = { param($data)
                $data.Exchange.Bookings.Analysis.BookingsMailboxes -gt 0
            }
            Severity    = "Medium"
            Description = "Bookings calendars and configurations need recreation"
            Recommendation = "Document Bookings pages. Plan for manual recreation in target tenant."
        }
        @{
            Id          = "OP-009"
            Category    = "Operations"
            Name        = "Stream Classic Videos"
            Condition   = { param($data)
                $data.SharePoint.Stream.Analysis.ClassicVideos -gt 0
            }
            Severity    = "High"
            Description = "Stream Classic videos must be migrated to Stream on SharePoint"
            Recommendation = "Plan Stream Classic to SharePoint migration before tenant migration."
        }
        @{
            Id          = "OP-010"
            Category    = "Operations"
            Name        = "Forms with External Sharing"
            Condition   = { param($data)
                $data.Teams.Forms.Analysis.ExternallySharedForms -gt 0
            }
            Severity    = "Medium"
            Description = "Forms shared externally will need new sharing links"
            Recommendation = "Document externally shared forms. Plan to redistribute new links post-migration."
        }

        # ============================================
        # ADDITIONAL GOTCHAS - DNS & Mail Flow
        # ============================================
        @{
            Id          = "MF-001"
            Category    = "Infrastructure"
            Name        = "DNS Cutover Planning (MX/SPF/DKIM/DMARC)"
            Condition   = { param($data) $data.Exchange.Mailboxes.Analysis.TotalMailboxes -gt 0 }
            Severity    = "Critical"
            Description = "DNS records (MX, SPF, DKIM, DMARC) must be carefully cutover to avoid mail flow disruption and spam filtering issues"
            Recommendation = "Plan DNS cutover sequence with reduced TTLs. Have rollback DNS values ready."
            RemediationSteps = @(
                "1. DOCUMENT CURRENT DNS: Export all MX, SPF, TXT (DKIM/DMARC), and autodiscover records"
                "2. REDUCE TTL: 48 hours before cutover, reduce TTL on all mail-related DNS records to 300 seconds"
                "3. GENERATE TARGET DKIM: In target tenant, enable DKIM signing and note the CNAME records needed"
                "4. PREPARE NEW SPF: Draft new SPF record including target tenant: 'v=spf1 include:spf.protection.outlook.com -all'"
                "5. CUTOVER SEQUENCE: a) Add DKIM CNAMEs, b) Update SPF, c) Switch MX records, d) Update DMARC"
                "6. MONITOR: Watch Message Trace in both tenants for 24-48 hours post-cutover"
                "7. ROLLBACK PLAN: Keep old MX values documented - can revert within TTL window if issues arise"
            )
            Tools = @("DNS Management Console", "MXToolbox", "Exchange Admin Center Message Trace")
            EstimatedEffort = "1 day planning + 2-4 hour cutover window"
            Prerequisites = @("Target tenant fully configured", "Mailboxes migrated", "Test users validated")
        }
        @{
            Id          = "MF-002"
            Category    = "Infrastructure"
            Name        = "Mail Flow Coexistence Free/Busy"
            Condition   = { param($data) $data.Exchange.Mailboxes.Analysis.TotalMailboxes -gt 0 }
            Severity    = "High"
            Description = "During coexistence period, Free/Busy lookups between tenants will fail without Organization Relationships"
            Recommendation = "Configure cross-tenant Organization Relationships and OAuth for Free/Busy sharing"
            RemediationSteps = @(
                "1. ENABLE ORG RELATIONSHIP IN SOURCE: New-OrganizationRelationship -Name 'Target Tenant' -DomainNames 'target.onmicrosoft.com' -FreeBusyAccessEnabled $true -FreeBusyAccessLevel AvailabilityOnly"
                "2. ENABLE ORG RELATIONSHIP IN TARGET: Mirror the configuration pointing to source tenant"
                "3. CONFIGURE OAUTH: Set up OAuth authentication between tenants for enhanced availability"
                "4. TEST: Use Outlook to check free/busy for mailboxes in opposite tenant"
                "5. DOCUMENT COEXISTENCE PERIOD: Define how long coexistence will last and communicate to users"
                "6. POST-MIGRATION: Remove Organization Relationships after all mailboxes migrated"
            )
            Tools = @("Exchange Online PowerShell", "Organization Relationships", "Test-OAuthConnectivity")
            EstimatedEffort = "4-8 hours"
            Prerequisites = @("Both tenants accessible", "Admin credentials for both")
        }
        @{
            Id          = "MF-003"
            Category    = "Operations"
            Name        = "Mailbox Migration Throttling"
            Condition   = { param($data) $data.Exchange.Mailboxes.Analysis.TotalMailboxes -gt 100 }
            Severity    = "High"
            Description = "Microsoft throttles cross-tenant mailbox migrations. Large migrations require batching and extended timelines."
            Recommendation = "Plan batched migration waves with 100-200 mailboxes per batch. Schedule during off-peak hours."
            RemediationSteps = @(
                "1. BATCH PLANNING: Group users by department/location into batches of 100-200"
                "2. CALCULATE TIMELINE: Expect 1-2GB per mailbox per hour average throughput"
                "3. SCHEDULE OFF-PEAK: Start batches Friday evening, validate Monday morning"
                "4. MONITOR PROGRESS: Use Get-MoveRequest | Get-MoveRequestStatistics to track"
                "5. HANDLE FAILURES: Large items (>150MB) may fail - use BadItemLimit parameter"
                "6. INCREMENTAL SYNC: Use SuspendWhenReadyToComplete for delta sync before cutover"
                "7. THROTTLING RESPONSE: If throttled, reduce concurrent migrations and retry"
            )
            Tools = @("Exchange Online PowerShell", "Migration batches", "BitTitan MigrationWiz (optional)")
            EstimatedEffort = "Varies by mailbox count - plan 1 week per 500 mailboxes"
            Prerequisites = @("Migration endpoint configured", "Batch schedules approved", "User communication plan")
        }

        # ============================================
        # ADDITIONAL GOTCHAS - Exchange Permissions
        # ============================================
        @{
            Id          = "EX-001"
            Category    = "Data"
            Name        = "Shared Mailbox Permissions Loss"
            Condition   = { param($data)
                $data.Exchange.Mailboxes.Mailboxes | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" } | Measure-Object | Select-Object -ExpandProperty Count
            }
            Severity    = "Critical"
            Description = "Shared mailbox Full Access, Send-As, and Send-on-Behalf permissions do not migrate automatically"
            Recommendation = "Export all shared mailbox permissions and re-apply in target tenant post-migration"
            RemediationSteps = @(
                "1. EXPORT FULL ACCESS: Get-Mailbox -RecipientTypeDetails SharedMailbox | Get-MailboxPermission | Where {`$_.User -ne 'NT AUTHORITY\\SELF'} | Export-Csv SharedMbxFullAccess.csv"
                "2. EXPORT SEND-AS: Get-Mailbox -RecipientTypeDetails SharedMailbox | Get-RecipientPermission | Export-Csv SharedMbxSendAs.csv"
                "3. EXPORT SEND-ON-BEHALF: Get-Mailbox -RecipientTypeDetails SharedMailbox | Select Identity,GrantSendOnBehalfTo | Export-Csv SharedMbxSendOnBehalf.csv"
                "4. MIGRATE MAILBOXES: Migrate shared mailboxes with user batches"
                "5. RE-APPLY FULL ACCESS: Add-MailboxPermission -Identity <mailbox> -User <user> -AccessRights FullAccess -InheritanceType All"
                "6. RE-APPLY SEND-AS: Add-RecipientPermission -Identity <mailbox> -Trustee <user> -AccessRights SendAs"
                "7. RE-APPLY SEND-ON-BEHALF: Set-Mailbox -Identity <mailbox> -GrantSendOnBehalfTo <user>"
                "8. VALIDATE: Test from Outlook that users can access and send from shared mailboxes"
            )
            Tools = @("Exchange Online PowerShell", "Permission export scripts")
            EstimatedEffort = "4-8 hours depending on shared mailbox count"
            Prerequisites = @("Permission exports complete before migration", "Users created in target")
        }
        @{
            Id          = "EX-002"
            Category    = "Data"
            Name        = "Delegate/Calendar Permissions Loss"
            Condition   = { param($data) $data.Exchange.Mailboxes.Analysis.TotalMailboxes -gt 0 }
            Severity    = "High"
            Description = "Calendar delegate permissions, folder-level permissions, and custom folder shares do not migrate"
            Recommendation = "Document and export calendar/folder permissions. Users may need to re-grant access post-migration."
            RemediationSteps = @(
                "1. EXPORT CALENDAR DELEGATES: Get-Mailbox -ResultSize Unlimited | ForEach { Get-CalendarProcessing -Identity `$_.Identity | Select Identity,ResourceDelegates }"
                "2. EXPORT FOLDER PERMISSIONS: Get-Mailbox | ForEach { Get-MailboxFolderPermission -Identity `"`$(`$_.Identity):\\Calendar`" | Export-Csv CalendarPerms.csv -Append }"
                "3. IDENTIFY EXECUTIVE ASSISTANTS: Focus on executives with delegates - these are business critical"
                "4. COMMUNICATE: Warn users that they may need to re-share calendars after migration"
                "5. RE-APPLY DELEGATES: Set-CalendarProcessing -Identity <mailbox> -ResourceDelegates <users>"
                "6. RE-APPLY FOLDER PERMS: Add-MailboxFolderPermission -Identity 'user:\\Calendar' -User <delegate> -AccessRights Editor"
                "7. VALIDATE: Have delegates confirm calendar access works correctly"
            )
            Tools = @("Exchange Online PowerShell", "Outlook delegate settings")
            EstimatedEffort = "2-4 hours + user validation time"
            Prerequisites = @("Permission exports complete", "Executive assistant list identified")
        }
        @{
            Id          = "EX-003"
            Category    = "Data"
            Name        = "Distribution Group Migration"
            Condition   = { param($data)
                $data.Exchange.DistributionGroups.Groups.Count -gt 0
            }
            Severity    = "High"
            Description = "Distribution groups must be recreated in target. Membership, ownership, and mail properties need restoration."
            Recommendation = "Export DL configurations and recreate in target before mailbox migration"
            RemediationSteps = @(
                "1. EXPORT DL LIST: Get-DistributionGroup -ResultSize Unlimited | Export-Csv DistributionGroups.csv"
                "2. EXPORT MEMBERSHIP: Get-DistributionGroup | ForEach { Get-DistributionGroupMember -Identity `$_.Identity | Select @{N='Group';E={`$_.Identity}},* } | Export-Csv DLMembers.csv"
                "3. EXPORT SETTINGS: Get-DistributionGroup | Select Name,PrimarySmtpAddress,ManagedBy,MemberJoinRestriction,MemberDepartRestriction,RequireSenderAuthenticationEnabled | Export-Csv DLSettings.csv"
                "4. CREATE IN TARGET: New-DistributionGroup -Name <name> -PrimarySmtpAddress <email> -ManagedBy <owner>"
                "5. ADD MEMBERS: Add-DistributionGroupMember -Identity <group> -Member <user>"
                "6. APPLY SETTINGS: Set-DistributionGroup with appropriate restrictions and settings"
                "7. UPDATE SMTP: Ensure proxyAddresses include all aliases from source"
                "8. VALIDATE: Send test email to DL and verify all members receive"
            )
            Tools = @("Exchange Online PowerShell", "CSV processing scripts")
            EstimatedEffort = "4-8 hours depending on DL count"
            Prerequisites = @("All DL members must exist in target tenant first")
        }
        @{
            Id          = "EX-004"
            Category    = "Data"
            Name        = "Calendar Metadata and Meeting Links"
            Condition   = { param($data) $data.Exchange.Mailboxes.Analysis.TotalMailboxes -gt 0 }
            Severity    = "Medium"
            Description = "Teams meeting links in calendar items will point to source tenant. Recurring meeting organizers may need to resend invites."
            Recommendation = "Document recurring meetings with external attendees. Plan Teams meeting link updates."
            RemediationSteps = @(
                "1. IDENTIFY RECURRING MEETINGS: Export calendar items with recurrence patterns, especially those with external attendees"
                "2. TEAMS MEETING LINKS: Understand that existing Teams links will still work briefly but should be regenerated"
                "3. COMMUNICATE TO ORGANIZERS: Meeting organizers should send updates to recurring series post-migration"
                "4. RESOURCE ROOMS: Room calendars with recurring bookings may need special handling"
                "5. AUTOMATED SCRIPT: Consider script to identify meetings extending past migration date"
                "6. EXTERNAL ATTENDEES: External attendees will receive updates with new Teams links"
                "7. POST-MIGRATION: Monitor for meeting join failures and proactively fix critical meetings"
            )
            Tools = @("Exchange Online PowerShell", "Graph API for calendar queries")
            EstimatedEffort = "2-4 hours planning + organizer actions"
            Prerequisites = @("Meeting organizer communication plan")
        }

        # ============================================
        # ADDITIONAL GOTCHAS - Teams Specific
        # ============================================
        @{
            Id          = "TM-001"
            Category    = "Data"
            Name        = "Teams Chat History Not Migrated"
            Condition   = { param($data) $data.Teams.Teams.Analysis.TotalTeams -gt 0 }
            Severity    = "Critical"
            Description = "1:1 and group chat history does NOT migrate in cross-tenant migrations. Only channel messages migrate with Teams."
            Recommendation = "Export chat history before migration. Set user expectations clearly."
            RemediationSteps = @(
                "1. SET EXPECTATIONS: Clearly communicate to users that private chats will NOT transfer"
                "2. USER EXPORT: Guide users to export important chats via Teams > Settings > Export or Content Search"
                "3. EDISCOVERY EXPORT: For compliance, export all chats via Content Search before decommissioning source"
                "4. ARCHIVE CHATS: Users should screenshot or save critical conversation content"
                "5. THIRD-PARTY TOOLS: Consider tools like AvePoint or BitTitan for chat backup if budget allows"
                "6. COMPLIANCE HOLD: If regulatory requirement, ensure chat data is preserved in source tenant archive"
                "7. DOCUMENT DECISION: Get sign-off that business accepts chat history will not migrate"
            )
            Tools = @("Teams Export", "eDiscovery Content Search", "Third-party migration tools")
            EstimatedEffort = "2-4 hours for export + user communication"
            Prerequisites = @("User communication completed", "Compliance sign-off obtained")
        }
        @{
            Id          = "TM-002"
            Category    = "Data"
            Name        = "Teams Meeting Recordings Location"
            Condition   = { param($data) $data.Teams.Teams.Analysis.TotalTeams -gt 0 }
            Severity    = "High"
            Description = "Teams meeting recordings stored in OneDrive/SharePoint will be in source tenant. New recordings go to target."
            Recommendation = "Document recording locations. Plan to migrate or archive important recordings."
            RemediationSteps = @(
                "1. LOCATE RECORDINGS: Recordings are in organizer's OneDrive > Recordings folder or SharePoint > Recordings"
                "2. IDENTIFY CRITICAL RECORDINGS: Work with departments to flag recordings that must be preserved"
                "3. MIGRATE RECORDINGS: Include OneDrive Recordings folder in SharePoint migration scope"
                "4. UPDATE LINKS: Meeting chat links to recordings will break - need to reshare in target"
                "5. RETENTION: Apply retention policy to recording locations to prevent accidental deletion"
                "6. COMMUNICATE: Tell users old recordings accessible via source until decommissioned"
                "7. STREAM CLASSIC: If using Stream Classic, separate migration needed (see Stream gotcha)"
            )
            Tools = @("SharePoint Migration Tool", "OneDrive for Business")
            EstimatedEffort = "Included in OneDrive migration"
            Prerequisites = @("Recording inventory complete", "Critical recordings flagged")
        }
        @{
            Id          = "TM-003"
            Category    = "Integration"
            Name        = "Teams Channel Tabs and Connectors"
            Condition   = { param($data) $data.Teams.Teams.Analysis.TotalTeams -gt 0 }
            Severity    = "High"
            Description = "Channel tabs (website, document, Power BI, Planner, etc.) do not migrate. Connectors and bots need reconfiguration."
            Recommendation = "Document all channel tabs and connectors. Plan manual reconfiguration in target."
            RemediationSteps = @(
                "1. INVENTORY TABS: For each Team, document all channel tabs including type and target URL/document"
                "2. INVENTORY CONNECTORS: Get-Team | ForEach { Get-TeamsChannelConnector (if available) } - or manual review"
                "3. CATEGORIZE TABS: Website tabs - just re-add URL; Document tabs - need new document link; App tabs - reinstall app"
                "4. PLANNER TABS: Planner plans must be migrated separately, then tab re-added"
                "5. POWER BI TABS: Re-add tabs pointing to migrated Power BI reports"
                "6. RECREATE CONNECTORS: Webhooks and connectors must be recreated - get URLs from source first"
                "7. BOTS: Custom bots need app registration update to point to target tenant"
                "8. VALIDATE: After migration, verify each critical channel has its tabs restored"
            )
            Tools = @("Teams Admin Center", "Microsoft Graph API for Teams")
            EstimatedEffort = "2-8 hours depending on tab complexity"
            Prerequisites = @("Tab and connector inventory complete")
        }
        @{
            Id          = "TM-004"
            Category    = "Data"
            Name        = "Teams Wiki Content"
            Condition   = { param($data) $data.Teams.Teams.Analysis.TotalTeams -gt 0 }
            Severity    = "Medium"
            Description = "Teams Wiki content is being deprecated but existing wikis need migration consideration"
            Recommendation = "Export Wiki content before migration. Consider migrating to OneNote or SharePoint pages."
            RemediationSteps = @(
                "1. IDENTIFY WIKIS: Review each Team channel for Wiki tabs with content"
                "2. EXPORT CONTENT: Wiki content can be exported via SharePoint - located in Teams Wiki Data library"
                "3. CONVERT FORMAT: Consider converting important wikis to OneNote notebooks or SharePoint pages"
                "4. MIGRATE DATA: Include Wiki Data library in SharePoint migration scope"
                "5. UPDATE TABS: After migration, add new Wiki tab or OneNote tab pointing to migrated content"
                "6. COMMUNICATE: Inform users about Wiki deprecation and new location"
            )
            Tools = @("SharePoint Migration Tool", "OneNote")
            EstimatedEffort = "1-2 hours per Team with significant Wiki content"
            Prerequisites = @("Wiki content inventory")
        }

        # ============================================
        # ADDITIONAL GOTCHAS - SharePoint/OneDrive
        # ============================================
        @{
            Id          = "SP-001"
            Category    = "Data"
            Name        = "SharePoint Permissions Inheritance Break"
            Condition   = { param($data) $data.SharePoint.Sites.Analysis.SharePointSites -gt 0 }
            Severity    = "Critical"
            Description = "Unique permissions on subsites, libraries, folders, and items may not migrate correctly. Inheritance can reset."
            Recommendation = "Document unique permissions before migration. Validate and restore post-migration."
            RemediationSteps = @(
                "1. IDENTIFY BROKEN INHERITANCE: Use PnP PowerShell: Get-PnPList | Where {`$_.HasUniqueRoleAssignments}"
                "2. EXPORT PERMISSIONS: Get-PnPSiteCollectionAdmin, Get-PnPGroup, Get-PnPGroupMember for each site"
                "3. DOCUMENT ITEM-LEVEL: For critical libraries, export item-level permissions"
                "4. CHOOSE MIGRATION TOOL: SharePoint Migration Tool preserves most permissions; verify tool capabilities"
                "5. TEST MIGRATION: Pilot with permission-heavy site first"
                "6. VALIDATE POST-MIGRATION: Spot-check unique permission locations"
                "7. RESTORE AS NEEDED: Re-apply permissions using Set-PnPListItemPermission or Set-PnPFolderPermission"
                "8. AUDIT: Compare permission reports before/after migration"
            )
            Tools = @("PnP PowerShell", "SharePoint Migration Tool", "Permission reports")
            EstimatedEffort = "4-16 hours depending on complexity"
            Prerequisites = @("Permission audit complete", "Critical sites identified")
        }
        @{
            Id          = "SP-002"
            Category    = "Data"
            Name        = "SharePoint Metadata and Version History"
            Condition   = { param($data) $data.SharePoint.Sites.Analysis.SharePointSites -gt 0 }
            Severity    = "High"
            Description = "Custom metadata columns, content types, and version history may not fully migrate depending on tool used"
            Recommendation = "Validate migration tool supports metadata preservation. Document version history requirements."
            RemediationSteps = @(
                "1. INVENTORY METADATA: Document custom site columns, content types, and managed metadata term sets"
                "2. VERSION REQUIREMENTS: Determine business requirement for version history (all, recent X, none)"
                "3. MIGRATION TOOL SETTINGS: Configure tool to preserve Created/Modified dates and version history"
                "4. TERM STORE: Managed Metadata term store must be migrated first or recreated in target"
                "5. CONTENT TYPES: Site content types need recreation; consider Content Type Hub"
                "6. TEST MIGRATION: Validate metadata intact on test documents"
                "7. LOOKUP COLUMNS: Lookup columns may break if referenced list not yet migrated"
                "8. POST-VALIDATION: Compare item properties before/after for sample documents"
            )
            Tools = @("SharePoint Migration Tool", "PnP PowerShell", "Sharegate (optional)")
            EstimatedEffort = "4-8 hours for inventory + migration validation"
            Prerequisites = @("Metadata inventory complete", "Term store documented")
        }
        @{
            Id          = "SP-003"
            Category    = "Data"
            Name        = "OneDrive Shortcuts and Synced Folders"
            Condition   = { param($data) $data.SharePoint.OneDrive.Sites.Count -gt 0 }
            Severity    = "High"
            Description = "OneDrive shortcuts to SharePoint ('Add shortcut to OneDrive') will break. Locally synced folders need resync."
            Recommendation = "Document OneDrive shortcuts. Plan sync client reconfiguration post-migration."
            RemediationSteps = @(
                "1. INVENTORY SHORTCUTS: Shortcuts are stored in OneDrive but point to SharePoint - these WILL break"
                "2. COMMUNICATE: Warn users their OneDrive shortcuts will need to be recreated"
                "3. SYNC CLIENT: OneDrive sync client will need to be signed out and resigned into target tenant"
                "4. UNLINK FIRST: Before migration, users should unlink OneDrive sync (Settings > Account > Unlink)"
                "5. CLEAN CACHE: Delete local OneDrive cache folder to avoid conflicts"
                "6. RELINK POST-MIGRATION: Sign into OneDrive with target tenant credentials"
                "7. RECREATE SHORTCUTS: Users manually re-add shortcuts to frequently accessed SharePoint libraries"
                "8. FILES ON-DEMAND: Ensure Files On-Demand is enabled to avoid re-downloading all content"
            )
            Tools = @("OneDrive sync client", "User communication")
            EstimatedEffort = "15-30 minutes per user"
            Prerequisites = @("User communication sent", "Shortcut documentation complete")
        }
        @{
            Id          = "SP-004"
            Category    = "Data"
            Name        = "External Sharing Links Invalidation"
            Condition   = { param($data) $data.SharePoint.Sites.Analysis.ExternalSharingEnabled -gt 0 }
            Severity    = "High"
            Description = "All external sharing links (Anyone links, specific people links) will be invalidated after migration"
            Recommendation = "Document critical external shares. Plan to regenerate and redistribute links."
            RemediationSteps = @(
                "1. INVENTORY EXTERNAL SHARES: Get-PnPSharingLinks or Sharing Reports in SharePoint Admin"
                "2. IDENTIFY CRITICAL: Work with departments to identify business-critical external shares"
                "3. COMMUNICATE EXTERNALLY: Warn external partners that links will stop working"
                "4. MIGRATE CONTENT: External sharing links do not migrate - only content does"
                "5. REGENERATE LINKS: After migration, create new sharing links for critical files/folders"
                "6. REDISTRIBUTE: Send new links to external partners"
                "7. GUEST ACCESS: External users with guest accounts need reinvitation (see guest gotcha)"
                "8. AUDIT: Review external sharing settings in target match source configuration"
            )
            Tools = @("SharePoint Admin Center", "PnP PowerShell sharing reports")
            EstimatedEffort = "2-4 hours + external communication"
            Prerequisites = @("External share inventory", "Partner communication plan")
        }

        # ============================================
        # ADDITIONAL GOTCHAS - Devices & Intune
        # ============================================
        @{
            Id          = "DV-001"
            Category    = "Infrastructure"
            Name        = "Intune Device Re-enrollment Required"
            Condition   = { param($data) $data.EntraID.Devices.Analysis.IntuneManaged -gt 0 }
            Severity    = "Critical"
            Description = "Intune-managed devices must be unenrolled from source and re-enrolled in target. MDM policies will not transfer."
            Recommendation = "Plan phased device re-enrollment. Export and recreate all Intune policies in target."
            RemediationSteps = @(
                "1. EXPORT INTUNE POLICIES: Use Intune backup tools or Graph API to export all policies"
                "2. RECREATE IN TARGET: Manually recreate or use IntuneBackupAndRestore to restore policies"
                "3. TEST POLICIES: Validate policies on test devices before mass enrollment"
                "4. UNENROLL DEVICES: Remove devices from source Intune - Company Portal > Settings > Remove device"
                "5. AZURE AD UNJOIN: For AAD joined devices, run dsregcmd /leave"
                "6. RE-ENROLL: Join to target AAD and enroll in target Intune"
                "7. VALIDATE COMPLIANCE: Check device shows compliant in target Intune"
                "8. CONDITIONAL ACCESS: Ensure target CA policies recognize newly enrolled devices"
            )
            Tools = @("Intune", "IntuneBackupAndRestore", "dsregcmd", "Company Portal")
            EstimatedEffort = "15-30 minutes per device (phased rollout)"
            Prerequisites = @("Intune policies recreated in target", "User communication complete")
        }
        @{
            Id          = "DV-002"
            Category    = "Infrastructure"
            Name        = "AutoPilot Profile Migration"
            Condition   = { param($data) $data.EntraID.Devices.Analysis.AutoPilotRegistered -gt 0 }
            Severity    = "Critical"
            Description = "Windows AutoPilot device registrations are tenant-specific. Devices must be deregistered and re-registered."
            Recommendation = "Export AutoPilot device list. Plan hardware hash re-import to target tenant."
            RemediationSteps = @(
                "1. EXPORT DEVICE LIST: In Intune > Devices > Windows enrollment > Devices, export all AutoPilot devices"
                "2. EXPORT HARDWARE HASHES: Download CSV with hardware hashes from source tenant"
                "3. DEREGISTER FROM SOURCE: Delete devices from source tenant AutoPilot"
                "4. WAIT FOR SYNC: Microsoft backend may take up to 24 hours to release device"
                "5. IMPORT TO TARGET: Upload hardware hash CSV to target tenant AutoPilot"
                "6. ASSIGN PROFILES: Create/assign deployment profiles in target tenant"
                "7. DEVICE RESET: Devices need factory reset to pick up new AutoPilot profile"
                "8. VALIDATE: Test OOBE experience on sample device with target tenant"
            )
            Tools = @("Intune AutoPilot", "Hardware Hash CSV", "Windows Recovery")
            EstimatedEffort = "1-2 hours for export/import + reset time per device"
            Prerequisites = @("AutoPilot profiles created in target", "Device reset schedule approved")
        }
        @{
            Id          = "DV-003"
            Category    = "Operations"
            Name        = "Client Profile and Cached Credentials"
            Condition   = { param($data) $data.EntraID.Users.Analysis.LicensedUsers -gt 0 }
            Severity    = "High"
            Description = "Users will need to sign out of all Office apps and re-authenticate. Cached credentials will cause errors."
            Recommendation = "Prepare user guide for clearing credentials. Plan for increased helpdesk calls."
            RemediationSteps = @(
                "1. DOCUMENT STEPS: Create user guide for signing out and clearing cached credentials"
                "2. OUTLOOK: File > Office Account > Sign Out. Close Outlook. Delete %localappdata%\\Microsoft\\Outlook\\*.ost"
                "3. TEAMS: Sign out from Teams. Clear Teams cache: %appdata%\\Microsoft\\Teams"
                "4. OFFICE APPS: Sign out from File > Account in each app. May need to clear credential manager"
                "5. CREDENTIAL MANAGER: Control Panel > Credential Manager > Windows Credentials > Remove Office entries"
                "6. BROWSER: Clear browser cache and cookies for Microsoft sites"
                "7. ONEDRIVE: Unlink and relink OneDrive sync client"
                "8. HELPDESK PREP: Brief helpdesk on common authentication issues and resolutions"
            )
            Tools = @("User documentation", "Credential Manager", "Clear cache scripts")
            EstimatedEffort = "15-30 minutes per user"
            Prerequisites = @("User documentation prepared", "Helpdesk briefed")
        }
        @{
            Id          = "DV-004"
            Category    = "Identity"
            Name        = "Multi-Factor Authentication Disruption"
            Condition   = { param($data) $data.EntraID.Users.Analysis.LicensedUsers -gt 0 }
            Severity    = "High"
            Description = "MFA registrations are tenant-specific. Users must re-register MFA methods in target tenant."
            Recommendation = "Plan MFA re-enrollment. Consider temporary MFA bypass during cutover window."
            RemediationSteps = @(
                "1. INVENTORY MFA: Get-MgUserAuthenticationMethod to document current MFA methods per user"
                "2. CONFIGURE TARGET: Enable MFA in target tenant with same methods as source"
                "3. TEMPORARY ACCESS PASS: Consider using TAP for initial sign-in to allow MFA registration"
                "4. COMMUNICATE: Tell users they will need to set up MFA again (Authenticator app, phone)"
                "5. RE-REGISTRATION: Users go to aka.ms/mfasetup in target tenant to register"
                "6. FIDO2/WINDOWS HELLO: Hardware keys and biometric methods need re-registration"
                "7. CONDITIONAL ACCESS: Ensure CA policies allow MFA registration grace period"
                "8. HELPDESK: Prepare for increased MFA-related support calls"
            )
            Tools = @("Entra ID Authentication Methods", "Temporary Access Pass", "aka.ms/mfasetup")
            EstimatedEffort = "5-10 minutes per user"
            Prerequisites = @("MFA policies configured in target", "User communication sent")
        }

        # ============================================
        # ADDITIONAL GOTCHAS - Applications & API
        # ============================================
        @{
            Id          = "AP-001"
            Category    = "Integration"
            Name        = "App Registrations OAuth Token Reset"
            Condition   = { param($data) $data.EntraID.Applications.Applications.Count -gt 0 }
            Severity    = "Critical"
            Description = "App registrations are tenant-specific. All OAuth tokens will be invalid. Apps need re-registration in target."
            Recommendation = "Document all app registrations. Plan recreation and secret rotation in target."
            RemediationSteps = @(
                "1. INVENTORY APPS: Get-MgApplication -All | Export-Csv AppRegistrations.csv"
                "2. DOCUMENT CONFIG: For each app, note: Redirect URIs, API permissions, secrets/certificates, owners"
                "3. EXPORT MANIFESTS: Download app manifests from Azure portal for reference"
                "4. CREATE IN TARGET: New-MgApplication with same configuration"
                "5. GENERATE NEW SECRETS: Create new client secrets/certificates in target"
                "6. UPDATE APPLICATIONS: All applications using these registrations need configuration updates"
                "7. ADMIN CONSENT: Grant admin consent for API permissions in target tenant"
                "8. TEST: Validate each application authenticates successfully against target"
            )
            Tools = @("Microsoft Graph PowerShell", "Azure Portal App Registrations")
            EstimatedEffort = "1-2 hours per complex app, 15-30 min for simple apps"
            Prerequisites = @("App inventory complete", "App owner contacts available")
        }
        @{
            Id          = "AP-002"
            Category    = "Integration"
            Name        = "Third-Party API Integrations Failure"
            Condition   = { param($data) $data.EntraID.Applications.Applications.Count -gt 0 }
            Severity    = "High"
            Description = "Third-party applications integrated via OAuth/SAML will stop working until reconfigured for target tenant"
            Recommendation = "Inventory all third-party integrations. Coordinate with vendors for reconfiguration."
            RemediationSteps = @(
                "1. INVENTORY INTEGRATIONS: List all third-party apps with M365 integration (HRIS, CRM, etc.)"
                "2. IDENTIFY AUTH TYPE: Document whether each uses OAuth, SAML SSO, or API keys"
                "3. CONTACT VENDORS: Engage vendor support early for enterprise apps"
                "4. SAML APPS: Update SAML metadata in third-party with target tenant identifiers"
                "5. OAUTH APPS: Update app registrations and provide new client credentials"
                "6. ENTERPRISE APPS: Recreate Enterprise App configurations in target Entra ID"
                "7. TEST INTEGRATIONS: Validate each integration before cutover"
                "8. ROLLBACK PLAN: Have vendor contacts ready for urgent issues during cutover"
            )
            Tools = @("Entra ID Enterprise Applications", "Vendor documentation")
            EstimatedEffort = "Varies significantly by app complexity"
            Prerequisites = @("Vendor contacts identified", "Integration inventory complete")
        }
        @{
            Id          = "AP-003"
            Category    = "Integration"
            Name        = "Power BI Workspace Ownership"
            Condition   = { param($data) $data.PowerPlatform.PowerBI.Workspaces.Count -gt 0 }
            Severity    = "High"
            Description = "Power BI workspaces, reports, and datasets need migration. Data connections and gateway configs need recreation."
            Recommendation = "Export Power BI content. Plan workspace recreation with proper ownership."
            RemediationSteps = @(
                "1. INVENTORY WORKSPACES: Document all workspaces, reports, dashboards, and datasets"
                "2. IDENTIFY OWNERS: Record workspace owners and members for permission recreation"
                "3. EXPORT PBIX: Download .pbix files for reports that can be downloaded"
                "4. DOCUMENT DATA SOURCES: Record all data source connections and credentials"
                "5. GATEWAY CONFIG: If using on-premises gateway, plan gateway installation in target"
                "6. CREATE WORKSPACES: Recreate workspace structure in target tenant"
                "7. PUBLISH REPORTS: Upload .pbix files and republish to target workspaces"
                "8. RECONFIGURE DATA: Update data source credentials and refresh schedules"
            )
            Tools = @("Power BI Admin Portal", "Power BI Desktop", "On-premises Gateway")
            EstimatedEffort = "4-16 hours depending on workspace count"
            Prerequisites = @("Power BI inventory complete", "Gateway decisions made")
        }

        # ============================================
        # ADDITIONAL GOTCHAS - Identity & Licensing
        # ============================================
        @{
            Id          = "LI-001"
            Category    = "Operations"
            Name        = "UPN vs Primary SMTP Address Mismatch"
            Condition   = { param($data)
                $users = $data.EntraID.Users.Users
                $mismatched = $users | Where-Object { $_.UserPrincipalName -and $_.Mail -and $_.UserPrincipalName.Split('@')[0] -ne $_.Mail.Split('@')[0] }
                ($mismatched | Measure-Object).Count -gt 0
            }
            Severity    = "High"
            Description = "Users with different UPN and primary email address require special attention during identity matching"
            Recommendation = "Document UPN vs email mismatches. Decide on target naming convention."
            RemediationSteps = @(
                "1. IDENTIFY MISMATCHES: Get-MgUser -All | Where {`$_.UserPrincipalName.Split('@')[0] -ne `$_.Mail.Split('@')[0]} | Export-Csv UPNMismatch.csv"
                "2. ANALYZE PATTERNS: Determine if mismatches are intentional (maiden names, preferred names) or historical"
                "3. DECIDE CONVENTION: Choose whether target will use UPN or email as primary identifier"
                "4. PLAN MAPPING: Create mapping table for identity matching in target"
                "5. AAD CONNECT: If using soft match, ensure proxyAddresses are correctly mapped"
                "6. COMMUNICATE: Inform affected users if their sign-in name will change"
                "7. UPDATE SYSTEMS: Any systems using old UPN format need updates"
                "8. VALIDATE: Test authentication for users with mismatches"
            )
            Tools = @("Microsoft Graph PowerShell", "Excel for mapping")
            EstimatedEffort = "2-4 hours for analysis"
            Prerequisites = @("User export complete", "Naming convention approved")
        }
        @{
            Id          = "LI-002"
            Category    = "Operations"
            Name        = "License Entitlement Verification"
            Condition   = { param($data) $data.EntraID.Users.Analysis.LicensedUsers -gt 0 }
            Severity    = "High"
            Description = "License counts and SKUs in target must match source requirements. Direct vs group-based licensing needs planning."
            Recommendation = "Audit license usage. Ensure target has equivalent licenses before migration."
            RemediationSteps = @(
                "1. EXPORT LICENSE USAGE: Get-MgSubscribedSku | Select SkuPartNumber,ConsumedUnits,PrepaidUnits"
                "2. MAP SKUS: Create mapping between source and target license SKUs"
                "3. PROCURE LICENSES: Ensure target tenant has sufficient license capacity"
                "4. GROUP-BASED LICENSING: If using GBL, recreate license groups in target"
                "5. DIRECT ASSIGNMENT: For direct-assigned licenses, plan bulk assignment"
                "6. SERVICE PLANS: Note which service plans are enabled/disabled per license"
                "7. ASSIGN IN TARGET: Assign licenses before or during user migration"
                "8. VALIDATE: Compare licensed user counts between tenants"
            )
            Tools = @("Microsoft Graph PowerShell", "M365 Admin Center", "License CSVs")
            EstimatedEffort = "4-8 hours"
            Prerequisites = @("License inventory complete", "Target licenses procured")
        }
        @{
            Id          = "LI-003"
            Category    = "Operations"
            Name        = "Domain Verification Timing"
            Condition   = { param($data) $true }  # Always check
            Severity    = "Critical"
            Description = "Custom domains cannot exist in two tenants simultaneously. Domain removal/addition timing is critical."
            Recommendation = "Plan domain cutover window carefully. Reduce DNS TTLs beforehand."
            RemediationSteps = @(
                "1. LIST DOMAINS: Get-MgDomain | Select Id,IsDefault,IsVerified"
                "2. REDUCE TTL: 48 hours before cutover, reduce DNS TTL to 300 seconds"
                "3. PREPARE TARGET: Have domain verification TXT record ready for target tenant"
                "4. REMOVE FROM SOURCE: Remove-MgDomain (or via admin center) - this breaks email routing!"
                "5. ADD TO TARGET: New-MgDomain then verify with TXT record"
                "6. UPDATE MX: Immediately update MX records to point to target tenant"
                "7. VERIFY ROUTING: Test mail flow to domain lands in target tenant"
                "8. TIMING: Plan for 15-30 minute window where domain is unverified in either tenant"
            )
            Tools = @("Microsoft Graph PowerShell", "DNS Management", "M365 Admin Center")
            EstimatedEffort = "2-4 hours for cutover window"
            Prerequisites = @("All mailboxes migrated", "DNS access ready", "Off-hours window scheduled")
        }
        @{
            Id          = "LI-004"
            Category    = "Identity"
            Name        = "Security Group vs M365 Group Confusion"
            Condition   = { param($data)
                $data.EntraID.Groups.Analysis.TotalGroups -gt 0
            }
            Severity    = "Medium"
            Description = "Security groups and M365 groups have different migration paths. Mail-enabled security groups add complexity."
            Recommendation = "Classify all groups by type. Plan appropriate migration method for each type."
            RemediationSteps = @(
                "1. CLASSIFY GROUPS: Export groups with their type (Security, M365, Mail-enabled Security, Distribution)"
                "2. M365 GROUPS: These have associated Teams/SharePoint - migrate with those workloads"
                "3. SECURITY GROUPS: Need recreation in target - Get-MgGroup -Filter 'securityEnabled eq true'"
                "4. MAIL-ENABLED: Must decide - convert to M365 group or recreate as DL + security group"
                "5. DYNAMIC GROUPS: Membership rules need recreation - export MembershipRule property"
                "6. NESTED GROUPS: Document group nesting - recreate in correct order (leaf groups first)"
                "7. PERMISSIONS: Groups used in permissions (SharePoint, CA, etc.) must exist before those configs"
                "8. VALIDATE: Compare group counts and membership between tenants"
            )
            Tools = @("Microsoft Graph PowerShell", "Entra ID Admin Center")
            EstimatedEffort = "4-8 hours depending on group count"
            Prerequisites = @("Group inventory and classification complete")
        }

        # ============================================
        # ADDITIONAL GOTCHAS - Compliance & Operations
        # ============================================
        @{
            Id          = "CO-001"
            Category    = "Operations"
            Name        = "Missing Rollback/Contingency Plan"
            Condition   = { param($data) $true }  # Always flag this
            Severity    = "Critical"
            Description = "Without a tested rollback plan, migration failures can cause extended outages"
            Recommendation = "Document detailed rollback procedures for each migration phase. Test on pilot."
            RemediationSteps = @(
                "1. DEFINE ROLLBACK TRIGGERS: What conditions would trigger a rollback decision?"
                "2. MX ROLLBACK: Keep old MX records documented - can revert within DNS TTL window"
                "3. MAILBOX ROLLBACK: For mailboxes, reverse migration direction is possible but costly"
                "4. IDENTITY ROLLBACK: If AAD Connect, can switch sync back to source tenant"
                "5. DOCUMENT CONTACTS: List who makes rollback decision and how to reach them"
                "6. TIME WINDOWS: Define point of no return for each phase"
                "7. DATA BACKUP: Ensure source tenant data is backed up/retained during migration"
                "8. TEST ROLLBACK: During pilot, actually test rolling back at least one user"
            )
            Tools = @("Runbook documentation", "Communication plan")
            EstimatedEffort = "8-16 hours for planning and documentation"
            Prerequisites = @("Rollback plan approved by stakeholders")
        }
        @{
            Id          = "CO-002"
            Category    = "Operations"
            Name        = "Insufficient Pilot Testing"
            Condition   = { param($data) $true }  # Always check
            Severity    = "High"
            Description = "Skipping or rushing pilot testing leads to unexpected issues during full migration"
            Recommendation = "Conduct thorough pilot migration with diverse user sample. Document all issues."
            RemediationSteps = @(
                "1. SELECT PILOT GROUP: Choose 10-20 users representing different roles, departments, and use cases"
                "2. INCLUDE EDGE CASES: Include users with shared mailboxes, delegates, mobile devices, etc."
                "3. EXECUTE FULL PROCESS: Run complete migration process for pilot group"
                "4. USER ACCEPTANCE: Have pilot users test all their workflows and applications"
                "5. DOCUMENT ISSUES: Record all issues encountered and resolutions"
                "6. REFINE PROCESS: Update migration procedures based on pilot learnings"
                "7. PILOT DURATION: Allow 1-2 weeks of pilot operation before full migration"
                "8. GO/NO-GO: Formal decision point based on pilot success before proceeding"
            )
            Tools = @("Pilot tracking spreadsheet", "User feedback forms")
            EstimatedEffort = "1-2 weeks for pilot"
            Prerequisites = @("Pilot user group identified and willing")
        }
        @{
            Id          = "CO-003"
            Category    = "Operations"
            Name        = "Poor User Communication/Change Management"
            Condition   = { param($data) $true }  # Always flag
            Severity    = "High"
            Description = "Inadequate user communication leads to confusion, resistance, and increased support burden"
            Recommendation = "Develop comprehensive communication plan with multiple touchpoints"
            RemediationSteps = @(
                "1. STAKEHOLDER ANALYSIS: Identify all affected groups and their communication needs"
                "2. COMMUNICATION SCHEDULE: Plan messages at -6 weeks, -3 weeks, -1 week, -1 day, day-of, +1 day, +1 week"
                "3. MULTIPLE CHANNELS: Use email, SharePoint news, Teams announcements, town halls"
                "4. SELF-SERVICE CONTENT: Create FAQ, how-to guides, video tutorials"
                "5. MANAGER BRIEFING: Equip managers to answer team questions"
                "6. HELPDESK PREP: Brief support teams on expected issues and resolutions"
                "7. FEEDBACK CHANNEL: Provide way for users to report issues and get help"
                "8. POST-MIGRATION: Follow up with tips and support resources"
            )
            Tools = @("Communication templates", "SharePoint news", "Email")
            EstimatedEffort = "8-16 hours for planning and content creation"
            Prerequisites = @("Communication plan approved", "Content created and reviewed")
        }
        @{
            Id          = "CO-004"
            Category    = "Compliance"
            Name        = "Audit Logs Not Preserved"
            Condition   = { param($data) $true }  # Always check
            Severity    = "High"
            Description = "Unified audit logs in source tenant will not be available after decommissioning. May have compliance implications."
            Recommendation = "Export audit logs before source tenant decommission. Consider third-party archival."
            RemediationSteps = @(
                "1. DETERMINE REQUIREMENTS: Check retention requirements (regulatory, legal, internal)"
                "2. EXPORT AUDIT LOGS: Search-UnifiedAuditLog -StartDate <date> -EndDate <date> | Export-Csv"
                "3. MAILBOX AUDITS: Export mailbox audit logs separately if needed"
                "4. SIGN-IN LOGS: Export Entra ID sign-in logs via Graph API or portal"
                "5. ACTIVITY LOGS: Export Azure activity logs if using Azure resources"
                "6. ARCHIVE STORAGE: Store exports in compliant long-term storage (Azure Blob, etc.)"
                "7. RETENTION PERIOD: Maintain for required retention period post-migration"
                "8. DOCUMENT: Record what was exported, when, and where stored"
            )
            Tools = @("Exchange Online PowerShell", "Microsoft Graph API", "Azure Storage")
            EstimatedEffort = "4-8 hours for export"
            Prerequisites = @("Retention requirements documented", "Archive storage prepared")
        }
        # ============================================
        # POWER PLATFORM & POWER BI RULES
        # ============================================
        @{
            Id          = "PP-001"
            Category    = "Integration"
            Name        = "Power BI Workspaces Without Premium Capacity"
            Condition   = { param($data)
                $data.PowerBI.Workspaces.Analysis.TotalWorkspaces -gt 0 -and
                ($data.PowerBI.Capacities.Analysis.PremiumCapacities -eq 0 -or
                 $null -eq $data.PowerBI.Capacities.Analysis.PremiumCapacities)
            }
            Severity    = "High"
            Description = "Power BI workspaces exist but no Premium capacity is configured — shared capacity workspaces have limitations"
            Recommendation = "Assess whether Premium capacity is needed in target tenant. Plan license procurement before migration."
            RemediationSteps = @(
                "1. AUDIT WORKSPACE USAGE: Review which workspaces need large dataset support or dedicated compute"
                "2. ASSESS PRO LICENSES: Confirm all content creators/viewers have Power BI Pro licenses in target"
                "3. EVALUATE PREMIUM NEED: Premium required for: paginated reports, datasets >1GB, deployment pipelines"
                "4. PROVISION CAPACITY: If needed, purchase Power BI Premium P-SKU or Fabric capacity in target tenant"
                "5. ASSIGN WORKSPACES: After provisioning, assign premium-eligible workspaces to capacity"
                "6. MIGRATE REPORTS: Export .pbix files and upload to target workspace"
                "7. RECONFIGURE DATA SOURCES: Update connection strings, credentials, and refresh schedules"
                "8. VALIDATE REFRESH: Test scheduled refresh for each dataset post-migration"
            )
            Tools = @("Power BI Admin Portal", "Power BI REST API", "MicrosoftPowerBIMgmt module")
            EstimatedEffort = "1-2 days per workspace depending on complexity"
            Prerequisites = @("Power BI Pro licenses confirmed", "Premium capacity decision made")
        }
        @{
            Id          = "PP-002"
            Category    = "Integration"
            Name        = "No Power Platform DLP Policies"
            Condition   = { param($data)
                $data.Dynamics365.Environments.Analysis.TotalEnvironments -gt 0 -and
                ($data.Dynamics365.DLPPolicies.Analysis.TotalDLPPolicies -eq 0 -or
                 $null -eq $data.Dynamics365.DLPPolicies.Analysis.TotalDLPPolicies)
            }
            Severity    = "High"
            Description = "Power Platform environments exist but no DLP policies are configured — data can flow freely between any connectors"
            Recommendation = "Create Power Platform DLP policies in target tenant before enabling users. Define Business, Non-Business, and Blocked connector groups."
            RemediationSteps = @(
                "1. INVENTORY CONNECTORS IN USE: Review connector usage across apps and flows before writing policies"
                "2. DEFINE CONNECTOR GROUPS: Categorize connectors as Business (internal data), Non-Business (external), Blocked"
                "3. CREATE TENANT-WIDE POLICY: Apply a base policy to All Environments in Power Platform Admin Center"
                "4. BLOCK HIGH-RISK CONNECTORS: At minimum, block consumer-grade connectors (Gmail, Twitter, etc.) from business environments"
                "5. ADD ENVIRONMENT POLICIES: For sensitive environments, add additional policies with tighter restrictions"
                "6. TEST WITH EXISTING APPS: Verify existing apps/flows still function after policies applied"
                "7. REPLICATE IN TARGET: Recreate same DLP policy structure in target tenant before migration"
                "8. DOCUMENT POLICY EXCEPTIONS: Record any approved exceptions for compliance documentation"
            )
            Tools = @("Power Platform Admin Center", "Microsoft.PowerApps.Administration.PowerShell")
            EstimatedEffort = "4-8 hours initial setup, ongoing governance required"
            Prerequisites = @("Power Platform Admin role", "Connector inventory complete")
        }
        @{
            Id          = "PP-003"
            Category    = "Integration"
            Name        = "Dynamics 365 Licensed Users"
            Condition   = { param($data)
                $data.Dynamics365.Users.Analysis.TotalDynamicsUsers -gt 0
            }
            Severity    = "High"
            Description = "Users have Dynamics 365 licenses — security roles, business units, and data access must be recreated in target environment"
            Recommendation = "Document all D365 user security roles and business unit assignments. Plan security model recreation before migrating users."
            RemediationSteps = @(
                "1. EXPORT USER SECURITY ROLES: Export all user-to-security-role mappings via Advanced Find or Dataverse API"
                "2. EXPORT BUSINESS UNITS: Document business unit hierarchy and user membership"
                "3. EXPORT FIELD SECURITY PROFILES: Query FieldSecurityProfile entity to document column-level permissions"
                "4. RECREATE SECURITY MODEL: In target D365 environment, recreate business unit structure first"
                "5. IMPORT SECURITY ROLES: Import or recreate security roles before user migration"
                "6. ASSIGN LICENSES IN TARGET: Ensure Dynamics 365 licenses are provisioned for all users in target"
                "7. BULK-ASSIGN USER ROLES: After user migration, reassign security roles via Dataverse API or XrmToolBox"
                "8. TEST ACCESS: Verify representative users can access correct records post-migration"
            )
            Tools = @("Power Platform Admin Center", "XrmToolBox", "Dataverse Web API", "Advanced Find")
            EstimatedEffort = "2-5 days depending on security model complexity"
            Prerequisites = @("Target D365 environment created", "D365 licenses in target", "Security model documented")
        }

        @{
            Id          = "CO-005"
            Category    = "Operations"
            Name        = "Backup/Restore Strategy Missing"
            Condition   = { param($data) $true }  # Always check
            Severity    = "Critical"
            Description = "Native M365 recoverability has limits. Migration without backup strategy risks data loss."
            Recommendation = "Implement backup solution for M365 data before migration. Validate restore capability."
            RemediationSteps = @(
                "1. ASSESS NATIVE LIMITS: Understand M365 retention (deleted items, version history, recycle bin)"
                "2. IDENTIFY GAPS: Determine what's not covered by native retention"
                "3. EVALUATE SOLUTIONS: Consider Veeam, AvePoint, Druva, etc. for M365 backup"
                "4. BACKUP SOURCE: Take full backup of source tenant before migration"
                "5. BACKUP TARGET: Continue backup coverage in target tenant post-migration"
                "6. TEST RESTORE: Validate ability to restore mailbox, OneDrive, SharePoint, Teams"
                "7. DOCUMENT RPO/RTO: Define recovery point and time objectives"
                "8. COMMUNICATE: Ensure users know recovery options and process"
            )
            Tools = @("Third-party backup solutions", "Native retention policies")
            EstimatedEffort = "Varies by solution - 8-40 hours implementation"
            Prerequisites = @("Backup solution selected and licensed", "Restore testing complete")
        }
    )
}
#endregion

#region Analysis Functions
function Invoke-GotchaAnalysis {
    <#
    .SYNOPSIS
        Performs comprehensive gotcha analysis on collected data
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeAllRules
    )

    Write-Log -Message "Starting gotcha analysis..." -Level Info

    $rules = Get-AnalysisRules
    $triggeredRules = [System.Collections.ArrayList]@()
    $analysisResults = @{
        Timestamp       = Get-Date
        RulesEvaluated  = $rules.Count
        RulesTriggered  = 0
        BySeverity      = @{
            Critical      = [System.Collections.ArrayList]@()
            High          = [System.Collections.ArrayList]@()
            Medium        = [System.Collections.ArrayList]@()
            Low           = [System.Collections.ArrayList]@()
            Informational = [System.Collections.ArrayList]@()
        }
        ByCategory      = @{}
        RiskScore       = 0
        MaxRiskScore    = 0
    }

    foreach ($rule in $rules) {
        try {
            $triggered = & $rule.Condition -data $CollectedData

            if ($triggered -and $triggered -ne $false -and $triggered -ne 0) {
                $triggeredRule = @{
                    Id               = $rule.Id
                    Category         = $rule.Category
                    Name             = $rule.Name
                    Severity         = $rule.Severity
                    Description      = $rule.Description
                    Recommendation   = $rule.Recommendation
                    RemediationSteps = if ($rule.RemediationSteps) { $rule.RemediationSteps } else { @() }
                    Tools            = if ($rule.Tools) { $rule.Tools } else { @() }
                    EstimatedEffort  = if ($rule.EstimatedEffort) { $rule.EstimatedEffort } else { "Not estimated" }
                    Prerequisites    = if ($rule.Prerequisites) { $rule.Prerequisites } else { @() }
                    TriggeredValue   = $triggered
                }

                $null = $triggeredRules.Add($triggeredRule)
                $null = $analysisResults.BySeverity[$rule.Severity].Add($triggeredRule)

                if (-not $analysisResults.ByCategory.ContainsKey($rule.Category)) {
                    $analysisResults.ByCategory[$rule.Category] = [System.Collections.ArrayList]@()
                }
                $null = $analysisResults.ByCategory[$rule.Category].Add($triggeredRule)
            }
        }
        catch {
            Write-Log -Message "Error evaluating rule $($rule.Id): $_" -Level Warning
        }
    }

    $analysisResults.RulesTriggered = $triggeredRules.Count
    $analysisResults.TriggeredRules = $triggeredRules

    # Calculate risk score
    $riskScore = 0
    $maxScore = 0

    foreach ($rule in $triggeredRules) {
        $severityWeight = $script:SeverityWeights[$rule.Severity]
        $categoryWeight = $script:RiskCategories[$rule.Category].Weight
        $riskScore += ($severityWeight * $categoryWeight)
    }

    # Max possible score
    foreach ($rule in $rules) {
        $severityWeight = $script:SeverityWeights[$rule.Severity]
        $categoryWeight = $script:RiskCategories[$rule.Category].Weight
        $maxScore += ($severityWeight * $categoryWeight)
    }

    $analysisResults.RiskScore = [math]::Round($riskScore, 2)
    $analysisResults.MaxRiskScore = [math]::Round($maxScore, 2)
    $analysisResults.RiskPercentage = if ($maxScore -gt 0) {
        [math]::Round(($riskScore / $maxScore) * 100, 2)
    } else { 0 }

    # Calculate risk level
    $analysisResults.RiskLevel = switch ($analysisResults.RiskPercentage) {
        { $_ -ge 75 } { "Critical"; break }
        { $_ -ge 50 } { "High"; break }
        { $_ -ge 25 } { "Medium"; break }
        default { "Low" }
    }

    Write-Log -Message "Analysis complete: $($triggeredRules.Count) of $($rules.Count) rules triggered. Risk level: $($analysisResults.RiskLevel)" -Level Info

    return $analysisResults
}

function Get-MigrationPriorities {
    <#
    .SYNOPSIS
        Generates prioritized action list based on analysis
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $AnalysisResults
    )

    $priorities = @{
        Immediate = [System.Collections.ArrayList]@()       # Critical items that block migration
        PreMigration = [System.Collections.ArrayList]@()    # Must be addressed before migration starts
        DuringMigration = [System.Collections.ArrayList]@() # Address during migration phases
        PostMigration = [System.Collections.ArrayList]@()   # Can be addressed after core migration
    }

    foreach ($rule in $AnalysisResults.TriggeredRules) {
        $priority = switch ($rule.Severity) {
            "Critical" { "Immediate" }
            "High"     { "PreMigration" }
            "Medium"   { "DuringMigration" }
            default    { "PostMigration" }
        }

        $null = $priorities[$priority].Add(@{
            Id             = $rule.Id
            Name           = $rule.Name
            Category       = $rule.Category
            Description    = $rule.Description
            Recommendation = $rule.Recommendation
        })
    }

    return $priorities
}

function Get-MigrationRoadmap {
    <#
    .SYNOPSIS
        Generates a suggested migration roadmap based on analysis
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $AnalysisResults,

        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData
    )

    $roadmap = @{
        Phase1_Discovery = @{
            Name     = "Discovery & Assessment"
            Duration = "2-4 weeks"
            Tasks    = [System.Collections.ArrayList]@(
                "Complete tenant discovery (DONE)"
                "Review and validate all gotchas identified"
                "Stakeholder alignment on migration scope"
                "Legal/compliance review for data handling"
            )
            Blockers = $AnalysisResults.BySeverity.Critical | ForEach-Object { $_.Name }
        }
        Phase2_Planning = @{
            Name     = "Detailed Planning"
            Duration = "3-6 weeks"
            Tasks    = [System.Collections.ArrayList]@(
                "Create detailed migration runbooks"
                "Establish target tenant and baseline configuration"
                "Set up coexistence if required"
                "Prepare user communication plan"
            )
        }
        Phase3_Preparation = @{
            Name     = "Pre-Migration Preparation"
            Duration = "2-4 weeks"
            Tasks    = [System.Collections.ArrayList]@(
                "Deploy and configure target tenant services"
                "Set up hybrid coexistence (if applicable)"
                "Migrate policies and configurations"
                "Pilot testing with small user group"
            )
        }
        Phase4_Migration = @{
            Name     = "Migration Execution"
            Duration = "Varies by scope"
            Tasks    = [System.Collections.ArrayList]@(
                "Batch user and mailbox migration"
                "SharePoint and OneDrive content migration"
                "Teams migration"
                "Application reconfiguration"
            )
        }
        Phase5_Validation = @{
            Name     = "Validation & Cutover"
            Duration = "1-2 weeks"
            Tasks    = [System.Collections.ArrayList]@(
                "Validate all migrated content and access"
                "Perform DNS cutover"
                "Decommission source tenant services"
                "Final user acceptance testing"
            )
        }
        Phase6_PostMigration = @{
            Name     = "Post-Migration"
            Duration = "2-4 weeks"
            Tasks    = [System.Collections.ArrayList]@(
                "Address post-migration items"
                "User training and support"
                "Performance optimization"
                "Project closure and documentation"
            )
        }
    }

    # Add specific tasks based on gotchas
    foreach ($rule in $AnalysisResults.TriggeredRules) {
        switch ($rule.Category) {
            "Identity" {
                if ($roadmap.Phase3_Preparation.Tasks -notcontains "Configure identity infrastructure") {
                    $null = $roadmap.Phase3_Preparation.Tasks.Add("Configure identity infrastructure")
                }
            }
            "Compliance" {
                if ($roadmap.Phase2_Planning.Tasks -notcontains "Legal review of compliance requirements") {
                    $null = $roadmap.Phase2_Planning.Tasks.Add("Legal review of compliance requirements")
                }
            }
            "Integration" {
                if ($roadmap.Phase4_Migration.Tasks -notcontains "Application SSO reconfiguration") {
                    $null = $roadmap.Phase4_Migration.Tasks.Add("Application SSO reconfiguration")
                }
            }
        }
    }

    return $roadmap
}

function Get-ComplexityScore {
    <#
    .SYNOPSIS
        Calculates overall migration complexity score based on comprehensive factors
    .DESCRIPTION
        Evaluates 15+ factors across identity, data, devices, compliance, and operations
        to produce a weighted complexity score from 0-100
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $true)]
        $AnalysisResults
    )

    # Calculate derived values
    $licensedUsers = $CollectedData.EntraID.Users.Analysis.LicensedUsers
    $sharedMailboxes = @($CollectedData.Exchange.Mailboxes.Mailboxes | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" }).Count
    $hybridDevices = $CollectedData.EntraID.Devices.Analysis.HybridJoined
    $intuneDevices = $CollectedData.EntraID.Devices.Analysis.IntuneManaged
    $guestUsers = $CollectedData.EntraID.Users.Analysis.GuestUsers
    $teamsCount = $CollectedData.Teams.Teams.Analysis.TotalTeams
    $criticalGotchas = ($AnalysisResults.BySeverity.Critical | Measure-Object).Count
    $highGotchas = ($AnalysisResults.BySeverity.High | Measure-Object).Count

    $factors = @{
        # Identity Factors (25% total weight)
        LicensedUsers = @{
            Value       = $licensedUsers
            Weight      = 0.10
            Description = "Licensed users to migrate"
            Thresholds  = @{ Low = 100; Medium = 500; High = 2000 }
        }
        HybridIdentity = @{
            Value       = if ($CollectedData.HybridIdentity.AADConnect.Configuration.OnPremisesSyncEnabled) { 100 } else { 0 }
            Weight      = 0.10
            Description = "Hybrid identity (AAD Connect)"
            Thresholds  = @{ Low = 0; Medium = 50; High = 75 }
        }
        GuestUsers = @{
            Value       = $guestUsers
            Weight      = 0.05
            Description = "Guest users requiring reinvitation"
            Thresholds  = @{ Low = 20; Medium = 100; High = 500 }
        }

        # Data Factors (30% total weight)
        MailboxCount = @{
            Value       = $CollectedData.Exchange.Mailboxes.Analysis.TotalMailboxes
            Weight      = 0.10
            Description = "Total mailboxes to migrate"
            Thresholds  = @{ Low = 100; Medium = 500; High = 2000 }
        }
        SharedMailboxes = @{
            Value       = $sharedMailboxes
            Weight      = 0.05
            Description = "Shared mailboxes (permissions complexity)"
            Thresholds  = @{ Low = 10; Medium = 50; High = 200 }
        }
        SharePointSites = @{
            Value       = $CollectedData.SharePoint.Sites.Analysis.SharePointSites
            Weight      = 0.08
            Description = "SharePoint sites to migrate"
            Thresholds  = @{ Low = 25; Medium = 100; High = 500 }
        }
        TeamsCount = @{
            Value       = $teamsCount
            Weight      = 0.07
            Description = "Microsoft Teams to migrate"
            Thresholds  = @{ Low = 20; Medium = 100; High = 500 }
        }

        # Device Factors (15% total weight)
        HybridDevices = @{
            Value       = $hybridDevices
            Weight      = 0.08
            Description = "Hybrid Azure AD joined devices"
            Thresholds  = @{ Low = 50; Medium = 200; High = 1000 }
        }
        IntuneManagedDevices = @{
            Value       = $intuneDevices
            Weight      = 0.07
            Description = "Intune managed devices (re-enrollment needed)"
            Thresholds  = @{ Low = 50; Medium = 200; High = 1000 }
        }

        # Integration Factors (10% total weight)
        AppRegistrations = @{
            Value       = $CollectedData.EntraID.Applications.Analysis.TotalApplications
            Weight      = 0.05
            Description = "App registrations to recreate"
            Thresholds  = @{ Low = 10; Medium = 50; High = 200 }
        }
        EnterpriseApps = @{
            Value       = $CollectedData.EntraID.Applications.Analysis.TotalServicePrincipals
            Weight      = 0.05
            Description = "Enterprise applications/integrations"
            Thresholds  = @{ Low = 25; Medium = 100; High = 300 }
        }

        # Compliance Factors (10% total weight)
        ComplianceIssues = @{
            Value       = ($AnalysisResults.ByCategory["Compliance"] | Measure-Object).Count
            Weight      = 0.10
            Description = "Compliance-related gotchas triggered"
            Thresholds  = @{ Low = 1; Medium = 3; High = 6 }
        }

        # Risk Factors (10% total weight) - Based on discovered gotchas
        CriticalGotchas = @{
            Value       = $criticalGotchas
            Weight      = 0.06
            Description = "Critical severity issues found"
            Thresholds  = @{ Low = 0; Medium = 2; High = 5 }
        }
        HighGotchas = @{
            Value       = $highGotchas
            Weight      = 0.04
            Description = "High severity issues found"
            Thresholds  = @{ Low = 2; Medium = 5; High = 10 }
        }
    }

    $totalScore = 0
    $breakdown = @{}

    foreach ($factor in $factors.GetEnumerator()) {
        $value = $factor.Value.Value
        if ($null -eq $value) { $value = 0 }
        $thresholds = $factor.Value.Thresholds
        $weight = $factor.Value.Weight

        $score = switch ($value) {
            { $_ -le $thresholds.Low } { 25; break }
            { $_ -le $thresholds.Medium } { 50; break }
            { $_ -le $thresholds.High } { 75; break }
            default { 100 }
        }

        $weightedScore = $score * $weight
        $totalScore += $weightedScore

        $breakdown[$factor.Key] = @{
            Description   = $factor.Value.Description
            RawValue      = $value
            Score         = $score
            Weight        = $weight
            WeightedScore = [math]::Round($weightedScore, 2)
        }
    }

    $complexityLevel = switch ($totalScore) {
        { $_ -le 30 } { "Low"; break }
        { $_ -le 50 } { "Medium"; break }
        { $_ -le 70 } { "High"; break }
        default { "Very High" }
    }

    # Generate summary
    $topFactors = $breakdown.GetEnumerator() |
        Sort-Object { $_.Value.WeightedScore } -Descending |
        Select-Object -First 5

    return @{
        TotalScore      = [math]::Round($totalScore, 2)
        ComplexityLevel = $complexityLevel
        Breakdown       = $breakdown
        TopFactors      = $topFactors
        Summary         = @{
            LicensedUsers     = $licensedUsers
            SharedMailboxes   = $sharedMailboxes
            TotalMailboxes    = $CollectedData.Exchange.Mailboxes.Analysis.TotalMailboxes
            SharePointSites   = $CollectedData.SharePoint.Sites.Analysis.SharePointSites
            Teams             = $teamsCount
            HybridIdentity    = $CollectedData.HybridIdentity.AADConnect.Configuration.OnPremisesSyncEnabled
            HybridDevices     = $hybridDevices
            IntuneDevices     = $intuneDevices
            CriticalIssues    = $criticalGotchas
            HighIssues        = $highGotchas
            TotalGotchas      = $AnalysisResults.RulesTriggered
        }
    }
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Invoke-GotchaAnalysis',
    'Get-MigrationPriorities',
    'Get-MigrationRoadmap',
    'Get-ComplexityScore',
    'Get-AnalysisRules'
)
