# M365 Migration Gotchas Reference

Complete reference of all 80 migration risk detection rules checked by the Gotcha Analysis Engine.

**Total Rules: 80** | Critical: 18 | High: 36 | Medium: 22

---

## Identity & Access Management (11 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| ID-001 | Synced Users Assessment | High | Synced users > 0 |
| ID-002 | Guest User Presence | Medium | Guest users > 100 |
| ID-003 | Hybrid Devices | Critical | Hybrid AD joined devices > 0 |
| ID-004 | Conditional Access Complexity | High | CA policies > 20 |
| ID-005 | Custom Roles | Medium | Custom role definitions > 0 |
| ID-006 | Cloud-Only Users with ImmutableId | High | Cloud-only users with ImmutableId > 0 |
| ID-007 | Users with External Identity Providers | High | Users with federated external IdP > 0 |
| ID-008 | Privileged Identity Management (PIM) Active | High | Privileged role assignments > 0 |
| ID-009 | Service Accounts Without Password Expiry | Medium | Users with DisablePasswordExpiration > 0 |
| ID-010 | Multiple UPN Suffixes | Medium | UPN suffixes > 3 |
| ID-011 | Azure AD B2B Direct Federation | Critical | Verified domains with Federated auth exists |

### ID-001: Synced Users Assessment
- **Severity:** High
- **Description:** On-premises synced users require special handling for ImmutableID preservation.
- **Recommendation:** Document sync configuration and plan ImmutableID strategy.
- **Remediation:**
  1. Export source data: `Get-MgUser -All | Select DisplayName, UserPrincipalName, OnPremisesImmutableId, OnPremisesSyncEnabled | Export-Csv SyncedUsers.csv`
  2. Document AAD Connect config on the AAD Connect server
  3. Choose strategy: Hard Match (preserve ImmutableID) or Soft Match (match by UPN/ProxyAddresses)
  4. For Hard Match: Pre-create users in target with same ImmutableID
  5. Configure target AAD Connect with same OU filtering and sync rules
  6. Validate sync in target tenant
  7. Cutover: Disable sync in source, wait 72hrs, enable in target
- **Tools:** Azure AD Connect, Microsoft Graph PowerShell, ADSyncTools
- **Effort:** 2-5 days depending on user count

### ID-002: Guest User Presence
- **Severity:** Medium
- **Description:** Large number of guest users will require reinvitation in target tenant.
- **Recommendation:** Automate guest reinvitation process using PowerShell scripts.
- **Remediation:**
  1. Export guest list with `Get-MgUser -Filter "userType eq 'Guest'"`
  2. Document permissions and group memberships for each guest
  3. Prepare bulk invitation script using `New-MgInvitation`
  4. Restore permissions after guests accept
  5. Notify guests about new invitation
  6. Validate guest count and permissions between tenants
- **Tools:** Microsoft Graph PowerShell, Excel for tracking
- **Effort:** 1-2 days + guest acceptance time

### ID-003: Hybrid Devices
- **Severity:** Critical
- **Description:** Hybrid Azure AD joined devices must be unjoined and rejoined to target.
- **Recommendation:** Plan phased device migration with user communication strategy.
- **Remediation:**
  1. Inventory devices: `Get-MgDevice -Filter "trustType eq 'ServerAd'"`
  2. Create Service Connection Point for target tenant
  3. Unjoin from source: `dsregcmd /leave` on each device
  4. Deploy automated unjoin via Intune/SCCM
  5. Trigger rejoin after SCP updated
  6. Verify join status with `dsregcmd /status`
  7. Update Conditional Access policies in target
- **Tools:** dsregcmd.exe, Group Policy, Intune/SCCM
- **Effort:** 3-7 days for phased rollout

### ID-004: Conditional Access Complexity
- **Severity:** High
- **Description:** Complex Conditional Access policy set requires careful recreation.
- **Recommendation:** Export and document all CA policies. Plan phased deployment in target.
- **Remediation:**
  1. Export all policies to JSON
  2. Export named locations
  3. Document all group/app/user dependencies
  4. Create named locations in target first
  5. Import policies in Report-Only mode
  6. Map all object IDs to target equivalents
  7. Test with pilot group, enable incrementally
- **Tools:** Microsoft Graph PowerShell, CA Documentation Workbook, Sign-in Logs
- **Effort:** 3-5 days

### ID-005: Custom Roles
- **Severity:** Medium
- **Description:** Custom role definitions must be recreated in target tenant.
- **Recommendation:** Export custom role definitions and permissions. Test in target before migration.
- **Remediation:**
  1. Export custom roles and permissions
  2. Document role assignments
  3. Create roles in target with exported permissions
  4. Verify permissions match
  5. Reassign roles after user migration
- **Tools:** Microsoft Graph PowerShell, Entra ID Portal
- **Effort:** 1-2 days

### ID-006: Cloud-Only Users with ImmutableId
- **Severity:** High
- **Description:** Cloud-only users with ImmutableId set may indicate previous sync or manual configuration.
- **Recommendation:** Review and document these users. ImmutableId preservation requires special handling.

### ID-007: Users with External Identity Providers
- **Severity:** High
- **Description:** Users with external identity provider federation require special migration planning.
- **Recommendation:** Document external IdP configurations. Plan for IdP trust recreation in target.

### ID-008: Privileged Identity Management (PIM) Active
- **Severity:** High
- **Description:** PIM role assignments detected - need careful recreation in target.
- **Recommendation:** Export PIM configurations. Plan for PIM setup before role assignment migration.

### ID-009: Service Accounts Without Password Expiry
- **Severity:** Medium
- **Description:** Service accounts with disabled password expiry need security review.
- **Recommendation:** Document service accounts. Review if managed identities can replace them in target.

### ID-010: Multiple UPN Suffixes
- **Severity:** Medium
- **Description:** Multiple UPN suffixes in use - all domains must be verified in target tenant.
- **Recommendation:** Document all UPN suffixes. Plan domain verification sequence in target.

### ID-011: Azure AD B2B Direct Federation
- **Severity:** Critical
- **Description:** B2B direct federation requires partner coordination during migration.
- **Recommendation:** Document all B2B federation partners. Coordinate cutover timing with partners.

---

## Data & Content Migration (10 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| DT-001 | Large Mailbox Detection | Medium | Mailboxes > 50GB exist |
| DT-002 | Public Folders Present | High | Public folders enabled |
| DT-003 | Large SharePoint Sites | High | Sites > 100GB exist |
| DT-004 | Archive Mailboxes | Medium | Archive-enabled mailboxes > 0 |
| DT-005 | In-Place Archive Mailboxes | High | Archive-enabled mailboxes > 100 |
| DT-006 | OneDrive Large Storage Users | Medium | OneDrive users > 50GB exist |
| DT-007 | SharePoint Custom Solutions | Critical | Custom solutions > 0 |
| DT-008 | SharePoint Workflows Active | High | Active workflows > 0 |
| DT-009 | External Sharing Enabled | Medium | External sharing enabled |
| DT-010 | Teams with External Members | Medium | Teams with guests > 0 |

### DT-001: Large Mailbox Detection
- **Severity:** Medium
- **Description:** Large mailboxes require extended migration windows.
- **Recommendation:** Plan incremental mailbox migration. Consider archive strategy.
- **Remediation:**
  1. Identify mailboxes > 50GB
  2. Calculate migration time (~1GB/hour for cross-tenant)
  3. Enable archive and move older items first
  4. Create separate migration batch for large mailboxes
  5. Schedule off-hours, monitor progress
  6. Use incremental sync: initial sync, then delta syncs until cutover
- **Tools:** Exchange Online PowerShell, Migration Batch cmdlets, BitTitan MigrationWiz (optional)
- **Effort:** 2-4 days per batch of large mailboxes

### DT-002: Public Folders Present
- **Severity:** High
- **Description:** Public folders require specialized migration approach.
- **Recommendation:** Plan batch migration. Consider modernization to M365 Groups.
- **Remediation:**
  1. Inventory public folders with sizes and item counts
  2. Check size limits (max 1M items, 50GB per folder)
  3. Decide: migrate as-is or modernize to M365 Groups/Shared Mailboxes
  4. For migration: create mapping CSV, lock source folders, create migration batch
  5. For modernization: use PF to Groups migration tool in EAC
  6. Recreate permissions in target
- **Tools:** Exchange Online PowerShell, EAC Migration wizard
- **Effort:** 3-7 days

### DT-003: Large SharePoint Sites
- **Severity:** High
- **Description:** Large SharePoint sites (>100GB) require incremental migration.
- **Recommendation:** Use pre-staging approach. Plan for extended migration windows.
- **Remediation:**
  1. Identify large sites and analyze content
  2. Run SharePoint Migration Assessment Tool (SMAT) for blockers
  3. Start pre-staging weeks before cutover
  4. Enable incremental syncs daily until cutover
  5. Handle special content: workflows, custom solutions, InfoPath
  6. Create user mapping file, validate post-migration
- **Tools:** SPMT, ShareGate, SMAT, PnP PowerShell
- **Effort:** 1-2 weeks for pre-staging plus cutover

### DT-004: Archive Mailboxes
- **Severity:** Medium
- **Description:** Archive mailboxes require separate migration consideration.
- **Recommendation:** Plan archive migration. Verify licensing in target tenant.
- **Effort:** Additional 50% time on top of primary mailbox migration

### DT-005: In-Place Archive Mailboxes
- **Severity:** High
- **Description:** Large number of archive mailboxes (>100) requires extended migration windows.
- **Recommendation:** Plan incremental archive migration. Consider archive-first approach.

### DT-006: OneDrive Large Storage Users
- **Severity:** Medium
- **Description:** Users with OneDrive storage exceeding 50GB require extended migration.
- **Recommendation:** Identify large OneDrive users. Plan pre-staging approach.

### DT-007: SharePoint Custom Solutions
- **Severity:** Critical
- **Description:** Custom SharePoint solutions (SPFx, Apps) require redevelopment or migration.
- **Recommendation:** Inventory all custom solutions. Plan for solution migration or rebuild.

### DT-008: SharePoint Workflows Active
- **Severity:** High
- **Description:** Active SharePoint workflows (2010/2013/Power Automate) need migration planning.
- **Recommendation:** Document all workflows. Plan for Power Automate rebuild.

### DT-009: External Sharing Enabled
- **Severity:** Medium
- **Description:** External sharing configuration must be replicated in target tenant.
- **Recommendation:** Document sharing policies. External shares need re-establishing post-migration.

### DT-010: Teams with External Members
- **Severity:** Medium
- **Description:** Teams with external members require guest reinvitation in target.
- **Recommendation:** Document Teams guest membership. Plan guest re-invitation process.

---

## Compliance & Security (10 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| CP-001 | Litigation Hold Active | Critical | Mailboxes on litigation hold > 0 |
| CP-002 | Active eDiscovery Cases | Critical | Active eDiscovery cases > 0 |
| CP-003 | Sensitivity Labels with Protection | Critical | Protected labels > 0 |
| CP-004 | Retention Policies Active | High | Enabled retention policies > 0 |
| CP-005 | DLP Policies Enforced | High | Enforced DLP policies > 0 |
| CP-006 | Information Barriers Configured | Critical | IB policies enabled > 0 |
| CP-007 | Communication Compliance Policies | High | Active CC policies > 0 |
| CP-008 | Records Management Labels | Critical | Record labels > 0 |
| CP-009 | Insider Risk Management | High | IRM policies active > 0 |
| CP-010 | Audit Log Retention Policies | Medium | Custom audit retention exists |

### CP-001: Litigation Hold Active
- **Severity:** Critical
- **Description:** Mailboxes under litigation hold must maintain legal compliance throughout migration.
- **Recommendation:** Coordinate with legal. Maintain chain of custody documentation.
- **Remediation:**
  1. Engage legal immediately - get written approval before any migration
  2. Document all holds with dates, owners, and durations
  3. Create chain of custody documentation with timestamps
  4. Enable hold in target BEFORE migration
  5. Migrate with verification - item counts must match
  6. Retain source mailbox as inactive for legal protection period
  7. Provide legal with signed attestation that hold was maintained
- **Tools:** Exchange Online PowerShell, eDiscovery portal
- **Effort:** Variable - legal coordination is the bottleneck

### CP-002: Active eDiscovery Cases
- **Severity:** Critical
- **Description:** Active eDiscovery cases require legal coordination before migration.
- **Recommendation:** Engage legal team. Document all case data and holds.

### CP-003: Sensitivity Labels with Protection
- **Severity:** Critical
- **Description:** Protected content requires Azure RMS migration planning.
- **Recommendation:** Plan for label GUID preservation or content re-protection.

### CP-004: Retention Policies Active
- **Severity:** High
- **Description:** Retention policies must be recreated to maintain compliance.
- **Recommendation:** Document all policies. Recreate before content migration.

### CP-005: DLP Policies Enforced
- **Severity:** High
- **Description:** DLP policies in enforcement mode need recreation.
- **Recommendation:** Export DLP configurations. Deploy in test mode initially.

### CP-006: Information Barriers Configured
- **Severity:** Critical
- **Description:** Information barriers require recreation before user migration.
- **Recommendation:** Export IB policies and segments. Recreate in target before migrating users.

### CP-007: Communication Compliance Policies
- **Severity:** High
- **Description:** Communication compliance policies must be recreated in target.
- **Recommendation:** Document all CC policies including custom classifiers. Plan recreation.

### CP-008: Records Management Labels
- **Severity:** Critical
- **Description:** Records management labels with regulatory requirements need careful migration.
- **Recommendation:** Engage legal/compliance team. Document retention schedules and disposition.

### CP-009: Insider Risk Management
- **Severity:** High
- **Description:** Insider risk policies and alerts need recreation in target.
- **Recommendation:** Document IRM policies. Historical alerts cannot be migrated. Plan fresh baseline.

### CP-010: Audit Log Retention Policies
- **Severity:** Medium
- **Description:** Custom audit log retention policies exist.
- **Recommendation:** Document retention periods. Export audit logs before migration if required.

---

## Applications & Integrations (11 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| IN-001 | Enterprise Applications | High | Service principals > 50 |
| IN-002 | Custom Teams Apps | High | Custom Teams apps > 0 |
| IN-003 | Power Platform Environments | Critical | Environments > 0 |
| IN-004 | Power BI Gateways | Critical | Gateways > 0 |
| IN-005 | Expiring App Credentials | High | Secrets or certs expiring within 90 days |
| IN-006 | Power Automate Flows | High | Flows > 50 |
| IN-007 | Power Apps Applications | High | Apps > 20 |
| IN-008 | Dataverse Environments | Critical | Dataverse environments > 0 |
| IN-009 | Custom Connectors | High | Custom connectors > 0 |
| IN-010 | Azure Logic Apps Integration | Medium | Integrated Logic Apps > 0 |
| IN-011 | Third-Party MDM Integration | High | Third-party MDM detected |

### IN-001: Enterprise Applications
- **Severity:** High
- **Description:** Large number of enterprise apps require SSO reconfiguration.
- **Recommendation:** Document SSO configurations. Plan phased app migration.

### IN-002: Custom Teams Apps
- **Severity:** High
- **Description:** Custom Teams apps must be republished in target tenant.
- **Recommendation:** Export app packages. Update manifests for target tenant.

### IN-003: Power Platform Environments
- **Severity:** Critical
- **Description:** Dynamics 365/Power Platform requires specialized migration.
- **Recommendation:** Plan dedicated Power Platform migration project.

### IN-004: Power BI Gateways
- **Severity:** Critical
- **Description:** On-premises gateways must be reinstalled for target tenant.
- **Recommendation:** Document gateway data sources. Plan reinstallation.

### IN-005: Expiring App Credentials
- **Severity:** High
- **Description:** Application credentials expiring within 90 days.
- **Recommendation:** Renew credentials before migration.

### IN-006: Power Automate Flows
- **Severity:** High
- **Description:** Large number of Power Automate flows require migration planning.
- **Recommendation:** Inventory all flows. Plan for flow export/import or recreation.

### IN-007: Power Apps Applications
- **Severity:** High
- **Description:** Power Apps require migration with data sources and connections.
- **Recommendation:** Document all Power Apps and their data sources. Plan connection recreation.

### IN-008: Dataverse Environments
- **Severity:** Critical
- **Description:** Dataverse environments with business data require dedicated migration.
- **Recommendation:** Plan dedicated Dataverse migration project.

### IN-009: Custom Connectors
- **Severity:** High
- **Description:** Custom connectors need recreation in target tenant.
- **Recommendation:** Export custom connector definitions. Recreate and test.

### IN-010: Azure Logic Apps Integration
- **Severity:** Medium
- **Description:** Logic Apps with M365 connections need connection updates.
- **Recommendation:** Document Logic App connections. Plan for connection recreation.

### IN-011: Third-Party MDM Integration
- **Severity:** High
- **Description:** Third-party MDM integration requires reconfiguration.
- **Recommendation:** Document MDM integration settings. Coordinate with MDM vendor.

---

## Hybrid Infrastructure (10 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| IF-001 | Federation Active | Critical | Federation in use |
| IF-002 | Pass-Through Authentication | High | PTA enabled |
| IF-003 | Directory Sync Active | Critical | On-premises sync enabled |
| IF-004 | Mail Flow Connectors | High | Inbound or outbound connectors > 0 |
| IF-005 | Password Hash Sync Enabled | Medium | PHS enabled |
| IF-006 | Seamless SSO Configured | High | Seamless SSO enabled |
| IF-007 | Multiple AAD Connect Servers | Medium | Staging server count > 0 |
| IF-008 | Group Writeback Enabled | High | Group writeback enabled |
| IF-009 | Device Writeback Enabled | High | Device writeback enabled |
| IF-010 | Exchange Hybrid Configuration | Critical | Hybrid enabled |

### IF-001: Federation Active
- **Severity:** Critical
- **Description:** Federated domains require careful cutover planning.
- **Recommendation:** Document ADFS config. Plan federation cutover strategy.

### IF-002: Pass-Through Authentication
- **Severity:** High
- **Description:** PTA agents must be deployed for target tenant.
- **Recommendation:** Plan PTA agent deployment. Consider staged rollout.

### IF-003: Directory Sync Active
- **Severity:** Critical
- **Description:** Azure AD Connect configuration must be addressed.
- **Recommendation:** Plan AAD Connect migration: new install or staged approach.

### IF-004: Mail Flow Connectors
- **Severity:** High
- **Description:** Mail flow connectors require recreation in target.
- **Recommendation:** Document connector configurations. Plan mail flow cutover.

### IF-005: Password Hash Sync Enabled
- **Severity:** Medium
- **Description:** Password Hash Sync configuration needs recreation for target.
- **Recommendation:** Document PHS configuration. Plan staged rollout in target.

### IF-006: Seamless SSO Configured
- **Severity:** High
- **Description:** Seamless SSO requires computer account recreation for target.
- **Recommendation:** Plan Seamless SSO cutover. New computer accounts needed.

### IF-007: Multiple AAD Connect Servers
- **Severity:** Medium
- **Description:** Multiple AAD Connect servers (staging mode) exist.
- **Recommendation:** Document all AAD Connect servers. Plan migration approach.

### IF-008: Group Writeback Enabled
- **Severity:** High
- **Description:** Group writeback to on-premises AD is configured.
- **Recommendation:** Document writeback configuration. Plan for writeback setup in target.

### IF-009: Device Writeback Enabled
- **Severity:** High
- **Description:** Device writeback to on-premises AD is configured.
- **Recommendation:** Document device writeback. Plan for Windows Hello for Business.

### IF-010: Exchange Hybrid Configuration
- **Severity:** Critical
- **Description:** Exchange hybrid configuration requires careful decommissioning.
- **Recommendation:** Document hybrid config. Plan for removal or reconfiguration.

---

## Operational Readiness (10 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| OP-001 | Teams Phone System | Critical | Phone system enabled |
| OP-002 | Private Channels | High | Private channels > 0 |
| OP-003 | Hub Sites Configured | Medium | Hub sites > 0 |
| OP-004 | Shared Channels with External Participants | Critical | Shared channels with external > 0 |
| OP-005 | Teams Templates in Use | Medium | Custom templates > 0 |
| OP-006 | Viva Insights Configured | Medium | Viva Insights enabled |
| OP-007 | Planner Plans Active | High | Plans > 100 |
| OP-008 | Bookings Configured | Medium | Bookings mailboxes > 0 |
| OP-009 | Stream Classic Videos | High | Classic videos > 0 |
| OP-010 | Forms with External Sharing | Medium | Externally shared forms > 0 |

### OP-001: Teams Phone System
- **Severity:** Critical
- **Description:** Teams Phone System requires dedicated migration project.
- **Recommendation:** Plan separate telephony migration. Document all configurations.

### OP-002: Private Channels
- **Severity:** High
- **Description:** Private channels have separate SharePoint sites that need individual migration.
- **Recommendation:** Plan private channel migration separately. Document memberships.

### OP-003: Hub Sites Configured
- **Severity:** Medium
- **Description:** Hub site associations must be recreated in target.
- **Recommendation:** Document hub site hierarchy. Recreate before site migration.

### OP-004: Shared Channels with External Participants
- **Severity:** Critical
- **Description:** Shared channels with external organizations require B2B direct connect.
- **Recommendation:** Document all shared channel relationships. Coordinate with external orgs.

### OP-005: Teams Templates in Use
- **Severity:** Medium
- **Description:** Custom Teams templates must be recreated in target tenant.
- **Recommendation:** Export template definitions. Recreate templates before team provisioning.

### OP-006: Viva Insights Configured
- **Severity:** Medium
- **Description:** Viva Insights historical data cannot be migrated.
- **Recommendation:** Document Insights configuration. Historical analytics will reset.

### OP-007: Planner Plans Active
- **Severity:** High
- **Description:** Large number of Planner plans (>100) require migration.
- **Recommendation:** Inventory Planner usage. Plans migrate with M365 Groups.

### OP-008: Bookings Configured
- **Severity:** Medium
- **Description:** Bookings calendars and configurations need recreation.
- **Recommendation:** Document Bookings pages. Plan for manual recreation in target.

### OP-009: Stream Classic Videos
- **Severity:** High
- **Description:** Stream Classic videos must be migrated to Stream on SharePoint.
- **Recommendation:** Plan Stream Classic to SharePoint migration before tenant migration.

### OP-010: Forms with External Sharing
- **Severity:** Medium
- **Description:** Forms shared externally will need new sharing links.
- **Recommendation:** Document externally shared forms. Plan to redistribute new links.

---

## Mail Flow & DNS (3 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| MF-001 | DNS Cutover Planning | Critical | Mailboxes > 0 |
| MF-002 | Mail Flow Coexistence Free/Busy | High | Mailboxes > 0 |
| MF-003 | Mailbox Migration Throttling | High | Mailboxes > 100 |

### MF-001: DNS Cutover Planning (MX/SPF/DKIM/DMARC)
- **Severity:** Critical
- **Description:** DNS records must be carefully cutover to avoid mail flow disruption.
- **Recommendation:** Plan DNS cutover sequence with reduced TTLs. Have rollback ready.

### MF-002: Mail Flow Coexistence Free/Busy
- **Severity:** High
- **Description:** During coexistence, Free/Busy lookups between tenants fail without Organization Relationships.
- **Recommendation:** Configure cross-tenant Organization Relationships and OAuth.

### MF-003: Mailbox Migration Throttling
- **Severity:** High
- **Description:** Microsoft throttles cross-tenant migrations. Large migrations require batching.
- **Recommendation:** Plan batched migration waves with 100-200 mailboxes per batch.

---

## Exchange Permissions & Calendars (4 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| EX-001 | Shared Mailbox Permissions Loss | Critical | Shared mailboxes > 0 |
| EX-002 | Delegate/Calendar Permissions Loss | High | Mailboxes > 0 |
| EX-003 | Distribution Group Migration | High | Distribution groups > 0 |
| EX-004 | Calendar Metadata and Meeting Links | Medium | Mailboxes > 0 |

### EX-001: Shared Mailbox Permissions Loss
- **Severity:** Critical
- **Description:** Shared mailbox permissions (Full Access, Send As, Send on Behalf) do not migrate automatically.
- **Recommendation:** Export all shared mailbox permissions. Re-apply in target post-migration.

### EX-002: Delegate/Calendar Permissions Loss
- **Severity:** High
- **Description:** Calendar delegate and folder permissions do not migrate.
- **Recommendation:** Document and export calendar/folder permissions. Users may need to re-grant access.

### EX-003: Distribution Group Migration
- **Severity:** High
- **Description:** Distribution groups must be recreated in target.
- **Recommendation:** Export DL configurations. Recreate in target before mailbox migration.

### EX-004: Calendar Metadata and Meeting Links
- **Severity:** Medium
- **Description:** Teams meeting links in calendar items point to source tenant.
- **Recommendation:** Document recurring meetings. Plan Teams meeting link updates.

---

## Teams Specific (4 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| TM-001 | Teams Chat History Not Migrated | Critical | Teams > 0 |
| TM-002 | Teams Meeting Recordings Location | High | Teams > 0 |
| TM-003 | Teams Channel Tabs and Connectors | High | Teams > 0 |
| TM-004 | Teams Wiki Content | Medium | Teams > 0 |

### TM-001: Teams Chat History Not Migrated
- **Severity:** Critical
- **Description:** 1:1 and group chat history does NOT migrate in cross-tenant migrations.
- **Recommendation:** Export chat history before migration. Set user expectations clearly.

### TM-002: Teams Meeting Recordings Location
- **Severity:** High
- **Description:** Teams meeting recordings stored in OneDrive/SharePoint will be in source tenant.
- **Recommendation:** Document recording locations. Plan to migrate or archive important recordings.

### TM-003: Teams Channel Tabs and Connectors
- **Severity:** High
- **Description:** Channel tabs and connectors do not migrate. Need reconfiguration.
- **Recommendation:** Document all channel tabs and connectors. Plan manual reconfiguration.

### TM-004: Teams Wiki Content
- **Severity:** Medium
- **Description:** Teams Wiki content being deprecated but existing wikis need consideration.
- **Recommendation:** Export Wiki content before migration. Consider migrating to OneNote.

---

## SharePoint & OneDrive (4 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| SP-001 | SharePoint Permissions Inheritance Break | Critical | SharePoint sites > 0 |
| SP-002 | SharePoint Metadata and Version History | High | SharePoint sites > 0 |
| SP-003 | OneDrive Shortcuts and Synced Folders | High | OneDrive sites > 0 |
| SP-004 | External Sharing Links Invalidation | High | External sharing enabled |

### SP-001: SharePoint Permissions Inheritance Break
- **Severity:** Critical
- **Description:** Unique permissions may not migrate correctly. Inheritance can reset.
- **Recommendation:** Document unique permissions before migration. Validate post-migration.

### SP-002: SharePoint Metadata and Version History
- **Severity:** High
- **Description:** Custom metadata, content types, and version history may not fully migrate.
- **Recommendation:** Validate migration tool supports metadata. Document version requirements.

### SP-003: OneDrive Shortcuts and Synced Folders
- **Severity:** High
- **Description:** OneDrive shortcuts to SharePoint will break. Synced folders need resync.
- **Recommendation:** Document shortcuts. Plan sync client reconfiguration post-migration.

### SP-004: External Sharing Links Invalidation
- **Severity:** High
- **Description:** All external sharing links will be invalidated after migration.
- **Recommendation:** Document critical external shares. Plan to regenerate and redistribute links.

---

## Devices & Intune (4 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| DV-001 | Intune Device Re-enrollment Required | Critical | Intune managed devices > 0 |
| DV-002 | AutoPilot Profile Migration | Critical | AutoPilot registered > 0 |
| DV-003 | Client Profile and Cached Credentials | High | Licensed users > 0 |
| DV-004 | Multi-Factor Authentication Disruption | High | Licensed users > 0 |

### DV-001: Intune Device Re-enrollment Required
- **Severity:** Critical
- **Description:** Intune-managed devices must be unenrolled from source and re-enrolled in target.
- **Recommendation:** Plan phased device re-enrollment. Export and recreate all Intune policies.

### DV-002: AutoPilot Profile Migration
- **Severity:** Critical
- **Description:** Windows AutoPilot registrations are tenant-specific. Must deregister and re-register.
- **Recommendation:** Export AutoPilot device list. Plan hardware hash re-import to target.

### DV-003: Client Profile and Cached Credentials
- **Severity:** High
- **Description:** Users need to sign out of all Office apps and re-authenticate with target tenant.
- **Recommendation:** Prepare user guide for clearing credentials. Plan for increased helpdesk calls.

### DV-004: Multi-Factor Authentication Disruption
- **Severity:** High
- **Description:** MFA registrations are tenant-specific. Users must re-register in target.
- **Recommendation:** Plan MFA re-enrollment. Consider temporary MFA bypass during cutover.

---

## Applications & API (3 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| AP-001 | App Registrations OAuth Token Reset | Critical | App registrations > 0 |
| AP-002 | Third-Party API Integrations Failure | High | Applications > 0 |
| AP-003 | Power BI Workspace Ownership | High | Workspaces > 0 |

### AP-001: App Registrations OAuth Token Reset
- **Severity:** Critical
- **Description:** App registrations are tenant-specific. All OAuth tokens become invalid. Apps need re-registration.
- **Recommendation:** Document all app registrations. Plan recreation and secret rotation.

### AP-002: Third-Party API Integrations Failure
- **Severity:** High
- **Description:** Third-party apps with OAuth/SAML will stop working until reconfigured.
- **Recommendation:** Inventory all third-party integrations. Coordinate with vendors.

### AP-003: Power BI Workspace Ownership
- **Severity:** High
- **Description:** Power BI workspaces need migration. Connections and gateway configs need recreation.
- **Recommendation:** Export Power BI content. Plan workspace recreation with proper ownership.

---

## Identity & Licensing (4 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| LI-001 | UPN vs Primary SMTP Address Mismatch | High | UPN domain differs from mail domain |
| LI-002 | License Entitlement Verification | High | Licensed users > 0 |
| LI-003 | Domain Verification Timing | Critical | Always checked |
| LI-004 | Security Group vs M365 Group Confusion | Medium | Groups > 0 |

### LI-001: UPN vs Primary SMTP Address Mismatch
- **Severity:** High
- **Description:** Users with different UPN and primary email address require special attention.
- **Recommendation:** Document mismatches. Decide on target naming convention.

### LI-002: License Entitlement Verification
- **Severity:** High
- **Description:** License counts and SKUs in target must match source.
- **Recommendation:** Audit license usage. Ensure target has equivalent licenses.

### LI-003: Domain Verification Timing
- **Severity:** Critical
- **Description:** Custom domains cannot exist in two tenants simultaneously.
- **Recommendation:** Plan domain cutover window carefully. Reduce DNS TTLs in advance.

### LI-004: Security Group vs M365 Group Confusion
- **Severity:** Medium
- **Description:** Security groups and M365 groups have different migration paths.
- **Recommendation:** Classify all groups by type. Plan appropriate migration method for each.

---

## Compliance & Operations Planning (5 rules)

| ID | Name | Severity | Trigger Condition |
|----|------|----------|-------------------|
| CO-001 | Missing Rollback/Contingency Plan | Critical | Always flagged |
| CO-002 | Insufficient Pilot Testing | High | Always flagged |
| CO-003 | Poor User Communication/Change Management | High | Always flagged |
| CO-004 | Audit Logs Not Preserved | High | Always flagged |
| CO-005 | Backup/Restore Strategy Missing | Critical | Always flagged |

### CO-001: Missing Rollback/Contingency Plan
- **Severity:** Critical
- **Description:** Without a tested rollback plan, migration failures cause extended outages.
- **Recommendation:** Document detailed rollback procedures. Test on pilot group first.

### CO-002: Insufficient Pilot Testing
- **Severity:** High
- **Description:** Skipping or rushing pilot testing leads to unexpected issues during full migration.
- **Recommendation:** Conduct thorough pilot migration with diverse user sample.

### CO-003: Poor User Communication/Change Management
- **Severity:** High
- **Description:** Inadequate user communication leads to confusion and increased support burden.
- **Recommendation:** Develop comprehensive communication plan with multiple touchpoints.

### CO-004: Audit Logs Not Preserved
- **Severity:** High
- **Description:** Unified audit logs in source tenant will not be available after decommissioning.
- **Recommendation:** Export audit logs before source decommission. Archive for compliance requirements.

### CO-005: Backup/Restore Strategy Missing
- **Severity:** Critical
- **Description:** Native M365 recoverability has limits. Migration without backup risks data loss.
- **Recommendation:** Implement backup solution for M365 data before beginning migration.
