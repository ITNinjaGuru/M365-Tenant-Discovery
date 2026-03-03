# M365 Tenant Discovery & Migration Assessment Tool

A comprehensive PowerShell 7-based tool for discovering and analyzing Microsoft 365 tenant configurations to identify migration risks, gotchas, and recommendations for tenant-to-tenant migrations.

## Overview

This tool is designed for migration architects and IT professionals planning M365 tenant-to-tenant migrations. It performs deep analysis of source tenant configurations and uses AI-powered insights to identify potential issues before they become blockers.

## Key Features

- **Comprehensive Discovery**: Collects data from all major M365 workloads
- **60+ Gotcha Detection Rules**: Built-in analysis to identify migration risks
- **AI-Powered Insights**: Integration with GPT-5.2, Claude Opus 4.6, or Gemini 3 Pro
- **Professional Reports**: HTML reports for IT teams and executive leadership
- **Risk Scoring**: Quantified complexity and risk assessment
- **Migration Roadmap**: AI-generated migration prioritization and phasing

## Supported Workloads

| Workload | What's Collected |
|----------|-----------------|
| **Microsoft Entra ID** | Users, Groups, Devices, Apps, Conditional Access, Roles, Named Locations, Administrative Units, PIM |
| **Exchange Online** | Mailboxes, Distribution Lists, Public Folders, Transport Rules, Connectors, Hybrid Config, Journaling, Resources |
| **SharePoint Online** | Sites, Hub Sites, Term Store, Content Types, OneDrive, External Sharing, Site Designs |
| **Microsoft Teams** | Teams, Channels (Private/Shared), Policies, Apps, Phone System, Governance |
| **Power BI** | Workspaces, Gateways, Capacities, Datasets, Reports |
| **Dynamics 365 / Power Platform** | Environments, Power Apps, Power Automate, Connectors, DLP Policies, Dataverse Solutions |
| **Security & Compliance** | DLP, Retention, Sensitivity Labels, eDiscovery, Information Barriers, Insider Risk |
| **Hybrid Identity** | AAD Connect, Federation, PTA, Seamless SSO, Device Writeback, Application Proxy |

## Requirements

### PowerShell Version
- **PowerShell 7.0 or later** (required)
- Windows PowerShell 5.1 is NOT supported

### Required Modules
```powershell
# Install required modules
Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
Install-Module -Name MicrosoftTeams -Scope CurrentUser -Force
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
```

### Permissions Required

| Service | Required Role/Permission |
|---------|-------------------------|
| Entra ID | Global Reader or Directory Reader + Application Administrator |
| Exchange Online | Exchange Administrator or Organization Management |
| SharePoint Online | SharePoint Administrator |
| Teams | Teams Administrator |
| Security & Compliance | Compliance Administrator |
| Power BI | Power BI Administrator |
| Dynamics 365 | System Administrator + Power Platform Administrator |

### App Registration Setup (For Automated/Unattended Authentication)

To run the discovery tool without interactive authentication (e.g., scheduled tasks, CI/CD), create an Entra ID app registration:

1. **Create App Registration in Azure Portal**
   - Go to Azure Portal → Entra ID → App registrations → New registration
   - Name: `M365-Tenant-Discovery` (or your preference)
   - Supported account types: Single tenant
   - No redirect URI needed

2. **Add API Permissions**
   ```
   Microsoft Graph (Application permissions):
   - User.Read.All
   - Group.Read.All
   - Directory.Read.All
   - Device.Read.All
   - Application.Read.All
   - Policy.Read.All
   - RoleManagement.Read.All
   - Organization.Read.All

   SharePoint (Application permissions):
   - Sites.Read.All
   - Sites.FullControl.All (for complete site enumeration)
   ```

3. **Grant Admin Consent**
   - Click "Grant admin consent for [your org]"

4. **Create Certificate (REQUIRED for SharePoint)**

   SharePoint app-only access requires certificate authentication. Create a self-signed certificate:

   ```powershell
   # Create self-signed certificate (run as admin)
   $cert = New-SelfSignedCertificate -Subject "CN=M365TenantDiscovery" `
       -CertStoreLocation "Cert:\CurrentUser\My" `
       -KeyExportPolicy Exportable `
       -KeySpec Signature `
       -KeyLength 2048 `
       -KeyAlgorithm RSA `
       -HashAlgorithm SHA256 `
       -NotAfter (Get-Date).AddYears(2)

   # Export public key (.cer) for Azure upload
   Export-Certificate -Cert $cert -FilePath "M365TenantDiscovery.cer"

   # Note the thumbprint for your config
   Write-Host "Certificate Thumbprint: $($cert.Thumbprint)"
   ```

   Then upload the `.cer` file to your app registration:
   - Go to Certificates & secrets → Certificates → Upload certificate
   - Upload the `M365TenantDiscovery.cer` file

5. **Optional: Create Client Secret (for Graph/Teams only)**
   - Go to Certificates & secrets → New client secret
   - Copy the secret value (you won't see it again!)
   - Note: Client secrets do NOT work for SharePoint - use certificate

6. **Configure the Tool**
   - Copy `Config/discovery-config.sample.json` to `Config/discovery-config.json`
   - Update the Authentication section:
   ```json
   "Authentication": {
     "Method": "ServicePrincipal",
     "TenantId": "your-tenant-id",
     "ClientId": "your-app-client-id",
     "ClientSecret": "your-client-secret-for-graph",
     "CertificateThumbprint": "your-certificate-thumbprint",
     "SharePoint": {
       "AdminUrl": "https://yourtenant-admin.sharepoint.com"
     }
   }
   ```

7. **Run with Configuration**
   ```powershell
   .\Start-TenantDiscovery.ps1 -ConfigPath ".\Config\discovery-config.json"
   ```

> **Important**: SharePoint Online requires certificate authentication for Azure AD app-only access. Client secrets only work for legacy SharePoint Add-in apps.

## Quick Start

### Basic Usage
```powershell
# Navigate to the tool directory
cd path\to\tenantdiscovery-claude

# Run discovery with interactive authentication
.\Start-TenantDiscovery.ps1 -SharePointAdminUrl "https://contoso-admin.sharepoint.com"
```

### With AI Analysis
```powershell
# Using Claude Opus 4.6 (Recommended)
.\Start-TenantDiscovery.ps1 `
    -SharePointAdminUrl "https://contoso-admin.sharepoint.com" `
    -AIProvider "Opus4.6" `
    -AIApiKey $env:ANTHROPIC_API_KEY

# Using OpenAI GPT-5.2
.\Start-TenantDiscovery.ps1 `
    -SharePointAdminUrl "https://contoso-admin.sharepoint.com" `
    -AIProvider "GPT-5.2" `
    -AIApiKey $env:OPENAI_API_KEY

# Using Google Gemini 3 Pro
.\Start-TenantDiscovery.ps1 `
    -SharePointAdminUrl "https://contoso-admin.sharepoint.com" `
    -AIProvider "Gemini-3-Pro" `
    -AIApiKey $env:GOOGLE_API_KEY
```

### Selective Collection
```powershell
# Skip specific workloads for faster execution
.\Start-TenantDiscovery.ps1 `
    -SharePointAdminUrl "https://contoso-admin.sharepoint.com" `
    -SkipPowerBI `
    -SkipDynamics `
    -SkipSecurity
```

### Using Configuration File
```powershell
# Copy and customize configuration
Copy-Item .\Config\discovery-config.sample.json .\Config\discovery-config.json

# Run with configuration
.\Start-TenantDiscovery.ps1 -ConfigPath ".\Config\discovery-config.json"
```

## Parameters

| Parameter | Description | Required | Default |
|-----------|-------------|----------|---------|
| `-ConfigPath` | Path to JSON configuration file | No | None |
| `-OutputPath` | Directory for output files | No | ./Output |
| `-SharePointAdminUrl` | SharePoint admin center URL | Yes* | None |
| `-AIProvider` | AI provider: GPT-5.2, Opus4.6, or Gemini-3-Pro | No | None |
| `-AIApiKey` | API key for selected AI provider | Conditional | None |
| `-SkipAI` | Skip AI-powered analysis | No | False |
| `-SkipExchange` | Skip Exchange Online collection | No | False |
| `-SkipSharePoint` | Skip SharePoint Online collection | No | False |
| `-SkipTeams` | Skip Microsoft Teams collection | No | False |
| `-SkipPowerBI` | Skip Power BI collection | No | False |
| `-SkipDynamics` | Skip Dynamics 365/Power Platform collection | No | False |
| `-SkipSecurity` | Skip Security & Compliance collection | No | False |
| `-Interactive` | Use interactive authentication | No | True |

*Required only if SharePoint collection is not skipped

## Output Structure

```
Output/
└── Discovery_YYYYMMDD_HHMMSS/
    ├── Data/
    │   ├── TenantDiscovery_Full.json    # Complete collected data
    │   ├── EntraID.json
    │   ├── Exchange.json
    │   ├── SharePoint.json
    │   ├── Teams.json
    │   ├── PowerBI.json
    │   ├── Dynamics365.json
    │   ├── Security.json
    │   └── HybridIdentity.json
    ├── Reports/
    │   ├── IT_Technical_Report.html     # Detailed IT report with charts
    │   └── Executive_Summary.html       # Executive summary for leadership
    └── Logs/
        └── TenantDiscovery_*.log        # Execution log
```

## Migration Gotchas Detected (60+ Rules)

### Identity & Access (15+ rules)
- Synced users requiring ImmutableID preservation
- Cloud-only users with ImmutableId set
- Guest users needing reinvitation
- Hybrid Azure AD joined devices
- Complex Conditional Access policies
- Custom role definitions
- Privileged Identity Management (PIM) active
- External identity providers
- Multiple UPN suffixes
- B2B direct federation
- Service accounts without password expiry
- Azure AD Application Proxy configured

### Data & Content (15+ rules)
- Large mailboxes (>50GB)
- Public folders with large item count
- Archive mailboxes enabled
- Litigation holds
- In-place archive mailboxes
- Large OneDrive users (>50GB)
- SharePoint custom solutions
- Active workflows
- External sharing configurations
- Teams with external members

### Compliance & Security (12+ rules)
- Active eDiscovery cases
- Sensitivity labels with protection
- Retention policies in enforcement
- DLP policies in enforcement
- Information barriers configured
- Communication compliance policies
- Records management labels
- Insider risk management
- Audit log retention policies
- Journal rules configured

### Integrations & Applications (12+ rules)
- Enterprise applications with SSO
- Custom Teams apps
- Power Platform environments
- Power BI gateways
- Expiring application credentials
- Power Automate flows (especially active)
- Power Apps (Canvas and Model-Driven)
- Dataverse environments
- Custom connectors
- Azure Logic Apps integration
- Third-party MDM integration

### Hybrid Infrastructure (10+ rules)
- Federated domains
- Pass-through authentication
- Seamless SSO configured
- Azure AD Connect (multiple servers)
- Group writeback enabled
- Device writeback enabled
- Exchange hybrid configuration
- Password hash sync
- Migration endpoints active
- Organization relationships

### Operations (10+ rules)
- Teams Phone System
- Private channels
- Shared channels with external organizations
- Hub sites
- Teams templates
- Viva Insights
- Large Planner usage
- Bookings configured
- Stream Classic videos
- Forms with external sharing

## AI Provider Setup

### Claude Opus 4.6 (Anthropic) - Recommended
```powershell
# Set environment variable
$env:ANTHROPIC_API_KEY = "sk-ant-..."

# Or pass directly
.\Start-TenantDiscovery.ps1 -AIProvider "Opus4.6" -AIApiKey "sk-ant-..."
```
Get API key from [console.anthropic.com](https://console.anthropic.com)

### GPT-5.2 (OpenAI)
```powershell
$env:OPENAI_API_KEY = "sk-..."
.\Start-TenantDiscovery.ps1 -AIProvider "GPT-5.2" -AIApiKey $env:OPENAI_API_KEY
```
Get API key from [platform.openai.com](https://platform.openai.com)

### Gemini 3 Pro (Google)
```powershell
$env:GOOGLE_API_KEY = "..."
.\Start-TenantDiscovery.ps1 -AIProvider "Gemini-3-Pro" -AIApiKey $env:GOOGLE_API_KEY
```
Get API key from [aistudio.google.com](https://aistudio.google.com)

## Configuration File Reference

The configuration file (`discovery-config.json`) supports extensive customization:

```json
{
  "Collection": {
    "EntraID": { "Enabled": true, "MaxUsersToProcess": 50000 },
    "Exchange": { "Enabled": true, "IncludeHybridConfig": true },
    "SharePoint": { "AdminUrl": "https://contoso-admin.sharepoint.com" },
    "Teams": { "Enabled": true, "IncludePhoneSystem": true },
    "PowerBI": { "Enabled": true },
    "Dynamics365": { "Enabled": true, "IncludeDLPPolicies": true },
    "Security": { "Enabled": true },
    "HybridIdentity": { "Enabled": true, "IncludeApplicationProxy": true }
  },
  "AI": {
    "Enabled": true,
    "Provider": "Opus4.6",
    "Options": {
      "MaxTokens": 12000,
      "GenerateRemediationPlans": true,
      "GenerateExecutiveSummary": true,
      "IdentifyHiddenRisks": true
    }
  },
  "Reporting": {
    "GenerateITReport": true,
    "GenerateExecutiveReport": true,
    "IncludeCharts": true
  }
}
```

See `Config/discovery-config.sample.json` for all available options.

## Project Structure

```
tenantdiscovery-claude/
├── Start-TenantDiscovery.ps1           # Main orchestrator script (630 lines)
├── Modules/
│   ├── Core/
│   │   └── TenantDiscovery.Core.psm1   # Core utilities, logging, config
│   ├── EntraID/
│   │   └── TenantDiscovery.EntraID.psm1    # Users, groups, devices, apps, CA
│   ├── Exchange/
│   │   └── TenantDiscovery.Exchange.psm1   # Mailboxes, DLs, transport, hybrid
│   ├── SharePoint/
│   │   └── TenantDiscovery.SharePoint.psm1 # Sites, OneDrive, hub sites
│   ├── Teams/
│   │   └── TenantDiscovery.Teams.psm1      # Teams, channels, policies
│   ├── PowerBI/
│   │   └── TenantDiscovery.PowerBI.psm1    # Workspaces, gateways
│   ├── Dynamics365/
│   │   └── TenantDiscovery.Dynamics365.psm1 # Power Platform, D365
│   ├── Security/
│   │   └── TenantDiscovery.Security.psm1    # DLP, retention, labels
│   └── HybridIdentity/
│       └── TenantDiscovery.HybridIdentity.psm1 # AAD Connect, federation
├── Analysis/
│   ├── GotchaAnalysisEngine.psm1       # 60+ risk detection rules
│   └── AIIntegration.psm1              # GPT-5.2, Opus 4.6, Gemini 3 Pro
├── Reports/
│   └── ReportGenerator.psm1            # HTML report with Chart.js
├── Config/
│   └── discovery-config.sample.json    # Sample configuration
└── README.md
```

## HTML Report Features

### IT Technical Report
- Interactive charts (severity distribution, category breakdown)
- Detailed gotcha analysis with remediation steps
- Dependency mapping between issues
- PowerShell commands for investigation
- AI-generated detailed technical analysis

### Executive Summary Report
- Risk dashboard with visual indicators
- Business impact assessment
- Timeline and resource estimates
- Decision points for leadership
- AI-generated executive briefing

## Best Practices

1. **Run During Off-Peak Hours**: Large tenants may take 30-60 minutes to fully collect
2. **Use Global Reader Role**: Minimizes permission while allowing full discovery
3. **Enable AI Analysis**: Provides actionable insights beyond basic detection
4. **Review Reports Before Sharing**: Reports contain sensitive tenant information
5. **Validate AI Recommendations**: AI analysis should be reviewed by migration experts
6. **Run Multiple Times**: Execute before and during migration for comparison
7. **Archive Reports**: Keep discovery reports for post-migration validation

## Troubleshooting

### Authentication Issues
```powershell
# Clear all cached tokens
Disconnect-MgGraph -ErrorAction SilentlyContinue
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Disconnect-SPOService -ErrorAction SilentlyContinue

# Retry with fresh authentication
.\Start-TenantDiscovery.ps1 -SharePointAdminUrl "https://contoso-admin.sharepoint.com"
```

### Module Version Conflicts
```powershell
# Check versions
Get-Module Microsoft.Graph -ListAvailable
Get-Module ExchangeOnlineManagement -ListAvailable

# Update all modules
Update-Module Microsoft.Graph -Force
Update-Module ExchangeOnlineManagement -Force
Update-Module MicrosoftTeams -Force
```

### Timeout or Throttling
```powershell
# Increase timeout in configuration
{
  "Performance": {
    "TimeoutMinutes": 180,
    "RetryCount": 5,
    "RetryDelaySeconds": 10
  }
}
```

### AI Provider Errors
```powershell
# Test AI connection
.\Start-TenantDiscovery.ps1 -AIProvider "Opus4.6" -AIApiKey "your-key" -TestAIOnly
```

## Common Migration Scenarios

### Simple Cloud-to-Cloud
```powershell
.\Start-TenantDiscovery.ps1 `
    -SharePointAdminUrl "https://contoso-admin.sharepoint.com" `
    -AIProvider "Opus4.6" `
    -AIApiKey $env:ANTHROPIC_API_KEY
```

### Hybrid to Cloud
```powershell
# Include all hybrid identity checks
.\Start-TenantDiscovery.ps1 `
    -SharePointAdminUrl "https://contoso-admin.sharepoint.com" `
    -ConfigPath ".\Config\hybrid-focused-config.json" `
    -AIProvider "Opus4.6" `
    -AIApiKey $env:ANTHROPIC_API_KEY
```

### Large Enterprise (50k+ users)
```powershell
# Use configuration with higher limits
.\Start-TenantDiscovery.ps1 `
    -SharePointAdminUrl "https://contoso-admin.sharepoint.com" `
    -ConfigPath ".\Config\enterprise-config.json" `
    -OutputPath "D:\MigrationAssessment" `
    -AIProvider "Opus4.6" `
    -AIApiKey $env:ANTHROPIC_API_KEY
```

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 2.0.0 | Jan 2026 | Added 30+ new gotcha rules, enhanced AI prompts, Chart.js visualizations |
| 1.0.0 | Jan 2026 | Initial release with 30 gotcha rules |

## Support

For issues and feature requests, please open an issue in the repository.

## License

This tool is provided as-is for tenant migration assessment purposes.

## Disclaimer

This tool performs read-only operations and does not modify tenant configuration. However:
- Always ensure you have appropriate authorization before running discovery tools
- Reports contain sensitive tenant information - handle according to data classification policies
- AI-generated recommendations should be validated by experienced migration professionals
- The tool connects to multiple M365 services - ensure network access is available

---

**Version**: 2.0.0
**PowerShell**: 7.0+
**Last Updated**: January 2026
