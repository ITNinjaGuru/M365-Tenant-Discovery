#Requires -Version 7.0
<#
.SYNOPSIS
    AI Integration Module for M365 Migration Analysis
.DESCRIPTION
    Provides integration with multiple AI providers (GPT-5.2, Opus 4.6, Gemini-3-Pro, Gemini-3-Flash)
    for intelligent analysis of migration gotchas and recommendations.
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

#region Configuration
$script:AIProviders = @{
    "GPT-5.2" = @{
        Name     = "OpenAI GPT-5.2"
        Endpoint = "https://api.openai.com/v1/chat/completions"
        Model    = "gpt-5.2"
        Headers  = @{
            "Content-Type" = "application/json"
        }
        AuthHeader = "Authorization"
        AuthPrefix = "Bearer "
    }
    "Opus4.6" = @{
        Name     = "Anthropic Claude Opus 4.6"
        Endpoint = "https://api.anthropic.com/v1/messages"
        Model    = "claude-opus-4-6"
        Headers  = @{
            "Content-Type"      = "application/json"
            "anthropic-version" = "2023-06-01"
        }
        AuthHeader = "x-api-key"
        AuthPrefix = ""
    }
    "Gemini-3-Pro" = @{
        Name       = "Google Gemini 3 Pro"
        Endpoint   = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3-pro-preview:generateContent"
        Model      = "gemini-3-pro-preview"
        Headers    = @{
            "Content-Type" = "application/json"
        }
        AuthHeader = "url-param"
        AuthPrefix = ""
    }
    "Gemini-3-Flash" = @{
        Name       = "Google Gemini 3 Flash"
        Endpoint   = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent"
        Model      = "gemini-3-flash-preview"
        Headers    = @{
            "Content-Type" = "application/json"
        }
        AuthHeader = "url-param"
        AuthPrefix = ""
    }
}

$script:CurrentProvider = $null
$script:APIKey = $null
#endregion

#region Provider Configuration
function Set-AIProvider {
    <#
    .SYNOPSIS
        Configures the AI provider for analysis
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("GPT-5.2", "Opus4.6", "Gemini-3-Pro", "Gemini-3-Flash")]
        [string]$Provider,

        [Parameter(Mandatory = $true)]
        [string]$APIKey,

        [Parameter(Mandatory = $false)]
        [string]$CustomEndpoint
    )

    $script:CurrentProvider = $script:AIProviders[$Provider]
    $script:APIKey = $APIKey

    if ($CustomEndpoint) {
        $script:CurrentProvider.Endpoint = $CustomEndpoint
    }

    Write-Log -Message "AI Provider configured: $($script:CurrentProvider.Name)" -Level Info

    return @{
        Provider = $Provider
        Name     = $script:CurrentProvider.Name
        Endpoint = $script:CurrentProvider.Endpoint
    }
}

function Get-AIProvider {
    <#
    .SYNOPSIS
        Returns current AI provider configuration
    #>
    if ($script:CurrentProvider) {
        return @{
            Name     = $script:CurrentProvider.Name
            Model    = $script:CurrentProvider.Model
            Endpoint = $script:CurrentProvider.Endpoint
        }
    }
    return $null
}

function Test-AIConnection {
    <#
    .SYNOPSIS
        Tests connectivity to the configured AI provider
    #>
    [CmdletBinding()]
    param()

    if (-not $script:CurrentProvider -or -not $script:APIKey) {
        return @{
            Success = $false
            Error   = "AI provider not configured. Call Set-AIProvider first."
        }
    }

    try {
        $testPrompt = "Respond with 'OK' if you can read this message."
        $response = Invoke-AIRequest -Prompt $testPrompt -MaxTokens 10

        return @{
            Success  = $true
            Provider = $script:CurrentProvider.Name
            Response = $response
        }
    }
    catch {
        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}
#endregion

#region Core AI Functions
function Invoke-AIRequest {
    <#
    .SYNOPSIS
        Sends a request to the configured AI provider
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter(Mandatory = $false)]
        [string]$SystemPrompt,

        [Parameter(Mandatory = $false)]
        [int]$MaxTokens = 4096,

        [Parameter(Mandatory = $false)]
        [double]$Temperature = 0.7
    )

    if (-not $script:CurrentProvider -or -not $script:APIKey) {
        throw "AI provider not configured. Call Set-AIProvider first."
    }

    $headers = $script:CurrentProvider.Headers.Clone()

    # Determine the endpoint URL - Gemini uses URL parameter auth
    $endpoint = $script:CurrentProvider.Endpoint
    if ($script:CurrentProvider.AuthHeader -eq "url-param") {
        # Gemini uses API key in URL parameter
        $endpoint = "$($script:CurrentProvider.Endpoint)?key=$script:APIKey"
    } else {
        # Other providers use header-based auth
        $headers[$script:CurrentProvider.AuthHeader] = "$($script:CurrentProvider.AuthPrefix)$script:APIKey"
    }

    $body = switch -Wildcard ($script:CurrentProvider.Model) {
        "gpt-*" {
            @{
                model                = $script:CurrentProvider.Model
                messages             = @(
                    if ($SystemPrompt) {
                        @{ role = "system"; content = $SystemPrompt }
                    }
                    @{ role = "user"; content = $Prompt }
                )
                max_completion_tokens = $MaxTokens
                temperature          = $Temperature
            }
        }
        "claude-*" {
            @{
                model       = $script:CurrentProvider.Model
                max_tokens  = $MaxTokens
                system      = if ($SystemPrompt) { $SystemPrompt } else { "You are an expert Microsoft 365 migration consultant." }
                messages    = @(
                    @{ role = "user"; content = $Prompt }
                )
            }
        }
        "gemini-*" {
            @{
                contents = @(
                    @{
                        parts = @(
                            @{ text = if ($SystemPrompt) { "$SystemPrompt`n`n$Prompt" } else { $Prompt } }
                        )
                    }
                )
                generationConfig = @{
                    maxOutputTokens = $MaxTokens
                    temperature     = $Temperature
                }
            }
        }
    }

    try {
        # Ensure model name is clean - convert to string explicitly and assign
        $cleanModelName = $script:CurrentProvider.Model -replace '[`''""]', ''
        if ($body.ContainsKey('model')) {
            $body['model'] = $cleanModelName
        }

        $jsonBody = $body | ConvertTo-Json -Depth 10

        $response = Invoke-RestMethod -Uri $endpoint `
            -Method Post `
            -Headers $headers `
            -Body $jsonBody `
            -ContentType "application/json"

        # Extract response text based on provider with proper null checks
        $responseText = switch -Wildcard ($script:CurrentProvider.Model) {
            "gpt-*" {
                if ($response.choices -and $response.choices.Count -gt 0 -and $response.choices[0].message) {
                    $response.choices[0].message.content
                } else {
                    throw "Invalid GPT response structure: $($response | ConvertTo-Json -Depth 3 -Compress)"
                }
            }
            "claude-*" {
                if ($response.content -and $response.content.Count -gt 0) {
                    $response.content[0].text
                } else {
                    throw "Invalid Claude response structure: $($response | ConvertTo-Json -Depth 3 -Compress)"
                }
            }
            "gemini-*" {
                if ($response.candidates -and $response.candidates.Count -gt 0 -and
                    $response.candidates[0].content -and
                    $response.candidates[0].content.parts -and
                    $response.candidates[0].content.parts.Count -gt 0) {
                    $response.candidates[0].content.parts[0].text
                } elseif ($response.error) {
                    throw "Gemini API error: $($response.error.message)"
                } else {
                    throw "Invalid Gemini response structure: $($response | ConvertTo-Json -Depth 3 -Compress)"
                }
            }
        }

        return $responseText
    }
    catch {
        Write-Log -Message "AI request failed: $_" -Level Error
        throw
    }
}
#endregion

#region Analysis Functions
function Get-AIGotchaAnalysis {
    <#
    .SYNOPSIS
        Uses AI to provide detailed analysis of discovered gotchas
    .DESCRIPTION
        Sends three sequential API calls to avoid token limits:
        - Batch 1 (8K): Executive Summary, Critical Path, Remediation Plans
        - Batch 2 (6K): Hidden Risks, Waves, Coexistence, Resources, Go/No-Go
        - Batch 3 (10K): Complete Migration Playbook with PowerShell scripts
        Returns markdown for report rendering and separate playbook for script extraction.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Gotchas,

        [Parameter(Mandatory = $true)]
        [hashtable]$TenantContext
    )

    Write-Log -Message "Generating AI analysis of gotchas (batched mode)..." -Level Info

    $systemPrompt = @"
ROLE
You are a senior enterprise cloud migration architect with deep expertise in identity, messaging, collaboration, automation, and security systems.

OBJECTIVE
Analyze the provided tenant data and discovered issues to produce a decisive, actionable migration assessment.

RULES
- Do NOT ask follow-up questions.
- Use ONLY the provided data.
- If data is missing, list assumptions clearly.
- Prioritize highest-risk items first.
- Provide prescriptive commands and complete working scripts when needed.
- Avoid generic advice.
- Do NOT generate ASCII art diagrams - use clean numbered lists for sequencing.
- Format output as structured markdown with clear sections, tables, and PowerShell code blocks.
"@

    # Group gotchas by severity for better analysis
    $criticalGotchas = $Gotchas | Where-Object { $_.Severity -eq "Critical" }
    $highGotchas = $Gotchas | Where-Object { $_.Severity -eq "High" }
    $mediumGotchas = $Gotchas | Where-Object { $_.Severity -eq "Medium" }
    $lowGotchas = $Gotchas | Where-Object { $_.Severity -eq "Low" }

    $formatGotcha = {
        param($g)
        "- **[$($g.Category)] $($g.Title)** (Severity: $($g.Severity))`n  - Issue: $($g.Description)`n  - Affected Count: $($g.AffectedCount)`n  - Initial Recommendation: $($g.Recommendation)"
    }

    $criticalSection = if ($criticalGotchas) { ($criticalGotchas | ForEach-Object { & $formatGotcha $_ }) -join "`n" } else { "None identified" }
    $highSection = if ($highGotchas) { ($highGotchas | ForEach-Object { & $formatGotcha $_ }) -join "`n" } else { "None identified" }
    $mediumSection = if ($mediumGotchas) { ($mediumGotchas | ForEach-Object { & $formatGotcha $_ }) -join "`n" } else { "None identified" }
    $lowSection = if ($lowGotchas) { ($lowGotchas | ForEach-Object { & $formatGotcha $_ }) -join "`n" } else { "None identified" }

    # Common context block for all batches
    $tenantContextBlock = @"
## Source Tenant Context

| Metric | Value | Migration Implication |
|--------|-------|----------------------|
| Licensed Users | $($TenantContext.UserCount) | $(if($TenantContext.UserCount -gt 500){"Large user base - consider batched migration"}else{"Manageable user count"}) |
| Shared Mailboxes | $($TenantContext.SharedMailboxCount) | Require permission mapping in target |
| Guest Users | $($TenantContext.GuestCount) | $(if($TenantContext.GuestCount -gt 50){"Guest reinvitation automation needed"}else{"Manual guest handling feasible"}) (not migrated, need reinvitation) |
| Hybrid Identity | $($TenantContext.HybridEnabled) | $(if($TenantContext.HybridEnabled){"Requires AAD Connect migration planning"}else{"Cloud-only simplifies identity migration"}) |
| Exchange Mailboxes | $($TenantContext.MailboxCount) | $(if($TenantContext.MailboxCount -gt 500){"Extended mailbox migration window needed"}else{"Standard migration timeline"}) |
| SharePoint Sites | $($TenantContext.SiteCount) | Site-by-site migration required |
| Teams | $($TenantContext.TeamCount) | $(if($TenantContext.TeamCount -gt 100){"Significant Teams estate"}else{"Moderate Teams footprint"}) |
| Synced Users | $($TenantContext.SyncedUserCount) | ImmutableId strategy required |
| Hybrid Devices | $($TenantContext.HybridDeviceCount) | Device registration cutover needed |
| Power Platform Environments | $($TenantContext.D365Environments) | $(if($TenantContext.D365Environments -gt 0){"Dedicated Power Platform migration workstream required"}else{"No Power Platform environments detected"}) |
| Power Apps | $($TenantContext.D365Apps) | $(if($TenantContext.D365Apps -gt 0){"Apps require export and connection recreation in target"}else{"No Power Apps detected"}) |
| Power Automate Flows | $($TenantContext.D365Flows) | $(if($TenantContext.D365Flows -gt 0){"Flows require export and connection recreation in target"}else{"No flows detected"}) |
| Power BI Workspaces | $($TenantContext.PowerBIWorkspaces) | $(if($TenantContext.PowerBIWorkspaces -gt 0){"Reports and datasets must be republished to target tenant workspaces"}else{"No Power BI workspaces detected"}) |
| Power BI Gateways | $($TenantContext.PowerBIGateways) | $(if($TenantContext.PowerBIGateways -gt 0){"On-premises data gateways must be reinstalled and reconfigured for target tenant"}else{"No on-premises gateways to migrate"}) |

## Discovered Issues Summary
- **Critical Issues**: $($criticalGotchas.Count)
- **High Priority Issues**: $($highGotchas.Count)
- **Medium Priority Issues**: $($mediumGotchas.Count)
- **Low Priority Issues**: $($lowGotchas.Count)
- **Total Issues**: $($Gotchas.Count)

## CRITICAL ISSUES (Migration Blockers)
$criticalSection

## HIGH PRIORITY ISSUES (Pre-Migration Resolution Required)
$highSection

## MEDIUM PRIORITY ISSUES (Address During Migration)
$mediumSection

## LOW PRIORITY ISSUES (Post-Migration Tasks)
$lowSection
"@

    #region BATCH 1: Executive Summary, Critical Path, Remediation Plans
    $prompt1 = @"
$tenantContextBlock

---

Provide the following sections in markdown:

## 1. EXECUTIVE SUMMARY
3-5 sentence assessment of migration complexity, key risks, and readiness (RED/AMBER/GREEN).

## 2. CRITICAL PATH ANALYSIS
Identify which items block others and must be completed in sequence.

## 3. MIGRATION SEQUENCING
For critical and high severity items, provide the required order of operations as a numbered list. State which items must be completed before others and why.

## 4. DETAILED REMEDIATION PLANS
For each CRITICAL and HIGH severity item provide:
- **Root Cause**: Why this blocks migration (2-3 sentences)
- **Step-by-Step Procedure**: Numbered steps with exact PowerShell commands in code blocks
- **Validation Commands**: PowerShell to verify success
- **Estimated Effort**: Time estimate
- **Rollback Commands**: PowerShell to undo if needed
"@
    #endregion

    #region BATCH 2: Hidden Risks, Waves, Go/No-Go
    $prompt2 = @"
$tenantContextBlock

---

Provide the following sections in markdown:

## 5. HIDDEN RISKS
Additional risks not detected but commonly encountered with this tenant configuration.

## 6. MIGRATION WAVE RECOMMENDATION
How to batch users/workloads into waves based on dependencies and risk. Provide specific wave groupings.

## 7. COEXISTENCE REQUIREMENTS
What coexistence capabilities are needed during migration and for how long.

## 8. RESOURCE & SKILL MATRIX
| Role | Required Skills | FTE Estimate | Duration |
|------|----------------|--------------|----------|
(Complete with realistic estimates)

## 9. TOOL RECOMMENDATIONS
Specific migration tools for each workload with licensing considerations.

## 10. GO/NO-GO CRITERIA
Measurable criteria that must be met before migration proceeds. Format as checklist.
"@
    #endregion

    #region BATCH 3: Playbook with Scripts
    $prompt3 = @"
$tenantContextBlock

---

## 11. MIGRATION PLAYBOOK / RUNBOOK

Provide a complete migration runbook. All scripts must be complete and ready to run.

### 11.1 Pre-Migration Scripts

```powershell
# Script 1: Export Users and Groups
# Include: Module imports, variables, error handling, progress output
```

```powershell
# Script 2: Export Permissions and Delegations
```

```powershell
# Script 3: Pre-Migration Health Check
```

### 11.2 Migration Execution Scripts

```powershell
# Script 4: Identity Preparation (ImmutableId mapping, UPN changes)
```

```powershell
# Script 5: Exchange Migration Batch Commands
```

```powershell
# Script 6: SharePoint/OneDrive Migration Setup
```

```powershell
# Script 7: Teams Migration Procedures
```

### 11.3 Post-Migration Validation Scripts

```powershell
# Script 8: User Access Verification
```

```powershell
# Script 9: Mail Flow Testing
```

```powershell
# Script 10: Permission Validation
```

### 11.4 Rollback Scripts

```powershell
# Script 11: Emergency Rollback Procedures
```

SCRIPT REQUIREMENTS:
- Module imports at top
- Variable declarations section
- Try/Catch error handling
- Write-Progress or Write-Host for status
- Comments explaining each section
- No placeholder values - use realistic examples
"@
    #endregion

    try {
        # Execute Batch 1
        Write-Log -Message "AI Analysis Batch 1/3: Executive Summary & Remediation Plans..." -Level Info
        $batch1 = Invoke-AIRequest -Prompt $prompt1 -SystemPrompt $systemPrompt -MaxTokens 8000 -Temperature 0.5
        Write-Log -Message "Batch 1 complete." -Level Info

        Start-Sleep -Milliseconds 1000

        # Execute Batch 2
        Write-Log -Message "AI Analysis Batch 2/3: Planning & Resources..." -Level Info
        $batch2 = Invoke-AIRequest -Prompt $prompt2 -SystemPrompt $systemPrompt -MaxTokens 6000 -Temperature 0.5
        Write-Log -Message "Batch 2 complete." -Level Info

        Start-Sleep -Milliseconds 1000

        # Execute Batch 3 (Playbook - kept separate for script extraction)
        Write-Log -Message "AI Analysis Batch 3/3: Migration Playbook & Scripts..." -Level Info
        $batch3 = Invoke-AIRequest -Prompt $prompt3 -SystemPrompt $systemPrompt -MaxTokens 10000 -Temperature 0.5
        Write-Log -Message "Batch 3 complete." -Level Info

        # Combine analysis (batches 1 & 2) using direct string concatenation - playbook kept separate
        # Use [string] cast to ensure AI responses are strings (not arrays) before concatenation
        $combinedAnalysis = [string]$batch1 + "`n`n---`n`n" + [string]$batch2 + @"

---

## 11. MIGRATION PLAYBOOK / RUNBOOK

The complete migration playbook with ready-to-execute PowerShell scripts has been generated and saved to separate files in the **Scripts/** folder. Review and customize these scripts before execution.
"@

        return @{
            Success   = $true
            Analysis  = $combinedAnalysis
            Playbook  = $batch3
            Provider  = $script:CurrentProvider.Name
            Timestamp = Get-Date
            BatchInfo = @{
                TotalBatches = 3
                Batch1Tokens = 8000
                Batch2Tokens = 6000
                Batch3Tokens = 10000
            }
            GotchaCount = @{
                Critical = $criticalGotchas.Count
                High     = $highGotchas.Count
                Medium   = $mediumGotchas.Count
                Low      = $lowGotchas.Count
            }
        }
    }
    catch {
        Write-Log -Message "AI gotcha analysis failed: $_" -Level Error
        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}

function Get-AIExecutiveSummary {
    <#
    .SYNOPSIS
        Generates an executive-level summary using AI
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $true)]
        $AnalysisResults,

        [Parameter(Mandatory = $true)]
        $ComplexityScore
    )

    Write-Log -Message "Generating AI executive summary..." -Level Info

    $systemPrompt = @"
You are a senior IT Strategy Consultant and former CIO presenting to C-level executives and board members. You specialize in digital transformation and have presented migration assessments to Fortune 500 leadership teams.

Your communication style:
- Business outcome focused, not technology focused
- Clear ROI and risk quantification
- Actionable strategic recommendations
- Honest assessment of challenges without creating unnecessary alarm
- Professional, confident, and authoritative tone

IMPORTANT GUIDELINES:
1. Translate technical issues into business impact (productivity, security, compliance, cost)
2. Frame risks in terms executives understand (revenue impact, regulatory exposure, competitive disadvantage)
3. Provide clear decision points with options and tradeoffs
4. Include realistic timeline ranges, not specific dates
5. Quantify where possible (e.g., "affects 40% of workforce", not "many users")
6. Avoid jargon - if technical terms are necessary, explain them briefly
7. Lead with the most important information
8. Be direct about what needs executive decision vs. what IT can handle

Format as a polished executive briefing suitable for board presentation with professional markdown.
"@

    # Calculate derived metrics for business context with null safety
    $totalUsers = if ($CollectedData.EntraID.Users.Analysis) { $CollectedData.EntraID.Users.Analysis.TotalUsers } else { 0 }
    $licensedUsers = if ($CollectedData.EntraID.Users.Analysis) { $CollectedData.EntraID.Users.Analysis.LicensedUsers } else { 0 }
    $guestUsers = if ($CollectedData.EntraID.Users.Analysis) { $CollectedData.EntraID.Users.Analysis.GuestUsers } else { 0 }
    $totalMailboxes = if ($CollectedData.Exchange.Mailboxes.Analysis) { $CollectedData.Exchange.Mailboxes.Analysis.TotalMailboxes } else { 0 }
    $totalSites = if ($CollectedData.SharePoint.Sites.Analysis) { $CollectedData.SharePoint.Sites.Analysis.SharePointSites } else { 0 }
    $totalTeams = if ($CollectedData.Teams.Teams.Analysis) { $CollectedData.Teams.Teams.Analysis.TotalTeams } else { 0 }
    $storageGB = if ($CollectedData.SharePoint.Sites.Analysis) { $CollectedData.SharePoint.Sites.Analysis.TotalStorageUsedGB } else { 0 }
    $hybridEnabled = if ($CollectedData.HybridIdentity.AADConnect.Configuration) { $CollectedData.HybridIdentity.AADConnect.Configuration.OnPremisesSyncEnabled } else { $false }
    $syncedUsers = if ($CollectedData.EntraID.Users.Analysis) { $CollectedData.EntraID.Users.Analysis.SyncedUsers } else { 0 }
    $hybridDevices = if ($CollectedData.EntraID.Devices.Analysis) { $CollectedData.EntraID.Devices.Analysis.HybridJoined } else { 0 }

    # Determine organization size category based on LICENSED users (not total directory)
    $orgSize = switch ($licensedUsers) {
        { $_ -gt 5000 } { "Enterprise (5,000+ licensed users)"; break }
        { $_ -gt 1000 } { "Large Organization (1,000-5,000 licensed users)"; break }
        { $_ -gt 250 }  { "Mid-Size Organization (250-1,000 licensed users)"; break }
        default         { "Small Organization (under 250 licensed users)" }
    }

    # Get shared mailbox count with null safety
    $sharedMailboxes = if ($CollectedData.Exchange.Mailboxes.Mailboxes) { @($CollectedData.Exchange.Mailboxes.Mailboxes | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" }).Count } else { 0 }

    # Calculate risk metrics with null safety
    $criticalCount = if ($AnalysisResults -and $AnalysisResults.BySeverity -and $AnalysisResults.BySeverity.Critical) { $AnalysisResults.BySeverity.Critical.Count } else { 0 }
    $highCount = if ($AnalysisResults -and $AnalysisResults.BySeverity -and $AnalysisResults.BySeverity.High) { $AnalysisResults.BySeverity.High.Count } else { 0 }
    $mediumCount = if ($AnalysisResults -and $AnalysisResults.BySeverity -and $AnalysisResults.BySeverity.Medium) { $AnalysisResults.BySeverity.Medium.Count } else { 0 }
    $lowCount = if ($AnalysisResults -and $AnalysisResults.BySeverity -and $AnalysisResults.BySeverity.Low) { $AnalysisResults.BySeverity.Low.Count } else { 0 }
    $totalIssues = if ($AnalysisResults -and $AnalysisResults.RulesTriggered) { $AnalysisResults.RulesTriggered } else { 0 }
    $riskLevel = if ($AnalysisResults -and $AnalysisResults.RiskLevel) { $AnalysisResults.RiskLevel } else { "Unknown" }

    # Build critical issues list safely
    $criticalIssuesList = if ($AnalysisResults -and $AnalysisResults.BySeverity -and $AnalysisResults.BySeverity.Critical) {
        ($AnalysisResults.BySeverity.Critical | ForEach-Object { "- **$($_.Name)**: $($_.Description)" }) -join "`n"
    } else { "No critical issues identified" }

    # Build category breakdown safely
    $categoryBreakdown = if ($AnalysisResults -and $AnalysisResults.ByCategory -and $AnalysisResults.ByCategory.Keys) {
        ($AnalysisResults.ByCategory.Keys | ForEach-Object { "- **$_**: $($AnalysisResults.ByCategory[$_].Count) findings" }) -join "`n"
    } else { "No category breakdown available" }

    # Infrastructure metrics with null safety
    $federatedDomains = if ($CollectedData.HybridIdentity.Federation.Analysis) { $CollectedData.HybridIdentity.Federation.Analysis.FederatedDomains } else { 0 }
    $litigationHoldCount = if ($CollectedData.Exchange.Mailboxes.Analysis) { $CollectedData.Exchange.Mailboxes.Analysis.LitigationHold } else { 0 }

    $prompt = @"
# Microsoft 365 Tenant Migration - Executive Briefing

## Organization Profile

| Dimension | Current State | Migration Implication |
|-----------|--------------|----------------------|
| Organization Size | $orgSize | $licensedUsers licensed users to migrate |
| Shared Mailboxes | $sharedMailboxes shared mailboxes | Require permission mapping in target |
| External Collaboration | $guestUsers guest users | Not migrated - require reinvitation in target tenant |
| Email Infrastructure | $totalMailboxes mailboxes total | Core productivity system - zero downtime required |
| Collaboration Footprint | $totalTeams Teams environments | Team continuity critical |
| Digital Content | $totalSites SharePoint sites ($storageGB GB) | Significant content migration effort |
| Identity Architecture | $(if($hybridEnabled){"Hybrid (On-premises AD synchronized)"}else{"Cloud-native"}) | $(if($hybridEnabled){"Complex identity transition required"}else{"Simplified identity migration"}) |
| Hybrid Workforce | $syncedUsers synced users, $hybridDevices hybrid devices | $(if($hybridDevices -gt 0){"Device re-registration impact on users"}else{"Minimal device impact"}) |

## Assessment Results Summary

### Overall Migration Readiness
- **Complexity Rating**: $($ComplexityScore.ComplexityLevel) ($($ComplexityScore.TotalScore)/100)
- **Risk Classification**: $riskLevel
- **Assessment Status**: $(if($criticalCount -gt 0){"REQUIRES REMEDIATION - Critical blockers identified"}elseif($highCount -gt 5){"CONDITIONAL PROCEED - Significant preparation needed"}else{"READY FOR PLANNING - Standard migration complexity"})

### Issue Distribution
| Priority | Count | Business Impact |
|----------|-------|-----------------|
| Critical (Blockers) | $criticalCount | Migration cannot proceed until resolved |
| High Priority | $highCount | Must be addressed before migration start |
| Medium Priority | $mediumCount | Address during migration execution |
| Low Priority | $lowCount | Post-migration optimization |

### Critical Issues Requiring Executive Attention
$criticalIssuesList

### Areas of Concern by Business Function
$categoryBreakdown

### Infrastructure Considerations
- **On-Premises Integration**: $(if($hybridEnabled){"Active directory synchronization in place - requires coordinated cutover"}else{"No on-premises dependencies"})
- **Federation Status**: $(if($federatedDomains -gt 0){"$federatedDomains federated domain(s) - authentication transition planning required"}else{"No federation dependencies"})
- **Compliance Holdings**: $(if($litigationHoldCount -gt 0){"$litigationHoldCount mailboxes under legal hold - legal coordination required"}else{"No active legal holds"})

---

Please generate a comprehensive executive summary addressing:

## 1. EXECUTIVE OVERVIEW (2-3 paragraphs)
Summarize the migration assessment in plain business language. What are we migrating, what's the current state, and what should leadership know?

## 2. BUSINESS IMPACT ASSESSMENT

### Operational Continuity
How will day-to-day business operations be affected? What's the expected user experience during migration?

### Productivity Impact
Quantify potential productivity impact (e.g., "Users may experience X during Y period")

### Partner & Customer Impact
Any external-facing implications?

## 3. RISK ASSESSMENT FOR LEADERSHIP

Present the top 5 risks in business terms:
| Risk | Business Impact | Likelihood | Mitigation Available |
|------|-----------------|------------|---------------------|

## 4. STRATEGIC RECOMMENDATIONS

Provide 5 prioritized recommendations for leadership:
1. [Highest Priority]
2. ...

## 5. DECISION POINTS FOR LEADERSHIP

What decisions does the executive team need to make?
- Decision 1: [Description] - Options: A, B, C
- Timeline for decision: [When]

## 6. INVESTMENT CONSIDERATIONS

Without providing specific dollar amounts, outline:
- Categories of investment required (licensing, tools, professional services, internal resources)
- Cost optimization opportunities
- ROI factors to consider

## 7. TIMELINE OVERVIEW

Provide a high-level timeline with phases:
| Phase | Duration Range | Key Milestones |
|-------|---------------|----------------|

## 8. SUCCESS METRICS

How will we measure successful migration?
- Business continuity metrics
- User satisfaction indicators
- Technical validation criteria

## 9. NEXT STEPS

Immediate actions following this briefing (next 30 days)
"@

    try {
        $summary = Invoke-AIRequest -Prompt $prompt -SystemPrompt $systemPrompt -MaxTokens 10000 -Temperature 0.7

        return @{
            Success   = $true
            Summary   = $summary
            Provider  = $script:CurrentProvider.Name
            Timestamp = Get-Date
            Metrics   = @{
                TotalUsers    = $totalUsers
                OrgSize       = $orgSize
                RiskLevel     = $AnalysisResults.RiskLevel
                Complexity    = $ComplexityScore.ComplexityLevel
                CriticalCount = $criticalCount
            }
        }
    }
    catch {
        Write-Log -Message "AI executive summary generation failed: $_" -Level Error
        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}

function Get-AIRemediationPlan {
    <#
    .SYNOPSIS
        Generates detailed remediation plan using AI
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Gotcha
    )

    $systemPrompt = @"
You are a Microsoft 365 migration technical specialist.
Provide detailed, step-by-step remediation guidance.
Include PowerShell commands where applicable.
Format using markdown with clear sections and code blocks.
"@

    $prompt = @"
## Issue Details
- **Title**: $($Gotcha.Title)
- **Category**: $($Gotcha.Category)
- **Severity**: $($Gotcha.Severity)
- **Description**: $($Gotcha.Description)
- **Affected Count**: $($Gotcha.AffectedCount)

Please provide:
1. **Technical Background**: Explain why this is an issue
2. **Impact Assessment**: What happens if not addressed
3. **Pre-requisites**: What's needed before remediation
4. **Step-by-Step Remediation**:
   - Include specific PowerShell commands where applicable
   - Note any required permissions
   - Include verification steps
5. **Rollback Plan**: How to undo if needed
6. **Success Validation**: How to confirm resolution
7. **Estimated Effort**: Time and resources required
"@

    try {
        $plan = Invoke-AIRequest -Prompt $prompt -SystemPrompt $systemPrompt -MaxTokens 4000

        return @{
            Success = $true
            Plan    = $plan
            GotchaId = $Gotcha.Id
        }
    }
    catch {
        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}

function Get-AITenantComparison {
    <#
    .SYNOPSIS
        Generates AI-powered comparison analysis between source findings and target recommendations
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$SourceData,

        [Parameter(Mandatory = $false)]
        [hashtable]$TargetRequirements
    )

    $systemPrompt = @"
You are a Microsoft 365 migration architect performing gap analysis.
Compare source tenant configuration against migration best practices.
Identify gaps, risks, and required changes.
"@

    $prompt = @"
## Source Tenant Analysis

### Identity Configuration
- Users: $($SourceData.EntraID.Users.Analysis.TotalUsers)
- Synced: $($SourceData.EntraID.Users.Analysis.SyncedUsers)
- Cloud-only: $($SourceData.EntraID.Users.Analysis.CloudOnlyUsers)
- Groups: $($SourceData.EntraID.Groups.Analysis.TotalGroups)
- Devices: $($SourceData.EntraID.Devices.Analysis.TotalDevices)
- Hybrid Devices: $($SourceData.EntraID.Devices.Analysis.HybridJoined)

### Workload Configuration
- Mailboxes: $($SourceData.Exchange.Mailboxes.Analysis.TotalMailboxes)
- SharePoint Sites: $($SourceData.SharePoint.Sites.Analysis.SharePointSites)
- Teams: $($SourceData.Teams.Teams.Analysis.TotalTeams)

### Security & Compliance
- Conditional Access Policies: $($SourceData.EntraID.ConditionalAccess.Analysis.TotalPolicies)
- DLP Policies: $($SourceData.Security.DLPPolicies.Analysis.TotalPolicies)
- Retention Policies: $($SourceData.Security.RetentionPolicies.Analysis.TotalPolicies)
- Sensitivity Labels: $($SourceData.Security.SensitivityLabels.Analysis.TotalLabels)

Please provide:
1. **Configuration Gaps**: What needs to change for target tenant
2. **Compatibility Assessment**: What migrates cleanly vs. requires recreation
3. **Dependency Mapping**: What must be migrated in what order
4. **Coexistence Requirements**: What's needed during migration
5. **Post-Migration Cleanup**: What to remove/decommission after migration
"@

    try {
        $comparison = Invoke-AIRequest -Prompt $prompt -SystemPrompt $systemPrompt -MaxTokens 5000

        return @{
            Success    = $true
            Comparison = $comparison
        }
    }
    catch {
        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}
#endregion

#region Batch Processing
function Invoke-AIBatchAnalysis {
    <#
    .SYNOPSIS
        Processes multiple gotchas with AI analysis in batches
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Gotchas,

        [Parameter(Mandatory = $false)]
        [int]$BatchSize = 5
    )

    $results = @()
    $batches = [math]::Ceiling($Gotchas.Count / $BatchSize)

    for ($i = 0; $i -lt $batches; $i++) {
        $startIndex = $i * $BatchSize
        $batch = $Gotchas | Select-Object -Skip $startIndex -First $BatchSize

        Write-Log -Message "Processing batch $($i + 1) of $batches..." -Level Info

        foreach ($gotcha in $batch) {
            try {
                $plan = Get-AIRemediationPlan -Gotcha $gotcha
                $results += @{
                    GotchaId = $gotcha.Id
                    Title    = $gotcha.Title
                    Plan     = $plan.Plan
                    Success  = $plan.Success
                }
            }
            catch {
                $results += @{
                    GotchaId = $gotcha.Id
                    Title    = $gotcha.Title
                    Success  = $false
                    Error    = $_.Exception.Message
                }
            }

            # Rate limiting
            Start-Sleep -Milliseconds 500
        }
    }

    return $results
}
#endregion

#region Script Extraction
function Save-MigrationPlaybook {
    <#
    .SYNOPSIS
        Extracts PowerShell scripts from AI-generated playbook and saves to individual files
    .DESCRIPTION
        Parses the markdown playbook content, extracts PowerShell code blocks,
        and saves each script to a separate .ps1 file in the specified output directory.
        Injects actual tenant data (TenantId, UserCount, etc.) to replace placeholders.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PlaybookContent,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $false)]
        [hashtable]$CollectedData
    )

    Write-Log -Message "Extracting migration scripts from playbook..." -Level Info

    # Create Scripts subdirectory
    $scriptsPath = Join-Path $OutputPath "Scripts"
    if (-not (Test-Path $scriptsPath)) {
        New-Item -ItemType Directory -Path $scriptsPath -Force | Out-Null
    }

    # Script mapping - associates script numbers/names with output filenames
    $scriptMapping = @{
        "Script 1"  = "Export-UsersAndGroups.ps1"
        "Script 2"  = "Export-PermissionsAndDelegations.ps1"
        "Script 3"  = "Pre-MigrationHealthCheck.ps1"
        "Script 4"  = "Prepare-Identity.ps1"
        "Script 5"  = "Start-ExchangeMigration.ps1"
        "Script 6"  = "Start-SharePointMigration.ps1"
        "Script 7"  = "Start-TeamsMigration.ps1"
        "Script 8"  = "Test-UserAccess.ps1"
        "Script 9"  = "Test-MailFlow.ps1"
        "Script 10" = "Test-Permissions.ps1"
        "Script 11" = "Invoke-EmergencyRollback.ps1"
        "Export Users" = "Export-UsersAndGroups.ps1"
        "Export Permissions" = "Export-PermissionsAndDelegations.ps1"
        "Health Check" = "Pre-MigrationHealthCheck.ps1"
        "Pre-Migration Health" = "Pre-MigrationHealthCheck.ps1"
        "Identity Preparation" = "Prepare-Identity.ps1"
        "ImmutableId" = "Prepare-Identity.ps1"
        "Exchange Migration" = "Start-ExchangeMigration.ps1"
        "SharePoint" = "Start-SharePointMigration.ps1"
        "OneDrive" = "Start-SharePointMigration.ps1"
        "Teams Migration" = "Start-TeamsMigration.ps1"
        "User Access" = "Test-UserAccess.ps1"
        "Mail Flow" = "Test-MailFlow.ps1"
        "Permission Validation" = "Test-Permissions.ps1"
        "Rollback" = "Invoke-EmergencyRollback.ps1"
        "Emergency" = "Invoke-EmergencyRollback.ps1"
    }

    # Extract all PowerShell code blocks
    $codeBlockPattern = '```powershell\s*([\s\S]*?)```'
    $matches = [regex]::Matches($PlaybookContent, $codeBlockPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    $savedScripts = @()
    $scriptIndex = 1

    foreach ($match in $matches) {
        $scriptContent = $match.Groups[1].Value.Trim()

        # Skip empty or placeholder scripts
        if ([string]::IsNullOrWhiteSpace($scriptContent) -or $scriptContent.Length -lt 50) {
            continue
        }

        # Try to determine the script name from content or context
        $fileName = $null

        # Look for script name hints in the content or preceding text
        $precedingText = ""
        $matchStart = $match.Index
        if ($matchStart -gt 200) {
            $precedingText = $PlaybookContent.Substring($matchStart - 200, 200)
        } elseif ($matchStart -gt 0) {
            $precedingText = $PlaybookContent.Substring(0, $matchStart)
        }

        # Check for script number or name in preceding text
        foreach ($key in $scriptMapping.Keys) {
            if ($precedingText -match [regex]::Escape($key) -or $scriptContent -match "# $key") {
                $fileName = $scriptMapping[$key]
                break
            }
        }

        # If no match found, use generic naming
        if (-not $fileName) {
            $fileName = "MigrationScript-$scriptIndex.ps1"
        }

        $filePath = Join-Path $scriptsPath $fileName

        # Inject tenant data into script (replace placeholder values)
        $injectedContent = $scriptContent
        if ($CollectedData) {
            $tenantId = if ($CollectedData.Metadata.TenantId) { $CollectedData.Metadata.TenantId } elseif ($CollectedData.TenantInfo.TenantId) { $CollectedData.TenantInfo.TenantId } else { "" }
            $tenantName = if ($CollectedData.Metadata.TenantName) { $CollectedData.Metadata.TenantName } elseif ($CollectedData.TenantInfo.DisplayName) { $CollectedData.TenantInfo.DisplayName } else { "YourTenant" }
            $licensedUsers = if ($CollectedData.EntraID.Users.Analysis.LicensedUsers) { $CollectedData.EntraID.Users.Analysis.LicensedUsers } else { 0 }
            $adminUrl = if ($CollectedData.SharePoint.TenantConfig.AdminUrl) { $CollectedData.SharePoint.TenantConfig.AdminUrl } else { "https://$($tenantName)-admin.sharepoint.com" }

            # Replace common placeholders
            $injectedContent = $injectedContent -replace '\$TenantId\s*=\s*[''"][^''"]*[''"]', "`$TenantId = '$tenantId'"
            $injectedContent = $injectedContent -replace '\$TenantId\s*=\s*\$null', "`$TenantId = '$tenantId'"
            $injectedContent = $injectedContent -replace '\$tenantId\s*=\s*[''"][^''"]*[''"]', "`$tenantId = '$tenantId'"
            $injectedContent = $injectedContent -replace '\$tenantId\s*=\s*\$null', "`$tenantId = '$tenantId'"
            $injectedContent = $injectedContent -replace '"[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}".*#.*tenant', "'$tenantId' # Your tenant ID"
            $injectedContent = $injectedContent -replace 'https://.*-admin\.sharepoint\.com', $adminUrl
            $injectedContent = $injectedContent -replace 'https://.*\.sharepoint\.com', "https://$tenantName.sharepoint.com"
            $injectedContent = $injectedContent -replace '\$UserCount\s*=\s*\d+', "`$UserCount = $licensedUsers"
        }

        # Add header comment to script
        $header = @"
#Requires -Version 7.0
<#
.SYNOPSIS
    M365 Migration Script - $($fileName -replace '\.ps1$', '')
.DESCRIPTION
    Auto-generated migration script from AI analysis.
    Review and customize before execution.
.NOTES
    Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    Source: M365 Tenant Discovery Tool - AI Migration Playbook
    TenantId: $(if ($CollectedData -and $CollectedData.Metadata.TenantId) { $CollectedData.Metadata.TenantId } else { "Not specified" })
#>

"@

        $fullContent = $header + $injectedContent

        # Save script (avoid duplicates by checking if file exists with same content)
        if (-not (Test-Path $filePath) -or (Get-Content $filePath -Raw) -ne $fullContent) {
            $fullContent | Out-File -FilePath $filePath -Encoding UTF8 -Force
            $savedScripts += $fileName
            Write-Log -Message "Saved script: $fileName" -Level Info
        }

        $scriptIndex++
    }

    # Also save the complete playbook as markdown for reference
    $playbookPath = Join-Path $scriptsPath "Migration-Playbook.md"
    $PlaybookContent | Out-File -FilePath $playbookPath -Encoding UTF8 -Force
    Write-Log -Message "Saved complete playbook: Migration-Playbook.md" -Level Info

    return @{
        Success = $true
        ScriptsPath = $scriptsPath
        SavedScripts = $savedScripts
        PlaybookPath = $playbookPath
    }
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Set-AIProvider',
    'Get-AIProvider',
    'Test-AIConnection',
    'Invoke-AIRequest',
    'Get-AIGotchaAnalysis',
    'Get-AIExecutiveSummary',
    'Get-AIRemediationPlan',
    'Get-AITenantComparison',
    'Invoke-AIBatchAnalysis',
    'Save-MigrationPlaybook'
)
