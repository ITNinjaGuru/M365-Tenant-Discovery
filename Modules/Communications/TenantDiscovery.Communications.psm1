#Requires -Version 7.0
<#
.SYNOPSIS
    Communication Plan Generator for M365 Migration
.DESCRIPTION
    Generates comprehensive communication packages for M365 tenant migrations including:
    - Branded email templates for all-company announcements
    - SharePoint Online news post content
    - Visual migration timelines
    - AI-powered content generation for non-technical audiences
.NOTES
    Author: M365 Migration Team
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

# Import AI Integration module
$aiPath = Join-Path $PSScriptRoot ".." ".." "Analysis" "AIIntegration.psm1"
if (Test-Path $aiPath) {
    Import-Module $aiPath -Force -Global
}

#region Helper Functions
function Get-JsonFromAIResponse {
    <#
    .SYNOPSIS
        Extracts JSON from AI responses that may include markdown code blocks
    #>
    param([string]$Response)

    # Remove markdown code blocks if present
    $json = $Response

    # Handle ```json ... ``` format
    if ($json -match '```json\s*([\s\S]*?)\s*```') {
        $json = $matches[1]
    }
    # Handle ``` ... ``` format (no language specified)
    elseif ($json -match '```\s*([\s\S]*?)\s*```') {
        $json = $matches[1]
    }

    # Trim whitespace
    $json = $json.Trim()

    return $json
}
#endregion

#region Template Paths
$script:TemplatesPath = Join-Path $PSScriptRoot "Templates"
#endregion

#region Communication Types
$script:CommunicationPhases = @{
    PreAnnouncement = @{
        Name        = "Pre-Announcement"
        Description = "Initial awareness communication 4-6 weeks before migration"
        Timing      = -42  # days before migration
    }
    DetailedNotice = @{
        Name        = "Detailed Migration Notice"
        Description = "Comprehensive details about what to expect 2-3 weeks before"
        Timing      = -21
    }
    OneWeekReminder = @{
        Name        = "One Week Reminder"
        Description = "Final preparation reminder with checklist"
        Timing      = -7
    }
    DayBefore = @{
        Name        = "Day Before Migration"
        Description = "Last minute reminders and expectations"
        Timing      = -1
    }
    MigrationDay = @{
        Name        = "Migration Day"
        Description = "Migration in progress notification"
        Timing      = 0
    }
    Completion = @{
        Name        = "Migration Complete"
        Description = "Success announcement with next steps"
        Timing      = 1
    }
    PostMigration = @{
        Name        = "Post-Migration Follow-up"
        Description = "Check-in and support resources"
        Timing      = 7
    }
}
#endregion

#region Main Functions
function New-CommunicationPlan {
    <#
    .SYNOPSIS
        Generates a complete communication plan based on migration analysis
    .DESCRIPTION
        Creates a comprehensive communication package including email templates,
        SharePoint news posts, and visual timelines using AI-generated content
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AnalysisResults,

        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $false)]
        [datetime]$TargetMigrationDate = (Get-Date).AddDays(60),

        [Parameter(Mandatory = $false)]
        [hashtable]$Branding,

        [Parameter(Mandatory = $false)]
        [string]$OutputPath = "./Output/Communications"
    )

    Write-Log -Message "Generating communication plan..." -Level Info

    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    # Set default branding if not provided
    if (-not $Branding) {
        $Branding = @{
            CompanyName    = "Your Organization"
            LogoUrl        = $null
            PrimaryColor   = "#0078d4"
            SecondaryColor = "#106ebe"
        }
    }

    # Generate migration summary for AI prompts
    $migrationSummary = Get-MigrationSummaryForAI -AnalysisResults $AnalysisResults -CollectedData $CollectedData

    # Generate all communication components
    $communicationPlan = @{
        GeneratedDate     = Get-Date
        TargetMigration   = $TargetMigrationDate
        Branding          = $Branding
        Summary           = $migrationSummary
        EmailTemplates    = @{}
        SharePointPosts   = @{}
        Timeline          = $null
        SupportResources  = $null
    }

    Write-Log -Message "Generating email templates..." -Level Info
    foreach ($phase in $script:CommunicationPhases.Keys) {
        $phaseDate = $TargetMigrationDate.AddDays($script:CommunicationPhases[$phase].Timing)
        $communicationPlan.EmailTemplates[$phase] = New-EmailTemplate `
            -Phase $phase `
            -MigrationSummary $migrationSummary `
            -Branding $Branding `
            -ScheduledDate $phaseDate
    }

    Write-Log -Message "Generating SharePoint news posts..." -Level Info
    foreach ($phase in $script:CommunicationPhases.Keys) {
        $phaseDate = $TargetMigrationDate.AddDays($script:CommunicationPhases[$phase].Timing)
        $communicationPlan.SharePointPosts[$phase] = New-SharePointNewsPost `
            -Phase $phase `
            -MigrationSummary $migrationSummary `
            -Branding $Branding `
            -ScheduledDate $phaseDate
    }

    Write-Log -Message "Generating visual timeline..." -Level Info
    $communicationPlan.Timeline = New-MigrationTimeline `
        -AnalysisResults $AnalysisResults `
        -TargetMigrationDate $TargetMigrationDate `
        -Branding $Branding

    Write-Log -Message "Generating support resources..." -Level Info
    $communicationPlan.SupportResources = New-SupportResourcesPage `
        -MigrationSummary $migrationSummary `
        -Branding $Branding

    # Export all files
    Export-CommunicationPlan -Plan $communicationPlan -OutputPath $OutputPath

    Write-Log -Message "Communication plan generated successfully at $OutputPath" -Level Success

    return $communicationPlan
}

function Get-MigrationSummaryForAI {
    <#
    .SYNOPSIS
        Creates a summary of migration data for AI content generation
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AnalysisResults,

        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData
    )

    $summary = @{
        UserCount            = 0   # Licensed users to migrate (not guests/unlicensed)
        SharedMailboxCount   = 0   # Shared mailboxes to migrate
        TotalDirectoryUsers  = 0   # Total users in directory (for reference)
        GuestCount           = 0   # Guest users (need reinvitation, not migration)
        MailboxCount         = 0   # User mailboxes
        SharePointSites      = 0   # SharePoint sites (team sites, comm sites - NOT OneDrive)
        OneDriveSites        = 0   # OneDrive personal sites (separate from SharePoint)
        TeamsCount           = 0
        RiskLevel            = $AnalysisResults.RiskLevel
        CriticalIssues       = @()
        HighIssues           = @()
        AffectedServices     = @()
    }

    # Extract user count - ONLY licensed users (excludes guests and unlicensed accounts)
    if ($CollectedData.EntraID.Users) {
        $summary.UserCount = $CollectedData.EntraID.Users.Analysis.LicensedUsers
        $summary.TotalDirectoryUsers = $CollectedData.EntraID.Users.Analysis.TotalUsers
        $summary.GuestCount = $CollectedData.EntraID.Users.Analysis.GuestUsers
    }

    # Extract mailbox count - separate user mailboxes from shared
    if ($CollectedData.Exchange.Mailboxes.Mailboxes) {
        $allMailboxes = $CollectedData.Exchange.Mailboxes.Mailboxes
        $sharedMailboxes = @($allMailboxes | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" })
        $userMailboxes = @($allMailboxes | Where-Object { $_.RecipientTypeDetails -eq "UserMailbox" })
        $summary.MailboxCount = $userMailboxes.Count
        $summary.SharedMailboxCount = $sharedMailboxes.Count
    }

    # Extract SharePoint sites (EXCLUDES OneDrive - those are personal storage, not team sites)
    if ($CollectedData.SharePoint.Sites.Analysis) {
        # Use the analysis count which already excludes OneDrive sites
        $summary.SharePointSites = $CollectedData.SharePoint.Sites.Analysis.SharePointSites
        $summary.OneDriveSites = $CollectedData.SharePoint.Sites.Analysis.OneDriveSites
    }
    elseif ($CollectedData.SharePoint.Sites.Sites) {
        # Fallback: manually filter out OneDrive sites (Template SPSPERS#10 or URL contains -my.sharepoint.com/personal)
        $allSites = $CollectedData.SharePoint.Sites.Sites
        $sharePointOnly = @($allSites | Where-Object {
            $_.Template -ne "SPSPERS#10" -and $_.Url -notlike "*-my.sharepoint.com/personal/*"
        })
        $oneDriveOnly = @($allSites | Where-Object {
            $_.Template -eq "SPSPERS#10" -or $_.Url -like "*-my.sharepoint.com/personal/*"
        })
        $summary.SharePointSites = $sharePointOnly.Count
        $summary.OneDriveSites = $oneDriveOnly.Count
    }

    # Extract Teams
    if ($CollectedData.Teams.Teams) {
        $summary.TeamsCount = ($CollectedData.Teams.Teams.Teams | Measure-Object).Count
    }

    # Get critical and high issues for communication focus
    if ($AnalysisResults.BySeverity) {
        $summary.CriticalIssues = @($AnalysisResults.BySeverity.Critical | ForEach-Object { $_.Name })
        $summary.HighIssues = @($AnalysisResults.BySeverity.High | ForEach-Object { $_.Name })
    }

    # Determine affected services
    $services = @()
    if ($summary.MailboxCount -gt 0) { $services += "Email (Outlook)" }
    if ($summary.SharePointSites -gt 0) { $services += "SharePoint" }
    if ($summary.OneDriveSites -gt 0) { $services += "OneDrive" }
    if ($summary.TeamsCount -gt 0) { $services += "Microsoft Teams" }
    if ($CollectedData.EntraID.Devices) { $services += "Device Access" }
    $summary.AffectedServices = $services

    return $summary
}
#endregion

#region Email Template Functions
function New-EmailTemplate {
    <#
    .SYNOPSIS
        Generates a branded HTML email template for a specific migration phase
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Phase,

        [Parameter(Mandatory = $true)]
        [hashtable]$MigrationSummary,

        [Parameter(Mandatory = $true)]
        [hashtable]$Branding,

        [Parameter(Mandatory = $true)]
        [datetime]$ScheduledDate
    )

    $phaseInfo = $script:CommunicationPhases[$Phase]

    # Generate AI content for this phase
    $aiContent = Get-AIEmailContent -Phase $Phase -PhaseInfo $phaseInfo -MigrationSummary $MigrationSummary

    # Build the HTML template
    $template = Get-EmailHTMLTemplate `
        -Phase $Phase `
        -PhaseInfo $phaseInfo `
        -Content $aiContent `
        -Branding $Branding `
        -ScheduledDate $ScheduledDate `
        -MigrationSummary $MigrationSummary

    return @{
        Phase         = $Phase
        PhaseName     = $phaseInfo.Name
        ScheduledDate = $ScheduledDate
        Subject       = $aiContent.Subject
        HTMLContent   = $template
        PlainText     = $aiContent.PlainText
    }
}

function Get-AIEmailContent {
    <#
    .SYNOPSIS
        Uses AI to generate email content appropriate for non-technical users
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Phase,

        [Parameter(Mandatory = $true)]
        [hashtable]$PhaseInfo,

        [Parameter(Mandatory = $true)]
        [hashtable]$MigrationSummary
    )

    $systemPrompt = @"
You are an expert corporate communications writer specializing in IT change management.
Write clear, friendly, and professional content for non-technical business users.
Avoid jargon. Use simple language. Focus on what users need to DO and how it AFFECTS them.
Be reassuring but honest about any inconveniences. Always include clear action items.
"@

    $userPrompt = @"
Write an all-company email for the "$($PhaseInfo.Name)" phase of our Microsoft 365 migration.

Context:
- We are migrating $($MigrationSummary.UserCount) licensed users to a new Microsoft 365 tenant
- $($MigrationSummary.SharedMailboxCount) shared mailboxes will also be migrated
- Affected services: $($MigrationSummary.AffectedServices -join ", ")
- $($MigrationSummary.MailboxCount) user mailboxes will be migrated
- $($MigrationSummary.SharePointSites) SharePoint team/communication sites will be moved
- $($MigrationSummary.OneDriveSites) OneDrive personal storage accounts will be migrated
- $($MigrationSummary.TeamsCount) Microsoft Teams will be migrated
- Overall risk level: $($MigrationSummary.RiskLevel)

Phase description: $($PhaseInfo.Description)

Please provide:
1. A concise email subject line (max 60 characters)
2. A friendly greeting
3. Main message body (3-4 paragraphs) that explains:
   - What is happening and why
   - How it affects employees
   - What they need to do (if anything)
   - Timeline and next steps
4. A clear call-to-action section
5. Reassuring closing

Format the response as JSON with keys: subject, greeting, body, callToAction, closing
"@

    try {
        $response = Invoke-AIRequest -Prompt $userPrompt -SystemPrompt $systemPrompt -MaxTokens 2000
        $jsonText = Get-JsonFromAIResponse -Response $response
        $content = $jsonText | ConvertFrom-Json

        # Generate plain text version
        $plainText = @"
$($content.subject)

$($content.greeting)

$($content.body)

$($content.callToAction)

$($content.closing)
"@

        return @{
            Subject      = $content.subject
            Greeting     = $content.greeting
            Body         = $content.body
            CallToAction = $content.callToAction
            Closing      = $content.closing
            PlainText    = $plainText
        }
    }
    catch {
        Write-Log -Message "AI content generation failed, using fallback template for $Phase" -Level Warning
        return Get-FallbackEmailContent -Phase $Phase -PhaseInfo $PhaseInfo -MigrationSummary $MigrationSummary
    }
}

function Get-FallbackEmailContent {
    <#
    .SYNOPSIS
        Provides fallback email content if AI generation fails
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Phase,

        [Parameter(Mandatory = $true)]
        [hashtable]$PhaseInfo,

        [Parameter(Mandatory = $true)]
        [hashtable]$MigrationSummary
    )

    $templates = @{
        PreAnnouncement = @{
            Subject      = "Upcoming Microsoft 365 Migration - What You Need to Know"
            Greeting     = "Dear Team,"
            Body         = "We are writing to inform you about an upcoming important change to our technology systems. In the coming weeks, we will be migrating to a new Microsoft 365 environment. This migration will help us improve our collaboration tools, enhance security, and provide you with better technology experiences.`n`nThis change will affect your email, Microsoft Teams, SharePoint, and OneDrive. While we are working hard to minimize any disruption, you may notice some changes during the transition.`n`nOver the next few weeks, we will share more detailed information about the timeline, what to expect, and any actions you may need to take."
            CallToAction = "No action is required at this time. Please watch for future communications with more details."
            Closing      = "Thank you for your patience and cooperation. If you have any questions, please reach out to the IT Help Desk."
        }
        DetailedNotice = @{
            Subject      = "Microsoft 365 Migration - Important Details and Timeline"
            Greeting     = "Dear Team,"
            Body         = "As previously announced, we are migrating to a new Microsoft 365 environment. We now have more details to share about the timeline and what you can expect.`n`nThe migration is scheduled to take place over the coming weeks. During this time, you may experience brief interruptions to email and other Microsoft 365 services.`n`nYour emails, files, and Teams conversations will be moved to the new environment. Most of this will happen automatically, but there are a few things you may need to do."
            CallToAction = "Please take the following steps before the migration:`n- Save any work in progress`n- Note any important calendar appointments`n- Ensure your files are saved to OneDrive or SharePoint"
            Closing      = "We appreciate your cooperation during this transition. More information will follow as we get closer to the migration date."
        }
        OneWeekReminder = @{
            Subject      = "One Week Until Microsoft 365 Migration - Action Required"
            Greeting     = "Dear Team,"
            Body         = "The Microsoft 365 migration is now just one week away. We want to remind you of what to expect and ensure you are prepared.`n`nDuring the migration, you may experience temporary disruptions to email and other Microsoft 365 services. Most services will continue to work, but there may be brief periods where you cannot access certain features."
            CallToAction = "Please complete these steps this week:`n- Review and clean up your mailbox`n- Ensure important files are saved to OneDrive or SharePoint`n- Note your current passwords - you may need to re-enter them`n- Contact IT with any concerns"
            Closing      = "Thank you for preparing for this migration. Your cooperation helps ensure a smooth transition for everyone."
        }
        DayBefore = @{
            Subject      = "Microsoft 365 Migration Tomorrow - Final Reminders"
            Greeting     = "Dear Team,"
            Body         = "The Microsoft 365 migration begins tomorrow. Here are your final reminders and what to expect.`n`nThe migration will begin during off-peak hours to minimize disruption. You may notice some services are temporarily unavailable. This is normal and expected."
            CallToAction = "Before you leave today:`n- Save and close all Microsoft 365 applications`n- Ensure all important work is saved`n- Sign out of Teams and Outlook`n- Do not send large files or important emails until migration is complete"
            Closing      = "We will send an update once the migration is complete. Thank you for your patience."
        }
        MigrationDay = @{
            Subject      = "Microsoft 365 Migration In Progress"
            Greeting     = "Dear Team,"
            Body         = "The Microsoft 365 migration is now in progress. Our IT team is working to complete the transition as quickly as possible.`n`nDuring this time, you may experience intermittent access to email, Teams, SharePoint, and OneDrive. This is temporary and expected."
            CallToAction = "Please be patient during the migration. Avoid:`n- Sending large attachments`n- Making changes to critical documents`n- Creating new Teams or SharePoint sites"
            Closing      = "We will notify you as soon as the migration is complete. Thank you for your understanding."
        }
        Completion = @{
            Subject      = "Microsoft 365 Migration Complete - Welcome to Your New Environment!"
            Greeting     = "Dear Team,"
            Body         = "Great news! The Microsoft 365 migration has been completed successfully. You now have access to your new Microsoft 365 environment.`n`nAll your emails, files, and Teams have been moved. You may need to sign in again to your applications with your credentials."
            CallToAction = "Please take these steps now:`n- Sign in to Outlook and verify your email is accessible`n- Open Teams and check your conversations`n- Access SharePoint and OneDrive to confirm your files`n- Report any issues to IT Help Desk immediately"
            Closing      = "Congratulations on completing the migration! If you encounter any issues, our IT team is ready to help."
        }
        PostMigration = @{
            Subject      = "Microsoft 365 Migration - One Week Check-In"
            Greeting     = "Dear Team,"
            Body         = "It has been one week since our Microsoft 365 migration. We hope you are settling into the new environment well.`n`nWe want to check in and ensure everything is working as expected. If you are experiencing any issues, please let us know right away."
            CallToAction = "Please take a moment to:`n- Verify all your emails and files are accessible`n- Test your Teams and SharePoint access`n- Report any outstanding issues to IT`n- Complete the migration feedback survey (link to follow)"
            Closing      = "Thank you for your patience and cooperation throughout this migration. Your feedback helps us improve future projects."
        }
    }

    $template = $templates[$Phase]
    $plainText = "$($template.Subject)`n`n$($template.Greeting)`n`n$($template.Body)`n`n$($template.CallToAction)`n`n$($template.Closing)"

    return @{
        Subject      = $template.Subject
        Greeting     = $template.Greeting
        Body         = $template.Body
        CallToAction = $template.CallToAction
        Closing      = $template.Closing
        PlainText    = $plainText
    }
}

function Get-EmailHTMLTemplate {
    <#
    .SYNOPSIS
        Generates the complete HTML email template
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Phase,

        [Parameter(Mandatory = $true)]
        [hashtable]$PhaseInfo,

        [Parameter(Mandatory = $true)]
        [hashtable]$Content,

        [Parameter(Mandatory = $true)]
        [hashtable]$Branding,

        [Parameter(Mandatory = $true)]
        [datetime]$ScheduledDate,

        [Parameter(Mandatory = $true)]
        [hashtable]$MigrationSummary
    )

    $logoHtml = if ($Branding.LogoUrl) {
        "<img src=`"$($Branding.LogoUrl)`" alt=`"$($Branding.CompanyName)`" style=`"max-height: 60px; margin-bottom: 20px;`">"
    } else {
        "<h1 style=`"color: $($Branding.PrimaryColor); margin: 0; font-size: 24px;`">$($Branding.CompanyName)</h1>"
    }

    # Convert body paragraphs to HTML
    $bodyHtml = ($Content.Body -split "`n`n" | ForEach-Object {
        "<p style=`"margin: 0 0 16px 0; line-height: 1.6;`">$_</p>"
    }) -join "`n"

    # Convert call to action to HTML list if it contains items
    $ctaHtml = if ($Content.CallToAction -match "^-|^\*|^\d\.") {
        $items = $Content.CallToAction -split "`n" | Where-Object { $_ -match "\S" } | ForEach-Object {
            $item = $_ -replace "^[-*]\s*", "" -replace "^\d+\.\s*", ""
            "<li style=`"margin: 8px 0;`">$item</li>"
        }
        "<ul style=`"margin: 16px 0; padding-left: 24px;`">`n$($items -join "`n")`n</ul>"
    } else {
        "<p style=`"margin: 0 0 16px 0; line-height: 1.6;`">$($Content.CallToAction)</p>"
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$($Content.Subject)</title>
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f5f5f5;">
    <table role="presentation" style="width: 100%; border-collapse: collapse;">
        <tr>
            <td style="padding: 40px 20px;">
                <table role="presentation" style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                    <!-- Header -->
                    <tr>
                        <td style="background: linear-gradient(135deg, $($Branding.PrimaryColor) 0%, $($Branding.SecondaryColor) 100%); padding: 30px; text-align: center;">
                            $logoHtml
                            <p style="color: rgba(255,255,255,0.9); margin: 10px 0 0 0; font-size: 14px;">$($PhaseInfo.Name)</p>
                        </td>
                    </tr>

                    <!-- Date Banner -->
                    <tr>
                        <td style="background-color: #f8f9fa; padding: 15px 30px; border-bottom: 1px solid #e9ecef;">
                            <table role="presentation" style="width: 100%;">
                                <tr>
                                    <td style="color: #6c757d; font-size: 13px;">
                                        <strong>Scheduled:</strong> $($ScheduledDate.ToString("dddd, MMMM d, yyyy"))
                                    </td>
                                    <td style="text-align: right; color: #6c757d; font-size: 13px;">
                                        Migration Update
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- Main Content -->
                    <tr>
                        <td style="padding: 30px;">
                            <p style="margin: 0 0 20px 0; font-size: 16px; color: #333;">$($Content.Greeting)</p>

                            <div style="color: #444; font-size: 15px;">
                                $bodyHtml
                            </div>

                            <!-- Call to Action Box -->
                            <div style="background-color: #e7f3ff; border-left: 4px solid $($Branding.PrimaryColor); padding: 20px; margin: 24px 0; border-radius: 0 4px 4px 0;">
                                <h3 style="margin: 0 0 12px 0; color: $($Branding.PrimaryColor); font-size: 16px;">What You Need to Do</h3>
                                <div style="color: #333; font-size: 14px;">
                                    $ctaHtml
                                </div>
                            </div>

                            <p style="margin: 20px 0 0 0; color: #444; font-size: 15px; line-height: 1.6;">$($Content.Closing)</p>
                        </td>
                    </tr>

                    <!-- Quick Stats (for relevant phases) -->
                    $(if ($Phase -in @("PreAnnouncement", "DetailedNotice")) {
                    @"
                    <tr>
                        <td style="padding: 0 30px 30px 30px;">
                            <table role="presentation" style="width: 100%; background-color: #f8f9fa; border-radius: 8px; overflow: hidden;">
                                <tr>
                                    <td style="padding: 20px; text-align: center; border-right: 1px solid #e9ecef;">
                                        <div style="font-size: 28px; font-weight: bold; color: $($Branding.PrimaryColor);">$($MigrationSummary.UserCount)</div>
                                        <div style="font-size: 12px; color: #6c757d; text-transform: uppercase;">Users</div>
                                    </td>
                                    <td style="padding: 20px; text-align: center; border-right: 1px solid #e9ecef;">
                                        <div style="font-size: 28px; font-weight: bold; color: $($Branding.PrimaryColor);">$($MigrationSummary.MailboxCount)</div>
                                        <div style="font-size: 12px; color: #6c757d; text-transform: uppercase;">Mailboxes</div>
                                    </td>
                                    <td style="padding: 20px; text-align: center;">
                                        <div style="font-size: 28px; font-weight: bold; color: $($Branding.PrimaryColor);">$($MigrationSummary.TeamsCount)</div>
                                        <div style="font-size: 12px; color: #6c757d; text-transform: uppercase;">Teams</div>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
"@
                    })

                    <!-- Footer -->
                    <tr>
                        <td style="background-color: #f8f9fa; padding: 20px 30px; border-top: 1px solid #e9ecef;">
                            <table role="presentation" style="width: 100%;">
                                <tr>
                                    <td style="color: #6c757d; font-size: 13px;">
                                        <p style="margin: 0 0 8px 0;"><strong>Need Help?</strong></p>
                                        <p style="margin: 0;">Contact the IT Help Desk</p>
                                    </td>
                                    <td style="text-align: right; color: #6c757d; font-size: 12px;">
                                        <p style="margin: 0;">$($Branding.CompanyName)</p>
                                        <p style="margin: 4px 0 0 0;">Microsoft 365 Migration Team</p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
"@

    return $html
}
#endregion

#region SharePoint News Functions
function New-SharePointNewsPost {
    <#
    .SYNOPSIS
        Generates SharePoint Online news post content for a migration phase
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Phase,

        [Parameter(Mandatory = $true)]
        [hashtable]$MigrationSummary,

        [Parameter(Mandatory = $true)]
        [hashtable]$Branding,

        [Parameter(Mandatory = $true)]
        [datetime]$ScheduledDate
    )

    $phaseInfo = $script:CommunicationPhases[$Phase]

    # Generate AI content for SharePoint
    $aiContent = Get-AISharePointContent -Phase $Phase -PhaseInfo $phaseInfo -MigrationSummary $MigrationSummary

    # Build the SharePoint HTML content
    $htmlContent = Get-SharePointNewsHTML `
        -Phase $Phase `
        -PhaseInfo $phaseInfo `
        -Content $aiContent `
        -Branding $Branding `
        -ScheduledDate $ScheduledDate `
        -MigrationSummary $MigrationSummary

    return @{
        Phase         = $Phase
        PhaseName     = $phaseInfo.Name
        ScheduledDate = $ScheduledDate
        Title         = $aiContent.Title
        Summary       = $aiContent.Summary
        HTMLContent   = $htmlContent
        BannerImage   = "migration-banner-$Phase.png"
        Category      = "IT Updates"
        IsFeatured    = $Phase -in @("PreAnnouncement", "MigrationDay", "Completion")
    }
}

function Get-AISharePointContent {
    <#
    .SYNOPSIS
        Uses AI to generate SharePoint news post content
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Phase,

        [Parameter(Mandatory = $true)]
        [hashtable]$PhaseInfo,

        [Parameter(Mandatory = $true)]
        [hashtable]$MigrationSummary
    )

    $systemPrompt = @"
You are an expert internal communications writer creating SharePoint news articles.
Write engaging, scannable content optimized for busy employees.
Use headers, bullet points, and short paragraphs.
Include helpful visuals descriptions where appropriate.
Focus on employee impact and clear action items.
"@

    $userPrompt = @"
Write a SharePoint news post for the "$($PhaseInfo.Name)" phase of our Microsoft 365 migration.

Context:
- Migrating $($MigrationSummary.UserCount) licensed users plus $($MigrationSummary.SharedMailboxCount) shared mailboxes
- Affected: $($MigrationSummary.AffectedServices -join ", ")
- $($MigrationSummary.MailboxCount) user mailboxes, $($MigrationSummary.SharePointSites) SharePoint sites, $($MigrationSummary.OneDriveSites) OneDrive accounts, $($MigrationSummary.TeamsCount) Teams
- Risk level: $($MigrationSummary.RiskLevel)

Phase: $($PhaseInfo.Description)

Provide a JSON response with:
1. title: Engaging headline (max 80 chars)
2. summary: 2 sentence description for news feed
3. intro: Opening paragraph
4. sections: Array of {header, content} for main content sections
5. timeline: Key dates/milestones if relevant
6. faqs: Array of {question, answer} (3-5 common questions)
7. callToAction: What employees should do
"@

    try {
        $response = Invoke-AIRequest -Prompt $userPrompt -SystemPrompt $systemPrompt -MaxTokens 3000
        $jsonText = Get-JsonFromAIResponse -Response $response
        $content = $jsonText | ConvertFrom-Json
        return $content
    }
    catch {
        Write-Log -Message "AI content generation failed, using fallback for SharePoint $Phase" -Level Warning
        return Get-FallbackSharePointContent -Phase $Phase -PhaseInfo $PhaseInfo -MigrationSummary $MigrationSummary
    }
}

function Get-FallbackSharePointContent {
    <#
    .SYNOPSIS
        Provides fallback SharePoint content if AI generation fails
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Phase,

        [Parameter(Mandatory = $true)]
        [hashtable]$PhaseInfo,

        [Parameter(Mandatory = $true)]
        [hashtable]$MigrationSummary
    )

    $templates = @{
        PreAnnouncement = @{
            Title       = "Upcoming Changes: Microsoft 365 Migration Announced"
            Summary     = "We are migrating to a new Microsoft 365 environment. Here's what you need to know about this important update."
            Intro       = "We are excited to announce an upcoming migration to a new Microsoft 365 environment. This change will improve our technology infrastructure and provide you with enhanced collaboration tools."
            Sections    = @(
                @{ Header = "Why Are We Doing This?"; Content = "This migration will help us improve security, enhance our collaboration capabilities, and provide you with a better technology experience." }
                @{ Header = "What Will Change?"; Content = "Your email, Teams, SharePoint, and OneDrive will be moved to the new environment. While things will look similar, you may notice some improvements." }
                @{ Header = "Timeline"; Content = "The migration will occur over the coming weeks. We will provide more specific dates as we finalize our plans." }
            )
            FAQs        = @(
                @{ Question = "Will I lose any emails or files?"; Answer = "No, all your data will be carefully migrated to the new environment." }
                @{ Question = "Will I need a new password?"; Answer = "You may need to re-enter your password after migration, but your credentials will remain the same." }
                @{ Question = "What should I do now?"; Answer = "Nothing is required at this time. Watch for future updates with more details." }
            )
            CallToAction = "No action required at this time. Stay tuned for more information."
        }
        DetailedNotice = @{
            Title       = "Microsoft 365 Migration: Detailed Timeline and Preparation Guide"
            Summary     = "Get all the details about our upcoming Microsoft 365 migration, including timeline, what to expect, and how to prepare."
            Intro       = "The Microsoft 365 migration is approaching. This article provides everything you need to know to prepare."
            Sections    = @(
                @{ Header = "Migration Timeline"; Content = "The migration will begin soon. Here are the key milestones you should be aware of." }
                @{ Header = "What to Expect"; Content = "During the migration, you may experience brief interruptions to email and other services. This is temporary." }
                @{ Header = "How to Prepare"; Content = "Review your important emails and files. Ensure everything is saved to OneDrive or SharePoint." }
            )
            FAQs        = @(
                @{ Question = "How long will the migration take?"; Answer = "The main migration will occur over a weekend to minimize disruption." }
                @{ Question = "Can I work during the migration?"; Answer = "We recommend completing critical work before the migration window." }
                @{ Question = "Who should I contact with questions?"; Answer = "Reach out to the IT Help Desk for any concerns." }
            )
            CallToAction = "Review your files and prepare for the migration. Contact IT with any questions."
        }
    }

    # Return template for phase or generic fallback
    if ($templates.ContainsKey($Phase)) {
        return $templates[$Phase]
    }

    return @{
        Title       = "$($PhaseInfo.Name) - Microsoft 365 Migration Update"
        Summary     = "Important update about our Microsoft 365 migration. Please read for the latest information."
        Intro       = "Here is the latest update on our Microsoft 365 migration project."
        Sections    = @(@{ Header = "Update"; Content = $PhaseInfo.Description })
        FAQs        = @(@{ Question = "Need help?"; Answer = "Contact the IT Help Desk for assistance." })
        CallToAction = "Please review this update and contact IT if you have questions."
    }
}

function Get-SharePointNewsHTML {
    <#
    .SYNOPSIS
        Generates SharePoint-optimized HTML content
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Phase,

        [Parameter(Mandatory = $true)]
        [hashtable]$PhaseInfo,

        [Parameter(Mandatory = $true)]
        $Content,

        [Parameter(Mandatory = $true)]
        [hashtable]$Branding,

        [Parameter(Mandatory = $true)]
        [datetime]$ScheduledDate,

        [Parameter(Mandatory = $true)]
        [hashtable]$MigrationSummary
    )

    # Build sections HTML
    $sectionsHtml = ""
    if ($Content.sections) {
        foreach ($section in $Content.sections) {
            $sectionsHtml += @"
            <div class="section" style="margin-bottom: 32px;">
                <h2 style="color: $($Branding.PrimaryColor); font-size: 22px; margin-bottom: 12px; border-bottom: 2px solid $($Branding.SecondaryColor); padding-bottom: 8px;">$($section.header)</h2>
                <p style="color: #333; line-height: 1.7; font-size: 15px;">$($section.content)</p>
            </div>
"@
        }
    }

    # Build FAQs HTML
    $faqsHtml = ""
    if ($Content.faqs) {
        $faqsHtml = @"
        <div class="faqs" style="background-color: #f8f9fa; padding: 24px; border-radius: 8px; margin: 32px 0;">
            <h2 style="color: $($Branding.PrimaryColor); font-size: 22px; margin-bottom: 20px;">Frequently Asked Questions</h2>
"@
        foreach ($faq in $Content.faqs) {
            $faqsHtml += @"
            <div class="faq-item" style="margin-bottom: 20px;">
                <h3 style="color: #333; font-size: 16px; margin-bottom: 8px; font-weight: 600;">$($faq.question)</h3>
                <p style="color: #666; font-size: 14px; line-height: 1.6; margin: 0;">$($faq.answer)</p>
            </div>
"@
        }
        $faqsHtml += "</div>"
    }

    $html = @"
<div class="news-article" style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 800px; margin: 0 auto;">
    <!-- Hero Section -->
    <div class="hero" style="background: linear-gradient(135deg, $($Branding.PrimaryColor) 0%, $($Branding.SecondaryColor) 100%); color: white; padding: 40px; border-radius: 8px; margin-bottom: 32px; text-align: center;">
        <span style="background-color: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 20px; font-size: 12px; text-transform: uppercase; letter-spacing: 1px;">$($PhaseInfo.Name)</span>
        <h1 style="font-size: 32px; margin: 20px 0 10px 0; font-weight: 600;">$($Content.title)</h1>
        <p style="opacity: 0.9; margin: 0; font-size: 16px;">$($ScheduledDate.ToString("MMMM d, yyyy"))</p>
    </div>

    <!-- Quick Stats -->
    <div class="stats-row" style="display: flex; justify-content: space-around; background-color: #fff; padding: 24px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); margin-bottom: 32px;">
        <div class="stat" style="text-align: center;">
            <div style="font-size: 36px; font-weight: bold; color: $($Branding.PrimaryColor);">$($MigrationSummary.UserCount)</div>
            <div style="font-size: 13px; color: #6c757d; text-transform: uppercase;">Users Affected</div>
        </div>
        <div class="stat" style="text-align: center;">
            <div style="font-size: 36px; font-weight: bold; color: $($Branding.PrimaryColor);">$($MigrationSummary.MailboxCount)</div>
            <div style="font-size: 13px; color: #6c757d; text-transform: uppercase;">Mailboxes</div>
        </div>
        <div class="stat" style="text-align: center;">
            <div style="font-size: 36px; font-weight: bold; color: $($Branding.PrimaryColor);">$($MigrationSummary.TeamsCount)</div>
            <div style="font-size: 13px; color: #6c757d; text-transform: uppercase;">Teams</div>
        </div>
        <div class="stat" style="text-align: center;">
            <div style="font-size: 36px; font-weight: bold; color: $($Branding.PrimaryColor);">$($MigrationSummary.SharePointSites)</div>
            <div style="font-size: 13px; color: #6c757d; text-transform: uppercase;">SharePoint Sites</div>
        </div>
        <div class="stat" style="text-align: center;">
            <div style="font-size: 36px; font-weight: bold; color: $($Branding.PrimaryColor);">$($MigrationSummary.OneDriveSites)</div>
            <div style="font-size: 13px; color: #6c757d; text-transform: uppercase;">OneDrive Accounts</div>
        </div>
    </div>

    <!-- Introduction -->
    <div class="intro" style="font-size: 18px; color: #444; line-height: 1.8; margin-bottom: 32px; padding: 0 16px;">
        $($Content.intro)
    </div>

    <!-- Main Content Sections -->
    <div class="content" style="padding: 0 16px;">
        $sectionsHtml
    </div>

    <!-- Call to Action -->
    <div class="cta" style="background: linear-gradient(135deg, $($Branding.PrimaryColor) 0%, $($Branding.SecondaryColor) 100%); color: white; padding: 32px; border-radius: 8px; margin: 32px 0; text-align: center;">
        <h2 style="margin: 0 0 16px 0; font-size: 24px;">What You Need to Do</h2>
        <p style="font-size: 16px; margin: 0; opacity: 0.95;">$($Content.callToAction)</p>
    </div>

    <!-- FAQs -->
    $faqsHtml

    <!-- Support Footer -->
    <div class="support" style="background-color: #e7f3ff; padding: 24px; border-radius: 8px; text-align: center; margin-top: 32px;">
        <h3 style="color: $($Branding.PrimaryColor); margin: 0 0 12px 0;">Need Help?</h3>
        <p style="margin: 0; color: #333;">Contact the <strong>IT Help Desk</strong> for assistance with any migration-related questions.</p>
    </div>
</div>
"@

    return $html
}
#endregion

#region Timeline Functions
function New-MigrationTimeline {
    <#
    .SYNOPSIS
        Generates a visual migration timeline using Chart.js
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AnalysisResults,

        [Parameter(Mandatory = $true)]
        [datetime]$TargetMigrationDate,

        [Parameter(Mandatory = $true)]
        [hashtable]$Branding
    )

    # Generate timeline milestones
    $milestones = @(
        @{
            Phase       = "Planning"
            StartOffset = -60
            Duration    = 14
            Color       = "#6c757d"
            Icon        = "clipboard-list"
            Tasks       = @("Discovery complete", "Risk assessment", "Communication plan")
        }
        @{
            Phase       = "Preparation"
            StartOffset = -46
            Duration    = 21
            Color       = "#17a2b8"
            Icon        = "cog"
            Tasks       = @("Environment setup", "User training", "Pilot testing")
        }
        @{
            Phase       = "Pre-Migration"
            StartOffset = -25
            Duration    = 18
            Color       = "#ffc107"
            Icon        = "exclamation-triangle"
            Tasks       = @("Final preparations", "User communications", "Backup verification")
        }
        @{
            Phase       = "Migration"
            StartOffset = -7
            Duration    = 7
            Color       = $Branding.PrimaryColor
            Icon        = "sync"
            Tasks       = @("Data migration", "Service cutover", "Validation testing")
        }
        @{
            Phase       = "Stabilization"
            StartOffset = 0
            Duration    = 14
            Color       = "#28a745"
            Icon        = "check-circle"
            Tasks       = @("Issue resolution", "User support", "Performance monitoring")
        }
        @{
            Phase       = "Optimization"
            StartOffset = 14
            Duration    = 21
            Color       = "#20c997"
            Icon        = "chart-line"
            Tasks       = @("Feature adoption", "Training follow-up", "Documentation")
        }
    )

    # Calculate actual dates
    foreach ($milestone in $milestones) {
        $milestone.StartDate = $TargetMigrationDate.AddDays($milestone.StartOffset)
        $milestone.EndDate = $milestone.StartDate.AddDays($milestone.Duration)
    }

    # Generate timeline HTML with Chart.js
    $html = New-TimelineHTML -Milestones $milestones -Branding $Branding -TargetDate $TargetMigrationDate

    return @{
        Milestones    = $milestones
        TargetDate    = $TargetMigrationDate
        HTMLContent   = $html
    }
}

function New-TimelineHTML {
    <#
    .SYNOPSIS
        Generates interactive timeline HTML with Chart.js visualization
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Milestones,

        [Parameter(Mandatory = $true)]
        [hashtable]$Branding,

        [Parameter(Mandatory = $true)]
        [datetime]$TargetDate
    )

    # Build datasets for Chart.js Gantt-style chart
    $datasets = @()
    $index = 0
    foreach ($milestone in $Milestones) {
        $datasets += @{
            label           = $milestone.Phase
            data            = @(@{
                x = @($milestone.StartDate.ToString("yyyy-MM-dd"), $milestone.EndDate.ToString("yyyy-MM-dd"))
                y = $milestone.Phase
            })
            backgroundColor = $milestone.Color
            borderRadius    = 4
            barPercentage   = 0.6
        }
        $index++
    }

    $datasetsJson = $datasets | ConvertTo-Json -Depth 10

    # Build milestone cards with dark theme
    $milestoneCards = ""
    foreach ($milestone in $Milestones) {
        $tasksList = ($milestone.Tasks | ForEach-Object { "<li>$_</li>" }) -join ""
        $milestoneCards += @"
        <div class="milestone-card">
            <div class="milestone-header">
                <div class="milestone-icon" style="background: linear-gradient(135deg, $($milestone.Color), $($milestone.Color)aa); box-shadow: 0 0 20px $($milestone.Color)66;">
                    <span>&#9679;</span>
                </div>
                <div>
                    <h3>$($milestone.Phase)</h3>
                    <p class="date-range">$($milestone.StartDate.ToString("MMM d")) - $($milestone.EndDate.ToString("MMM d, yyyy"))</p>
                </div>
            </div>
            <ul>
                $tasksList
            </ul>
        </div>
"@
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Migration Timeline - $($Branding.CompanyName)</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-adapter-date-fns"></script>
    <style>
        :root {
            --bg-primary: #0f172a;
            --bg-secondary: #1e293b;
            --bg-tertiary: #334155;
            --bg-glass: rgba(30, 41, 59, 0.8);
            --text-primary: #f1f5f9;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --border-glass: rgba(148, 163, 184, 0.1);
            --accent-cyan: #06b6d4;
            --accent-purple: #8b5cf6;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, var(--bg-primary) 0%, #0c1929 50%, var(--bg-primary) 100%);
            min-height: 100vh;
            color: var(--text-primary);
        }
        body::before {
            content: '';
            position: fixed;
            top: 0; left: 0; width: 100%; height: 100%;
            background: radial-gradient(ellipse at 20% 20%, rgba(6, 182, 212, 0.1) 0%, transparent 50%),
                        radial-gradient(ellipse at 80% 80%, rgba(139, 92, 246, 0.1) 0%, transparent 50%);
            pointer-events: none;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 40px 20px;
            position: relative;
            z-index: 1;
        }
        .header {
            background: linear-gradient(135deg, rgba(6, 182, 212, 0.2), rgba(139, 92, 246, 0.2));
            backdrop-filter: blur(20px);
            border: 1px solid var(--border-glass);
            padding: 50px;
            border-radius: 24px;
            margin-bottom: 40px;
            text-align: center;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3), 0 0 60px rgba(6, 182, 212, 0.1);
        }
        .header h1 {
            font-size: 2.8em;
            margin-bottom: 12px;
            background: linear-gradient(135deg, var(--text-primary), var(--accent-cyan));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        .header p { color: var(--text-secondary); font-size: 1.2em; }
        .target-date {
            display: inline-block;
            background: linear-gradient(135deg, var(--accent-cyan), var(--accent-purple));
            padding: 14px 28px;
            border-radius: 30px;
            margin-top: 25px;
            font-size: 1.2em;
            font-weight: 600;
            box-shadow: 0 4px 20px rgba(6, 182, 212, 0.4);
        }
        .chart-container {
            background: var(--bg-glass);
            backdrop-filter: blur(20px);
            border: 1px solid var(--border-glass);
            border-radius: 20px;
            padding: 35px;
            margin-bottom: 40px;
            box-shadow: 8px 8px 16px rgba(0, 0, 0, 0.4), -4px -4px 12px rgba(255, 255, 255, 0.02);
        }
        .chart-title {
            color: var(--text-primary);
            font-size: 1.6em;
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 2px solid rgba(6, 182, 212, 0.3);
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .chart-title::before {
            content: '';
            width: 4px;
            height: 28px;
            background: linear-gradient(180deg, var(--accent-cyan), var(--accent-purple));
            border-radius: 2px;
        }
        .milestones-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
            gap: 24px;
            margin-top: 30px;
        }
        .milestone-card {
            background: var(--bg-glass);
            backdrop-filter: blur(20px);
            border: 1px solid var(--border-glass);
            border-radius: 16px;
            padding: 24px;
            box-shadow: 8px 8px 16px rgba(0, 0, 0, 0.3), -4px -4px 12px rgba(255, 255, 255, 0.02);
            transition: all 0.3s ease;
        }
        .milestone-card:hover {
            transform: translateY(-4px);
            box-shadow: 8px 8px 16px rgba(0, 0, 0, 0.3), -4px -4px 12px rgba(255, 255, 255, 0.02), 0 0 30px rgba(6, 182, 212, 0.2);
        }
        .milestone-header {
            display: flex;
            align-items: center;
            margin-bottom: 16px;
        }
        .milestone-icon {
            width: 48px;
            height: 48px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: 16px;
            font-size: 20px;
            color: white;
        }
        .milestone-card h3 {
            color: var(--text-primary);
            font-size: 1.1em;
            margin: 0;
        }
        .milestone-card .date-range {
            color: var(--text-muted);
            font-size: 0.9em;
            margin-top: 4px;
        }
        .milestone-card ul {
            margin: 0;
            padding-left: 24px;
            color: var(--text-secondary);
        }
        .milestone-card li {
            margin: 8px 0;
            font-size: 0.95em;
        }
        .legend {
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
            justify-content: center;
            margin-top: 25px;
        }
        .legend-item {
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 0.9em;
            color: var(--text-secondary);
        }
        .legend-color {
            width: 18px;
            height: 18px;
            border-radius: 6px;
            box-shadow: 0 0 10px currentColor;
        }
        .section-title {
            color: var(--text-primary);
            font-size: 1.6em;
            margin-bottom: 25px;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .section-title::before {
            content: '';
            width: 4px;
            height: 28px;
            background: linear-gradient(180deg, var(--accent-cyan), var(--accent-purple));
            border-radius: 2px;
        }
        ::-webkit-scrollbar { width: 10px; }
        ::-webkit-scrollbar-track { background: var(--bg-secondary); }
        ::-webkit-scrollbar-thumb { background: var(--bg-tertiary); border-radius: 5px; }
    </style>
</head>
<body>
    <div class="container">
        <!-- Header -->
        <div class="header">
            <h1>Microsoft 365 Migration Timeline</h1>
            <p>$($Branding.CompanyName) - Complete Migration Schedule</p>
            <div class="target-date">
                Target Go-Live: $($TargetDate.ToString("MMMM d, yyyy"))
            </div>
        </div>

        <!-- Gantt Chart -->
        <div class="chart-container">
            <h2 class="chart-title">Project Timeline</h2>
            <div style="position: relative; height: 400px;">
                <canvas id="timelineChart"></canvas>
            </div>
            <div class="legend">
                $(foreach ($m in $Milestones) {
                    "<div class='legend-item'><div class='legend-color' style='background-color: $($m.Color); color: $($m.Color);'></div>$($m.Phase)</div>"
                })
            </div>
        </div>

        <!-- Milestone Cards -->
        <h2 class="section-title">Phase Details</h2>
        <div class="milestones-grid">
            $milestoneCards
        </div>
    </div>

    <script>
        const ctx = document.getElementById('timelineChart').getContext('2d');

        // Prepare data for horizontal bar chart (Gantt style)
        const phases = [$( ($Milestones | ForEach-Object { "'$($_.Phase)'" }) -join ", " )];
        const startDates = [$( ($Milestones | ForEach-Object { "new Date('$($_.StartDate.ToString("yyyy-MM-dd"))')" }) -join ", " )];
        const durations = [$( ($Milestones | ForEach-Object { $_.Duration }) -join ", " )];
        const colors = [$( ($Milestones | ForEach-Object { "'$($_.Color)'" }) -join ", " )];

        const today = new Date();

        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: phases,
                datasets: [{
                    label: 'Duration (days)',
                    data: phases.map((phase, i) => ({
                        x: [startDates[i], new Date(startDates[i].getTime() + durations[i] * 24 * 60 * 60 * 1000)],
                        y: phase
                    })),
                    backgroundColor: colors,
                    borderRadius: 6,
                    borderSkipped: false,
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const start = new Date(context.raw.x[0]).toLocaleDateString();
                                const end = new Date(context.raw.x[1]).toLocaleDateString();
                                return start + ' - ' + end;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        type: 'time',
                        time: {
                            unit: 'week',
                            displayFormats: {
                                week: 'MMM d'
                            }
                        },
                        title: {
                            display: true,
                            text: 'Timeline',
                            color: '#94a3b8'
                        },
                        grid: {
                            color: 'rgba(148, 163, 184, 0.1)'
                        },
                        ticks: {
                            color: '#94a3b8'
                        }
                    },
                    y: {
                        title: {
                            display: true,
                            text: 'Project Phases',
                            color: '#94a3b8'
                        },
                        grid: {
                            display: false
                        },
                        ticks: {
                            color: '#f1f5f9'
                        }
                    }
                }
            }
        });
    </script>
</body>
</html>
"@

    return $html
}
#endregion

#region Support Resources
function New-SupportResourcesPage {
    <#
    .SYNOPSIS
        Generates a support resources page for end users
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$MigrationSummary,

        [Parameter(Mandatory = $true)]
        [hashtable]$Branding
    )

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Migration Support Resources - $($Branding.CompanyName)</title>
    <style>
        :root {
            --bg-primary: #0f172a;
            --bg-secondary: #1e293b;
            --bg-tertiary: #334155;
            --bg-glass: rgba(30, 41, 59, 0.8);
            --text-primary: #f1f5f9;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --border-glass: rgba(148, 163, 184, 0.1);
            --accent-cyan: #06b6d4;
            --accent-purple: #8b5cf6;
            --accent-green: #10b981;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, var(--bg-primary) 0%, #0c1929 50%, var(--bg-primary) 100%);
            min-height: 100vh;
            color: var(--text-primary);
        }
        body::before {
            content: '';
            position: fixed;
            top: 0; left: 0; width: 100%; height: 100%;
            background: radial-gradient(ellipse at 30% 20%, rgba(6, 182, 212, 0.08) 0%, transparent 50%),
                        radial-gradient(ellipse at 70% 80%, rgba(139, 92, 246, 0.08) 0%, transparent 50%);
            pointer-events: none;
        }
        .container {
            max-width: 1000px;
            margin: 0 auto;
            padding: 40px 20px;
            position: relative;
            z-index: 1;
        }
        .header {
            background: linear-gradient(135deg, rgba(6, 182, 212, 0.2), rgba(139, 92, 246, 0.2));
            backdrop-filter: blur(20px);
            border: 1px solid var(--border-glass);
            padding: 50px;
            border-radius: 24px;
            margin-bottom: 40px;
            text-align: center;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3), 0 0 60px rgba(6, 182, 212, 0.1);
        }
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 12px;
            background: linear-gradient(135deg, var(--text-primary), var(--accent-cyan));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        .header p { color: var(--text-secondary); font-size: 1.15em; }
        .section {
            background: var(--bg-glass);
            backdrop-filter: blur(20px);
            border: 1px solid var(--border-glass);
            border-radius: 20px;
            padding: 35px;
            margin-bottom: 28px;
            box-shadow: 8px 8px 16px rgba(0, 0, 0, 0.3), -4px -4px 12px rgba(255, 255, 255, 0.02);
        }
        .section h2 {
            color: var(--text-primary);
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 2px solid rgba(6, 182, 212, 0.3);
            font-size: 1.4em;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .section h2::before {
            content: '';
            width: 4px;
            height: 24px;
            background: linear-gradient(180deg, var(--accent-cyan), var(--accent-purple));
            border-radius: 2px;
        }
        .resource-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 20px;
        }
        .resource-card {
            background: rgba(51, 65, 85, 0.5);
            border-radius: 12px;
            padding: 24px;
            border-left: 4px solid var(--accent-cyan);
            transition: all 0.3s ease;
        }
        .resource-card:hover {
            transform: translateX(4px);
            box-shadow: 0 0 25px rgba(6, 182, 212, 0.2);
        }
        .resource-card h3 { margin-bottom: 10px; color: var(--text-primary); font-size: 1.1em; }
        .resource-card p { color: var(--text-secondary); font-size: 0.95em; line-height: 1.6; }
        .faq-item {
            border-bottom: 1px solid var(--border-glass);
            padding: 20px 0;
        }
        .faq-item:last-child { border-bottom: none; }
        .faq-question {
            font-weight: 600;
            color: var(--accent-cyan);
            margin-bottom: 10px;
            font-size: 1.05em;
        }
        .faq-answer { color: var(--text-secondary); line-height: 1.7; }
        .contact-box {
            background: linear-gradient(135deg, rgba(6, 182, 212, 0.3), rgba(139, 92, 246, 0.3));
            backdrop-filter: blur(20px);
            border: 1px solid rgba(6, 182, 212, 0.3);
            padding: 40px;
            border-radius: 20px;
            text-align: center;
            margin-top: 40px;
            box-shadow: 0 0 40px rgba(6, 182, 212, 0.15);
        }
        .contact-box h2 { margin-bottom: 16px; font-size: 1.6em; }
        .contact-box p { color: var(--text-secondary); margin-bottom: 8px; }
        .contact-box .cta {
            display: inline-block;
            background: linear-gradient(135deg, var(--accent-cyan), var(--accent-purple));
            padding: 14px 32px;
            border-radius: 30px;
            margin-top: 20px;
            font-size: 1.1em;
            font-weight: 600;
            box-shadow: 0 4px 20px rgba(6, 182, 212, 0.4);
        }
        .checklist {
            list-style: none;
            padding: 0;
        }
        .checklist li {
            padding: 16px 0 16px 40px;
            position: relative;
            border-bottom: 1px solid var(--border-glass);
            color: var(--text-secondary);
            transition: all 0.2s ease;
        }
        .checklist li:hover {
            color: var(--text-primary);
            padding-left: 44px;
        }
        .checklist li:before {
            content: "☐";
            position: absolute;
            left: 0;
            color: var(--accent-cyan);
            font-size: 20px;
            text-shadow: 0 0 10px var(--accent-cyan);
        }
        ::-webkit-scrollbar { width: 10px; }
        ::-webkit-scrollbar-track { background: var(--bg-secondary); }
        ::-webkit-scrollbar-thumb { background: var(--bg-tertiary); border-radius: 5px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Migration Support Resources</h1>
            <p>Everything you need for a smooth transition to your new Microsoft 365 environment</p>
        </div>

        <!-- Quick Links -->
        <div class="section">
            <h2>Quick Start Guides</h2>
            <div class="resource-grid">
                <div class="resource-card">
                    <h3>Email (Outlook)</h3>
                    <p>How to access your email after migration, set up your profile, and restore your settings.</p>
                </div>
                <div class="resource-card">
                    <h3>Microsoft Teams</h3>
                    <p>Reconnecting to Teams, accessing your chats, and joining your existing teams.</p>
                </div>
                <div class="resource-card">
                    <h3>SharePoint & OneDrive</h3>
                    <p>Accessing your files, syncing folders, and navigating the new environment.</p>
                </div>
                <div class="resource-card">
                    <h3>Mobile Devices</h3>
                    <p>Setting up Outlook, Teams, and OneDrive on your phone or tablet.</p>
                </div>
            </div>
        </div>

        <!-- Pre-Migration Checklist -->
        <div class="section">
            <h2>Pre-Migration Checklist</h2>
            <ul class="checklist">
                <li>Save any work in progress to OneDrive or SharePoint</li>
                <li>Note any important calendar appointments for the migration window</li>
                <li>Export any local Outlook rules you want to keep</li>
                <li>Make note of any shared mailbox access you have</li>
                <li>Save any important Teams chat messages</li>
                <li>Close all Microsoft 365 applications before migration begins</li>
                <li>Ensure you know your password (you may need to re-enter it)</li>
            </ul>
        </div>

        <!-- Post-Migration Checklist -->
        <div class="section">
            <h2>Post-Migration Checklist</h2>
            <ul class="checklist">
                <li>Sign in to Outlook and verify your email is accessible</li>
                <li>Check that your calendar appointments are visible</li>
                <li>Open Microsoft Teams and verify your teams and chats</li>
                <li>Access OneDrive and check your files</li>
                <li>Test sending and receiving email</li>
                <li>Reconfigure any custom Outlook rules if needed</li>
                <li>Update saved passwords in your browser if prompted</li>
            </ul>
        </div>

        <!-- FAQs -->
        <div class="section">
            <h2>Frequently Asked Questions</h2>
            <div class="faq-item">
                <div class="faq-question">Will I lose any emails or files during the migration?</div>
                <div class="faq-answer">No, all your emails, files, and Teams data will be carefully migrated to the new environment. We perform extensive validation to ensure nothing is lost.</div>
            </div>
            <div class="faq-item">
                <div class="faq-question">Will my password change?</div>
                <div class="faq-answer">Your password will remain the same. However, you may need to re-enter it when you first sign in to applications after the migration.</div>
            </div>
            <div class="faq-item">
                <div class="faq-question">What about my calendar appointments?</div>
                <div class="faq-answer">All your calendar appointments, including recurring meetings, will be migrated. Meeting links may be updated, and organizers will send updated invites if needed.</div>
            </div>
            <div class="faq-item">
                <div class="faq-question">Will my Teams chats and files be available?</div>
                <div class="faq-answer">Yes, your Teams conversations, channels, and files will be migrated. You may need to sign out and back in to Teams after the migration.</div>
            </div>
            <div class="faq-item">
                <div class="faq-question">What if I experience issues after migration?</div>
                <div class="faq-answer">Contact the IT Help Desk immediately. We have additional support staff available during and after the migration to help resolve any issues quickly.</div>
            </div>
            <div class="faq-item">
                <div class="faq-question">Do I need to do anything to my mobile devices?</div>
                <div class="faq-answer">You may need to remove and re-add your email account on mobile devices. The IT Help Desk can assist with this if needed.</div>
            </div>
        </div>

        <!-- Contact -->
        <div class="contact-box">
            <h2>Need Help?</h2>
            <p>Our IT Help Desk is here to support you throughout the migration.</p>
            <div class="cta">Contact the IT Help Desk</div>
        </div>
    </div>
</body>
</html>
"@

    return @{
        HTMLContent = $html
        Title       = "Migration Support Resources"
    }
}
#endregion

#region Export Functions
function Export-CommunicationPlan {
    <#
    .SYNOPSIS
        Exports all communication plan components to files
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Plan,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    # Create subdirectories
    $emailsPath = Join-Path $OutputPath "Emails"
    $spoPath = Join-Path $OutputPath "SharePointPosts"
    $timelinePath = Join-Path $OutputPath "Timeline"
    $resourcesPath = Join-Path $OutputPath "Resources"

    @($emailsPath, $spoPath, $timelinePath, $resourcesPath) | ForEach-Object {
        if (-not (Test-Path $_)) {
            New-Item -Path $_ -ItemType Directory -Force | Out-Null
        }
    }

    # Export email templates
    Write-Log -Message "Exporting email templates..." -Level Info
    foreach ($phase in $Plan.EmailTemplates.Keys) {
        $template = $Plan.EmailTemplates[$phase]
        $htmlFile = Join-Path $emailsPath "$phase-email.html"
        $txtFile = Join-Path $emailsPath "$phase-email.txt"

        $template.HTMLContent | Out-File -FilePath $htmlFile -Encoding utf8
        $template.PlainText | Out-File -FilePath $txtFile -Encoding utf8
    }

    # Export SharePoint posts
    Write-Log -Message "Exporting SharePoint news posts..." -Level Info
    foreach ($phase in $Plan.SharePointPosts.Keys) {
        $post = $Plan.SharePointPosts[$phase]
        $htmlFile = Join-Path $spoPath "$phase-news.html"
        $post.HTMLContent | Out-File -FilePath $htmlFile -Encoding utf8
    }

    # Export timeline
    Write-Log -Message "Exporting timeline..." -Level Info
    $timelineFile = Join-Path $timelinePath "migration-timeline.html"
    $Plan.Timeline.HTMLContent | Out-File -FilePath $timelineFile -Encoding utf8

    # Export support resources
    Write-Log -Message "Exporting support resources..." -Level Info
    $resourcesFile = Join-Path $resourcesPath "support-resources.html"
    $Plan.SupportResources.HTMLContent | Out-File -FilePath $resourcesFile -Encoding utf8

    # Create index/summary file
    $summaryFile = Join-Path $OutputPath "communication-plan-summary.html"
    New-CommunicationPlanSummary -Plan $Plan -OutputPath $OutputPath | Out-File -FilePath $summaryFile -Encoding utf8

    # Export plan as JSON for reference
    $jsonFile = Join-Path $OutputPath "communication-plan.json"
    $planData = @{
        GeneratedDate   = $Plan.GeneratedDate.ToString("yyyy-MM-dd HH:mm:ss")
        TargetMigration = $Plan.TargetMigration.ToString("yyyy-MM-dd")
        Branding        = $Plan.Branding
        Summary         = $Plan.Summary
        Phases          = $script:CommunicationPhases
    }
    $planData | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonFile -Encoding utf8

    Write-Log -Message "All communication files exported to $OutputPath" -Level Success
}

function New-CommunicationPlanSummary {
    <#
    .SYNOPSIS
        Creates an HTML summary/index page for the communication plan
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Plan,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    $emailLinks = ""
    foreach ($phase in $script:CommunicationPhases.Keys | Sort-Object { $script:CommunicationPhases[$_].Timing }) {
        $phaseInfo = $script:CommunicationPhases[$phase]
        $template = $Plan.EmailTemplates[$phase]
        $emailLinks += @"
        <tr>
            <td>$($phaseInfo.Name)</td>
            <td>$($template.ScheduledDate.ToString("MMM d, yyyy"))</td>
            <td>
                <a href="Emails/$phase-email.html">HTML</a> |
                <a href="Emails/$phase-email.txt">Plain Text</a>
            </td>
        </tr>
"@
    }

    $spoLinks = ""
    foreach ($phase in $script:CommunicationPhases.Keys | Sort-Object { $script:CommunicationPhases[$_].Timing }) {
        $phaseInfo = $script:CommunicationPhases[$phase]
        $post = $Plan.SharePointPosts[$phase]
        $featured = if ($post.IsFeatured) { "<span class='featured-badge'>Featured</span>" } else { "" }
        $spoLinks += @"
        <tr>
            <td>$($phaseInfo.Name) $featured</td>
            <td>$($post.ScheduledDate.ToString("MMM d, yyyy"))</td>
            <td>
                <a href="SharePointPosts/$phase-news.html">View Content</a>
            </td>
        </tr>
"@
    }

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Communication Plan Summary - $($Plan.Branding.CompanyName)</title>
    <style>
        :root {
            --bg-primary: #0f172a;
            --bg-secondary: #1e293b;
            --bg-tertiary: #334155;
            --bg-glass: rgba(30, 41, 59, 0.8);
            --text-primary: #f1f5f9;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --border-glass: rgba(148, 163, 184, 0.1);
            --accent-cyan: #06b6d4;
            --accent-purple: #8b5cf6;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, var(--bg-primary) 0%, #0c1929 50%, var(--bg-primary) 100%);
            min-height: 100vh;
            color: var(--text-primary);
        }
        body::before {
            content: '';
            position: fixed;
            top: 0; left: 0; width: 100%; height: 100%;
            background: radial-gradient(ellipse at 25% 25%, rgba(6, 182, 212, 0.08) 0%, transparent 50%),
                        radial-gradient(ellipse at 75% 75%, rgba(139, 92, 246, 0.08) 0%, transparent 50%);
            pointer-events: none;
        }
        .container { max-width: 1200px; margin: 0 auto; padding: 40px 20px; position: relative; z-index: 1; }
        .header {
            background: linear-gradient(135deg, rgba(6, 182, 212, 0.2), rgba(139, 92, 246, 0.2));
            backdrop-filter: blur(20px);
            border: 1px solid var(--border-glass);
            padding: 50px;
            border-radius: 24px;
            margin-bottom: 40px;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3), 0 0 60px rgba(6, 182, 212, 0.1);
        }
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 12px;
            background: linear-gradient(135deg, var(--text-primary), var(--accent-cyan));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        .header p { color: var(--text-secondary); margin-top: 8px; }
        .header .target-date {
            display: inline-block;
            background: linear-gradient(135deg, var(--accent-cyan), var(--accent-purple));
            padding: 12px 24px;
            border-radius: 30px;
            margin-top: 20px;
            font-size: 1.1em;
            font-weight: 600;
            box-shadow: 0 4px 20px rgba(6, 182, 212, 0.4);
        }
        .stats {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
            margin-bottom: 40px;
        }
        .stat {
            background: var(--bg-glass);
            backdrop-filter: blur(20px);
            border: 1px solid var(--border-glass);
            padding: 30px;
            border-radius: 16px;
            text-align: center;
            box-shadow: 8px 8px 16px rgba(0, 0, 0, 0.3), -4px -4px 12px rgba(255, 255, 255, 0.02);
            transition: all 0.3s ease;
        }
        .stat:hover {
            transform: translateY(-4px);
            box-shadow: 8px 8px 16px rgba(0, 0, 0, 0.3), -4px -4px 12px rgba(255, 255, 255, 0.02), 0 0 30px rgba(6, 182, 212, 0.2);
        }
        .stat-value {
            font-size: 3em;
            font-weight: 700;
            background: linear-gradient(135deg, var(--accent-cyan), var(--accent-purple));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        .stat-label {
            font-size: 0.85em;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-top: 8px;
        }
        .section {
            background: var(--bg-glass);
            backdrop-filter: blur(20px);
            border: 1px solid var(--border-glass);
            border-radius: 20px;
            padding: 35px;
            margin-bottom: 28px;
            box-shadow: 8px 8px 16px rgba(0, 0, 0, 0.3), -4px -4px 12px rgba(255, 255, 255, 0.02);
        }
        .section h2 {
            color: var(--text-primary);
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 2px solid rgba(6, 182, 212, 0.3);
            font-size: 1.4em;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .section h2::before {
            content: '';
            width: 4px;
            height: 24px;
            background: linear-gradient(180deg, var(--accent-cyan), var(--accent-purple));
            border-radius: 2px;
        }
        table { width: 100%; border-collapse: separate; border-spacing: 0; }
        th {
            text-align: left;
            padding: 16px;
            background: var(--bg-tertiary);
            color: var(--accent-cyan);
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.85em;
            letter-spacing: 0.5px;
        }
        th:first-child { border-radius: 10px 0 0 0; }
        th:last-child { border-radius: 0 10px 0 0; }
        td {
            padding: 16px;
            background: rgba(51, 65, 85, 0.3);
            border-bottom: 1px solid var(--border-glass);
            color: var(--text-secondary);
        }
        tr:hover td {
            background: rgba(6, 182, 212, 0.1);
            color: var(--text-primary);
        }
        a {
            color: var(--accent-cyan);
            text-decoration: none;
            transition: all 0.2s ease;
        }
        a:hover {
            color: var(--accent-purple);
            text-shadow: 0 0 10px var(--accent-cyan);
        }
        .resource-list { list-style: none; padding: 0; }
        .resource-list li {
            padding: 20px 0;
            border-bottom: 1px solid var(--border-glass);
        }
        .resource-list li:last-child { border-bottom: none; }
        .resource-list a { font-size: 1.1em; font-weight: 500; }
        .resource-list p { color: var(--text-muted); font-size: 0.9em; margin-top: 6px; }
        .featured-badge {
            background: linear-gradient(135deg, #10b981, #34d399);
            color: white;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 0.75em;
            margin-left: 10px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        ::-webkit-scrollbar { width: 10px; }
        ::-webkit-scrollbar-track { background: var(--bg-secondary); }
        ::-webkit-scrollbar-thumb { background: var(--bg-tertiary); border-radius: 5px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Communication Plan Summary</h1>
            <p>Microsoft 365 Migration - $($Plan.Branding.CompanyName)</p>
            <div class="target-date">Target Migration: $($Plan.TargetMigration.ToString("MMMM d, yyyy"))</div>
        </div>

        <div class="stats">
            <div class="stat">
                <div class="stat-value">$($Plan.Summary.UserCount)</div>
                <div class="stat-label">Users</div>
            </div>
            <div class="stat">
                <div class="stat-value">$($Plan.Summary.MailboxCount)</div>
                <div class="stat-label">Mailboxes</div>
            </div>
            <div class="stat">
                <div class="stat-value">$($Plan.Summary.TeamsCount)</div>
                <div class="stat-label">Teams</div>
            </div>
            <div class="stat">
                <div class="stat-value">7</div>
                <div class="stat-label">Communications</div>
            </div>
        </div>

        <div class="section">
            <h2>Email Templates</h2>
            <table>
                <thead>
                    <tr>
                        <th>Phase</th>
                        <th>Scheduled Date</th>
                        <th>Downloads</th>
                    </tr>
                </thead>
                <tbody>
                    $emailLinks
                </tbody>
            </table>
        </div>

        <div class="section">
            <h2>SharePoint News Posts</h2>
            <table>
                <thead>
                    <tr>
                        <th>Phase</th>
                        <th>Scheduled Date</th>
                        <th>Content</th>
                    </tr>
                </thead>
                <tbody>
                    $spoLinks
                </tbody>
            </table>
        </div>

        <div class="section">
            <h2>Additional Resources</h2>
            <ul class="resource-list">
                <li>
                    <a href="Timeline/migration-timeline.html">Migration Timeline (Visual)</a>
                    <p>Interactive Gantt chart showing all migration phases and milestones</p>
                </li>
                <li>
                    <a href="Resources/support-resources.html">Support Resources Page</a>
                    <p>FAQs, checklists, and help resources for end users</p>
                </li>
                <li>
                    <a href="communication-plan.json">Plan Data (JSON)</a>
                    <p>Machine-readable plan data for automation</p>
                </li>
            </ul>
        </div>
    </div>
</body>
</html>
"@
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'New-CommunicationPlan',
    'New-EmailTemplate',
    'New-SharePointNewsPost',
    'New-MigrationTimeline',
    'New-SupportResourcesPage',
    'Export-CommunicationPlan'
)
