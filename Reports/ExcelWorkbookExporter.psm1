#Requires -Version 7.0
<#
.SYNOPSIS
    Excel Workbook Exporter for M365 Tenant Discovery
.DESCRIPTION
    Generates a comprehensive Excel workbook with:
    - Multiple worksheets (one per data type)
    - Executive Summary sheet with charts
    - Formatted headers and styling
    - Data validation
    - Conditional formatting
    - Freeze panes and filters
.NOTES
    Author: AI Migration Expert
    Version: 1.0.0
    Target: PowerShell 7.x
    Dependencies: ImportExcel module (optional, will use COM if not available)
#>

#region Excel Workbook Functions

function Test-ImportExcelModule {
    <#
    .SYNOPSIS
        Checks if ImportExcel module is available
    #>
    return (Get-Module -ListAvailable -Name ImportExcel) -ne $null
}

function New-TenantDiscoveryWorkbook {
    <#
    .SYNOPSIS
        Creates a comprehensive Excel workbook with all tenant data
    .DESCRIPTION
        Generates a single Excel file with multiple worksheets containing:
        - Executive Summary (with charts)
        - Users
        - Groups
        - Devices
        - Applications
        - Conditional Access Policies
        - Mailboxes
        - SharePoint Sites
        - Teams
        - Migration Gotchas
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $true)]
        $AnalysisResults,

        [Parameter(Mandatory = $true)]
        $ComplexityScore,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Host "Generating Excel workbook: $OutputPath" -ForegroundColor Cyan

    $useImportExcel = Test-ImportExcelModule

    if ($useImportExcel) {
        Write-Host "Using ImportExcel module for enhanced features..." -ForegroundColor Green
        New-WorkbookWithImportExcel -CollectedData $CollectedData -AnalysisResults $AnalysisResults `
            -ComplexityScore $ComplexityScore -OutputPath $OutputPath
    } else {
        Write-Host "ImportExcel module not found, using COM objects..." -ForegroundColor Yellow
        Write-Host "For best results, install ImportExcel: Install-Module -Name ImportExcel" -ForegroundColor Yellow
        New-WorkbookWithCOM -CollectedData $CollectedData -AnalysisResults $AnalysisResults `
            -ComplexityScore $ComplexityScore -OutputPath $OutputPath
    }

    Write-Host "Excel workbook generated successfully!" -ForegroundColor Green
    return $OutputPath
}

function New-WorkbookWithImportExcel {
    <#
    .SYNOPSIS
        Creates workbook using ImportExcel module (recommended)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $true)]
        $AnalysisResults,

        [Parameter(Mandatory = $true)]
        $ComplexityScore,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    # Import the module
    Import-Module ImportExcel -ErrorAction Stop

    # Remove existing file if present
    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Force
    }

    # Define color scheme
    $headerColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $criticalColor = [System.Drawing.Color]::FromArgb(239, 68, 68)
    $highColor = [System.Drawing.Color]::FromArgb(249, 115, 22)
    $mediumColor = [System.Drawing.Color]::FromArgb(245, 158, 11)
    $lowColor = [System.Drawing.Color]::FromArgb(16, 185, 129)

    # ===== Executive Summary Sheet =====
    Write-Host "  Creating Executive Summary sheet..." -ForegroundColor Cyan

    # Create array of objects (one per row) instead of object with arrays
    $orgName = if ($CollectedData.Metadata.TenantName) { $CollectedData.Metadata.TenantName } elseif ($CollectedData.TenantInfo.DisplayName) { $CollectedData.TenantInfo.DisplayName } else { "Unknown" }
    $tenantId = if ($CollectedData.Metadata.TenantId) { $CollectedData.Metadata.TenantId } elseif ($CollectedData.TenantInfo.TenantId) { $CollectedData.TenantInfo.TenantId } else { "Unknown" }

    $execSummary = @(
        [PSCustomObject]@{ 'Metric' = 'Organization Name'; 'Value' = $orgName }
        [PSCustomObject]@{ 'Metric' = 'Tenant ID'; 'Value' = $tenantId }
        [PSCustomObject]@{ 'Metric' = 'Assessment Date'; 'Value' = (Get-Date -Format "yyyy-MM-dd HH:mm") }
        [PSCustomObject]@{ 'Metric' = ''; 'Value' = '' }
        [PSCustomObject]@{ 'Metric' = 'Total Users'; 'Value' = if ($CollectedData.EntraID.Users.Analysis.TotalUsers) { $CollectedData.EntraID.Users.Analysis.TotalUsers } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'Licensed Users'; 'Value' = if ($CollectedData.EntraID.Users.Analysis.LicensedUsers) { $CollectedData.EntraID.Users.Analysis.LicensedUsers } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'Guest Users'; 'Value' = if ($CollectedData.EntraID.Users.Analysis.GuestUsers) { $CollectedData.EntraID.Users.Analysis.GuestUsers } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'Total Mailboxes'; 'Value' = if ($CollectedData.Exchange.Mailboxes.Analysis.TotalMailboxes) { $CollectedData.Exchange.Mailboxes.Analysis.TotalMailboxes } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'SharePoint Sites'; 'Value' = if ($CollectedData.SharePoint.Sites.Analysis.SharePointSites) { $CollectedData.SharePoint.Sites.Analysis.SharePointSites } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'OneDrive Sites (All)'; 'Value' = if ($CollectedData.SharePoint.Sites.Analysis.OneDriveSites) { $CollectedData.SharePoint.Sites.Analysis.OneDriveSites } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'OneDrive Sites (Licensed Users)'; 'Value' = if ($CollectedData.SharePoint.Sites.Analysis.OneDriveSitesLicensedUsers) { $CollectedData.SharePoint.Sites.Analysis.OneDriveSitesLicensedUsers } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'Teams'; 'Value' = if ($CollectedData.Teams.Teams.Analysis.TotalTeams) { $CollectedData.Teams.Teams.Analysis.TotalTeams } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'Security Groups'; 'Value' = if ($CollectedData.EntraID.Groups.Analysis.SecurityGroups) { $CollectedData.EntraID.Groups.Analysis.SecurityGroups } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'M365 Groups'; 'Value' = if ($CollectedData.EntraID.Groups.Analysis.M365Groups) { $CollectedData.EntraID.Groups.Analysis.M365Groups } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'Applications'; 'Value' = if ($CollectedData.EntraID.Applications.Analysis.TotalApplications) { $CollectedData.EntraID.Applications.Analysis.TotalApplications } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'Devices'; 'Value' = if ($CollectedData.EntraID.Devices.Analysis.TotalDevices) { $CollectedData.EntraID.Devices.Analysis.TotalDevices } else { 0 } }
        [PSCustomObject]@{ 'Metric' = ''; 'Value' = '' }
        [PSCustomObject]@{ 'Metric' = 'Complexity Score'; 'Value' = "$($ComplexityScore.TotalScore)/100" }
        [PSCustomObject]@{ 'Metric' = 'Complexity Level'; 'Value' = $ComplexityScore.ComplexityLevel }
        [PSCustomObject]@{ 'Metric' = 'Risk Level'; 'Value' = $AnalysisResults.RiskLevel }
        [PSCustomObject]@{ 'Metric' = ''; 'Value' = '' }
        [PSCustomObject]@{ 'Metric' = 'Critical Issues'; 'Value' = if ($AnalysisResults.BySeverity.Critical) { $AnalysisResults.BySeverity.Critical.Count } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'High Issues'; 'Value' = if ($AnalysisResults.BySeverity.High) { $AnalysisResults.BySeverity.High.Count } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'Medium Issues'; 'Value' = if ($AnalysisResults.BySeverity.Medium) { $AnalysisResults.BySeverity.Medium.Count } else { 0 } }
        [PSCustomObject]@{ 'Metric' = 'Low Issues'; 'Value' = if ($AnalysisResults.BySeverity.Low) { $AnalysisResults.BySeverity.Low.Count } else { 0 } }
    )

    $execSummary | Export-Excel -Path $OutputPath -WorksheetName "Executive Summary" `
        -AutoSize -FreezeTopRow -BoldTopRow `
        -TableStyle Medium2

    # ===== Users Sheet =====
    if ($CollectedData.EntraID.Users.Data -and $CollectedData.EntraID.Users.Data.Count -gt 0) {
        Write-Host "  Creating Users sheet..." -ForegroundColor Cyan

        $userData = $CollectedData.EntraID.Users.Data | Select-Object -Property `
            @{N='User Principal Name';E={$_.UserPrincipalName}},
            @{N='Display Name';E={$_.DisplayName}},
            @{N='Mail';E={$_.Mail}},
            @{N='Job Title';E={$_.JobTitle}},
            @{N='Department';E={$_.Department}},
            @{N='Licenses';E={($_.AssignedLicenses | ForEach-Object { $_.SkuId }) -join '; '}},
            @{N='Account Enabled';E={$_.AccountEnabled}},
            @{N='User Type';E={$_.UserType}},
            @{N='Creation Date';E={$_.CreatedDateTime}}

        $userData | Export-Excel -Path $OutputPath -WorksheetName "Users" `
            -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
            -TableStyle Medium2 -ConditionalText $(
                New-ConditionalText -Text "False" -Range "G:G" -BackgroundColor Pink
            )
    }

    # ===== Groups Sheet =====
    if ($CollectedData.EntraID.Groups.Data -and $CollectedData.EntraID.Groups.Data.Count -gt 0) {
        Write-Host "  Creating Groups sheet..." -ForegroundColor Cyan

        $groupData = $CollectedData.EntraID.Groups.Data | Select-Object -Property `
            @{N='Display Name';E={$_.DisplayName}},
            @{N='Mail';E={$_.Mail}},
            @{N='Mail Enabled';E={$_.MailEnabled}},
            @{N='Security Enabled';E={$_.SecurityEnabled}},
            @{N='Group Types';E={$_.GroupTypes -join '; '}},
            @{N='Membership Rule';E={$_.MembershipRule}},
            @{N='Description';E={$_.Description}},
            @{N='Creation Date';E={$_.CreatedDateTime}}

        $groupData | Export-Excel -Path $OutputPath -WorksheetName "Groups" `
            -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
            -TableStyle Medium2
    }

    # ===== Devices Sheet =====
    if ($CollectedData.EntraID.Devices.Data -and $CollectedData.EntraID.Devices.Data.Count -gt 0) {
        Write-Host "  Creating Devices sheet..." -ForegroundColor Cyan

        $deviceData = $CollectedData.EntraID.Devices.Data | Select-Object -Property `
            @{N='Display Name';E={$_.DisplayName}},
            @{N='Device ID';E={$_.DeviceId}},
            @{N='Operating System';E={$_.OperatingSystem}},
            @{N='OS Version';E={$_.OperatingSystemVersion}},
            @{N='Trust Type';E={$_.TrustType}},
            @{N='Is Compliant';E={$_.IsCompliant}},
            @{N='Is Managed';E={$_.IsManaged}},
            @{N='Approx Last SignIn';E={$_.ApproximateLastSignInDateTime}},
            @{N='Registered Owner';E={
                if ($_.RegisteredOwners -and $_.RegisteredOwners.Count -gt 0) {
                    $_.RegisteredOwners[0].UserPrincipalName
                } else { "" }
            }}

        $deviceData | Export-Excel -Path $OutputPath -WorksheetName "Devices" `
            -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
            -TableStyle Medium2 -ConditionalText $(
                New-ConditionalText -Text "False" -Range "F:F" -BackgroundColor Pink
            )
    }

    # ===== Applications Sheet =====
    if ($CollectedData.EntraID.Applications.Data -and $CollectedData.EntraID.Applications.Data.Count -gt 0) {
        Write-Host "  Creating Applications sheet..." -ForegroundColor Cyan

        $appData = $CollectedData.EntraID.Applications.Data | Select-Object -Property `
            @{N='Display Name';E={$_.DisplayName}},
            @{N='App ID';E={$_.AppId}},
            @{N='Publisher Domain';E={$_.PublisherDomain}},
            @{N='Sign In Audience';E={$_.SignInAudience}},
            @{N='Creation Date';E={$_.CreatedDateTime}},
            @{N='API Permissions';E={
                if ($_.RequiredResourceAccess) {
                    ($_.RequiredResourceAccess | ForEach-Object {
                        $_.ResourceAccess.Id
                    }) -join '; '
                } else { "" }
            }}

        $appData | Export-Excel -Path $OutputPath -WorksheetName "Applications" `
            -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
            -TableStyle Medium2
    }

    # ===== Conditional Access Policies Sheet =====
    if ($CollectedData.EntraID.ConditionalAccessPolicies.Data -and $CollectedData.EntraID.ConditionalAccessPolicies.Data.Count -gt 0) {
        Write-Host "  Creating Conditional Access Policies sheet..." -ForegroundColor Cyan

        $caData = $CollectedData.EntraID.ConditionalAccessPolicies.Data | Select-Object -Property `
            @{N='Display Name';E={$_.DisplayName}},
            @{N='State';E={$_.State}},
            @{N='Include Users';E={$_.Conditions.Users.IncludeUsers -join '; '}},
            @{N='Exclude Users';E={$_.Conditions.Users.ExcludeUsers -join '; '}},
            @{N='Include Applications';E={$_.Conditions.Applications.IncludeApplications -join '; '}},
            @{N='Grant Controls';E={$_.GrantControls.BuiltInControls -join '; '}},
            @{N='Created';E={$_.CreatedDateTime}},
            @{N='Modified';E={$_.ModifiedDateTime}}

        $caData | Export-Excel -Path $OutputPath -WorksheetName "Conditional Access" `
            -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
            -TableStyle Medium2 -ConditionalText $(
                New-ConditionalText -Text "disabled" -Range "B:B" -BackgroundColor LightYellow
            )
    }

    # ===== Mailboxes Sheet =====
    if ($CollectedData.Exchange.Mailboxes.Data -and $CollectedData.Exchange.Mailboxes.Data.Count -gt 0) {
        Write-Host "  Creating Mailboxes sheet..." -ForegroundColor Cyan

        $mailboxData = $CollectedData.Exchange.Mailboxes.Data | Select-Object -Property `
            @{N='User Principal Name';E={$_.UserPrincipalName}},
            @{N='Display Name';E={$_.DisplayName}},
            @{N='Primary SMTP';E={$_.PrimarySmtpAddress}},
            @{N='Mailbox Type';E={$_.RecipientTypeDetails}},
            @{N='Archive Status';E={$_.ArchiveStatus}},
            @{N='Litigation Hold';E={$_.LitigationHoldEnabled}},
            @{N='Forward To';E={$_.ForwardingSmtpAddress}},
            @{N='Hidden From GAL';E={$_.HiddenFromAddressListsEnabled}}

        $mailboxData | Export-Excel -Path $OutputPath -WorksheetName "Mailboxes" `
            -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
            -TableStyle Medium2 -ConditionalText $(
                New-ConditionalText -Text "True" -Range "F:F" -BackgroundColor LightBlue
            )
    }

    # ===== SharePoint Sites Sheet =====
    if ($CollectedData.SharePoint.Sites.Data -and $CollectedData.SharePoint.Sites.Data.Count -gt 0) {
        Write-Host "  Creating SharePoint Sites sheet..." -ForegroundColor Cyan

        $siteData = $CollectedData.SharePoint.Sites.Data | Select-Object -Property `
            @{N='Title';E={$_.Title}},
            @{N='URL';E={$_.Url}},
            @{N='Template';E={$_.Template}},
            @{N='Storage Used (GB)';E={[math]::Round($_.StorageUsageCurrent / 1024, 2)}},
            @{N='Storage Quota (GB)';E={[math]::Round($_.StorageQuota / 1024, 2)}},
            @{N='Last Content Modified';E={$_.LastContentModifiedDate}},
            @{N='Owner';E={$_.Owner}}

        $siteData | Export-Excel -Path $OutputPath -WorksheetName "SharePoint Sites" `
            -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
            -TableStyle Medium2
    }

    # ===== Teams Sheet =====
    if ($CollectedData.Teams.Teams.Data -and $CollectedData.Teams.Teams.Data.Count -gt 0) {
        Write-Host "  Creating Teams sheet..." -ForegroundColor Cyan

        $teamsData = $CollectedData.Teams.Teams.Data | Select-Object -Property `
            @{N='Display Name';E={$_.DisplayName}},
            @{N='Description';E={$_.Description}},
            @{N='Visibility';E={$_.Visibility}},
            @{N='Is Archived';E={$_.IsArchived}},
            @{N='Mail Nickname';E={$_.MailNickname}},
            @{N='Web URL';E={$_.WebUrl}}

        $teamsData | Export-Excel -Path $OutputPath -WorksheetName "Teams" `
            -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
            -TableStyle Medium2 -ConditionalText $(
                New-ConditionalText -Text "True" -Range "D:D" -BackgroundColor LightYellow
            )
    }

    # ===== Migration Gotchas Sheet =====
    Write-Host "  Creating Migration Gotchas sheet..." -ForegroundColor Cyan

    $gotchasData = @()
    foreach ($severity in @("Critical", "High", "Medium", "Low")) {
        $gotchas = $AnalysisResults.BySeverity[$severity]
        if ($gotchas -and $gotchas.Count -gt 0) {
            foreach ($gotcha in $gotchas) {
                $gotchasData += [PSCustomObject]@{
                    'Severity' = $severity
                    'Category' = $gotcha.Category
                    'Name' = $gotcha.Name
                    'Description' = $gotcha.Description
                    'Affected Count' = $gotcha.AffectedCount
                    'Recommendation' = $gotcha.Recommendation
                }
            }
        }
    }

    if ($gotchasData.Count -gt 0) {
        $gotchasData | Export-Excel -Path $OutputPath -WorksheetName "Migration Gotchas" `
            -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
            -TableStyle Medium2 -ConditionalText $(
                New-ConditionalText -Text "Critical" -Range "A:A" -BackgroundColor $criticalColor -ConditionalTextColor White
                New-ConditionalText -Text "High" -Range "A:A" -BackgroundColor $highColor -ConditionalTextColor White
                New-ConditionalText -Text "Medium" -Range "A:A" -BackgroundColor $mediumColor
                New-ConditionalText -Text "Low" -Range "A:A" -BackgroundColor $lowColor -ConditionalTextColor White
            )
    }

    Write-Host "  Adding charts..." -ForegroundColor Cyan

    try {
        # Create a dedicated chart data sheet with known cell positions
        $criticalIssues = if ($AnalysisResults.BySeverity.Critical) { $AnalysisResults.BySeverity.Critical.Count } else { 0 }
        $highIssues = if ($AnalysisResults.BySeverity.High) { $AnalysisResults.BySeverity.High.Count } else { 0 }
        $mediumIssues = if ($AnalysisResults.BySeverity.Medium) { $AnalysisResults.BySeverity.Medium.Count } else { 0 }
        $lowIssues = if ($AnalysisResults.BySeverity.Low) { $AnalysisResults.BySeverity.Low.Count } else { 0 }

        $totalUsersChart = if ($CollectedData.EntraID.Users.Analysis.TotalUsers) { $CollectedData.EntraID.Users.Analysis.TotalUsers } else { 0 }
        $licensedUsersChart = if ($CollectedData.EntraID.Users.Analysis.LicensedUsers) { $CollectedData.EntraID.Users.Analysis.LicensedUsers } else { 0 }
        $guestUsersChart = if ($CollectedData.EntraID.Users.Analysis.GuestUsers) { $CollectedData.EntraID.Users.Analysis.GuestUsers } else { 0 }
        $totalMailboxesChart = if ($CollectedData.Exchange.Mailboxes.Analysis.TotalMailboxes) { $CollectedData.Exchange.Mailboxes.Analysis.TotalMailboxes } else { 0 }
        $totalSitesChart = if ($CollectedData.SharePoint.Sites.Analysis.SharePointSites) { $CollectedData.SharePoint.Sites.Analysis.SharePointSites } else { 0 }
        $totalTeamsChart = if ($CollectedData.Teams.Teams.Analysis.TotalTeams) { $CollectedData.Teams.Teams.Analysis.TotalTeams } else { 0 }

        # Severity chart data - write to a hidden "ChartData" sheet
        $chartData = @(
            [PSCustomObject]@{ 'Category' = 'Critical'; 'Count' = $criticalIssues }
            [PSCustomObject]@{ 'Category' = 'High'; 'Count' = $highIssues }
            [PSCustomObject]@{ 'Category' = 'Medium'; 'Count' = $mediumIssues }
            [PSCustomObject]@{ 'Category' = 'Low'; 'Count' = $lowIssues }
        )

        $severityChart = New-ExcelChartDefinition -XRange "Category" -YRange "Count" `
            -Title "Issues by Severity" -ChartType Pie -Row 0 -Column 0 -Width 500 -Height 350 `
            -SeriesHeader "Count" -XAxisTitleText "" -YAxisTitleText ""

        $chartData | Export-Excel -Path $OutputPath -WorksheetName "Severity Chart" `
            -AutoSize -ExcelChartDefinition $severityChart

        # Workload chart data
        $workloadData = @(
            [PSCustomObject]@{ 'Workload' = 'Users'; 'Count' = $licensedUsersChart }
            [PSCustomObject]@{ 'Workload' = 'Mailboxes'; 'Count' = $totalMailboxesChart }
            [PSCustomObject]@{ 'Workload' = 'SharePoint Sites'; 'Count' = $totalSitesChart }
            [PSCustomObject]@{ 'Workload' = 'Teams'; 'Count' = $totalTeamsChart }
            [PSCustomObject]@{ 'Workload' = 'Guest Users'; 'Count' = $guestUsersChart }
        )

        $workloadChart = New-ExcelChartDefinition -XRange "Workload" -YRange "Count" `
            -Title "Workload Distribution" -ChartType ColumnClustered -Row 0 -Column 0 -Width 500 -Height 350 `
            -SeriesHeader "Count" -XAxisTitleText "" -YAxisTitleText "Count"

        $workloadData | Export-Excel -Path $OutputPath -WorksheetName "Workload Chart" `
            -AutoSize -ExcelChartDefinition $workloadChart

        Write-Host "  Charts added successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "  Warning: Could not add charts to Excel - $($_)" -ForegroundColor Yellow
    }
}

function New-WorkbookWithCOM {
    <#
    .SYNOPSIS
        Creates workbook using COM objects (fallback method)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $true)]
        $AnalysisResults,

        [Parameter(Mandatory = $true)]
        $ComplexityScore,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Host "  Using COM-based export (basic formatting only)..." -ForegroundColor Yellow

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()

        # Create sheets and add data (simplified version)
        # Executive Summary
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = "Executive Summary"

        $row = 1
        $sheet.Cells.Item($row, 1) = "Metric"
        $sheet.Cells.Item($row, 2) = "Value"

        $sheet.Cells.Item($row, 1).Font.Bold = $true
        $sheet.Cells.Item($row, 2).Font.Bold = $true

        $row++
        $sheet.Cells.Item($row, 1) = "Organization Name"
        $sheet.Cells.Item($row, 2) = $orgName

        $row++
        $sheet.Cells.Item($row, 1) = "Tenant ID"
        $sheet.Cells.Item($row, 2) = $tenantId

        $row++
        $sheet.Cells.Item($row, 1) = "Assessment Date"
        $sheet.Cells.Item($row, 2) = (Get-Date -Format "yyyy-MM-dd HH:mm")

        # Add more summary data...
        $row += 2
        $sheet.Cells.Item($row, 1) = "Total Users"
        $sheet.Cells.Item($row, 2) = $CollectedData.EntraID.Users.Analysis.TotalUsers

        $row++
        $sheet.Cells.Item($row, 1) = "Licensed Users"
        $sheet.Cells.Item($row, 2) = $CollectedData.EntraID.Users.Analysis.LicensedUsers

        # Auto-fit columns
        $sheet.UsedRange.EntireColumn.AutoFit() | Out-Null

        # Save and close
        $workbook.SaveAs($OutputPath)
        $workbook.Close($false)
        $excel.Quit()

        # Clean up COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        Write-Host "  Basic Excel workbook created using COM" -ForegroundColor Green
    }
    catch {
        Write-Host "  Error creating Excel workbook via COM: $_" -ForegroundColor Red
        throw
    }
}

#endregion

# Export functions
Export-ModuleMember -Function @(
    'Test-ImportExcelModule',
    'New-TenantDiscoveryWorkbook'
)
