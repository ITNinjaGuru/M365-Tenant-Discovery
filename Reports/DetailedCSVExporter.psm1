#Requires -Version 7.0
<#
.SYNOPSIS
    Detailed CSV Exporter for M365 Tenant Discovery Tool

.DESCRIPTION
    Exports comprehensive CSV files with detailed information about mailboxes,
    users, and other resources that may not be captured in the main reports.
    Designed for record-keeping and detailed analysis.

.NOTES
    Author: M365 Discovery Tool
    Version: 1.0.0
#>

function Export-DetailedMailboxCSV {
    <#
    .SYNOPSIS
        Exports detailed mailbox information to CSV
    .DESCRIPTION
        Creates a comprehensive CSV file with all mailbox details including:
        - Primary SMTP Address
        - All Email Addresses
        - UPN (UserPrincipalName)
        - Alias
        - Mailbox Statistics (TotalItemSize, ItemCount)
        - Retention Hold Status
        - Archive Mailbox Status
        - Archive Mailbox Size
        - Litigation Hold Status
        - Last User Action Time
        - Mailbox Type
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Host "  Exporting detailed mailbox CSV..." -ForegroundColor Cyan

    # Check if we have mailbox data
    if (-not $CollectedData.Exchange -or -not $CollectedData.Exchange.Mailboxes -or -not $CollectedData.Exchange.Mailboxes.Data) {
        Write-Host "    No mailbox data available. Skipping mailbox CSV export." -ForegroundColor Yellow
        return $null
    }

    $mailboxes = $CollectedData.Exchange.Mailboxes.Data

    if ($mailboxes.Count -eq 0) {
        Write-Host "    No mailboxes found. Skipping mailbox CSV export." -ForegroundColor Yellow
        return $null
    }

    # Build detailed mailbox records
    $detailedMailboxes = @()

    foreach ($mailbox in $mailboxes) {
        # Get mailbox statistics if available
        $stats = $CollectedData.Exchange.MailboxStatistics.Data | Where-Object { $_.Identity -eq $mailbox.Identity -or $_.DisplayName -eq $mailbox.DisplayName } | Select-Object -First 1

        # Build email addresses list (all addresses)
        $emailAddresses = if ($mailbox.EmailAddresses) {
            ($mailbox.EmailAddresses | Where-Object { $_ -like '*@*' }) -join '; '
        } else {
            ""
        }

        # Extract primary SMTP
        $primarySmtp = if ($mailbox.PrimarySmtpAddress) {
            $mailbox.PrimarySmtpAddress
        } elseif ($mailbox.EmailAddresses) {
            $primary = $mailbox.EmailAddresses | Where-Object { $_ -like 'SMTP:*' } | Select-Object -First 1
            if ($primary) { $primary -replace '^SMTP:', '' } else { "" }
        } else {
            ""
        }

        # Get archive info if available
        $archiveEnabled = if ($null -ne $mailbox.ArchiveStatus) {
            $mailbox.ArchiveStatus -ne "None"
        } else {
            $false
        }

        $archiveSize = ""
        if ($archiveEnabled -and $stats -and $stats.ArchiveTotalItemSize) {
            $archiveSize = $stats.ArchiveTotalItemSize
        }

        # Build record
        $record = [PSCustomObject]@{
            'Display Name'                = $mailbox.DisplayName
            'User Principal Name'         = $mailbox.UserPrincipalName
            'Primary SMTP Address'        = $primarySmtp
            'Alias'                       = $mailbox.Alias
            'Email Addresses (All)'       = $emailAddresses
            'Mailbox Type'                = $mailbox.RecipientTypeDetails
            'Total Item Size'             = if ($stats) { $stats.TotalItemSize } else { "" }
            'Item Count'                  = if ($stats) { $stats.ItemCount } else { "" }
            'Deleted Item Size'           = if ($stats) { $stats.TotalDeletedItemSize } else { "" }
            'Deleted Item Count'          = if ($stats) { $stats.DeletedItemCount } else { "" }
            'Last User Action Time'       = if ($stats) { $stats.LastUserActionTime } else { "" }
            'Litigation Hold Enabled'     = $mailbox.LitigationHoldEnabled
            'In-Place Holds'              = if ($mailbox.InPlaceHolds) { ($mailbox.InPlaceHolds -join '; ') } else { "" }
            'Retention Policy'            = $mailbox.RetentionPolicy
            'Retention Hold Enabled'      = $mailbox.RetentionHoldEnabled
            'Archive Enabled'             = $archiveEnabled
            'Archive Status'              = $mailbox.ArchiveStatus
            'Archive Size'                = $archiveSize
            'Archive Database'            = $mailbox.ArchiveDatabase
            'Archive Quota'               = $mailbox.ArchiveQuota
            'Archive Warning Quota'       = $mailbox.ArchiveWarningQuota
            'Prohibit Send Quota'         = $mailbox.ProhibitSendQuota
            'Prohibit Send Receive Quota' = $mailbox.ProhibitSendReceiveQuota
            'Issue Warning Quota'         = $mailbox.IssueWarningQuota
            'Hidden From Address Lists'   = $mailbox.HiddenFromAddressListsEnabled
            'When Created'                = $mailbox.WhenCreated
            'When Changed'                = $mailbox.WhenChanged
            'Mailbox Database'            = $mailbox.Database
        }

        $detailedMailboxes += $record
    }

    # Export to CSV
    $csvPath = Join-Path -Path $OutputPath -ChildPath "DetailedMailboxes.csv"
    $detailedMailboxes | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host "    Detailed mailbox CSV exported: $csvPath" -ForegroundColor Green
    Write-Host "    Total mailboxes exported: $($detailedMailboxes.Count)" -ForegroundColor Gray

    return $csvPath
}

function Export-DetailedUserCSV {
    <#
    .SYNOPSIS
        Exports detailed user information to CSV
    .DESCRIPTION
        Creates a comprehensive CSV file with all user details including:
        - UPN, Display Name, Mail
        - All assigned licenses
        - Account status
        - Sign-in activity
        - MFA status
        - Group memberships
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Host "  Exporting detailed user CSV..." -ForegroundColor Cyan

    # Check if we have user data
    if (-not $CollectedData.EntraID -or -not $CollectedData.EntraID.Users -or -not $CollectedData.EntraID.Users.Data) {
        Write-Host "    No user data available. Skipping user CSV export." -ForegroundColor Yellow
        return $null
    }

    $users = $CollectedData.EntraID.Users.Data

    if ($users.Count -eq 0) {
        Write-Host "    No users found. Skipping user CSV export." -ForegroundColor Yellow
        return $null
    }

    # Build detailed user records
    $detailedUsers = @()

    foreach ($user in $users) {
        # Get assigned licenses
        $licenses = if ($user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) {
            ($user.AssignedLicenses | ForEach-Object { $_.SkuId }) -join '; '
        } else {
            ""
        }

        # Build record
        $record = [PSCustomObject]@{
            'Display Name'              = $user.DisplayName
            'User Principal Name'       = $user.UserPrincipalName
            'Mail'                      = $user.Mail
            'User Type'                 = $user.UserType
            'Account Enabled'           = $user.AccountEnabled
            'Created DateTime'          = $user.CreatedDateTime
            'Job Title'                 = $user.JobTitle
            'Department'                = $user.Department
            'Office Location'           = $user.OfficeLocation
            'City'                      = $user.City
            'State'                     = $user.State
            'Country'                   = $user.Country
            'Phone'                     = $user.BusinessPhones -join '; '
            'Mobile Phone'              = $user.MobilePhone
            'Assigned Licenses (SKUs)'  = $licenses
            'License Count'             = if ($user.AssignedLicenses) { $user.AssignedLicenses.Count } else { 0 }
            'Proxy Addresses'           = if ($user.ProxyAddresses) { ($user.ProxyAddresses -join '; ') } else { "" }
            'On-Premises Sync Enabled'  = $user.OnPremisesSyncEnabled
            'Last Password Change'      = $user.LastPasswordChangeDateTime
            'Password Policies'         = $user.PasswordPolicies
        }

        $detailedUsers += $record
    }

    # Export to CSV
    $csvPath = Join-Path -Path $OutputPath -ChildPath "DetailedUsers.csv"
    $detailedUsers | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host "    Detailed user CSV exported: $csvPath" -ForegroundColor Green
    Write-Host "    Total users exported: $($detailedUsers.Count)" -ForegroundColor Gray

    return $csvPath
}

function Export-DetailedGroupCSV {
    <#
    .SYNOPSIS
        Exports detailed group information to CSV
    .DESCRIPTION
        Creates a comprehensive CSV file with all group details including membership counts
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Host "  Exporting detailed group CSV..." -ForegroundColor Cyan

    # Check if we have group data
    if (-not $CollectedData.EntraID -or -not $CollectedData.EntraID.Groups -or -not $CollectedData.EntraID.Groups.Data) {
        Write-Host "    No group data available. Skipping group CSV export." -ForegroundColor Yellow
        return $null
    }

    $groups = $CollectedData.EntraID.Groups.Data

    if ($groups.Count -eq 0) {
        Write-Host "    No groups found. Skipping group CSV export." -ForegroundColor Yellow
        return $null
    }

    # Build detailed group records
    $detailedGroups = @()

    foreach ($group in $groups) {
        # Build record
        $record = [PSCustomObject]@{
            'Display Name'         = $group.DisplayName
            'Mail'                 = $group.Mail
            'Mail Enabled'         = $group.MailEnabled
            'Mail Nickname'        = $group.MailNickname
            'Group Type'           = $group.GroupTypes -join '; '
            'Security Enabled'     = $group.SecurityEnabled
            'Description'          = $group.Description
            'Created DateTime'     = $group.CreatedDateTime
            'Visibility'           = $group.Visibility
            'Member Count'         = if ($group.Members) { $group.Members.Count } else { 0 }
            'Owner Count'          = if ($group.Owners) { $group.Owners.Count } else { 0 }
            'On-Premises Sync'     = $group.OnPremisesSyncEnabled
            'Proxy Addresses'      = if ($group.ProxyAddresses) { ($group.ProxyAddresses -join '; ') } else { "" }
        }

        $detailedGroups += $record
    }

    # Export to CSV
    $csvPath = Join-Path -Path $OutputPath -ChildPath "DetailedGroups.csv"
    $detailedGroups | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host "    Detailed group CSV exported: $csvPath" -ForegroundColor Green
    Write-Host "    Total groups exported: $($detailedGroups.Count)" -ForegroundColor Gray

    return $csvPath
}

function Export-AllDetailedCSVs {
    <#
    .SYNOPSIS
        Exports all detailed CSV files
    .DESCRIPTION
        Main function that exports all available detailed CSV files
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$CollectedData,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Host "`nExporting Detailed CSV Files..." -ForegroundColor Cyan
    Write-Host "=" * 80 -ForegroundColor Gray

    $exportedFiles = @()

    # Create CSV subdirectory
    $csvOutputPath = Join-Path -Path $OutputPath -ChildPath "DetailedCSV"
    if (-not (Test-Path -Path $csvOutputPath)) {
        New-Item -Path $csvOutputPath -ItemType Directory -Force | Out-Null
    }

    # Export mailboxes
    $mailboxCsv = Export-DetailedMailboxCSV -CollectedData $CollectedData -OutputPath $csvOutputPath
    if ($mailboxCsv) {
        $exportedFiles += $mailboxCsv
    }

    # Export users
    $userCsv = Export-DetailedUserCSV -CollectedData $CollectedData -OutputPath $csvOutputPath
    if ($userCsv) {
        $exportedFiles += $userCsv
    }

    # Export groups
    $groupCsv = Export-DetailedGroupCSV -CollectedData $CollectedData -OutputPath $csvOutputPath
    if ($groupCsv) {
        $exportedFiles += $groupCsv
    }

    Write-Host "`n  ✓ Detailed CSV export complete!" -ForegroundColor Green
    Write-Host "  Total files exported: $($exportedFiles.Count)" -ForegroundColor Gray
    Write-Host "  Location: $csvOutputPath" -ForegroundColor Gray

    return $exportedFiles
}

# Export module members
Export-ModuleMember -Function @(
    'Export-DetailedMailboxCSV',
    'Export-DetailedUserCSV',
    'Export-DetailedGroupCSV',
    'Export-AllDetailedCSVs'
)
