#Requires -Version 7.0
<#
.SYNOPSIS
    Migration Script Runner - Executes AI-generated migration scripts with proper isolation
.DESCRIPTION
    Runs each migration script in an isolated PowerShell process to avoid module version conflicts.
    This prevents "assembly already loaded" errors when running multiple scripts in sequence.
.PARAMETER ScriptsPath
    Path to the Scripts folder containing the generated migration scripts
.PARAMETER ExecutionMode
    'Sequential' (default) - Run scripts one at a time, waiting for each to complete
    'Parallel' - Run scripts in parallel (use with caution - may cause resource contention)
.PARAMETER LogPath
    Optional path to save execution logs
.EXAMPLE
    .\Run-MigrationScripts.ps1 -ScriptsPath ".\Scripts"

.EXAMPLE
    .\Run-MigrationScripts.ps1 -ScriptsPath ".\Scripts" -LogPath ".\migration-log.txt"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$ScriptsPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Sequential", "Parallel")]
    [string]$ExecutionMode = "Sequential",

    [Parameter(Mandatory = $false)]
    [string]$LogPath
)

Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  M365 Migration Script Runner" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""

# Get all .ps1 scripts in order (excluding this runner and playbook reference)
$scripts = Get-ChildItem -Path $ScriptsPath -Filter "*.ps1" -File |
    Where-Object { $_.Name -notlike "*Run-*" -and $_.Name -notlike "*Migration-Playbook*" } |
    Sort-Object Name

if ($scripts.Count -eq 0) {
    Write-Host "No migration scripts found in: $ScriptsPath" -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $($scripts.Count) migration script(s) to execute:" -ForegroundColor White
$scripts | ForEach-Object { Write-Host "  • $($_.Name)" -ForegroundColor Gray }
Write-Host ""

# Prepare log file if specified
$logContent = @()
if ($LogPath) {
    $logContent += "═══════════════════════════════════════════════════════════════════════"
    $logContent += "  M365 Migration Scripts Execution Log"
    $logContent += "  Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $logContent += "═══════════════════════════════════════════════════════════════════════"
    $logContent += ""
}

$successCount = 0
$failureCount = 0
$results = @()

# Execute scripts based on mode
if ($ExecutionMode -eq "Sequential") {
    Write-Host "Execution Mode: SEQUENTIAL (each script runs in isolated process)" -ForegroundColor Cyan
    Write-Host ""

    foreach ($script in $scripts) {
        $scriptPath = $script.FullName
        $scriptName = $script.Name

        Write-Host "▶ Running: $scriptName" -ForegroundColor Yellow

        try {
            # Run script in isolated PowerShell process to avoid module conflicts
            $output = & pwsh -NoProfile -ExecutionPolicy Bypass -File $scriptPath 2>&1

            if ($LASTEXITCODE -eq 0) {
                Write-Host "  ✓ Completed successfully" -ForegroundColor Green
                $successCount++
                $status = "SUCCESS"
            } else {
                Write-Host "  ✗ Failed with exit code: $LASTEXITCODE" -ForegroundColor Red
                if ($output) {
                    $output | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
                }
                $failureCount++
                $status = "FAILED"
            }
        }
        catch {
            Write-Host "  ✗ Error executing script: $_" -ForegroundColor Red
            $failureCount++
            $status = "ERROR"
        }

        $results += @{
            Script = $scriptName
            Status = $status
            Time   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }

        if ($LogPath) {
            $logContent += "Script: $scriptName"
            $logContent += "Status: $status"
            $logContent += "Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            $logContent += ""
        }

        Write-Host ""
    }
}
else {
    Write-Host "Execution Mode: PARALLEL (scripts run concurrently - use with caution)" -ForegroundColor Cyan
    Write-Host ""

    $jobs = @()
    foreach ($script in $scripts) {
        $scriptPath = $script.FullName
        $scriptName = $script.Name

        Write-Host "▶ Queuing: $scriptName" -ForegroundColor Yellow

        $job = Start-Job -ScriptBlock {
            param($scriptPath)
            & pwsh -NoProfile -ExecutionPolicy Bypass -File $scriptPath
        } -ArgumentList $scriptPath -Name $scriptName

        $jobs += $job
    }

    Write-Host ""
    Write-Host "Waiting for all scripts to complete..." -ForegroundColor Cyan
    Write-Host ""

    foreach ($job in $jobs) {
        $jobResult = Receive-Job -Job $job -Wait -AutoRemoveJob
        $jobName = $job.Name

        if ($job.State -eq "Completed" -and $job.Error.Count -eq 0) {
            Write-Host "✓ $jobName: Completed" -ForegroundColor Green
            $successCount++
            $status = "SUCCESS"
        } else {
            Write-Host "✗ $jobName: Failed" -ForegroundColor Red
            if ($jobResult) {
                $jobResult | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
            }
            $failureCount++
            $status = "FAILED"
        }

        $results += @{
            Script = $jobName
            Status = $status
            Time   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }

        if ($LogPath) {
            $logContent += "Script: $jobName"
            $logContent += "Status: $status"
            $logContent += "Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            $logContent += ""
        }
    }
}

# Summary
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  Execution Summary" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""
Write-Host "Total Scripts:  $($scripts.Count)" -ForegroundColor White
Write-Host "Successful:     $successCount" -ForegroundColor Green
Write-Host "Failed:         $failureCount" -ForegroundColor $(if ($failureCount -gt 0) { "Red" } else { "Green" })
Write-Host ""

if ($results) {
    Write-Host "Details:" -ForegroundColor White
    $results | ForEach-Object {
        $color = if ($_.Status -eq "SUCCESS") { "Green" } else { "Red" }
        Write-Host "  [$($_.Status)] $($_.Script)" -ForegroundColor $color
    }
    Write-Host ""
}

if ($LogPath) {
    $logContent += ""
    $logContent += "═══════════════════════════════════════════════════════════════════════"
    $logContent += "  Summary"
    $logContent += "  Total: $($scripts.Count) | Success: $successCount | Failed: $failureCount"
    $logContent += "  Completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $logContent += "═══════════════════════════════════════════════════════════════════════"

    $logContent | Out-File -FilePath $LogPath -Encoding UTF8
    Write-Host "Log saved to: $LogPath" -ForegroundColor Gray
    Write-Host ""
}

# Exit with appropriate code
exit $(if ($failureCount -gt 0) { 1 } else { 0 })
