#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Troubleshoots why Windows spontaneously rebooted.
.DESCRIPTION
    Analyzes event logs to identify the cause of unexpected system reboots,
    including BSODs, power failures, Windows Updates, and hardware errors.
    Can analyze events leading up to crashes and install monitoring tools.
.PARAMETER Days
    Number of days to look back in event logs. Default is 7.
.PARAMETER InstallMonitoring
    Prompts to install monitoring tools for tracking system state.
.PARAMETER AnalyzePreCrash
    Minutes before each crash to analyze for warning signs. Default is 10.
.EXAMPLE
    .\Get-RebootReason.ps1
    .\Get-RebootReason.ps1 -Days 30 -AnalyzePreCrash 15
    .\Get-RebootReason.ps1 -InstallMonitoring
#>

param(
    [int]$Days = 7,
    [switch]$InstallMonitoring,
    [int]$AnalyzePreCrash = 10
)

$ErrorActionPreference = "SilentlyContinue"
$StartDate = (Get-Date).AddDays(-$Days)

Write-Host "`n===== WINDOWS REBOOT TROUBLESHOOTER =====" -ForegroundColor Cyan
Write-Host "Analyzing events from the last $Days days...`n" -ForegroundColor Gray

# --- 1. Unexpected Shutdowns (Dirty Reboots) ---
Write-Host "[1] UNEXPECTED SHUTDOWNS (Event ID 6008)" -ForegroundColor Yellow
$unexpectedShutdowns = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    Id = 6008
    StartTime = $StartDate
} -MaxEvents 10 2>$null

if ($unexpectedShutdowns) {
    foreach ($event in $unexpectedShutdowns) {
        $msg = $event.Message -replace "`r`n", " "
        Write-Host "  $($event.TimeCreated): $msg" -ForegroundColor Red
    }
} else {
    Write-Host "  No unexpected shutdowns found." -ForegroundColor Green
}

# --- 2. Kernel-Power Critical Events (BSOD / Power Loss) ---
Write-Host "`n[2] KERNEL-POWER CRITICAL EVENTS (Event ID 41)" -ForegroundColor Yellow
$kernelPower = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    ProviderName = 'Microsoft-Windows-Kernel-Power'
    Id = 41
    StartTime = $StartDate
} -MaxEvents 10 2>$null

$crashTimes = @()

if ($kernelPower) {
    foreach ($event in $kernelPower) {
        $bugCheckCode = ($event.Properties[0]).Value
        $powerButtonTimestamp = ($event.Properties[1]).Value
        $sleepInProgress = ($event.Properties[2]).Value
        Write-Host "  $($event.TimeCreated)" -ForegroundColor Red
        Write-Host "    BugcheckCode: $bugCheckCode | SleepInProgress: $sleepInProgress" -ForegroundColor Gray

        $crashTimes += $event.TimeCreated

        # Interpret common bugcheck codes
        switch ($bugCheckCode) {
            0 {
                Write-Host "    -> POWER LOSS or HARD SHUTDOWN (no BSOD)" -ForegroundColor Magenta
                Write-Host "       Possible causes: PSU failure, power outage, overheating shutdown," -ForegroundColor DarkGray
                Write-Host "       power button held, or motherboard power delivery issue" -ForegroundColor DarkGray
            }
            159 { Write-Host "    -> HAL_INITIALIZATION_FAILED" -ForegroundColor Magenta }
            209 { Write-Host "    -> VIDEO_DXGKRNL_FATAL_ERROR (Graphics driver)" -ForegroundColor Magenta }
            239 { Write-Host "    -> CRITICAL_PROCESS_DIED" -ForegroundColor Magenta }
            278 { Write-Host "    -> KERNEL_SECURITY_CHECK_FAILURE" -ForegroundColor Magenta }
            default {
                if ($bugCheckCode -ne 0) {
                    Write-Host "    -> Bugcheck code: 0x$($bugCheckCode.ToString('X'))" -ForegroundColor Magenta
                }
            }
        }
    }
} else {
    Write-Host "  No Kernel-Power critical events found." -ForegroundColor Green
}

# --- 3. BugCheck / BSOD Details ---
Write-Host "`n[3] BUGCHECK / BSOD DETAILS (Event ID 1001)" -ForegroundColor Yellow
$bugChecks = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    ProviderName = 'Microsoft-Windows-WER-SystemErrorReporting'
    Id = 1001
    StartTime = $StartDate
} -MaxEvents 10 2>$null

if ($bugChecks) {
    foreach ($event in $bugChecks) {
        Write-Host "  $($event.TimeCreated)" -ForegroundColor Red
        $crashTimes += $event.TimeCreated

        # Extract bugcheck parameters from message
        if ($event.Message -match 'bugcheck.*?(\(0x[0-9A-Fa-f]+.*?\))') {
            Write-Host "    $($Matches[0])" -ForegroundColor Gray
        }
        if ($event.Message -match 'dump.*?saved.*?:(.+?\.dmp)') {
            Write-Host "    Dump file: $($Matches[1].Trim())" -ForegroundColor Gray
        }
    }
} else {
    Write-Host "  No BugCheck events found." -ForegroundColor Green
}

# --- 4. PRE-CRASH EVENT ANALYSIS ---
Write-Host "`n[4] PRE-CRASH EVENT ANALYSIS" -ForegroundColor Yellow

if ($crashTimes.Count -gt 0) {
    Write-Host "  Analyzing $AnalyzePreCrash minutes before each crash..." -ForegroundColor Gray

    $crashTimes = $crashTimes | Sort-Object -Unique | Select-Object -First 5

    foreach ($crashTime in $crashTimes) {
        Write-Host "`n  === Events before crash at $crashTime ===" -ForegroundColor Cyan
        $preStart = $crashTime.AddMinutes(-$AnalyzePreCrash)

        # Thermal/Power throttling events
        $thermalEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            ProviderName = 'Microsoft-Windows-Kernel-Processor-Power'
            StartTime = $preStart
            EndTime = $crashTime
        } -MaxEvents 5 2>$null

        if ($thermalEvents) {
            Write-Host "    [THERMAL/POWER THROTTLING]" -ForegroundColor Red
            foreach ($e in $thermalEvents) {
                Write-Host "      $($e.TimeCreated): $($e.Message.Substring(0, [Math]::Min(100, $e.Message.Length)))..." -ForegroundColor DarkYellow
            }
        }

        # Disk errors before crash
        $diskErrors = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            ProviderName = 'disk'
            StartTime = $preStart
            EndTime = $crashTime
        } -MaxEvents 5 2>$null

        if ($diskErrors) {
            Write-Host "    [DISK ERRORS]" -ForegroundColor Red
            foreach ($e in $diskErrors) {
                Write-Host "      $($e.TimeCreated): $($e.Message -replace '`r`n', ' ')" -ForegroundColor DarkYellow
            }
        }

        # NTFS errors
        $ntfsErrors = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            ProviderName = 'Microsoft-Windows-Ntfs'
            Level = 2,3  # Error, Warning
            StartTime = $preStart
            EndTime = $crashTime
        } -MaxEvents 5 2>$null

        if ($ntfsErrors) {
            Write-Host "    [NTFS ERRORS]" -ForegroundColor Red
            foreach ($e in $ntfsErrors) {
                Write-Host "      $($e.TimeCreated): $($e.Message -replace '`r`n', ' ')" -ForegroundColor DarkYellow
            }
        }

        # Memory/Resource exhaustion
        $resourceEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            Id = 2004  # Resource exhaustion
            StartTime = $preStart
            EndTime = $crashTime
        } -MaxEvents 3 2>$null

        if ($resourceEvents) {
            Write-Host "    [RESOURCE EXHAUSTION]" -ForegroundColor Red
            foreach ($e in $resourceEvents) {
                Write-Host "      $($e.TimeCreated): Low memory/resources detected" -ForegroundColor DarkYellow
            }
        }

        # Application crashes before system crash
        $appCrashes = Get-WinEvent -FilterHashtable @{
            LogName = 'Application'
            ProviderName = 'Application Error'
            StartTime = $preStart
            EndTime = $crashTime
        } -MaxEvents 5 2>$null

        if ($appCrashes) {
            Write-Host "    [APPLICATION CRASHES]" -ForegroundColor Red
            foreach ($e in $appCrashes) {
                if ($e.Message -match 'Faulting application name:\s*(\S+)') {
                    Write-Host "      $($e.TimeCreated): $($Matches[1]) crashed" -ForegroundColor DarkYellow
                }
            }
        }

        # Driver issues
        $driverEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            Level = 2  # Error
            StartTime = $preStart
            EndTime = $crashTime
        } -MaxEvents 20 2>$null | Where-Object { $_.Message -match 'driver|device' } | Select-Object -First 5

        if ($driverEvents) {
            Write-Host "    [DRIVER/DEVICE ERRORS]" -ForegroundColor Red
            foreach ($e in $driverEvents) {
                $shortMsg = ($e.Message -split "`n")[0].Substring(0, [Math]::Min(80, ($e.Message -split "`n")[0].Length))
                Write-Host "      $($e.TimeCreated): $shortMsg" -ForegroundColor DarkYellow
            }
        }

        # Volmgr events (often precede crashes)
        $volmgrEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            ProviderName = 'volmgr'
            StartTime = $preStart
            EndTime = $crashTime
        } -MaxEvents 3 2>$null

        if ($volmgrEvents) {
            Write-Host "    [VOLUME MANAGER]" -ForegroundColor Red
            foreach ($e in $volmgrEvents) {
                Write-Host "      $($e.TimeCreated): Crash dump configuration issue" -ForegroundColor DarkYellow
            }
        }

        # Check for nothing found
        if (-not $thermalEvents -and -not $diskErrors -and -not $ntfsErrors -and
            -not $resourceEvents -and -not $appCrashes -and -not $driverEvents -and -not $volmgrEvents) {
            Write-Host "    No warning events found before this crash" -ForegroundColor Green
            Write-Host "    -> Suggests sudden power loss or hardware failure" -ForegroundColor DarkGray
        }
    }
} else {
    Write-Host "  No crash times to analyze." -ForegroundColor Green
}

# --- 5. POWER CONFIGURATION ANALYSIS ---
Write-Host "`n[5] POWER CONFIGURATION ANALYSIS" -ForegroundColor Yellow

# Check sleep/hibernate issues
$sleepEvents = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    ProviderName = 'Microsoft-Windows-Kernel-Power'
    Id = 42, 107, 506  # Sleep entry, resume from sleep, sleep transition errors
    StartTime = $StartDate
} -MaxEvents 10 2>$null | Where-Object { $_.LevelDisplayName -eq 'Error' -or $_.LevelDisplayName -eq 'Warning' }

if ($sleepEvents) {
    Write-Host "  Sleep/Hibernate issues detected:" -ForegroundColor Red
    foreach ($e in $sleepEvents) {
        Write-Host "    $($e.TimeCreated): $($e.Message.Substring(0, [Math]::Min(80, $e.Message.Length)))" -ForegroundColor DarkYellow
    }
}

# Active power scheme
$powerScheme = powercfg /getactivescheme 2>$null
if ($powerScheme) {
    Write-Host "  Active Power Plan: $($powerScheme -replace 'Power Scheme GUID: [a-f0-9-]+\s*', '')" -ForegroundColor Cyan
}

# Check for wake timers and fast startup
$fastStartup = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name HiberbootEnabled -ErrorAction SilentlyContinue).HiberbootEnabled
Write-Host "  Fast Startup: $(if ($fastStartup -eq 1) { 'Enabled (can cause issues)' } else { 'Disabled' })" -ForegroundColor $(if ($fastStartup -eq 1) { 'Yellow' } else { 'Green' })

# --- 6. HARDWARE HEALTH INDICATORS ---
Write-Host "`n[6] HARDWARE HEALTH INDICATORS" -ForegroundColor Yellow

# WHEA errors (hardware)
$wheaErrors = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    ProviderName = 'Microsoft-Windows-WHEA-Logger'
    StartTime = $StartDate
} -MaxEvents 10 2>$null

if ($wheaErrors) {
    Write-Host "  WHEA Hardware Errors:" -ForegroundColor Red
    foreach ($event in $wheaErrors) {
        $shortMsg = ($event.Message -split "`n")[0]
        Write-Host "    $($event.TimeCreated): $shortMsg" -ForegroundColor DarkYellow

        # Try to identify error type
        if ($event.Message -match 'processor') { Write-Host "      -> CPU issue detected" -ForegroundColor Magenta }
        if ($event.Message -match 'memory|cache') { Write-Host "      -> Memory/Cache issue detected" -ForegroundColor Magenta }
        if ($event.Message -match 'pci') { Write-Host "      -> PCI/Bus issue detected" -ForegroundColor Magenta }
    }
} else {
    Write-Host "  No WHEA hardware errors found." -ForegroundColor Green
}

# SMART disk warnings
$smartEvents = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    Id = 7, 11, 15, 51, 52, 55, 153  # Various disk failure events
    StartTime = $StartDate
} -MaxEvents 10 2>$null

if ($smartEvents) {
    Write-Host "  Disk Health Warnings:" -ForegroundColor Red
    foreach ($e in $smartEvents) {
        Write-Host "    $($e.TimeCreated): Event $($e.Id) - Disk issue" -ForegroundColor DarkYellow
    }
}

# --- 7. Planned Shutdowns / Restarts ---
Write-Host "`n[7] PLANNED SHUTDOWNS/RESTARTS (Event ID 1074)" -ForegroundColor Yellow
$plannedShutdowns = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    Id = 1074
    StartTime = $StartDate
} -MaxEvents 10 2>$null

if ($plannedShutdowns) {
    foreach ($event in $plannedShutdowns) {
        $process = ($event.Properties[0]).Value
        $reason = ($event.Properties[2]).Value
        $user = ($event.Properties[6]).Value
        Write-Host "  $($event.TimeCreated)" -ForegroundColor Cyan
        Write-Host "    Process: $process | User: $user" -ForegroundColor Gray
        Write-Host "    Reason: $reason" -ForegroundColor Gray
    }
} else {
    Write-Host "  No planned shutdown events found." -ForegroundColor Green
}

# --- 8. Windows Update Reboots ---
Write-Host "`n[8] WINDOWS UPDATE INITIATED REBOOTS" -ForegroundColor Yellow
$wuReboots = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    ProviderName = 'Microsoft-Windows-WindowsUpdateClient'
    StartTime = $StartDate
} -MaxEvents 20 2>$null | Where-Object { $_.Message -match 'restart|reboot' }

if ($wuReboots) {
    foreach ($event in $wuReboots) {
        $shortMsg = ($event.Message -split "`n")[0]
        Write-Host "  $($event.TimeCreated): $shortMsg" -ForegroundColor Cyan
    }
} else {
    Write-Host "  No Windows Update reboot events found." -ForegroundColor Green
}

# --- 9. System Boot Times ---
Write-Host "`n[9] RECENT SYSTEM BOOT TIMES" -ForegroundColor Yellow
$bootEvents = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    Id = 6005, 6006, 6009  # Event log started, stopped, boot info
    StartTime = $StartDate
} -MaxEvents 20 2>$null | Sort-Object TimeCreated -Descending

if ($bootEvents) {
    $grouped = $bootEvents | Group-Object { $_.TimeCreated.Date }
    foreach ($group in $grouped | Select-Object -First 7) {
        $boots = ($group.Group | Where-Object { $_.Id -eq 6005 }).Count
        $shutdowns = ($group.Group | Where-Object { $_.Id -eq 6006 }).Count
        Write-Host "  $($group.Name): $boots boot(s), $shutdowns clean shutdown(s)" -ForegroundColor Cyan
    }
}

# --- 10. Current System Uptime ---
Write-Host "`n[10] CURRENT SYSTEM UPTIME" -ForegroundColor Yellow
$os = Get-CimInstance Win32_OperatingSystem
$uptime = (Get-Date) - $os.LastBootUpTime
Write-Host "  Last Boot: $($os.LastBootUpTime)" -ForegroundColor Cyan
Write-Host "  Uptime: $($uptime.Days) days, $($uptime.Hours) hours, $($uptime.Minutes) minutes" -ForegroundColor Cyan

# --- 11. Memory Dump Files ---
Write-Host "`n[11] MEMORY DUMP FILES" -ForegroundColor Yellow
$dumpPath = "$env:SystemRoot\Minidump"
$fullDump = "$env:SystemRoot\MEMORY.DMP"

if (Test-Path $fullDump) {
    $dumpInfo = Get-Item $fullDump
    Write-Host "  Full dump: $fullDump" -ForegroundColor Red
    Write-Host "    Size: $('{0:N2}' -f ($dumpInfo.Length / 1MB)) MB | Date: $($dumpInfo.LastWriteTime)" -ForegroundColor Gray
}

if (Test-Path $dumpPath) {
    $miniDumps = Get-ChildItem $dumpPath -Filter "*.dmp" | Sort-Object LastWriteTime -Descending | Select-Object -First 5
    if ($miniDumps) {
        Write-Host "  Recent minidumps in $dumpPath`:" -ForegroundColor Gray
        foreach ($dump in $miniDumps) {
            Write-Host "    $($dump.Name) - $($dump.LastWriteTime)" -ForegroundColor Red
        }
    }
} else {
    Write-Host "  No minidump folder found." -ForegroundColor Green
}

# --- 12. RELIABILITY HISTORY SUMMARY ---
Write-Host "`n[12] RELIABILITY HISTORY (Application Failures)" -ForegroundColor Yellow
$relEvents = Get-WinEvent -FilterHashtable @{
    LogName = 'Application'
    ProviderName = 'Windows Error Reporting'
    StartTime = $StartDate
} -MaxEvents 20 2>$null

if ($relEvents) {
    $appFailures = @{}
    foreach ($e in $relEvents) {
        if ($e.Message -match 'Fault bucket.*?Application Name:\s*(\S+)') {
            $app = $Matches[1]
            if ($appFailures.ContainsKey($app)) { $appFailures[$app]++ } else { $appFailures[$app] = 1 }
        }
    }
    if ($appFailures.Count -gt 0) {
        Write-Host "  Frequently crashing applications:" -ForegroundColor Red
        $appFailures.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 5 | ForEach-Object {
            Write-Host "    $($_.Key): $($_.Value) crash(es)" -ForegroundColor DarkYellow
        }
    }
} else {
    Write-Host "  No application failures found." -ForegroundColor Green
}

# --- Summary ---
Write-Host "`n===== SUMMARY =====" -ForegroundColor Cyan
$issuesFound = @()

if ($unexpectedShutdowns) { $issuesFound += "Unexpected shutdowns detected (dirty reboots)" }
if ($kernelPower) { $issuesFound += "Kernel-Power critical events (BSOD or power loss)" }
if ($bugChecks) { $issuesFound += "BugCheck/BSOD events recorded" }
if ($wheaErrors) { $issuesFound += "WHEA hardware errors (CPU/RAM/Bus issues)" }
if ($smartEvents) { $issuesFound += "Disk health warnings detected" }
if ($sleepEvents) { $issuesFound += "Sleep/Hibernate transition failures" }

if ($issuesFound.Count -gt 0) {
    Write-Host "Issues found:" -ForegroundColor Red
    foreach ($issue in $issuesFound) {
        Write-Host "  - $issue" -ForegroundColor Red
    }
    Write-Host "`nRecommendations:" -ForegroundColor Yellow
    Write-Host "  1. Analyze dump files with WinDbg or BlueScreenView" -ForegroundColor Gray
    Write-Host "  2. Run memory diagnostics: mdsched.exe" -ForegroundColor Gray
    Write-Host "  3. Check disk health: wmic diskdrive get status" -ForegroundColor Gray
    Write-Host "  4. Check thermals with HWiNFO64 or similar" -ForegroundColor Gray
    Write-Host "  5. Update drivers (especially GPU and chipset)" -ForegroundColor Gray
    Write-Host "  6. Consider running: sfc /scannow and DISM /Online /Cleanup-Image /RestoreHealth" -ForegroundColor Gray
    Write-Host "`n  Run with -InstallMonitoring to set up proactive monitoring tools" -ForegroundColor Cyan
} else {
    Write-Host "No critical reboot issues detected in the last $Days days." -ForegroundColor Green
    if ($plannedShutdowns) {
        Write-Host "Recent reboots appear to be planned (updates, user-initiated, etc.)" -ForegroundColor Cyan
    }
}

# --- MONITORING TOOLS INSTALLATION ---
if ($InstallMonitoring) {
    Write-Host "`n===== MONITORING TOOLS INSTALLATION =====" -ForegroundColor Cyan

    # Check for winget
    $hasWinget = Get-Command winget -ErrorAction SilentlyContinue

    Write-Host "`nAvailable monitoring tools to install:" -ForegroundColor Yellow
    Write-Host "  1. BlueScreenView    - Analyze BSOD minidump files (NirSoft)" -ForegroundColor Gray
    Write-Host "  2. HWiNFO64          - Hardware monitoring (temps, voltages, fan speeds)" -ForegroundColor Gray
    Write-Host "  3. CrystalDiskInfo   - Disk SMART health monitoring" -ForegroundColor Gray
    Write-Host "  4. WhoCrashed        - Automated crash dump analysis" -ForegroundColor Gray
    Write-Host "  5. Reliability Monitor Shortcut" -ForegroundColor Gray
    Write-Host "  6. Create Scheduled Event Log Monitor Task" -ForegroundColor Gray
    Write-Host "  0. Skip / Exit" -ForegroundColor Gray

    $selection = Read-Host "`nEnter numbers to install (comma-separated, e.g., 1,2,3)"
    $choices = $selection -split ',' | ForEach-Object { $_.Trim() }

    foreach ($choice in $choices) {
        switch ($choice) {
            "1" {
                Write-Host "`nInstalling BlueScreenView..." -ForegroundColor Cyan
                if ($hasWinget) {
                    winget install --id NirSoft.BlueScreenView --accept-source-agreements --accept-package-agreements 2>$null
                } else {
                    Write-Host "  Download from: https://www.nirsoft.net/utils/blue_screen_view.html" -ForegroundColor Yellow
                    Start-Process "https://www.nirsoft.net/utils/blue_screen_view.html"
                }
            }
            "2" {
                Write-Host "`nInstalling HWiNFO64..." -ForegroundColor Cyan
                if ($hasWinget) {
                    winget install --id REALiX.HWiNFO --accept-source-agreements --accept-package-agreements 2>$null
                } else {
                    Write-Host "  Download from: https://www.hwinfo.com/download/" -ForegroundColor Yellow
                    Start-Process "https://www.hwinfo.com/download/"
                }
            }
            "3" {
                Write-Host "`nInstalling CrystalDiskInfo..." -ForegroundColor Cyan
                if ($hasWinget) {
                    winget install --id CrystalDewWorld.CrystalDiskInfo --accept-source-agreements --accept-package-agreements 2>$null
                } else {
                    Write-Host "  Download from: https://crystalmark.info/en/software/crystaldiskinfo/" -ForegroundColor Yellow
                    Start-Process "https://crystalmark.info/en/software/crystaldiskinfo/"
                }
            }
            "4" {
                Write-Host "`nWhoCrashed (free version):" -ForegroundColor Cyan
                Write-Host "  Download from: https://www.resplendence.com/whocrashed" -ForegroundColor Yellow
                Start-Process "https://www.resplendence.com/whocrashed"
            }
            "5" {
                Write-Host "`nOpening Reliability Monitor..." -ForegroundColor Cyan
                Start-Process "perfmon" -ArgumentList "/rel"
            }
            "6" {
                Write-Host "`nCreating Scheduled Event Log Monitor..." -ForegroundColor Cyan

                # Create a script to monitor for critical events
                $monitorScript = @'
# Reboot Monitor Script - Logs critical events to a file
$logPath = "$env:USERPROFILE\RebootMonitor.log"
$lastCheck = (Get-Date).AddHours(-1)

$criticalEvents = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    Level = 1,2  # Critical, Error
    StartTime = $lastCheck
} -MaxEvents 50 2>$null | Where-Object {
    $_.ProviderName -match 'Kernel-Power|WHEA|disk|Ntfs|volmgr'
}

if ($criticalEvents) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logPath -Value "`n=== $timestamp ==="
    foreach ($e in $criticalEvents) {
        Add-Content -Path $logPath -Value "$($e.TimeCreated) [$($e.ProviderName)] $($e.Id): $(($e.Message -split "`n")[0])"
    }
}
'@
                $scriptPath = "$env:USERPROFILE\RebootMonitorTask.ps1"
                $monitorScript | Out-File -FilePath $scriptPath -Encoding UTF8

                # Create scheduled task
                $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
                $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1)
                $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

                Register-ScheduledTask -TaskName "RebootMonitor" -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null

                Write-Host "  Created scheduled task 'RebootMonitor'" -ForegroundColor Green
                Write-Host "  Logs will be saved to: $env:USERPROFILE\RebootMonitor.log" -ForegroundColor Gray
                Write-Host "  Task runs hourly to check for critical events" -ForegroundColor Gray
            }
            "0" {
                Write-Host "Skipping tool installation." -ForegroundColor Gray
            }
        }
    }

    Write-Host "`n--- Additional Manual Recommendations ---" -ForegroundColor Yellow
    Write-Host "  - Enable automatic minidump: System Properties > Advanced > Startup and Recovery" -ForegroundColor Gray
    Write-Host "  - Install Windows Debugging Tools (WinDbg) from Windows SDK for advanced analysis" -ForegroundColor Gray
    Write-Host "  - Consider running: verifier /standard /all (Driver Verifier - use with caution)" -ForegroundColor Gray
}

Write-Host ""
