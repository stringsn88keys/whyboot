#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Analyzes patterns across multiple unexpected reboots to identify correlating services.
.DESCRIPTION
    Queries Windows Event Logs across every unexpected reboot in the specified period and
    uses Service Control Manager events (ID 7036) to reconstruct which services were
    running at each crash time. Compares crash sessions against clean boot sessions to
    surface unusual patterns.

    Reports:
      - Individual services that are always or usually running at crash time
      - Services that appear in crashes but rarely/never in clean sessions
      - Services that are always stopped during crashes
      - Service state changes in the N minutes before each crash
      - Service pairs (and combinations) that co-occur across multiple crashes
      - Crash-only pairs: combinations never seen in clean sessions
.PARAMETER Days
    How many days back to search for unexpected reboots. Default: 60.
.PARAMETER PreCrashMinutes
    Minutes before each crash to flag as "pre-crash activity". Default: 30.
.PARAMETER MinCrashesForPattern
    Minimum number of crashes a service must appear in to be reported. Default: 2.
.PARAMETER MaxCleanSessions
    Maximum clean sessions to sample for the baseline comparison. Default: 10.
.PARAMETER NoOllama
    Skip the AI analysis step entirely.
.PARAMETER OllamaModel
    Override the Ollama model from config.json (used as fallback if Claude/Copilot unavailable).
.PARAMETER OllamaUrl
    Override the Ollama API URL from config.json.
.EXAMPLE
    .\Analyze-RebootPatterns.ps1
    .\Analyze-RebootPatterns.ps1 -Days 90 -PreCrashMinutes 60
    .\Analyze-RebootPatterns.ps1 -NoOllama -MinCrashesForPattern 1
#>
param(
    [int]$Days                 = 60,
    [int]$PreCrashMinutes      = 30,
    [int]$MinCrashesForPattern = 2,
    [int]$MaxCleanSessions     = 10,
    [switch]$NoOllama,
    [string]$OllamaModel,
    [string]$OllamaUrl
)

$ErrorActionPreference = "SilentlyContinue"

# ── Load config ────────────────────────────────────────────────────────────────
$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$configPath = Join-Path $scriptDir "config.json"
$config     = $null
if (Test-Path $configPath) {
    try { $config = Get-Content $configPath -Raw | ConvertFrom-Json } catch {}
}
if (-not $OllamaModel) {
    $OllamaModel = if ($config -and $config.OllamaModel) { $config.OllamaModel } else { "qwen3:4b" }
}
if (-not $OllamaUrl) {
    $OllamaUrl = if ($config -and $config.OllamaUrl) { $config.OllamaUrl } else { "http://localhost:11434" }
}

# Load shared AI helper (Claude → Copilot → Ollama priority)
. (Join-Path $scriptDir "AI-Helper.ps1")

$StartDate = (Get-Date).AddDays(-$Days)

Write-Host "`n===== REBOOT PATTERN ANALYZER =====" -ForegroundColor Cyan
Write-Host "Period : $StartDate  →  $(Get-Date)" -ForegroundColor Gray
Write-Host "Looking for unexpected reboots across the last $Days days...`n" -ForegroundColor Gray

# ── Helper: strip Unicode directional marks from event messages ────────────────
function Remove-UnicodeMarks ([string]$s) {
    return $s -replace [char]0x200E,'' -replace [char]0x200F,'' `
               -replace [char]0x202A,'' -replace [char]0x202B,'' `
               -replace [char]0x202C,'' -replace [char]0x202D,'' `
               -replace [char]0x202E,''
}

# ── Helper: extract actual crash time from a 6008 event message ───────────────
# 6008 fires at the next boot; its message contains the prior crash timestamp.
function Get-CrashTimeFrom6008 ([object]$Event) {
    $msg = Remove-UnicodeMarks $Event.Message
    # "The previous system shutdown at 5:52:51 PM on 2/19/2026 was unexpected."
    if ($msg -match 'shutdown at (.+?) on (.+?) was unexpected') {
        $timeStr = $Matches[1].Trim()
        $dateStr = ($Matches[2].Trim()) -replace '\s+', ' '
        try { return [DateTime]::Parse("$dateStr $timeStr") } catch {}
    }
    # Fallback via Properties[0]=time, [1]=date
    if ($Event.Properties.Count -ge 2) {
        try { return [DateTime]::Parse("$($Event.Properties[1].Value) $($Event.Properties[0].Value)") } catch {}
    }
    # Final fallback: estimate 5 minutes before this boot logged the event
    return $Event.TimeCreated.AddMinutes(-5)
}

# ── Helper: parse a 7036 Service Control Manager event ────────────────────────
# Returns @{Name=...; State=...} or $null
function Parse-SCMEvent ([object]$Event) {
    if ($Event.Message -match 'The (.+?) service entered the (running|stopped) state') {
        return @{ Name = $Matches[1].Trim(); State = $Matches[2].ToLower() }
    }
    if ($Event.Properties.Count -ge 2) {
        $state = "$($Event.Properties[1].Value)".ToLower()
        if ($state -match '^(running|stopped)$') {
            return @{ Name = "$($Event.Properties[0].Value)"; State = $state }
        }
    }
    return $null
}

# ── Helper: build service-state snapshot from a slice of 7036 events ──────────
# Returns hashtable: serviceName -> last-seen state ("running"|"stopped")
function Get-ServiceSnapshot ([object[]]$SCMSlice) {
    $map = @{}
    foreach ($evt in $SCMSlice) {
        $p = Parse-SCMEvent $evt
        if ($p) { $map[$p.Name] = $p.State }
    }
    return $map
}

# ══════════════════════════════════════════════════════════════════════════════
# 1. GATHER BOOT/CRASH EVENTS
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "[1] GATHERING BOOT AND CRASH EVENTS" -ForegroundColor Yellow

# Pull slightly further back so we can find boot-starts for early sessions
$extendedStart = $StartDate.AddDays(-7)

$allBootStarts = Get-WinEvent -FilterHashtable @{ LogName='System'; Id=6005; StartTime=$extendedStart } -MaxEvents 500 2>$null | Sort-Object TimeCreated
$allCleanStops = Get-WinEvent -FilterHashtable @{ LogName='System'; Id=6006; StartTime=$extendedStart } -MaxEvents 500 2>$null | Sort-Object TimeCreated
$allDirtyStops = Get-WinEvent -FilterHashtable @{ LogName='System'; Id=6008; StartTime=$StartDate  } -MaxEvents 100 2>$null | Sort-Object TimeCreated
$allKP41       = Get-WinEvent -FilterHashtable @{
    LogName='System'; ProviderName='Microsoft-Windows-Kernel-Power'; Id=41; StartTime=$StartDate
} -MaxEvents 100 2>$null | Sort-Object TimeCreated

Write-Host "  Boot events (6005):           $($allBootStarts.Count)" -ForegroundColor Gray
Write-Host "  Clean shutdowns (6006):       $($allCleanStops.Count)" -ForegroundColor Gray
Write-Host "  Unexpected shutdowns (6008):  $($allDirtyStops.Count)" -ForegroundColor Gray
Write-Host "  Kernel-Power 41:              $($allKP41.Count)" -ForegroundColor Gray

if (-not $allDirtyStops -or $allDirtyStops.Count -eq 0) {
    Write-Host "`n  No unexpected shutdowns found in the last $Days days." -ForegroundColor Green
    Write-Host "  Try a larger -Days value (e.g., -Days 90)." -ForegroundColor Gray
    exit
}

# ══════════════════════════════════════════════════════════════════════════════
# 2. BUILD CRASH SESSION RECORDS
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "`n[2] BUILDING CRASH SESSION RECORDS" -ForegroundColor Yellow

$crashSessions = @()

foreach ($dirty in $allDirtyStops) {
    $bootAfterCrash = $dirty.TimeCreated          # 6008 fires at the next boot
    $crashTime      = Get-CrashTimeFrom6008 $dirty

    # Match KP41 fired at approximately the same boot (within a few minutes)
    $kp41 = $allKP41 | Where-Object {
        $_.TimeCreated -ge $bootAfterCrash.AddMinutes(-3) -and
        $_.TimeCreated -le $bootAfterCrash.AddMinutes(8)
    } | Select-Object -First 1

    $bugCheckCode = if ($kp41) { [int]($kp41.Properties[0].Value) } else { -1 }
    $bugCheckHex  = switch ($bugCheckCode) {
        -1  { "N/A (no KP41 matched)" }
         0  { "0x0 (power loss / hard reset)" }
        default { "0x$($bugCheckCode.ToString('X'))" }
    }

    # Session start = last 6005 BEFORE the crash time
    $priorBoot = $allBootStarts | Where-Object { $_.TimeCreated -lt $crashTime } |
                 Sort-Object TimeCreated -Descending | Select-Object -First 1
    $sessionStart = if ($priorBoot) { $priorBoot.TimeCreated } else { $crashTime.AddHours(-12) }

    $crashSessions += [PSCustomObject]@{
        CrashTime     = $crashTime
        BootTime      = $bootAfterCrash
        SessionStart  = $sessionStart
        SessionLength = $crashTime - $sessionStart
        BugCheckCode  = $bugCheckCode
        BugCheckHex   = $bugCheckHex
        # Populated in Section 4:
        RunningAtCrash  = @()
        StoppedAtCrash  = @()
        PreCrashChanges = @()
    }
}

# Remove duplicates: if the same hour/crash appears twice (multiple 6008s for one crash),
# keep only the earliest record per unique crash-hour.
$crashSessions = $crashSessions | Sort-Object CrashTime |
    Group-Object { $_.CrashTime.ToString("yyyyMMddHHmm").Substring(0,11) } |
    ForEach-Object { $_.Group | Select-Object -First 1 }

$totalCrashes = $crashSessions.Count
Write-Host "  Unique crash sessions: $totalCrashes" -ForegroundColor Red

foreach ($cs in $crashSessions) {
    Write-Host ("  [{0}] Crash {1}  BugCheck: {2}" -f
        ($crashSessions.IndexOf($cs) + 1), $cs.CrashTime, $cs.BugCheckHex) -ForegroundColor Red
    Write-Host ("       Session: {0} → crash  ({1:N1}h uptime)" -f
        $cs.SessionStart, $cs.SessionLength.TotalHours) -ForegroundColor Gray
}

# ══════════════════════════════════════════════════════════════════════════════
# 3. BUILD CLEAN SESSION BASELINE
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "`n[3] BUILDING CLEAN SESSION BASELINE" -ForegroundColor Yellow

$cleanSessions = @()

for ($i = 0; $i -lt $allBootStarts.Count -and $cleanSessions.Count -lt $MaxCleanSessions; $i++) {
    $boot = $allBootStarts[$i]

    # Find the next clean stop after this boot
    $cleanStop = $allCleanStops | Where-Object { $_.TimeCreated -gt $boot.TimeCreated } |
                 Sort-Object TimeCreated | Select-Object -First 1
    if (-not $cleanStop) { continue }

    # Skip if a dirty stop occurred in this window (crashed session, not clean)
    $hasDirty = $allDirtyStops | Where-Object {
        $_.TimeCreated -gt $boot.TimeCreated -and $_.TimeCreated -lt $cleanStop.TimeCreated
    } | Select-Object -First 1
    if ($hasDirty) { continue }

    # Skip if another boot occurred before the clean stop (incomplete session)
    $nextBoot = if ($i + 1 -lt $allBootStarts.Count) { $allBootStarts[$i + 1] } else { $null }
    if ($nextBoot -and $nextBoot.TimeCreated -lt $cleanStop.TimeCreated) { continue }

    $cleanSessions += [PSCustomObject]@{
        SessionStart  = $boot.TimeCreated
        SessionEnd    = $cleanStop.TimeCreated
        SessionLength = $cleanStop.TimeCreated - $boot.TimeCreated
        RunningAtStop = @()
    }
}

$totalClean = $cleanSessions.Count
Write-Host "  Clean sessions found: $totalClean" -ForegroundColor Green
foreach ($cs in $cleanSessions | Select-Object -First 5) {
    Write-Host ("  Clean: {0} → {1}  ({2:N1}h)" -f
        $cs.SessionStart, $cs.SessionEnd, $cs.SessionLength.TotalHours) -ForegroundColor Gray
}
if ($totalClean -gt 5) { Write-Host "  ... and $($totalClean - 5) more" -ForegroundColor DarkGray }

# ══════════════════════════════════════════════════════════════════════════════
# 4. COLLECT SERVICE STATE DATA (Event 7036)
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "`n[4] COLLECTING SERVICE STATE DATA (SCM Event 7036)" -ForegroundColor Yellow
Write-Host "  Loading Service Control Manager events..." -ForegroundColor Gray

# Find the earliest session boundary we need data from
$allSessionStarts = @($crashSessions | ForEach-Object { $_.SessionStart }) +
                    @($cleanSessions | ForEach-Object { $_.SessionStart })
$earliestNeeded   = ($allSessionStarts | Sort-Object | Select-Object -First 1)
if (-not $earliestNeeded) { $earliestNeeded = $StartDate }

$allSCM = Get-WinEvent -FilterHashtable @{
    LogName      = 'System'
    ProviderName = 'Service Control Manager'
    Id           = 7036
    StartTime    = ([DateTime]$earliestNeeded).AddMinutes(-5)
} -MaxEvents 15000 2>$null | Sort-Object TimeCreated

Write-Host "  Loaded $($allSCM.Count) service state-change events." -ForegroundColor Gray

# Populate crash sessions
foreach ($cs in $crashSessions) {
    $sessionSCM = @($allSCM | Where-Object {
        $_.TimeCreated -ge $cs.SessionStart -and $_.TimeCreated -le $cs.CrashTime
    })

    $snap = Get-ServiceSnapshot $sessionSCM
    $cs.RunningAtCrash = @($snap.GetEnumerator() | Where-Object { $_.Value -eq 'running' } |
                           ForEach-Object { $_.Key } | Sort-Object)
    $cs.StoppedAtCrash = @($snap.GetEnumerator() | Where-Object { $_.Value -eq 'stopped' } |
                           ForEach-Object { $_.Key } | Sort-Object)

    # Pre-crash window: service changes in the last PreCrashMinutes before crash
    $preCrashStart = $cs.CrashTime.AddMinutes(-$PreCrashMinutes)
    $changes = @()
    foreach ($evt in ($sessionSCM | Where-Object { $_.TimeCreated -ge $preCrashStart })) {
        $p = Parse-SCMEvent $evt
        if ($p) {
            $changes += [PSCustomObject]@{
                Time        = $evt.TimeCreated
                ServiceName = $p.Name
                State       = $p.State
                MinsBefore  = [Math]::Round(($cs.CrashTime - $evt.TimeCreated).TotalMinutes, 1)
            }
        }
    }
    $cs.PreCrashChanges = $changes

    Write-Host ("  Crash {0}: {1} running, {2} stopped, {3} pre-crash changes" -f
        $cs.CrashTime, $cs.RunningAtCrash.Count, $cs.StoppedAtCrash.Count, $changes.Count) -ForegroundColor Gray
}

# Populate clean sessions
foreach ($cs in $cleanSessions) {
    $sessionSCM = @($allSCM | Where-Object {
        $_.TimeCreated -ge $cs.SessionStart -and $_.TimeCreated -le $cs.SessionEnd
    })
    $snap = Get-ServiceSnapshot $sessionSCM
    $cs.RunningAtStop = @($snap.GetEnumerator() | Where-Object { $_.Value -eq 'running' } |
                          ForEach-Object { $_.Key } | Sort-Object)
}

# ══════════════════════════════════════════════════════════════════════════════
# 5. INDIVIDUAL SERVICE PATTERN ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "`n[5] INDIVIDUAL SERVICE PATTERN ANALYSIS" -ForegroundColor Yellow

# Frequency counts
$crashFreq   = @{}   # serviceName -> count of crash sessions it was running in
$cleanFreq   = @{}   # serviceName -> count of clean sessions it was running in
$stoppedFreq = @{}   # serviceName -> count of crash sessions it was stopped in

foreach ($cs in $crashSessions) {
    foreach ($svc in $cs.RunningAtCrash) {
        $crashFreq[$svc] = ($(if ($crashFreq[$svc] -ne $null) { $crashFreq[$svc] } else { 0 })) + 1
    }
    foreach ($svc in $cs.StoppedAtCrash) {
        $stoppedFreq[$svc] = ($(if ($stoppedFreq[$svc] -ne $null) { $stoppedFreq[$svc] } else { 0 })) + 1
    }
}
foreach ($cs in $cleanSessions) {
    foreach ($svc in $cs.RunningAtStop) {
        $cleanFreq[$svc] = ($(if ($cleanFreq[$svc] -ne $null) { $cleanFreq[$svc] } else { 0 })) + 1
    }
}

# Build categorized lists
$alwaysRunning   = [System.Collections.Generic.List[object]]::new()
$usuallyRunning  = [System.Collections.Generic.List[object]]::new()
$crashSpecific   = [System.Collections.Generic.List[object]]::new()
$alwaysStopped   = [System.Collections.Generic.List[string]]::new()

foreach ($svc in $crashFreq.Keys) {
    $cf      = $crashFreq[$svc]
    $clean   = if ($null -ne $cleanFreq[$svc]) { $cleanFreq[$svc] } else { 0 }
    $pct     = [Math]::Round($cf / $totalCrashes * 100)
    $cleanPct= if ($totalClean -gt 0) { [Math]::Round($clean / $totalClean * 100) } else { -1 }

    $item = [PSCustomObject]@{
        Service   = $svc
        CrashCount= $cf
        CrashPct  = $pct
        CleanCount= $clean
        CleanPct  = $cleanPct
    }

    if ($cf -eq $totalCrashes -and $cf -ge $MinCrashesForPattern) {
        $alwaysRunning.Add($item)
    } elseif ($cf -ge [Math]::Ceiling($totalCrashes * 0.75) -and $cf -ge $MinCrashesForPattern) {
        $usuallyRunning.Add($item)
    }

    # Crash-specific: appears in >=50% of crashes but <=25% of clean sessions
    if ($cf -ge $MinCrashesForPattern -and $pct -ge 50 -and $totalClean -gt 0 -and $cleanPct -le 25) {
        $crashSpecific.Add($item)
    }
}

foreach ($svc in $stoppedFreq.Keys) {
    if ($stoppedFreq[$svc] -eq $totalCrashes -and $totalCrashes -ge $MinCrashesForPattern) {
        $alwaysStopped.Add($svc) | Out-Null
    }
}

# Sort
$alwaysRunning  = @($alwaysRunning  | Sort-Object CleanPct)
$usuallyRunning = @($usuallyRunning | Sort-Object CrashCount -Descending)
$crashSpecific  = @($crashSpecific  | Sort-Object { $_.CrashPct - $_.CleanPct } -Descending)

# Display
Write-Host "`n  ── ALWAYS RUNNING during unexpected reboots ($totalCrashes/$totalCrashes) ──" -ForegroundColor Red
if ($alwaysRunning.Count -gt 0) {
    foreach ($item in $alwaysRunning) {
        $base = if ($totalClean -gt 0) { "  [clean: $($item.CleanCount)/$totalClean ($($item.CleanPct)%)]" } else { "" }
        $flag = if ($totalClean -gt 0 -and $item.CleanPct -le 25) { "  *** UNUSUAL ***" } else { "" }
        $col  = if ($flag) { 'Magenta' } else { 'Red' }
        Write-Host "    [!] $($item.Service)$base$flag" -ForegroundColor $col
    }
} else {
    Write-Host "    (no single service appeared in every crash session)" -ForegroundColor Green
}

Write-Host "`n  ── USUALLY RUNNING (75%+ of crashes, min $MinCrashesForPattern) ──" -ForegroundColor Yellow
if ($usuallyRunning.Count -gt 0) {
    foreach ($item in $usuallyRunning | Select-Object -First 20) {
        $base = if ($totalClean -gt 0) { "  [clean: $($item.CleanCount)/$totalClean ($($item.CleanPct)%)]" } else { "" }
        Write-Host "    $($item.Service): $($item.CrashCount)/$totalCrashes ($($item.CrashPct)%)$base" -ForegroundColor Yellow
    }
    if ($usuallyRunning.Count -gt 20) {
        Write-Host "    ... and $($usuallyRunning.Count - 20) more (see report file)" -ForegroundColor DarkGray
    }
} else {
    Write-Host "    (none)" -ForegroundColor Gray
}

if ($totalClean -gt 0 -and $crashSpecific.Count -gt 0) {
    Write-Host "`n  ── CRASH-SPECIFIC (high crash rate, low clean-session rate) ──" -ForegroundColor Magenta
    foreach ($item in $crashSpecific | Select-Object -First 15) {
        Write-Host ("    [!!] {0}: crashes={1}%  clean={2}%" -f
            $item.Service, $item.CrashPct, $item.CleanPct) -ForegroundColor Magenta
    }
}

Write-Host "`n  ── ALWAYS STOPPED during unexpected reboots ──" -ForegroundColor Cyan
if ($alwaysStopped.Count -gt 0) {
    foreach ($svc in $alwaysStopped | Sort-Object) {
        Write-Host "    [ ] $svc" -ForegroundColor Cyan
    }
} else {
    Write-Host "    (none)" -ForegroundColor Gray
}

# ══════════════════════════════════════════════════════════════════════════════
# 6. PRE-CRASH SERVICE ACTIVITY PATTERNS
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "`n[6] PRE-CRASH SERVICE ACTIVITY (last $PreCrashMinutes min before each crash)" -ForegroundColor Yellow

# Count unique service:state transitions that appear across multiple crashes
$preCrashCounts = @{}
foreach ($cs in $crashSessions) {
    $seen = @{}
    foreach ($chg in $cs.PreCrashChanges) {
        $key = "$($chg.ServiceName):$($chg.State)"
        if (-not $seen[$key]) {
            $seen[$key]          = $true
            $preCrashCounts[$key] = (if ($null -ne $preCrashCounts[$key]) { $preCrashCounts[$key] } else { 0 }) + 1
        }
    }
}

$repeatedPreCrash = @($preCrashCounts.GetEnumerator() |
    Where-Object { $_.Value -ge $MinCrashesForPattern } |
    Sort-Object Value -Descending)

if ($repeatedPreCrash.Count -gt 0) {
    Write-Host ("  Service state changes seen in {0}+ crashes within {1}min of crash:" -f
        $MinCrashesForPattern, $PreCrashMinutes) -ForegroundColor Red
    foreach ($item in $repeatedPreCrash) {
        $parts = $item.Key -split ':'
        $col   = if ($parts[1] -eq 'running') { 'Yellow' } else { 'Cyan' }
        Write-Host ("    [{0}] {1}: {2}/{3} crashes" -f
            $parts[1].ToUpper(), $parts[0], $item.Value, $totalCrashes) -ForegroundColor $col
    }
} else {
    Write-Host "  No service activity repeated in $MinCrashesForPattern+ crashes before the event." -ForegroundColor Green
}

# ══════════════════════════════════════════════════════════════════════════════
# 7. SERVICE COMBINATION / CO-OCCURRENCE ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "`n[7] SERVICE COMBINATION ANALYSIS" -ForegroundColor Yellow

# Candidate services: appeared as "running" in at least MinCrashesForPattern crashes
$candidates = @($crashFreq.GetEnumerator() |
    Where-Object { $_.Value -ge $MinCrashesForPattern } |
    ForEach-Object { $_.Key })

$pairCrashCount = @{}
$pairCleanCount = @{}

if ($totalCrashes -ge 2) {
    Write-Host ("  Evaluating pairs among {0} candidate services..." -f $candidates.Count) -ForegroundColor Gray

    foreach ($cs in $crashSessions) {
        $active = @($cs.RunningAtCrash | Where-Object { $_ -in $candidates })
        for ($i = 0; $i -lt $active.Count; $i++) {
            for ($j = $i + 1; $j -lt $active.Count; $j++) {
                $k = "$($active[$i])|$($active[$j])"
                $pairCrashCount[$k] = (if ($null -ne $pairCrashCount[$k]) { $pairCrashCount[$k] } else { 0 }) + 1
            }
        }
    }

    if ($totalClean -gt 0) {
        foreach ($cs in $cleanSessions) {
            $active = @($cs.RunningAtStop | Where-Object { $_ -in $candidates })
            for ($i = 0; $i -lt $active.Count; $i++) {
                for ($j = $i + 1; $j -lt $active.Count; $j++) {
                    $k = "$($active[$i])|$($active[$j])"
                    $pairCleanCount[$k] = (if ($null -ne $pairCleanCount[$k]) { $pairCleanCount[$k] } else { 0 }) + 1
                }
            }
        }
    }

    # Categorise
    $alwaysPairs    = [System.Collections.Generic.List[object]]::new()
    $usuallyPairs   = [System.Collections.Generic.List[object]]::new()
    $crashOnlyPairs = [System.Collections.Generic.List[object]]::new()

    foreach ($pair in $pairCrashCount.GetEnumerator()) {
        $cc = $pair.Value
        if ($cc -lt $MinCrashesForPattern) { continue }
        $cl     = if ($null -ne $pairCleanCount[$pair.Key]) { $pairCleanCount[$pair.Key] } else { 0 }
        $parts  = $pair.Key -split '\|'
        $obj    = [PSCustomObject]@{
            Service1   = $parts[0]
            Service2   = $parts[1]
            CrashCount = $cc
            CleanCount = $cl
        }
        if ($cc -eq $totalCrashes)                                   { $alwaysPairs.Add($obj)    | Out-Null }
        elseif ($cc -ge $MinCrashesForPattern)                       { $usuallyPairs.Add($obj)   | Out-Null }
        if ($cc -ge $MinCrashesForPattern -and $cl -eq 0 -and $totalClean -gt 0) {
            $crashOnlyPairs.Add($obj) | Out-Null
        }
    }

    $alwaysPairs    = @($alwaysPairs    | Sort-Object CrashCount -Descending)
    $usuallyPairs   = @($usuallyPairs   | Sort-Object CrashCount -Descending)
    $crashOnlyPairs = @($crashOnlyPairs | Sort-Object CrashCount -Descending)

    Write-Host "`n  ── PAIRS always co-running ($totalCrashes/$totalCrashes crashes) ──" -ForegroundColor Red
    if ($alwaysPairs.Count -gt 0) {
        foreach ($p in $alwaysPairs | Select-Object -First 20) {
            $ci = if ($totalClean -gt 0) { "  [clean: $($p.CleanCount)/$totalClean]" } else { "" }
            Write-Host "    [!] $($p.Service1)  +  $($p.Service2)$ci" -ForegroundColor Red
        }
        if ($alwaysPairs.Count -gt 20) {
            Write-Host "    ... and $($alwaysPairs.Count - 20) more pairs (see report)" -ForegroundColor DarkGray
        }
    } else {
        Write-Host "    (none)" -ForegroundColor Gray
    }

    if ($usuallyPairs.Count -gt 0) {
        Write-Host "`n  ── PAIRS in $MinCrashesForPattern+ crashes (not all) ──" -ForegroundColor Yellow
        foreach ($p in $usuallyPairs | Select-Object -First 15) {
            $ci = if ($totalClean -gt 0) { "  [clean: $($p.CleanCount)/$totalClean]" } else { "" }
            Write-Host "    $($p.Service1)  +  $($p.Service2): $($p.CrashCount)/$totalCrashes$ci" -ForegroundColor Yellow
        }
    }

    if ($totalClean -gt 0 -and $crashOnlyPairs.Count -gt 0) {
        Write-Host "`n  ── CRASH-ONLY PAIRS (never seen in $totalClean clean sessions) ──" -ForegroundColor Magenta
        foreach ($p in $crashOnlyPairs | Select-Object -First 15) {
            Write-Host ("    [!!] {0}  +  {1}:  {2}/{3} crashes, 0/{4} clean" -f
                $p.Service1, $p.Service2, $p.CrashCount, $totalCrashes, $totalClean) -ForegroundColor Magenta
        }
    }
} else {
    Write-Host "  Need at least 2 crash sessions for combination analysis ($totalCrashes found)." -ForegroundColor Gray
    $alwaysPairs = @(); $usuallyPairs = @(); $crashOnlyPairs = @()
}

# ══════════════════════════════════════════════════════════════════════════════
# 8. CRASH-BY-CRASH DETAIL
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "`n[8] CRASH-BY-CRASH DETAIL" -ForegroundColor Yellow

for ($i = 0; $i -lt $crashSessions.Count; $i++) {
    $cs = $crashSessions[$i]
    Write-Host ("`n  [{0}] Crash at {1}   BugCheck: {2}" -f ($i + 1), $cs.CrashTime, $cs.BugCheckHex) -ForegroundColor Red
    Write-Host ("       Session: {0} → crash  ({1:N1}h uptime)" -f
        $cs.SessionStart, $cs.SessionLength.TotalHours) -ForegroundColor Gray
    Write-Host "       Services tracked running at crash: $($cs.RunningAtCrash.Count)" -ForegroundColor Gray

    if ($cs.PreCrashChanges.Count -gt 0) {
        Write-Host "       Service state changes in last $PreCrashMinutes min:" -ForegroundColor DarkYellow
        foreach ($chg in $cs.PreCrashChanges | Sort-Object MinsBefore -Descending) {
            Write-Host ("         {0,5:N1}m before: [{1}] {2}" -f
                $chg.MinsBefore, $chg.State.ToUpper(), $chg.ServiceName) -ForegroundColor DarkYellow
        }
    } else {
        Write-Host "       No service changes in last $PreCrashMinutes min before crash" -ForegroundColor DarkGray
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 9. WRITE REPORT FILES
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "`n[9] GENERATING REPORT FILES" -ForegroundColor Yellow

$ts         = Get-Date -Format "yyyyMMdd-HHmmss"
$reportTxt  = Join-Path $scriptDir "reboot-patterns-$ts.txt"
$reportMd   = Join-Path $scriptDir "reboot-patterns-$ts.md"
$sep        = "=" * 80

# ── TEXT REPORT ────────────────────────────────────────────────────────────────
$t = [System.Text.StringBuilder]::new()
$nl = "`r`n"

$t.Append($sep + $nl) | Out-Null
$t.Append("  REBOOT PATTERN ANALYSIS$nl") | Out-Null
$t.Append("  Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')$nl") | Out-Null
$t.Append("  Period    : Last $Days days ($StartDate → $(Get-Date -Format 'yyyy-MM-dd'))$nl") | Out-Null
$t.Append($sep + $nl + $nl) | Out-Null

$t.Append("OVERVIEW$nl") | Out-Null
$t.Append("  Unexpected reboots  : $totalCrashes$nl") | Out-Null
$t.Append("  Clean sessions      : $totalClean (baseline)$nl") | Out-Null
$t.Append("  Pre-crash window    : $PreCrashMinutes minutes$nl") | Out-Null
$t.Append("  Min crashes/pattern : $MinCrashesForPattern$nl$nl") | Out-Null

$t.Append("UNEXPECTED REBOOTS$nl") | Out-Null
for ($i = 0; $i -lt $crashSessions.Count; $i++) {
    $cs = $crashSessions[$i]
    $t.Append("  [$($i+1)] $($cs.CrashTime)  BugCheck: $($cs.BugCheckHex)$nl") | Out-Null
    $t.Append("       Session: $($cs.SessionStart) → crash  ($([Math]::Round($cs.SessionLength.TotalHours,1))h uptime)$nl") | Out-Null
    $t.Append("       Running services tracked: $($cs.RunningAtCrash.Count)$nl") | Out-Null
}
$t.Append($nl) | Out-Null

# Individual patterns
$t.Append("INDIVIDUAL SERVICE PATTERNS$nl$nl") | Out-Null

$t.Append("  ALWAYS RUNNING during crashes ($totalCrashes/$totalCrashes):$nl") | Out-Null
if ($alwaysRunning.Count -gt 0) {
    foreach ($item in $alwaysRunning) {
        $base = if ($totalClean -gt 0) { "  [clean: $($item.CleanCount)/$totalClean ($($item.CleanPct)%)]" } else { "" }
        $flag = if ($totalClean -gt 0 -and $item.CleanPct -le 25) { "  *** UNUSUAL ***" } else { "" }
        $t.Append("    [!] $($item.Service)$base$flag$nl") | Out-Null
    }
} else {
    $t.Append("    (none)$nl") | Out-Null
}
$t.Append($nl) | Out-Null

$t.Append("  USUALLY RUNNING (75%+ of crashes, min $MinCrashesForPattern):$nl") | Out-Null
if ($usuallyRunning.Count -gt 0) {
    foreach ($item in $usuallyRunning) {
        $base = if ($totalClean -gt 0) { "  [clean: $($item.CleanCount)/$totalClean ($($item.CleanPct)%)]" } else { "" }
        $t.Append("    $($item.Service): $($item.CrashCount)/$totalCrashes ($($item.CrashPct)%)$base$nl") | Out-Null
    }
} else {
    $t.Append("    (none)$nl") | Out-Null
}
$t.Append($nl) | Out-Null

if ($totalClean -gt 0 -and $crashSpecific.Count -gt 0) {
    $t.Append("  CRASH-SPECIFIC SERVICES (high crash rate, low clean rate):$nl") | Out-Null
    foreach ($item in $crashSpecific | Select-Object -First 20) {
        $t.Append("    [!!] $($item.Service):  crash=$($item.CrashPct)%  clean=$($item.CleanPct)%$nl") | Out-Null
    }
    $t.Append($nl) | Out-Null
}

$t.Append("  ALWAYS STOPPED during crashes:$nl") | Out-Null
if ($alwaysStopped.Count -gt 0) {
    foreach ($svc in $alwaysStopped | Sort-Object) {
        $t.Append("    [ ] $svc$nl") | Out-Null
    }
} else {
    $t.Append("    (none)$nl") | Out-Null
}
$t.Append($nl) | Out-Null

# Pre-crash activity
$t.Append("PRE-CRASH SERVICE ACTIVITY (within $PreCrashMinutes min of crash)$nl") | Out-Null
if ($repeatedPreCrash.Count -gt 0) {
    foreach ($item in $repeatedPreCrash) {
        $parts = $item.Key -split ':'
        $t.Append("  [$($parts[1].ToUpper())] $($parts[0]):  $($item.Value)/$totalCrashes crashes$nl") | Out-Null
    }
} else {
    $t.Append("  (no repeated pre-crash service activity)$nl") | Out-Null
}
$t.Append($nl) | Out-Null

# Combination patterns
$t.Append("SERVICE COMBINATION PATTERNS$nl$nl") | Out-Null
if ($totalCrashes -ge 2) {
    $t.Append("  ALWAYS CO-RUNNING ($totalCrashes/$totalCrashes crashes):$nl") | Out-Null
    if ($alwaysPairs.Count -gt 0) {
        foreach ($p in $alwaysPairs | Select-Object -First 30) {
            $ci = if ($totalClean -gt 0) { "  [clean: $($p.CleanCount)/$totalClean]" } else { "" }
            $t.Append("    [!] $($p.Service1)  +  $($p.Service2)$ci$nl") | Out-Null
        }
    } else {
        $t.Append("    (none)$nl") | Out-Null
    }
    $t.Append($nl) | Out-Null

    $t.Append("  CO-RUNNING in $MinCrashesForPattern+ crashes (not all):$nl") | Out-Null
    if ($usuallyPairs.Count -gt 0) {
        foreach ($p in $usuallyPairs | Select-Object -First 30) {
            $ci = if ($totalClean -gt 0) { "  [clean: $($p.CleanCount)/$totalClean]" } else { "" }
            $t.Append("    $($p.Service1)  +  $($p.Service2):  $($p.CrashCount)/$totalCrashes$ci$nl") | Out-Null
        }
    } else {
        $t.Append("    (none)$nl") | Out-Null
    }
    $t.Append($nl) | Out-Null

    if ($totalClean -gt 0 -and $crashOnlyPairs.Count -gt 0) {
        $t.Append("  CRASH-ONLY PAIRS (never seen in $totalClean clean sessions):$nl") | Out-Null
        foreach ($p in $crashOnlyPairs | Select-Object -First 20) {
            $t.Append("    [!!] $($p.Service1)  +  $($p.Service2):  $($p.CrashCount)/$totalCrashes crashes, 0/$totalClean clean$nl") | Out-Null
        }
        $t.Append($nl) | Out-Null
    }
}

# Per-crash detail
$t.Append("CRASH SESSION DETAIL$nl") | Out-Null
for ($i = 0; $i -lt $crashSessions.Count; $i++) {
    $cs = $crashSessions[$i]
    $t.Append("$nl  Crash #$($i+1)$nl") | Out-Null
    $t.Append("    Time        : $($cs.CrashTime)$nl") | Out-Null
    $t.Append("    Boot after  : $($cs.BootTime)$nl") | Out-Null
    $t.Append("    BugCheck    : $($cs.BugCheckHex)$nl") | Out-Null
    $t.Append("    Session     : $($cs.SessionStart) → crash ($([Math]::Round($cs.SessionLength.TotalHours,1))h)$nl") | Out-Null
    $t.Append("    Running services ($($cs.RunningAtCrash.Count)):$nl") | Out-Null
    foreach ($svc in $cs.RunningAtCrash) { $t.Append("      + $svc$nl") | Out-Null }
    if ($cs.StoppedAtCrash.Count -gt 0) {
        $t.Append("    Stopped services ($($cs.StoppedAtCrash.Count)):$nl") | Out-Null
        foreach ($svc in $cs.StoppedAtCrash) { $t.Append("      - $svc$nl") | Out-Null }
    }
    if ($cs.PreCrashChanges.Count -gt 0) {
        $t.Append("    Service changes in last $PreCrashMinutes min:$nl") | Out-Null
        foreach ($chg in $cs.PreCrashChanges | Sort-Object MinsBefore -Descending) {
            $t.Append("      $($chg.MinsBefore)m before: [$($chg.State.ToUpper())] $($chg.ServiceName)$nl") | Out-Null
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 10. AI PATTERN ANALYSIS (Claude → Copilot → Ollama)
# ══════════════════════════════════════════════════════════════════════════════
$ollamaAnalysis = ""
$aiProvider     = "none"
if (-not $NoOllama) {
    Write-Host "`n[10] QUERYING AI FOR PATTERN ANALYSIS (Claude → Copilot → Ollama)" -ForegroundColor Yellow

    $alwaysStr = if ($alwaysRunning.Count -gt 0) {
        ($alwaysRunning | ForEach-Object {
            $b = if ($totalClean -gt 0) { " [clean: $($_.CleanCount)/$totalClean ($($_.CleanPct)%)]" } else { "" }
            "  $($_.Service)$b"
        }) -join "`n"
    } else { "  (none)" }

    $crashSpecStr = if ($crashSpecific.Count -gt 0) {
        ($crashSpecific | Select-Object -First 12 | ForEach-Object {
            "  $($_.Service) (crash=$($_.CrashPct)%, clean=$($_.CleanPct)%)"
        }) -join "`n"
    } else { "  (none)" }

    $preCrashStr = if ($repeatedPreCrash.Count -gt 0) {
        ($repeatedPreCrash | ForEach-Object {
            $parts = $_.Key -split ':'
            "  [$($parts[1].ToUpper())] $($parts[0]): $($_.Value)/$totalCrashes crashes"
        }) -join "`n"
    } else { "  (none)" }

    $crashOnlyStr = if ($crashOnlyPairs.Count -gt 0) {
        ($crashOnlyPairs | Select-Object -First 10 | ForEach-Object {
            "  $($_.Service1) + $($_.Service2) ($($_.CrashCount)/$totalCrashes crashes)"
        }) -join "`n"
    } else { "  (none)" }

    $crashDetailStr = ($crashSessions | ForEach-Object {
        "  - $($_.CrashTime)  BugCheck=$($_.BugCheckHex)  Uptime=$([Math]::Round($_.SessionLength.TotalHours,1))h  Services=$($_.RunningAtCrash.Count)"
    }) -join "`n"

    $prompt = @"
You are a Windows system stability expert. Below is pattern data extracted from $totalCrashes unexpected system reboots over the last $Days days, compared against $totalClean clean shutdown sessions as a baseline.

Data source: Service Control Manager Event 7036 (service state transitions) was used to reconstruct which services were running at crash time within each boot session.

Note: Only services that had at least one state-change event logged during a session appear in the data. Services with no 7036 events (e.g., core kernel components) are not included.

═══ SERVICES ALWAYS RUNNING DURING CRASHES ($totalCrashes/$totalCrashes):
$alwaysStr

═══ CRASH-SPECIFIC SERVICES (high crash rate, low clean-session rate):
$crashSpecStr

═══ SERVICE CHANGES IN LAST $PreCrashMinutes MINUTES BEFORE CRASHES:
$preCrashStr

═══ SERVICE PAIRS THAT ONLY APPEAR DURING CRASHES (0/$totalClean clean sessions):
$crashOnlyStr

═══ CRASH TIMELINE:
$crashDetailStr

Please provide a structured analysis:

1. MOST SUSPICIOUS PATTERNS: Which services or combinations stand out as most likely to be contributing to instability? Explain why based on the data (e.g., high crash rate + low clean rate, pre-crash timing, combination uniqueness).

2. ALWAYS-RUNNING SERVICES TO INVESTIGATE: Among the "always running" services, which warrant further investigation, especially if they have a low clean-session percentage? Which are likely benign Windows services?

3. PRE-CRASH SIGNALS: What do the service changes just before crashes suggest about the crash trigger or sequence of events?

4. ACTIONABLE RECOMMENDATIONS: Specific steps to investigate or resolve the issue - e.g., check/disable/update a specific service, look for service interactions, examine service dependencies.

5. CONFIDENCE & CAVEATS: Given the sample size ($totalCrashes crashes, $totalClean clean sessions), how confident are you in these patterns? What additional data would help confirm or refute the patterns?

Be specific and prioritise the most actionable findings.
"@

    $aiResult       = Invoke-AIAnalysis -Prompt $prompt -Config $config -OllamaModel $OllamaModel -OllamaUrl $OllamaUrl
    $ollamaAnalysis = $aiResult.Response
    $aiProvider     = $aiResult.Provider
    if ($aiResult.Success) {
        Write-Host "  Analysis complete via $aiProvider." -ForegroundColor Green
    } else {
        Write-Host "  AI analysis unavailable." -ForegroundColor Red
    }

    $t.Append("$nl$sep$nl  AI PATTERN ANALYSIS (provider: $aiProvider)$nl$sep$nl$nl") | Out-Null
    $t.Append($ollamaAnalysis + $nl) | Out-Null
}

# Write text report
$t.ToString() | Out-File -FilePath $reportTxt -Encoding UTF8

# ── MARKDOWN REPORT ────────────────────────────────────────────────────────────
$md = [System.Text.StringBuilder]::new()
$md.Append("# Reboot Pattern Analysis$nl") | Out-Null
$md.Append("**Generated:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $nl") | Out-Null
$md.Append("**Period:** Last $Days days  $nl") | Out-Null
$md.Append("**Crashes analyzed:** $totalCrashes | **Clean sessions (baseline):** $totalClean$nl$nl") | Out-Null

$md.Append("## Unexpected Reboots$nl$nl") | Out-Null
for ($i = 0; $i -lt $crashSessions.Count; $i++) {
    $cs = $crashSessions[$i]
    $md.Append("### Crash #$($i+1) - $($cs.CrashTime)$nl") | Out-Null
    $md.Append("| Field | Value |$nl|-------|-------|$nl") | Out-Null
    $md.Append("| BugCheck | ``$($cs.BugCheckHex)`` |$nl") | Out-Null
    $md.Append("| Session start | $($cs.SessionStart) |$nl") | Out-Null
    $md.Append("| Uptime before crash | $([Math]::Round($cs.SessionLength.TotalHours,1))h |$nl") | Out-Null
    $md.Append("| Services running at crash | $($cs.RunningAtCrash.Count) |$nl") | Out-Null
    if ($cs.PreCrashChanges.Count -gt 0) {
        $md.Append("$nl**Pre-crash service changes (last $PreCrashMinutes min):**$nl$nl") | Out-Null
        foreach ($chg in $cs.PreCrashChanges | Sort-Object MinsBefore -Descending) {
            $md.Append("- $($chg.MinsBefore)m before: ``[$($chg.State.ToUpper())]`` $($chg.ServiceName)$nl") | Out-Null
        }
    }
    $md.Append($nl) | Out-Null
}

$md.Append("## Individual Service Patterns$nl$nl") | Out-Null

$md.Append("### Always Running During Crashes ($totalCrashes/$totalCrashes)$nl$nl") | Out-Null
if ($alwaysRunning.Count -gt 0) {
    $cleanHdr = if ($totalClean -gt 0) { " Clean Count | Clean % | Unusual? |" } else { "" }
    $cleanDiv = if ($totalClean -gt 0) { " ----------- | ------- | -------- |" } else { "" }
    $md.Append("| Service |$cleanHdr$nl| ------- |$cleanDiv$nl") | Out-Null
    foreach ($item in $alwaysRunning) {
        $unusual = if ($totalClean -gt 0 -and $item.CleanPct -le 25) { "**YES**" } else { "No" }
        $cleanCols = if ($totalClean -gt 0) { " $($item.CleanCount)/$totalClean | $($item.CleanPct)% | $unusual |" } else { "" }
        $md.Append("| $($item.Service) |$cleanCols$nl") | Out-Null
    }
} else {
    $md.Append("*(no service appeared in every crash session)*$nl") | Out-Null
}
$md.Append($nl) | Out-Null

$md.Append("### Usually Running (75%+ of crashes)$nl$nl") | Out-Null
if ($usuallyRunning.Count -gt 0) {
    $cleanHdr = if ($totalClean -gt 0) { " Clean Count | Clean % |" } else { "" }
    $cleanDiv = if ($totalClean -gt 0) { " ----------- | ------- |" } else { "" }
    $md.Append("| Service | Crash Count | Crash % |$cleanHdr$nl| ------- | ----------- | ------- |$cleanDiv$nl") | Out-Null
    foreach ($item in $usuallyRunning) {
        $cleanCols = if ($totalClean -gt 0) { " $($item.CleanCount)/$totalClean | $($item.CleanPct)% |" } else { "" }
        $md.Append("| $($item.Service) | $($item.CrashCount)/$totalCrashes | $($item.CrashPct)% |$cleanCols$nl") | Out-Null
    }
} else {
    $md.Append("*(none)*$nl") | Out-Null
}
$md.Append($nl) | Out-Null

if ($totalClean -gt 0 -and $crashSpecific.Count -gt 0) {
    $md.Append("### Crash-Specific Services$nl$nl") | Out-Null
    $md.Append("| Service | Crash Rate | Clean Rate |$nl| ------- | ---------- | ---------- |$nl") | Out-Null
    foreach ($item in $crashSpecific | Select-Object -First 20) {
        $md.Append("| $($item.Service) | $($item.CrashPct)% | $($item.CleanPct)% |$nl") | Out-Null
    }
    $md.Append($nl) | Out-Null
}

$md.Append("### Always Stopped During Crashes$nl$nl") | Out-Null
if ($alwaysStopped.Count -gt 0) {
    foreach ($svc in $alwaysStopped | Sort-Object) { $md.Append("- $svc$nl") | Out-Null }
} else {
    $md.Append("*(none)*$nl") | Out-Null
}
$md.Append($nl) | Out-Null

$md.Append("## Pre-Crash Service Activity$nl$nl") | Out-Null
$md.Append("*Service state changes within $PreCrashMinutes minutes before each crash, seen in $MinCrashesForPattern+ crashes.*$nl$nl") | Out-Null
if ($repeatedPreCrash.Count -gt 0) {
    $md.Append("| State | Service | Crash Count |$nl| ----- | ------- | ----------- |$nl") | Out-Null
    foreach ($item in $repeatedPreCrash) {
        $parts = $item.Key -split ':'
        $md.Append("| $($parts[1].ToUpper()) | $($parts[0]) | $($item.Value)/$totalCrashes |$nl") | Out-Null
    }
} else {
    $md.Append("*(no repeated pre-crash service activity)*$nl") | Out-Null
}
$md.Append($nl) | Out-Null

$md.Append("## Service Combination Patterns$nl$nl") | Out-Null
if ($totalCrashes -ge 2) {
    $md.Append("### Always Co-Running ($totalCrashes/$totalCrashes crashes)$nl$nl") | Out-Null
    if ($alwaysPairs.Count -gt 0) {
        $cleanHdr = if ($totalClean -gt 0) { " Clean Count |" } else { "" }
        $cleanDiv = if ($totalClean -gt 0) { " ----------- |" } else { "" }
        $md.Append("| Service 1 | Service 2 |$cleanHdr$nl| --------- | --------- |$cleanDiv$nl") | Out-Null
        foreach ($p in $alwaysPairs | Select-Object -First 30) {
            $ci = if ($totalClean -gt 0) { " $($p.CleanCount)/$totalClean |" } else { "" }
            $md.Append("| $($p.Service1) | $($p.Service2) |$ci$nl") | Out-Null
        }
    } else {
        $md.Append("*(none)*$nl") | Out-Null
    }
    $md.Append($nl) | Out-Null

    if ($totalClean -gt 0 -and $crashOnlyPairs.Count -gt 0) {
        $md.Append("### Crash-Only Pairs (never in $totalClean clean sessions)$nl$nl") | Out-Null
        $md.Append("| Service 1 | Service 2 | Crash Count |$nl| --------- | --------- | ----------- |$nl") | Out-Null
        foreach ($p in $crashOnlyPairs | Select-Object -First 20) {
            $md.Append("| $($p.Service1) | $($p.Service2) | $($p.CrashCount)/$totalCrashes |$nl") | Out-Null
        }
        $md.Append($nl) | Out-Null
    }

    if ($usuallyPairs.Count -gt 0) {
        $md.Append("### Co-Running in $MinCrashesForPattern+ Crashes (not all)$nl$nl") | Out-Null
        $cleanHdr = if ($totalClean -gt 0) { " Clean Count |" } else { "" }
        $cleanDiv = if ($totalClean -gt 0) { " ----------- |" } else { "" }
        $md.Append("| Service 1 | Service 2 | Crash Count |$cleanHdr$nl| --------- | --------- | ----------- |$cleanDiv$nl") | Out-Null
        foreach ($p in $usuallyPairs | Select-Object -First 30) {
            $ci = if ($totalClean -gt 0) { " $($p.CleanCount)/$totalClean |" } else { "" }
            $md.Append("| $($p.Service1) | $($p.Service2) | $($p.CrashCount)/$totalCrashes |$ci$nl") | Out-Null
        }
        $md.Append($nl) | Out-Null
    }
}

if ($ollamaAnalysis -and $aiProvider -ne 'none') {
    $md.Append("## AI Pattern Analysis$nl") | Out-Null
    $md.Append("*Provider: $aiProvider*$nl$nl") | Out-Null
    $md.Append($ollamaAnalysis + $nl) | Out-Null
}

$md.ToString() | Out-File -FilePath $reportMd -Encoding UTF8

Write-Host "  Text report : $reportTxt" -ForegroundColor Green
Write-Host "  MD report   : $reportMd" -ForegroundColor Green

# ══════════════════════════════════════════════════════════════════════════════
# SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "`n===== PATTERN ANALYSIS COMPLETE =====" -ForegroundColor Cyan
Write-Host ("  Crashes analyzed      : {0}" -f $totalCrashes) -ForegroundColor Gray
Write-Host ("  Clean sessions        : {0}" -f $totalClean) -ForegroundColor Gray
Write-Host ("  Always-running svcs   : {0}" -f $alwaysRunning.Count) -ForegroundColor $(if ($alwaysRunning.Count   -gt 0) { 'Red'     } else { 'Green' })
Write-Host ("  Crash-specific svcs   : {0}" -f $crashSpecific.Count)  -ForegroundColor $(if ($crashSpecific.Count   -gt 0) { 'Magenta' } else { 'Green' })
Write-Host ("  Pre-crash patterns    : {0}" -f $repeatedPreCrash.Count) -ForegroundColor $(if ($repeatedPreCrash.Count -gt 0) { 'Yellow'  } else { 'Green' })
Write-Host ("  Always co-run pairs   : {0}" -f $alwaysPairs.Count)    -ForegroundColor $(if ($alwaysPairs.Count     -gt 0) { 'Red'     } else { 'Green' })
Write-Host ("  Crash-only pairs      : {0}" -f $crashOnlyPairs.Count)  -ForegroundColor $(if ($crashOnlyPairs.Count  -gt 0) { 'Magenta' } else { 'Green' })
Write-Host ""
Write-Host "Reports:" -ForegroundColor Cyan
Write-Host "  $reportTxt" -ForegroundColor Gray
Write-Host "  $reportMd" -ForegroundColor Gray
Write-Host ""
Write-Host "Analysis performed by: $aiProvider" -ForegroundColor Cyan
