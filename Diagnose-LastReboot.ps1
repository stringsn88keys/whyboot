#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Diagnoses the last system reboot using event logs, AI analysis, and web search.
.DESCRIPTION
    Finds the last reboot timestamp, gathers event log entries within a configurable
    time window, sends the data to an AI for analysis (Claude → Copilot → Ollama
    in priority order), performs a DuckDuckGo web search for similar issues, and
    writes a diagnosis file.
.PARAMETER OllamaModel
    The ollama model to use for analysis. Defaults to the value in config.json
    (set by Setup-Whyboot.ps1), or "qwen3:4b" if no config exists.
.PARAMETER OllamaUrl
    The base URL for the ollama API. Defaults to the value in config.json,
    or "http://localhost:11434" if no config exists.
.PARAMETER WindowSeconds
    Number of seconds before and after the last boot time to search for events. Default is 10.
.EXAMPLE
    .\Diagnose-LastReboot.ps1
    .\Diagnose-LastReboot.ps1 -OllamaModel "mistral" -WindowSeconds 30
#>

param(
    [string]$OllamaModel,
    [string]$OllamaUrl,
    [int]$WindowSeconds = 10
)

$ErrorActionPreference = "SilentlyContinue"

# Load defaults from config.json (written by Setup-Whyboot.ps1)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$configPath = Join-Path $scriptDir "config.json"
$config = $null
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

# ── Helper: parse the actual crash time out of an Event 6008 message ──────────
# 6008 fires at the *next* boot; the crash timestamp is embedded in its message.
function Get-CrashTimeFrom6008 ([object]$Event) {
    $msg = $Event.Message -replace [char]0x200E,'' -replace [char]0x200F,'' `
                          -replace [char]0x202A,'' -replace [char]0x202B,'' `
                          -replace [char]0x202C,'' -replace [char]0x202D,'' `
                          -replace [char]0x202E,''
    if ($msg -match 'shutdown at (.+?) on (.+?) was unexpected') {
        try { return [DateTime]::Parse("$($Matches[2].Trim()) $($Matches[1].Trim())") } catch {}
    }
    if ($Event.Properties.Count -ge 2) {
        try { return [DateTime]::Parse("$($Event.Properties[1].Value) $($Event.Properties[0].Value)") } catch {}
    }
    return $Event.TimeCreated.AddMinutes(-5)
}

# Fetch the most recent Kernel-Power 41 early - used to decide the analysis window below
$kernelPower41 = Get-WinEvent -FilterHashtable @{
    LogName      = 'System'
    ProviderName = 'Microsoft-Windows-Kernel-Power'
    Id           = 41
} -MaxEvents 1 2>$null

# ── 1. Find the last reboot timestamp ──────────────────────────────────────────

Write-Host "`n===== DIAGNOSE LAST REBOOT =====" -ForegroundColor Cyan

$os = Get-CimInstance Win32_OperatingSystem
$lastBoot = $os.LastBootUpTime

Write-Host "Last boot time (WMI): $lastBoot" -ForegroundColor Gray

# Check Event ID 6005 (EventLog service started = boot completed)
$bootEvent = Get-WinEvent -FilterHashtable @{
    LogName   = 'System'
    Id        = 6005
} -MaxEvents 1 2>$null

if ($bootEvent) {
    Write-Host "EventLog service start (6005): $($bootEvent.TimeCreated)" -ForegroundColor Gray
}

# Check Event ID 6008 (unexpected shutdown)
$dirtyShutdown = Get-WinEvent -FilterHashtable @{
    LogName   = 'System'
    Id        = 6008
} -MaxEvents 1 2>$null

# Check Event ID 1074 (planned shutdown/restart)
$plannedShutdown = Get-WinEvent -FilterHashtable @{
    LogName   = 'System'
    Id        = 1074
} -MaxEvents 1 2>$null

$shutdownType = "Unknown"
if ($dirtyShutdown -and $plannedShutdown) {
    if ($dirtyShutdown.TimeCreated -gt $plannedShutdown.TimeCreated) {
        $shutdownType = "Unexpected (dirty)"
        Write-Host "Last shutdown was UNEXPECTED at $($dirtyShutdown.TimeCreated)" -ForegroundColor Red
    } else {
        $shutdownType = "Planned"
        $process = ($plannedShutdown.Properties[0]).Value
        $reason  = ($plannedShutdown.Properties[2]).Value
        Write-Host "Last shutdown was PLANNED at $($plannedShutdown.TimeCreated)" -ForegroundColor Green
        Write-Host "  Process: $process | Reason: $reason" -ForegroundColor Gray
    }
} elseif ($dirtyShutdown) {
    $shutdownType = "Unexpected (dirty)"
    Write-Host "Last shutdown was UNEXPECTED at $($dirtyShutdown.TimeCreated)" -ForegroundColor Red
} elseif ($plannedShutdown) {
    $shutdownType = "Planned"
    Write-Host "Last shutdown was PLANNED at $($plannedShutdown.TimeCreated)" -ForegroundColor Green
}

# ── 2. Gather events within the time window ────────────────────────────────────
#
# Kernel-Power 41 indicates the system was not cleanly shut down - most often the
# user force-reset an unresponsive PC.  In that case a narrow window around the
# current boot misses everything that led to the hang, so we instead analyse the
# full session from the previous boot up to the crash time.

$usedLastShutdownWindow = $false
$windowDescription      = ""

if ($kernelPower41 -and $dirtyShutdown) {
    $crashTime = Get-CrashTimeFrom6008 $dirtyShutdown
    $prevBoot  = Get-WinEvent -FilterHashtable @{
        LogName = 'System'
        Id      = 6005
    } -MaxEvents 100 2>$null |
        Where-Object { $_.TimeCreated -lt $crashTime } |
        Sort-Object TimeCreated -Descending |
        Select-Object -First 1

    if ($prevBoot) {
        $windowStart            = $prevBoot.TimeCreated
        $windowEnd              = $crashTime
        $usedLastShutdownWindow = $true
        $windowDescription      = "full session since last boot ($($prevBoot.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')) → crash $($crashTime.ToString('yyyy-MM-dd HH:mm:ss')))"
        Write-Host "`nKernel-Power 41 detected - likely unresponsive PC hard-reset by user." -ForegroundColor Magenta
        Write-Host "Analyzing full session since last boot instead of $WindowSeconds`s window..." -ForegroundColor Yellow
        Write-Host "  Session : $($prevBoot.TimeCreated)  →  crash : $crashTime" -ForegroundColor Gray
    } else {
        $windowStart       = $lastBoot.AddSeconds(-$WindowSeconds)
        $windowEnd         = $lastBoot.AddSeconds($WindowSeconds)
        $windowDescription = "$($WindowSeconds * 2)s window ($windowStart to $windowEnd)"
        Write-Host "`nGathering events within $WindowSeconds seconds of boot..." -ForegroundColor Yellow
    }
} else {
    $windowStart       = $lastBoot.AddSeconds(-$WindowSeconds)
    $windowEnd         = $lastBoot.AddSeconds($WindowSeconds)
    $windowDescription = "$($WindowSeconds * 2)s window ($windowStart to $windowEnd)"
    Write-Host "`nGathering events within $WindowSeconds seconds of boot..." -ForegroundColor Yellow
}

# General events (errors, warnings, critical) from System log
$systemEvents = Get-WinEvent -FilterHashtable @{
    LogName   = 'System'
    Level     = 1, 2, 3  # Critical, Error, Warning
    StartTime = $windowStart
    EndTime   = $windowEnd
} -MaxEvents 50 2>$null

# General events from Application log
$appEvents = Get-WinEvent -FilterHashtable @{
    LogName   = 'Application'
    Level     = 1, 2, 3
    StartTime = $windowStart
    EndTime   = $windowEnd
} -MaxEvents 50 2>$null

# Specific key events for the AI prompt (broader search, not limited to the window)
# $kernelPower41 was already fetched above for window-type determination.

$event6008 = Get-WinEvent -FilterHashtable @{
    LogName   = 'System'
    Id        = 6008
} -MaxEvents 1 2>$null

$event1074 = Get-WinEvent -FilterHashtable @{
    LogName   = 'System'
    Id        = 1074
} -MaxEvents 1 2>$null

$wheaErrors = Get-WinEvent -FilterHashtable @{
    LogName      = 'System'
    ProviderName = 'Microsoft-Windows-WHEA-Logger'
    StartTime    = $windowStart
    EndTime      = $windowEnd
} -MaxEvents 10 2>$null

# Combine and sort all window events
$allEvents = @()
if ($systemEvents) { $allEvents += $systemEvents }
if ($appEvents)    { $allEvents += $appEvents }
$allEvents = $allEvents | Sort-Object TimeCreated

$eventCount = $allEvents.Count
Write-Host "Found $eventCount events in the $windowDescription." -ForegroundColor Gray

# Format events into structured text
$eventBlock = ""
foreach ($evt in $allEvents) {
    $msg = ($evt.Message -replace "`r`n", " ").Trim()
    if ($msg.Length -gt 300) { $msg = $msg.Substring(0, 300) + "..." }
    $eventBlock += "[$($evt.TimeCreated)] [$($evt.LogName)] [$($evt.ProviderName)] ID=$($evt.Id) Level=$($evt.LevelDisplayName)`n"
    $eventBlock += "  $msg`n`n"
}

# Add key reboot-indicator events if they exist
$keyEventsBlock = ""
if ($kernelPower41) {
    $bugCheckCode = ($kernelPower41.Properties[0]).Value
    $keyEventsBlock += "KERNEL-POWER 41 (Critical): BugcheckCode=$bugCheckCode at $($kernelPower41.TimeCreated)`n"
}
if ($event6008) {
    $keyEventsBlock += "EVENT 6008 (Unexpected Shutdown): $($event6008.TimeCreated) - $($event6008.Message -replace "`r`n", " ")`n"
}
if ($event1074) {
    $process1074 = ($event1074.Properties[0]).Value
    $reason1074  = ($event1074.Properties[2]).Value
    $keyEventsBlock += "EVENT 1074 (Planned Shutdown): Process=$process1074, Reason=$reason1074 at $($event1074.TimeCreated)`n"
}
if ($wheaErrors) {
    foreach ($w in $wheaErrors) {
        $keyEventsBlock += "WHEA ERROR: $($w.TimeCreated) - $(($w.Message -split "`n")[0])`n"
    }
}

# ── 3. Query AI for analysis (Claude → Copilot → Ollama) ──────────────────────

Write-Host "`nQuerying AI for analysis (Claude → Copilot → Ollama)..." -ForegroundColor Yellow

$systemInfo = @"
Computer: $($os.CSName)
OS: $($os.Caption) $($os.Version)
Last Boot: $lastBoot
Shutdown Type: $shutdownType
"@

$ollamaPrompt = @"
You are a Windows system diagnostics expert. Analyze the following event log data from around the time of the last system reboot and provide a diagnosis.

SYSTEM INFO:
$systemInfo

KEY REBOOT-RELATED EVENTS:
$keyEventsBlock

ALL EVENTS IN WINDOW ($windowDescription):
$eventBlock

Based on these events, please provide:
1. The most likely cause of the reboot
2. Whether it was a clean or dirty shutdown
3. Any concerning patterns or errors
4. Recommended actions to prevent future unexpected reboots
"@

$aiResult      = Invoke-AIAnalysis -Prompt $ollamaPrompt -Config $config -OllamaModel $OllamaModel -OllamaUrl $OllamaUrl
$ollamaAnalysis = $aiResult.Response
$ollamaSuccess  = $aiResult.Success
$aiProvider     = $aiResult.Provider

if ($aiResult.Success) {
    Write-Host "Analysis complete via $aiProvider." -ForegroundColor Green
} else {
    Write-Host "AI analysis unavailable." -ForegroundColor Red
}

# ── 4. DuckDuckGo web search ──────────────────────────────────────────────────

Write-Host "`nSearching DuckDuckGo for similar issues..." -ForegroundColor Yellow

# Build search query from key findings
$searchTerms = @("Windows reboot")
if ($kernelPower41) {
    $searchTerms += "Kernel-Power 41"
    $bugCode = ($kernelPower41.Properties[0]).Value
    if ($bugCode -ne 0) { $searchTerms += "bugcheck 0x$($bugCode.ToString('X'))" }
}
if ($event6008) { $searchTerms += "Event 6008 unexpected shutdown" }
if ($wheaErrors) { $searchTerms += "WHEA hardware error" }

# Add any notable provider names from the window events
$notableProviders = $allEvents | Where-Object {
    $_.LevelDisplayName -eq 'Critical' -or $_.LevelDisplayName -eq 'Error'
} | Select-Object -ExpandProperty ProviderName -Unique | Select-Object -First 2
foreach ($prov in $notableProviders) {
    if ($prov -and $prov -notmatch 'EventLog|Service Control Manager') {
        $searchTerms += $prov
    }
}

$searchQuery = ($searchTerms | Select-Object -First 5) -join " "
Write-Host "Search query: $searchQuery" -ForegroundColor Gray

$webResults = ""
$webSuccess = $false

try {
    $encoded = [System.Uri]::EscapeDataString($searchQuery)
    $ddgResponse = Invoke-WebRequest -Uri "https://html.duckduckgo.com/html/?q=$encoded" -UseBasicParsing -TimeoutSec 30

    # Parse result links and snippets
    $resultLinks = @()
    $resultSnippets = @()

    # Extract result titles/links
    $linkMatches = [regex]::Matches($ddgResponse.Content, '<a[^>]*class="result__a"[^>]*href="([^"]*)"[^>]*>(.*?)</a>')
    foreach ($m in $linkMatches) {
        $rawUrl = $m.Groups[1].Value
        # DDG wraps links in a redirect; extract the actual URL from the uddg parameter
        if ($rawUrl -match 'uddg=([^&]+)') {
            $rawUrl = [System.Uri]::UnescapeDataString($Matches[1])
        }
        $resultLinks += @{
            Url   = $rawUrl -replace '&amp;', '&'
            Title = ($m.Groups[2].Value -replace '<[^>]+>', '').Trim()
        }
    }

    # Extract snippets
    $snippetMatches = [regex]::Matches($ddgResponse.Content, '<a[^>]*class="result__snippet"[^>]*>(.*?)</a>')
    foreach ($m in $snippetMatches) {
        $resultSnippets += ($m.Groups[1].Value -replace '<[^>]+>', '' -replace '&#x27;', "'" -replace '&amp;', '&').Trim()
    }

    $resultCount = [Math]::Min($resultLinks.Count, 5)
    if ($resultCount -gt 0) {
        $webSuccess = $true
        for ($i = 0; $i -lt $resultCount; $i++) {
            $snippet = if ($i -lt $resultSnippets.Count) { $resultSnippets[$i] } else { "" }
            $webResults += "$($i + 1). $($resultLinks[$i].Title)`n"
            $webResults += "   URL: $($resultLinks[$i].Url)`n"
            if ($snippet) { $webResults += "   $snippet`n" }
            $webResults += "`n"
        }
        Write-Host "Found $resultCount web results." -ForegroundColor Green
    } else {
        $webResults = "(No search results found)"
        Write-Host "No search results found." -ForegroundColor Gray
    }
} catch {
    $webResults = "(Web search failed: $_)"
    Write-Host "DuckDuckGo search failed: $_" -ForegroundColor Red
}

# Feed web results to AI for synthesis if both succeeded
$webSynthesis = ""
if ($ollamaSuccess -and $webSuccess) {
    Write-Host "Synthesizing web results with AI..." -ForegroundColor Yellow
    $synthesisPrompt = @"
Based on your earlier analysis of the reboot event logs and these web search results for similar issues, provide a brief synthesis of what the web community says about this type of reboot issue and any additional recommendations.

WEB SEARCH RESULTS:
$webResults

Your earlier analysis concluded: $($ollamaAnalysis.Substring(0, [Math]::Min(500, $ollamaAnalysis.Length)))

Provide a concise synthesis (3-5 bullet points) of relevant web findings and how they relate to this specific case.
"@
    $synthResult  = Invoke-AIAnalysis -Prompt $synthesisPrompt -Config $config -OllamaModel $OllamaModel -OllamaUrl $OllamaUrl
    $webSynthesis = $synthResult.Response
    if ($synthResult.Success) {
        Write-Host "Web synthesis complete via $($synthResult.Provider)." -ForegroundColor Green
    } else {
        Write-Host "Web synthesis unavailable." -ForegroundColor Red
    }
}

# ── 5. Write diagnosis file ───────────────────────────────────────────────────

$rebootTimestamp = $lastBoot.ToString("yyyyMMdd-HHmmss")
$outputFile = Join-Path (Get-Location) "reboot-$rebootTimestamp-diagnosis.txt"

$uptime = (Get-Date) - $lastBoot

$report = @"
================================================================================
  REBOOT DIAGNOSIS REPORT
  Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
================================================================================

SYSTEM INFORMATION
  Computer:      $($os.CSName)
  OS:            $($os.Caption) $($os.Version)
  Last Boot:     $lastBoot
  Current Uptime: $($uptime.Days)d $($uptime.Hours)h $($uptime.Minutes)m
  Shutdown Type: $shutdownType

================================================================================
  RAW EVENT LOG DATA ($windowDescription)
================================================================================

KEY REBOOT-RELATED EVENTS:
$(if ($keyEventsBlock) { $keyEventsBlock } else { "(None found)" })

ALL EVENTS IN WINDOW ($eventCount events):
$(if ($eventBlock) { $eventBlock } else { "(No events found in window)" })

================================================================================
  AI ANALYSIS (provider: $aiProvider)
================================================================================

$ollamaAnalysis

================================================================================
  WEB SEARCH FINDINGS
================================================================================

Search Query: $searchQuery

$webResults

$(if ($webSynthesis) { @"
--- Web Synthesis ---
$webSynthesis
"@ })

================================================================================
  COMBINED DIAGNOSIS SUMMARY
================================================================================

Reboot Time:    $lastBoot
Shutdown Type:  $shutdownType
Events Found:   $eventCount events ($windowDescription)
AI Status:      $(if ($ollamaSuccess) { "Analysis complete ($aiProvider)" } else { "Unavailable" })
Web Search:     $(if ($webSuccess) { "Results found" } else { "Unavailable" })

$(if ($ollamaSuccess) { @"
AI DIAGNOSIS:
$ollamaAnalysis
"@ } else { @"
NOTE: AI analysis was unavailable. Configure a Claude API key (ANTHROPIC_API_KEY),
a GitHub Copilot token (GITHUB_TOKEN), or start Ollama ('ollama serve') for analysis.
"@ })
"@

$report | Out-File -FilePath $outputFile -Encoding UTF8
Write-Host "`nDiagnosis written to: $outputFile" -ForegroundColor Green

# Also write a markdown version
$mdFile = Join-Path (Get-Location) "reboot-$rebootTimestamp-diagnosis.md"

$mdReport = @"
# Reboot Diagnosis Report
**Generated:** $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

## System Information
| Field | Value |
|-------|-------|
| Computer | $($os.CSName) |
| OS | $($os.Caption) $($os.Version) |
| Last Boot | $lastBoot |
| Current Uptime | $($uptime.Days)d $($uptime.Hours)h $($uptime.Minutes)m |
| Shutdown Type | $shutdownType |

## Raw Event Log Data
**Window:** $windowDescription

### Key Reboot-Related Events
$(if ($keyEventsBlock) { $keyEventsBlock } else { "*(None found)*" })

### All Events in Window ($eventCount events)
``````
$(if ($eventBlock) { $eventBlock.TrimEnd() } else { "(No events found in window)" })
``````

## AI Analysis

*Provider: $aiProvider*

$ollamaAnalysis

## Web Search Findings
**Search Query:** ``$searchQuery``

$webResults

$(if ($webSynthesis) { @"
### Web Synthesis
$webSynthesis
"@ })

## Combined Diagnosis Summary
| Field | Value |
|-------|-------|
| Reboot Time | $lastBoot |
| Shutdown Type | $shutdownType |
| Events Found | $eventCount events ($windowDescription) |
| AI Status | $(if ($ollamaSuccess) { "Analysis complete ($aiProvider)" } else { "Unavailable" }) |
| Web Search | $(if ($webSuccess) { "Results found" } else { "Unavailable" }) |

$(if ($ollamaSuccess) { @"
### AI Diagnosis
$ollamaAnalysis
"@ } else { @"
> **Note:** AI analysis was unavailable. Configure a Claude API key (`ANTHROPIC_API_KEY`),
> a GitHub Copilot token (`GITHUB_TOKEN`), or start Ollama (`ollama serve`) for analysis.
"@ })
"@

$mdReport | Out-File -FilePath $mdFile -Encoding UTF8
Write-Host "Markdown version: $mdFile" -ForegroundColor Green
Write-Host ""
Write-Host "Analysis performed by: $aiProvider" -ForegroundColor Cyan
