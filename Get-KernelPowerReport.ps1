#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Queries the last N Kernel-Power events and writes a self-contained HTML report.
.DESCRIPTION
    Collects Microsoft-Windows-Kernel-Power events from the System log, then
    produces an HTML report containing:
      - Summary statistics
      - Event ID breakdown table
      - Histogram: events by hour of day (0-23)
      - Histogram: events by day of week
      - Histogram: events by calendar week
      - Full scrollable event listing
.PARAMETER Count
    Maximum number of events to retrieve. Default: 100.
.PARAMETER OutputPath
    Path for the HTML report. Defaults to kernel-power-report.html in the
    current directory.
.PARAMETER NoBrowser
    Skip opening the report in the default browser after writing.
.EXAMPLE
    .\Get-KernelPowerReport.ps1
    .\Get-KernelPowerReport.ps1 -Count 250 -OutputPath C:\Temp\power.html
#>
param(
    [int]    $Count      = 100,
    [string] $OutputPath = "",
    [switch] $NoBrowser
)

$ErrorActionPreference = "SilentlyContinue"

if (-not $OutputPath) {
    $OutputPath = Join-Path (Get-Location) "kernel-power-report.html"
}

Write-Host "`n===== KERNEL POWER EVENT REPORT =====" -ForegroundColor Cyan
Write-Host "Querying last $Count Kernel-Power events..." -ForegroundColor Yellow

# =============================================================================
# 1. Collect events
# =============================================================================
$events = Get-WinEvent -FilterHashtable @{
    LogName      = 'System'
    ProviderName = 'Microsoft-Windows-Kernel-Power'
} -MaxEvents $Count -ErrorAction SilentlyContinue | Sort-Object TimeCreated

if (-not $events -or $events.Count -eq 0) {
    Write-Host "No Kernel-Power events found." -ForegroundColor Red
    exit 1
}

Write-Host "Found $($events.Count) events." -ForegroundColor Green

# =============================================================================
# 2. Compute histograms
# =============================================================================

# Hour of day (0-23)
$hourBuckets = @{}
0..23 | ForEach-Object { $hourBuckets[$_] = 0 }
foreach ($e in $events) { $hourBuckets[$e.TimeCreated.Hour]++ }

# Day of week (0=Sunday .. 6=Saturday)
$dowBuckets = @{}
0..6 | ForEach-Object { $dowBuckets[$_] = 0 }
foreach ($e in $events) { $dowBuckets[[int]$e.TimeCreated.DayOfWeek]++ }
$dowNames = @('Sun','Mon','Tue','Wed','Thu','Fri','Sat')

# Calendar week - keyed by the Monday that starts each ISO week
$weekBuckets = [System.Collections.Generic.SortedDictionary[datetime,int]]::new()
foreach ($e in $events) {
    $dow    = [int]$e.TimeCreated.DayOfWeek          # 0=Sun
    $offset = if ($dow -eq 0) { -6 } else { 1 - $dow }
    $monday = $e.TimeCreated.Date.AddDays($offset)
    if ($weekBuckets.ContainsKey($monday)) { $weekBuckets[$monday]++ }
    else { $weekBuckets[$monday] = 1 }
}
# Fill any missing weeks between first and last with zero counts
$firstMonday = $weekBuckets.Keys | Select-Object -First 1
$lastMonday  = $weekBuckets.Keys | Select-Object -Last  1
$cursor = $firstMonday
while ($cursor -le $lastMonday) {
    if (-not $weekBuckets.ContainsKey($cursor)) { $weekBuckets[$cursor] = 0 }
    $cursor = $cursor.AddDays(7)
}
$weekLabels = $weekBuckets.Keys | ForEach-Object { $_.ToString('MM/dd') }
$weekValues = $weekBuckets.Values | ForEach-Object { [int]$_ }
$peakWeek   = ($weekBuckets.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key

# Event ID
$idBuckets = @{}
foreach ($e in $events) {
    if (-not $idBuckets.ContainsKey($e.Id)) { $idBuckets[$e.Id] = 0 }
    $idBuckets[$e.Id]++
}
$idSorted = $idBuckets.GetEnumerator() | Sort-Object Key

# =============================================================================
# 3. Known event ID descriptions
# =============================================================================
$knownIds = @{
    41  = "System rebooted without clean shutdown (Kernel-Power 41)"
    42  = "Entering sleep / hibernate"
    105 = "Power source changed (AC/battery)"
    107 = "Resumed from sleep"
    131 = "Boot configuration change"
    137 = "NTP synchronisation"
    172 = "Kernel power policy change"
    506 = "Connected standby entry"
    507 = "Connected standby exit"
}

# =============================================================================
# 4. SVG bar chart helper
# =============================================================================
function New-SvgBarChart {
    param(
        [int[]]   $Values,
        [string[]]$Labels,
        [string]  $BarColor     = "#4f93d8",
        [string]  $XCaption     = "",
        [switch]  $RotateLabels            # use for charts with many narrow bars
    )

    $svgW    = 880
    $svgH    = if ($RotateLabels) { 290 } else { 260 }
    $padL    = 48
    $padR    = 16
    $padT    = 18
    $padB    = if ($RotateLabels) { 80 } else { 52 }

    $chartW  = $svgW - $padL - $padR
    $chartH  = $svgH - $padT - $padB

    $n       = $Labels.Count
    $maxVal  = ($Values | Measure-Object -Maximum).Maximum
    if ($maxVal -lt 1) { $maxVal = 1 }

    $slotW   = $chartW / $n
    $barW    = [Math]::Max(4, [Math]::Floor($slotW * 0.65))

    # Y-axis ticks: 4 intervals
    $yTicks  = 4
    $tickStep = [Math]::Ceiling($maxVal / $yTicks)
    $yMax    = $tickStep * $yTicks

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<svg xmlns='http://www.w3.org/2000/svg' width='$svgW' height='$svgH' ")
    [void]$sb.Append("style='font-family:Consolas,monospace;font-size:11px;display:block'>")

    # Chart background
    [void]$sb.Append("<rect x='0' y='0' width='$svgW' height='$svgH' fill='#0f0f1e' rx='6'/>")

    # Gridlines + Y labels
    for ($g = 0; $g -le $yTicks; $g++) {
        $yVal = $tickStep * $g
        $y    = $padT + [int]($chartH - ($chartH * $yVal / $yMax))
        [void]$sb.Append("<line x1='$padL' y1='$y' x2='$($padL + $chartW)' y2='$y' stroke='#222240' stroke-width='1'/>")
        [void]$sb.Append("<text x='$($padL - 6)' y='$($y + 4)' fill='#555' text-anchor='end'>$yVal</text>")
    }

    # Bars + X labels
    for ($i = 0; $i -lt $n; $i++) {
        $v    = $Values[$i]
        $barH = if ($yMax -gt 0) { [int]($chartH * $v / $yMax) } else { 0 }
        $cx   = $padL + [int]($slotW * $i + $slotW / 2)
        $bx   = $cx - [int]($barW / 2)
        $by   = $padT + $chartH - $barH

        if ($barH -gt 0) {
            $opacity = if ($v -gt 0) { "dd" } else { "33" }
            [void]$sb.Append("<rect x='$bx' y='$by' width='$barW' height='$barH' fill='$BarColor$opacity' rx='2'/>")
            [void]$sb.Append("<text x='$cx' y='$($by - 3)' fill='#bbb' text-anchor='middle'>$v</text>")
        }

        if ($RotateLabels) {
            $lx = $cx; $ly = $padT + $chartH + 12
            [void]$sb.Append("<text x='$lx' y='$ly' fill='#777' text-anchor='end' ")
            [void]$sb.Append("transform='rotate(-45 $lx $ly)'>$($Labels[$i])</text>")
        } else {
            [void]$sb.Append("<text x='$cx' y='$($padT + $chartH + 14)' fill='#777' text-anchor='middle'>$($Labels[$i])</text>")
        }
    }

    # Axes
    [void]$sb.Append("<line x1='$padL' y1='$padT' x2='$padL' y2='$($padT + $chartH)' stroke='#444' stroke-width='1'/>")
    [void]$sb.Append("<line x1='$padL' y1='$($padT + $chartH)' x2='$($padL + $chartW)' y2='$($padT + $chartH)' stroke='#444' stroke-width='1'/>")

    if ($XCaption) {
        $cx = $padL + [int]($chartW / 2)
        [void]$sb.Append("<text x='$cx' y='$($svgH - 4)' fill='#555' text-anchor='middle' font-size='11'>$XCaption</text>")
    }

    [void]$sb.Append("</svg>")
    return $sb.ToString()
}

# Build hour chart
$hourValues = 0..23 | ForEach-Object { $hourBuckets[$_] }
$hourLabels = 0..23 | ForEach-Object { $_.ToString("00") }
$hourChart  = New-SvgBarChart -Values $hourValues -Labels $hourLabels `
                              -BarColor "#4f93d8" -XCaption "Hour of Day (00 = midnight)"

# Build day-of-week chart
$dowValues = 0..6 | ForEach-Object { $dowBuckets[$_] }
$dowChart  = New-SvgBarChart -Values $dowValues -Labels $dowNames `
                             -BarColor "#7ed8a4" -XCaption "Day of Week"

# Build weekly chart (rotate labels when there are more than 10 weeks)
$weekChart = New-SvgBarChart -Values $weekValues -Labels $weekLabels `
                             -BarColor "#c47ed8" -XCaption "Week starting (MM/DD)" `
                             -RotateLabels:($weekLabels.Count -gt 10)

# =============================================================================
# 5. Build HTML fragments
# =============================================================================
$levelColors = @{
    'Critical'    = '#e05252'
    'Error'       = '#e07c52'
    'Warning'     = '#d4a843'
    'Information' = '#5ba8d4'
    'Verbose'     = '#777'
}

# Event ID summary rows
$idRows = [System.Text.StringBuilder]::new()
foreach ($kv in $idSorted) {
    $id   = [int]$kv.Key
    $cnt  = $kv.Value
    $desc = if ($knownIds.ContainsKey($id)) { $knownIds[$id] } else { "-" }
    $pct  = [Math]::Round($cnt / $events.Count * 100, 1)
    [void]$idRows.Append("<tr>")
    [void]$idRows.Append("<td class='mono accent'>$id</td>")
    [void]$idRows.Append("<td>$desc</td>")
    [void]$idRows.Append("<td class='mono right teal'>$cnt</td>")
    [void]$idRows.Append("<td class='right dim'>$pct%</td>")
    [void]$idRows.Append("</tr>`n")
}

# All-events table rows (newest first)
$eventRows = [System.Text.StringBuilder]::new()
foreach ($e in ($events | Sort-Object TimeCreated -Descending)) {
    $lvl   = $e.LevelDisplayName
    $color = if ($levelColors.ContainsKey($lvl)) { $levelColors[$lvl] } else { "#aaa" }
    $raw   = $e.Message -replace "<","&lt;" -replace ">","&gt;"
    $first = ($raw -split "`n")[0].Trim()
    if ($first.Length -gt 180) { $first = $first.Substring(0,180) + "..." }
    $desc  = if ($knownIds.ContainsKey($e.Id)) { $knownIds[$e.Id] } else { "" }
    [void]$eventRows.Append("<tr>")
    [void]$eventRows.Append("<td class='mono dim nowrap'>$($e.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'))</td>")
    [void]$eventRows.Append("<td class='mono accent center'>$($e.Id)</td>")
    [void]$eventRows.Append("<td class='small dim'>$desc</td>")
    [void]$eventRows.Append("<td class='nowrap' style='color:$color'>$lvl</td>")
    [void]$eventRows.Append("<td class='small'>$first</td>")
    [void]$eventRows.Append("</tr>`n")
}

# Summary stats
$earliest   = ($events | Select-Object -First 1).TimeCreated
$latest     = ($events | Select-Object -Last  1).TimeCreated
$span       = $latest - $earliest
$spanStr    = "$($span.Days)d $($span.Hours)h $($span.Minutes)m"
$peakHour   = ($hourBuckets.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
$peakDow    = ($dowBuckets.GetEnumerator()  | Sort-Object Value -Descending | Select-Object -First 1).Key
$peakDowStr = $dowNames[$peakDow]
$os         = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue

# =============================================================================
# 6. Assemble HTML
# =============================================================================
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Kernel Power Report - $($os.CSName)</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0d0d1a;color:#c8cfe0;font-family:'Segoe UI',sans-serif;font-size:14px;line-height:1.5;padding:28px 32px}
a{color:#4f93d8}
h1{color:#7ec8e8;font-size:1.55em;font-weight:600;margin-bottom:4px}
h2{color:#7ec8e8;font-size:1.0em;font-weight:600;text-transform:uppercase;letter-spacing:1px;margin:32px 0 12px;border-bottom:1px solid #1e1e3a;padding-bottom:6px}
.meta{color:#555;font-size:0.82em;margin-bottom:28px}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px;margin-bottom:28px}
.card{background:#131325;border:1px solid #1e1e3a;border-radius:7px;padding:14px 16px}
.card .lbl{color:#555;font-size:0.75em;text-transform:uppercase;letter-spacing:1px}
.card .val{color:#dde8f0;font-size:1.35em;font-weight:700;margin-top:3px}
.card .val.sm{font-size:0.95em}
.chart-box{background:#131325;border:1px solid #1e1e3a;border-radius:7px;padding:18px;margin-bottom:24px}
.chart-box .cap{color:#778;font-size:0.8em;margin-bottom:10px}
.tbl-box{background:#131325;border:1px solid #1e1e3a;border-radius:7px;overflow:auto;margin-bottom:24px}
.tbl-box.scroll{max-height:540px}
table{width:100%;border-collapse:collapse}
th{background:#0a0a14;color:#556;font-size:0.78em;text-transform:uppercase;letter-spacing:1px;padding:9px 12px;text-align:left;position:sticky;top:0;font-weight:normal}
td{padding:7px 12px;border-bottom:1px solid #141428;vertical-align:top}
tr:last-child td{border-bottom:none}
tr:hover td{background:#161630}
.mono{font-family:Consolas,monospace}
.accent{color:#e0a84f}
.teal{color:#4fd4a4}
.dim{color:#666}
.small{font-size:0.82em}
.center{text-align:center}
.right{text-align:right}
.nowrap{white-space:nowrap}
</style>
</head>
<body>

<h1>Kernel Power Event Report</h1>
<div class="meta">
  Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
  &nbsp;&bull;&nbsp; System: $($os.CSName)
  &nbsp;&bull;&nbsp; OS: $($os.Caption) $($os.Version)
</div>

<div class="cards">
  <div class="card"><div class="lbl">Events</div><div class="val">$($events.Count)</div></div>
  <div class="card"><div class="lbl">Unique IDs</div><div class="val">$($idBuckets.Count)</div></div>
  <div class="card"><div class="lbl">Date Range</div><div class="val sm">$($earliest.ToString('yyyy-MM-dd'))</div></div>
  <div class="card"><div class="lbl">Span</div><div class="val sm">$spanStr</div></div>
  <div class="card"><div class="lbl">Peak Hour</div><div class="val">$('{0:D2}' -f $peakHour):00</div></div>
  <div class="card"><div class="lbl">Peak Day</div><div class="val">$peakDowStr</div></div>
  <div class="card"><div class="lbl">Peak Week</div><div class="val sm">$($peakWeek.ToString('yyyy-MM-dd'))</div></div>
</div>

<h2>Event ID Breakdown</h2>
<div class="tbl-box">
<table>
<thead><tr><th>ID</th><th>Description</th><th style="text-align:right">Count</th><th style="text-align:right">%</th></tr></thead>
<tbody>$($idRows.ToString())</tbody>
</table>
</div>

<h2>Histogram - Hour of Day</h2>
<div class="chart-box">
<div class="cap">Events by clock hour (00 = midnight). Reveals whether events cluster at boot, sleep, or business hours.</div>
$hourChart
</div>

<h2>Histogram - Day of Week</h2>
<div class="chart-box">
<div class="cap">Events by day of week. Persistent spikes on specific days may indicate scheduled tasks or workload patterns.</div>
$dowChart
</div>

<h2>Histogram - Weekly Event Count</h2>
<div class="chart-box">
<div class="cap">Total events per calendar week (Mon-Sun). Each bar is one week; the label shows the Monday start date. Gaps appear as zero-count bars.</div>
$weekChart
</div>

<h2>All Events (newest first)</h2>
<div class="tbl-box scroll">
<table>
<thead><tr><th>Timestamp</th><th>ID</th><th>Description</th><th>Level</th><th>Message</th></tr></thead>
<tbody>$($eventRows.ToString())</tbody>
</table>
</div>

</body>
</html>
"@

$html | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "Report written to: $OutputPath" -ForegroundColor Green

if (-not $NoBrowser) {
    try { Start-Process $OutputPath } catch {}
}
