#Requires -Version 5.1
<#
.SYNOPSIS
    Claude Code Stop hook: validates .ps1 files for encoding issues and variable case consistency.
.DESCRIPTION
    Checks every .ps1 file in the project for:
      1. Problematic Unicode characters that break the PowerShell parser when the file
         is read as CP-1252 (em-dash, smart quotes, non-breaking space, etc.)
      2. The same variable name used with inconsistent casing across the file
         (e.g. $MyVar vs $myvar vs $MYVAR)
    On failure, writes a description of each issue to stderr and exits 2 so Claude
    fixes the problems before the task is considered complete.
#>

# ── Read stdin JSON (provided by Claude Code) ────────────────────────────────
$inputJson = $null
try {
    $raw = [Console]::In.ReadToEnd()
    if ($raw.Trim()) { $inputJson = $raw | ConvertFrom-Json }
} catch {}

# Prevent infinite loops: if Claude is already retrying because of this hook, let it stop.
if ($inputJson -and $inputJson.stop_hook_active) { exit 0 }

$projectDir = if ($inputJson -and $inputJson.cwd) { $inputJson.cwd } else { Get-Location }

$ps1Files = Get-ChildItem -Path $projectDir -Filter "*.ps1" -File -ErrorAction SilentlyContinue
$issues   = [System.Collections.Generic.List[string]]::new()

# Built-in / automatic PowerShell variables — skip these for the case-check.
$skipVars = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(
        '_', 'true', 'false', 'null', 'this', 'input', 'args', 'error',
        'host', 'home', 'pid', 'profile', 'pshome', 'psitem',
        'lastexitcode', 'matches', 'myinvocation', 'psscriptroot',
        'pscommandpath', 'psboundparameters', 'pscmdlet',
        'stacktrace', 'shellid', 'ofs', 'foreach', 'switch',
        'executioncontext', 'verbosepreference', 'debugpreference',
        'erroractionpreference', 'warningpreference', 'informationpreference'
    ),
    [System.StringComparer]::OrdinalIgnoreCase
)

foreach ($file in $ps1Files) {
    $content = Get-Content -Path $file.FullName -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
    if (-not $content) { continue }
    $lines = $content -split "`n"

    # ── Check 1: Characters that are invalid / problematic in CP-1252 ─────────
    # These look fine in a UTF-8 editor but cause parse errors when PowerShell
    # reads the file without a BOM (falls back to the system ANSI codepage).
    $badChars = [ordered]@{
        [char]0x2014 = 'em-dash U+2014          -> use plain hyphen -'
        [char]0x2013 = 'en-dash U+2013          -> use plain hyphen -'
        [char]0x201C = 'left double-quote U+201C -> use straight quote "'
        [char]0x201D = 'right double-quote U+201D-> use straight quote "'
        [char]0x2018 = "left single-quote U+2018 -> use straight apostrophe '"
        [char]0x2019 = "right single-quote U+2019-> use straight apostrophe '"
        [char]0x2026 = 'ellipsis U+2026         -> use three dots ...'
        [char]0x00A0 = 'non-breaking space U+00A0-> use regular space'
    }

    foreach ($ch in $badChars.Keys) {
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i].IndexOf($ch) -ge 0) {
                $issues.Add("  $($file.Name):$($i + 1): $($badChars[$ch])")
            }
        }
    }

    # ── Check 2: Inconsistent variable casing ─────────────────────────────────
    # Match $VarName but exclude scope-prefixed vars ($env:, $global:, etc.)
    # and single-char specials ($_, $?, $^).
    $varMatches = [regex]::Matches($content, '\$([A-Za-z][A-Za-z0-9_]+)(?!:)')
    $byLower    = @{}

    foreach ($m in $varMatches) {
        $name = $m.Groups[1].Value
        $key  = $name.ToLower()
        if ($skipVars.Contains($key)) { continue }
        if (-not $byLower.ContainsKey($key)) {
            $byLower[$key] = [System.Collections.Generic.HashSet[string]]::new()
        }
        [void]$byLower[$key].Add($name)
    }

    foreach ($key in $byLower.Keys) {
        if ($byLower[$key].Count -gt 1) {
            $forms = ($byLower[$key] | Sort-Object | ForEach-Object { '$' + $_ }) -join ', '
            $issues.Add("  $($file.Name): variable case mismatch: $forms")
        }
    }
}

if ($issues.Count -gt 0) {
    $msg = @(
        "PS1 quality issues found - please fix before finishing:",
        ""
    ) + $issues + @("")
    [Console]::Error.WriteLine($msg -join "`n")
    exit 2
}

exit 0
