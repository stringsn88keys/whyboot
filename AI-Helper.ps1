<#
.SYNOPSIS
    Shared AI analysis helper. Tries Claude Code CLI → Claude API → Copilot API → Ollama.
.DESCRIPTION
    Dot-source this file and call Invoke-AIAnalysis with a prompt and config object.
    Returns a hashtable with Response, Provider, and Success fields.

    Priority order:
      1. Claude Code CLI  (`claude -p`)  - uses existing Claude Code auth, no key needed
      2. Claude API       - requires ClaudeApiKey in config.json or ANTHROPIC_API_KEY env var
      3. Copilot API      - requires CopilotToken in config.json or GITHUB_TOKEN env var
      4. Ollama           - local fallback, requires `ollama serve` to be running

    Note: `gh copilot` CLI is intentionally skipped - its `explain`/`suggest` subcommands
    are designed for shell command help only, not general text analysis.
#>

function Invoke-AIAnalysis {
    param(
        [string]$Prompt,
        [object]$Config,
        [string]$OllamaModel = "qwen3:4b",
        [string]$OllamaUrl   = "http://localhost:11434"
    )

    $result = @{ Response = ""; Provider = "none"; Success = $false }

    # ── 1. Claude Code CLI ────────────────────────────────────────────────────
    # Uses `claude -p` (print/pipe mode). No API key needed -leverages existing
    # Claude Code auth. Prompt is written to a temp file to safely handle long inputs.
    #
    # When running as Administrator the PATH is stripped of user-profile dirs, so
    # Get-Command may not find claude even though it is installed. Fall back to the
    # known per-user install locations if the PATH lookup fails.
    $claudeCmd = Get-Command claude -ErrorAction SilentlyContinue
    if (-not $claudeCmd) {
        # When elevated (RunAsAdministrator), PATH drops user-profile entries.
        # Build a candidate list: per-env paths + every user profile under C:\Users.
        $candidates = [System.Collections.Generic.List[string]]::new()
        foreach ($suffix in @('.local\bin\claude.exe', 'AppData\Roaming\npm\claude.cmd',
                              'AppData\Local\Programs\claude\claude.exe')) {
            $candidates.Add("$env:USERPROFILE\$suffix")
        }
        foreach ($prof in (Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue)) {
            $candidates.Add("$($prof.FullName)\.local\bin\claude.exe")
            $candidates.Add("$($prof.FullName)\AppData\Roaming\npm\claude.cmd")
        }
        foreach ($c in $candidates) {
            if (Test-Path $c -ErrorAction SilentlyContinue) {
                $claudeCmd = [PSCustomObject]@{ Source = $c }
                Write-Host "  Found claude at: $c" -ForegroundColor Gray
                break
            }
        }
        if (-not $claudeCmd) {
            Write-Host "  claude not found in PATH or common install locations." -ForegroundColor DarkYellow
        }
    }
    if ($claudeCmd) {
        Write-Host "  Trying Claude Code CLI..." -ForegroundColor Gray
        $tmpPrompt = [System.IO.Path]::GetTempFileName()
        try {
            $Prompt | Out-File -FilePath $tmpPrompt -Encoding UTF8
            # Clear CLAUDECODE so this works even when called from inside a Claude Code session
            $savedClaudeCode = $env:CLAUDECODE
            $env:CLAUDECODE = $null
            $output = Get-Content $tmpPrompt | & $claudeCmd.Source -p --no-session-persistence 2>&1
            $env:CLAUDECODE = $savedClaudeCode
            if ($LASTEXITCODE -eq 0 -and $output) {
                $result.Response = $output -join "`n"
                $result.Provider = "Claude Code CLI"
                $result.Success  = $true
                return $result
            }
            Write-Host "  Claude Code CLI returned no output (exit $LASTEXITCODE)." -ForegroundColor DarkYellow
        } catch {
            Write-Host "  Claude Code CLI failed: $_" -ForegroundColor DarkYellow
        } finally {
            $env:CLAUDECODE = $savedClaudeCode
            Remove-Item $tmpPrompt -ErrorAction SilentlyContinue
        }
    }

    # ── 2. Claude API ─────────────────────────────────────────────────────────
    $claudeKey = if ($Config -and $Config.ClaudeApiKey) { $Config.ClaudeApiKey } `
                 else { $env:ANTHROPIC_API_KEY }
    if ($claudeKey) {
        $claudeModel = if ($Config -and $Config.ClaudeModel) { $Config.ClaudeModel } `
                       else { "claude-opus-4-6" }
        Write-Host "  Trying Claude API ($claudeModel)..." -ForegroundColor Gray
        try {
            $body = @{
                model      = $claudeModel
                max_tokens = 4096
                messages   = @(@{ role = "user"; content = $Prompt })
            } | ConvertTo-Json -Depth 10

            $headers = @{
                "x-api-key"         = $claudeKey
                "anthropic-version" = "2023-06-01"
            }
            $resp = Invoke-RestMethod -Uri "https://api.anthropic.com/v1/messages" `
                        -Method Post -Headers $headers -Body $body `
                        -ContentType "application/json" -TimeoutSec 120
            $result.Response = $resp.content[0].text
            $result.Provider = "Claude API ($claudeModel)"
            $result.Success  = $true
            return $result
        } catch {
            Write-Host "  Claude API unavailable: $_" -ForegroundColor DarkYellow
        }
    }

    # ── 3. Copilot API ────────────────────────────────────────────────────────
    $copilotToken = if ($Config -and $Config.CopilotToken) { $Config.CopilotToken } `
                    else { $env:GITHUB_TOKEN }
    if ($copilotToken) {
        Write-Host "  Trying GitHub Copilot API..." -ForegroundColor Gray
        try {
            $body = @{
                model    = "gpt-4o"
                messages = @(@{ role = "user"; content = $Prompt })
            } | ConvertTo-Json -Depth 10

            $headers = @{
                "Authorization"          = "Bearer $copilotToken"
                "Copilot-Integration-Id" = "vscode-chat"
                "Editor-Version"         = "vscode/1.85.1"
            }
            $resp = Invoke-RestMethod -Uri "https://api.githubcopilot.com/chat/completions" `
                        -Method Post -Headers $headers -Body $body `
                        -ContentType "application/json" -TimeoutSec 120
            $result.Response = $resp.choices[0].message.content
            $result.Provider = "GitHub Copilot API (gpt-4o)"
            $result.Success  = $true
            return $result
        } catch {
            Write-Host "  Copilot API unavailable: $_" -ForegroundColor DarkYellow
        }
    }

    # ── 4. Ollama (local fallback) ────────────────────────────────────────────
    Write-Host "  Trying Ollama ($OllamaModel at $OllamaUrl)..." -ForegroundColor Gray
    try {
        $body = @{
            model  = $OllamaModel
            prompt = $Prompt
            stream = $false
        } | ConvertTo-Json -Depth 10

        $resp = Invoke-RestMethod -Uri "$OllamaUrl/api/generate" -Method Post `
                    -Body $body -ContentType "application/json" -TimeoutSec 300
        $result.Response = $resp.response
        $result.Provider = "Ollama ($OllamaModel)"
        $result.Success  = $true
    } catch {
        $result.Response = "(AI analysis unavailable - all providers failed. Last error: $_)"
        $result.Provider = "none"
        Write-Host "  Ollama unavailable: $_" -ForegroundColor Red
        Write-Host "  To enable analysis: install Claude Code, add a Claude/Copilot API key to config.json, or run 'ollama serve'." -ForegroundColor Gray
    }

    return $result
}
