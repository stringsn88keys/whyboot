# whyboot

Windows reboot diagnostic toolkit. Answers the question: *why did my PC just reboot?*

whyboot combines Windows Event Log analysis, a local LLM (via [ollama](https://ollama.com)), and web search to produce a detailed diagnosis report every time your system reboots unexpectedly.

## Prerequisites

- Windows 10/11
- PowerShell 5.1+ (run as Administrator)
- [ollama](https://ollama.com) (installed automatically by setup)

## Quick Start

```powershell
# 1. Run setup (detects hardware, installs ollama, pulls the best model)
.\Setup-Whyboot.ps1

# 2. After an unexpected reboot, diagnose it
.\Diagnose-LastReboot.ps1
```

## Scripts

### `Setup-Whyboot.ps1`

Detects your CPU, RAM, and GPU (including VRAM), selects the largest ollama model that fits your hardware, installs ollama if needed, pulls the model, and writes `config.json`.

```powershell
.\Setup-Whyboot.ps1          # interactive
.\Setup-Whyboot.ps1 -Force   # skip confirmation prompts
```

**Model selection** is automatic based on available memory:

| Model | Memory Needed | Notes |
|-------|--------------|-------|
| `qwen3:235b` | 150 GB | Flagship MoE, extreme hardware |
| `qwen3:30b` | 20 GB | 30B MoE (3B active), rivals 32B dense |
| `qwen3:14b` | 10 GB | Great quality |
| `qwen3:8b` | 5 GB | Good quality |
| `qwen3:4b` | 3 GB | Solid quality |
| `qwen3:1.7b` | 2 GB | Runs on most systems |
| `qwen3:0.6b` | 1 GB | Runs on anything |

### `Diagnose-LastReboot.ps1`

The main diagnostic script. Finds the last reboot, gathers event log entries around that time, sends the data to ollama for AI analysis, searches the web for similar issues, and writes a timestamped report file.

```powershell
.\Diagnose-LastReboot.ps1
.\Diagnose-LastReboot.ps1 -OllamaModel "mistral" -WindowSeconds 30
```

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-OllamaModel` | from `config.json` | Override the ollama model |
| `-OllamaUrl` | `http://localhost:11434` | Override the ollama API URL |
| `-WindowSeconds` | `10` | Seconds before/after boot to search for events |

Output is saved to the current directory as both `reboot-YYYYMMDD-HHmmss-diagnosis.txt` and `reboot-YYYYMMDD-HHmmss-diagnosis.md`.

### `Analyze-RebootPatterns.ps1`

Looks across **multiple** unexpected reboots to find services and service combinations that consistently correlate with crashes. Uses Service Control Manager Event 7036 to reconstruct which services were running at the time of each crash, then compares those crash sessions against clean shutdown sessions as a baseline.

```powershell
.\Analyze-RebootPatterns.ps1
.\Analyze-RebootPatterns.ps1 -Days 90 -PreCrashMinutes 60
.\Analyze-RebootPatterns.ps1 -NoOllama -MinCrashesForPattern 1
```

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-Days` | `60` | Days to look back for unexpected reboots |
| `-PreCrashMinutes` | `30` | Minutes before each crash to flag as pre-crash activity |
| `-MinCrashesForPattern` | `2` | Min crashes a service must appear in to be reported |
| `-MaxCleanSessions` | `10` | Max clean sessions to use as baseline |
| `-NoOllama` | off | Skip the AI pattern analysis |
| `-OllamaModel` / `-OllamaUrl` | from `config.json` | Override ollama settings |

**What it detects:**

- **Always-running services** — present in every crash session (100%). Flagged as unusual if rarely seen in clean sessions.
- **Usually-running services** — present in 75%+ of crashes. Shown with clean-session comparison.
- **Crash-specific services** — high crash rate but low clean-session rate; strongest signal of a problematic service.
- **Always-stopped services** — consistently absent from running state at crash time.
- **Pre-crash service changes** — service state transitions seen in multiple crash sessions within the last N minutes before each crash.
- **Always co-running pairs** — service combinations that appear together in every crash.
- **Crash-only pairs** — service combinations that appear in crashes but never in clean sessions.

Output is saved as `reboot-patterns-YYYYMMDD-HHmmss.txt` and `.md`.

> **Note:** Only services that had at least one state-change event (SCM Event 7036) during a session are tracked. Core kernel components that never emit 7036 events will not appear.

### `Get-KernelPowerReport.ps1`

Queries the last N `Microsoft-Windows-Kernel-Power` events and writes a self-contained HTML report. Includes summary statistics, an event ID breakdown table, and histograms by hour of day, day of week, and calendar week, followed by a full scrollable event listing.

```powershell
.\Get-KernelPowerReport.ps1
.\Get-KernelPowerReport.ps1 -Count 250 -OutputPath C:\Temp\power.html
.\Get-KernelPowerReport.ps1 -NoBrowser
```

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-Count` | `100` | Maximum number of events to retrieve |
| `-OutputPath` | `kernel-power-report.html` | Path for the HTML report |
| `-NoBrowser` | off | Skip opening the report in the default browser after writing |

### `AI-Helper.ps1`

Shared dot-source library used by the other scripts. Provides `Invoke-AIAnalysis`, which tries AI providers in priority order until one succeeds:

1. **Claude Code CLI** (`claude -p`) — uses existing Claude Code auth, no key needed
2. **Claude API** — requires `ClaudeApiKey` in `config.json` or `ANTHROPIC_API_KEY` env var
3. **Copilot API** — requires `CopilotToken` in `config.json` or `GITHUB_TOKEN` env var
4. **Ollama** — local fallback, requires `ollama serve` to be running

Not intended to be run directly; dot-sourced by the diagnostic scripts.

## Example Diagnosis

Below is a real diagnosis from February 17, 2026, where an unexpected reboot was traced to a driver issue.

**System:** Windows 11 Pro, 24 GB VRAM (qwen3:30b model)

**What happened:** The system rebooted without cleanly shutting down (dirty shutdown). Two events were captured in the 10-second window around boot:

| Event | Time | Severity | Meaning |
|-------|------|----------|---------|
| Kernel-Power 41 (BugcheckCode=0) | 12:20:14 | Critical | System crashed without a proper shutdown |
| Intel I225-V Network Disconnect | 12:20:15 | Warning | Network dropped as a *consequence* of the crash |

**AI Diagnosis Summary:**
- **Root cause:** A critical system crash (bugcheck) triggered by a faulty driver or hardware issue -- not a power loss or planned action.
- **BugcheckCode=0** indicates a generic system crash, typically caused by driver failure, memory corruption, or hardware instability.
- **The network disconnect was a symptom**, not the cause. The system rebooted before the network stack could recover.
- **An older Event 1074** (planned OS upgrade from 6 days prior) was ruled out as irrelevant.

**Recommended actions from the diagnosis:**
1. Update the Intel I225-V Ethernet driver (most likely culprit)
2. Run `sfc /scannow` to check for corrupted system files
3. If crashes persist, analyze minidumps with BlueScreenView and run Windows Memory Diagnostic (`mdsched.exe`)

> See the full diagnosis report (with web search corroboration) in [`example-diagnosis.md`](example-diagnosis.md).

## How It Works

1. **Boot detection** -- queries WMI (`Win32_OperatingSystem`) for the last boot time
2. **Event collection** -- pulls Critical/Error/Warning events from System and Application logs within a configurable window around boot, plus key reboot indicators (Event 41, 6008, 1074, WHEA errors)
3. **LLM analysis** -- sends the event data to a local ollama instance for expert interpretation
4. **Web search** -- queries DuckDuckGo for community knowledge about the specific error pattern
5. **Synthesis** -- feeds the web results back to the LLM for a combined diagnosis
6. **Report** -- writes everything to a timestamped text file
