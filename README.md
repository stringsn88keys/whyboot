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

### `Get-RebootReason.ps1`

A broader reboot troubleshooter that scans the last N days of event logs for patterns across multiple categories: unexpected shutdowns, kernel-power events, BSODs, hardware errors, disk health, Windows Update reboots, and more.

```powershell
.\Get-RebootReason.ps1
.\Get-RebootReason.ps1 -Days 30 -AnalyzePreCrash 15
.\Get-RebootReason.ps1 -InstallMonitoring
```

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-Days` | `7` | Days to look back in event logs |
| `-AnalyzePreCrash` | `10` | Minutes before each crash to analyze |
| `-InstallMonitoring` | off | Offer to install monitoring tools (BlueScreenView, HWiNFO64, CrystalDiskInfo, etc.) |

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
