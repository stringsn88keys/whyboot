#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Sets up whyboot by detecting hardware, selecting the best ollama model,
    installing ollama, pulling the model, and writing config.json.
.DESCRIPTION
    Detects system CPU, RAM, and GPU (including VRAM) to select the largest
    ollama model that will run well on the current hardware. Installs ollama
    via winget if not already present, pulls the selected model, and saves
    the configuration to config.json.
.PARAMETER Force
    Skip confirmation prompts.
.EXAMPLE
    .\Setup-Whyboot.ps1
    .\Setup-Whyboot.ps1 -Force
#>

param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$configPath = Join-Path $scriptDir "config.json"

Write-Host "`n===== WHYBOOT SETUP =====" -ForegroundColor Cyan

# ── 1. Detect hardware ────────────────────────────────────────────────────────

Write-Host "`n[1] Detecting hardware..." -ForegroundColor Yellow

# RAM
$totalRamGB = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB)
Write-Host "  RAM: ${totalRamGB} GB" -ForegroundColor Gray

# CPU
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
$cpuName = $cpu.Name.Trim()
$cpuCores = $cpu.NumberOfCores
$cpuThreads = $cpu.NumberOfLogicalProcessors
Write-Host "  CPU: $cpuName ($cpuCores cores, $cpuThreads threads)" -ForegroundColor Gray

# GPU - query all video controllers
$gpus = Get-CimInstance Win32_VideoController
$bestGpu = $null
$bestVramGB = 0

foreach ($gpu in $gpus) {
    $vramBytes = $gpu.AdapterRAM
    # AdapterRAM is a uint32 so caps at 4GB; for larger GPUs check registry
    $vramGB = 0

    if ($vramBytes -and $vramBytes -gt 0) {
        $vramGB = [math]::Round($vramBytes / 1GB, 1)
    }

    # Try registry for accurate VRAM on modern GPUs (AdapterRAM overflows at 4GB)
    $regVram = $null
    try {
        $regPath = "HKLM:\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}"
        $subKeys = Get-ChildItem $regPath -ErrorAction SilentlyContinue
        foreach ($key in $subKeys) {
            $props = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
            if ($props.DriverDesc -and $props.DriverDesc -eq $gpu.Name.Trim()) {
                # qwMemorySize is a 64-bit value that correctly reports VRAM >4GB
                $qwMemSize = $props.'HardwareInformation.qwMemorySize'
                if ($qwMemSize) {
                    $regVram = [math]::Round([int64]$qwMemSize / 1GB, 1)
                }

                # Fallback to MemorySize (32-bit, caps at 4GB but better than nothing)
                if (-not $regVram) {
                    $memSize = $props.'HardwareInformation.MemorySize'
                    if ($memSize) {
                        $regVram = [math]::Round([int64]$memSize / 1GB, 1)
                    }
                }
                break
            }
        }
    } catch {}

    if ($regVram -and $regVram -gt $vramGB) {
        $vramGB = $regVram
    }

    $gpuName = $gpu.Name.Trim()
    Write-Host "  GPU: $gpuName (VRAM: ${vramGB} GB)" -ForegroundColor Gray

    if ($vramGB -gt $bestVramGB) {
        $bestVramGB = $vramGB
        $bestGpu = $gpu
    }
}

if (-not $bestGpu) {
    Write-Host "  GPU: No dedicated GPU detected" -ForegroundColor Gray
}

# ── 2. Select the best model ─────────────────────────────────────────────────

Write-Host "`n[2] Selecting best model for this hardware..." -ForegroundColor Yellow

# Determine if we have a capable GPU (NVIDIA/AMD discrete)
$hasNvidiaGpu = $false
$hasAmdGpu = $false
if ($bestGpu) {
    $hasNvidiaGpu = $bestGpu.Name -match 'NVIDIA|GeForce|RTX|GTX|Quadro|Tesla'
    $hasAmdGpu = $bestGpu.Name -match 'AMD|Radeon RX'
}
$hasDiscreteGpu = $hasNvidiaGpu -or $hasAmdGpu

# Effective memory for model selection:
# - With GPU: VRAM is the bottleneck (model must fit in VRAM)
# - CPU-only: system RAM is the limit (but need headroom for OS)
if ($hasDiscreteGpu -and $bestVramGB -ge 2) {
    $effectiveMemGB = $bestVramGB
    $accelerator = "GPU ($($bestGpu.Name.Trim()))"
} else {
    $effectiveMemGB = [math]::Max(0, $totalRamGB - 4)  # reserve 4GB for OS
    $accelerator = "CPU only"
}

Write-Host "  Accelerator: $accelerator" -ForegroundColor Gray
Write-Host "  Effective memory for models: ${effectiveMemGB} GB" -ForegroundColor Gray

# Model selection table (model name -> approximate memory needed in GB)
# Sorted from largest to smallest; pick the biggest that fits
# Qwen3 MoE models (30B-A3B) are preferred where they fit — they activate
# only a fraction of their parameters so they punch well above their size.
$modelTiers = @(
    @{ Model = "qwen3:235b";  NeedGB = 150; Desc = "235B MoE (22B active) - flagship, needs extreme hardware" }
    @{ Model = "qwen3:30b";   NeedGB = 20;  Desc = "30B MoE (3B active) - rivals 32B dense, very efficient" }
    @{ Model = "qwen3:14b";   NeedGB = 10;  Desc = "14B dense - great quality, needs 12GB+ VRAM or 16GB+ RAM" }
    @{ Model = "qwen3:8b";    NeedGB = 5;   Desc = "8B dense - good quality, needs 8GB+ VRAM or 12GB+ RAM" }
    @{ Model = "qwen3:4b";    NeedGB = 3;   Desc = "4B dense - solid quality, rivals Qwen2.5-72B on benchmarks" }
    @{ Model = "qwen3:1.7b";  NeedGB = 2;   Desc = "1.7B dense - decent quality, runs on most systems" }
    @{ Model = "qwen3:0.6b";  NeedGB = 1;   Desc = "0.6B dense - basic quality, runs on anything" }
)

$selectedModel = $null
foreach ($tier in $modelTiers) {
    if ($effectiveMemGB -ge $tier.NeedGB) {
        $selectedModel = $tier
        break
    }
}

if (-not $selectedModel) {
    $selectedModel = $modelTiers[-1]  # fallback to smallest
}

Write-Host "`n  Selected model: $($selectedModel.Model)" -ForegroundColor Green
Write-Host "  $($selectedModel.Desc)" -ForegroundColor Gray
Write-Host "  Requires ~$($selectedModel.NeedGB) GB, you have ~${effectiveMemGB} GB available" -ForegroundColor Gray

if (-not $Force) {
    $confirm = Read-Host "`nProceed with '$($selectedModel.Model)'? (Y/n)"
    if ($confirm -and $confirm -notmatch '^[Yy]') {
        Write-Host "Setup cancelled." -ForegroundColor Yellow
        exit 0
    }
}

# ── 3. Install ollama ─────────────────────────────────────────────────────────

Write-Host "`n[3] Checking ollama installation..." -ForegroundColor Yellow

$ollamaCmd = Get-Command ollama -ErrorAction SilentlyContinue

if ($ollamaCmd) {
    Write-Host "  ollama is already installed: $($ollamaCmd.Source)" -ForegroundColor Green
} else {
    Write-Host "  ollama not found. Installing..." -ForegroundColor Yellow

    $hasWinget = Get-Command winget -ErrorAction SilentlyContinue
    if ($hasWinget) {
        Write-Host "  Installing via winget..." -ForegroundColor Gray
        winget install --id Ollama.Ollama --accept-source-agreements --accept-package-agreements
        if ($LASTEXITCODE -ne 0) {
            Write-Host "  winget install failed. Trying direct download..." -ForegroundColor Yellow
        } else {
            # Refresh PATH
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
            $ollamaCmd = Get-Command ollama -ErrorAction SilentlyContinue
        }
    }

    if (-not $ollamaCmd) {
        Write-Host "  Downloading ollama installer..." -ForegroundColor Gray
        $installerPath = Join-Path $env:TEMP "OllamaSetup.exe"
        Invoke-WebRequest -Uri "https://ollama.com/download/OllamaSetup.exe" -OutFile $installerPath -UseBasicParsing
        Write-Host "  Running installer..." -ForegroundColor Gray
        Start-Process -FilePath $installerPath -Wait
        Remove-Item $installerPath -ErrorAction SilentlyContinue

        # Refresh PATH
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        $ollamaCmd = Get-Command ollama -ErrorAction SilentlyContinue
    }

    if ($ollamaCmd) {
        Write-Host "  ollama installed successfully." -ForegroundColor Green
    } else {
        Write-Host "  ERROR: Could not find ollama after install. You may need to restart your terminal." -ForegroundColor Red
        Write-Host "  After restarting, run: ollama pull $($selectedModel.Model)" -ForegroundColor Yellow
    }
}

# ── 4. Pull the model ─────────────────────────────────────────────────────────

Write-Host "`n[4] Pulling model '$($selectedModel.Model)'..." -ForegroundColor Yellow

if ($ollamaCmd) {
    # Ensure ollama is serving
    $ollamaRunning = $false
    try {
        Invoke-RestMethod -Uri "http://localhost:11434/api/tags" -TimeoutSec 5 | Out-Null
        $ollamaRunning = $true
    } catch {
        Write-Host "  Starting ollama serve..." -ForegroundColor Gray
        Start-Process "ollama" -ArgumentList "serve" -WindowStyle Hidden
        Start-Sleep -Seconds 3
        try {
            Invoke-RestMethod -Uri "http://localhost:11434/api/tags" -TimeoutSec 5 | Out-Null
            $ollamaRunning = $true
        } catch {
            Write-Host "  WARNING: Could not start ollama serve. Start it manually." -ForegroundColor Yellow
        }
    }

    if ($ollamaRunning) {
        # Check if the model is already downloaded
        $modelAlreadyPulled = $false
        try {
            $tags = Invoke-RestMethod -Uri "http://localhost:11434/api/tags" -TimeoutSec 5
            # Match model name (with or without :latest tag)
            $wantedName = $selectedModel.Model
            foreach ($m in $tags.models) {
                if ($m.name -eq $wantedName -or $m.name -eq "${wantedName}:latest" -or $m.model -eq $wantedName -or $m.model -eq "${wantedName}:latest") {
                    $modelAlreadyPulled = $true
                    break
                }
            }
        } catch {}

        if ($modelAlreadyPulled) {
            Write-Host "  Model '$($selectedModel.Model)' is already downloaded. Skipping pull." -ForegroundColor Green
        } else {
            Write-Host "  This may take a while depending on your connection..." -ForegroundColor Gray
            & ollama pull $selectedModel.Model
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  Model pulled successfully." -ForegroundColor Green
            } else {
                Write-Host "  Model pull failed. Run manually: ollama pull $($selectedModel.Model)" -ForegroundColor Red
            }
        }
    }
} else {
    Write-Host "  Skipping pull (ollama not available). After installing, run:" -ForegroundColor Yellow
    Write-Host "    ollama pull $($selectedModel.Model)" -ForegroundColor Gray
}

# ── 5. Write config.json ──────────────────────────────────────────────────────

Write-Host "`n[5] Writing configuration..." -ForegroundColor Yellow

$config = @{
    OllamaModel = $selectedModel.Model
    OllamaUrl   = "http://localhost:11434"
} | ConvertTo-Json

$config | Out-File -FilePath $configPath -Encoding UTF8
Write-Host "  Config saved to: $configPath" -ForegroundColor Green
Write-Host "  Model: $($selectedModel.Model)" -ForegroundColor Gray

# ── Done ──────────────────────────────────────────────────────────────────────

Write-Host "`n===== SETUP COMPLETE =====" -ForegroundColor Cyan
Write-Host "Run .\Diagnose-LastReboot.ps1 to diagnose your last reboot.`n" -ForegroundColor Gray
