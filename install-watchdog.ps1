# =============================================================================
# Install Chrome Watchdog as Scheduled Task
# Keeps Chrome running in maximized mode with persistent D: profile
# =============================================================================

$ErrorActionPreference = 'Continue'

# Self-elevate if not running as Administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Elevating to Administrator..." -ForegroundColor Yellow
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -Wait
    exit
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Chrome Watchdog Installation" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# =============================================================================
# Step 1: Create Watchdog Script
# =============================================================================
Write-Host "[1/3] Creating Chrome Watchdog script..." -ForegroundColor Yellow

$watchdogScript = @'
# Chrome Watchdog - Keeps Chrome running MAXIMIZED with D: persistent profile
# Monitors window state and forces maximized if minimized/windowed

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WindowHelper {
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    public const int SW_MAXIMIZE = 3;
    public const int SW_RESTORE = 9;
}
"@

$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$profileDir = if (Test-Path "D:\") { "D:\ChromeProfile" } else { "C:\ChromeProfile" }

if (-not (Test-Path $profileDir)) {
    New-Item -ItemType Directory -Path $profileDir -Force | Out-Null
}

# Function to clean Chrome crash/restore data
function Clear-ChromeRestoreData {
    # Stop Chrome first if running
    $wasRunning = $false
    $chromeProcs = Get-Process chrome -ErrorAction SilentlyContinue
    if ($chromeProcs) {
        $wasRunning = $true
        $chromeProcs | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
    
    # Clear Preferences exit_type
    $prefsFile = Join-Path $profileDir "Default\Preferences"
    if (Test-Path $prefsFile) {
        try {
            $content = Get-Content $prefsFile -Raw
            $content = $content -replace '"exit_type"\s*:\s*"[^"]*"', '"exit_type":"Normal"'
            $content = $content -replace '"exited_cleanly"\s*:\s*false', '"exited_cleanly":true'
            $content | Set-Content $prefsFile -Force
        } catch {}
    }
    
    # Delete Session Restore files completely
    $sessionDir = Join-Path $profileDir "Default"
    $filesToDelete = @(
        "Current Session",
        "Current Tabs",
        "Last Session",
        "Last Tabs",
        "Session Storage",
        "Sessions"
    )
    foreach ($file in $filesToDelete) {
        $path = Join-Path $sessionDir $file
        if (Test-Path $path) {
            Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    
    # Delete Crashpad folder
    $crashpadDir = Join-Path $profileDir "Crashpad"
    if (Test-Path $crashpadDir) {
        Remove-Item $crashpadDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    return $wasRunning
}

# Function to start Chrome maximized
function Start-ChromeMaximized {
    # Check if Chrome is already running - avoid duplicate instances
    $existingChrome = Get-Process chrome -ErrorAction SilentlyContinue
    if ($existingChrome) {
        Write-Host "[WATCHDOG] Chrome already running - skipping start" -ForegroundColor Yellow
        return
    }
    
    Write-Host "[WATCHDOG] Starting Chrome..." -ForegroundColor Cyan
    $argString = "--start-maximized --user-data-dir=`"$profileDir`""
    Start-Process $chromePath -ArgumentList $argString
    Start-Sleep -Seconds 3
}

# Function to check if Chrome-only mode is active
function Is-ChromeOnlyMode {
    return (Test-Path "C:\ProgramData\msbmc-chrome-only.flag")
}

# Function to ensure Chrome window is maximized
function Ensure-ChromeMaximized {
    $chromeProcs = Get-Process chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
    
    foreach ($proc in $chromeProcs) {
        $hwnd = $proc.MainWindowHandle
        if ($hwnd -ne 0) {
            $isMinimized = [WindowHelper]::IsIconic($hwnd)
            $isMaximized = [WindowHelper]::IsZoomed($hwnd)
            
            if ($isMinimized -or (-not $isMaximized)) {
                # Restore if minimized, then maximize
                if ($isMinimized) {
                    [WindowHelper]::ShowWindow($hwnd, [WindowHelper]::SW_RESTORE) | Out-Null
                    Start-Sleep -Milliseconds 200
                }
                [WindowHelper]::ShowWindow($hwnd, [WindowHelper]::SW_MAXIMIZE) | Out-Null
                [WindowHelper]::SetForegroundWindow($hwnd) | Out-Null
            }
        }
    }
}

# Initial cleanup on first run
Clear-ChromeRestoreData | Out-Null

# Main watchdog loop
while ($true) {
    $chromeProc = Get-Process chrome -ErrorAction SilentlyContinue
    
    if ($chromeProc) {
        # Chrome running - ensure it's maximized and in foreground
        Ensure-ChromeMaximized
    } else {
        # Chrome NOT running
        # Only restart if Chrome-only mode is active
        if (Is-ChromeOnlyMode) {
            Write-Host "[WATCHDOG] Chrome closed in Chrome-only mode - restarting..." -ForegroundColor Yellow
            Start-ChromeMaximized
        }
        # In maintenance mode, do nothing - let user work freely
    }
    
    Start-Sleep -Seconds 3
}
'@

$watchdogPath = "C:\ProgramData\chrome-watchdog.ps1"
$watchdogScript | Set-Content $watchdogPath -Force

Write-Host "[OK] Watchdog script created at $watchdogPath" -ForegroundColor Green

# =============================================================================
# Step 2: Remove old NSSM service if exists
# =============================================================================
Write-Host ""
Write-Host "[2/3] Cleaning up old service..." -ForegroundColor Yellow

$nssmPath = "C:\ProgramData\chocolatey\bin\nssm.exe"
$existingService = Get-Service -Name ChromeWatchdog -ErrorAction SilentlyContinue
if ($existingService) {
    Write-Host "  Removing old ChromeWatchdog service..." -ForegroundColor Cyan
    if (Test-Path $nssmPath) {
        & $nssmPath stop ChromeWatchdog 2>$null
        & $nssmPath remove ChromeWatchdog confirm 2>$null
    } else {
        Stop-Service ChromeWatchdog -Force -ErrorAction SilentlyContinue
        sc.exe delete ChromeWatchdog 2>$null
    }
    Start-Sleep -Seconds 2
}

# =============================================================================
# Step 3: Register as Scheduled Task (runs in user session)
# =============================================================================
Write-Host ""
Write-Host "[3/3] Registering ChromeWatchdog scheduled task..." -ForegroundColor Yellow

# Create VBScript wrapper to hide PowerShell window completely
$vbsWrapperPath = "C:\ProgramData\chrome-watchdog-launcher.vbs"
$vbsContent = @"
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File ""$watchdogPath""", 0, False
"@
$vbsContent | Set-Content $vbsWrapperPath -Force

$taskName = "MSBMC-ChromeWatchdog"

$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "  Removing existing task..." -ForegroundColor Cyan
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

$action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument "`"$vbsWrapperPath`""
$trigger = New-ScheduledTaskTrigger -AtLogOn -User 'msbmc'
$principal = New-ScheduledTaskPrincipal -UserId 'msbmc' -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null

Write-Host "[OK] ChromeWatchdog scheduled task registered" -ForegroundColor Green

# Disable by default - will be enabled only in Chrome-only mode
Write-Host "  Disabling task (enabled only in Chrome-only mode)..." -ForegroundColor Cyan
Disable-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue | Out-Null
Write-Host "[OK] Task created (disabled by default)" -ForegroundColor Green

$task = Get-ScheduledTask -TaskName $taskName
Write-Host "[OK] Task status: $($task.State)" -ForegroundColor Green

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Chrome Watchdog Installation Complete" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Task Details:" -ForegroundColor Yellow
Write-Host "  Name: $taskName" -ForegroundColor White
Write-Host "  Script: $watchdogPath" -ForegroundColor White
Write-Host "  Trigger: At logon (msbmc user)" -ForegroundColor White
Write-Host "  Status: DISABLED (enabled by configure-kiosk.ps1 in Chrome-only mode)" -ForegroundColor Cyan
Write-Host ""
Write-Host "Behavior:" -ForegroundColor Yellow
Write-Host "  - Task DISABLED by default" -ForegroundColor White
Write-Host "  - ENABLED automatically when entering Chrome-only mode" -ForegroundColor White
Write-Host "  - Keeps Chrome running maximized" -ForegroundColor White
Write-Host "  - Restarts Chrome if closed" -ForegroundColor White
Write-Host ""
