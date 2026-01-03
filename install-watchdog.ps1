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
    [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll")] public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [DllImport("user32.dll")] public static extern IntPtr FindWindow(string className, string windowName);
    
    public const int SW_MAXIMIZE = 3;
    public const int SW_RESTORE = 9;
    public const int SW_HIDE = 0;
    public const int GWL_STYLE = -16;
    public const int WS_CAPTION = 0x00C00000;
    public const int WS_THICKFRAME = 0x00040000;
    
    public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    public const uint SWP_NOSIZE = 0x0001;
    public const uint SWP_NOMOVE = 0x0002;
    public const uint SWP_SHOWWINDOW = 0x0040;
    public const uint SWP_FRAMECHANGED = 0x0020;
    
    // Hide taskbar
    public static void HideTaskbar() {
        IntPtr hwnd = FindWindow("Shell_TrayWnd", null);
        if (hwnd != IntPtr.Zero) {
            ShowWindow(hwnd, SW_HIDE);
        }
    }
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

# Function to ensure Chrome window is borderless, fullscreen, and topmost
function Ensure-ChromeFullscreen {
    $chromeProcs = Get-Process chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
    
    foreach ($proc in $chromeProcs) {
        $hwnd = $proc.MainWindowHandle
        if ($hwnd -ne 0) {
            # Remove window borders (caption + thick frame)
            $currentStyle = [WindowHelper]::GetWindowLong($hwnd, [WindowHelper]::GWL_STYLE)
            $newStyle = $currentStyle -band (-bnot [WindowHelper]::WS_CAPTION) -band (-bnot [WindowHelper]::WS_THICKFRAME)
            [WindowHelper]::SetWindowLong($hwnd, [WindowHelper]::GWL_STYLE, $newStyle) | Out-Null
            
            # Move to exact fullscreen position
            [WindowHelper]::MoveWindow($hwnd, 0, 0, 1920, 1080, $true) | Out-Null
            
            # Set as topmost (always on top)
            [WindowHelper]::SetWindowPos(
                $hwnd,
                [WindowHelper]::HWND_TOPMOST,
                0, 0, 0, 0,
                [WindowHelper]::SWP_NOSIZE -bor [WindowHelper]::SWP_NOMOVE -bor [WindowHelper]::SWP_SHOWWINDOW -bor [WindowHelper]::SWP_FRAMECHANGED
            ) | Out-Null
            
            # Bring to foreground
            [WindowHelper]::SetForegroundWindow($hwnd) | Out-Null
        }
    }
    
    # Also ensure taskbar stays hidden
    [WindowHelper]::HideTaskbar()
}

# Initial cleanup on first run
Clear-ChromeRestoreData | Out-Null

# Main watchdog loop
while ($true) {
    $chromeProc = Get-Process chrome -ErrorAction SilentlyContinue
    
    if ($chromeProc) {
        # Chrome running - ensure it's borderless, fullscreen, and topmost
        Ensure-ChromeFullscreen
    } else {
        # Chrome NOT running
        # Only restart if Chrome-only mode is active
        if (Is-ChromeOnlyMode) {
            Write-Host "[WATCHDOG] Chrome closed in Chrome-only mode - restarting..." -ForegroundColor Yellow
            Start-ChromeMaximized
            Start-Sleep -Seconds 3
            Ensure-ChromeFullscreen
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
