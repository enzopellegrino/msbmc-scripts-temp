# =============================================================================
# Reset Kiosk Mode State
# =============================================================================
# Ripristina lo stato normale del sistema rimuovendo tutte le modifiche kiosk

param()

$ErrorActionPreference = "Continue"

# Elevate if needed
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Restarting as Administrator..." -ForegroundColor Yellow
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   Reset Kiosk Mode State" -ForegroundColor Cyan  
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Stop and disable watchdog
Write-Host "[1/6] Stopping Chrome Watchdog..." -ForegroundColor Yellow
try {
    Stop-ScheduledTask -TaskName "MSBMC-ChromeWatchdog" -ErrorAction SilentlyContinue
    Disable-ScheduledTask -TaskName "MSBMC-ChromeWatchdog" -ErrorAction SilentlyContinue | Out-Null
    Write-Host "   [OK] Watchdog disabled" -ForegroundColor Green
} catch {
    Write-Host "   [SKIP] Watchdog task not found" -ForegroundColor Gray
}

# Remove flag file
Write-Host "[2/6] Removing flag file..." -ForegroundColor Yellow
Remove-Item -Path "C:\ProgramData\msbmc-chrome-only.flag" -Force -ErrorAction SilentlyContinue
Write-Host "   [OK] Flag removed" -ForegroundColor Green

# Show desktop icons
Write-Host "[3/6] Showing desktop icons..." -ForegroundColor Yellow
$desktopRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
if (Test-Path $desktopRegPath) {
    Set-ItemProperty -Path $desktopRegPath -Name "HideIcons" -Value 0 -Force -ErrorAction SilentlyContinue
    Write-Host "   [OK] Desktop icons enabled" -ForegroundColor Green
} else {
    Write-Host "   [SKIP] Registry path not found" -ForegroundColor Gray
}

# Reset taskbar auto-hide via registry (if exists)
Write-Host "[4/6] Resetting taskbar..." -ForegroundColor Yellow
$taskbarRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3"
if (Test-Path $taskbarRegPath) {
    try {
        $settings = Get-ItemProperty -Path $taskbarRegPath -Name "Settings" -ErrorAction Stop
        $bytes = $settings.Settings
        if ($bytes[8] -ne 0x00) {
            $bytes[8] = 0x00  # Disable auto-hide
            Set-ItemProperty -Path $taskbarRegPath -Name "Settings" -Value $bytes -Force
            Write-Host "   [OK] Taskbar auto-hide disabled" -ForegroundColor Green
        } else {
            Write-Host "   [OK] Taskbar already visible" -ForegroundColor Green
        }
    } catch {
        Write-Host "   [SKIP] Could not modify taskbar registry" -ForegroundColor Gray
    }
} else {
    Write-Host "   [SKIP] Taskbar registry path not found" -ForegroundColor Gray
}

# Show taskbar via API
Write-Host "[5/6] Ensuring taskbar is visible..." -ForegroundColor Yellow
try {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class TaskbarReset {
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string className, string windowName);
    [DllImport("user32.dll")]
    public static extern int ShowWindow(IntPtr hwnd, int command);
    public static void Show() {
        IntPtr hwnd = FindWindow("Shell_TrayWnd", null);
        if (hwnd != IntPtr.Zero) {
            ShowWindow(hwnd, 1);  // SW_SHOWNORMAL
        }
    }
}
"@
    [TaskbarReset]::Show()
    Write-Host "   [OK] Taskbar shown via API" -ForegroundColor Green
} catch {
    Write-Host "   [SKIP] Type already exists or API failed" -ForegroundColor Gray
}

# Refresh explorer to apply all changes
Write-Host "[6/6] Refreshing explorer..." -ForegroundColor Yellow
try {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class ShellNotify {
    [DllImport("shell32.dll")]
    public static extern void SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
}
"@
    [ShellNotify]::SHChangeNotify(0x8000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    Write-Host "   [OK] Shell refreshed" -ForegroundColor Green
} catch {
    Write-Host "   [SKIP] Type already exists" -ForegroundColor Gray
}

# Kill Chrome if running
Write-Host "" -ForegroundColor Yellow
Write-Host "Killing Chrome processes..." -ForegroundColor Yellow
Stop-Process -Name "chrome" -Force -ErrorAction SilentlyContinue
Write-Host "   [OK] Chrome stopped" -ForegroundColor Green

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "   RESET COMPLETE" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "System restored to normal state." -ForegroundColor Cyan
Write-Host "Desktop icons, taskbar, and Chrome should now be normal." -ForegroundColor Cyan
Write-Host ""
Write-Host "You can now test the kiosk mode with:" -ForegroundColor Yellow
Write-Host "   C:\MSBMC\Scripts\configure-kiosk.ps1" -ForegroundColor White
Write-Host ""
