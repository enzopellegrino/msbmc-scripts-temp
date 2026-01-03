# =============================================================================
# MSBMC Chrome-Only Mode Configuration
# =============================================================================
# 
# Funzionamento:
#   - Hotkey 3x ESC: Toggle tra modalità Chrome-only e modalità normale
#   - In Chrome-only: Chrome è la shell, niente explorer/taskbar/desktop
#   - Password richiesta per tornare alla modalità normale
#
# =============================================================================

param(
    [switch]$FromTask
)

$ErrorActionPreference = "Stop"

# If not admin, restart elevated
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    if (-not $FromTask) {
        Write-Host "Restarting as Administrator..." -ForegroundColor Yellow
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    }
    exit
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   MSBMC Chrome-Only Mode Setup" -ForegroundColor Cyan  
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$KioskPassword = "msbmc2024"
$ChromeProfileDir = if (Test-Path "D:\ChromeProfile") { "D:\ChromeProfile" } else { "C:\ChromeProfile" }

# =============================================================================
# STEP 1: Create the Chrome-Only Toggle Script
# =============================================================================
Write-Host "[1/3] Creating chrome-only-toggle.ps1..." -ForegroundColor Yellow

$toggleScriptPath = "C:\ProgramData\msbmc-chrome-only-toggle.ps1"

$toggleScriptContent = @'

param(
    [switch]$Enable,
    [switch]$Disable
)

$ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$ChromeProfileDir = "D:\ChromeProfile"  # Always use D:\ persistent profile

if ($Enable) {
    Write-Host "[CHROME-ONLY] Enabling Chrome-only mode..." -ForegroundColor Green
    
    # Verify profile exists
    if (-not (Test-Path $ChromeProfileDir)) {
        Write-Host "[ERROR] Chrome profile not found at $ChromeProfileDir" -ForegroundColor Red
        Write-Host "Please run setup-ami-interactive.ps1 first to create the profile." -ForegroundColor Yellow
        return
    }
    
    Write-Host "[INFO] Using Chrome profile: $ChromeProfileDir" -ForegroundColor Cyan
    
    # NON cambiamo la shell - explorer rimane attivo per gestire processi di sistema
    # Invece nascondiamo taskbar + desktop icons + forziamo Chrome sempre in primo piano
    
    # Create flag file
    Set-Content -Path "C:\ProgramData\msbmc-chrome-only.flag" -Value "enabled"
    
    # Enable Chrome Watchdog (mantiene Chrome sempre attivo e in primo piano)
    Enable-ScheduledTask -TaskName "MSBMC-ChromeWatchdog" -ErrorAction SilentlyContinue | Out-Null
    
    # Start watchdog BEFORE Chrome (così monitora subito)
    Start-ScheduledTask -TaskName "MSBMC-ChromeWatchdog" -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    
    # Hide desktop icons via registry
    $desktopRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    if (-not (Test-Path $desktopRegPath)) {
        New-Item -Path $desktopRegPath -Force | Out-Null
    }
    Set-ItemProperty -Path $desktopRegPath -Name "HideIcons" -Value 1 -Force
    
    # Apply registry changes without killing explorer
    $code = @"
using System;
using System.Runtime.InteropServices;
public class Shell {
    [DllImport("shell32.dll")]
    public static extern void SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
}
public class Taskbar {
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string className, string windowName);
    [DllImport("user32.dll")]
    public static extern int ShowWindow(IntPtr hwnd, int command);
    public static void Hide() {
        IntPtr hwnd = FindWindow("Shell_TrayWnd", null);
        if (hwnd != IntPtr.Zero) {
            ShowWindow(hwnd, 0);  // SW_HIDE
        }
    }
}
"@
    
    if (-not ([System.Management.Automation.PSTypeName]'Shell').Type) {
        Add-Type -TypeDefinition $code
    }
    
    # Refresh shell to apply icon hide
    [Shell]::SHChangeNotify(0x8000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    Start-Sleep -Seconds 1
    
    # Hide taskbar
    [Taskbar]::Hide()
    
    # Start Chrome MAXIMIZED
    $chromeArgs = @(
        "--start-maximized"
        "--user-data-dir=$ChromeProfileDir"
        "--disable-session-crashed-bubble"
        "--disable-infobars"
        "--no-first-run"
        "--disable-background-mode"
        "https://espn.com"
    )
    Start-Process $ChromePath -ArgumentList $chromeArgs
    
    Write-Host "[CHROME-ONLY] Done! Chrome maximized, desktop locked." -ForegroundColor Green
    Write-Host "[CHROME-ONLY] Press ESC 3x + password 'msbmc2024' to exit." -ForegroundColor Yellow
    
} elseif ($Disable) {
    Write-Host "[CHROME-ONLY] Disabling Chrome-only mode..." -ForegroundColor Yellow
    
    # Disable Chrome Watchdog
    Disable-ScheduledTask -TaskName "MSBMC-ChromeWatchdog" -ErrorAction SilentlyContinue | Out-Null
    Stop-ScheduledTask -TaskName "MSBMC-ChromeWatchdog" -ErrorAction SilentlyContinue
    
    # Remove flag file
    Remove-Item -Path "C:\ProgramData\msbmc-chrome-only.flag" -Force -ErrorAction SilentlyContinue
    
    # Show desktop icons
    $desktopRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    if (-not (Test-Path $desktopRegPath)) {
        New-Item -Path $desktopRegPath -Force | Out-Null
    }
    Set-ItemProperty -Path $desktopRegPath -Name "HideIcons" -Value 0 -Force
    
    # Show taskbar
    $code = @"
using System;
using System.Runtime.InteropServices;
public class TaskbarShow {
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
public class ShellRefresh {
    [DllImport("shell32.dll")]
    public static extern void SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
}
"@
    
    if (-not ([System.Management.Automation.PSTypeName]'TaskbarShow').Type) {
        Add-Type -TypeDefinition $code
    }
    
    [TaskbarShow]::Show()
    
    # Refresh shell to show icons
    [ShellRefresh]::SHChangeNotify(0x8000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    
    # Kill Chrome
    Stop-Process -Name "chrome" -Force -ErrorAction SilentlyContinue
    
    Write-Host "[CHROME-ONLY] Normal mode restored." -ForegroundColor Green
    
} else {
    Write-Host "Usage: msbmc-chrome-only-toggle.ps1 -Enable | -Disable"
}
'@

$toggleScriptContent | Set-Content -Path $toggleScriptPath -Encoding UTF8 -Force
Write-Host "   Created: $toggleScriptPath" -ForegroundColor Green

# =============================================================================
# STEP 2: Create ESC Monitor (always running, handles 3x ESC)
# =============================================================================
Write-Host "[2/3] Creating ESC monitor script..." -ForegroundColor Yellow

$escMonitorPath = "C:\ProgramData\msbmc-esc-monitor.ps1"

$escMonitorContent = @'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class KeyboardHook {
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);
}
"@

$KioskPassword = "msbmc2024"
$VK_ESCAPE = 0x1B
$escCount = 0
$lastEscTime = [DateTime]::MinValue
$escWindow = [TimeSpan]::FromSeconds(2)
$escReleased = $true

function Is-ChromeOnlyMode {
    return (Test-Path "C:\ProgramData\msbmc-chrome-only.flag")
}

while ($true) {
    Start-Sleep -Milliseconds 50
    
    $keyState = [KeyboardHook]::GetAsyncKeyState($VK_ESCAPE)
    $escPressed = ($keyState -band 0x8000) -ne 0
    
    if ($escPressed -and $escReleased) {
        $escReleased = $false
        $now = Get-Date
        
        if (($now - $lastEscTime) -gt $escWindow) {
            $escCount = 1
        } else {
            $escCount++
        }
        $lastEscTime = $now
        
        if ($escCount -ge 3) {
            $escCount = 0
            
            if (Is-ChromeOnlyMode) {
                # In Chrome-only mode: ask password to exit
                $inputPassword = [Microsoft.VisualBasic.Interaction]::InputBox(
                    "Enter password to exit Chrome-only mode:",
                    "MSBMC Maintenance",
                    ""
                )
                
                if ($inputPassword -eq $KioskPassword) {
                    & "C:\ProgramData\msbmc-chrome-only-toggle.ps1" -Disable
                }
            } else {
                # In normal mode: enable Chrome-only mode
                & "C:\ProgramData\msbmc-chrome-only-toggle.ps1" -Enable
            }
        }
    }
    
    if (-not $escPressed) {
        $escReleased = $true
    }
}
'@

$escMonitorContent | Set-Content -Path $escMonitorPath -Encoding UTF8 -Force
Write-Host "   Created: $escMonitorPath" -ForegroundColor Green

# =============================================================================
# STEP 3: Create Scheduled Task to run ESC monitor at logon
# =============================================================================
Write-Host "[3/3] Creating scheduled task..." -ForegroundColor Yellow

$taskName = "MSBMC-EscMonitor"
$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# Create VBScript to run hidden
$vbsPath = "C:\ProgramData\msbmc-esc-monitor-launcher.vbs"
$vbsContent = @'
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""C:\ProgramData\msbmc-esc-monitor.ps1""", 0, False
'@
$vbsContent | Set-Content -Path $vbsPath -Encoding ASCII -Force

$action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$vbsPath`""
$trigger = New-ScheduledTaskTrigger -AtLogOn -User "msbmc"
$principal = New-ScheduledTaskPrincipal -UserId "msbmc" -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
Write-Host "   Created: $taskName (runs at logon)" -ForegroundColor Green

# =============================================================================
# Summary
# =============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "   SETUP COMPLETE" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "How it works:" -ForegroundColor Cyan
Write-Host "   Normal mode → ESC 3x → Chrome-only mode" -ForegroundColor White
Write-Host "   Chrome-only → ESC 3x → Password → Normal mode" -ForegroundColor White
Write-Host ""
Write-Host "Password: $KioskPassword" -ForegroundColor Yellow
Write-Host ""
Write-Host "Manual:" -ForegroundColor Cyan
Write-Host "   Enable:  & '$toggleScriptPath' -Enable" -ForegroundColor Gray
Write-Host "   Disable: & '$toggleScriptPath' -Disable" -ForegroundColor Gray
Write-Host ""

# Start the ESC monitor now
Write-Host "Starting ESC monitor..." -ForegroundColor Yellow
Start-Process wscript.exe -ArgumentList "`"$vbsPath`"" -WindowStyle Hidden
Write-Host "[OK] ESC monitor running!" -ForegroundColor Green
Write-Host ""

# Enable Chrome-only mode automatically
Write-Host "Enabling Chrome-only mode as default..." -ForegroundColor Yellow
Write-Host "  (instances will boot directly to Chrome with espn.com)" -ForegroundColor Cyan
& $toggleScriptPath -Enable
Write-Host "[OK] Chrome-only mode enabled!" -ForegroundColor Green
Write-Host ""
Write-Host "After reboot, Chrome will be the only visible application." -ForegroundColor Cyan
Write-Host "Press ESC 3x + password 'msbmc2024' to restore normal mode." -ForegroundColor Cyan
Write-Host ""
