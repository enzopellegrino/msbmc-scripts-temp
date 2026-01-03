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
    if (-not ([System.Management.Automation.PSTypeName]'Shell').Type) {
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
        Add-Type -TypeDefinition $code
    }
    
    # Refresh shell to apply icon hide
    [Shell]::SHChangeNotify(0x8000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    
    # PERSISTENT taskbar auto-hide via registry
    $taskbarRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3"
    if (Test-Path $taskbarRegPath) {
        $settings = Get-ItemProperty -Path $taskbarRegPath -Name "Settings" -ErrorAction SilentlyContinue
        if ($settings) {
            $bytes = $settings.Settings
            # Byte 8: 0x03 = auto-hide enabled (0x02 causes issues on some systems)
            $bytes[8] = 0x03
            Set-ItemProperty -Path $taskbarRegPath -Name "Settings" -Value $bytes -Force
        }
    }
    
    # Also mark taskbar as small icons (takes less space if it appears)
    Set-ItemProperty -Path $desktopRegPath -Name "TaskbarSmallIcons" -Value 1 -Force -ErrorAction SilentlyContinue
    
    Start-Sleep -Seconds 1
    
    # Hide taskbar immediately via API (temporary until reboot, then registry takes over)
    [Taskbar]::Hide()
    
    # Disable Windows key to prevent Start menu access
    $explorerPolicyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    if (-not (Test-Path $explorerPolicyPath)) {
        New-Item -Path $explorerPolicyPath -Force | Out-Null
    }
    Set-ItemProperty -Path $explorerPolicyPath -Name "NoWinKeys" -Value 1 -Force
    
    Write-Host "   [OK] Windows key disabled + taskbar set to auto-hide" -ForegroundColor Green
    
    # Start Chrome WITHOUT maximize (we'll position it manually to cover full screen)
    Write-Host "[INFO] Starting Chrome with profile: $ChromeProfileDir" -ForegroundColor Cyan
    
    # Check if Chrome is already running - don't start multiple instances
    $existingChrome = Get-Process chrome -ErrorAction SilentlyContinue
    if (-not $existingChrome) {
        # Start Chrome in windowed mode with slightly larger dimensions to cover borders
        # Position at -8,-8 to hide window borders/title bar edges
        $argString = "--user-data-dir=`"$ChromeProfileDir`" --window-size=1936,1096 --window-position=-8,-8"
        Start-Process $ChromePath -ArgumentList $argString
        Start-Sleep -Seconds 5
        
        # Remove window borders using SetWindowLong to make it borderless
        if (-not ([System.Management.Automation.PSTypeName]'WindowStyle').Type) {
            Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WindowStyle {
    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    [DllImport("user32.dll")]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    public const int GWL_STYLE = -16;
    public const int WS_CAPTION = 0x00C00000;
    public const int WS_THICKFRAME = 0x00040000;
}
"@
        }
        
        # Get Chrome window and make it borderless + resize to exact full screen
        Start-Sleep -Seconds 2
        $chromeWindows = Get-Process chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
        if ($chromeWindows) {
            $mainWindow = $chromeWindows[0].MainWindowHandle
            
            # Remove window caption and frame (borders)
            $currentStyle = [WindowStyle]::GetWindowLong($mainWindow, [WindowStyle]::GWL_STYLE)
            $newStyle = $currentStyle -band (-bnot [WindowStyle]::WS_CAPTION) -band (-bnot [WindowStyle]::WS_THICKFRAME)
            [WindowStyle]::SetWindowLong($mainWindow, [WindowStyle]::GWL_STYLE, $newStyle) | Out-Null
            
            Start-Sleep -Milliseconds 500
            
            # Now move to exact position covering entire screen
            [WindowStyle]::MoveWindow($mainWindow, 0, 0, 1920, 1080, $true) | Out-Null
            [WindowStyle]::SetForegroundWindow($mainWindow) | Out-Null
        }
        
        Write-Host "   [OK] Chrome borderless and positioned to cover full screen" -ForegroundColor Green
    } else {
        Write-Host "   [SKIP] Chrome already running" -ForegroundColor Yellow
    }
    
    Write-Host "[CHROME-ONLY] Done! Chrome maximized covering full screen, desktop locked." -ForegroundColor Green
    Write-Host "[CHROME-ONLY] Windows key disabled. Press ESC 3x + password to exit." -ForegroundColor Yellow
    
} elseif ($Disable) {
    Write-Host "[CHROME-ONLY] Disabling Chrome-only mode..." -ForegroundColor Yellow
    
    # Re-enable Windows key
    $explorerPolicyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    if (Test-Path $explorerPolicyPath) {
        Remove-ItemProperty -Path $explorerPolicyPath -Name "NoWinKeys" -Force -ErrorAction SilentlyContinue
    }
    
    # Disable taskbar auto-hide
    $taskbarRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3"
    if (Test-Path $taskbarRegPath) {
        $settings = Get-ItemProperty -Path $taskbarRegPath -Name "Settings" -ErrorAction SilentlyContinue
        if ($settings) {
            $bytes = $settings.Settings
            # Byte 8: 0x02 = always visible
            $bytes[8] = 0x02
            Set-ItemProperty -Path $taskbarRegPath -Name "Settings" -Value $bytes -Force
        }
    }
    
    Write-Host "   [OK] Windows key re-enabled + taskbar restored" -ForegroundColor Green
    
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

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

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

function Show-PasswordDialog {
    param([string]$Message, [string]$Title)
    
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" Height="200" Width="400" 
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="20"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="20"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Text="$Message" TextWrapping="Wrap" FontSize="14"/>
        <PasswordBox Grid.Row="2" Name="PasswordBox" FontSize="14" Height="30"/>
        
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="OkButton" Content="OK" Width="80" Height="30" Margin="0,0,10,0" IsDefault="True"/>
            <Button Name="CancelButton" Content="Cancel" Width="80" Height="30" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    $passwordBox = $window.FindName("PasswordBox")
    $okButton = $window.FindName("OkButton")
    $cancelButton = $window.FindName("CancelButton")
    
    $okButton.Add_Click({
        $window.Tag = $passwordBox.Password
        $window.DialogResult = $true
        $window.Close()
    })
    
    $cancelButton.Add_Click({
        $window.DialogResult = $false
        $window.Close()
    })
    
    $passwordBox.Focus()
    $result = $window.ShowDialog()
    
    if ($result) {
        return $window.Tag
    }
    return $null
}

function Show-InfoDialog {
    param([string]$Message, [string]$Title)
    
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="$Title" Height="150" Width="400" 
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="20"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Text="$Message" TextWrapping="Wrap" FontSize="14" VerticalAlignment="Center"/>
        <Button Grid.Row="2" Content="OK" Width="80" Height="30" HorizontalAlignment="Center" IsDefault="True"/>
    </Grid>
</Window>
"@
    
    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $window.ShowDialog() | Out-Null
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
                # In Chrome-only mode: ask password to exit to maintenance mode
                $inputPassword = Show-PasswordDialog -Message "Chrome-only mode is active. Enter password to switch to maintenance mode (taskbar and desktop will be visible):" -Title "MSBMC - Exit Chrome-Only Mode"
                
                if ($inputPassword -eq $KioskPassword) {
                    & "C:\ProgramData\msbmc-chrome-only-toggle.ps1" -Disable
                    Start-Sleep -Seconds 2
                    Show-InfoDialog -Message "Maintenance mode enabled. Taskbar and desktop icons are now visible. Press ESC 3x again to return to Chrome-only mode." -Title "MSBMC - Maintenance Mode"
                } elseif ($inputPassword) {
                    Show-InfoDialog -Message "Incorrect password. Chrome-only mode remains active." -Title "MSBMC - Access Denied"
                }
            } else {
                # In normal mode: enable Chrome-only mode
                & "C:\ProgramData\msbmc-chrome-only-toggle.ps1" -Enable
                Start-Sleep -Seconds 2
                Show-InfoDialog -Message "Chrome-only mode enabled. Only Chrome is visible. Press ESC 3x + password to return to maintenance mode." -Title "MSBMC - Chrome-Only Mode"
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
Write-Host "[3/4] Creating ESC monitor task..." -ForegroundColor Yellow

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
# STEP 4: Create Startup Task to Restore Chrome-Only Mode
# =============================================================================
Write-Host "[4/4] Creating Chrome-only mode restore task..." -ForegroundColor Yellow

$restoreScriptPath = "C:\ProgramData\msbmc-restore-chrome-only.ps1"
$restoreScriptContent = @'
# This script runs at logon to ALWAYS restore Chrome-only mode
# Even if operator exited before reboot, we return to Chrome-only mode

$logFile = "C:\MSBMC\Logs\restore-chrome-only.log"
$logDir = Split-Path $logFile -Parent
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Restore task started" | Add-Content $logFile

Start-Sleep -Seconds 10  # Wait for desktop, watchdog, and other services to stabilize

"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Calling toggle -Enable" | Add-Content $logFile

# ALWAYS enable Chrome-only mode at boot
# Operator can temporarily exit with ESC 3x + password, but reboot resets to Chrome-only
try {
    & "C:\ProgramData\msbmc-chrome-only-toggle.ps1" -Enable 2>&1 | Add-Content $logFile
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Toggle completed successfully" | Add-Content $logFile
} catch {
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Toggle failed: $_" | Add-Content $logFile
}
'@
$restoreScriptContent | Set-Content -Path $restoreScriptPath -Encoding UTF8 -Force

$restoreTaskName = "MSBMC-RestoreChromeOnly"
$existingRestoreTask = Get-ScheduledTask -TaskName $restoreTaskName -ErrorAction SilentlyContinue
if ($existingRestoreTask) {
    Unregister-ScheduledTask -TaskName $restoreTaskName -Confirm:$false
}

$restoreVbsPath = "C:\ProgramData\msbmc-restore-chrome-only-launcher.vbs"
$restoreVbsContent = @'
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""C:\ProgramData\msbmc-restore-chrome-only.ps1""", 0, False
'@
$restoreVbsContent | Set-Content -Path $restoreVbsPath -Encoding ASCII -Force

$restoreAction = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$restoreVbsPath`""
$restoreTrigger = New-ScheduledTaskTrigger -AtLogOn -User "msbmc"
$restorePrincipal = New-ScheduledTaskPrincipal -UserId "msbmc" -LogonType Interactive -RunLevel Highest
$restoreSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -TaskName $restoreTaskName -Action $restoreAction -Trigger $restoreTrigger -Principal $restorePrincipal -Settings $restoreSettings -Force | Out-Null
Write-Host "   Created: $restoreTaskName (restores Chrome-only mode at boot)" -ForegroundColor Green

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
