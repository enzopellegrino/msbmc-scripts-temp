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
Write-Host "[1/5] Creating chrome-only-toggle.ps1..." -ForegroundColor Yellow

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
    
    # Create flag file
    Set-Content -Path "C:\ProgramData\msbmc-chrome-only.flag" -Value "enabled"
    
    # Enable and start Kiosk Guard (blocks Windows key + maintains fullscreen)
    Enable-ScheduledTask -TaskName "MSBMC-KioskGuard" -ErrorAction SilentlyContinue | Out-Null
    Start-ScheduledTask -TaskName "MSBMC-KioskGuard" -ErrorAction SilentlyContinue
    
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
    
    # Load Windows API types
    if (-not ([System.Management.Automation.PSTypeName]'KioskHelper').Type) {
        $code = @"
using System;
using System.Runtime.InteropServices;

public class KioskHelper {
    [DllImport("shell32.dll")]
    public static extern void SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
    
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string className, string windowName);
    
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
    
    [DllImport("user32.dll")]
    public static extern int ShowWindow(IntPtr hwnd, int command);
    
    [DllImport("user32.dll")]
    public static extern bool EnableWindow(IntPtr hwnd, bool enable);
    
    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    
    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    
    [DllImport("user32.dll")]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    
    public const int SW_HIDE = 0;
    public const int SW_SHOW = 5;
    public const int GWL_STYLE = -16;
    public const int GWL_EXSTYLE = -20;
    public const int WS_CAPTION = 0x00C00000;
    public const int WS_THICKFRAME = 0x00040000;
    public const int WS_EX_APPWINDOW = 0x00040000;
    public const int WS_EX_TOOLWINDOW = 0x00000080;
    
    public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
    public const uint SWP_NOSIZE = 0x0001;
    public const uint SWP_NOMOVE = 0x0002;
    public const uint SWP_SHOWWINDOW = 0x0040;
    public const uint SWP_FRAMECHANGED = 0x0020;
    
    public static void HideTaskbar() {
        // Hide main taskbar
        IntPtr hwnd = FindWindow("Shell_TrayWnd", null);
        if (hwnd != IntPtr.Zero) {
            ShowWindow(hwnd, SW_HIDE);
            EnableWindow(hwnd, false);  // Disable clicks
        }
        // Hide secondary taskbar (multi-monitor)
        IntPtr hwnd2 = FindWindow("Shell_SecondaryTrayWnd", null);
        if (hwnd2 != IntPtr.Zero) {
            ShowWindow(hwnd2, SW_HIDE);
            EnableWindow(hwnd2, false);
        }
        // Hide Start button separately
        IntPtr startBtn = FindWindow("Button", "Start");
        if (startBtn != IntPtr.Zero) {
            ShowWindow(startBtn, SW_HIDE);
        }
    }
    
    public static void ShowTaskbar() {
        IntPtr hwnd = FindWindow("Shell_TrayWnd", null);
        if (hwnd != IntPtr.Zero) {
            EnableWindow(hwnd, true);
            ShowWindow(hwnd, SW_SHOW);
        }
        IntPtr hwnd2 = FindWindow("Shell_SecondaryTrayWnd", null);
        if (hwnd2 != IntPtr.Zero) {
            EnableWindow(hwnd2, true);
            ShowWindow(hwnd2, SW_SHOW);
        }
        IntPtr startBtn = FindWindow("Button", "Start");
        if (startBtn != IntPtr.Zero) {
            ShowWindow(startBtn, SW_SHOW);
        }
    }
}
"@
        Add-Type -TypeDefinition $code
    }
    
    # Refresh shell to apply icon hide
    [KioskHelper]::SHChangeNotify(0x8000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    
    # HIDE AND DISABLE taskbar completely
    Start-Sleep -Seconds 1
    [KioskHelper]::HideTaskbar()
    
    # Disable Windows key via Scancode Map (requires logoff but we also use policy)
    $explorerPolicyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    if (-not (Test-Path $explorerPolicyPath)) {
        New-Item -Path $explorerPolicyPath -Force | Out-Null
    }
    Set-ItemProperty -Path $explorerPolicyPath -Name "NoWinKeys" -Value 1 -Force
    
    # Also set system-wide policy (more reliable)
    $systemPolicyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    if (-not (Test-Path $systemPolicyPath)) {
        New-Item -Path $systemPolicyPath -Force | Out-Null
    }
    Set-ItemProperty -Path $systemPolicyPath -Name "NoWinKeys" -Value 1 -Force -ErrorAction SilentlyContinue
    
    Write-Host "   [OK] Taskbar hidden + disabled, Windows key blocked" -ForegroundColor Green
    
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
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    
    public const int GWL_STYLE = -16;
    public const int WS_CAPTION = 0x00C00000;
    public const int WS_THICKFRAME = 0x00040000;

    public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    public const uint SWP_NOSIZE = 0x0001;
    public const uint SWP_NOMOVE = 0x0002;
    public const uint SWP_SHOWWINDOW = 0x0040;
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
            
            # Now move to exact position covering entire screen and force always-on-top
            [WindowStyle]::MoveWindow($mainWindow, 0, 0, 1920, 1080, $true) | Out-Null
            [WindowStyle]::SetForegroundWindow($mainWindow) | Out-Null
            [WindowStyle]::SetWindowPos(
                $mainWindow,
                [WindowStyle]::HWND_TOPMOST,
                0,
                0,
                0,
                0,
                [WindowStyle]::SWP_NOSIZE -bor [WindowStyle]::SWP_NOMOVE -bor [WindowStyle]::SWP_SHOWWINDOW
            ) | Out-Null
        }
        
        Write-Host "   [OK] Chrome borderless and positioned to cover full screen" -ForegroundColor Green
    } else {
        Write-Host "   [SKIP] Chrome already running" -ForegroundColor Yellow
    }
    
    Write-Host "[CHROME-ONLY] Done! Chrome maximized covering full screen, desktop locked." -ForegroundColor Green
    Write-Host "[CHROME-ONLY] Windows key disabled. Press ESC 3x + password to exit." -ForegroundColor Yellow
    
} elseif ($Disable) {
    Write-Host "[CHROME-ONLY] Disabling Chrome-only mode..." -ForegroundColor Yellow
    
    # Re-enable Windows key (both user and system)
    $explorerPolicyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    if (Test-Path $explorerPolicyPath) {
        Remove-ItemProperty -Path $explorerPolicyPath -Name "NoWinKeys" -Force -ErrorAction SilentlyContinue
    }
    $systemPolicyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    if (Test-Path $systemPolicyPath) {
        Remove-ItemProperty -Path $systemPolicyPath -Name "NoWinKeys" -Force -ErrorAction SilentlyContinue
    }
    
    Write-Host "   [OK] Windows key re-enabled" -ForegroundColor Green
    
    # Stop and disable Kiosk Guard (stops keyboard hook)
    Stop-ScheduledTask -TaskName "MSBMC-KioskGuard" -ErrorAction SilentlyContinue
    Disable-ScheduledTask -TaskName "MSBMC-KioskGuard" -ErrorAction SilentlyContinue | Out-Null
    # Kill any running kiosk guard process
    Get-Process powershell -ErrorAction SilentlyContinue | Where-Object { 
        $_.CommandLine -like "*kiosk-guard*" 
    } | Stop-Process -Force -ErrorAction SilentlyContinue
    
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
    
    # Show taskbar using KioskHelper (already loaded)
    if ([System.Management.Automation.PSTypeName]'KioskHelper'.Type) {
        [KioskHelper]::ShowTaskbar()
        [KioskHelper]::SHChangeNotify(0x8000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    } else {
        # Fallback if KioskHelper not loaded
        $code = @"
using System;
using System.Runtime.InteropServices;
public class TaskbarShow {
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string className, string windowName);
    [DllImport("user32.dll")]
    public static extern int ShowWindow(IntPtr hwnd, int command);
    [DllImport("user32.dll")]
    public static extern bool EnableWindow(IntPtr hwnd, bool enable);
    public static void Show() {
        IntPtr hwnd = FindWindow("Shell_TrayWnd", null);
        if (hwnd != IntPtr.Zero) {
            EnableWindow(hwnd, true);
            ShowWindow(hwnd, 5);  // SW_SHOW
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
    }
    
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
Write-Host "[2/5] Creating ESC monitor script..." -ForegroundColor Yellow

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
Write-Host "[3/5] Creating ESC monitor task..." -ForegroundColor Yellow

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
# STEP 4: Create Kiosk Guard Script (blocks Windows key + fullscreen Chrome)
# =============================================================================
Write-Host "[4/5] Creating Kiosk Guard script..." -ForegroundColor Yellow

$kioskGuardPath = "C:\ProgramData\msbmc-kiosk-guard.ps1"
$kioskGuardContent = @'
# MSBMC Kiosk Guard - Blocks Windows key and maintains fullscreen Chrome

Add-Type -AssemblyName System.Windows.Forms

Add-Type @"
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

public class KioskGuard {
    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
    
    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
    
    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);
    
    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
    
    [DllImport("kernel32.dll")]
    private static extern IntPtr GetModuleHandle(string lpModuleName);
    
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string className, string windowName);
    
    [DllImport("user32.dll")]
    public static extern int ShowWindow(IntPtr hwnd, int command);
    
    [DllImport("user32.dll")]
    public static extern bool EnableWindow(IntPtr hwnd, bool enable);
    
    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    
    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool SystemParametersInfo(int uiAction, int uiParam, ref RECT pvParam, int fWinIni);
    
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left, Top, Right, Bottom;
    }
    
    private const int WH_KEYBOARD_LL = 13;
    private const int VK_LWIN = 0x5B;
    private const int VK_RWIN = 0x5C;
    
    public const int SW_HIDE = 0;
    public const int SW_SHOW = 5;
    public const int GWL_STYLE = -16;
    public const int WS_CAPTION = 0x00C00000;
    public const int WS_THICKFRAME = 0x00040000;
    public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
    public const uint SWP_SHOWWINDOW = 0x0040;
    public const uint SWP_FRAMECHANGED = 0x0020;
    public const int SPI_SETWORKAREA = 0x002F;
    public const int SPIF_SENDCHANGE = 0x02;
    
    private static IntPtr _hookID = IntPtr.Zero;
    private static LowLevelKeyboardProc _proc = HookCallback;
    
    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam) {
        if (nCode >= 0) {
            int vkCode = Marshal.ReadInt32(lParam);
            if (vkCode == VK_LWIN || vkCode == VK_RWIN) {
                return (IntPtr)1;
            }
        }
        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }
    
    public static void StartHook() {
        using (Process curProcess = Process.GetCurrentProcess())
        using (ProcessModule curModule = curProcess.MainModule) {
            _hookID = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, GetModuleHandle(curModule.ModuleName), 0);
        }
    }
    
    public static void StopHook() {
        if (_hookID != IntPtr.Zero) {
            UnhookWindowsHookEx(_hookID);
            _hookID = IntPtr.Zero;
        }
    }
    
    public static void HideTaskbar() {
        IntPtr hwnd = FindWindow("Shell_TrayWnd", null);
        if (hwnd != IntPtr.Zero) { ShowWindow(hwnd, SW_HIDE); EnableWindow(hwnd, false); }
        IntPtr hwnd2 = FindWindow("Shell_SecondaryTrayWnd", null);
        if (hwnd2 != IntPtr.Zero) { ShowWindow(hwnd2, SW_HIDE); EnableWindow(hwnd2, false); }
    }
    
    public static void ShowTaskbar() {
        IntPtr hwnd = FindWindow("Shell_TrayWnd", null);
        if (hwnd != IntPtr.Zero) { EnableWindow(hwnd, true); ShowWindow(hwnd, SW_SHOW); }
    }
    
    public static void SetFullWorkArea(int w, int h) {
        RECT r = new RECT(); r.Left = 0; r.Top = 0; r.Right = w; r.Bottom = h;
        SystemParametersInfo(SPI_SETWORKAREA, 0, ref r, SPIF_SENDCHANGE);
    }
    
    public static void MakeChromeFullscreen(IntPtr hwnd, int w, int h) {
        if (hwnd == IntPtr.Zero) return;
        int style = GetWindowLong(hwnd, GWL_STYLE);
        SetWindowLong(hwnd, GWL_STYLE, style & ~WS_CAPTION & ~WS_THICKFRAME);
        SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, w, h, SWP_SHOWWINDOW | SWP_FRAMECHANGED);
        SetForegroundWindow(hwnd);
    }
    
    public static void MakeChromeNormal(IntPtr hwnd) {
        if (hwnd == IntPtr.Zero) return;
        SetWindowPos(hwnd, HWND_NOTOPMOST, 100, 100, 1600, 900, SWP_SHOWWINDOW | SWP_FRAMECHANGED);
    }
}
"@

function Is-ChromeOnlyMode { return (Test-Path "C:\ProgramData\msbmc-chrome-only.flag") }

$sw = 1920
$sh = 1080

[KioskGuard]::StartHook()
[KioskGuard]::SetFullWorkArea($sw, $sh)

while ($true) {
    if (Is-ChromeOnlyMode) {
        [KioskGuard]::HideTaskbar()
        $cp = Get-Process chrome -EA SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
        if ($cp) { [KioskGuard]::MakeChromeFullscreen($cp[0].MainWindowHandle, $sw, $sh) }
    }
    Start-Sleep -Seconds 2
}
'@
$kioskGuardContent | Set-Content -Path $kioskGuardPath -Encoding UTF8 -Force
Write-Host "   Created: $kioskGuardPath" -ForegroundColor Green

# Create scheduled task for Kiosk Guard
$kioskGuardTaskName = "MSBMC-KioskGuard"
$existingKioskGuardTask = Get-ScheduledTask -TaskName $kioskGuardTaskName -ErrorAction SilentlyContinue
if ($existingKioskGuardTask) {
    Unregister-ScheduledTask -TaskName $kioskGuardTaskName -Confirm:$false
}

$kioskGuardVbsPath = "C:\ProgramData\msbmc-kiosk-guard-launcher.vbs"
$kioskGuardVbsContent = @'
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""C:\ProgramData\msbmc-kiosk-guard.ps1""", 0, False
'@
$kioskGuardVbsContent | Set-Content -Path $kioskGuardVbsPath -Encoding ASCII -Force

$kioskGuardAction = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$kioskGuardVbsPath`""
$kioskGuardTrigger = New-ScheduledTaskTrigger -AtLogOn -User "msbmc"
$kioskGuardPrincipal = New-ScheduledTaskPrincipal -UserId "msbmc" -LogonType Interactive -RunLevel Highest
$kioskGuardSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask -TaskName $kioskGuardTaskName -Action $kioskGuardAction -Trigger $kioskGuardTrigger -Principal $kioskGuardPrincipal -Settings $kioskGuardSettings -Force | Out-Null
Write-Host "   Created: $kioskGuardTaskName (keyboard hook + fullscreen)" -ForegroundColor Green

# =============================================================================
# STEP 5: Create Startup Task to Restore Chrome-Only Mode
# =============================================================================
Write-Host "[5/5] Creating Chrome-only mode restore task..." -ForegroundColor Yellow

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
