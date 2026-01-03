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

# SIMPLIFIED toggle - KioskGuard does all the heavy lifting
$toggleScriptContent = @'

param(
    [switch]$Enable,
    [switch]$Disable
)

$ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$ChromeProfileDir = "D:\ChromeProfile"
$flagFile = "C:\ProgramData\msbmc-chrome-only.flag"

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class ShellNotify {
    [DllImport("shell32.dll")]
    public static extern void SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
}
"@

if ($Enable) {
    Write-Host "[CHROME-ONLY] Enabling Chrome-only mode..." -ForegroundColor Green
    
    # Verify profile exists
    if (-not (Test-Path $ChromeProfileDir)) {
        Write-Host "[ERROR] Chrome profile not found at $ChromeProfileDir" -ForegroundColor Red
        return
    }
    
    # Create flag file - KioskGuard monitors this
    Set-Content -Path $flagFile -Value "enabled"
    Write-Host "   [OK] Flag file created" -ForegroundColor Green
    
    # Hide desktop icons
    $desktopRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    if (-not (Test-Path $desktopRegPath)) {
        New-Item -Path $desktopRegPath -Force | Out-Null
    }
    Set-ItemProperty -Path $desktopRegPath -Name "HideIcons" -Value 1 -Force
    [ShellNotify]::SHChangeNotify(0x8000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    Write-Host "   [OK] Desktop icons hidden" -ForegroundColor Green
    
    # Start Chrome if not running (KioskGuard will position it)
    $existing = Get-Process chrome -ErrorAction SilentlyContinue
    if (-not $existing) {
        Start-Process $ChromePath -ArgumentList "--user-data-dir=`"$ChromeProfileDir`""
        Write-Host "   [OK] Chrome started" -ForegroundColor Green
    }
    
    Write-Host "[CHROME-ONLY] Mode enabled - KioskGuard will enforce lockdown" -ForegroundColor Cyan
    Write-Host "[CHROME-ONLY] Press ESC 3x + password to exit" -ForegroundColor Yellow
    
} elseif ($Disable) {
    Write-Host "[CHROME-ONLY] Disabling Chrome-only mode..." -ForegroundColor Yellow
    
    # Remove flag file - KioskGuard will detect and restore normal mode
    Remove-Item -Path $flagFile -Force -ErrorAction SilentlyContinue
    Write-Host "   [OK] Flag file removed" -ForegroundColor Green
    
    # Show desktop icons
    $desktopRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Set-ItemProperty -Path $desktopRegPath -Name "HideIcons" -Value 0 -Force
    [ShellNotify]::SHChangeNotify(0x8000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    Write-Host "   [OK] Desktop icons shown" -ForegroundColor Green
    
    Write-Host "[CHROME-ONLY] Maintenance mode - KioskGuard will restore taskbar" -ForegroundColor Cyan
    
} else {
    Write-Host "Usage: msbmc-chrome-only-toggle.ps1 -Enable | -Disable"
}
'@

$toggleScriptContent | Set-Content -Path $toggleScriptPath -Encoding UTF8 -Force
Write-Host "   Created: $toggleScriptPath" -ForegroundColor Green
# =============================================================================
# STEP 2: Create ESC Monitor (always running, handles 3x ESC)
# =============================================================================
Write-Host "[2/4] Creating ESC monitor script..." -ForegroundColor Yellow

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
# STEP 4: Create Kiosk Guard Script (blocks Windows key + fullscreen Chrome)
# =============================================================================
Write-Host "[4/4] Creating Kiosk Guard script and scheduled task..." -ForegroundColor Yellow

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
Write-Host "   Created: $kioskGuardTaskName (runs at logon, monitors flag file)" -ForegroundColor Green

# =============================================================================
# Summary
# =============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "   SETUP COMPLETE" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "How it works:" -ForegroundColor Cyan
Write-Host "   - KioskGuard runs at logon and monitors the flag file" -ForegroundColor White
Write-Host "   - When flag exists: blocks Win key, hides taskbar, Chrome fullscreen" -ForegroundColor White
Write-Host "   - When flag removed (via ESC 3x): restores normal mode" -ForegroundColor White
Write-Host ""
Write-Host "   Normal mode -> ESC 3x -> Chrome-only mode" -ForegroundColor White
Write-Host "   Chrome-only -> ESC 3x -> Password -> Maintenance mode" -ForegroundColor White
Write-Host ""
Write-Host "Password: $KioskPassword" -ForegroundColor Yellow
Write-Host ""
Write-Host "Manual:" -ForegroundColor Cyan
Write-Host "   Enable:  & '$toggleScriptPath' -Enable" -ForegroundColor Gray
Write-Host "   Disable: & '$toggleScriptPath' -Disable" -ForegroundColor Gray
Write-Host ""

# Start KioskGuard now
Write-Host "Starting KioskGuard..." -ForegroundColor Yellow
Start-Process wscript.exe -ArgumentList "`"$kioskGuardVbsPath`"" -WindowStyle Hidden
Write-Host "[OK] KioskGuard running!" -ForegroundColor Green

# Start the ESC monitor now
Write-Host "Starting ESC monitor..." -ForegroundColor Yellow
Start-Process wscript.exe -ArgumentList "`"$vbsPath`"" -WindowStyle Hidden
Write-Host "[OK] ESC monitor running!" -ForegroundColor Green
Write-Host ""

# Enable Chrome-only mode automatically
Write-Host "Enabling Chrome-only mode as default..." -ForegroundColor Yellow
& $toggleScriptPath -Enable
Write-Host ""
Write-Host "[OK] Chrome-only mode enabled!" -ForegroundColor Green
Write-Host ""
Write-Host "After reboot, Chrome will be the only visible application." -ForegroundColor Cyan
Write-Host "Press ESC 3x + password 'msbmc2024' to restore maintenance mode." -ForegroundColor Cyan
Write-Host ""
