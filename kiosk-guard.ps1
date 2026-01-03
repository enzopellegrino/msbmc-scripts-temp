# =============================================================================
# MSBMC Kiosk Guard - Complete kiosk management (SINGLE SCRIPT)
# =============================================================================
# This script handles EVERYTHING:
# 1. Blocks Windows key (low-level keyboard hook) - only in chrome-only mode
# 2. Hides/shows taskbar based on mode
# 3. Sets work area to full screen / restores it
# 4. Keeps Chrome borderless, fullscreen, topmost (chrome-only mode)
# 5. Restarts Chrome if closed (chrome-only mode)
# 6. Restores normal behavior in maintenance mode
# =============================================================================

Add-Type -AssemblyName System.Windows.Forms

Add-Type @"
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

public class KioskGuard {
    // Keyboard hook
    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
    
    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
    
    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);
    
    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
    
    [DllImport("kernel32.dll")]
    private static extern IntPtr GetModuleHandle(string lpModuleName);
    
    // Window management
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
    public static extern bool IsIconic(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    
    // Work area
    [DllImport("user32.dll")]
    public static extern bool SystemParametersInfo(int uiAction, int uiParam, ref RECT pvParam, int fWinIni);
    
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left, Top, Right, Bottom;
    }
    
    // Constants
    private const int WH_KEYBOARD_LL = 13;
    private const int VK_LWIN = 0x5B;
    private const int VK_RWIN = 0x5C;
    
    public const int SW_HIDE = 0;
    public const int SW_SHOW = 5;
    public const int SW_RESTORE = 9;
    public const int GWL_STYLE = -16;
    public const int WS_CAPTION = 0x00C00000;
    public const int WS_THICKFRAME = 0x00040000;
    
    public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
    public const uint SWP_SHOWWINDOW = 0x0040;
    public const uint SWP_FRAMECHANGED = 0x0020;
    
    public const int SPI_SETWORKAREA = 0x002F;
    public const int SPIF_SENDCHANGE = 0x02;
    
    // Hook state - can be toggled
    private static IntPtr _hookID = IntPtr.Zero;
    private static LowLevelKeyboardProc _proc = HookCallback;
    private static bool _blockWinKey = true;
    
    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam) {
        if (nCode >= 0 && _blockWinKey) {
            int vkCode = Marshal.ReadInt32(lParam);
            if (vkCode == VK_LWIN || vkCode == VK_RWIN) {
                return (IntPtr)1; // Block
            }
        }
        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }
    
    public static void StartHook() {
        if (_hookID == IntPtr.Zero) {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule) {
                _hookID = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }
    }
    
    public static void SetBlockWinKey(bool block) {
        _blockWinKey = block;
    }
    
    public static void HideTaskbar() {
        IntPtr hwnd = FindWindow("Shell_TrayWnd", null);
        if (hwnd != IntPtr.Zero) {
            ShowWindow(hwnd, SW_HIDE);
            EnableWindow(hwnd, false);
        }
        IntPtr hwnd2 = FindWindow("Shell_SecondaryTrayWnd", null);
        if (hwnd2 != IntPtr.Zero) {
            ShowWindow(hwnd2, SW_HIDE);
            EnableWindow(hwnd2, false);
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
    }
    
    public static void SetFullWorkArea(int w, int h) {
        RECT r = new RECT(); r.Left = 0; r.Top = 0; r.Right = w; r.Bottom = h;
        SystemParametersInfo(SPI_SETWORKAREA, 0, ref r, SPIF_SENDCHANGE);
    }
    
    public static void RestoreWorkArea(int w, int h, int taskbarH) {
        RECT r = new RECT(); r.Left = 0; r.Top = 0; r.Right = w; r.Bottom = h - taskbarH;
        SystemParametersInfo(SPI_SETWORKAREA, 0, ref r, SPIF_SENDCHANGE);
    }
    
    public static void MakeChromeFullscreen(IntPtr hwnd, int w, int h) {
        if (hwnd == IntPtr.Zero) return;
        
        // Restore if minimized
        if (IsIconic(hwnd)) {
            ShowWindow(hwnd, SW_RESTORE);
        }
        
        // Remove borders
        int style = GetWindowLong(hwnd, GWL_STYLE);
        SetWindowLong(hwnd, GWL_STYLE, style & ~WS_CAPTION & ~WS_THICKFRAME);
        
        // Position fullscreen and topmost
        SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, w, h, SWP_SHOWWINDOW | SWP_FRAMECHANGED);
        SetForegroundWindow(hwnd);
    }
    
    public static bool IsMinimized(IntPtr hwnd) {
        return hwnd != IntPtr.Zero && IsIconic(hwnd);
    }
    
    public static void MakeChromeNormal(IntPtr hwnd) {
        if (hwnd == IntPtr.Zero) return;
        SetWindowPos(hwnd, HWND_NOTOPMOST, 50, 50, 1800, 950, SWP_SHOWWINDOW | SWP_FRAMECHANGED);
    }
}
"@

# Configuration
$screenWidth = 1920
$screenHeight = 1080
$taskbarHeight = 40
$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$chromeProfileDir = "D:\ChromeProfile"
$flagFile = "C:\ProgramData\msbmc-chrome-only.flag"
$logFile = "C:\MSBMC\Logs\kiosk-guard.log"

# Ensure log directory exists
$logDir = Split-Path $logFile -Parent
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$timestamp] $Message" | Add-Content $logFile
    Write-Host "[$timestamp] $Message"
}

# State
$lastMode = $null

function Is-ChromeOnlyMode {
    return (Test-Path $flagFile)
}

function Start-Chrome {
    Write-Log "Start-Chrome called..."
    $existing = Get-Process chrome -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Log "Chrome already running (PID: $($existing[0].Id))"
        return
    }
    
    Write-Log "Launching Chrome from: $chromePath"
    Write-Log "Profile dir: $chromeProfileDir"
    
    if (-not (Test-Path $chromePath)) {
        Write-Log "ERROR: Chrome not found at $chromePath"
        return
    }
    
    try {
        $proc = Start-Process $chromePath -ArgumentList "--user-data-dir=`"$chromeProfileDir`"" -PassThru
        Write-Log "Chrome started with PID: $($proc.Id)"
        Start-Sleep -Seconds 3
    } catch {
        Write-Log "ERROR starting Chrome: $_"
    }
}

function Get-ChromeWindow {
    $procs = Get-Process chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
    if ($procs) { return $procs[0].MainWindowHandle }
    return [IntPtr]::Zero
}

function Wait-ForDesktop {
    Write-Log "Waiting for desktop to be ready..."
    $maxWait = 60
    $waited = 0
    while ($waited -lt $maxWait) {
        $explorer = Get-Process explorer -ErrorAction SilentlyContinue
        $taskbar = [KioskGuard]::FindWindow("Shell_TrayWnd", $null)
        if ($explorer -and $taskbar -ne [IntPtr]::Zero) {
            Write-Log "Desktop ready (explorer running, taskbar found)"
            Start-Sleep -Seconds 3  # Extra wait for stability
            return $true
        }
        Start-Sleep -Seconds 1
        $waited++
    }
    Write-Log "Timeout waiting for desktop"
    return $false
}

# ============================================================================
# STARTUP
# ============================================================================
Write-Log "KioskGuard starting..."

# Small delay for system stability at boot
Start-Sleep -Seconds 5

# AUTO-CREATE flag file at boot (default to chrome-only mode)
if (-not (Test-Path $flagFile)) {
    Write-Log "Flag file not found - creating (default chrome-only mode)"
    Set-Content -Path $flagFile -Value "enabled" -Force
}

# Start keyboard hook
[KioskGuard]::StartHook()
Write-Log "Keyboard hook installed"

# Start Chrome FIRST (before applying restrictions)
if (-not (Get-Process chrome -ErrorAction SilentlyContinue)) {
    Write-Log "Starting Chrome..."
    Start-Chrome
    Start-Sleep -Seconds 5
}

# Now apply lockdown
Write-Log "Applying chrome-only lockdown..."
[KioskGuard]::SetBlockWinKey($true)
[KioskGuard]::HideTaskbar()
Start-Sleep -Seconds 1
[KioskGuard]::SetFullWorkArea($screenWidth, $screenHeight)

# Make Chrome fullscreen
$hwnd = Get-ChromeWindow
if ($hwnd -ne [IntPtr]::Zero) {
    Write-Log "Making Chrome fullscreen..."
    [KioskGuard]::MakeChromeFullscreen($hwnd, $screenWidth, $screenHeight)
} else {
    Write-Log "WARNING: Chrome window not found"
}

$lastMode = $true  # Start in chrome-only mode
Write-Log "Initial setup complete - entering main loop"

# ============================================================================
# MAIN LOOP
# ============================================================================
while ($true) {
    $chromeOnly = Is-ChromeOnlyMode
    
    # Detect mode change
    if ($chromeOnly -ne $lastMode) {
        if ($chromeOnly) {
            Write-Log ">>> Chrome-only mode ENABLED <<<"
            [KioskGuard]::SetBlockWinKey($true)
            [KioskGuard]::SetFullWorkArea($screenWidth, $screenHeight)
            [KioskGuard]::HideTaskbar()
        } else {
            Write-Log ">>> Maintenance mode ENABLED <<<"
            [KioskGuard]::SetBlockWinKey($false)
            [KioskGuard]::ShowTaskbar()
            [KioskGuard]::RestoreWorkArea($screenWidth, $screenHeight, $taskbarHeight)
            
            # Make Chrome normal window
            $hwnd = Get-ChromeWindow
            if ($hwnd -ne [IntPtr]::Zero) {
                [KioskGuard]::MakeChromeNormal($hwnd)
            }
        }
        $lastMode = $chromeOnly
    }
    
    # Chrome-only mode: enforce lockdown continuously
    if ($chromeOnly) {
        # Keep taskbar hidden and work area full
        [KioskGuard]::HideTaskbar()
        [KioskGuard]::SetFullWorkArea($screenWidth, $screenHeight)
        
        # Check Chrome
        $hwnd = Get-ChromeWindow
        if ($hwnd -ne [IntPtr]::Zero) {
            # Check if minimized and log
            if ([KioskGuard]::IsMinimized($hwnd)) {
                Write-Log "Chrome was minimized - restoring..."
            }
            # Chrome running - keep it fullscreen and topmost
            [KioskGuard]::MakeChromeFullscreen($hwnd, $screenWidth, $screenHeight)
        } else {
            # Chrome not running - restart it
            Write-Log "Chrome closed - restarting..."
            Start-Chrome
            Start-Sleep -Seconds 2
            $hwnd = Get-ChromeWindow
            if ($hwnd -ne [IntPtr]::Zero) {
                [KioskGuard]::MakeChromeFullscreen($hwnd, $screenWidth, $screenHeight)
            }
        }
    }
    
    Start-Sleep -Seconds 2
}

