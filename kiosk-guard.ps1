# =============================================================================
# MSBMC Kiosk Guard - Blocks Windows key and maintains fullscreen Chrome
# =============================================================================
# This script runs continuously and:
# 1. Blocks ALL Windows key presses (including virtual from noVNC)
# 2. Keeps taskbar hidden and disabled
# 3. Ensures Chrome covers the ENTIRE screen (including taskbar area)
# =============================================================================

Add-Type @"
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

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
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    // Work area modification
    [DllImport("user32.dll")]
    public static extern bool SystemParametersInfo(int uiAction, int uiParam, ref RECT pvParam, int fWinIni);
    
    [DllImport("user32.dll")]
    public static extern bool SystemParametersInfo(int uiAction, int uiParam, IntPtr pvParam, int fWinIni);
    
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
    
    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_SYSKEYDOWN = 0x0104;
    private const int VK_LWIN = 0x5B;
    private const int VK_RWIN = 0x5C;
    
    public const int SW_HIDE = 0;
    public const int GWL_STYLE = -16;
    public const int WS_CAPTION = 0x00C00000;
    public const int WS_THICKFRAME = 0x00040000;
    
    public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    public const uint SWP_SHOWWINDOW = 0x0040;
    public const uint SWP_FRAMECHANGED = 0x0020;
    
    public const int SPI_SETWORKAREA = 0x002F;
    public const int SPIF_SENDCHANGE = 0x02;
    
    private static IntPtr _hookID = IntPtr.Zero;
    private static LowLevelKeyboardProc _proc = HookCallback;
    
    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam) {
        if (nCode >= 0) {
            int vkCode = Marshal.ReadInt32(lParam);
            
            // Block Windows keys
            if (vkCode == VK_LWIN || vkCode == VK_RWIN) {
                return (IntPtr)1; // Block the key
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
    
    public static void SetFullWorkArea(int width, int height) {
        RECT rect = new RECT();
        rect.Left = 0;
        rect.Top = 0;
        rect.Right = width;
        rect.Bottom = height;
        SystemParametersInfo(SPI_SETWORKAREA, 0, ref rect, SPIF_SENDCHANGE);
    }
    
    public static void MakeChromeFullscreen(IntPtr hwnd, int width, int height) {
        if (hwnd == IntPtr.Zero) return;
        
        // Remove window borders
        int style = GetWindowLong(hwnd, GWL_STYLE);
        style = style & ~WS_CAPTION & ~WS_THICKFRAME;
        SetWindowLong(hwnd, GWL_STYLE, style);
        
        // Position at 0,0 with exact dimensions
        SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, width, height, SWP_SHOWWINDOW | SWP_FRAMECHANGED);
        SetForegroundWindow(hwnd);
    }
}
"@

# Check if Chrome-only mode is enabled
function Is-ChromeOnlyMode {
    return (Test-Path "C:\ProgramData\msbmc-chrome-only.flag")
}

# Screen dimensions
$screenWidth = 1920
$screenHeight = 1080

# Start the keyboard hook to block Windows key
[KioskGuard]::StartHook()
Write-Host "[KIOSK-GUARD] Keyboard hook installed - Windows key blocked" -ForegroundColor Green

# Set work area to full screen (removes taskbar reservation)
[KioskGuard]::SetFullWorkArea($screenWidth, $screenHeight)
Write-Host "[KIOSK-GUARD] Work area set to full screen: ${screenWidth}x${screenHeight}" -ForegroundColor Green

# Main loop
while ($true) {
    if (Is-ChromeOnlyMode) {
        # Hide taskbar
        [KioskGuard]::HideTaskbar()
        
        # Find Chrome window and make it fullscreen
        $chromeProcs = Get-Process chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
        if ($chromeProcs) {
            $hwnd = $chromeProcs[0].MainWindowHandle
            [KioskGuard]::MakeChromeFullscreen($hwnd, $screenWidth, $screenHeight)
        }
    }
    
    Start-Sleep -Seconds 2
}
