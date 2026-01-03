# =============================================================================
# Set Display Resolution - Force 1920x1080 resolution at boot
# =============================================================================
# This script runs at user login to ensure the display is set to 1920x1080
# Required for consistent Chrome window sizing and capture quality

param(
    [switch]$FromTask  # When called from scheduled task, only set resolution
)

# Self-elevate if not running as Administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin -and -not $FromTask) {
    Write-Host "Elevating to Administrator..." -ForegroundColor Yellow
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -FromTask" -Verb RunAs -Wait
    exit
}

Write-Host "Setting display resolution to 1920x1080..." -ForegroundColor Cyan

try {
    # Add Windows.Graphics.Display type for resolution changes
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class DisplaySettings {
    [DllImport("user32.dll")]
    public static extern int ChangeDisplaySettings(ref DEVMODE devMode, int flags);
    
    [StructLayout(LayoutKind.Sequential)]
    public struct DEVMODE {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmDeviceName;
        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;
        public int dmPositionX;
        public int dmPositionY;
        public int dmDisplayOrientation;
        public int dmDisplayFixedOutput;
        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmFormName;
        public short dmLogPixels;
        public int dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;
        public int dmICMMethod;
        public int dmICMIntent;
        public int dmMediaType;
        public int dmDitherType;
        public int dmReserved1;
        public int dmReserved2;
        public int dmPanningWidth;
        public int dmPanningHeight;
    }
}
"@ -ErrorAction SilentlyContinue

    # Create DEVMODE structure for 1920x1080
    $devMode = New-Object DisplaySettings+DEVMODE
    $devMode.dmSize = [System.Runtime.InteropServices.Marshal]::SizeOf($devMode)
    $devMode.dmPelsWidth = 1920
    $devMode.dmPelsHeight = 1080
    $devMode.dmBitsPerPel = 32
    $devMode.dmFields = 0x1C0000  # DM_PELSWIDTH | DM_PELSHEIGHT | DM_BITSPERPEL

    # Apply resolution change
    $result = [DisplaySettings]::ChangeDisplaySettings([ref]$devMode, 0)
    
    if ($result -eq 0) {
        Write-Host "[OK] Display resolution set to 1920x1080" -ForegroundColor Green
    } else {
        Write-Host "[WARNING] ChangeDisplaySettings returned $result" -ForegroundColor Yellow
        
        # Fallback: Try using QRes if available
        if (Test-Path "C:\Windows\System32\QRes.exe") {
            & QRes.exe /x:1920 /y:1080
            Write-Host "[OK] Resolution set using QRes" -ForegroundColor Green
        }
    }
}
catch {
    Write-Host "[ERROR] Failed to set resolution: $_" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

# Only create scheduled task if NOT called from task (avoid infinite loop)
if (-not $FromTask) {
    # Create scheduled task to run at user login for persistence
    Write-Host ""
    Write-Host "Creating scheduled task for auto-resolution..." -ForegroundColor Yellow

    $taskName = "MSBMC-SetResolution"
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    if ($existingTask) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    }

    # Create VBScript wrapper to hide PowerShell window completely
    $vbsWrapperPath = "C:\ProgramData\msbmc-resolution-launcher.vbs"
    $vbsContent = @"
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""C:\MSBMC\Scripts\set-resolution.ps1"" -FromTask", 0, False
"@
    $vbsContent | Set-Content $vbsWrapperPath -Force

    $action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$vbsWrapperPath`""
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User "msbmc"
    $principal = New-ScheduledTaskPrincipal -UserId "msbmc" -LogonType Interactive -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

    try {
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force -ErrorAction Stop | Out-Null
        Write-Host "[OK] Scheduled task created - resolution will persist after reboot" -ForegroundColor Green
    } catch {
        Write-Host "[ERROR] Failed to create scheduled task: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "[INFO] Resolution is set for current session but may not persist" -ForegroundColor Yellow
    }
} else {
    Write-Host "[OK] Resolution set (running from scheduled task)" -ForegroundColor Green
}
