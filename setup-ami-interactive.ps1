# =============================================================================
# MSBMC AMI Configuration - Interactive Setup
# Configures base AMI to production-ready final AMI
# =============================================================================
# 
# USAGE:
#   This script must be run AS MSBMC USER after first login
#   Scripts are already in C:\MSBMC\Scripts\ (copied by Packer)
#   Run this script to configure the AMI:
#     Set-ExecutionPolicy Bypass -Scope Process -Force
#     C:\MSBMC\Scripts\setup-ami-interactive.ps1
#
# =============================================================================

param(
    [switch]$SkipConfirm
)

$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "MSBMC AMI Configuration"

try {

# =============================================================================
# Verify running as msbmc user (NOT Administrator)
# =============================================================================
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
if ($currentUser -notlike "*msbmc") {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "ERROR: Wrong User Context" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "This script MUST be run as 'msbmc' user, not Administrator." -ForegroundColor Yellow
    Write-Host "Current user: $currentUser" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Instructions:" -ForegroundColor Cyan
    Write-Host "  1. Log out from Administrator" -ForegroundColor White
    Write-Host "  2. Log in as 'msbmc' (auto-login should be configured)" -ForegroundColor White
    Write-Host "  3. Run this script again from msbmc user session" -ForegroundColor White
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit 1
}

Write-Host ""
Write-Host "Running as user: $currentUser" -ForegroundColor Green
Write-Host ""

# =============================================================================
# Setup: Scripts directory (already in AMI)
# =============================================================================
$scriptsDir = "C:\MSBMC\Scripts"

# =============================================================================
# Helper Functions - Check BASE AMI Prerequisites (from Packer)
# =============================================================================
function Test-BaseAMIPrerequisites {
    $missingItems = @()
    
    # Check Chrome installed
    $chromePath = "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe"
    if (-not (Test-Path $chromePath)) {
        $missingItems += "Chrome not installed at $chromePath"
    }
    
    # Check ChromeDriver
    $chromeDriverPath = "C:\ProgramData\chocolatey\bin\chromedriver.exe"
    if (-not (Test-Path $chromeDriverPath)) {
        $missingItems += "ChromeDriver not found at $chromeDriverPath"
    }
    
    # Check FFmpeg
    $ffmpegPath = "C:\ProgramData\chocolatey\bin\ffmpeg.exe"
    if (-not (Test-Path $ffmpegPath)) {
        $missingItems += "FFmpeg not installed"
    }
    
    # Check Python
    $pythonPaths = @(
        "C:\Python314\python.exe",
        "C:\Python313\python.exe",
        "C:\Python312\python.exe",
        "C:\Python311\python.exe"
    )
    $pythonPath = $pythonPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $pythonPath) {
        $missingItems += "Python not installed (checked Python 3.11-3.14)"
    }
    
    # Check TightVNC
    $vncPath = "C:\Program Files\TightVNC\tvnserver.exe"
    if (-not (Test-Path $vncPath)) {
        $missingItems += "TightVNC not installed"
    }
    
    # Check noVNC directory
    $novncPath = "C:\noVNC"
    if (-not (Test-Path $novncPath)) {
        $missingItems += "noVNC not installed at $novncPath"
    }
    
    # Check MSBMC scripts directory
    if (-not (Test-Path $scriptsDir)) {
        $missingItems += "MSBMC Scripts directory not found at $scriptsDir"
    }
    
    # Check msbmc user exists
    $msbmcUser = Get-LocalUser -Name "msbmc" -ErrorAction SilentlyContinue
    if (-not $msbmcUser) {
        $missingItems += "User 'msbmc' does not exist"
    }
    
    return @{
        Success = ($missingItems.Count -eq 0)
        MissingItems = $missingItems
    }
}

# =============================================================================
# Helper Functions - Check current configuration state
# =============================================================================
function Test-ResolutionConfigured {
    $task = Get-ScheduledTask -TaskName "MSBMC-SetResolution" -ErrorAction SilentlyContinue
    if (-not $task) { return $false }
    
    # Check if task is configured for msbmc user
    if ($task.Principal.UserId -ne "msbmc") { return $false }
    
    # Check if VBS wrapper exists
    $vbsWrapper = "C:\ProgramData\msbmc-resolution-launcher.vbs"
    return (Test-Path $vbsWrapper)
}

function Test-VBAudioInstalled {
    # Check if VB-Audio device exists
    $audio = Get-WmiObject Win32_SoundDevice -ErrorAction SilentlyContinue | 
        Where-Object { $_.Name -like "*VB-Audio*" -or $_.Name -like "*CABLE*" }
    
    if (-not $audio) { return $false }
    
    # Check if driver files exist
    $driverPath = "C:\Program Files\VB\CABLE\VBCABLE_Setup_x64.exe"
    return (Test-Path $driverPath)
}

function Test-ChromeConfigured {
    # Check if Chrome profile directory exists
    $profileDir = if (Test-Path "D:\ChromeProfile") { "D:\ChromeProfile" } else { "C:\ChromeProfile" }
    if (-not (Test-Path $profileDir)) { return $false }
    
    # Check if Default profile was created by Packer
    $defaultProfile = Join-Path $profileDir "Default"
    if (-not (Test-Path $defaultProfile)) { return $false }
    
    # Check for Packer marker file
    $markerFile = Join-Path $profileDir ".msbmc-initialized"
    if (-not (Test-Path $markerFile)) { return $false }
    
    # Check if Chrome is set as default browser in registry
    $httpHandler = Get-ItemProperty "HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice" -ErrorAction SilentlyContinue
    if (-not $httpHandler -or $httpHandler.ProgId -notlike "*Chrome*") { 
        # Profile exists but default browser not set yet (will be fixed in Step 3)
        return $false
    }
    
    return $true
}

function Test-ChromeWatchdogInstalled {
    $task = Get-ScheduledTask -TaskName "MSBMC-ChromeWatchdog" -ErrorAction SilentlyContinue
    if (-not $task) { return $false }
    
    # Verify task is DISABLED by default (only enabled when Chrome-only mode active)
    if ($task.State -ne 'Disabled' -and $task.State -ne 'Ready') { return $false }
    
    # Check if watchdog script exists
    $watchdogScript = Join-Path $scriptsDir "install-watchdog.ps1"
    return (Test-Path $watchdogScript)
}

function Test-KioskConfigured {
    # Check if ESC monitor task exists
    $escTask = Get-ScheduledTask -TaskName "MSBMC-EscMonitor" -ErrorAction SilentlyContinue
    if (-not $escTask) { return $false }
    
    # Check if toggle script exists in ProgramData (created by configure-kiosk.ps1)
    if (-not (Test-Path "C:\ProgramData\msbmc-chrome-only-toggle.ps1")) { return $false }
    
    # Check if ESC monitor script exists
    if (-not (Test-Path "C:\ProgramData\msbmc-esc-monitor.ps1")) { return $false }
    
    # Verify Chrome-only mode is NOT active (flag file should not exist)
    $flagFile = "C:\ProgramData\msbmc-chrome-only.flag"
    if (Test-Path $flagFile) {
        Write-Host "[WARN] Chrome-only mode is currently ACTIVE (unexpected)" -ForegroundColor Yellow
    }
    
    # Verify Shell registry is NOT set to Chrome (should be default explorer.exe)
    $shellValue = Get-ItemProperty "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" -Name "Shell" -ErrorAction SilentlyContinue
    if ($shellValue -and $shellValue.Shell -like "*chrome*") {
        Write-Host "[WARN] Shell is set to Chrome (Chrome-only mode active)" -ForegroundColor Yellow
        return $false
    }
    
    return $true
}

# =============================================================================
# Detect current configuration state
# =============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "MSBMC AMI Configuration - Interactive" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "User: $currentUser" -ForegroundColor Green
Write-Host ""

# =============================================================================
# Step 0: Verify BASE AMI Prerequisites
# =============================================================================
Write-Host "Verifying BASE AMI prerequisites..." -ForegroundColor Yellow
Write-Host ""

$baseCheck = Test-BaseAMIPrerequisites
if (-not $baseCheck.Success) {
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "ERROR: BASE AMI Incomplete" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "The following prerequisites are missing:" -ForegroundColor Yellow
    foreach ($item in $baseCheck.MissingItems) {
        Write-Host "  [X] $item" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "This means the Packer build did not complete successfully." -ForegroundColor Yellow
    Write-Host "Please rebuild the BASE AMI using Packer before running this script." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit 1
}

Write-Host "[OK] Chrome installed" -ForegroundColor Green
Write-Host "[OK] ChromeDriver installed" -ForegroundColor Green
Write-Host "[OK] FFmpeg installed" -ForegroundColor Green
Write-Host "[OK] Python installed" -ForegroundColor Green
Write-Host "[OK] TightVNC installed" -ForegroundColor Green
Write-Host "[OK] noVNC installed" -ForegroundColor Green
Write-Host "[OK] MSBMC scripts present" -ForegroundColor Green
Write-Host "[OK] User msbmc exists" -ForegroundColor Green
Write-Host ""
Write-Host "BASE AMI prerequisites OK!" -ForegroundColor Green
Write-Host ""

# =============================================================================
# Check current configuration status
# =============================================================================
Write-Host "Checking current configuration..." -ForegroundColor Yellow
Write-Host ""

$resolutionDone = Test-ResolutionConfigured
$vbaudioDone = Test-VBAudioInstalled
$chromeDone = Test-ChromeConfigured
$watchdogDone = Test-ChromeWatchdogInstalled
$kioskDone = Test-KioskConfigured

# Display status
$status1 = if ($resolutionDone) { "[OK]" } else { "[  ]" }
$color1 = if ($resolutionDone) { "Green" } else { "Yellow" }
$status2 = if ($vbaudioDone) { "[OK]" } else { "[  ]" }
$color2 = if ($vbaudioDone) { "Green" } else { "Yellow" }
$status3 = if ($chromeDone) { "[OK]" } else { "[  ]" }
$color3 = if ($chromeDone) { "Green" } else { "Yellow" }
$status4 = if ($watchdogDone) { "[OK]" } else { "[  ]" }
$color4 = if ($watchdogDone) { "Green" } else { "Yellow" }
$status5 = if ($kioskDone) { "[OK]" } else { "[  ]" }
$color5 = if ($kioskDone) { "Green" } else { "Yellow" }

Write-Host "Configuration Status:" -ForegroundColor Cyan
Write-Host "  $status1 1. Display resolution 1920x1080" -ForegroundColor $color1
Write-Host "  $status2 2. VB-Audio Virtual Cable" -ForegroundColor $color2
Write-Host "  $status3 3. Chrome default browser + profile" -ForegroundColor $color3
Write-Host "  $status4 4. Chrome Watchdog service" -ForegroundColor $color4
Write-Host "  $status5 5. Kiosk Mode toggle system" -ForegroundColor $color5
Write-Host ""

# Check if all done
if ($resolutionDone -and $vbaudioDone -and $chromeDone -and $watchdogDone -and $kioskDone) {
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "All Configuration Steps Completed!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  1. Test Chrome-only mode toggle (ESC 3x)" -ForegroundColor White
    Write-Host "  2. Test VNC access via noVNC (port 6080)" -ForegroundColor White
    Write-Host "  3. Verify VB-Audio appears in sound devices" -ForegroundColor White
    Write-Host "  4. Create FINAL AMI from EC2 console" -ForegroundColor White
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit 0
}

# Ask to continue
if (-not $SkipConfirm) {
    Write-Host "Do you want to continue with missing steps? (y/n): " -ForegroundColor Yellow -NoNewline
    $response = Read-Host
    $response = $response.Trim().ToLower()
    if ($response -ne 'y' -and $response -ne 'yes') {
        Write-Host "Configuration cancelled." -ForegroundColor Yellow
        exit 0
    }
}

Write-Host ""
Write-Host "Starting configuration..." -ForegroundColor Green
Write-Host ""

# =============================================================================
# Step 1: Display Resolution
# =============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "[1/5] Display Resolution 1920x1080" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($resolutionDone) {
    Write-Host "[SKIP] Resolution already configured" -ForegroundColor Green
    Write-Host ""
    # Verify current resolution
    Add-Type -AssemblyName System.Windows.Forms
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen
    $currentRes = "$($screen.Bounds.Width)x$($screen.Bounds.Height)"
    Write-Host "Current resolution: $currentRes" -ForegroundColor Cyan
} else {
    $resolutionScript = Join-Path $scriptsDir "set-resolution.ps1"
    if (Test-Path $resolutionScript) {
        Write-Host "Setting resolution to 1920x1080..." -ForegroundColor Yellow
        Write-Host "This will create a scheduled task for persistent resolution." -ForegroundColor Cyan
        Write-Host ""
        
        & $resolutionScript
        
        Write-Host ""
        Write-Host "Verifying resolution..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        
        if (Test-ResolutionConfigured) {
            Write-Host "[OK] Resolution configured successfully" -ForegroundColor Green
            
            # Show current resolution
            Add-Type -AssemblyName System.Windows.Forms
            $screen = [System.Windows.Forms.Screen]::PrimaryScreen
            $currentRes = "$($screen.Bounds.Width)x$($screen.Bounds.Height)"
            Write-Host "Current resolution: $currentRes" -ForegroundColor Cyan
        } else {
            Write-Host "[WARN] Resolution task created but verification failed" -ForegroundColor Yellow
        }
    } else {
        Write-Host "[ERROR] Script set-resolution.ps1 not found" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Press any key to continue..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

# =============================================================================
# Step 2: VB-Audio Virtual Cable
# =============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "[2/5] VB-Audio Virtual Cable" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($vbaudioDone) {
    Write-Host "[SKIP] VB-Audio already installed" -ForegroundColor Green
    Write-Host ""
    # Show VB-Audio device
    $vbDevice = Get-WmiObject Win32_SoundDevice | Where-Object { $_.Name -like "*VB-Audio*" -or $_.Name -like "*CABLE*" }
    if ($vbDevice) {
        Write-Host "Device: $($vbDevice.Name)" -ForegroundColor Cyan
    }
} else {
    $vbaudioScript = Join-Path $scriptsDir "install-vbaudio.ps1"
    if (Test-Path $vbaudioScript) {
        Write-Host "Installing VB-Audio Virtual Cable..." -ForegroundColor Yellow
        Write-Host "This will install audio routing driver for capture." -ForegroundColor Cyan
        Write-Host ""
        
        & $vbaudioScript
        
        Write-Host ""
        Write-Host "Verifying installation..." -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        
        if (Test-VBAudioInstalled) {
            Write-Host "[OK] VB-Audio installed successfully" -ForegroundColor Green
            
            $vbDevice = Get-WmiObject Win32_SoundDevice | Where-Object { $_.Name -like "*VB-Audio*" -or $_.Name -like "*CABLE*" }
            if ($vbDevice) {
                Write-Host "Device: $($vbDevice.Name)" -ForegroundColor Cyan
            }
        } else {
            Write-Host "[WARN] VB-Audio installation may require reboot" -ForegroundColor Yellow
            Write-Host "The driver will be active after next reboot." -ForegroundColor Cyan
        }
    } else {
        Write-Host "[ERROR] Script install-vbaudio.ps1 not found" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Press any key to continue..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

# =============================================================================
# Step 3: Chrome Configuration
# =============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "[3/5] Chrome Configuration" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if profile exists (created by Packer)
$profileDir = if (Test-Path "D:\ChromeProfile") { "D:\ChromeProfile" } else { "C:\ChromeProfile" }
$markerFile = Join-Path $profileDir ".msbmc-initialized"

if ((Test-Path $profileDir) -and (Test-Path $markerFile)) {
    Write-Host "[OK] Chrome profile already created by Packer" -ForegroundColor Green
    Write-Host "Profile directory: $profileDir" -ForegroundColor Cyan
    Write-Host ""
    
    # Set Chrome as default browser (HKCU for current user)
    Write-Host "Setting Chrome as default browser for current user..." -ForegroundColor Yellow
    try {
        $protocols = @('http', 'https', 'ftp')
        foreach ($protocol in $protocols) {
            $regPath = "HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\$protocol\UserChoice"
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force | Out-Null
            }
            Set-ItemProperty -Path $regPath -Name "ProgId" -Value "ChromeHTML" -Force -ErrorAction SilentlyContinue
        }
        
        $htmlRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice"
        if (-not (Test-Path $htmlRegPath)) {
            New-Item -Path $htmlRegPath -Force | Out-Null
        }
        Set-ItemProperty -Path $htmlRegPath -Name "ProgId" -Value "ChromeHTML" -Force -ErrorAction SilentlyContinue
        
        Write-Host "[OK] Chrome set as default browser" -ForegroundColor Green
    } catch {
        Write-Host "[WARN] Failed to set default browser: $_" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "VPN Extensions:" -ForegroundColor Yellow
    Write-Host "  Extensions can be installed manually from Chrome Web Store:" -ForegroundColor Cyan
    Write-Host "  - NordVPN: https://chrome.google.com/webstore/detail/fjoaledfpmneenckfbpdfhkmimnjocfa" -ForegroundColor Gray
    Write-Host "  - Surfshark: https://chrome.google.com/webstore/detail/ailoabdmgclmfmhdagmlohpjlbpffblp" -ForegroundColor Gray
    Write-Host ""
    
    $installExtensions = Read-Host "Do you want to open Chrome to install VPN extensions now? (y/n)"
    if ($installExtensions -eq 'y' -or $installExtensions -eq 'yes') {
        Write-Host ""
        Write-Host "Opening Chrome with extension pages..." -ForegroundColor Cyan
        
        $chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
        $nordvpnUrl = "https://chrome.google.com/webstore/detail/fjoaledfpmneenckfbpdfhkmimnjocfa"
        $surfsharkUrl = "https://chrome.google.com/webstore/detail/ailoabdmgclmfmhdagmlohpjlbpffblp"
        
        Start-Process $chromePath -ArgumentList @(
            "--user-data-dir=$profileDir",
            $nordvpnUrl,
            $surfsharkUrl
        )
        
        Write-Host ""
        Write-Host "Please install extensions in Chrome, then press any key to continue..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
    
    # Create desktop shortcut for Chrome with persistent profile
    Write-Host ""
    Write-Host "Creating desktop shortcut..." -ForegroundColor Yellow
    try {
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $shortcutPath = Join-Path $desktopPath "Chrome (MSBMC Profile).lnk"
        
        if (-not (Test-Path $shortcutPath)) {
            $WScriptShell = New-Object -ComObject WScript.Shell
            $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
            $shortcut.Arguments = "--user-data-dir=`"$profileDir`""
            $shortcut.IconLocation = "C:\Program Files\Google\Chrome\Application\chrome.exe,0"
            $shortcut.Description = "Chrome with MSBMC persistent profile"
            $shortcut.Save()
            
            Write-Host "[OK] Desktop shortcut created" -ForegroundColor Green
            Write-Host "    Use this shortcut to always launch Chrome with persistent profile" -ForegroundColor Cyan
        } else {
            Write-Host "[OK] Desktop shortcut already exists" -ForegroundColor Green
        }
        
        # Remove default Chrome shortcuts (not using persistent profile)
        Write-Host ""
        Write-Host "Cleaning up default Chrome shortcuts..." -ForegroundColor Yellow
        
        $shortcutsToRemove = @(
            "Google Chrome.lnk",
            "Chrome.lnk"
        )
        
        foreach ($shortcut in $shortcutsToRemove) {
            $path = Join-Path $desktopPath $shortcut
            if (Test-Path $path) {
                Remove-Item $path -Force
                Write-Host "  Removed: $shortcut" -ForegroundColor Cyan
            }
        }
        
        # Unpin Chrome from taskbar (remove taskbar shortcut)
        Write-Host ""
        Write-Host "Unpinning Chrome from taskbar..." -ForegroundColor Yellow
        try {
            $taskbarPath = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
            Get-ChildItem $taskbarPath -Filter "*Chrome*.lnk" | ForEach-Object {
                Remove-Item $_.FullName -Force
                Write-Host "  Removed taskbar pin: $($_.Name)" -ForegroundColor Cyan
            }
        } catch {
            Write-Host "  [INFO] No taskbar pins to remove" -ForegroundColor Gray
        }
        
        Write-Host "[OK] Chrome shortcuts cleaned up" -ForegroundColor Green
        
    } catch {
        Write-Host "[WARN] Could not create desktop shortcut: $_" -ForegroundColor Yellow
    }
    
} else {
    Write-Host "[ERROR] Chrome profile not found!" -ForegroundColor Red
    Write-Host ""
    Write-Host "The Chrome profile should have been created by Packer (setup-chrome-profile.ps1)." -ForegroundColor Yellow
    Write-Host "This indicates the BASE AMI was not built correctly." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Expected location: $profileDir" -ForegroundColor Cyan
    Write-Host "Expected marker: $(Join-Path $profileDir '.msbmc-initialized')" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Please rebuild the BASE AMI using Packer before continuing." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to continue to next step..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

Write-Host ""
Write-Host "Press any key to continue..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

# =============================================================================
# Step 4: Chrome Watchdog Service
# =============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "[4/5] Chrome Watchdog Service" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($watchdogDone) {
    Write-Host "[SKIP] Chrome Watchdog already installed" -ForegroundColor Green
    Write-Host ""
    $task = Get-ScheduledTask -TaskName "MSBMC-ChromeWatchdog" -ErrorAction SilentlyContinue
    if ($task) {
        Write-Host "Task state: $($task.State)" -ForegroundColor Cyan
        Write-Host "Note: Task is DISABLED by default, enabled only in Chrome-only mode" -ForegroundColor Cyan
    }
} else {
    $watchdogScript = Join-Path $scriptsDir "install-watchdog.ps1"
    if (Test-Path $watchdogScript) {
        Write-Host "Installing Chrome Watchdog..." -ForegroundColor Yellow
        Write-Host "This will create a background task to keep Chrome maximized." -ForegroundColor Cyan
        Write-Host "Task is DISABLED by default (enabled only in Chrome-only mode)." -ForegroundColor Cyan
        Write-Host ""
        
        & $watchdogScript
        
        Write-Host ""
        Write-Host "Verifying installation..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        
        $task = Get-ScheduledTask -TaskName "MSBMC-ChromeWatchdog" -ErrorAction SilentlyContinue
        if ($task) {
            Write-Host "[OK] Chrome Watchdog installed successfully" -ForegroundColor Green
            Write-Host "Task state: $($task.State)" -ForegroundColor Cyan
        } else {
            Write-Host "[ERROR] Watchdog task was not created" -ForegroundColor Red
        }
    } else {
        Write-Host "[ERROR] Script install-watchdog.ps1 not found" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Press any key to continue..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

# =============================================================================
# Step 5: Kiosk Mode Toggle System
# =============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "[5/5] Kiosk Mode Toggle System" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($kioskDone) {
    Write-Host "[SKIP] Kiosk Mode already configured" -ForegroundColor Green
    Write-Host ""
    
    # Check current state
    $flagFile = "C:\ProgramData\msbmc-chrome-only.flag"
    if (Test-Path $flagFile) {
        Write-Host "Current mode: Chrome-only ACTIVE" -ForegroundColor Yellow
        Write-Host "Press ESC 3x + enter password 'msbmc2024' to exit" -ForegroundColor Cyan
    } else {
        Write-Host "Current mode: Normal desktop (Chrome-only DISABLED)" -ForegroundColor Green
        Write-Host "Press ESC 3x to enable Chrome-only mode" -ForegroundColor Cyan
    }
} else {
    $kioskScript = Join-Path $scriptsDir "configure-kiosk.ps1"
    if (Test-Path $kioskScript) {
        Write-Host "Configuring Kiosk Mode toggle system..." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "This will install:" -ForegroundColor Cyan
        Write-Host "  - ESC keyboard monitor (runs at logon)" -ForegroundColor White
        Write-Host "  - Chrome-only mode toggle script" -ForegroundColor White
        Write-Host "  - Password protection for exit (msbmc2024)" -ForegroundColor White
        Write-Host ""
        Write-Host "Toggle behavior:" -ForegroundColor Yellow
        Write-Host "  Normal → Chrome-only:  Press ESC 3 times" -ForegroundColor White
        Write-Host "  Chrome-only → Normal:  Press ESC 3 times + enter 'msbmc2024'" -ForegroundColor White
        Write-Host ""
        
        & $kioskScript
        
        Write-Host ""
        Write-Host "Verifying installation..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        
        $escTask = Get-ScheduledTask -TaskName "MSBMC-EscMonitor" -ErrorAction SilentlyContinue
        $toggleScript = "C:\ProgramData\msbmc-chrome-only-toggle.ps1"
        $escMonitorScript = "C:\ProgramData\msbmc-esc-monitor.ps1"
        
        if ($escTask -and (Test-Path $toggleScript) -and (Test-Path $escMonitorScript)) {
            Write-Host "[OK] Kiosk Mode configured successfully" -ForegroundColor Green
            Write-Host ""
            Write-Host "ESC Monitor task state: $($escTask.State)" -ForegroundColor Cyan
            
            $flagFile = "C:\ProgramData\msbmc-chrome-only.flag"
            if (Test-Path $flagFile) {
                Write-Host "Current mode: Chrome-only ACTIVE" -ForegroundColor Yellow
            } else {
                Write-Host "Current mode: Normal desktop (Chrome-only DISABLED)" -ForegroundColor Green
            }
        } else {
            Write-Host "[ERROR] Kiosk Mode configuration incomplete" -ForegroundColor Red
            if (-not $escTask) {
                Write-Host "  Missing: MSBMC-EscMonitor scheduled task" -ForegroundColor Red
            }
            if (-not (Test-Path $toggleScript)) {
                Write-Host "  Missing: msbmc-chrome-only-toggle.ps1 script" -ForegroundColor Red
            }
            if (-not (Test-Path $escMonitorScript)) {
                Write-Host "  Missing: msbmc-esc-monitor.ps1 script" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "[ERROR] Script configure-kiosk.ps1 not found" -ForegroundColor Red
    }
}

# =============================================================================
# Final Summary
# =============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Configuration Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Re-check all configurations
Write-Host "Final Status:" -ForegroundColor Yellow
Write-Host ""

$finalResolution = Test-ResolutionConfigured
$finalVBAudio = Test-VBAudioInstalled
$finalChrome = Test-ChromeConfigured
$finalWatchdog = Test-ChromeWatchdogInstalled
$finalKiosk = Test-KioskConfigured

$s1 = if ($finalResolution) { "[OK]" } else { "[  ]" }
$s2 = if ($finalVBAudio) { "[OK]" } else { "[  ]" }
$s3 = if ($finalChrome) { "[OK]" } else { "[  ]" }
$s4 = if ($finalWatchdog) { "[OK]" } else { "[  ]" }
$s5 = if ($finalKiosk) { "[OK]" } else { "[  ]" }

$c1 = if ($finalResolution) { "Green" } else { "Red" }
$c2 = if ($finalVBAudio) { "Green" } else { "Red" }
$c3 = if ($finalChrome) { "Green" } else { "Red" }
$c4 = if ($finalWatchdog) { "Green" } else { "Red" }
$c5 = if ($finalKiosk) { "Green" } else { "Red" }

Write-Host "  $s1 Display Resolution" -ForegroundColor $c1
Write-Host "  $s2 VB-Audio Virtual Cable" -ForegroundColor $c2
Write-Host "  $s3 Chrome Configuration" -ForegroundColor $c3
Write-Host "  $s4 Chrome Watchdog" -ForegroundColor $c4
Write-Host "  $s5 Kiosk Mode Toggle" -ForegroundColor $c5
Write-Host ""

# Show active MSBMC tasks
Write-Host "Active MSBMC Scheduled Tasks:" -ForegroundColor Yellow
$msbmcTasks = Get-ScheduledTask | Where-Object { $_.TaskName -like "*MSBMC*" }
if ($msbmcTasks) {
    foreach ($task in $msbmcTasks) {
        $stateColor = switch ($task.State) {
            'Ready' { 'Green' }
            'Running' { 'Cyan' }
            'Disabled' { 'Yellow' }
            default { 'Gray' }
        }
        Write-Host "  - $($task.TaskName): $($task.State)" -ForegroundColor $stateColor
    }
} else {
    Write-Host "  (none found)" -ForegroundColor Gray
}
Write-Host ""

# Check if reboot required
$rebootRequired = $false
if ($finalVBAudio -and -not (Get-WmiObject Win32_SoundDevice | Where-Object { $_.Name -like "*VB-Audio*" })) {
    $rebootRequired = $true
}

if ($rebootRequired) {
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "REBOOT REQUIRED" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "A reboot is required to:" -ForegroundColor White
    Write-Host "  - Activate VB-Audio driver" -ForegroundColor White
    Write-Host "  - Apply all system changes" -ForegroundColor White
    Write-Host "  - Enable scheduled tasks" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Do you want to reboot now? (y/n): " -ForegroundColor Yellow -NoNewline
    $rebootResponse = Read-Host
    $rebootResponse = $rebootResponse.Trim().ToLower()
    
    if ($rebootResponse -eq 'y' -or $rebootResponse -eq 'yes') {
        Write-Host ""
        Write-Host "System will reboot in 10 seconds..." -ForegroundColor Yellow
        Write-Host "Press Ctrl+C to cancel" -ForegroundColor Gray
        Start-Sleep -Seconds 10
        Restart-Computer -Force
    } else {
        Write-Host ""
        Write-Host "Please reboot manually when ready:" -ForegroundColor Cyan
        Write-Host "  Restart-Computer -Force" -ForegroundColor White
    }
} else {
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "System Ready" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
}

Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "  1. Test Chrome-only mode: Press ESC 3 times" -ForegroundColor White
Write-Host "  2. Exit Chrome-only mode: Press ESC 3 times + enter 'msbmc2024'" -ForegroundColor White
Write-Host "  3. Test VNC access: http://<instance-ip>:6080" -ForegroundColor White
Write-Host "  4. Verify VB-Audio: Check sound devices in Control Panel" -ForegroundColor White
Write-Host "  5. Create FINAL AMI: EC2 Console → Actions → Image → Create Image" -ForegroundColor White
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "ERROR: Configuration Failed" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Error details:" -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Script location:" -ForegroundColor Yellow
    Write-Host $_.InvocationInfo.ScriptName -ForegroundColor Gray
    Write-Host "Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Stack trace:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    Write-Host ""
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "  1. Check that all scripts exist in C:\MSBMC\Scripts\" -ForegroundColor White
    Write-Host "  2. Verify BASE AMI was built successfully with Packer" -ForegroundColor White
    Write-Host "  3. Ensure running as 'msbmc' user (not Administrator)" -ForegroundColor White
    Write-Host "  4. Check Windows Event Viewer for detailed errors" -ForegroundColor White
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit 1
}
