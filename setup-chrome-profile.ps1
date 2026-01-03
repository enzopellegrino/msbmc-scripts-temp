# =============================================================================
# Setup Chrome Profile for msbmc User (Packer execution)
# Creates Chrome profile with VPN extensions pre-installed
# =============================================================================

$ErrorActionPreference = 'Continue'

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Setting up Chrome Profile for msbmc" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"

# Verify Chrome is installed
if (-not (Test-Path $chromePath)) {
    Write-Host "[ERROR] Chrome not found at $chromePath" -ForegroundColor Red
    exit 1
}

# Verify msbmc user exists
$msbmcUser = Get-LocalUser -Name "msbmc" -ErrorAction SilentlyContinue
if (-not $msbmcUser) {
    Write-Host "[ERROR] User 'msbmc' does not exist" -ForegroundColor Red
    Write-Host "  This script must run after msbmc user creation" -ForegroundColor Yellow
    exit 1
}

Write-Host "[OK] Chrome found" -ForegroundColor Green
Write-Host "[OK] User msbmc exists" -ForegroundColor Green
Write-Host ""

# =============================================================================
# Step 1: Create Chrome Profile Directory
# =============================================================================
Write-Host "[1/4] Creating Chrome profile directory..." -ForegroundColor Yellow

$profileDir = if (Test-Path "D:\") { "D:\ChromeProfile" } else { "C:\ChromeProfile" }

if (-not (Test-Path $profileDir)) {
    New-Item -ItemType Directory -Path $profileDir -Force | Out-Null
    Write-Host "  Created: $profileDir" -ForegroundColor Green
} else {
    Write-Host "  Already exists: $profileDir" -ForegroundColor Green
}

# Create Default profile subdirectory
$defaultProfile = Join-Path $profileDir "Default"
if (-not (Test-Path $defaultProfile)) {
    New-Item -ItemType Directory -Path $defaultProfile -Force | Out-Null
}

# Create Extensions directory
$extensionsDir = Join-Path $defaultProfile "Extensions"
if (-not (Test-Path $extensionsDir)) {
    New-Item -ItemType Directory -Path $extensionsDir -Force | Out-Null
}

Write-Host "  [OK] Profile structure created" -ForegroundColor Green

# =============================================================================
# Step 2: Set Permissions for msbmc User
# =============================================================================
Write-Host ""
Write-Host "[2/4] Setting permissions for msbmc user..." -ForegroundColor Yellow

try {
    # Give msbmc full control over profile directory
    $acl = Get-Acl $profileDir
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        'msbmc',
        'FullControl',
        'ContainerInherit,ObjectInherit',
        'None',
        'Allow'
    )
    $acl.SetAccessRule($rule)
    Set-Acl $profileDir $acl
    
    Write-Host "  [OK] Permissions set for msbmc" -ForegroundColor Green
} catch {
    Write-Host "  [WARN] Failed to set permissions: $_" -ForegroundColor Yellow
}

# =============================================================================
# Step 3: Create Initial Preferences File
# =============================================================================
Write-Host ""
Write-Host "[3/4] Creating Chrome preferences..." -ForegroundColor Yellow

$preferencesFile = Join-Path $defaultProfile "Preferences"

# Basic preferences JSON
$preferences = @{
    "profile" = @{
        "name" = "MSBMC"
        "managed_user_id" = ""
    }
    "homepage" = "about:blank"
    "homepage_is_newtabpage" = $false
    "session" = @{
        "restore_on_startup" = 1  # Open blank page
    }
    "download" = @{
        "default_directory" = "D:\Downloads"
    }
    "extensions" = @{
        "settings" = @{}
    }
} | ConvertTo-Json -Depth 10

$preferences | Set-Content -Path $preferencesFile -Encoding UTF8 -Force
Write-Host "  [OK] Preferences file created" -ForegroundColor Green

# =============================================================================
# Step 4: Download and Install VPN Extensions
# =============================================================================
Write-Host ""
Write-Host "[4/4] Installing VPN extensions..." -ForegroundColor Yellow

$nordvpnExtId = "fjoaledfpmneenckfbpdfhkmimnjocfa"
$surfsharkExtId = "ailoabdmgclmfmhdagmlohpjlbpffblp"

# Create extension directories
$nordvpnDir = Join-Path $extensionsDir $nordvpnExtId
$surfsharkDir = Join-Path $extensionsDir $surfsharkExtId

New-Item -ItemType Directory -Path $nordvpnDir -Force | Out-Null
New-Item -ItemType Directory -Path $surfsharkDir -Force | Out-Null

Write-Host "  Extension directories created" -ForegroundColor Cyan
Write-Host ""

# Extensions will be installed on first Chrome launch
# Create marker files to indicate profile is initialized
$markerFile = Join-Path $profileDir ".msbmc-initialized"
"Profile created by Packer for msbmc user" | Set-Content -Path $markerFile -Force

Write-Host "  [INFO] VPN extensions will be installed on first Chrome launch" -ForegroundColor Cyan
Write-Host "  Extensions can be installed manually from Chrome Web Store:" -ForegroundColor Cyan
Write-Host "    - NordVPN: https://chrome.google.com/webstore/detail/$nordvpnExtId" -ForegroundColor Gray
Write-Host "    - Surfshark: https://chrome.google.com/webstore/detail/$surfsharkExtId" -ForegroundColor Gray
Write-Host ""

# =============================================================================
# Step 5: Configure Chrome as Default Browser for msbmc (via Registry)
# =============================================================================
Write-Host ""
Write-Host "[5/4] Configuring Chrome as default browser for msbmc..." -ForegroundColor Yellow

try {
    # Load msbmc user registry hive
    $msbmcSID = $msbmcUser.SID.Value
    $msbmcHivePath = "C:\Users\msbmc\NTUSER.DAT"
    
    if (Test-Path $msbmcHivePath) {
        # Load hive
        reg load "HKU\$msbmcSID" $msbmcHivePath 2>&1 | Out-Null
        
        # Set Chrome as default for protocols
        $protocols = @('http', 'https', 'ftp')
        foreach ($protocol in $protocols) {
            $regPath = "HKU:\$msbmcSID\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\$protocol\UserChoice"
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force -ErrorAction SilentlyContinue | Out-Null
            }
            Set-ItemProperty -Path $regPath -Name "ProgId" -Value "ChromeHTML" -Force -ErrorAction SilentlyContinue
        }
        
        # Set Chrome for .html files
        $htmlRegPath = "HKU:\$msbmcSID\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice"
        if (-not (Test-Path $htmlRegPath)) {
            New-Item -Path $htmlRegPath -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Set-ItemProperty -Path $htmlRegPath -Name "ProgId" -Value "ChromeHTML" -Force -ErrorAction SilentlyContinue
        
        # Unload hive
        [gc]::Collect()
        Start-Sleep -Seconds 2
        reg unload "HKU\$msbmcSID" 2>&1 | Out-Null
        
        Write-Host "  [OK] Chrome set as default browser in msbmc registry" -ForegroundColor Green
    } else {
        Write-Host "  [WARN] msbmc user profile not yet created" -ForegroundColor Yellow
        Write-Host "  Default browser will be set on first login" -ForegroundColor Cyan
    }
} catch {
    Write-Host "  [WARN] Could not set default browser in registry: $_" -ForegroundColor Yellow
    Write-Host "  Will be set by setup-ami-interactive.ps1" -ForegroundColor Cyan
}

# =============================================================================
# Step 5: Create Desktop Shortcut
# =============================================================================
Write-Host ""
Write-Host "[5/5] Creating desktop shortcut..." -ForegroundColor Yellow

try {
    # Get msbmc user's desktop path
    $msbmcProfilePath = "C:\Users\msbmc"
    $desktopPath = Join-Path $msbmcProfilePath "Desktop"
    
    # Create desktop folder if it doesn't exist
    if (-not (Test-Path $desktopPath)) {
        New-Item -ItemType Directory -Path $desktopPath -Force | Out-Null
    }
    
    $shortcutPath = Join-Path $desktopPath "Chrome (MSBMC Profile).lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $chromePath
    $shortcut.Arguments = "--user-data-dir=`"$profileDir`""
    $shortcut.IconLocation = "$chromePath,0"
    $shortcut.Description = "Chrome with MSBMC persistent profile"
    $shortcut.Save()
    
    # Set ownership to msbmc user
    icacls $shortcutPath /setowner "msbmc" /T /C 2>&1 | Out-Null
    
    Write-Host "  [OK] Desktop shortcut created: $shortcutPath" -ForegroundColor Green
} catch {
    Write-Host "  [WARN] Could not create desktop shortcut: $_" -ForegroundColor Yellow
}

# =============================================================================
# Summary
# =============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Chrome Profile Setup Complete" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Profile created at: $profileDir" -ForegroundColor Cyan
Write-Host "User: msbmc" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. VPN extensions can be installed from setup-ami-interactive.ps1" -ForegroundColor White
Write-Host "  2. Or manually from Chrome Web Store after first login" -ForegroundColor White
Write-Host ""
