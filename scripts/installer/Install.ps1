<#
.SYNOPSIS
    Self-contained installer for the Excel Sheet Navigator add-in.
    Runs on any machine - no Visual Studio, no build, no administrator rights.

.DESCRIPTION
    Copies the add-in into the current user's profile, trusts its signing
    certificate for the current user, and registers it with Excel (all under
    HKCU / the current-user certificate stores - nothing machine-wide).

    Intended to be launched by double-clicking Install.bat.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

try {
    $here   = Split-Path -Parent $MyInvocation.MyCommand.Path
    $appSrc = Join-Path $here 'app'
    $cer    = Join-Path $here 'ExcelSheetNavigator.cer'
    $dest   = Join-Path $env:LOCALAPPDATA 'ExcelSheetNavigator'

    Write-Host ''
    Write-Host 'Excel Sheet Navigator - Installer' -ForegroundColor Cyan
    Write-Host '================================='
    Write-Host ''

    if (-not (Test-Path $appSrc)) {
        throw "The 'app' folder is missing. Please extract the WHOLE ZIP first, then run Install.bat again."
    }
    if (-not (Test-Path $cer)) {
        throw "The certificate file 'ExcelSheetNavigator.cer' is missing from the package."
    }

    # 1) Check for the Office runtime the add-in needs (informational only).
    $vstoKeys = @(
        'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R',
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4R',
        'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4',
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4'
    )
    $hasVsto = $false
    foreach ($k in $vstoKeys) { if (Test-Path $k) { $hasVsto = $true; break } }
    if (-not $hasVsto) {
        Write-Host 'NOTE: The Visual Studio 2010 Tools for Office Runtime was not detected.' -ForegroundColor Yellow
        Write-Host '      If the panel does not appear in Excel, install it (free) from:'    -ForegroundColor Yellow
        Write-Host '      https://aka.ms/vstoredist'                                          -ForegroundColor Yellow
        Write-Host ''
    }

    # 2) Excel must be closed so files are not locked.
    if (Get-Process -Name excel -ErrorAction SilentlyContinue) {
        Write-Host 'Excel is currently open. Please CLOSE Excel completely,' -ForegroundColor Yellow
        Write-Host 'then press Enter to continue...' -ForegroundColor Yellow
        [void](Read-Host)
    }

    # 3) Copy the program files into the user profile.
    Write-Host '  Copying program files...'
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    Copy-Item -Path (Join-Path $appSrc '*') -Destination $dest -Recurse -Force

    # 4) Trust the add-in's certificate (current user only).
    Write-Host '  Trusting the add-in certificate...'
    Import-Certificate -FilePath $cer -CertStoreLocation 'Cert:\CurrentUser\TrustedPublisher' | Out-Null
    Import-Certificate -FilePath $cer -CertStoreLocation 'Cert:\CurrentUser\Root' | Out-Null

    # 5) Register the add-in with Excel (current user).
    Write-Host '  Registering with Excel...'
    $key = 'HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelSheetNavigator'
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name 'FriendlyName' -Value 'Excel Sheet Navigator' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name 'Description'  -Value 'Dockable worksheet navigation panel.' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $key -Name 'Manifest'     -Value ((Join-Path $dest 'ExcelSheetNavigator.vsto') + '|vstolocal') -PropertyType String -Force | Out-Null

    Write-Host ''
    Write-Host 'Installed successfully!' -ForegroundColor Green
    Write-Host 'Open Excel - the Sheet Navigator panel appears on the right.'
    Write-Host 'You can show or hide it from:  Home tab > Navigator > Sheet Navigator.'
    Write-Host ''
}
catch {
    Write-Host ''
    Write-Host ('Installation failed: ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host 'If a file was in use, fully close Excel and run Install.bat again.'
    Write-Host ''
    exit 1
}
