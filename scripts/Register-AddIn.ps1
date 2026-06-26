<#
.SYNOPSIS
    Registers the Excel Sheet Navigator VSTO add-in for the current user (no admin rights required).

.DESCRIPTION
    - Trusts the development signing certificate (adds it to the current user's
      Trusted Publishers and Root stores) so Office loads the signed add-in
      without a security prompt.
    - Writes the HKCU add-in registration so Excel discovers and loads the add-in.

    This performs a local "ClickOnce vstolocal" install that points directly at the
    build output, which is ideal for development and side-by-side testing.

.PARAMETER Configuration
    Build configuration to register (Debug or Release). Defaults to Debug.

.EXAMPLE
    .\Register-AddIn.ps1
    .\Register-AddIn.ps1 -Configuration Release
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug'
)

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot  = Split-Path -Parent $scriptDir
$vstoPath  = Join-Path $repoRoot "ExcelSheetNavigator\bin\$Configuration\ExcelSheetNavigator.vsto"
$cerPath   = Join-Path $scriptDir 'ExcelSheetNavigator.cer'

if (-not (Test-Path $vstoPath)) {
    throw "Deployment manifest not found: $vstoPath`nBuild the solution first (Configuration=$Configuration)."
}
if (-not (Test-Path $cerPath)) {
    throw "Signing certificate not found: $cerPath"
}

# 1. Trust the publisher certificate for the current user.
foreach ($store in @('TrustedPublisher', 'Root')) {
    Write-Host "Trusting certificate in CurrentUser\$store ..."
    Import-Certificate -FilePath $cerPath -CertStoreLocation "Cert:\CurrentUser\$store" | Out-Null
}

# 2. Register the add-in under the current user's Excel add-ins.
$addinKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelSheetNavigator'
if (-not (Test-Path $addinKey)) {
    New-Item -Path $addinKey -Force | Out-Null
}

New-ItemProperty -Path $addinKey -Name 'FriendlyName' -Value 'Excel Sheet Navigator' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name 'Description'  -Value 'Dockable worksheet navigation panel.' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $addinKey -Name 'Manifest'     -Value ("$vstoPath|vstolocal") -PropertyType String -Force | Out-Null

Write-Host ''
Write-Host 'Excel Sheet Navigator registered successfully.' -ForegroundColor Green
Write-Host "Manifest: $vstoPath|vstolocal"
Write-Host 'Restart Excel to load the add-in.'
