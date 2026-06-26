<#
.SYNOPSIS
    Removes the Excel Sheet Navigator add-in for the current user.
    Intended to be launched by double-clicking Uninstall.bat.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'SilentlyContinue'

Write-Host ''
Write-Host 'Excel Sheet Navigator - Uninstaller' -ForegroundColor Cyan
Write-Host '==================================='
Write-Host ''

if (Get-Process -Name excel -ErrorAction SilentlyContinue) {
    Write-Host 'Excel is open. Please CLOSE Excel completely,' -ForegroundColor Yellow
    Write-Host 'then press Enter to continue...' -ForegroundColor Yellow
    [void](Read-Host)
}

Write-Host '  Removing Excel registration...'
Remove-Item 'HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelSheetNavigator' -Recurse -Force

Write-Host '  Removing program files...'
$dest = Join-Path $env:LOCALAPPDATA 'ExcelSheetNavigator'
if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }

Write-Host '  Removing trusted certificate...'
$thumb = 'FFE84A0F4BF082EB0D87E9A64AF121FB474BB7D4'
foreach ($store in 'TrustedPublisher', 'Root') {
    $p = "Cert:\CurrentUser\$store\$thumb"
    if (Test-Path $p) { Remove-Item $p -Force }
}

Write-Host ''
Write-Host 'Uninstalled. Restart Excel if it is open.' -ForegroundColor Green
Write-Host ''
