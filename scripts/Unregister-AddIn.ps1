<#
.SYNOPSIS
    Unregisters the Excel Sheet Navigator VSTO add-in for the current user.

.DESCRIPTION
    Removes the HKCU add-in registration and (optionally) removes the development
    signing certificate from the current user's Trusted Publishers and Root stores.

.PARAMETER KeepCertificate
    If set, the development signing certificate is left in the trust stores.

.EXAMPLE
    .\Unregister-AddIn.ps1
    .\Unregister-AddIn.ps1 -KeepCertificate
#>
[CmdletBinding()]
param(
    [switch]$KeepCertificate
)

$ErrorActionPreference = 'Stop'

$addinKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelSheetNavigator'
if (Test-Path $addinKey) {
    Remove-Item -Path $addinKey -Recurse -Force
    Write-Host 'Add-in registration removed.' -ForegroundColor Green
}
else {
    Write-Host 'Add-in registration was not present.'
}

if (-not $KeepCertificate) {
    $thumbprint = 'FFE84A0F4BF082EB0D87E9A64AF121FB474BB7D4'
    foreach ($store in @('TrustedPublisher', 'Root')) {
        $certPath = "Cert:\CurrentUser\$store\$thumbprint"
        if (Test-Path $certPath) {
            Remove-Item -Path $certPath -Force
            Write-Host "Removed development certificate from CurrentUser\$store."
        }
    }
}

Write-Host 'Restart Excel to complete removal.'
