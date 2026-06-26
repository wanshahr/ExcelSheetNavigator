<#
.SYNOPSIS
    Builds the add-in and produces a ready-to-share installer package (folder + ZIP).

.DESCRIPTION
    1. Builds the solution in the requested configuration.
    2. Collects the add-in binaries + manifests, the public certificate and the
       double-click installer into  dist\ExcelSheetNavigator\.
    3. Zips it to  dist\ExcelSheetNavigator.zip.

    Hand the ZIP to anyone: they extract it and double-click Install.bat.

.PARAMETER Configuration
    Build configuration to package (Debug or Release). Defaults to Release.

.EXAMPLE
    .\scripts\Build-Package.ps1
    .\scripts\Build-Package.ps1 -Configuration Debug
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot  = Split-Path -Parent $scriptDir
$solution  = Join-Path $repoRoot 'ExcelSheetNavigator.sln'

# --- Locate MSBuild -------------------------------------------------------
$msbuild = 'C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\MSBuild.exe'
if (-not (Test-Path $msbuild)) {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (Test-Path $vswhere) {
        $msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' |
                    Select-Object -First 1
    }
}
if (-not $msbuild -or -not (Test-Path $msbuild)) {
    throw 'MSBuild.exe could not be found. Open a Developer PowerShell or install the VS Office workload.'
}

# --- Build ----------------------------------------------------------------
Write-Host "Building $Configuration ..." -ForegroundColor Cyan
& $msbuild $solution /p:Configuration=$Configuration /t:Rebuild /v:minimal /nologo
if ($LASTEXITCODE -ne 0) { throw "Build failed (exit code $LASTEXITCODE)." }

# --- Lay out the package --------------------------------------------------
$binDir   = Join-Path $repoRoot "ExcelSheetNavigator\bin\$Configuration"
$distRoot = Join-Path $repoRoot 'dist'
$pkgDir   = Join-Path $distRoot 'ExcelSheetNavigator'
$appDir   = Join-Path $pkgDir 'app'

if (Test-Path $pkgDir) { Remove-Item $pkgDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $appDir | Out-Null

Write-Host 'Collecting add-in files...' -ForegroundColor Cyan
foreach ($pattern in '*.dll', '*.config', '*.manifest', '*.vsto') {
    Get-ChildItem -Path $binDir -Filter $pattern -File |
        Copy-Item -Destination $appDir -Force
}

# Public certificate (trusted by the installer) + the installer itself.
Copy-Item (Join-Path $scriptDir 'ExcelSheetNavigator.cer') -Destination $pkgDir -Force
Copy-Item (Join-Path $scriptDir 'installer\*')             -Destination $pkgDir -Recurse -Force

# --- Zip ------------------------------------------------------------------
$zipPath = Join-Path $distRoot 'ExcelSheetNavigator.zip'
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path $pkgDir -DestinationPath $zipPath -Force

Write-Host ''
Write-Host 'Package ready to share:' -ForegroundColor Green
Write-Host "  Folder: $pkgDir"
Write-Host "  ZIP:    $zipPath"
Write-Host ''
Write-Host 'Recipients just extract the ZIP and double-click Install.bat.'
