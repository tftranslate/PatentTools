[CmdletBinding()]
param(
    [string]$Version,
    [switch]$SkipDotmBuild
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$buildDir = Join-Path $root "build"

# Read version from VERSION file if not provided
if ([string]::IsNullOrWhiteSpace($Version)) {
    $versionFile = Join-Path $root "VERSION"
    if (Test-Path $versionFile) {
        $Version = Get-Content $versionFile -Raw | Select-Object -First 1 | ForEach-Object { $_.Trim() }
        Write-Host "Version from VERSION file: $Version" -ForegroundColor Gray
    }
    else {
        throw "VERSION file not found at $versionFile"
    }
}

Write-Host "`n=== PatentTools Build Pipeline ===" -ForegroundColor Cyan
Write-Host "Version: $Version" -ForegroundColor Gray

# Check ISCC.exe
$isccPath = "C:\Program Files\Inno Setup 7\ISCC.exe"
if (-not (Test-Path $isccPath)) {
    $fallbackPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    if (Test-Path $fallbackPath) {
        Write-Host "⚠️  Inno Setup 7 not found, using version 6..." -ForegroundColor Yellow
        $isccPath = $fallbackPath
    } else {
        Write-Host "❌ Inno Setup not found!" -ForegroundColor Red
        exit 1
    }
}

# Build .dotm if needed
$dotmPath = Join-Path $buildDir "PatentTools.dotm"

if (-not (Test-Path $dotmPath)) {
    Write-Host "`n=== Step 1: Building DOTM template ===" -ForegroundColor Yellow
    & powershell.exe -NoProfile -ExecutionPolicy Bypass `
        -File (Join-Path $root "scripts\Build-PatentTools.ps1") -Version $Version
    
    if ($LASTEXITCODE -ne 0) { throw "DOTM build failed!" }
    Write-Host "✅ .dotm created with version $Version" -ForegroundColor Green
}
else {
    Write-Host "`n=== Step 1: .dotm already exists (skipping rebuild) ===" -ForegroundColor Gray
}

Write-Host "`n=== Step 2: Compiling installer ===" -ForegroundColor Yellow

$issFile = Join-Path $root "installers\inno-setup\PatentToolsInstaller.iss"

$outputBaseName = 'PatentTools-v' + $Version + '-Setup'


# Update the .iss file with correct output filename before compiling
$issContent = Get-Content $issFile -Raw
$regexPattern = 'OutputBaseFilename=PatentTools_.*?_Setup'
$replacement = "OutputBaseFilename=$outputBaseName"
Write-Host "Replacing '$regexPattern' with '$replacement'" -ForegroundColor Gray
$issContent = $issContent -replace $regexPattern, $replacement
$issTempPath = Join-Path (Split-Path $issFile) "PatentToolsInstaller_temp.iss"
Set-Content -Path $issTempPath -Value $issContent

try {
    Write-Host "Compiler options: /Qp OutputBaseFilename=$outputBaseName" -ForegroundColor Gray
    
    # Call Inno Setup - pass version and scope 8always user currently) to preprocessor
    # Pass /DIsUser=1 for User installer (no UAC), don't pass or /DIsUser=0 for AllUsers
    $isUserValue = 1
    $preprocessorParams = @("/Qp", "/DAPP_VERSION=$Version")
    if ($isUserValue) {
        $preprocessorParams += "/DIsUser=$isUserValue"
    }
    & $isccPath @preprocessorParams $issTempPath
} finally {
    Remove-Item $issTempPath -Force
}

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Inno Setup compilation failed!" -ForegroundColor Red
    exit 1
}

Write-Host "`n=== ✅ Build completed ===" -ForegroundColor Green
$installerPath = Join-Path $buildDir "$outputBaseName.exe"
if (Test-Path $installerPath) {
    Write-Host "Installer created:" -ForegroundColor Cyan
    Write-Host "  $installerPath" -ForegroundColor White
    $size = [math]::Round((Get-Item $installerPath).Length / 1KB, 2)
    Write-Host "  Size: $([math]::Round($size, 0)) KB" -ForegroundColor Gray
}
