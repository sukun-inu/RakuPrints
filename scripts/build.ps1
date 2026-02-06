# Build script for RakuPrint
# Usage: .\scripts\build.ps1 [-SkipInstaller]

param(
    [switch]$SkipInstaller
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot

Push-Location $ProjectRoot
try {
    Write-Host "=== RakuPrint Build Script ===" -ForegroundColor Cyan

    # Clean previous build
    Write-Host "`n[1/3] Cleaning previous build..." -ForegroundColor Yellow
    if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
    if (Test-Path "build") { Remove-Item -Recurse -Force "build" }

    # Build with PyInstaller
    Write-Host "`n[2/3] Building with PyInstaller..." -ForegroundColor Yellow
    python -m PyInstaller "RakuPrint.spec" --noconfirm

    if ($LASTEXITCODE -ne 0) {
        Write-Host "PyInstaller build failed!" -ForegroundColor Red
        exit 1
    }
    Write-Host "PyInstaller build completed!" -ForegroundColor Green

    # Build installer with Inno Setup
    if (-not $SkipInstaller) {
        Write-Host "`n[3/3] Building installer with Inno Setup..." -ForegroundColor Yellow
        
        $iscc = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
        if (-not (Test-Path $iscc)) {
            $iscc = "C:\Program Files\Inno Setup 6\ISCC.exe"
        }
        
        if (Test-Path $iscc) {
            & $iscc "installer.iss"
            if ($LASTEXITCODE -eq 0) {
                Write-Host "Installer build completed!" -ForegroundColor Green
            } else {
                Write-Host "Installer build failed!" -ForegroundColor Red
            }
        } else {
            Write-Host "Inno Setup not found. Skipping installer build." -ForegroundColor Yellow
            Write-Host "Install from: https://jrsoftware.org/isdl.php" -ForegroundColor Gray
        }
    } else {
        Write-Host "`n[3/3] Skipping installer build..." -ForegroundColor Yellow
    }

    Write-Host "`n=== Build Complete ===" -ForegroundColor Cyan
    Write-Host "Output:" -ForegroundColor White
    Write-Host "  dist\RakuPrint\RakuPrint.exe" -ForegroundColor Gray
    if (-not $SkipInstaller) {
        Write-Host "  installer_output\RakuPrint_Setup_*.exe" -ForegroundColor Gray
    }
}
finally {
    Pop-Location
}
