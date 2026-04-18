# Build script for RakuPrint
# Usage: .\scripts\build.ps1 [-SkipInstaller]

param(
    [switch]$SkipInstaller
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$InstallerScript = Join-Path $ProjectRoot "installer.iss"

Push-Location $ProjectRoot
try {
    Write-Host "=== RakuPrint Build Script ===" -ForegroundColor Cyan

    # Clean previous build
    Write-Host "`n[1/3] Cleaning previous build..." -ForegroundColor Yellow
    if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
    if (Test-Path "build") { Remove-Item -Recurse -Force "build" }

    # Build with PyInstaller
    Write-Host "`n[2/3] Building with PyInstaller..." -ForegroundColor Yellow
    if (-not (Test-Path $VenvPython)) {
        throw "Virtual environment python not found: $VenvPython"
    }
    & $VenvPython -m PyInstaller "RakuPrint.spec" --noconfirm

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
            if (-not (Test-Path $InstallerScript)) {
                throw "Installer script not found: $InstallerScript"
            }
            & $iscc $InstallerScript
            if ($LASTEXITCODE -eq 0) {
                Write-Host "Installer build completed!" -ForegroundColor Green

                # Generate SHA256SUMS.txt for integrity verification
                $installerExe = Get-ChildItem "installer_output\*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($installerExe) {
                    $hash = & $VenvPython -c "import hashlib,sys; print(hashlib.sha256(open(sys.argv[1],'rb').read()).hexdigest())" $installerExe.FullName
                    "$hash  $($installerExe.Name)" | Out-File -FilePath "installer_output\SHA256SUMS.txt" -Encoding utf8 -NoNewline
                    Write-Host "SHA256SUMS.txt generated: $hash" -ForegroundColor Green
                }
            } else {
                throw "Installer build failed! (ISCC exit code: $LASTEXITCODE)"
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
        Write-Host "  installer_output\SHA256SUMS.txt" -ForegroundColor Gray
        Write-Host "" -ForegroundColor White
        Write-Host "Upload both files to the GitHub release." -ForegroundColor Yellow
    }
}
finally {
    Pop-Location
}
