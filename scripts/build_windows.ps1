<#
.SYNOPSIS
    Repeatable Build Script for GST & Accounting Ledger Reconciliation Tool (Windows x64).

.DESCRIPTION
    1. Validates Python 3.12 (64-bit) installation.
    2. Initializes or updates virtual environment (.venv).
    3. Installs dependencies from Requirements.txt.
    4. Executes the automated test suite (tests\test_ledger_recon.py).
    5. Builds the standalone Windows executable using PyInstaller.
    6. Verifies executable creation and generates checksum.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $ProjectRoot

Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " Building GST & Accounting Ledger Reconciliation Workstation" -ForegroundColor Cyan
Write-Host " Project Root: $ProjectRoot" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan

# 1. Verify Python
Write-Host "[1/6] Checking Python installation..." -ForegroundColor Yellow
$PythonCmd = Get-Command "py" -ErrorAction SilentlyContinue
if ($PythonCmd) {
    $PythonExe = "py -3.12"
} else {
    $PythonCmd = Get-Command "python" -ErrorAction SilentlyContinue
    if (-not $PythonCmd) {
        Write-Error "Python is not found on PATH. Please install Python 3.12 64-bit."
        exit 1
    }
    $PythonExe = "python"
}

# 2. Virtual Environment
Write-Host "[2/6] Preparing virtual environment (.venv)..." -ForegroundColor Yellow
$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $VenvPython)) {
    Write-Host "Creating new virtual environment at .venv..."
    & py -3.12 -m venv .venv
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to create virtual environment."
        exit 1
    }
}

# 3. Install Dependencies
Write-Host "[3/6] Installing dependencies from Requirements.txt..." -ForegroundColor Yellow
& $VenvPython -m pip install --quiet --upgrade pip
& $VenvPython -m pip install --quiet -r Requirements.txt
if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to install required dependencies."
    exit 1
}

# 4. Run Automated Test Suite
Write-Host "[4/6] Executing automated test suite..." -ForegroundColor Yellow
& $VenvPython -m unittest tests/test_ledger_recon.py -v
if ($LASTEXITCODE -ne 0) {
    Write-Error "Automated test suite failed! Halting build."
    exit 1
}

# 5. Build Executable with PyInstaller
Write-Host "[5/6] Building standalone executable with PyInstaller..." -ForegroundColor Yellow
& $VenvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name "GST-Reconciliation-Tool" `
    --paths "src" `
    --hidden-import openpyxl `
    --hidden-import pandas `
    --hidden-import xlrd `
    --hidden-import reconcile `
    --hidden-import ledger_recon `
    src\gst_recon_gui.py

if ($LASTEXITCODE -ne 0) {
    Write-Error "PyInstaller build failed."
    exit 1
}

# 6. Verify Output
$TargetExe = Join-Path $ProjectRoot "dist\GST-Reconciliation-Tool.exe"
if (Test-Path $TargetExe) {
    $ExeItem = Get-Item $TargetExe
    $Hash = (Get-FileHash -Path $TargetExe -Algorithm SHA256).Hash
    Write-Host "========================================================================" -ForegroundColor Green
    Write-Host " BUILD SUCCESSFUL!" -ForegroundColor Green
    Write-Host " Executable Path : $($ExeItem.FullName)" -ForegroundColor Green
    Write-Host " File Size       : $($ExeItem.Length) bytes ($([Math]::Round($ExeItem.Length / 1MB, 2)) MB)" -ForegroundColor Green
    Write-Host " SHA-256 Checksum: $Hash" -ForegroundColor Green
    Write-Host "========================================================================" -ForegroundColor Green
} else {
    Write-Error "Target executable was not found at $TargetExe!"
    exit 1
}

