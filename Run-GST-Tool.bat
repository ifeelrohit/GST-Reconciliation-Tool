@echo off
setlocal
cd /d "%~dp0"

REM 1. Try local project virtual environment pythonw
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "src\gst_recon_gui.py"
    exit /b 0
)

REM 2. Try py launcher
where py >nul 2>nul
if %errorlevel% equ 0 (
    start "" py -3.12 "src\gst_recon_gui.py"
    exit /b 0
)

REM 3. Try system pythonw
where pythonw >nul 2>nul
if %errorlevel% equ 0 (
    start "" pythonw "src\gst_recon_gui.py"
    exit /b 0
)

REM 4. Fallback to python
where python >nul 2>nul
if %errorlevel% equ 0 (
    start "" python "src\gst_recon_gui.py"
    exit /b 0
)

echo [ERROR] Python was not found on this system.
echo Please run dist\GST-Reconciliation-Tool.exe or install Python 3.12.
pause
exit /b 1
