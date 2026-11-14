@echo off
setlocal
echo ======================================================================
echo  GST ^& Accounting Ledger Reconciliation Tool - Automatic Web Setup
echo ======================================================================
echo.
echo [1/3] Downloading latest workstation from GitHub Releases...
set "DEST=%USERPROFILE%\GST-Reconciliation-Tool"
set "ZIP=%TEMP%\GST-Reconciliation-Tool-x64.zip"
set "URL=https://github.com/ifeelrohit/GST-Reconciliation-Tool/releases/latest/download/GST-Reconciliation-Tool-Windows-x64.zip"

powershell -NoProfile -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile('%URL%', '%ZIP%')"
if %errorlevel% neq 0 (
    echo [ERROR] Failed to download package. Please check your internet connection.
    pause
    exit /b 1
)

echo [2/3] Extracting files to: %DEST%
if not exist "%DEST%" mkdir "%DEST%"
powershell -NoProfile -Command "Expand-Archive -Path '%ZIP%' -DestinationPath '%DEST%' -Force"
del /f /q "%ZIP%" >nul 2>nul

echo [3/3] Creating Desktop Shortcut...
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), 'GST Reconciliation Tool.lnk')); $s.TargetPath = '%DEST%\GST-Reconciliation-Tool.exe'; $s.WorkingDirectory = '%DEST%'; $s.Save()" >nul 2>nul

echo.
echo ======================================================================
echo  Installation Complete! Launching tool...
echo ======================================================================
start "" "%DEST%\GST-Reconciliation-Tool.exe"
exit /b 0
