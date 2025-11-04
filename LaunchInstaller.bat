@echo off
echo ========================================
echo PowerShell Module Installer
echo ========================================
echo.
echo This will install required PowerShell modules for Microsoft 365 management.
echo.
echo Press any key to continue or Ctrl+C to cancel...
pause >nul
echo.
echo Starting installation...
echo.

SET scriptPath=%~dp0PowerShellInstaller.ps1
powershell.exe -ExecutionPolicy Bypass -File "%scriptPath%" -SkipExecutionPolicy

echo.
echo ========================================
echo Installation complete!
echo ========================================
echo.
echo Check the log file in your Documents folder for details.
echo You may need to restart PowerShell for changes to take effect.
echo.
pause
