@echo off
setlocal EnableExtensions DisableDelayedExpansion
title Excel Skills - Desktop Excel Automation
echo.
echo Excel Skills - Excel 2010+ Automation
echo This tool creates new output files. Existing originals are not modified.
echo.
where powershell.exe >nul 2>&1
if errorlevel 1 (
  echo PowerShell could not be found on this Windows computer.
  exit /b 2
)
powershell.exe -NoProfile -File "%~dp0ExcelSkills.ps1" -Action Menu
set "RESULT=%ERRORLEVEL%"
echo.
if not "%RESULT%"=="0" echo The operation failed. Review the error above.
pause
exit /b %RESULT%
