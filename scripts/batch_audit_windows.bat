@echo off
setlocal EnableExtensions DisableDelayedExpansion
title Excel Skills - Batch XLSX Audit
echo.
echo Excel Skills - read-only audit
echo ------------------------------
set /p "SOURCE=Enter the folder containing Excel files: "
if not defined SOURCE (
  echo No folder entered.
  exit /b 2
)
set "SOURCE=%SOURCE:"=%"
if not exist "%SOURCE%\." (
  echo Folder not found: "%SOURCE%"
  exit /b 2
)
set /p "LOGFILE=Enter output log path, or press Enter for audit-results.txt: "
if not defined LOGFILE set "LOGFILE=%~dp0audit-results.txt"
set "LOGFILE=%LOGFILE:"=%"
set "PYRUN=python"
where py >nul 2>&1
if not errorlevel 1 set "PYRUN=py -3"
set "COUNT=0"
set "FAILED=0"
> "%LOGFILE%" echo Excel Skills workbook audit results
if errorlevel 1 (
  echo Cannot create log: "%LOGFILE%"
  exit /b 3
)
for /r "%SOURCE%" %%F in (*.xlsx) do (
  echo Checking "%%~fF"
  >> "%LOGFILE%" echo.
  %PYRUN% "%~dp0audit_xlsx.py" "%%~fF" >> "%LOGFILE%" 2>&1
  if errorlevel 1 set "FAILED=1"
  set /a COUNT+=1 >nul
)
echo.
echo Scanned Excel files: %COUNT%
if "%FAILED%"=="1" (
  echo Some files failed. Review "%LOGFILE%"
  exit /b 1
)
echo Report saved to "%LOGFILE%"
exit /b 0
