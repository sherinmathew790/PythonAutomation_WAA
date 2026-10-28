@echo off
setlocal enabledelayedexpansion

echo ===================================
echo Cleaning Log Files
echo ===================================
echo.

REM Check if logs directory exists
if not exist "logs" (
    echo No logs directory found.
    echo Creating logs directory...
    mkdir logs
    goto :end
)

echo Checking for log files...

REM Count existing log files
set "log_count=0"
for %%F in (logs\*.log) do set /a "log_count+=1"

if %log_count% equ 0 (
    echo No log files found to clean.
    goto :end
)

echo Found %log_count% log files.
choice /C YN /M "Do you want to delete all log files"

if !errorlevel! equ 2 (
    echo Operation cancelled by user.
    goto :end
)

REM Delete all log files
echo Deleting log files...
del /f /q "logs\*.log" 2>nul
if !errorlevel! equ 0 (
    echo Successfully deleted all log files.
) else (
    echo Warning: Some files could not be deleted.
)

:end
echo.
echo Cleanup process completed.
echo.
pause
endlocal