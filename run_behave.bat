@echo off
setlocal enabledelayedexpansion
 
REM Get the current directory of the script
set SCRIPT_DIR=%~dp0
 
REM Set PYTHONPATH relative to the script directory
set PYTHONPATH=%SCRIPT_DIR%..\innoproj-general
 
echo Starting Behave tests...
echo Script directory: %SCRIPT_DIR%
echo PYTHONPATH: %PYTHONPATH%
 
REM Activate virtual environment (force activation regardless of current state)
echo Activating virtual environment...
if exist "%SCRIPT_DIR%venv\Scripts\activate.bat" (
    call "%SCRIPT_DIR%venv\Scripts\activate.bat"
) else (
    echo Error: Virtual environment not found at: %SCRIPT_DIR%venv
    echo Please run setup.bat first to create the virtual environment.
    pause
    exit /b 1
)
 
REM Verify Python is from virtual environment
for /f "tokens=*" %%i in ('where python') do (
    set PYTHON_PATH=%%i
    goto :found_python
)
:found_python
 
echo Using Python: %PYTHON_PATH%
python --version
 
REM Check if tests directory exists
if not exist "%SCRIPT_DIR%tests" (
    echo Error: Tests directory not found at: %SCRIPT_DIR%tests
    pause
    exit /b 1
)
 
REM Change to tests directory
echo Changing to tests directory...
cd "%SCRIPT_DIR%tests"
if !errorlevel! neq 0 (
    echo Error: Failed to change to tests directory.
    pause
    exit /b 1
)
 
REM Verify behave is available
where behave >nul 2>&1
if !errorlevel! neq 0 (
    echo Error: Behave not found. Please ensure it's installed in the virtual environment.
    echo Try running: pip install behave
    pause
    exit /b 1
)
 
REM Run Behave tests
echo.
echo Running Behave tests...
echo =====================================
behave --no-capture --format pretty
set BEHAVE_EXIT_CODE=!errorlevel!
 
REM Check Behave execution status
if !BEHAVE_EXIT_CODE! neq 0 (
    echo.
    echo Tests completed with errors. Exit code: !BEHAVE_EXIT_CODE!
    echo Please check the test results above.
) else (
    echo.
    echo All tests completed successfully!
)
 
REM Return to original directory
cd "%SCRIPT_DIR%"
 
echo.
echo Test execution completed.
echo Results are available in the logs folder.
pause
 
exit /b !BEHAVE_EXIT_CODE!
 
endlocal
 