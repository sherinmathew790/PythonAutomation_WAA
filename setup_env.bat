@echo off
setlocal enabledelayedexpansion

echo ===================================
echo Starting project setup...
echo ===================================

REM Define minimum required Python version
set MIN_PYTHON_VERSION=3.12.4

REM Check for Python command and version
set "PYTHON_CMD="
for %%C in (python py) do (
    %%C --version >nul 2>&1
    if !errorlevel! equ 0 (
        for /f "tokens=1,2" %%V in ('%%C --version 2^>^&1') do (
            set "PYTHON_VERSION=%%W"
            call :compare_versions "!PYTHON_VERSION!" "%MIN_PYTHON_VERSION%"
            if !errorlevel! equ 1 (
                echo Found Python !PYTHON_VERSION! using '%%C' but minimum required version is %MIN_PYTHON_VERSION%
            ) else (
                set "PYTHON_CMD=%%C"
                echo Found Python !PYTHON_VERSION! using '%%C'
                goto :python_found
            )
        )
    )
)

:python_not_found
echo.
echo Error: Compatible Python version not found.
echo Please ensure that:
echo  1. Python is installed and added to PATH
echo  2. Python version is %MIN_PYTHON_VERSION% or higher
echo  3. Using either 'python' or 'py' command
echo.
echo Current status:
where python 2>nul && python --version
where py 2>nul && py --version
pause
exit /b 1

:compare_versions
setlocal
set "ver1=%~1"
set "ver2=%~2"
for /f "tokens=1,2,3 delims=." %%a in ("%ver1%") do (
    set "major1=%%a"
    set "minor1=%%b"
    set "patch1=%%c"
)
for /f "tokens=1,2,3 delims=." %%a in ("%ver2%") do (
    set "major2=%%a"
    set "minor2=%%b"
    set "patch2=%%c"
)
if %major1% lss %major2% exit /b 1
if %major1% gtr %major2% exit /b 0
if %minor1% lss %minor2% exit /b 1
if %minor1% gtr %minor2% exit /b 0
if %patch1% lss %patch2% exit /b 1
exit /b 0

:python_found
echo Using Python command: !PYTHON_CMD!
echo Python version: !PYTHON_VERSION!

REM Check if virtual environment already exists
if exist venv (
    echo.
    echo Virtual environment already exists.
    choice /C YN /M "Do you want to remove and recreate it"
    if !errorlevel! equ 1 (
        echo.
        echo Removing existing virtual environment...
        rmdir /s /q venv
        if !errorlevel! neq 0 (
            echo Error: Failed to remove existing virtual environment.
            pause
            exit /b 1
        )
    ) else (
        echo.
        echo Setup cancelled by user.
        pause
        exit /b 0
    )
)

echo.
echo Creating virtual environment...
!PYTHON_CMD! -m venv venv
if !errorlevel! neq 0 (
    echo Error: Failed to create virtual environment.
    echo Please ensure you have venv module installed.
    pause
    exit /b 1
)

echo.
echo Activating virtual environment...
call venv\Scripts\activate
if !errorlevel! neq 0 (
    echo Error: Failed to activate virtual environment.
    pause
    exit /b 1
)

echo.
echo Upgrading pip...
!PYTHON_CMD! -m pip install --upgrade pip
if !errorlevel! neq 0 (
    echo Warning: Failed to upgrade pip, continuing with installation...
)

echo.
echo Installing dependencies...
pip install -r requirements.txt
if !errorlevel! neq 0 (
    echo Error: Failed to install dependencies.
    echo Please check requirements.txt and your internet connection.
    pause
    exit /b 1
)

echo.
echo ===================================
echo Setup completed successfully!
echo ===================================
echo.
echo Virtual environment created and dependencies installed.
echo Python version: !PYTHON_VERSION!
echo.
echo To activate the virtual environment:
echo     venv\Scripts\activate
echo.
echo To deactivate the virtual environment:
echo     deactivate
echo.
echo To run tests:
echo     run_behave.bat
echo ===================================

REM Keep the window open
pause

endlocal