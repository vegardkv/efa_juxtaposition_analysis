@echo off
REM EFA Juxtaposition Analysis Advanced Launcher
REM Copyright (C) 2025 John-Are Hansen
REM Licensed under GPL v3.0

setlocal enabledelayedexpansion

echo ========================================
echo EFA Juxtaposition Analysis
echo Advanced Launcher v1.0
echo ========================================
echo.

REM Determine the application directory
REM First try the current directory, then the standard location
set CURRENT_DIR=%~dp0
set APP_DIR=%CURRENT_DIR%
set STANDARD_DIR=c:\Appl\efa_uv_app
set REQUIRED_LIBS=numpy pandas matplotlib scipy shapely

echo Batch file location: %CURRENT_DIR%
echo Looking for application files...

REM Function to print colored output (if supported)
set "RED=[91m"
set "GREEN=[92m"
set "YELLOW=[93m"
set "BLUE=[94m"
set "RESET=[0m"

echo %BLUE%Performing system checks...%RESET%
echo.

REM Check if uv is installed and get version
echo [1/7] Checking uv installation...
where uv >nul 2>&1
if %errorlevel% neq 0 (
    echo %RED%✗ ERROR: uv is not installed or not in PATH%RESET%
    echo.
    echo Installation instructions:
    echo 1. Open Command Prompt as Administrator
    echo 2. Run: winget install --id=astral-sh.uv -e
    echo 3. Restart Command Prompt and try again
    echo.
    pause
    exit /b 1
)

REM Get uv version
for /f "tokens=*" %%i in ('uv --version 2^>nul') do set UV_VERSION=%%i
echo %GREEN%✓ uv is installed: %UV_VERSION%%RESET%

REM Check if application directory exists
echo [2/7] Checking application directory...

REM Check if Python file exists in current directory (where batch file is located)
if exist "%CURRENT_DIR%EFA_juxtaposition_v0p9p6.py" (
    set APP_DIR=%CURRENT_DIR%
    echo %GREEN%✓ Application found in current directory: %APP_DIR%%RESET%
    goto :check_python_file
)

REM Check if Python file exists in standard directory
if exist "%STANDARD_DIR%\EFA_juxtaposition_v0p9p6.py" (
    set APP_DIR=%STANDARD_DIR%
    echo %GREEN%✓ Application found in standard directory: %APP_DIR%%RESET%
    goto :check_python_file
)

REM If neither location has the Python file, show error
echo %RED%✗ ERROR: Application directory not found%RESET%
echo.
echo Searched locations:
echo   1. Current directory: %CURRENT_DIR%
echo   2. Standard directory: %STANDARD_DIR%
echo.
echo Setup instructions:
echo 1. Ensure you're running this batch file from the same directory as EFA_juxtaposition_v0p9p6.py
echo    OR
echo 2. Set up the standard installation:
echo    - Navigate to c:\Appl (create if it doesn't exist)
echo    - Run: uv init efa_uv_app
echo    - Copy the Python application files to c:\Appl\efa_uv_app\
echo.
pause
exit /b 1

:check_python_file

REM Check if Python application file exists
echo [3/7] Checking Python application file...
if not exist "%APP_DIR%EFA_juxtaposition_v0p9p6.py" (
    echo %RED%✗ ERROR: Python application file not found%RESET%
    echo Expected: %APP_DIR%EFA_juxtaposition_v0p9p6.py
    echo.
    echo Please download the file from GitHub and copy it to the application directory
    echo.
    pause
    exit /b 1
)

REM Get file size for verification
for %%A in ("%APP_DIR%EFA_juxtaposition_v0p9p6.py") do set FILE_SIZE=%%~zA
echo %GREEN%✓ Python application file found (Size: %FILE_SIZE% bytes)%RESET%

REM Check if pyproject.toml exists
echo [4/7] Checking uv project configuration...
if not exist "%APP_DIR%pyproject.toml" (
    echo %YELLOW%⚠ Warning: uv project configuration not found%RESET%
    echo Expected: %APP_DIR%pyproject.toml
    echo.
    echo This suggests the application was not properly initialized with uv
    echo The application may still work if dependencies are installed globally
    echo.
    echo To create a proper uv project:
    echo   1. Navigate to the parent directory of your application
    echo   2. Run: uv init [project_name]
    echo   3. Copy the Python files to the new project directory
    echo   4. Run: uv add numpy pandas matplotlib scipy shapely
    echo.
    set UV_PROJECT=false
) else (
    echo %GREEN%✓ uv project configuration found%RESET%
    set UV_PROJECT=true
)

REM Check Python version
echo [5/7] Checking Python version...
cd /d "%APP_DIR%"

if "%UV_PROJECT%"=="true" (
    for /f "tokens=*" %%i in ('uv run python --version 2^>nul') do set PYTHON_VERSION=%%i
) else (
    for /f "tokens=*" %%i in ('python --version 2^>nul') do set PYTHON_VERSION=%%i
)

if "!PYTHON_VERSION!"=="" (
    echo %YELLOW%⚠ Warning: Could not determine Python version%RESET%
) else (
    echo %GREEN%✓ Python version: !PYTHON_VERSION!%RESET%
)

REM Check if required dependencies are installed with detailed reporting
echo [6/7] Checking Python dependencies...

REM Create a more detailed dependency check script
echo import sys > detailed_check.py
echo import importlib >> detailed_check.py
echo. >> detailed_check.py
echo required_libs = ['numpy', 'pandas', 'matplotlib', 'scipy', 'shapely'] >> detailed_check.py
echo missing_libs = [] >> detailed_check.py
echo installed_libs = [] >> detailed_check.py
echo. >> detailed_check.py
echo for lib in required_libs: >> detailed_check.py
echo     try: >> detailed_check.py
echo         module = importlib.import_module(lib) >> detailed_check.py
echo         version = getattr(module, '__version__', 'unknown') >> detailed_check.py
echo         installed_libs.append(f'{lib} (v{version})') >> detailed_check.py
echo     except ImportError: >> detailed_check.py
echo         missing_libs.append(lib) >> detailed_check.py
echo. >> detailed_check.py
echo if missing_libs: >> detailed_check.py
echo     print('MISSING:' + ','.join(missing_libs)) >> detailed_check.py
echo     sys.exit(1) >> detailed_check.py
echo else: >> detailed_check.py
echo     print('INSTALLED:' + '|'.join(installed_libs)) >> detailed_check.py

REM Run the detailed dependency check
if "%UV_PROJECT%"=="true" (
    for /f "tokens=1,2 delims=:" %%a in ('uv run detailed_check.py 2^>nul') do (
        if "%%a"=="MISSING" (
            echo %RED%✗ ERROR: Missing required libraries: %%b%RESET%
            echo.
            echo Please install missing dependencies using:
            echo   uv add %%b
            echo.
            del detailed_check.py >nul 2>&1
            pause
            exit /b 1
        )
        if "%%a"=="INSTALLED" (
            echo %GREEN%✓ All required libraries are installed:%RESET%
            for %%c in (%%b) do (
                set lib=%%c
                set lib=!lib:|= !
                echo   - !lib!
            )
        )
    )
) else (
    for /f "tokens=1,2 delims=:" %%a in ('python detailed_check.py 2^>nul') do (
        if "%%a"=="MISSING" (
            echo %RED%✗ ERROR: Missing required libraries: %%b%RESET%
            echo.
            echo Please install missing dependencies using:
            echo   pip install %%b
            echo   OR set up a proper uv project and use: uv add %%b
            echo.
            del detailed_check.py >nul 2>&1
            pause
            exit /b 1
        )
        if "%%a"=="INSTALLED" (
            echo %GREEN%✓ All required libraries are installed:%RESET%
            for %%c in (%%b) do (
                set lib=%%c
                set lib=!lib:|= !
                echo   - !lib!
            )
        )
    )
)
del detailed_check.py >nul 2>&1

REM Final pre-launch check
echo [7/7] Performing final checks...
echo %GREEN%✓ All system checks passed%RESET%
echo.

REM Launch the application
echo %BLUE%========================================%RESET%
echo %BLUE%Launching EFA Juxtaposition Analysis...%RESET%
echo %BLUE%========================================%RESET%
echo.
echo %YELLOW%Starting application... (this may take a few moments)%RESET%
echo %YELLOW%To close the application, close the GUI window or press Ctrl+C%RESET%
echo.

REM Record start time
set START_TIME=%TIME%

if "%UV_PROJECT%"=="true" (
    uv run EFA_juxtaposition_v0p9p6.py
) else (
    python EFA_juxtaposition_v0p9p6.py
)

REM Check exit status
set EXIT_CODE=%errorlevel%
set END_TIME=%TIME%

echo.
echo %BLUE%========================================%RESET%
if %EXIT_CODE% equ 0 (
    echo %GREEN%✓ Application closed successfully%RESET%
    echo Session started: %START_TIME%
    echo Session ended: %END_TIME%
) else (
    echo %RED%✗ Application exited with error code: %EXIT_CODE%%RESET%
    echo.
    echo Common issues and solutions:
    echo - If GUI doesn't appear: Check if tkinter is installed
    echo - If import errors: Reinstall dependencies with uv add
    echo - If file errors: Verify application file integrity
    echo.
    echo For support, contact: jareh@equinor.com
)

echo %BLUE%========================================%RESET%
echo.
pause