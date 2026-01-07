@echo off
REM ============================================================================
REM EFA Juxtaposition Analysis - PyInstaller Build Script
REM ============================================================================
REM This script compiles EFA_juxtaposition_v0p9p9.py using PyInstaller
REM Creates a folder containing the executable and all dependencies
REM ============================================================================

echo ============================================================================
echo EFA Juxtaposition Analysis - PyInstaller Build
echo ============================================================================
echo.

REM Activate virtual environment
echo [1/5] Activating virtual environment...
call Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment!
    pause
    exit /b 1
)
echo Done.
echo.

REM Check if PyInstaller is installed
echo [2/5] Checking PyInstaller installation...
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo PyInstaller not found. Installing...
    pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: Failed to install PyInstaller!
        pause
        exit /b 1
    )
) else (
    echo PyInstaller is already installed.
)
echo.

REM Clean previous builds
echo [3/5] Cleaning previous builds...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "EFA_juxtaposition_v0p9p9.spec" del "EFA_juxtaposition_v0p9p9.spec"
echo Done.
echo.

REM Build with PyInstaller
echo [4/5] Building with PyInstaller (--onedir)...
echo This may take several minutes...
echo.

pyinstaller --onedir ^
    --name="EFA_Juxtaposition_v1.0.0" ^
    --windowed ^
    --icon=help_images\efa_icon.ico ^
    --add-data "test_data;test_data" ^
    --add-data "help_images;help_images" ^
    --hidden-import=tkinter ^
    --hidden-import=matplotlib ^
    --hidden-import=numpy ^
    --hidden-import=pandas ^
    --hidden-import=scipy ^
    --hidden-import=shapely ^
    --hidden-import=PIL ^
    --hidden-import=openpyxl ^
    --hidden-import=xlsxwriter ^
    --hidden-import=matplotlib.backends.backend_tkagg ^
    --collect-all=matplotlib ^
    --collect-all=shapely ^
    --noupx ^
    EFA_juxtaposition_v0p9p9.py

if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed!
    echo Check the output above for error details.
    pause
    exit /b 1
)
echo.
echo Build completed successfully!
echo.

REM Copy test_data if build succeeded
echo [5/5] Copying additional files to distribution folder...
if exist "test_data" (
    xcopy "test_data" "dist\EFA_Juxtaposition_v1.0.0\test_data\" /E /I /Y >nul
    echo Test data copied successfully.
) else (
    echo WARNING: test_data folder not found. Skipping...
)

if exist "help_images" (
    xcopy "help_images" "dist\EFA_Juxtaposition_v1.0.0\help_images\" /E /I /Y >nul
    echo Help images copied successfully.
) else (
    echo WARNING: help_images folder not found. Skipping...
)
echo.

echo ============================================================================
echo BUILD COMPLETE!
echo ============================================================================
echo.
echo Output location: dist\EFA_Juxtaposition_v1.0.0\
echo.
echo Directory contents:
dir "dist\EFA_Juxtaposition_v1.0.0" /B
echo.
echo Executable: dist\EFA_Juxtaposition_v1.0.0\EFA_Juxtaposition_v1.0.0.exe
echo.
echo TO RUN:
echo   1. Navigate to dist\EFA_Juxtaposition_v1.0.0\
echo   2. Double-click EFA_Juxtaposition_v1.0.0.exe
echo.
echo TO DISTRIBUTE:
echo   1. Zip the entire dist\EFA_Juxtaposition_v1.0.0\ folder
echo   2. Share the ZIP file
echo   3. Users extract and run the .exe file
echo.
echo ============================================================================
pause
