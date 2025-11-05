@echo off
REM ============================================================================
REM Build Script for Field Log Viewer v2
REM ============================================================================
REM This script creates a standalone executable for Field Log Viewer
REM including the settings folder with default configuration.
REM ============================================================================

echo.
echo ========================================================================
echo Building Field Log Viewer v2 Executable
echo ========================================================================
echo.

REM --- Version Configuration ---
SET VERSION=2.0
SET APP_NAME=Field_Log_Viewer_v%VERSION%

echo Version: %VERSION%
echo Output: %APP_NAME%.exe
echo.

REM --- Check if PyInstaller is installed ---
echo Checking for PyInstaller...
python -m PyInstaller --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: PyInstaller is not installed.
    echo.
    echo Installing PyInstaller...
    pip install pyinstaller
    IF %ERRORLEVEL% NEQ 0 (
        echo ERROR: Failed to install PyInstaller.
        echo Please install manually: pip install pyinstaller
        pause
        exit /b 1
    )
)
echo PyInstaller found!
echo.

REM --- Clean previous build artifacts ---
echo Cleaning previous build artifacts...
echo.
echo IMPORTANT: Close Field Log Viewer if it's running!
timeout /t 3 /nobreak >nul
IF EXIST "build" rmdir /s /q "build"
IF EXIST "dist\%APP_NAME%.exe" (
    echo Removing old executable...
    del /f /q "dist\%APP_NAME%.exe" 2>nul
    IF EXIST "dist\%APP_NAME%.exe" (
        echo.
        echo ERROR: Cannot delete old executable - it may be running!
        echo Please close Field_Log_Viewer_v2.0.exe and try again.
        pause
        exit /b 1
    )
)
IF EXIST "dist" rmdir /s /q "dist"
IF EXIST "%APP_NAME%.spec" del /q "%APP_NAME%.spec"
echo Clean complete.
echo.

REM --- Create settings folder structure if it doesn't exist ---
echo Checking settings folder...
IF NOT EXIST "flv_settings" (
    echo Creating default settings folder...
    mkdir "flv_settings"
    echo Default settings folder created.
)
echo.

REM --- Build the executable ---
echo ========================================================================
echo Building executable with PyInstaller...
echo ========================================================================
echo.
echo Options:
echo   - One file executable
echo   - Windowed (no console)
echo   - Including settings folder
echo.

echo.
echo NOTE: The next step can take 2-5 minutes. Please be patient!
echo       - Building PYZ archive: ~30 seconds
echo       - Building PKG archive: ~2-3 minutes (this is the slow part)
echo       - Building EXE: ~1 minute
echo.
echo If it appears stuck at "Building PKG", it's still working...
echo.

python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "%APP_NAME%" ^
    --add-data "flv_settings;flv_settings" ^
    --exclude-module matplotlib ^
    --exclude-module matplotlib.pyplot ^
    --exclude-module PIL ^
    --exclude-module plotly ^
    --exclude-module jinja2 ^
    --exclude-module pytest ^
    --exclude-module IPython ^
    --exclude-module notebook ^
    --exclude-module scipy ^
    --exclude-module sklearn ^
    --exclude-module pandas.tests ^
    --exclude-module pandas.plotting ^
    --exclude-module pandas.io.parquet ^
    --exclude-module pandas.io.feather ^
    "Field Log Viewer v2.py"

IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo ========================================================================
    echo ERROR: Build failed!
    echo ========================================================================
    echo.
    echo Check the error messages above for details.
    pause
    exit /b 1
)

echo.
echo ========================================================================
echo Build completed successfully!
echo ========================================================================
echo.

REM --- Check if executable was created ---
IF EXIST "dist\%APP_NAME%.exe" (
    echo Output location: dist\%APP_NAME%.exe
    echo File size: 
    dir "dist\%APP_NAME%.exe" | find "%APP_NAME%.exe"
    echo.
    echo ========================================================================
    echo Deployment Instructions:
    echo ========================================================================
    echo.
    echo 1. Copy the entire "dist" folder contents to deployment location
    echo 2. The executable includes the settings folder structure
    echo 3. Settings will be saved in the same folder as the executable
    echo.
    echo Note: The executable is portable and can be moved to any location.
    echo       Settings will auto-save in: [exe_location]\flv_settings\
    echo.
    echo ========================================================================
    echo.
    
    REM --- Optional: Create a clean deployment folder ---
    echo Creating deployment package...
    IF EXIST "%APP_NAME%" rmdir /s /q "%APP_NAME%"
    mkdir "%APP_NAME%"
    copy "dist\%APP_NAME%.exe" "%APP_NAME%\"
    
    IF EXIST "flv_settings" (
        xcopy "flv_settings" "%APP_NAME%\flv_settings\" /E /I /Y >nul
    )
    
    echo.
    echo Deployment package created in: %APP_NAME%\
    echo   - %APP_NAME%.exe
    echo   - flv_settings\ (folder with defaults)
    echo.
    echo This folder is ready to zip and send to colleagues!
    echo.
    
) ELSE (
    echo ERROR: Executable not found in dist folder!
    echo Build may have failed silently.
)

echo ========================================================================
echo Build process complete!
echo ========================================================================
echo.
pause
