@echo off
REM ========================================
REM Build Script for Online Logger v2.0
REM Creates an executable using PyInstaller
REM ========================================

echo.
echo ========================================
echo   Online Logger - Build Script
echo   Version: 2.0
echo ========================================
echo.

REM Check if PyInstaller is installed
echo [1/5] Checking PyInstaller installation...
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller is not installed!
    echo.
    echo Installing PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to install PyInstaller!
        echo Please install it manually: pip install pyinstaller
        pause
        exit /b 1
    )
)
echo PyInstaller is installed.
echo.

REM Clean previous build artifacts
echo [2/5] Cleaning previous build artifacts...
if exist "dist\Online_Log_SQL_Lite" (
    rmdir /s /q "dist\Online_Log_SQL_Lite"
    echo Removed old dist folder.
)
if exist "build" (
    rmdir /s /q "build"
    echo Removed old build folder.
)
if exist "Online_Log_SQL_Lite.spec" (
    del /q "Online_Log_SQL_Lite.spec"
    echo Removed old spec file.
)
echo.

REM Build the executable
echo [3/5] Building executable with PyInstaller...
echo.
echo Configuration:
echo - Name: Online Logger v2.0
echo - Icon: _repositoryfiles\OnlineLoggerLogo.ico
echo - Console: Enabled (for debugging)
echo - One Directory Mode
echo.

pyinstaller ^
    --name="Online_Logger_v2.0" ^
    --icon="_repositoryfiles\OnlineLoggerLogo.ico" ^
    --console ^
    --onedir ^
    --noconfirm ^
    --add-data "settings;settings" ^
    --add-data "_repositoryfiles;_repositoryfiles" ^
    --hidden-import=xlwings ^
    --hidden-import=customtkinter ^
    --hidden-import=PIL ^
    --hidden-import=watchdog ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    "Python Script\Online_Log_SQL_Lite.py"

if errorlevel 1 (
    echo.
    echo ERROR: Build failed!
    pause
    exit /b 1
)
echo.

REM Copy additional files to dist folder
echo [4/5] Copying additional files to distribution folder...

REM Copy settings folder with all JSON files and config subfolder
if exist "settings" (
    xcopy /E /I /Y "settings" "dist\Online_Logger_v2.0\settings" >nul
    echo Copied settings folder with project JSONs and config.
)

REM Copy Guide folder
if exist "Guide" (
    xcopy /E /I /Y "Guide" "dist\Online_Logger_v2.0\Guide" >nul
    echo Copied Guide folder.
)

REM Copy README and launcher
if exist "README.md" (
    copy /Y "README.md" "dist\Online_Logger_v2.0\" >nul
    echo Copied README.md
)

if exist "BUILD_README.md" (
    copy /Y "BUILD_README.md" "dist\Online_Logger_v2.0\" >nul
    echo Copied BUILD_README.md
)

if exist "Start_Online_Logger.bat" (
    copy /Y "Start_Online_Logger.bat" "dist\Online_Logger_v2.0\" >nul
    echo Copied launcher script
)


echo.

REM Show build results
echo [5/5] Build complete!
echo.
echo ========================================
echo   Build Summary
echo ========================================
echo.
echo Executable location:
echo   dist\Online_Logger_v2.0\Online_Logger_v2.0.exe
echo.
echo Console window: ENABLED (for debugging)
echo.
echo To disable console for production:
echo   Edit this script and remove --console flag
echo.

REM Ask if user wants to open the dist folder
echo.
choice /C YN /M "Open dist folder now"
if errorlevel 2 goto :end
if errorlevel 1 explorer "dist\Online_Logger_v2.0"

:end
echo.
echo Press any key to exit...
pause >nul
