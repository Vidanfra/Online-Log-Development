@echo off
REM Concurrent Access Test Runner for Online Logger & Field Log Viewer
REM This script tests if both programs can access the database simultaneously

echo ========================================================================
echo SQLite Concurrent Access Test
echo ========================================================================
echo.
echo This test verifies that Online Logger and Field Log Viewer can
echo work together without blocking or freezing issues.
echo.
echo Prerequisites:
echo   1. Online Logger has been run at least once
echo   2. SQLite Mirror has been enabled
echo   3. At least one log entry exists in the database
echo.
echo ========================================================================
echo.

REM Try to find Python
set PYTHON_CMD=python
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python not found in PATH
    echo.
    echo Please install Python 3.13 or add it to your PATH
    echo Python download: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Searching for database files...
echo.

REM Check common database locations
set DB_PATH=
if exist "SQL Database\CNO_600061_Online_Log_20251029.db" (
    set DB_PATH=SQL Database\CNO_600061_Online_Log_20251029.db
    echo Found: %DB_PATH%
) else if exist "SQL Database\Online_Log_SQLite.db" (
    set DB_PATH=SQL Database\Online_Log_SQLite.db
    echo Found: %DB_PATH%
) else if exist "SQL Database\fieldlog.db" (
    set DB_PATH=SQL Database\fieldlog.db
    echo Found: %DB_PATH%
) else (
    echo WARNING: No database found in common locations
    echo.
    echo Please ensure:
    echo   1. Online Logger has been run
    echo   2. SQLite Mirror is enabled
    echo   3. At least one log entry exists
    echo.
    echo Or manually specify database path:
    echo   test_concurrent_access.bat "path\to\your\database.db"
    echo.
    pause
    exit /b 1
)

echo.
echo Running concurrent access test...
echo ========================================================================
echo.

REM Run the test
if "%~1"=="" (
    python test_concurrent_access.py "%DB_PATH%"
) else (
    python test_concurrent_access.py "%~1"
)

echo.
echo ========================================================================
echo Test completed!
echo.
echo If all tests PASSED (green), the scripts are safe to deploy.
echo If any tests FAILED (red), review TESTING_CONCURRENT_ACCESS.md
echo ========================================================================
echo.
pause
