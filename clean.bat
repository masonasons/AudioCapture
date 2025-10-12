@echo off
REM AudioCapture Clean Script
REM This script removes all build artifacts

echo ========================================
echo AudioCapture Clean Script
echo ========================================
echo.

echo WARNING: This will delete the build directory and all build artifacts.
set /p confirm="Are you sure you want to continue? (Y/N): "

if /i not "%confirm%"=="Y" (
    echo Clean cancelled.
    pause
    exit /b 0
)

echo.
echo Cleaning build directory...

if exist "build" (
    rmdir /S /Q "build"
    echo Build directory deleted.
) else (
    echo Build directory does not exist.
)

echo.
echo ========================================
echo Clean complete!
echo ========================================
echo.

pause
