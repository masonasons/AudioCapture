@echo off
REM AudioCapture Build Script (CI/CD - No pauses)
REM This script builds the AudioCapture application for automated/CI builds

echo ========================================
echo AudioCapture Build Script (CI/CD)
echo ========================================
echo.

REM Check if build directory exists, if not create it
if not exist "build" (
    echo Creating build directory...
    mkdir build
    cd build

    echo Configuring CMake...
    "C:\Program Files\CMake\bin\cmake.exe" .. -G "Visual Studio 17 2022" -A x64 -DCMAKE_TOOLCHAIN_FILE=C:/vcpkg/scripts/buildsystems/vcpkg.cmake

    if errorlevel 1 (
        echo ERROR: CMake configuration failed!
        cd ..
        exit /b 1
    )
    cd ..
) else (
    echo Build directory already exists.
)

echo.
echo Building AudioCapture (Release)...
"C:\Program Files\CMake\bin\cmake.exe" --build build --config Release --target AudioCapture

if errorlevel 1 (
    echo.
    echo ERROR: Build failed!
    exit /b 1
)

echo.
echo ========================================
echo Build successful!
echo ========================================
echo.

REM Check if package directory exists, if not create it
if not exist "package" (
    echo Creating package directory...
    mkdir package
)

echo Copying executable to package folder...
copy /Y "build\bin\Release\AudioCapture.exe" "package\AudioCapture.exe"

if errorlevel 1 (
    echo WARNING: Failed to copy executable to package folder
    exit /b 1
) else (
    echo Successfully copied to package\AudioCapture.exe
)

echo.
echo Copying DLL dependencies to package folder...

REM Copy Opus DLL
if exist "C:\vcpkg\installed\x64-windows\bin\opus.dll" (
    copy /Y "C:\vcpkg\installed\x64-windows\bin\opus.dll" "package\opus.dll"
    echo Copied opus.dll
) else (
    echo WARNING: opus.dll not found
)

REM Copy Ogg DLL
if exist "C:\vcpkg\installed\x64-windows\bin\ogg.dll" (
    copy /Y "C:\vcpkg\installed\x64-windows\bin\ogg.dll" "package\ogg.dll"
    echo Copied ogg.dll
) else (
    echo WARNING: ogg.dll not found
)

echo.
echo ========================================
echo Build complete!
echo ========================================
echo Output: package\AudioCapture.exe
echo.

exit /b 0
