@echo off
REM AudioCapture Build Script
REM This script builds the AudioCapture application and copies it to the package folder

echo ========================================
echo AudioCapture Build Script
echo ========================================
echo.

REM Check if build directory exists and has CMakeCache.txt
if not exist "build\CMakeCache.txt" (
    echo Creating/Configuring build directory...
    if not exist "build" mkdir build
    cd build

    echo Configuring CMake...
    "C:\Program Files\CMake\bin\cmake.exe" .. -G "Visual Studio 17 2022" -A x64 -DCMAKE_TOOLCHAIN_FILE=C:/vcpkg/scripts/buildsystems/vcpkg.cmake

    if errorlevel 1 (
        echo ERROR: CMake configuration failed!
        cd ..
        pause
        exit /b 1
    )
    cd ..
) else (
    echo Build directory already configured.
)

echo.
echo Building AudioCapture (Release)...
"C:\Program Files\CMake\bin\cmake.exe" --build build --config Release --target AudioCapture

if errorlevel 1 (
    echo.
    echo ERROR: Build failed!
    pause
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

REM Copy FLAC DLLs
if exist "C:\vcpkg\installed\x64-windows\bin\FLAC.dll" (
    copy /Y "C:\vcpkg\installed\x64-windows\bin\FLAC.dll" "package\FLAC.dll"
    echo Copied FLAC.dll
) else (
    echo WARNING: FLAC.dll not found
)

if exist "C:\vcpkg\installed\x64-windows\bin\FLAC++.dll" (
    copy /Y "C:\vcpkg\installed\x64-windows\bin\FLAC++.dll" "package\FLAC++.dll"
    echo Copied FLAC++.dll
) else (
    echo WARNING: FLAC++.dll not found
)

echo.
echo ========================================
echo Build complete!
echo ========================================
echo Output: package\AudioCapture.exe
echo.

pause
