@echo off

REM Try to find cmake - prefer hardcoded path, fall back to PATH
set CMAKE_EXE="C:\Program Files\CMake\bin\cmake.exe"
if not exist %CMAKE_EXE% (
	echo CMake not found at hardcoded path, trying PATH...
	set CMAKE_EXE=cmake
)

REM Determine vcpkg toolchain path
if defined VCPKG_ROOT (
	set "VCPKG_TOOLCHAIN=%VCPKG_ROOT%/scripts/buildsystems/vcpkg.cmake"
) else (
	set "VCPKG_TOOLCHAIN=C:/vcpkg/scripts/buildsystems/vcpkg.cmake"
)

if not exist "build" (
	mkdir build
	cd build
	%CMAKE_EXE% .. -G "Visual Studio 17 2022" -A x64 -DCMAKE_TOOLCHAIN_FILE=%VCPKG_TOOLCHAIN% -DVCPKG_TARGET_TRIPLET=x64-windows-static-mt
	if errorlevel 1 (
		echo ERROR: CMake configuration failed
		cd ..
		exit /b 1
	)
	cd ..
)

%CMAKE_EXE% --build build --config Release --target AudioCapture
if errorlevel 1 (
	echo ERROR: Build failed
	exit /b 1
)

if not exist "package" mkdir package

copy /Y "build\bin\Release\AudioCapture.exe" "package\AudioCapture.exe" >nul
if errorlevel 1 (
	echo ERROR: Failed to copy executable
	exit /b 1
)

echo Libraries are statically linked - no DLLs needed.

exit /b 0
