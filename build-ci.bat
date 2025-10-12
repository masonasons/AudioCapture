@echo off
REM CI Build Script - follows build.bat configuration

if not exist "build" (
	mkdir build
	cd build
	"C:\Program Files\CMake\bin\cmake.exe" .. -G "Visual Studio 17 2022" -A x64 -DCMAKE_TOOLCHAIN_FILE=C:/vcpkg/scripts/buildsystems/vcpkg.cmake -DVCPKG_TARGET_TRIPLET=x64-windows-static-mt
	if errorlevel 1 (
		echo ERROR: CMake configuration failed
		cd ..
		exit /b 1
	)
	cd ..
)

"C:\Program Files\CMake\bin\cmake.exe" --build build --config Release --target AudioCapture
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
