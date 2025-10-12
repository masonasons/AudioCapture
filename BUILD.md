# AudioCapture Build Instructions

This document describes how to build the AudioCapture application.

## Prerequisites

Before building, ensure you have the following installed:

1. **Visual Studio 2022** (or Visual Studio 2019)
   - Install the "Desktop development with C++" workload
   - Download from: https://visualstudio.microsoft.com/

2. **CMake** (version 3.15 or later)
   - Download from: https://cmake.org/download/
   - Make sure it's installed at `C:\Program Files\CMake\bin\cmake.exe`
   - Or update the path in the build scripts

3. **vcpkg** (C++ package manager)
   - Clone from: https://github.com/Microsoft/vcpkg
   - Recommended location: `C:\vcpkg`
   - Bootstrap vcpkg: `.\vcpkg\bootstrap-vcpkg.bat`

4. **Required packages via vcpkg**:
   ```batch
   vcpkg install libflac:x64-windows
   vcpkg install opus:x64-windows
   vcpkg install ogg:x64-windows
   vcpkg install nlohmann-json:x64-windows
   ```

## Build Methods

### Method 1: Using the Build Script (Recommended)

Simply double-click **`build.bat`**

The script will:
1. Create the build directory if needed
2. Configure CMake with vcpkg
3. Build the Release configuration
4. Copy the executable to the `package` folder
5. Copy required DLLs (opus.dll, ogg.dll, FLAC.dll, FLAC++.dll) to the `package` folder

### Method 2: Manual Build

1. **Create build directory**:
   ```batch
   mkdir build
   cd build
   ```

2. **Configure CMake**:
   ```batch
   "C:\Program Files\CMake\bin\cmake.exe" .. -G "Visual Studio 17 2022" -A x64 -DCMAKE_TOOLCHAIN_FILE=C:/vcpkg/scripts/buildsystems/vcpkg.cmake
   ```

3. **Build the project**:
   ```batch
   "C:\Program Files\CMake\bin\cmake.exe" --build . --config Release --target AudioCapture
   ```

4. **Copy files to package**:
   ```batch
   copy bin\Release\AudioCapture.exe ..\package\AudioCapture.exe
   copy C:\vcpkg\installed\x64-windows\bin\opus.dll ..\package\opus.dll
   copy C:\vcpkg\installed\x64-windows\bin\ogg.dll ..\package\ogg.dll
   copy C:\vcpkg\installed\x64-windows\bin\FLAC.dll ..\package\FLAC.dll
   copy C:\vcpkg\installed\x64-windows\bin\FLAC++.dll ..\package\FLAC++.dll
   ```

## Output

After a successful build, you'll find:

- **`package\AudioCapture.exe`** - The main application
- **`package\opus.dll`** - Opus codec library
- **`package\ogg.dll`** - Ogg container library
- **`package\FLAC.dll`** - FLAC codec library
- **`package\FLAC++.dll`** - FLAC C++ wrapper library

You can run the application directly from the `package` folder.

## Creating Distribution Package

After building, you can create a distributable package:

1. Double-click **`package.bat`**

This will:
- Create a `dist\AudioCapture-v1.0.0` folder with all necessary files
- Include the executable, DLLs, and documentation
- Create a ZIP archive for easy distribution
- Generate a README.txt with usage instructions

The package includes:
- AudioCapture.exe
- opus.dll, ogg.dll, FLAC.dll, and FLAC++.dll
- README.txt (user instructions)
- LICENSE.txt (if available)
- AudioCaptures\ (default output folder)

## Cleaning Build Artifacts

To remove all build artifacts and start fresh:

- Double-click **`clean.bat`** to delete the build directory

## Troubleshooting

### CMake not found
- Verify CMake is installed at `C:\Program Files\CMake\bin\cmake.exe`
- Or update the path in the build scripts to match your installation

### vcpkg not found
- Verify vcpkg is installed at `C:\vcpkg`
- Or update `CMAKE_TOOLCHAIN_FILE` path in the build scripts
- Make sure you've run `bootstrap-vcpkg.bat` in the vcpkg directory

### Missing packages
- Ensure all required packages are installed via vcpkg:
  ```batch
  vcpkg list
  ```
- If missing, install them:
  ```batch
  vcpkg install libflac:x64-windows opus:x64-windows ogg:x64-windows nlohmann-json:x64-windows
  ```

### Build errors
- Make sure you have Visual Studio 2022 with C++ development tools
- Try cleaning and rebuilding: run `clean.bat` then `build.bat`
- Check that all prerequisites are installed correctly

## Configuration Options

### Changing vcpkg location
If vcpkg is installed in a different location, update the following in **build.bat** and **build-ci.bat**:

```batch
-DCMAKE_TOOLCHAIN_FILE=C:/your/path/to/vcpkg/scripts/buildsystems/vcpkg.cmake
```

And in **package.bat** update the DLL paths:
```batch
C:\your\path\to\vcpkg\installed\x64-windows\bin\opus.dll
C:\your\path\to\vcpkg\installed\x64-windows\bin\ogg.dll
C:\your\path\to\vcpkg\installed\x64-windows\bin\FLAC.dll
C:\your\path\to\vcpkg\installed\x64-windows\bin\FLAC++.dll
```

### Building Debug version

To build a debug version with debugging symbols, modify the build command:

```batch
"C:\Program Files\CMake\bin\cmake.exe" --build build --config Debug --target AudioCapture
```

The debug executable will be at `build\bin\Debug\AudioCapture.exe`

## System Requirements

- **OS**: Windows 10 Build 19045 or later (tested and working)
  - Build 20348+ recommended for optimal per-process audio capture
- **Architecture**: x64 (64-bit)
- **Compiler**: MSVC 2019 or later
