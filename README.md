# Audio Capture - Per-Process Audio Recording

A Windows application for capturing audio from individual processes using WASAPI (Windows Audio Session API). Records audio to WAV, MP3, Opus, or FLAC formats with an accessible Win32 interface.

## Full Disclosure

This was written with Claude Code. It has been tested and does function completely.

## Features

- **Per-Process Audio Capture**: Record audio from specific applications independently
- **Multiple Format Support**: Save recordings as WAV, MP3, Opus, or FLAC files
  - **WAV**: Uncompressed PCM audio (highest quality, largest size)
  - **MP3**: Compressed with configurable bitrate (128-320 kbps)
  - **Opus**: Modern codec with configurable bitrate (64-256 kbps)
  - **FLAC**: Lossless compression with configurable levels (0-8)
- **Multi-Process Recording**: Capture audio from multiple processes simultaneously using checkboxes
- **Silence Detection**: Optional skip silence feature to save disk space
- **Process Filtering**: Show only processes with active audio output
- **Window Title Display**: See window titles to easily identify processes
- **Accessible Win32 UI**: Standard Windows controls with full keyboard navigation and screen reader support
- **Real-time Monitoring**: View active recording sessions and data statistics
- **Settings Persistence**: Automatically saves and restores your preferences

## Requirements

### System Requirements
- Windows 10 Build 19045 or later (tested and working)
- Windows 10 Build 20348 or later (for optimal per-process capture)

### Build Requirements
- Visual Studio 2019 or later (with C++ Desktop Development workload)
- Windows SDK 10.0.19041.0 or later
- C++17 compatible compiler
- CMake 3.15 or later
- vcpkg (for managing dependencies)

### Dependencies (via vcpkg)
- libflac
- opus
- ogg
- nlohmann-json

## Building the Project

### Prerequisites

1. Install vcpkg dependencies:
```cmd
vcpkg install libflac:x64-windows opus:x64-windows ogg:x64-windows nlohmann-json:x64-windows
```

### Using the Build Script

The easiest way to build is using the provided batch script:

```cmd
build.bat
```

This will:
- Configure CMake with vcpkg integration
- Build the Release configuration
- Copy the executable and DLLs to the `package` folder

### Manual Build with Visual Studio

1. Open a Developer Command Prompt for Visual Studio

2. Navigate to the project directory and create a build directory:
```cmd
mkdir build
cd build
```

3. Generate Visual Studio project files with vcpkg:
```cmd
cmake .. -G "Visual Studio 17 2022" -A x64 -DCMAKE_TOOLCHAIN_FILE=C:/vcpkg/scripts/buildsystems/vcpkg.cmake
```

4. Build the project:
```cmd
cmake --build . --config Release
```

The executable will be in `build\bin\Release\AudioCapture.exe`

### Using CMake and Ninja

1. Open a Developer Command Prompt for Visual Studio

2. Navigate to the project directory and create a build directory:
```cmd
mkdir build
cd build
```

3. Configure and build:
```cmd
cmake .. -G Ninja -DCMAKE_BUILD_TYPE=Release
ninja
```

The executable will be in `build\bin\AudioCapture.exe`

## Usage

### Starting the Application

1. Run `AudioCapture.exe`
2. The main window will display a list of running processes with their window titles

### Basic Capture (Single Process)

1. **Navigate to a Process**: Use arrow keys or click on a process in the list
2. **Choose Output Format**: Select WAV, MP3, Opus, or FLAC from the dropdown
   - For MP3/Opus: Choose bitrate from the dropdown that appears
   - For FLAC: Choose compression level (0=fast, 8=best compression)
3. **Optional Settings**:
   - Check "Skip silence" to avoid recording silent audio
   - Check "Show only processes with active audio" to filter the list
4. **Set Output Location**: Use the Browse button to choose where recordings are saved (default: Documents\AudioCaptures)
5. **Start Recording**: Press Enter or click "Start Capture"
6. **Monitor Progress**: View active recordings in the "Active Recordings" list
7. **Stop Recording**: Select a recording and click "Stop Capture" (focus returns to process list)

### Multiple Simultaneous Captures

You can record multiple processes at once using checkboxes:
1. **Check Multiple Processes**: Click the checkbox next to each process you want to record
2. **Click "Start Capture"**: All checked processes will start recording simultaneously
3. **Each process records to its own file** with a unique timestamp
4. Stop individual recordings by selecting them in the "Active Recordings" list

### Filtering Processes

- **Show only processes with active audio**: Check this box to see only applications currently playing sound
- **Window titles**: The "Window Title" column helps identify processes (e.g., "YouTube - Chrome" vs "Gmail - Chrome")

### Output Files

Files are automatically named using the pattern:
```
{ProcessName}-YYYY_MM_DD-HH_MM_SS.{extension}
```

For example: `chrome-2025_10_12-14_30_45.flac`

## Technical Details

### Audio Capture Method

This application uses WASAPI (Windows Audio Session API) in loopback mode to capture audio. The implementation:
- Uses `IAudioClient` with `AUDCLNT_STREAMFLAGS_LOOPBACK` for capture
- Captures system-wide audio on Windows 10 (build < 20348)
- Can be extended for true per-process capture on Windows 10 Build 20348+ using Audio Graph API

### Supported Audio Formats

#### WAV (PCM)
- Uncompressed audio (32-bit float or 16-bit PCM)
- Highest quality, no loss
- Largest file size
- No additional codecs required
- Best for further editing

#### MP3
- Lossy compressed audio using Media Foundation
- Good quality at smaller file sizes
- Configurable bitrate: 128, 192, 256, or 320 kbps
- Default: 192 kbps
- Native Windows support, widely compatible

#### Opus
- Modern lossy codec optimized for internet streaming
- Excellent quality at low bitrates
- Configurable bitrate: 64, 96, 128, 192, or 256 kbps
- Default: 128 kbps
- Stored in OGG container format

#### FLAC
- Lossless compression (no quality loss)
- Typically 40-60% of WAV size
- Configurable compression levels 0-8
  - Level 0: Fastest encoding, larger files
  - Level 5: Default, good balance
  - Level 8: Slowest encoding, smallest files
- Ideal for archival and high-quality playback

### Architecture

The application is structured into several components:

- **AudioCapture**: WASAPI audio capture engine
- **ProcessEnumerator**: Enumerates running processes, window titles, and audio sessions
- **CaptureManager**: Manages multiple simultaneous capture sessions with silence detection
- **WavWriter**: Writes uncompressed WAV files
- **Mp3Encoder**: Encodes audio to MP3 using Media Foundation
- **OpusEncoder**: Encodes audio to Opus in OGG container
- **FlacEncoder**: Encodes audio to FLAC with configurable compression

## Limitations and Known Issues

### Current Limitations

1. **System-Wide Capture**: On Windows versions prior to Build 20348, the application captures all system audio, not just the selected process. The selected process serves as a label for organization.

2. **No Audio Preview**: The application doesn't provide real-time audio monitoring or visualization.

3. **Elevated Processes**: Cannot capture audio from processes running with higher privileges unless the application also runs elevated.

4. **First Refresh May Be Slow**: When you first click Refresh, the application fetches window titles for all processes, which can take a moment. Subsequent refreshes use cached data and are faster.

5. **Audio Session Detection**: The "Show only processes with active audio" filter checks for active audio sessions, which requires querying Windows Audio Session API and may add a slight delay.

### Potential Improvements

- Implement audio level meters and visualization
- Add pause/resume functionality
- Support for recording to multiple formats simultaneously
- Audio processing filters (volume normalization, noise reduction, etc.)
- Background worker thread for window title enumeration
- VU meter display for active recordings

## Accessibility

The application uses standard Win32 controls for full accessibility:
- List views with keyboard navigation and labeled columns
- Checkboxes for multi-selection with keyboard support
- Standard buttons with keyboard shortcuts (Tab/Enter navigation)
- Proper tab order for all controls
- ARIA-compliant labels for all UI elements
- Screen reader compatible (tested with NVDA/JAWS)
- Focus management (returns to process list after stopping capture)

## Troubleshooting

### No Audio Captured

- Ensure the target application is actually playing audio
- Check that your system audio is not muted
- Verify the application has permission to access audio devices
- Try running as Administrator if capturing from elevated processes

### Build Errors

- Ensure Windows SDK is properly installed
- Verify CMake version is 3.15 or later
- Check that you're using a C++17 compatible compiler
- Make sure all required Windows libraries are available

### Application Crashes

- Ensure you're running on Windows 10 or later
- Check that Media Foundation is available (required for MP3)
- Verify COM is properly initialized (automatic in this application)

## License

This project is provided as-is for educational and personal use.

## Credits

Built using:
- Windows Audio Session API (WASAPI)
- Media Foundation for MP3 encoding
- libFLAC for FLAC encoding
- libopus and libogg for Opus encoding
- nlohmann-json for settings persistence
- Win32 API for user interface
- CMake for build system
- vcpkg for dependency management
