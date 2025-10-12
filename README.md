# Audio Capture - Per-Process Audio Recording

A Windows application for capturing audio from individual processes using WASAPI (Windows Audio Session API). Records audio to WAV, MP3, Opus, or FLAC formats with an accessible Win32 interface.

## Full Disclosure

This was written with Claude Code. It has been tested and does function completely.

## Features

- **Per-Process Audio Capture**: Record audio from specific applications independently
- **System-Wide Audio Capture**: Option to capture all system audio simultaneously
- **Microphone Input Capture**: Record from microphone/line-in devices with seamless integration
- **Multi-Process Recording Modes**:
  - **Separate Files**: Each process records to its own file
  - **Combined File**: All processes mixed into a single output file
  - **Both**: Create individual files AND a combined mixed file
- **Real-time Audio Mixing**: Combine multiple audio streams into a single file
- **Audio Monitoring/Passthrough**: Send captured audio to another output device in real-time for monitoring
- **Monitor-Only Mode**: Listen to audio without recording it to disk
- **Multiple Format Support**: Save recordings as WAV, MP3, Opus, or FLAC files
  - **WAV**: Uncompressed PCM audio (highest quality, largest size)
  - **MP3**: Compressed with configurable bitrate (128-320 kbps)
  - **Opus**: Modern codec with configurable bitrate (64-256 kbps)
  - **FLAC**: Lossless compression with configurable levels (0-8)
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

### System-Wide Audio Capture

To capture all system audio at once:
1. **Select "[System Audio - All Processes]"**: This special entry at the top of the process list (PID 0) captures all system audio
2. **Start Capture**: Records everything your system is playing
3. Perfect for recording multiple applications together

### Audio Monitoring (Passthrough to Output Device)

You can send captured audio to another audio device in real-time:
1. **Enable "Monitor audio"**: Check the checkbox to activate passthrough
2. **Select Monitor Device**: Choose which output device to send audio to from the dropdown
3. **Start Capture**: Audio will play through both the original device and your selected monitor device
4. **Monitor-Only Mode**: Check "Monitor only - no recording" to listen without saving to disk
   - Recording format and output path controls are disabled in this mode
   - Useful for testing or temporary audio routing

### Multiple Simultaneous Captures with Recording Modes

You can record multiple processes at once with flexible output options:

#### Recording Mode Selection
Choose from the **Multi-process recording** dropdown:
- **Separate files**: Each process records to its own individual file
- **Combined file**: All processes are mixed into a single output file
- **Both**: Creates individual files AND a combined mixed file

#### Starting Multi-Process Capture
1. **Check Multiple Processes**: Click the checkbox next to each process you want to record
2. **Select Recording Mode**: Choose how you want the audio saved (Separate files / Combined file / Both)
3. **Click "Start Capture"**: All checked processes will start recording simultaneously
4. **Monitor Progress**: View all active recordings in the "Active Recordings" list
5. **Stop Individual Recordings**: Select a recording and click "Stop Capture"
6. **Stop All**: Click "Stop All" to end all active recordings at once

#### Recording Mode Examples
- **Separate files**: Recording Discord, Spotify, and Chrome creates three files:
  - `Discord-2025_10_12-14_30_45.opus`
  - `Spotify-2025_10_12-14_30_45.opus`
  - `chrome-2025_10_12-14_30_45.opus`

- **Combined file**: All three applications are mixed together into:
  - `Combined-2025_10_12-14_30_45.opus`

- **Both**: Creates all four files (three individual + one combined)

### Microphone Input Capture

You can capture microphone or line-in audio along with application audio:

#### Setting Up Microphone Capture
1. **Enable Microphone**: Check the "Capture microphone" checkbox
2. **Select Device**: Choose your microphone from the dropdown that appears
3. **Start Capture**: The microphone will be captured according to your recording mode

#### Microphone with Recording Modes
The microphone integrates seamlessly with multi-process recording modes:

- **Separate files mode**: Microphone records to its own file
  - Creates: `Microphone-2025_10_12-14_30_45.opus`
  - Each application also gets its own file

- **Combined file mode**: Microphone audio is mixed with application audio
  - Microphone audio is included in: `Combined-2025_10_12-14_30_45.opus`
  - No separate microphone file is created (appears as "Monitor Only" in the list)
  - Perfect for recording commentary over gameplay or music

- **Both mode**: Microphone creates its own file AND is included in the combined file
  - Creates: `Microphone-2025_10_12-14_30_45.opus` (separate mic file)
  - Microphone also mixed into: `Combined-2025_10_12-14_30_45.opus`
  - Best for maximum flexibility

#### Use Cases
- **Gaming Commentary**: Capture game audio + microphone in one file (Combined mode)
- **Music Recording**: Record DAW output + microphone vocals separately (Separate files mode)
- **Podcast Recording**: Capture multiple apps + microphone with both individual tracks and mixed output (Both mode)

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

This application uses WASAPI (Windows Audio Session API) to capture audio from multiple sources:

#### Loopback Capture (Application Audio)
- Uses `IAudioClient` with `AUDCLNT_STREAMFLAGS_LOOPBACK` for capturing application output
- Supports true per-process capture on Windows 10 Build 20348+ using `AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK`
- Falls back to system-wide capture on older Windows versions (build < 20348)

#### Input Device Capture (Microphone)
- Uses `IAudioClient` with capture mode (eCapture data flow direction) for microphone input
- Supports any WASAPI-compatible input device (microphones, line-in, etc.)
- Automatically handles device enumeration and format negotiation

#### Real-time Mixing
- Multiple audio streams are mixed in real-time using a dedicated mixer thread
- Automatic sample rate conversion and format matching
- 32-bit float PCM mixing for maximum quality and dynamic range
- Each stream can simultaneously record to its own file AND contribute to the mixed output

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

- **AudioCapture**: WASAPI audio capture engine with support for both loopback (application audio) and input device (microphone) capture, plus real-time passthrough
- **AudioDeviceEnumerator**: Enumerates available audio output devices (for monitoring) and input devices (microphones/line-in)
- **AudioMixer**: Real-time audio mixer that combines multiple audio streams with automatic resampling and format conversion
- **ProcessEnumerator**: Enumerates running processes, window titles, and audio sessions
- **CaptureManager**: Manages multiple simultaneous capture sessions with silence detection and coordinated mixing
- **WavWriter**: Writes uncompressed WAV files
- **Mp3Encoder**: Encodes audio to MP3 using Media Foundation
- **OpusEncoder**: Encodes audio to Opus in OGG container
- **FlacEncoder**: Encodes audio to FLAC with configurable compression

## Limitations and Known Issues

### Current Limitations

1. **System-Wide Capture**: On Windows versions prior to Build 20348, the application captures all system audio, not just the selected process. The selected process serves as a label for organization.

2. **Audio Monitoring Latency**: Real-time audio passthrough operates with approximately 100ms latency. This is optimized for minimal delay while maintaining stability.

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
- Lower latency passthrough options (experimental sub-50ms modes)

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
