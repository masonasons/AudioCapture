# Audio Capture - Per-Process Audio Recording

A Windows application for capturing audio from individual processes using WASAPI (Windows Audio Session API). Records audio to WAV, MP3, or Opus formats with an accessible Win32 interface.

## Full Disclosure

This was written with Claude Code. It has been tested and does function completely.

## Features

- **Per-Process Audio Capture**: Record audio from specific applications independently
- **Multiple Format Support**: Save recordings as WAV, MP3, or Opus files
- **Multi-Process Recording**: Capture audio from multiple processes simultaneously to different audio files
- **Accessible Win32 UI**: Standard Windows controls for full accessibility
- **Real-time Monitoring**: View active recording sessions and data statistics

## Requirements

### System Requirements
- Windows 10 Build 20348 or later (for optimal per-process capture)

### Build Requirements
- Visual Studio 2019 or later (with C++ Desktop Development workload)
- Windows SDK 10.0.19041.0 or later
- C++17 compatible compiler
- CMake 3.15 or later

## Building the Project

### Using Visual Studio

1. Open a Developer Command Prompt for Visual Studio

2. Navigate to the project directory:
3. Create a build directory:
```cmd
mkdir build
cd build
```

4. Generate Visual Studio project files:
```cmd
cmake .. -G "Visual Studio 17 2022" -A x64
```

5. Build the project:
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
2. The main window will display a list of running processes

### Capturing Audio

1. **Select a Process**: Click on a process in the top list view
2. **Choose Output Format**: Select WAV, MP3, or Opus from the dropdown
3. **Set Output Location**: Use the Browse button to choose where recordings are saved (default: Documents\AudioCaptures)
4. **Start Recording**: Click "Start Capture"
5. **Monitor Progress**: View active recordings in the bottom list view
6. **Stop Recording**: Select a recording and click "Stop Capture"

### Multiple Captures

You can record multiple processes simultaneously:
1. Start a capture for the first process
2. Select another process from the list
3. Click "Start Capture" again
4. Both recordings will appear in the active recordings list

### Output Files

Files are automatically named using the pattern:
```
{ProcessName}_{ProcessID}.{extension}
```

For example: `chrome_12345.mp3`

## Technical Details

### Audio Capture Method

This application uses WASAPI (Windows Audio Session API) in loopback mode to capture audio. The implementation:
- Uses `IAudioClient` with `AUDCLNT_STREAMFLAGS_LOOPBACK` for capture
- Captures system-wide audio on Windows 10 (build < 20348)
- Can be extended for true per-process capture on Windows 10 Build 20348+ using Audio Graph API

### Supported Audio Formats

#### WAV (PCM)
- Uncompressed audio
- Highest quality
- Largest file size
- No additional codecs required

#### MP3
- Compressed audio using Media Foundation
- Good quality at smaller file sizes
- Default bitrate: 192 kbps
- Native Windows support

#### Opus
- Modern compressed audio codec
- Currently saves as WAV format (placeholder)
- **Note**: Full Opus support requires integrating libopus library

### Architecture

The application is structured into several components:

- **AudioCapture**: WASAPI audio capture engine
- **ProcessEnumerator**: Lists running Windows processes
- **CaptureManager**: Manages multiple simultaneous capture sessions
- **WavWriter**: Writes uncompressed WAV files
- **Mp3Encoder**: Encodes audio to MP3 using Media Foundation
- **OpusEncoder**: Opus encoding (requires libopus)

## Limitations and Known Issues

### Current Limitations

1. **System-Wide Capture**: On Windows versions prior to Build 20348, the application captures all system audio, not just the selected process. The selected process serves as a label for organization.

2. **No Audio Preview**: The application doesn't provide real-time audio monitoring or visualization.

3. **Elevated Processes**: Cannot capture audio from processes running with higher privileges unless the application also runs elevated.

### Potential Improvements

- Implement audio level meters and visualization
- Add pause/resume functionality
- Support for recording to multiple formats simultaneously
- Audio processing filters (volume normalization, etc.)

## Accessibility

The application uses standard Win32 controls for full accessibility:
- List views with keyboard navigation
- Standard buttons with keyboard shortcuts
- Proper tab order for all controls
- Screen reader compatible

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
- Win32 API for user interface
- CMake for build system
