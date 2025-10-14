#include <windows.h>
#include <roapi.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <fstream>
#include <string>
#include <vector>
#include <memory>
#include <map>
#include <nlohmann/json.hpp>
#include "resource.h"
#include "CaptureManager.h"
#include "InputSourceManager.h"
#include "OutputDestinationManager.h"
#include "AudioDeviceEnumerator.h"
#include "WavFileDestination.h"
#include "Mp3FileDestination.h"
#include "OpusFileDestination.h"
#include "FlacFileDestination.h"
#include "DeviceOutputDestination.h"

using json = nlohmann::json;

#pragma comment(lib, "comctl32.lib")
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

// Global variables
HINSTANCE g_hInst;
HWND g_hWnd;

// UI Controls (in tab order)
HWND g_hInputSourcesLabel;     // "Input Sources" label
HWND g_hInputFilterCombo;      // Filter combo box for input sources
HWND g_hOutputDestsLabel;      // "Output Destinations" label
HWND g_hOutputFilterCombo;     // Filter combo box for output destinations
HWND g_hInputSourcesList;      // Checkable list of input sources
HWND g_hOutputDestsList;       // Checkable list of output destinations
HWND g_hBitrateEdit;           // Bitrate for MP3/Opus (kbps)
HWND g_hBitrateSpin;           // Spin control for bitrate
HWND g_hFlacCompressionEdit;   // FLAC compression level (0-8)
HWND g_hFlacCompressionSpin;   // Spin control for FLAC compression
HWND g_hCaptureModeGroup;      // Group box for capture mode
HWND g_hRadioSingleFile;       // Single file mode radio button
HWND g_hRadioMultipleFiles;    // Multiple files mode radio button
HWND g_hRadioBothModes;        // Both modes radio button
HWND g_hVolumeLabel;           // Volume label showing current source
HWND g_hVolumeSlider;          // Volume slider (0-100)
HWND g_hVolumeValue;           // Volume percentage display
HWND g_hOutputPath;
HWND g_hBrowseBtn;
HWND g_hRefreshBtn;
HWND g_hStartStopBtn;
HWND g_hSkipSilenceCheck;      // Skip silence checkbox
HWND g_hPauseResumeBtn;        // Pause/Resume button (only visible during file recording)
HWND g_hStatusText;
HACCEL g_hAccel;               // Accelerator table

// Managers
std::unique_ptr<CaptureManager> g_captureManager;
std::unique_ptr<InputSourceManager> g_sourceManager;
std::unique_ptr<OutputDestinationManager> g_destManager;
std::unique_ptr<AudioDeviceEnumerator> g_deviceEnumerator;

// State
std::vector<AvailableSource> g_availableSources;
std::vector<AudioDeviceInfo> g_availableOutputDevices;
std::vector<UINT32> g_activeSessionIds;  // Multiple sessions for multi-file mode
bool g_isCapturing = false;
bool g_useWinRT = false;
bool g_isFilesPaused = false;  // Track pause state for file destinations

// Volume settings per source (key = source ID, value = volume 0.0-1.0)
std::map<std::wstring, float> g_sourceVolumes;

// Active source tracking (key = source ID, value = InputSourcePtr) for real-time control
std::map<std::wstring, InputSourcePtr> g_activeSources;

// Active destination tracking (key = destination ID, value = OutputDestinationPtr) for real-time control
std::map<std::wstring, OutputDestinationPtr> g_activeDestinations;

// Active capture mode and settings (stored when capture starts, used for dynamic source/dest addition)
int g_activeCaptureMode = -1;  // -1 = not capturing, 0 = Single File, 1 = Multiple Files, 2 = Both Modes
std::vector<int> g_activeFileFormats;   // Which file formats are checked (0-3 for WAV/MP3/Opus/FLAC)
std::vector<int> g_activeDeviceIndices; // Which device indices are checked
std::wstring g_activeOutputPath;        // Output folder path
int g_activeBitrate;                    // Bitrate setting (in bps)
int g_activeFlacCompression;            // FLAC compression level (0-8)

// Session tracking: map sessionId to list of source IDs in that session
// This allows us to determine proper filenames when dynamically adding destinations
std::map<UINT32, std::vector<std::wstring>> g_sessionToSources;

// Store the audio format used for the current capture session
// This is needed for dynamically adding destinations during capture
std::vector<BYTE> g_activeCaptureFormat;

// CRITICAL: Track the last process ID that had a process-loopback audio client
// Windows needs time to fully release the process loopback interface after Stop
// If we try to re-initialize the same process too quickly, we get E_UNEXPECTED
DWORD g_lastProcessLoopbackPid = 0;
DWORD g_lastProcessLoopbackStopTime = 0;  // GetTickCount when stopped

// WORKAROUND for Windows 10 Build 19044 bug: Cache audio format from first initialization
// This allows us to skip creating tempSource on subsequent starts
std::vector<BYTE> g_cachedFormatForPid;  // Cached format for last successful process
DWORD g_cachedFormatPid = 0;  // PID this format belongs to

// Window class name
const wchar_t CLASS_NAME[] = L"AudioCaptureWindow";

// Destination type and index encoding for lParam - forward declarations
// Format: (type << 16) | index
// Type: 0 = File Format (WAV/MP3/Opus/FLAC), 1 = Device
// Index: For file formats: 0=WAV, 1=MP3, 2=Opus, 3=FLAC; For devices: device index in g_availableOutputDevices
inline LPARAM MakeDestinationParam(int type, int index) {
    return static_cast<LPARAM>((type << 16) | (index & 0xFFFF));
}

inline int GetDestinationType(LPARAM lParam) {
    return static_cast<int>(lParam >> 16);
}

inline int GetDestinationIndex(LPARAM lParam) {
    return static_cast<int>(lParam & 0xFFFF);
}

// Forward declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void InitializeControls(HWND hwnd);
void RefreshInputSources();
void RefreshOutputDestinations();
UINT32 CreatePerSourceSession(InputSourcePtr source, const WAVEFORMATEX* formatCopy, bool includeDevices);
void StartCapture();
void StopCapture();
void BrowseOutputFolder();
void UpdateStatus(const std::wstring& message);
void LoadSettings();
void SaveSettings();
std::wstring GetDefaultOutputPath();
std::wstring GetSettingsPath();
int GetBitrate();
int GetFlacCompression();
int GetCaptureMode();
std::wstring SanitizeFilename(const std::wstring& displayName);
void UpdateControlVisibility();
void UpdateVolumeControls();
void OnVolumeSliderChanged();
float GetSourceVolume(const std::wstring& sourceId);
void SetSourceVolume(const std::wstring& sourceId, float volume);

// WinMain entry point
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow) {
    g_hInst = hInstance;

    // Initialize COM - use MULTITHREADED for better audio capture performance
    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(hr)) {
        MessageBox(nullptr, L"Failed to initialize COM", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    // Try to initialize WinRT (Windows 10+)
    // Use __try/__except to handle delay-load DLL failure on Windows 7
    __try {
        hr = RoInitialize(RO_INIT_MULTITHREADED);
        if (SUCCEEDED(hr)) {
            g_useWinRT = true;
        }
    }
    __except (EXCEPTION_EXECUTE_HANDLER) {
        // WinRT not available (Windows 7 or older)
        g_useWinRT = false;
    }

    // Initialize common controls
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_LISTVIEW_CLASSES | ICC_STANDARD_CLASSES;
    InitCommonControlsEx(&icex);

    // Register window class
    WNDCLASS wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1));

    RegisterClass(&wc);

    // Create window
    g_hWnd = CreateWindowEx(
        0, CLASS_NAME, L"AudioCapture - Multi-Source Recording",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 1100, 700,
        nullptr, nullptr, hInstance, nullptr
    );

    if (g_hWnd == nullptr) {
        MessageBox(nullptr, L"Failed to create window", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    ShowWindow(g_hWnd, nCmdShow);
    UpdateWindow(g_hWnd);

    // Create accelerator table for keyboard shortcuts
    ACCEL accels[] = {
        { FVIRTKEY | FCONTROL, 'R', IDC_REFRESH_BTN },      // Ctrl+R = Refresh
        { FVIRTKEY | FCONTROL, 'S', IDC_START_STOP_BTN },   // Ctrl+S = Start/Stop
        { FVIRTKEY | FCONTROL, 'O', IDC_BROWSE_BTN },       // Ctrl+O = Open/Browse
        { FVIRTKEY, VK_F5, IDC_REFRESH_BTN },               // F5 = Refresh
    };
    g_hAccel = CreateAcceleratorTable(accels, ARRAYSIZE(accels));

    // Message loop with accelerator and dialog message support
    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0)) {
        // Check for accelerator keys first (Ctrl+R, Ctrl+S, etc.)
        if (!TranslateAccelerator(g_hWnd, g_hAccel, &msg)) {
            // Handle spacebar for list view checkboxes
            if (msg.message == WM_KEYDOWN && msg.wParam == VK_SPACE) {
                if (msg.hwnd == g_hInputSourcesList) {
                    int selected = ListView_GetNextItem(g_hInputSourcesList, -1, LVNI_FOCUSED);
                    if (selected != -1) {
                        BOOL checked = ListView_GetCheckState(g_hInputSourcesList, selected);
                        ListView_SetCheckState(g_hInputSourcesList, selected, !checked);
                        continue;  // Skip further processing
                    }
                }
                else if (msg.hwnd == g_hOutputDestsList) {
                    int selected = ListView_GetNextItem(g_hOutputDestsList, -1, LVNI_FOCUSED);
                    if (selected != -1) {
                        BOOL checked = ListView_GetCheckState(g_hOutputDestsList, selected);
                        ListView_SetCheckState(g_hOutputDestsList, selected, !checked);
                        continue;  // Skip further processing
                    }
                }
            }

            // Use IsDialogMessage for Tab navigation and other keyboard handling
            if (!IsDialogMessage(g_hWnd, &msg)) {
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
    }

    // Cleanup
    if (g_hAccel) {
        DestroyAcceleratorTable(g_hAccel);
    }

    // Cleanup
    if (g_captureManager) {
        g_captureManager->StopAll();
    }

    if (g_useWinRT) {
        __try {
            RoUninitialize();
        }
        __except (EXCEPTION_EXECUTE_HANDLER) {
            // Ignore if DLL unload fails
        }
    }
    CoUninitialize();

    return 0;
}

// Window procedure
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE:
        InitializeControls(hwnd);
        LoadSettings();
        RefreshInputSources();
        return 0;

    case WM_ACTIVATE:
        // Set focus to input sources list when window is activated
        if (LOWORD(wParam) != WA_INACTIVE && g_hInputSourcesList) {
            SetFocus(g_hInputSourcesList);
        }
        return 0;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDC_REFRESH_BTN) {
            RefreshInputSources();
            RefreshOutputDestinations();
        }
        else if (LOWORD(wParam) == IDC_START_STOP_BTN) {
            if (g_isCapturing) {
                StopCapture();
            } else {
                StartCapture();
            }
        }
        else if (LOWORD(wParam) == IDC_BROWSE_BTN) {
            BrowseOutputFolder();
        }
        else if (LOWORD(wParam) == IDC_PAUSE_RESUME_BTN) {
            // Toggle pause/resume state for file destinations
            if (g_isFilesPaused) {
                g_captureManager->ResumeFileDestinations();
                SetWindowText(g_hPauseResumeBtn, L"&Pause");
                g_isFilesPaused = false;
                UpdateStatus(L"File recording resumed");
            } else {
                g_captureManager->PauseFileDestinations();
                SetWindowText(g_hPauseResumeBtn, L"&Resume");
                g_isFilesPaused = true;
                UpdateStatus(L"File recording paused (device monitoring continues)");
            }
        }
        else if (LOWORD(wParam) == IDC_INPUT_FILTER_COMBO && HIWORD(wParam) == CBN_SELCHANGE) {
            // Filter combo box selection changed - refresh the input sources list
            RefreshInputSources();
        }
        else if (LOWORD(wParam) == IDC_OUTPUT_FILTER_COMBO && HIWORD(wParam) == CBN_SELCHANGE) {
            // Output filter combo box selection changed - refresh the output destinations list
            RefreshOutputDestinations();
        }
        return 0;

    case WM_NOTIFY:
        {
            LPNMHDR pnmh = (LPNMHDR)lParam;
            // Detect checkbox state changes in output destinations list
            if (pnmh->idFrom == IDC_OUTPUT_DESTS_LIST && pnmh->code == LVN_ITEMCHANGED) {
                LPNMLISTVIEW pnmv = (LPNMLISTVIEW)lParam;
                if (pnmv->uChanged & LVIF_STATE) {
                    // Checkbox state changed
                    UpdateControlVisibility();

                    // Handle real-time destination addition/removal during capture
                    if (g_isCapturing && (pnmv->uNewState & LVIS_STATEIMAGEMASK) != (pnmv->uOldState & LVIS_STATEIMAGEMASK)) {
                        int itemIndex = pnmv->iItem;
                        BOOL isNowChecked = ListView_GetCheckState(g_hOutputDestsList, itemIndex);


                        // Get the destination metadata from lParam
                        LVITEM lvi = {};
                        lvi.mask = LVIF_PARAM;
                        lvi.iItem = itemIndex;
                        if (!ListView_GetItem(g_hOutputDestsList, &lvi)) {
                            return 0;  // Failed to get item info
                        }

                        int destType = GetDestinationType(lvi.lParam);
                        int destIndex = GetDestinationIndex(lvi.lParam);

                        // Get output path and format info
                        wchar_t pathBuf[MAX_PATH];
                        GetWindowText(g_hOutputPath, pathBuf, MAX_PATH);
                        std::wstring outputPath = pathBuf;

                        if (isNowChecked) {
                            // Destination was checked - add to all active sessions
                            // Use the stored capture format from when capture was started
                            if (!g_activeCaptureFormat.empty()) {
                                const WAVEFORMATEX* formatCopy = reinterpret_cast<const WAVEFORMATEX*>(g_activeCaptureFormat.data());

                                int bitrate = GetBitrate() * 1000;
                                int flacCompression = GetFlacCompression();

                                // Lambda to create a destination with specific filename
                                auto CreateDestWithFilename = [&](const std::wstring& baseFilename) -> OutputDestinationPtr {
                                    OutputDestinationPtr dest;
                                    DestinationConfig config;
                                    config.useTimestamp = true;

                                    // Get skip silence setting from checkbox
                                    config.skipSilence = (SendMessage(g_hSkipSilenceCheck, BM_GETCHECK, 0, 0) == BST_CHECKED);
                                    config.silenceThreshold = 0.01f;  // 1% amplitude threshold
                                    config.silenceDurationMs = 1000;  // Skip after 1 second of silence

                                    // Create destination based on type and index from lParam
                                    if (destType == 0) {  // File format
                                        if (destIndex == 0) {  // WAV File
                                            config.outputPath = outputPath + L"\\" + baseFilename + L".wav";
                                            dest = std::make_shared<WavFileDestination>();
                                        }
                                        else if (destIndex == 1) {  // MP3 File
                                            config.outputPath = outputPath + L"\\" + baseFilename + L".mp3";
                                            config.bitrate = bitrate;
                                            dest = std::make_shared<Mp3FileDestination>();
                                        }
                                        else if (destIndex == 2) {  // Opus File
                                            config.outputPath = outputPath + L"\\" + baseFilename + L".opus";
                                            config.bitrate = bitrate;
                                            dest = std::make_shared<OpusFileDestination>();
                                        }
                                        else if (destIndex == 3) {  // FLAC File
                                            config.outputPath = outputPath + L"\\" + baseFilename + L".flac";
                                            config.compressionLevel = flacCompression;
                                            dest = std::make_shared<FlacFileDestination>();
                                        }
                                    }
                                    else if (destType == 1) {  // Audio Device
                                        if (destIndex >= 0 && destIndex < static_cast<int>(g_availableOutputDevices.size())) {
                                            dest = std::make_shared<DeviceOutputDestination>();
                                            config.outputPath = g_availableOutputDevices[destIndex].deviceId;
                                            config.friendlyName = g_availableOutputDevices[destIndex].friendlyName;
                                        }
                                    }

                                    if (dest && dest->Configure(formatCopy, config)) {
                                        return dest;
                                    }
                                    return nullptr;
                                };

                                // CRITICAL: Create a SEPARATE destination instance for EACH session
                                // with appropriate filename based on session type
                                bool addedToAny = false;
                                int destCreated = 0;
                                int destConfigured = 0;
                                int destAdded = 0;

                                for (size_t sessionIdx = 0; sessionIdx < g_activeSessionIds.size(); sessionIdx++) {
                                    UINT32 sessionId = g_activeSessionIds[sessionIdx];

                                    // Determine filename based on session type
                                    std::wstring baseFilename = L"capture";

                                    // In Mode 2 (Both), first session is ALWAYS the mixed session
                                    // Remaining sessions are per-source sessions
                                    bool isMixedSession = (g_activeCaptureMode == 2 && sessionIdx == 0);

                                    auto sessionSourcesIt = g_sessionToSources.find(sessionId);
                                    if (sessionSourcesIt != g_sessionToSources.end()) {
                                        const auto& sourceIds = sessionSourcesIt->second;

                                        // Use source-specific filename for per-source sessions
                                        if (!isMixedSession && sourceIds.size() == 1) {
                                            const std::wstring& sourceId = sourceIds[0];
                                            auto sourceIt = g_activeSources.find(sourceId);
                                            if (sourceIt != g_activeSources.end()) {
                                                std::wstring sourceName = sourceIt->second->GetMetadata().displayName;
                                                std::wstring sanitizedName = SanitizeFilename(sourceName);
                                                baseFilename = sanitizedName + L"_capture";
                                            }
                                        }
                                        // Otherwise use "capture" (mixed session or multi-source session)
                                    }

                                    OutputDestinationPtr dest = CreateDestWithFilename(baseFilename);
                                    destCreated++;
                                    if (dest) {
                                        destConfigured++;
                                        if (g_captureManager->AddOutputDestination(sessionId, dest)) {
                                            addedToAny = true;
                                            destAdded++;
                                            // Store each destination instance separately
                                            g_activeDestinations[dest->GetName()] = dest;
                                        }
                                    }
                                }

                                if (!addedToAny) {
                                    MessageBox(g_hWnd, L"Failed to add destination to any session", L"Warning", MB_OK | MB_ICONWARNING);
                                }
                            }
                        } else {
                            // Destination was unchecked - remove from all active sessions
                            // We need to find the actual destination name from g_activeDestinations
                            // because file destinations use timestamped paths as their names

                            // Search for destination by matching the output path pattern
                            std::wstring searchPattern;
                            if (destType == 0) {  // File format
                                if (destIndex == 0) searchPattern = L".wav";
                                else if (destIndex == 1) searchPattern = L".mp3";
                                else if (destIndex == 2) searchPattern = L".opus";
                                else if (destIndex == 3) searchPattern = L".flac";
                            }
                            else if (destType == 1) {  // Audio Device
                                if (destIndex >= 0 && destIndex < static_cast<int>(g_availableOutputDevices.size())) {
                                    searchPattern = g_availableOutputDevices[destIndex].friendlyName;
                                }
                            }

                            if (!searchPattern.empty()) {
                                // Find matching destination(s) in active destinations
                                std::vector<std::wstring> destIdsToRemove;
                                for (const auto& pair : g_activeDestinations) {
                                    const std::wstring& destName = pair.first;
                                    // For file formats, match by extension; for devices, match exact name
                                    if (destType == 0) {
                                        // File format - check if name ends with the extension
                                        if (destName.length() >= searchPattern.length() &&
                                            destName.substr(destName.length() - searchPattern.length()) == searchPattern) {
                                            destIdsToRemove.push_back(destName);
                                        }
                                    } else {
                                        // Device - match exact name
                                        if (destName == searchPattern) {
                                            destIdsToRemove.push_back(destName);
                                        }
                                    }
                                }

                                // Remove all matching destinations
                                for (const auto& destId : destIdsToRemove) {
                                    for (UINT32 sessionId : g_activeSessionIds) {
                                        g_captureManager->RemoveOutputDestination(sessionId, destId);
                                    }
                                    g_activeDestinations.erase(destId);
                                }
                            }
                        }
                    }
                }
            }
            // Detect focus and checkbox state changes in input sources list
            else if (pnmh->idFrom == IDC_INPUT_SOURCES_LIST && pnmh->code == LVN_ITEMCHANGED) {
                LPNMLISTVIEW pnmv = (LPNMLISTVIEW)lParam;
                // Update volume controls when focus changes or checkbox state changes
                if ((pnmv->uChanged & LVIF_STATE) &&
                    ((pnmv->uNewState & LVIS_FOCUSED) != (pnmv->uOldState & LVIS_FOCUSED) ||
                     (pnmv->uNewState & LVIS_STATEIMAGEMASK) != (pnmv->uOldState & LVIS_STATEIMAGEMASK))) {
                    UpdateVolumeControls();

                    // Handle real-time source addition/removal during capture
                    if (g_isCapturing && (pnmv->uNewState & LVIS_STATEIMAGEMASK) != (pnmv->uOldState & LVIS_STATEIMAGEMASK)) {
                        // Checkbox state changed during capture
                        int itemIndex = pnmv->iItem;
                        if (itemIndex >= 0 && itemIndex < static_cast<int>(g_availableSources.size())) {
                            BOOL isNowChecked = ListView_GetCheckState(g_hInputSourcesList, itemIndex);
                            const auto& sourceMetadata = g_availableSources[itemIndex];

                            if (isNowChecked) {
                                // Source was checked - add dynamically based on capture mode
                                auto source = g_sourceManager->CreateSource(sourceMetadata);
                                if (source) {
                                    // Apply volume setting
                                    float volume = GetSourceVolume(source->GetMetadata().id);
                                    source->SetVolume(volume);

                                    // Start the source temporarily to get its format
                                    if (!source->StartCapture()) {
                                        ListView_SetCheckState(g_hInputSourcesList, itemIndex, FALSE);
                                        MessageBox(g_hWnd, L"Failed to start source", L"Error", MB_OK | MB_ICONWARNING);
                                        return 0;
                                    }

                                    // Get format from the new source
                                    const WAVEFORMATEX* format = source->GetFormat();
                                    if (!format) {
                                        source->StopCapture();
                                        ListView_SetCheckState(g_hInputSourcesList, itemIndex, FALSE);
                                        MessageBox(g_hWnd, L"Cannot get audio format from source", L"Error", MB_OK | MB_ICONWARNING);
                                        return 0;
                                    }

                                    // Make a copy of the format
                                    UINT32 formatSize = sizeof(WAVEFORMATEX) + format->cbSize;
                                    std::vector<BYTE> formatBuffer(formatSize);
                                    memcpy(formatBuffer.data(), format, formatSize);
                                    const WAVEFORMATEX* formatCopy = reinterpret_cast<const WAVEFORMATEX*>(formatBuffer.data());

                                    // Stop the source - CaptureManager will start it properly
                                    source->StopCapture();

                                    bool addedToAny = false;
                                    bool hadFailure = false;

                                    // Mode 0: Single File - add to all existing sessions
                                    if (g_activeCaptureMode == 0) {
                                        for (UINT32 sessionId : g_activeSessionIds) {
                                            if (g_captureManager->AddInputSource(sessionId, source)) {
                                                addedToAny = true;
                                            } else {
                                                hadFailure = true;
                                            }
                                        }
                                    }
                                    // Mode 1: Multiple Files - create new per-source session
                                    else if (g_activeCaptureMode == 1) {
                                        // In Mode 1, include devices in per-source sessions (each source can be monitored)
                                        UINT32 sessionId = CreatePerSourceSession(source, formatCopy, true);
                                        if (sessionId != 0) {
                                            g_activeSessionIds.push_back(sessionId);
                                            addedToAny = true;
                                        }
                                    }
                                    // Mode 2: Both Modes - add to mixed session AND create per-source session
                                    else if (g_activeCaptureMode == 2) {
                                        // Add to first session (mixed session which has the devices)
                                        if (!g_activeSessionIds.empty()) {
                                            if (g_captureManager->AddInputSource(g_activeSessionIds[0], source)) {
                                                addedToAny = true;
                                            }
                                        }

                                        // Also create per-source session (no devices - mixed session has them)
                                        UINT32 sessionId = CreatePerSourceSession(source, formatCopy, false);
                                        if (sessionId != 0) {
                                            g_activeSessionIds.push_back(sessionId);
                                            addedToAny = true;
                                        }
                                    }

                                    // Store for real-time control only if successfully added
                                    if (addedToAny) {
                                        g_activeSources[source->GetMetadata().id] = source;

                                        // Update status
                                        std::wstring statusMsg = L"Added source: " + sourceMetadata.metadata.displayName;
                                        if (hadFailure) {
                                            statusMsg += L" (some sessions failed)";
                                        }
                                        UpdateStatus(statusMsg);
                                    } else {
                                        // Failed to add - uncheck the checkbox
                                        ListView_SetCheckState(g_hInputSourcesList, itemIndex, FALSE);
                                        MessageBox(g_hWnd, L"Failed to add source", L"Error", MB_OK | MB_ICONWARNING);
                                        UpdateStatus(L"Failed to add source: " + sourceMetadata.metadata.displayName);
                                    }
                                } else {
                                    // Failed to create source - uncheck the checkbox
                                    ListView_SetCheckState(g_hInputSourcesList, itemIndex, FALSE);
                                    UpdateStatus(L"Failed to create source: " + sourceMetadata.metadata.displayName);
                                }
                            } else {
                                // Source was unchecked - remove from sessions
                                bool removedAny = false;

                                // Mode 0: Remove from all sessions (current behavior)
                                if (g_activeCaptureMode == 0) {
                                    for (UINT32 sessionId : g_activeSessionIds) {
                                        if (g_captureManager->RemoveInputSource(sessionId, sourceMetadata.metadata.id)) {
                                            removedAny = true;
                                        }
                                    }
                                }
                                // Mode 1: Find and stop the per-source session
                                else if (g_activeCaptureMode == 1) {
                                    // In Mode 1, each source has its own session
                                    // Remove the source from all sessions (will find its session)
                                    std::vector<UINT32> sessionsToRemove;
                                    for (UINT32 sessionId : g_activeSessionIds) {
                                        if (g_captureManager->RemoveInputSource(sessionId, sourceMetadata.metadata.id)) {
                                            removedAny = true;
                                            // Stop this session since it only had this one source
                                            g_captureManager->StopCaptureSession(sessionId);
                                            sessionsToRemove.push_back(sessionId);
                                        }
                                    }
                                    // Remove from session IDs list
                                    for (UINT32 sid : sessionsToRemove) {
                                        g_activeSessionIds.erase(std::remove(g_activeSessionIds.begin(), g_activeSessionIds.end(), sid), g_activeSessionIds.end());
                                    }
                                }
                                // Mode 2: Remove from mixed session AND stop per-source session
                                else if (g_activeCaptureMode == 2) {
                                    // Remove from first session (mixed)
                                    if (!g_activeSessionIds.empty()) {
                                        if (g_captureManager->RemoveInputSource(g_activeSessionIds[0], sourceMetadata.metadata.id)) {
                                            removedAny = true;
                                        }
                                    }

                                    // Find and stop the per-source session (skip first session which is mixed)
                                    std::vector<UINT32> sessionsToRemove;
                                    for (size_t i = 1; i < g_activeSessionIds.size(); i++) {
                                        UINT32 sessionId = g_activeSessionIds[i];
                                        if (g_captureManager->RemoveInputSource(sessionId, sourceMetadata.metadata.id)) {
                                            // Stop this per-source session
                                            g_captureManager->StopCaptureSession(sessionId);
                                            sessionsToRemove.push_back(sessionId);
                                        }
                                    }
                                    // Remove from session IDs list
                                    for (UINT32 sid : sessionsToRemove) {
                                        g_activeSessionIds.erase(std::remove(g_activeSessionIds.begin(), g_activeSessionIds.end(), sid), g_activeSessionIds.end());
                                    }
                                }

                                // Remove from active sources tracking
                                g_activeSources.erase(sourceMetadata.metadata.id);

                                // Update status
                                if (removedAny) {
                                    UpdateStatus(L"Removed source: " + sourceMetadata.metadata.displayName);
                                }
                            }
                        }
                    }
                }
            }
        }
        return 0;

    case WM_HSCROLL:
        {
            // Handle volume slider changes
            if ((HWND)lParam == g_hVolumeSlider) {
                OnVolumeSliderChanged();
            }
        }
        return 0;

    case WM_SIZE:
        // Auto-resize controls
        {
            RECT rc;
            GetClientRect(hwnd, &rc);
            int width = rc.right - rc.left;
            int height = rc.bottom - rc.top;

            // Input filter combo box (left side, between label and list)
            SetWindowPos(g_hInputFilterCombo, nullptr, 10, 32, 250, 25, SWP_NOZORDER);

            // Input sources list (left half, below combo box)
            SetWindowPos(g_hInputSourcesList, nullptr, 10, 60, (width / 2) - 20, height - 150, SWP_NOZORDER);

            // Output filter combo box (right side, between label and list)
            SetWindowPos(g_hOutputFilterCombo, nullptr, (width / 2) + 10, 32, 200, 25, SWP_NOZORDER);

            // Output destinations list (right half, below combo box)
            SetWindowPos(g_hOutputDestsList, nullptr, (width / 2) + 10, 60, (width / 2) - 20, height - 150, SWP_NOZORDER);

            // Bottom controls
            SetWindowPos(g_hOutputPath, nullptr, 80, height - 80, width - 230, 25, SWP_NOZORDER);
            SetWindowPos(g_hBrowseBtn, nullptr, width - 140, height - 80, 60, 25, SWP_NOZORDER);
            SetWindowPos(g_hStartStopBtn, nullptr, width - 70, height - 80, 60, 25, SWP_NOZORDER);
            SetWindowPos(g_hStatusText, nullptr, 10, height - 45, width - 20, 35, SWP_NOZORDER);
        }
        return 0;

    case WM_CLOSE:
        SaveSettings();
        DestroyWindow(hwnd);
        return 0;

    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

// Initialize all UI controls
void InitializeControls(HWND hwnd) {
    // Create managers
    g_captureManager = std::make_unique<CaptureManager>();
    g_sourceManager = std::make_unique<InputSourceManager>();
    g_destManager = std::make_unique<OutputDestinationManager>();
    g_deviceEnumerator = std::make_unique<AudioDeviceEnumerator>();

    // Title labels (static text - not tabbable) - Store HWNDs for control
    g_hInputSourcesLabel = CreateWindow(L"STATIC", L"Input Sources (check to capture):",
        WS_VISIBLE | WS_CHILD,
        10, 10, 400, 20, hwnd, (HMENU)0x2000, g_hInst, nullptr);

    // Input filter combo box (between label and list)
    g_hInputFilterCombo = CreateWindowEx(
        0,
        L"COMBOBOX", nullptr,
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | CBS_DROPDOWNLIST | CBS_HASSTRINGS,
        10, 32, 250, 200, hwnd, (HMENU)IDC_INPUT_FILTER_COMBO, g_hInst, nullptr);

    // Populate the filter combo box
    SendMessage(g_hInputFilterCombo, CB_ADDSTRING, 0, (LPARAM)L"All");
    SendMessage(g_hInputFilterCombo, CB_ADDSTRING, 0, (LPARAM)L"Input Devices");

    // Only add process filters if WinRT is available (Windows 10 Build 19041+)
    if (g_useWinRT) {
        SendMessage(g_hInputFilterCombo, CB_ADDSTRING, 0, (LPARAM)L"Processes");
        SendMessage(g_hInputFilterCombo, CB_ADDSTRING, 0, (LPARAM)L"Processes with Audio Sessions Only");
    }

    SendMessage(g_hInputFilterCombo, CB_SETCURSEL, 0, 0);  // Default to "All"

    g_hOutputDestsLabel = CreateWindow(L"STATIC", L"Output Destinations (check to record/monitor):",
        WS_VISIBLE | WS_CHILD,
        560, 10, 400, 20, hwnd, (HMENU)0x2001, g_hInst, nullptr);

    // Output filter combo box (between label and list)
    g_hOutputFilterCombo = CreateWindowEx(
        0,
        L"COMBOBOX", nullptr,
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | CBS_DROPDOWNLIST | CBS_HASSTRINGS,
        560, 32, 200, 200, hwnd, (HMENU)IDC_OUTPUT_FILTER_COMBO, g_hInst, nullptr);

    // Populate the output filter combo box
    SendMessage(g_hOutputFilterCombo, CB_ADDSTRING, 0, (LPARAM)L"All");
    SendMessage(g_hOutputFilterCombo, CB_ADDSTRING, 0, (LPARAM)L"File Formats");
    SendMessage(g_hOutputFilterCombo, CB_ADDSTRING, 0, (LPARAM)L"Output Devices");
    SendMessage(g_hOutputFilterCombo, CB_SETCURSEL, 0, 0);  // Default to "All"

    // Settings labels and controls (on the right side under output destinations)
    // Initially hidden - will show when MP3/Opus/FLAC are selected
    CreateWindow(L"STATIC", L"Bitrate (kbps):",
        WS_CHILD,  // Hidden initially
        560, 545, 100, 20, hwnd, (HMENU)0x1000, g_hInst, nullptr);

    g_hBitrateEdit = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        L"EDIT", L"192",
        WS_CHILD | WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL,  // Hidden initially
        670, 543, 60, 22, hwnd, nullptr, g_hInst, nullptr);

    g_hBitrateSpin = CreateWindowEx(
        0, UPDOWN_CLASS, nullptr,
        WS_CHILD | UDS_SETBUDDYINT | UDS_ALIGNRIGHT | UDS_ARROWKEYS,  // Hidden initially
        0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
    SendMessage(g_hBitrateSpin, UDM_SETBUDDY, (WPARAM)g_hBitrateEdit, 0);
    SendMessage(g_hBitrateSpin, UDM_SETRANGE, 0, MAKELPARAM(320, 64));
    SendMessage(g_hBitrateSpin, UDM_SETPOS, 0, 192);

    CreateWindow(L"STATIC", L"FLAC Level:",
        WS_CHILD,  // Hidden initially
        750, 545, 80, 20, hwnd, (HMENU)0x1001, g_hInst, nullptr);

    g_hFlacCompressionEdit = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        L"EDIT", L"5",
        WS_CHILD | WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL,  // Hidden initially
        840, 543, 40, 22, hwnd, nullptr, g_hInst, nullptr);

    g_hFlacCompressionSpin = CreateWindowEx(
        0, UPDOWN_CLASS, nullptr,
        WS_CHILD | UDS_SETBUDDYINT | UDS_ALIGNRIGHT | UDS_ARROWKEYS,  // Hidden initially
        0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
    SendMessage(g_hFlacCompressionSpin, UDM_SETBUDDY, (WPARAM)g_hFlacCompressionEdit, 0);
    SendMessage(g_hFlacCompressionSpin, UDM_SETRANGE, 0, MAKELPARAM(8, 0));
    SendMessage(g_hFlacCompressionSpin, UDM_SETPOS, 0, 5);

    // Capture Mode Selection Group - positioned near output folder controls
    g_hCaptureModeGroup = CreateWindow(L"BUTTON", L"Capture Mode:",
        WS_VISIBLE | WS_CHILD | BS_GROUPBOX,
        900, 543, 180, 70, hwnd, (HMENU)IDC_CAPTURE_MODE_GROUP, g_hInst, nullptr);

    // Radio buttons for capture mode (inside group box)
    g_hRadioSingleFile = CreateWindow(L"BUTTON", L"Single &File (mixed)",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTORADIOBUTTON | WS_GROUP,
        910, 563, 160, 20, hwnd, (HMENU)IDC_RADIO_SINGLE_FILE, g_hInst, nullptr);

    g_hRadioMultipleFiles = CreateWindow(L"BUTTON", L"&Multiple Files",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTORADIOBUTTON,
        910, 583, 120, 20, hwnd, (HMENU)IDC_RADIO_MULTI_FILES, g_hInst, nullptr);

    g_hRadioBothModes = CreateWindow(L"BUTTON", L"Bot&h Modes",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTORADIOBUTTON,
        1030, 583, 100, 20, hwnd, (HMENU)IDC_RADIO_BOTH_MODES, g_hInst, nullptr);

    // Set default selection to Single File mode
    SendMessage(g_hRadioSingleFile, BM_SETCHECK, BST_CHECKED, 0);

    // Input sources list (checkable) - TAB ORDER 1
    g_hInputSourcesList = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        WC_LISTVIEW, L"Input Sources (check to capture)",
        WS_VISIBLE | WS_CHILD | WS_BORDER | WS_TABSTOP | LVS_REPORT | LVS_SINGLESEL,
        10, 35, 530, 500, hwnd, (HMENU)IDC_INPUT_SOURCES_LIST, g_hInst, nullptr);

    ListView_SetExtendedListViewStyle(g_hInputSourcesList, LVS_EX_CHECKBOXES | LVS_EX_FULLROWSELECT);

    // Add columns
    LVCOLUMN lvc = {};
    lvc.mask = LVCF_TEXT | LVCF_WIDTH;
    lvc.pszText = (LPWSTR)L"Source";
    lvc.cx = 350;
    ListView_InsertColumn(g_hInputSourcesList, 0, &lvc);

    lvc.pszText = (LPWSTR)L"Type";
    lvc.cx = 150;
    ListView_InsertColumn(g_hInputSourcesList, 1, &lvc);

    // Volume controls (below input sources list) - Initially hidden
    // Positioned entirely on the left side to avoid screen reader confusion with right side controls
    g_hVolumeLabel = CreateWindow(L"STATIC", L"Volume: (select a source)",
        WS_CHILD,  // Hidden initially
        10, 543, 240, 20, hwnd, (HMENU)IDC_VOLUME_LABEL, g_hInst, nullptr);

    g_hVolumeSlider = CreateWindowEx(
        0, TRACKBAR_CLASS, nullptr,
        WS_CHILD | WS_TABSTOP | TBS_HORZ | TBS_AUTOTICKS,  // Hidden initially
        260, 540, 200, 30, hwnd, (HMENU)IDC_VOLUME_SLIDER, g_hInst, nullptr);
    SendMessage(g_hVolumeSlider, TBM_SETRANGE, TRUE, MAKELONG(0, 100));
    SendMessage(g_hVolumeSlider, TBM_SETPOS, TRUE, 100);  // Default 100%

    g_hVolumeValue = CreateWindow(L"STATIC", L"",
        WS_CHILD,  // Hidden initially
        470, 543, 50, 20, hwnd, (HMENU)IDC_VOLUME_VALUE, g_hInst, nullptr);

    // Output destinations list (checkable) - TAB ORDER 2
    // Positioned below the filter combo box
    g_hOutputDestsList = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        WC_LISTVIEW, L"Output Destinations (check to record/monitor)",
        WS_VISIBLE | WS_CHILD | WS_BORDER | WS_TABSTOP | LVS_REPORT | LVS_SINGLESEL,
        560, 60, 530, 475, hwnd, (HMENU)IDC_OUTPUT_DESTS_LIST, g_hInst, nullptr);

    ListView_SetExtendedListViewStyle(g_hOutputDestsList, LVS_EX_CHECKBOXES | LVS_EX_FULLROWSELECT);

    lvc.pszText = (LPWSTR)L"Destination";
    lvc.cx = 250;
    ListView_InsertColumn(g_hOutputDestsList, 0, &lvc);

    lvc.pszText = (LPWSTR)L"Type";
    lvc.cx = 250;
    ListView_InsertColumn(g_hOutputDestsList, 1, &lvc);

    // Populate output destinations with file formats and audio devices
    RefreshOutputDestinations();

    // Bottom controls label (not tabbable) - ASSIGN UNIQUE ID
    CreateWindow(L"STATIC", L"Output Folder:",
        WS_VISIBLE | WS_CHILD,
        10, 550, 70, 20, hwnd, (HMENU)0x2002, g_hInst, nullptr);

    // Output path edit box - TAB ORDER 3
    g_hOutputPath = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        L"EDIT", GetDefaultOutputPath().c_str(),
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | ES_AUTOHSCROLL,
        80, 545, 880, 25, hwnd, (HMENU)IDC_OUTPUT_PATH, g_hInst, nullptr);

    // Browse button - TAB ORDER 4
    g_hBrowseBtn = CreateWindow(L"BUTTON", L"&Browse...",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_PUSHBUTTON,
        970, 545, 60, 25, hwnd, (HMENU)IDC_BROWSE_BTN, g_hInst, nullptr);

    // Skip silence checkbox - positioned above output path
    g_hSkipSilenceCheck = CreateWindow(L"BUTTON", L"Skip &Silence (files only)",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX,
        560, 543, 160, 20, hwnd, (HMENU)IDC_SKIP_SILENCE_CHECK, g_hInst, nullptr);

    // Refresh button - TAB ORDER 5
    g_hRefreshBtn = CreateWindow(L"BUTTON", L"&Refresh",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_PUSHBUTTON,
        10, 580, 80, 30, hwnd, (HMENU)IDC_REFRESH_BTN, g_hInst, nullptr);

    // Pause/Resume button - TAB ORDER 6 (initially hidden, shown only during file recording)
    g_hPauseResumeBtn = CreateWindow(L"BUTTON", L"&Pause",
        WS_CHILD | WS_TABSTOP | BS_PUSHBUTTON,  // Not WS_VISIBLE initially
        100, 580, 80, 30, hwnd, (HMENU)IDC_PAUSE_RESUME_BTN, g_hInst, nullptr);

    // Start/Stop button - TAB ORDER 7
    g_hStartStopBtn = CreateWindow(L"BUTTON", L"&Start",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_PUSHBUTTON | BS_DEFPUSHBUTTON,
        1040, 545, 50, 25, hwnd, (HMENU)IDC_START_STOP_BTN, g_hInst, nullptr);

    // Status text (not tabbable)
    g_hStatusText = CreateWindow(L"STATIC", L"Ready. Select sources and destinations, then click Start.\r\n\r\nKeyboard shortcuts: F5 or Ctrl+R = Refresh | Ctrl+S = Start/Stop | Ctrl+O = Browse | Space = Toggle checkbox",
        WS_VISIBLE | WS_CHILD | SS_LEFT,
        10, 620, 1080, 40, hwnd, nullptr, g_hInst, nullptr);

    // Set initial focus to input sources list
    SetFocus(g_hInputSourcesList);
}

// Refresh the input sources list
void RefreshInputSources() {
    ListView_DeleteAllItems(g_hInputSourcesList);
    g_availableSources.clear();

    UpdateStatus(L"Refreshing sources...");

    // Get selected filter from combo box
    int filterIndex = static_cast<int>(SendMessage(g_hInputFilterCombo, CB_GETCURSEL, 0, 0));
    // 0 = All, 1 = Input Devices, 2 = Processes, 3 = Processes with Audio Sessions Only

    bool includeProcesses = (filterIndex == 0 || filterIndex == 2 || filterIndex == 3) && g_useWinRT;  // Processes require WinRT (Windows 10+)
    bool includeInputDevices = (filterIndex == 0 || filterIndex == 1);
    bool includeSystemAudio = (filterIndex == 0);  // System audio only in "All" mode

    // Refresh all source types (no output devices - those are destinations)
    g_sourceManager->RefreshAvailableSources(
        includeProcesses,      // Include processes based on filter AND WinRT availability
        includeSystemAudio,    // Include system audio based on filter
        includeInputDevices,   // Include input devices based on filter
        false                  // Don't include output devices here - they go in destinations
    );

    g_availableSources = g_sourceManager->GetAvailableSources();

    // For "Processes with Audio Sessions Only" filter, we need to check each process
    if (filterIndex == 3) {
        std::vector<AvailableSource> filteredSources;
        for (const auto& source : g_availableSources) {
            if (source.metadata.type == InputSourceType::Process) {
                // Check if this process has active audio sessions
                // Use ProcessEnumerator to check for active audio
                ProcessInfo procInfo = g_sourceManager->FindProcessInfo(source.metadata.processId);
                if (procInfo.hasActiveAudio) {
                    filteredSources.push_back(source);
                }
            } else {
                // Keep non-process sources (shouldn't happen with this filter, but just in case)
                filteredSources.push_back(source);
            }
        }
        g_availableSources = filteredSources;
    }

    // Populate list
    LVITEM lvi = {};
    lvi.mask = LVIF_TEXT;
    int displayIndex = 0;

    for (size_t i = 0; i < g_availableSources.size(); i++) {
        const auto& source = g_availableSources[i];

        lvi.iItem = displayIndex++;
        lvi.iSubItem = 0;
        lvi.pszText = const_cast<LPWSTR>(source.metadata.displayName.c_str());
        ListView_InsertItem(g_hInputSourcesList, &lvi);

        // Set type column
        const wchar_t* typeStr = L"Unknown";
        switch (source.metadata.type) {
        case InputSourceType::Process:
            typeStr = L"Process";
            break;
        case InputSourceType::SystemAudio:
            typeStr = L"System Audio";
            break;
        case InputSourceType::InputDevice:
            typeStr = L"Microphone";
            break;
        }
        ListView_SetItemText(g_hInputSourcesList, lvi.iItem, 1, const_cast<LPWSTR>(typeStr));
    }

    UpdateStatus(L"Found " + std::to_wstring(g_availableSources.size()) + L" input sources. Ready to capture.");
}

// Refresh the output destinations list
void RefreshOutputDestinations() {
    ListView_DeleteAllItems(g_hOutputDestsList);

    // Get selected filter from combo box
    int filterIndex = static_cast<int>(SendMessage(g_hOutputFilterCombo, CB_GETCURSEL, 0, 0));
    if (filterIndex == CB_ERR) {
        filterIndex = 0;  // Default to "All" if nothing selected
    }
    // 0 = All, 1 = File Formats, 2 = Output Devices

    LVITEM lvi = {};
    lvi.mask = LVIF_TEXT | LVIF_PARAM;  // Include LPARAM for storing destination metadata
    int itemIndex = 0;

    // Add file format destinations (if filter allows)
    if (filterIndex == 0 || filterIndex == 1) {
        lvi.iItem = itemIndex++;
        lvi.iSubItem = 0;
        lvi.pszText = (LPWSTR)L"WAV File";
        lvi.lParam = MakeDestinationParam(0, 0);  // Type=0 (File), Index=0 (WAV)
        ListView_InsertItem(g_hOutputDestsList, &lvi);
        ListView_SetItemText(g_hOutputDestsList, lvi.iItem, 1, (LPWSTR)L"Uncompressed Audio");

        lvi.iItem = itemIndex++;
        lvi.pszText = (LPWSTR)L"MP3 File";
        lvi.lParam = MakeDestinationParam(0, 1);  // Type=0 (File), Index=1 (MP3)
        ListView_InsertItem(g_hOutputDestsList, &lvi);
        ListView_SetItemText(g_hOutputDestsList, lvi.iItem, 1, (LPWSTR)L"Compressed Audio");

        lvi.iItem = itemIndex++;
        lvi.pszText = (LPWSTR)L"Opus File";
        lvi.lParam = MakeDestinationParam(0, 2);  // Type=0 (File), Index=2 (Opus)
        ListView_InsertItem(g_hOutputDestsList, &lvi);
        ListView_SetItemText(g_hOutputDestsList, lvi.iItem, 1, (LPWSTR)L"Compressed Audio");

        lvi.iItem = itemIndex++;
        lvi.pszText = (LPWSTR)L"FLAC File";
        lvi.lParam = MakeDestinationParam(0, 3);  // Type=0 (File), Index=3 (FLAC)
        ListView_InsertItem(g_hOutputDestsList, &lvi);
        ListView_SetItemText(g_hOutputDestsList, lvi.iItem, 1, (LPWSTR)L"Lossless Compression");
    }

    // Enumerate and add audio output devices (if filter allows)
    if (filterIndex == 0 || filterIndex == 2) {
        // Only re-enumerate devices if needed (or if we cleared them)
        if (g_availableOutputDevices.empty() || filterIndex == 2) {
            g_availableOutputDevices.clear();
            if (g_deviceEnumerator->EnumerateDevices()) {
                g_availableOutputDevices = g_deviceEnumerator->GetDevices();
            }
        }

        for (size_t i = 0; i < g_availableOutputDevices.size(); i++) {
            const auto& device = g_availableOutputDevices[i];
            lvi.iItem = itemIndex++;
            lvi.pszText = const_cast<LPWSTR>(device.friendlyName.c_str());
            lvi.lParam = MakeDestinationParam(1, static_cast<int>(i));  // Type=1 (Device), Index=device index
            ListView_InsertItem(g_hOutputDestsList, &lvi);

            std::wstring typeStr = device.isDefault ? L"Audio Device (Default)" : L"Audio Device";
            ListView_SetItemText(g_hOutputDestsList, lvi.iItem, 1, const_cast<LPWSTR>(typeStr.c_str()));
        }
    }
}

// Helper function to create a per-source session with separate files
UINT32 CreatePerSourceSession(InputSourcePtr source, const WAVEFORMATEX* formatCopy, bool includeDevices) {
    std::vector<OutputDestinationPtr> sourceDestinations;

    // Get sanitized source name
    InputSourceMetadata metadata = source->GetMetadata();
    std::wstring sanitizedName = SanitizeFilename(metadata.displayName);

    // Lambda to create destination (copied from StartCapture)
    auto CreateDestination = [&](int formatIndex, const std::wstring& baseFilename) -> OutputDestinationPtr {
        OutputDestinationPtr dest;
        DestinationConfig config;
        config.useTimestamp = true;

        if (formatIndex == 0) {  // WAV
            config.outputPath = g_activeOutputPath + L"\\" + baseFilename + L".wav";
            dest = std::make_shared<WavFileDestination>();
        }
        else if (formatIndex == 1) {  // MP3
            config.outputPath = g_activeOutputPath + L"\\" + baseFilename + L".mp3";
            config.bitrate = g_activeBitrate;
            dest = std::make_shared<Mp3FileDestination>();
        }
        else if (formatIndex == 2) {  // Opus
            config.outputPath = g_activeOutputPath + L"\\" + baseFilename + L".opus";
            config.bitrate = g_activeBitrate;
            dest = std::make_shared<OpusFileDestination>();
        }
        else if (formatIndex == 3) {  // FLAC
            config.outputPath = g_activeOutputPath + L"\\" + baseFilename + L".flac";
            config.compressionLevel = g_activeFlacCompression;
            dest = std::make_shared<FlacFileDestination>();
        }

        if (dest && dest->Configure(formatCopy, config)) {
            return dest;
        }
        return nullptr;
    };

    // Create destinations with source-specific names
    // Use CURRENTLY checked outputs from UI, not g_activeFileFormats/g_activeDeviceIndices
    // (those only reflect what was checked at startup)
    int destCount = ListView_GetItemCount(g_hOutputDestsList);
    for (int i = 0; i < destCount; i++) {
        if (ListView_GetCheckState(g_hOutputDestsList, i)) {
            LVITEM lvi = {};
            lvi.mask = LVIF_PARAM;
            lvi.iItem = i;
            if (ListView_GetItem(g_hOutputDestsList, &lvi)) {
                int destType = GetDestinationType(lvi.lParam);
                int destIndex = GetDestinationIndex(lvi.lParam);

                if (destType == 0) {
                    // File format
                    std::wstring filename = sanitizedName + L"_capture";
                    auto dest = CreateDestination(destIndex, filename);
                    if (dest) {
                        sourceDestinations.push_back(dest);
                    }
                }
                else if (destType == 1 && includeDevices) {
                    // Audio device (only add if includeDevices=true to avoid duplicates)
                    if (destIndex >= 0 && destIndex < static_cast<int>(g_availableOutputDevices.size())) {
                        OutputDestinationPtr dest = std::make_shared<DeviceOutputDestination>();
                        DestinationConfig config;
                        config.outputPath = g_availableOutputDevices[destIndex].deviceId;
                        config.friendlyName = g_availableOutputDevices[destIndex].friendlyName;
                        if (dest->Configure(formatCopy, config)) {
                            sourceDestinations.push_back(dest);
                        }
                    }
                }
            }
        }
    }

    if (sourceDestinations.empty()) {
        return 0;  // No destinations created
    }

    // Create session for this source
    CaptureConfig config;
    config.sources.push_back(source);
    config.destinations = sourceDestinations;

    UINT32 sessionId = g_captureManager->StartCaptureSession(config);
    if (sessionId != 0) {
        // Track which source is in this per-source session
        g_sessionToSources[sessionId] = { source->GetMetadata().id };

        // Store destinations for real-time control
        for (const auto& dest : sourceDestinations) {
            g_activeDestinations[dest->GetName()] = dest;
        }
    }

    return sessionId;
}

// Start capturing
void StartCapture() {
    if (g_isCapturing) {
        return;
    }

    // Get output path
    wchar_t pathBuf[MAX_PATH];
    GetWindowText(g_hOutputPath, pathBuf, MAX_PATH);
    std::wstring outputPath = pathBuf;

    if (outputPath.empty()) {
        MessageBox(g_hWnd, L"Please specify an output folder", L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Collect indices of checked input sources (not the sources themselves yet)
    std::vector<int> checkedSourceIndices;
    int itemCount = ListView_GetItemCount(g_hInputSourcesList);

    for (int i = 0; i < itemCount; i++) {
        if (ListView_GetCheckState(g_hInputSourcesList, i)) {
            if (i < static_cast<int>(g_availableSources.size())) {
                checkedSourceIndices.push_back(i);
            }
        }
    }

    if (checkedSourceIndices.empty()) {
        MessageBox(g_hWnd, L"Please select at least one input source", L"Error", MB_OK | MB_ICONWARNING);
        return;
    }

    // Get item count for output destinations list
    itemCount = ListView_GetItemCount(g_hOutputDestsList);

    // Check if first source is a process and if we have cached format for it
    DWORD firstSourcePid = 0;
    bool useCachedFormat = false;
    InputSourcePtr tempSource;  // Will be null if using cached format

    if (g_availableSources[checkedSourceIndices[0]].metadata.type == InputSourceType::Process) {
        firstSourcePid = g_availableSources[checkedSourceIndices[0]].metadata.processId;

        // WORKAROUND: Check if we have cached format for this PID (Windows 10 Build 19044 bug)
        if (firstSourcePid == g_cachedFormatPid && !g_cachedFormatForPid.empty()) {
            useCachedFormat = true;
        }
    }

    std::vector<BYTE> formatBuffer;
    const WAVEFORMATEX* formatCopy = nullptr;

    if (useCachedFormat) {
        // Use cached format - no need to create tempSource!
        formatBuffer = g_cachedFormatForPid;
        formatCopy = reinterpret_cast<const WAVEFORMATEX*>(formatBuffer.data());
    } else {
        // Create first source temporarily to get format
        // CRITICAL: In Mode 2, this source will be reused to avoid duplicate process loopback initialization
        tempSource = g_sourceManager->CreateSource(g_availableSources[checkedSourceIndices[0]]);
        if (!tempSource || !tempSource->StartCapture()) {
            MessageBox(g_hWnd, L"Failed to initialize audio source", L"Error", MB_OK | MB_ICONERROR);
            return;
        }

        // Get format from first source (now available after StartCapture)
        const WAVEFORMATEX* format = tempSource->GetFormat();
        if (!format) {
            tempSource->StopCapture();
            MessageBox(g_hWnd, L"Failed to get audio format", L"Error", MB_OK | MB_ICONERROR);
            return;
        }

        // Make a copy of the format structure (including extra bytes for WAVEFORMATEXTENSIBLE)
        // so it remains valid after we destroy the temp source
        UINT32 formatSize = sizeof(WAVEFORMATEX) + format->cbSize;
        formatBuffer.resize(formatSize);
        memcpy(formatBuffer.data(), format, formatSize);
        formatCopy = reinterpret_cast<const WAVEFORMATEX*>(formatBuffer.data());

        // Cache the format for this PID (Windows 10 Build 19044 workaround)
        if (firstSourcePid != 0) {
            g_cachedFormatForPid = formatBuffer;
            g_cachedFormatPid = firstSourcePid;
        }

        // CRITICAL FIX: DON'T stop and reset tempSource yet!
        // In Mode 2, we need to reuse it to avoid duplicate process initialization
        // We'll stop it later if not reused
        // tempSource->StopCapture();
        // tempSource.reset();
    }

    // Get bitrate and FLAC compression from controls
    int bitrate = GetBitrate() * 1000;  // Convert kbps to bps
    int flacCompression = GetFlacCompression();

    // Get capture mode
    int captureMode = GetCaptureMode();

    // CRITICAL: Store capture mode and settings globally for dynamic source/dest addition
    g_activeCaptureMode = captureMode;
    g_activeOutputPath = outputPath;
    g_activeBitrate = bitrate;
    g_activeFlacCompression = flacCompression;
    g_activeCaptureFormat = formatBuffer;  // Store format for dynamic destination addition

    // Collect which file formats and devices are checked
    std::vector<int> checkedFileFormats;  // format indices 0-3 for WAV/MP3/Opus/FLAC
    std::vector<int> checkedDevices;      // device indices in g_availableOutputDevices

    for (int i = 0; i < itemCount; i++) {
        if (ListView_GetCheckState(g_hOutputDestsList, i)) {
            // Get the destination metadata from lParam
            LVITEM lvi = {};
            lvi.mask = LVIF_PARAM;
            lvi.iItem = i;
            if (ListView_GetItem(g_hOutputDestsList, &lvi)) {
                int destType = GetDestinationType(lvi.lParam);
                int destIndex = GetDestinationIndex(lvi.lParam);

                if (destType == 0) {  // File format
                    checkedFileFormats.push_back(destIndex);
                } else if (destType == 1) {  // Device
                    checkedDevices.push_back(destIndex);
                }
            }
        }
    }

    if (checkedFileFormats.empty() && checkedDevices.empty()) {
        // CRITICAL FIX: Must clean up tempSource before returning!
        if (tempSource) {
            tempSource->StopCapture();
            tempSource.reset();
            // Give Windows extra time to release the process loopback audio interface
            // Without this, the next Start attempt will fail with E_UNEXPECTED
            Sleep(500);
        }

        MessageBox(g_hWnd, L"Please select at least one output destination", L"Error", MB_OK | MB_ICONWARNING);
        return;
    }

    // Store checked formats and devices globally for dynamic addition
    g_activeFileFormats = checkedFileFormats;
    g_activeDeviceIndices = checkedDevices;

    // Lambda helper to create a destination based on format index and filename
    auto CreateDestination = [&](int formatIndex, const std::wstring& baseFilename) -> OutputDestinationPtr {
        OutputDestinationPtr dest;
        DestinationConfig config;
        config.useTimestamp = true;

        // Get skip silence setting from checkbox
        config.skipSilence = (SendMessage(g_hSkipSilenceCheck, BM_GETCHECK, 0, 0) == BST_CHECKED);
        config.silenceThreshold = 0.01f;  // 1% amplitude threshold
        config.silenceDurationMs = 1000;  // Skip after 1 second of silence

        if (formatIndex == 0) {  // WAV
            config.outputPath = outputPath + L"\\" + baseFilename + L".wav";
            dest = std::make_shared<WavFileDestination>();
        }
        else if (formatIndex == 1) {  // MP3
            config.outputPath = outputPath + L"\\" + baseFilename + L".mp3";
            config.bitrate = bitrate;
            dest = std::make_shared<Mp3FileDestination>();
        }
        else if (formatIndex == 2) {  // Opus
            config.outputPath = outputPath + L"\\" + baseFilename + L".opus";
            config.bitrate = bitrate;
            dest = std::make_shared<OpusFileDestination>();
        }
        else if (formatIndex == 3) {  // FLAC
            config.outputPath = outputPath + L"\\" + baseFilename + L".flac";
            config.compressionLevel = flacCompression;
            dest = std::make_shared<FlacFileDestination>();
        }

        if (dest && dest->Configure(formatCopy, config)) {
            return dest;
        }
        return nullptr;
    };

    // Lambda to create device destinations
    auto CreateDeviceDestinations = [&]() -> std::vector<OutputDestinationPtr> {
        std::vector<OutputDestinationPtr> devDests;
        for (int deviceIndex : checkedDevices) {
            if (deviceIndex >= 0 && deviceIndex < static_cast<int>(g_availableOutputDevices.size())) {
                OutputDestinationPtr dest = std::make_shared<DeviceOutputDestination>();
                DestinationConfig config;
                config.outputPath = g_availableOutputDevices[deviceIndex].deviceId;
                config.friendlyName = g_availableOutputDevices[deviceIndex].friendlyName;
                if (dest->Configure(formatCopy, config)) {
                    devDests.push_back(dest);
                }
            }
        }
        return devDests;
    };

    g_activeSessionIds.clear();
    g_activeSources.clear();  // Clear active source tracking
    g_activeDestinations.clear();  // Clear active destination tracking
    int totalDestinations = 0;

    // Mode 0: Single File (mixed) - all sources to one set of destinations
    if (captureMode == 0) {
        std::vector<OutputDestinationPtr> allDestinations;

        // Create one destination per checked format with "capture" name
        for (int formatIdx : checkedFileFormats) {
            auto dest = CreateDestination(formatIdx, L"capture");
            if (dest) {
                allDestinations.push_back(dest);
            }
        }

        // Add device destinations
        auto deviceDests = CreateDeviceDestinations();
        allDestinations.insert(allDestinations.end(), deviceDests.begin(), deviceDests.end());

        if (!allDestinations.empty()) {
            // Create fresh sources for this session
            std::vector<InputSourcePtr> sessionSources;
            for (int srcIdx : checkedSourceIndices) {
                InputSourcePtr source;

                // CRITICAL FIX: Reuse tempSource for the first source to avoid duplicate initialization
                if (srcIdx == checkedSourceIndices[0] && tempSource) {
                    source = tempSource;
                    tempSource.reset(); // Mark as consumed
                } else {
                    source = g_sourceManager->CreateSource(g_availableSources[srcIdx]);
                }

                if (source) {
                    // Apply volume setting
                    float volume = GetSourceVolume(source->GetMetadata().id);
                    source->SetVolume(volume);
                    sessionSources.push_back(source);
                    // Store for real-time control
                    g_activeSources[source->GetMetadata().id] = source;
                }
            }

            if (!sessionSources.empty()) {
                CaptureConfig config;
                config.sources = sessionSources;
                config.destinations = allDestinations;
                // No routing rules = all sources go to all destinations

                UINT32 sessionId = g_captureManager->StartCaptureSession(config);
                if (sessionId != 0) {
                    g_activeSessionIds.push_back(sessionId);
                    totalDestinations += static_cast<int>(allDestinations.size());

                    // Track which sources are in this session
                    std::vector<std::wstring> sourceIds;
                    for (const auto& src : sessionSources) {
                        sourceIds.push_back(src->GetMetadata().id);
                    }
                    g_sessionToSources[sessionId] = sourceIds;

                    // Store destinations for real-time control
                    for (const auto& dest : allDestinations) {
                        g_activeDestinations[dest->GetName()] = dest;
                    }
                }
            }
        }
    }
    // Mode 1: Multiple Files - separate files for each source
    else if (captureMode == 1) {
        // Create one session per source
        for (size_t i = 0; i < checkedSourceIndices.size(); i++) {
            int srcIdx = checkedSourceIndices[i];
            std::vector<OutputDestinationPtr> sourceDestinations;

            InputSourcePtr source;

            // CRITICAL FIX: Reuse tempSource for the first source to avoid duplicate initialization
            if (srcIdx == checkedSourceIndices[0] && tempSource) {
                source = tempSource;
                tempSource.reset(); // Mark as consumed
            } else {
                source = g_sourceManager->CreateSource(g_availableSources[srcIdx]);
            }

            if (!source) continue;

            // Apply volume setting
            float volume = GetSourceVolume(source->GetMetadata().id);
            source->SetVolume(volume);

            // Store for real-time control
            g_activeSources[source->GetMetadata().id] = source;

            // Get sanitized source name
            InputSourceMetadata metadata = source->GetMetadata();
            std::wstring sanitizedName = SanitizeFilename(metadata.displayName);

            // Create file destinations with source-specific names
            for (int formatIdx : checkedFileFormats) {
                std::wstring filename = sanitizedName + L"_capture";
                auto dest = CreateDestination(formatIdx, filename);
                if (dest) {
                    sourceDestinations.push_back(dest);
                }
            }

            // Add device destinations only to first source (avoid duplicates)
            if (i == 0) {
                auto deviceDests = CreateDeviceDestinations();
                sourceDestinations.insert(sourceDestinations.end(), deviceDests.begin(), deviceDests.end());
            }

            if (!sourceDestinations.empty()) {
                CaptureConfig config;
                config.sources.push_back(source);  // Only this fresh source
                config.destinations = sourceDestinations;

                UINT32 sessionId = g_captureManager->StartCaptureSession(config);
                if (sessionId != 0) {
                    g_activeSessionIds.push_back(sessionId);
                    totalDestinations += static_cast<int>(sourceDestinations.size());

                    // Track which source is in this per-source session
                    g_sessionToSources[sessionId] = { source->GetMetadata().id };

                    // Store destinations for real-time control
                    for (const auto& dest : sourceDestinations) {
                        g_activeDestinations[dest->GetName()] = dest;
                    }
                }
            }
        }
    }
    // Mode 2: Both Modes - mixed file + separate files per source
    // Solution: Create TWO sets of sessions with INDEPENDENT source instances:
    // 1. Mixed session with all sources → creates mixed file
    // 2. Per-source sessions with fresh source instances → creates per-source files
    else if (captureMode == 2) {
        // First: Create mixed session with all sources for the mixed file
        std::vector<InputSourcePtr> mixedSources;
        for (int srcIdx : checkedSourceIndices) {
            InputSourcePtr source;

            // CRITICAL: Reuse tempSource for the first source to avoid duplicate initialization
            if (srcIdx == checkedSourceIndices[0] && tempSource) {
                source = tempSource;
                tempSource.reset(); // Mark as consumed
            } else {
                source = g_sourceManager->CreateSource(g_availableSources[srcIdx]);
            }

            if (source) {
                // Apply volume setting
                float volume = GetSourceVolume(source->GetMetadata().id);
                source->SetVolume(volume);
                mixedSources.push_back(source);
                // Store for real-time control
                g_activeSources[source->GetMetadata().id] = source;
            }
        }

        if (!mixedSources.empty() && !checkedFileFormats.empty()) {
            // Create device destinations for monitoring
            std::vector<OutputDestinationPtr> mixedDestinations;
            auto deviceDests = CreateDeviceDestinations();
            mixedDestinations.insert(mixedDestinations.end(), deviceDests.begin(), deviceDests.end());

            CaptureConfig config;
            config.sources = mixedSources;
            config.destinations = mixedDestinations;  // Device destinations only

            // Enable mixed output to create the mixed file
            config.enableMixedOutput = true;
            config.mixedOutputPath = L"capture";

            // Determine format from first checked format
            int firstFormat = checkedFileFormats[0];
            if (firstFormat == 0) config.mixedOutputFormat = AudioFormat::WAV;
            else if (firstFormat == 1) config.mixedOutputFormat = AudioFormat::MP3;
            else if (firstFormat == 2) config.mixedOutputFormat = AudioFormat::OPUS;
            else if (firstFormat == 3) config.mixedOutputFormat = AudioFormat::FLAC;

            config.mixedOutputBitrate = 192000;

            UINT32 sessionId = g_captureManager->StartCaptureSession(config);
            if (sessionId != 0) {
                g_activeSessionIds.push_back(sessionId);
                totalDestinations += static_cast<int>(mixedDestinations.size()) + 1;  // +1 for mixed file

                // Track which sources are in this mixed session
                std::vector<std::wstring> sourceIds;
                for (const auto& src : mixedSources) {
                    sourceIds.push_back(src->GetMetadata().id);
                }
                g_sessionToSources[sessionId] = sourceIds;

                // Store destinations for real-time control
                for (const auto& dest : mixedDestinations) {
                    g_activeDestinations[dest->GetName()] = dest;
                }
            }
        }

        // Second: Create per-source sessions with FRESH source instances for per-source files
        for (size_t i = 0; i < checkedSourceIndices.size(); i++) {
            int srcIdx = checkedSourceIndices[i];
            std::vector<OutputDestinationPtr> sourceDestinations;

            // Create a FRESH source instance (do NOT reuse from mixed session!)
            InputSourcePtr source = g_sourceManager->CreateSource(g_availableSources[srcIdx]);
            if (!source) continue;

            // Apply volume setting
            float volume = GetSourceVolume(source->GetMetadata().id);
            source->SetVolume(volume);

            // Get sanitized source name
            InputSourceMetadata metadata = source->GetMetadata();
            std::wstring sanitizedName = SanitizeFilename(metadata.displayName);

            // Create file destinations with source-specific names
            for (int formatIdx : checkedFileFormats) {
                std::wstring filename = sanitizedName + L"_capture";
                auto dest = CreateDestination(formatIdx, filename);
                if (dest) {
                    sourceDestinations.push_back(dest);
                }
            }

            if (!sourceDestinations.empty()) {
                CaptureConfig config;
                config.sources.push_back(source);  // Only this fresh source
                config.destinations = sourceDestinations;

                UINT32 sessionId = g_captureManager->StartCaptureSession(config);
                if (sessionId != 0) {
                    g_activeSessionIds.push_back(sessionId);
                    totalDestinations += static_cast<int>(sourceDestinations.size());

                    // Track which source is in this per-source session
                    g_sessionToSources[sessionId] = { source->GetMetadata().id };

                    // Store destinations for real-time control
                    for (const auto& dest : sourceDestinations) {
                        g_activeDestinations[dest->GetName()] = dest;
                    }

                    // Store this fresh source (overwrites the mixed session source in g_activeSources)
                    // This is OK because we're using separate source instances
                    g_activeSources[source->GetMetadata().id] = source;
                }
            }
        }
    }

    // Clean up tempSource if it wasn't consumed by any mode
    if (tempSource) {
        tempSource->StopCapture();
        tempSource.reset();
    }

    if (g_activeSessionIds.empty()) {
        // CRITICAL FIX: Must clean up tempSource if no sessions were started
        if (tempSource) {
            tempSource->StopCapture();
            tempSource.reset();
            // Give Windows extra time to release the process loopback audio interface
            Sleep(500);
        }

        MessageBox(g_hWnd, L"Failed to start capture sessions", L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    g_isCapturing = true;
    SetWindowText(g_hStartStopBtn, L"Stop");

    // Hide capture mode controls during capture
    ShowWindow(g_hCaptureModeGroup, SW_HIDE);
    ShowWindow(g_hRadioSingleFile, SW_HIDE);
    ShowWindow(g_hRadioMultipleFiles, SW_HIDE);
    ShowWindow(g_hRadioBothModes, SW_HIDE);

    // Show pause/resume button only if there are file destinations
    if (!checkedFileFormats.empty()) {
        ShowWindow(g_hPauseResumeBtn, SW_SHOW);
        SetWindowText(g_hPauseResumeBtn, L"&Pause");
    } else {
        ShowWindow(g_hPauseResumeBtn, SW_HIDE);
    }

    UpdateStatus(L"Capturing " + std::to_wstring(checkedSourceIndices.size()) + L" source(s) to " +
                 std::to_wstring(totalDestinations) + L" destination(s) [" +
                 std::to_wstring(g_activeSessionIds.size()) + L" sessions]");
}

// Stop capturing
void StopCapture() {
    if (!g_isCapturing) {
        return;
    }

    // CRITICAL FIX: Track if we're stopping a process-loopback capture
    // We need to record the PID and timestamp so StartCapture can wait if needed
    bool hadProcessLoopback = false;
    DWORD processLoopbackPid = 0;
    for (const auto& [sourceId, source] : g_activeSources) {
        auto metadata = source->GetMetadata();
        if (metadata.type == InputSourceType::Process && metadata.processId != 0) {
            hadProcessLoopback = true;
            processLoopbackPid = metadata.processId;
            break;  // Only need to find one
        }
    }

    // Stop all active sessions
    for (UINT32 sessionId : g_activeSessionIds) {
        if (sessionId != 0) {
            g_captureManager->StopCaptureSession(sessionId);
        }
    }
    g_activeSessionIds.clear();
    g_activeSources.clear();  // Clear active source tracking
    g_activeDestinations.clear();  // Clear active destination tracking
    g_sessionToSources.clear();  // Clear session-to-source mapping

    // CRITICAL FIX: Give Windows time to fully release process loopback audio resources
    // Without this delay, immediately clicking Start again will fail because Windows
    // hasn't finished releasing the IAudioClient from the previous session
    // Windows 10 Build 19044 needs more time than newer builds
    Sleep(500);

    // Record when we stopped this process loopback capture
    if (hadProcessLoopback) {
        g_lastProcessLoopbackPid = processLoopbackPid;
        g_lastProcessLoopbackStopTime = GetTickCount();
    }

    // Clear stored capture mode and settings
    g_activeCaptureMode = -1;
    g_activeFileFormats.clear();
    g_activeDeviceIndices.clear();
    g_activeOutputPath.clear();
    g_activeBitrate = 0;
    g_activeFlacCompression = 0;
    g_activeCaptureFormat.clear();

    g_isCapturing = false;
    SetWindowText(g_hStartStopBtn, L"Start");

    // Hide pause/resume button when capture stops and reset state
    ShowWindow(g_hPauseResumeBtn, SW_HIDE);
    SetWindowText(g_hPauseResumeBtn, L"&Pause");  // Reset text
    g_isFilesPaused = false;  // Reset pause state

    // Show capture mode controls again after capture stops
    ShowWindow(g_hCaptureModeGroup, SW_SHOW);
    ShowWindow(g_hRadioSingleFile, SW_SHOW);
    ShowWindow(g_hRadioMultipleFiles, SW_SHOW);
    ShowWindow(g_hRadioBothModes, SW_SHOW);

    UpdateStatus(L"Capture stopped. Ready.");
}

// Browse for output folder
void BrowseOutputFolder() {
    BROWSEINFO bi = {};
    bi.hwndOwner = g_hWnd;
    bi.lpszTitle = L"Select Output Folder";
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;

    LPITEMIDLIST pidl = SHBrowseForFolder(&bi);
    if (pidl != nullptr) {
        wchar_t path[MAX_PATH];
        if (SHGetPathFromIDList(pidl, path)) {
            SetWindowText(g_hOutputPath, path);
        }
        CoTaskMemFree(pidl);
    }
}

// Update status text
void UpdateStatus(const std::wstring& message) {
    SetWindowText(g_hStatusText, message.c_str());
}

// Get default output path
std::wstring GetDefaultOutputPath() {
    wchar_t path[MAX_PATH];
    if (SHGetFolderPath(nullptr, CSIDL_MYDOCUMENTS, nullptr, 0, path) == S_OK) {
        return std::wstring(path) + L"\\AudioCapture";
    }
    return L"C:\\AudioCapture";
}

// Helper function to convert wstring to UTF-8 string
std::string WStringToUTF8(const std::wstring& wstr) {
    if (wstr.empty()) return std::string();

    int size_needed = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.size(), nullptr, 0, nullptr, nullptr);
    if (size_needed <= 0) return std::string();

    std::string result(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.size(), &result[0], size_needed, nullptr, nullptr);
    return result;
}

// Helper function to convert UTF-8 string to wstring
std::wstring UTF8ToWString(const std::string& str) {
    if (str.empty()) return std::wstring();

    int size_needed = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), (int)str.size(), nullptr, 0);
    if (size_needed <= 0) return std::wstring();

    std::wstring result(size_needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), (int)str.size(), &result[0], size_needed);
    return result;
}

// Get settings file path (AppData\Local\AudioCapture\settings.json)
std::wstring GetSettingsPath() {
    wchar_t path[MAX_PATH];
    if (SHGetFolderPath(nullptr, CSIDL_LOCAL_APPDATA, nullptr, 0, path) == S_OK) {
        std::wstring settingsDir = std::wstring(path) + L"\\AudioCapture";

        // Create directory if it doesn't exist
        // Check if it already exists first to avoid error on success
        DWORD attrib = GetFileAttributesW(settingsDir.c_str());
        if (attrib == INVALID_FILE_ATTRIBUTES || !(attrib & FILE_ATTRIBUTE_DIRECTORY)) {
            if (!CreateDirectoryW(settingsDir.c_str(), nullptr)) {
                DWORD error = GetLastError();
                if (error != ERROR_ALREADY_EXISTS) {
                    // Failed to create directory - log error
                    OutputDebugStringW((L"Failed to create settings directory: " + std::to_wstring(error) + L"\n").c_str());
                }
            }
        }

        return settingsDir + L"\\settings.json";
    }
    return L"settings.json";  // Fallback to current directory
}

// Load settings from JSON
void LoadSettings() {
    try {
        std::wstring settingsPath = GetSettingsPath();
        std::string pathStr = WStringToUTF8(settingsPath);

        std::ifstream file(pathStr, std::ios::binary);
        if (!file.is_open()) {
            // Settings file doesn't exist yet - this is normal on first run
            OutputDebugStringW(L"Settings file not found (first run?)\n");
            return;
        }

        json j;
        try {
            file >> j;
        }
        catch (const std::exception& e) {
            OutputDebugStringA(("Failed to parse settings JSON: " + std::string(e.what()) + "\n").c_str());
            return;
        }

        if (j.contains("outputPath")) {
            std::string outputPathUtf8 = j["outputPath"].get<std::string>();
            std::wstring path = UTF8ToWString(outputPathUtf8);
            SetWindowText(g_hOutputPath, path.c_str());
        }

        if (j.contains("bitrate")) {
            int bitrate = j["bitrate"].get<int>();
            SetWindowText(g_hBitrateEdit, std::to_wstring(bitrate).c_str());
            SendMessage(g_hBitrateSpin, UDM_SETPOS, 0, bitrate);
        }

        if (j.contains("flacCompression")) {
            int flacLevel = j["flacCompression"].get<int>();
            SetWindowText(g_hFlacCompressionEdit, std::to_wstring(flacLevel).c_str());
            SendMessage(g_hFlacCompressionSpin, UDM_SETPOS, 0, flacLevel);
        }

        if (j.contains("captureMode")) {
            int mode = j["captureMode"].get<int>();
            switch (mode) {
            case 0:
                SendMessage(g_hRadioSingleFile, BM_SETCHECK, BST_CHECKED, 0);
                SendMessage(g_hRadioMultipleFiles, BM_SETCHECK, BST_UNCHECKED, 0);
                SendMessage(g_hRadioBothModes, BM_SETCHECK, BST_UNCHECKED, 0);
                break;
            case 1:
                SendMessage(g_hRadioSingleFile, BM_SETCHECK, BST_UNCHECKED, 0);
                SendMessage(g_hRadioMultipleFiles, BM_SETCHECK, BST_CHECKED, 0);
                SendMessage(g_hRadioBothModes, BM_SETCHECK, BST_UNCHECKED, 0);
                break;
            case 2:
                SendMessage(g_hRadioSingleFile, BM_SETCHECK, BST_UNCHECKED, 0);
                SendMessage(g_hRadioMultipleFiles, BM_SETCHECK, BST_UNCHECKED, 0);
                SendMessage(g_hRadioBothModes, BM_SETCHECK, BST_CHECKED, 0);
                break;
            }
        }

        // Load source volumes
        if (j.contains("sourceVolumes") && j["sourceVolumes"].is_object()) {
            g_sourceVolumes.clear();
            for (auto& [key, value] : j["sourceVolumes"].items()) {
                try {
                    // Convert UTF-8 key to wide string
                    std::wstring sourceId = UTF8ToWString(key);
                    float volume = value.get<float>();

                    // Validate volume range
                    if (volume < 0.0f) volume = 0.0f;
                    if (volume > 1.0f) volume = 1.0f;

                    g_sourceVolumes[sourceId] = volume;
                }
                catch (const std::exception& e) {
                    OutputDebugStringA(("Error loading volume for source: " + std::string(e.what()) + "\n").c_str());
                }
            }
        }

        OutputDebugStringW(L"Settings loaded successfully\n");
    }
    catch (const std::exception& e) {
        OutputDebugStringA(("Error loading settings: " + std::string(e.what()) + "\n").c_str());
    }
    catch (...) {
        OutputDebugStringW(L"Unknown error loading settings\n");
    }
}

// Save settings to JSON
void SaveSettings() {
    try {
        json j;

        // Get output path and convert to UTF-8
        wchar_t pathBuf[MAX_PATH];
        GetWindowText(g_hOutputPath, pathBuf, MAX_PATH);
        std::wstring wPath = pathBuf;
        std::string pathUtf8 = WStringToUTF8(wPath);
        j["outputPath"] = pathUtf8;

        j["bitrate"] = GetBitrate();
        j["flacCompression"] = GetFlacCompression();
        j["captureMode"] = GetCaptureMode();

        // Save source volumes as JSON object (UTF-8 keys, float values)
        json volumesObj = json::object();
        for (const auto& [sourceId, volume] : g_sourceVolumes) {
            // Convert wide string source ID to UTF-8
            std::string sourceIdUtf8 = WStringToUTF8(sourceId);
            volumesObj[sourceIdUtf8] = volume;
        }
        j["sourceVolumes"] = volumesObj;

        // Get settings path and convert to UTF-8
        std::wstring settingsPath = GetSettingsPath();
        std::string settingsPathUtf8 = WStringToUTF8(settingsPath);

        // Open file in binary mode with explicit truncation
        std::ofstream file(settingsPathUtf8, std::ios::binary | std::ios::trunc);
        if (!file.is_open()) {
            DWORD error = GetLastError();
            OutputDebugStringW((L"Failed to open settings file for writing: " + std::to_wstring(error) + L"\n").c_str());
            return;
        }

        // Write JSON and explicitly flush
        std::string jsonStr = j.dump(4);
        file.write(jsonStr.c_str(), jsonStr.size());

        if (!file.good()) {
            OutputDebugStringW(L"Error writing settings JSON\n");
            return;
        }

        // Explicit flush and close
        file.flush();
        file.close();

        if (file.fail()) {
            OutputDebugStringW(L"Error flushing/closing settings file\n");
            return;
        }

        OutputDebugStringW(L"Settings saved successfully\n");
    }
    catch (const std::exception& e) {
        OutputDebugStringA(("Error saving settings: " + std::string(e.what()) + "\n").c_str());
    }
    catch (...) {
        OutputDebugStringW(L"Unknown error saving settings\n");
    }
}

// Get bitrate from control (in kbps)
int GetBitrate() {
    wchar_t buf[16];
    GetWindowText(g_hBitrateEdit, buf, 16);
    int bitrate = _wtoi(buf);
    if (bitrate < 64) bitrate = 64;
    if (bitrate > 320) bitrate = 320;
    return bitrate;
}

// Get FLAC compression level from control (0-8)
int GetFlacCompression() {
    wchar_t buf[16];
    GetWindowText(g_hFlacCompressionEdit, buf, 16);
    int level = _wtoi(buf);
    if (level < 0) level = 0;
    if (level > 8) level = 8;
    return level;
}

// Update visibility of bitrate and FLAC controls based on selected outputs
void UpdateControlVisibility() {
    int itemCount = ListView_GetItemCount(g_hOutputDestsList);
    bool showBitrate = false;
    bool showFlac = false;

    // Check which output formats are selected
    for (int i = 0; i < itemCount; i++) {
        if (ListView_GetCheckState(g_hOutputDestsList, i)) {
            // Get the destination metadata from lParam
            LVITEM lvi = {};
            lvi.mask = LVIF_PARAM;
            lvi.iItem = i;
            if (ListView_GetItem(g_hOutputDestsList, &lvi)) {
                int destType = GetDestinationType(lvi.lParam);
                int destIndex = GetDestinationIndex(lvi.lParam);

                if (destType == 0) {  // File format
                    if (destIndex == 1 || destIndex == 2) {  // MP3 or Opus
                        showBitrate = true;
                    }
                    if (destIndex == 3) {  // FLAC
                        showFlac = true;
                    }
                }
            }
        }
    }

    // Show/hide bitrate controls
    int bitrateShow = showBitrate ? SW_SHOW : SW_HIDE;
    ShowWindow(GetDlgItem(g_hWnd, 0x1000), bitrateShow);  // Static label
    ShowWindow(g_hBitrateEdit, bitrateShow);
    ShowWindow(g_hBitrateSpin, bitrateShow);

    // Show/hide FLAC controls
    int flacShow = showFlac ? SW_SHOW : SW_HIDE;
    ShowWindow(GetDlgItem(g_hWnd, 0x1001), flacShow);  // Static label
    ShowWindow(g_hFlacCompressionEdit, flacShow);
    ShowWindow(g_hFlacCompressionSpin, flacShow);
}

// Get selected capture mode
// Returns: 0 = Single File, 1 = Multiple Files, 2 = Both Modes
int GetCaptureMode() {
    if (SendMessage(g_hRadioSingleFile, BM_GETCHECK, 0, 0) == BST_CHECKED) {
        return 0;  // Single File Mode
    }
    else if (SendMessage(g_hRadioMultipleFiles, BM_GETCHECK, 0, 0) == BST_CHECKED) {
        return 1;  // Multiple Files Mode
    }
    else if (SendMessage(g_hRadioBothModes, BM_GETCHECK, 0, 0) == BST_CHECKED) {
        return 2;  // Both Modes
    }
    return 0;  // Default to Single File
}

// Sanitize source name for safe filename usage
std::wstring SanitizeFilename(const std::wstring& displayName) {
    std::wstring sanitized = displayName;

    // Remove common bracketed markers like [Input], [Output], [Loopback], etc.
    size_t pos = 0;
    while ((pos = sanitized.find(L'[')) != std::wstring::npos) {
        size_t endPos = sanitized.find(L']', pos);
        if (endPos != std::wstring::npos) {
            sanitized.erase(pos, endPos - pos + 1);
        } else {
            break;
        }
    }

    // Replace invalid filename characters with underscores
    const wchar_t invalidChars[] = L"<>:\"/\\|?*";
    for (wchar_t c : invalidChars) {
        std::replace(sanitized.begin(), sanitized.end(), c, L'_');
    }

    // Also replace control characters (0-31) and some other problematic chars
    for (size_t i = 0; i < sanitized.length(); i++) {
        if (sanitized[i] < 32 || sanitized[i] == 127) {
            sanitized[i] = L'_';
        }
    }

    // Trim leading/trailing whitespace and dots (Windows doesn't like them)
    size_t start = 0;
    size_t end = sanitized.length();

    while (start < end && (sanitized[start] == L' ' || sanitized[start] == L'\t' || sanitized[start] == L'.')) {
        start++;
    }

    while (end > start && (sanitized[end-1] == L' ' || sanitized[end-1] == L'\t' || sanitized[end-1] == L'.')) {
        end--;
    }

    sanitized = sanitized.substr(start, end - start);

    // Ensure we have a valid name (not empty)
    if (sanitized.empty()) {
        sanitized = L"capture";
    }

    // Truncate if too long (Windows has 255 char limit, leave room for path and extension)
    if (sanitized.length() > 100) {
        sanitized = sanitized.substr(0, 100);
    }

    return sanitized;
}

// Update volume controls based on focused input source
void UpdateVolumeControls() {
    // Get focused item from input sources list
    int focusedIndex = ListView_GetNextItem(g_hInputSourcesList, -1, LVNI_FOCUSED);

    // Check if item is valid and checked
    if (focusedIndex >= 0 && focusedIndex < static_cast<int>(g_availableSources.size())) {
        BOOL isChecked = ListView_GetCheckState(g_hInputSourcesList, focusedIndex);

        if (isChecked) {
            // Get source information
            const auto& source = g_availableSources[focusedIndex];

            // Show volume controls (not showing g_hVolumeValue anymore - percentage is in label)
            ShowWindow(g_hVolumeLabel, SW_SHOW);
            ShowWindow(g_hVolumeSlider, SW_SHOW);

            // Get volume from map (default 100 if not found)
            float volume = GetSourceVolume(source.metadata.id);
            int sliderPos = static_cast<int>(volume * 100.0f);

            // Clamp to valid range
            if (sliderPos < 0) sliderPos = 0;
            if (sliderPos > 100) sliderPos = 100;

            // Update label with source name AND percentage
            std::wstring labelText = L"Volume for: " + source.metadata.displayName + L" (" + std::to_wstring(sliderPos) + L"%)";
            SetWindowText(g_hVolumeLabel, labelText.c_str());

            // Update slider position
            SendMessage(g_hVolumeSlider, TBM_SETPOS, TRUE, sliderPos);
        } else {
            // Not checked - hide controls and clear text
            SetWindowText(g_hVolumeLabel, L"");
            ShowWindow(g_hVolumeLabel, SW_HIDE);
            ShowWindow(g_hVolumeSlider, SW_HIDE);
        }
    } else {
        // No valid selection - hide controls and clear text
        SetWindowText(g_hVolumeLabel, L"");
        ShowWindow(g_hVolumeLabel, SW_HIDE);
        ShowWindow(g_hVolumeSlider, SW_HIDE);
    }
}

// Handle volume slider changes
void OnVolumeSliderChanged() {
    // Get focused item from input sources list
    int focusedIndex = ListView_GetNextItem(g_hInputSourcesList, -1, LVNI_FOCUSED);

    // Check if item is valid and checked
    if (focusedIndex >= 0 && focusedIndex < static_cast<int>(g_availableSources.size())) {
        BOOL isChecked = ListView_GetCheckState(g_hInputSourcesList, focusedIndex);

        if (isChecked) {
            // Get slider position (0-100)
            int sliderPos = static_cast<int>(SendMessage(g_hVolumeSlider, TBM_GETPOS, 0, 0));

            // Clamp to valid range
            if (sliderPos < 0) sliderPos = 0;
            if (sliderPos > 100) sliderPos = 100;

            // Convert to float (0.0-1.0)
            float volume = sliderPos / 100.0f;

            // Get source ID and update volume
            const auto& source = g_availableSources[focusedIndex];
            SetSourceVolume(source.metadata.id, volume);

            // Update label with source name AND percentage
            std::wstring labelText = L"Volume for: " + source.metadata.displayName + L" (" + std::to_wstring(sliderPos) + L"%)";
            SetWindowText(g_hVolumeLabel, labelText.c_str());

            // If capturing, apply volume change to active sources in real-time
            if (g_isCapturing) {
                auto it = g_activeSources.find(source.metadata.id);
                if (it != g_activeSources.end() && it->second) {
                    it->second->SetVolume(volume);
                }
            }
        }
    }
}

// Get volume for a source (default 1.0 if not found)
float GetSourceVolume(const std::wstring& sourceId) {
    auto it = g_sourceVolumes.find(sourceId);
    if (it != g_sourceVolumes.end()) {
        return it->second;
    }
    return 1.0f;  // Default volume = 100%
}

// Set volume for a source
void SetSourceVolume(const std::wstring& sourceId, float volume) {
    // Clamp volume to valid range
    if (volume < 0.0f) volume = 0.0f;
    if (volume > 1.0f) volume = 1.0f;

    g_sourceVolumes[sourceId] = volume;
}
