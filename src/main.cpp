#include <windows.h>
#include <roapi.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <string>
#include <vector>
#include <fstream>
#include <nlohmann/json.hpp>
#include "resource.h"
#include "ProcessEnumerator.h"
#include "CaptureManager.h"
#include "AudioDeviceEnumerator.h"

using json = nlohmann::json;

#pragma comment(lib, "comctl32.lib")
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

// Global variables
HINSTANCE g_hInst;
HWND g_hWnd;
HWND g_hProcessList;
HWND g_hRefreshBtn;
HWND g_hStartBtn;
HWND g_hStopBtn;
HWND g_hStopAllBtn;
HWND g_hPauseAllBtn;
HWND g_hResumeAllBtn;
HWND g_hFormatCombo;
HWND g_hOutputPath;
HWND g_hBrowseBtn;
HWND g_hStatusText;
HWND g_hRecordingList;
HWND g_hMp3BitrateCombo;
HWND g_hOpusBitrateCombo;
HWND g_hMp3BitrateLabel;
HWND g_hOpusBitrateLabel;
HWND g_hProcessListLabel;
HWND g_hOutputPathLabel;
HWND g_hRecordingListLabel;
HWND g_hSkipSilenceCheckbox;
HWND g_hFlacCompressionCombo;
HWND g_hFlacCompressionLabel;
HWND g_hShowAudioOnlyCheckbox;
HWND g_hPassthroughCheckbox;
HWND g_hPassthroughDeviceCombo;
HWND g_hPassthroughDeviceLabel;
HWND g_hMonitorOnlyCheckbox;
HWND g_hRecordingModeCombo;
HWND g_hRecordingModeLabel;
HWND g_hMicrophoneCheckbox;
HWND g_hMicrophoneDeviceCombo;
HWND g_hMicrophoneDeviceLabel;

std::unique_ptr<ProcessEnumerator> g_processEnum;
std::unique_ptr<CaptureManager> g_captureManager;
std::unique_ptr<AudioDeviceEnumerator> g_audioDeviceEnum;
std::vector<ProcessInfo> g_processes;
bool g_useWinRT = false;  // Track whether we initialized with WinRT or COM

// Window class name
const wchar_t CLASS_NAME[] = L"AudioCaptureWindow";

// Forward declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void InitializeControls(HWND hwnd);
void RefreshProcessList();
void StartCapture();
void StopCapture();
void UpdateRecordingList();
void BrowseOutputFolder();
void OnFormatChanged();
std::wstring GetDefaultOutputPath();
std::wstring FormatFileSize(UINT64 bytes);
void LoadSettings();
void SaveSettings();
std::wstring GetSettingsFilePath();
std::string WStringToString(const std::wstring& wstr);
std::wstring StringToWString(const std::string& str);
void PopulatePassthroughDevices();
void OnPassthroughCheckboxChanged();
void OnMonitorOnlyCheckboxChanged();
void PopulateMicrophoneDevices();
void OnMicrophoneCheckboxChanged();

// Helper to detect Windows version
bool IsWindows8OrGreater() {
    typedef LONG (WINAPI* RtlGetVersionPtr)(PRTL_OSVERSIONINFOW);
    HMODULE hNtdll = GetModuleHandleW(L"ntdll.dll");
    if (!hNtdll) return false;

    RtlGetVersionPtr RtlGetVersion = (RtlGetVersionPtr)GetProcAddress(hNtdll, "RtlGetVersion");
    if (!RtlGetVersion) return false;

    RTL_OSVERSIONINFOW osvi = { 0 };
    osvi.dwOSVersionInfoSize = sizeof(osvi);
    if (RtlGetVersion(&osvi) == 0) {
        // Windows 8 is version 6.2, Windows 10 is 10.0
        return (osvi.dwMajorVersion > 6) || (osvi.dwMajorVersion == 6 && osvi.dwMinorVersion >= 2);
    }
    return false;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int nCmdShow) {
    g_hInst = hInstance;

    // Initialize COM/WinRT based on Windows version
    // Windows 8+ supports WinRT (RoInitialize)
    // Windows 7 requires traditional COM (CoInitializeEx)
    HRESULT hr;
    bool useWinRT = IsWindows8OrGreater();
    g_useWinRT = useWinRT;  // Store for cleanup later

    if (useWinRT) {
        // Initialize Windows Runtime (required for ActivateAudioInterfaceAsync on Windows 10+)
        // RoInitialize initializes both COM and WinRT
        // IMPORTANT: Must use RO_INIT_SINGLETHREADED for ActivateAudioInterfaceAsync to work
        hr = RoInitialize(RO_INIT_SINGLETHREADED);
        if (FAILED(hr) && hr != RPC_E_CHANGED_MODE && hr != S_FALSE) {
            char errMsg[256];
            sprintf_s(errMsg, "Failed to initialize Windows Runtime: 0x%08X", hr);
            MessageBoxA(nullptr, errMsg, "Error", MB_OK | MB_ICONERROR);
            return 0;
        }
    } else {
        // Windows 7 - use traditional COM initialization
        hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        if (FAILED(hr) && hr != RPC_E_CHANGED_MODE && hr != S_FALSE) {
            char errMsg[256];
            sprintf_s(errMsg, "Failed to initialize COM: 0x%08X", hr);
            MessageBoxA(nullptr, errMsg, "Error", MB_OK | MB_ICONERROR);
            return 0;
        }
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
    wc.hIcon = LoadIcon(nullptr, IDI_APPLICATION);

    RegisterClass(&wc);

    // Create window
    g_hWnd = CreateWindowEx(
        0,
        CLASS_NAME,
        L"Audio Capture - Per-Process Recording",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 800, 600,
        nullptr,
        nullptr,
        hInstance,
        nullptr
    );

    if (g_hWnd == nullptr) {
        return 0;
    }

    // Initialize global objects
    g_processEnum = std::make_unique<ProcessEnumerator>();
    g_captureManager = std::make_unique<CaptureManager>();
    g_audioDeviceEnum = std::make_unique<AudioDeviceEnumerator>();

    ShowWindow(g_hWnd, nCmdShow);
    UpdateWindow(g_hWnd);

    // Message loop with dialog message processing for tab navigation
    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0)) {
        if (!IsDialogMessage(g_hWnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    // Cleanup COM/Windows Runtime
    if (g_useWinRT) {
        RoUninitialize();
    } else {
        CoUninitialize();
    }

    return 0;
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE:
        InitializeControls(hwnd);
        LoadSettings();
        RefreshProcessList();
        // Set initial focus to the process list
        SetFocus(g_hProcessList);
        // Start timer for updating recording list (every 500ms)
        SetTimer(hwnd, 1, 500, nullptr);
        // Populate devices after window is fully created
        PostMessage(hwnd, WM_USER + 1, 0, 0);
        return 0;

    case WM_USER + 1:
        // Called after window is fully created
        PopulatePassthroughDevices();
        PopulateMicrophoneDevices();
        return 0;

    case WM_COMMAND: {
        int wmId = LOWORD(wParam);
        int wmEvent = HIWORD(wParam);

        switch (wmId) {
        case IDC_REFRESH_BTN:
            RefreshProcessList();
            break;

        case IDC_START_BTN:
            StartCapture();
            break;

        case IDC_STOP_BTN:
            StopCapture();
            break;

        case IDC_STOP_ALL_BTN:
            if (MessageBox(g_hWnd, L"Stop all active captures?", L"Confirm", MB_YESNO | MB_ICONQUESTION) == IDYES) {
                // Stop all individual captures FIRST (stops audio callbacks)
                g_captureManager->StopAllCaptures();
                // Then disable mixed recording (mixer thread can exit cleanly)
                g_captureManager->DisableMixedRecording();
                EnableWindow(g_hStopBtn, FALSE);
                EnableWindow(g_hStopAllBtn, FALSE);
                ShowWindow(g_hStopAllBtn, SW_HIDE);
                UpdateRecordingList();
                SetWindowText(g_hStatusText, L"All captures stopped.");
                SetFocus(g_hProcessList);
            }
            break;

        case IDC_PAUSE_ALL_BTN:
            g_captureManager->PauseAllCaptures();
            SetWindowText(g_hStatusText, L"All captures paused.");
            UpdateRecordingList();  // Update button states immediately
            SetFocus(g_hProcessList);
            break;

        case IDC_RESUME_ALL_BTN:
            g_captureManager->ResumeAllCaptures();
            SetWindowText(g_hStatusText, L"All captures resumed.");
            UpdateRecordingList();  // Update button states immediately
            SetFocus(g_hProcessList);
            break;

        case IDC_BROWSE_BTN:
            BrowseOutputFolder();
            break;

        case IDC_FORMAT_COMBO:
            if (wmEvent == CBN_SELCHANGE) {
                OnFormatChanged();
            }
            break;

        case IDC_SHOW_AUDIO_ONLY_CHECKBOX:
            if (wmEvent == BN_CLICKED) {
                RefreshProcessList();
            }
            break;

        case IDC_PASSTHROUGH_CHECKBOX:
            if (wmEvent == BN_CLICKED) {
                OnPassthroughCheckboxChanged();
            }
            break;

        case IDC_MONITOR_ONLY_CHECKBOX:
            if (wmEvent == BN_CLICKED) {
                OnMonitorOnlyCheckboxChanged();
            }
            break;

        case IDC_MICROPHONE_CHECKBOX:
            if (wmEvent == BN_CLICKED) {
                OnMicrophoneCheckboxChanged();
            }
            break;
        }
        return 0;
    }

    case WM_SIZE: {
        // Resize controls when window is resized
        RECT rcClient;
        GetClientRect(hwnd, &rcClient);
        int width = rcClient.right - rcClient.left;
        int height = rcClient.bottom - rcClient.top;

        // Adjust control positions and sizes
        SetWindowPos(g_hProcessListLabel, nullptr, 10, 10, 200, 20, SWP_NOZORDER);
        SetWindowPos(g_hProcessList, nullptr, 10, 30, width - 20, 180, SWP_NOZORDER);
        SetWindowPos(g_hRefreshBtn, nullptr, 10, 215, 100, 25, SWP_NOZORDER);
        SetWindowPos(g_hShowAudioOnlyCheckbox, nullptr, 120, 218, 280, 20, SWP_NOZORDER);
        SetWindowPos(g_hFormatCombo, nullptr, 10, 245, 100, 25, SWP_NOZORDER);
        SetWindowPos(g_hOutputPathLabel, nullptr, 10, 275, 100, 20, SWP_NOZORDER);
        SetWindowPos(g_hOutputPath, nullptr, 10, 295, width - 100, 25, SWP_NOZORDER);
        SetWindowPos(g_hBrowseBtn, nullptr, width - 85, 295, 75, 25, SWP_NOZORDER);
        SetWindowPos(g_hRecordingModeLabel, nullptr, 10, 330, 140, 20, SWP_NOZORDER);
        SetWindowPos(g_hRecordingModeCombo, nullptr, 150, 327, 150, 25, SWP_NOZORDER);
        SetWindowPos(g_hStartBtn, nullptr, 10, 360, 100, 30, SWP_NOZORDER);
        SetWindowPos(g_hStopBtn, nullptr, 120, 360, 100, 30, SWP_NOZORDER);
        SetWindowPos(g_hStopAllBtn, nullptr, 230, 360, 100, 30, SWP_NOZORDER);
        SetWindowPos(g_hPauseAllBtn, nullptr, 340, 360, 100, 30, SWP_NOZORDER);
        SetWindowPos(g_hResumeAllBtn, nullptr, 450, 360, 100, 30, SWP_NOZORDER);
        SetWindowPos(g_hRecordingListLabel, nullptr, 10, 400, 200, 20, SWP_NOZORDER);
        SetWindowPos(g_hRecordingList, nullptr, 10, 420, width - 20, height - 470, SWP_NOZORDER);
        SetWindowPos(g_hStatusText, nullptr, 10, height - 40, width - 20, 30, SWP_NOZORDER);
        return 0;
    }

    case WM_DESTROY:
        SaveSettings();
        g_captureManager->StopAllCaptures();
        PostQuitMessage(0);
        return 0;

    case WM_TIMER:
        if (wParam == 1) {
            // Update recording list periodically
            UpdateRecordingList();
        }
        return 0;
    }

    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

void InitializeControls(HWND hwnd) {
    // Process list label
    g_hProcessListLabel = CreateWindow(
        L"STATIC", L"Available Processes:",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        10, 10, 200, 20,
        hwnd, (HMENU)IDC_PROCESS_LIST_LABEL, g_hInst, nullptr
    );

    // Process list (ListView) - with checkboxes for multi-select
    g_hProcessList = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        WC_LISTVIEW,
        L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT,
        10, 30, 760, 180,
        hwnd,
        (HMENU)IDC_PROCESS_LIST,
        g_hInst,
        nullptr
    );

    // Enable checkboxes for the process list
    ListView_SetExtendedListViewStyle(g_hProcessList, LVS_EX_CHECKBOXES | LVS_EX_FULLROWSELECT);

    // Set up list view columns
    LVCOLUMN lvc;
    lvc.mask = LVCF_TEXT | LVCF_WIDTH;
    lvc.cx = 180;
    lvc.pszText = (LPWSTR)L"Process Name";
    ListView_InsertColumn(g_hProcessList, 0, &lvc);

    lvc.cx = 60;
    lvc.pszText = (LPWSTR)L"PID";
    ListView_InsertColumn(g_hProcessList, 1, &lvc);

    lvc.cx = 250;
    lvc.pszText = (LPWSTR)L"Window Title";
    ListView_InsertColumn(g_hProcessList, 2, &lvc);

    lvc.cx = 300;
    lvc.pszText = (LPWSTR)L"Path";
    ListView_InsertColumn(g_hProcessList, 3, &lvc);

    // Refresh button
    g_hRefreshBtn = CreateWindow(
        L"BUTTON", L"Refresh",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        10, 215, 100, 25,
        hwnd, (HMENU)IDC_REFRESH_BTN, g_hInst, nullptr
    );

    // Show audio only checkbox
    g_hShowAudioOnlyCheckbox = CreateWindow(
        L"BUTTON", L"Show only processes with active audio",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
        120, 218, 280, 20,
        hwnd, (HMENU)IDC_SHOW_AUDIO_ONLY_CHECKBOX, g_hInst, nullptr
    );

    // Format combo box
    g_hFormatCombo = CreateWindow(
        WC_COMBOBOX, L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
        10, 245, 100, 200,
        hwnd, (HMENU)IDC_FORMAT_COMBO, g_hInst, nullptr
    );
    SendMessage(g_hFormatCombo, CB_ADDSTRING, 0, (LPARAM)L"WAV");
    SendMessage(g_hFormatCombo, CB_ADDSTRING, 0, (LPARAM)L"MP3");
    SendMessage(g_hFormatCombo, CB_ADDSTRING, 0, (LPARAM)L"Opus");
    SendMessage(g_hFormatCombo, CB_ADDSTRING, 0, (LPARAM)L"FLAC");
    SendMessage(g_hFormatCombo, CB_SETCURSEL, 0, 0);

    // MP3 bitrate label
    g_hMp3BitrateLabel = CreateWindow(
        L"STATIC", L"MP3 Bitrate:",
        WS_CHILD | SS_LEFT,
        120, 248, 80, 20,
        hwnd, (HMENU)IDC_MP3_BITRATE_LABEL, g_hInst, nullptr
    );

    // MP3 bitrate combo box
    g_hMp3BitrateCombo = CreateWindow(
        WC_COMBOBOX, L"",
        WS_CHILD | WS_TABSTOP | CBS_DROPDOWNLIST,
        200, 245, 80, 200,
        hwnd, (HMENU)IDC_MP3_BITRATE_COMBO, g_hInst, nullptr
    );
    SendMessage(g_hMp3BitrateCombo, CB_ADDSTRING, 0, (LPARAM)L"128 kbps");
    SendMessage(g_hMp3BitrateCombo, CB_ADDSTRING, 0, (LPARAM)L"192 kbps");
    SendMessage(g_hMp3BitrateCombo, CB_ADDSTRING, 0, (LPARAM)L"256 kbps");
    SendMessage(g_hMp3BitrateCombo, CB_ADDSTRING, 0, (LPARAM)L"320 kbps");
    SendMessage(g_hMp3BitrateCombo, CB_SETCURSEL, 1, 0);  // Default to 192 kbps

    // Opus bitrate label
    g_hOpusBitrateLabel = CreateWindow(
        L"STATIC", L"Opus Bitrate:",
        WS_CHILD | SS_LEFT,
        120, 248, 80, 20,
        hwnd, (HMENU)IDC_OPUS_BITRATE_LABEL, g_hInst, nullptr
    );

    // Opus bitrate combo box
    g_hOpusBitrateCombo = CreateWindow(
        WC_COMBOBOX, L"",
        WS_CHILD | WS_TABSTOP | CBS_DROPDOWNLIST,
        200, 245, 80, 200,
        hwnd, (HMENU)IDC_OPUS_BITRATE_COMBO, g_hInst, nullptr
    );
    SendMessage(g_hOpusBitrateCombo, CB_ADDSTRING, 0, (LPARAM)L"64 kbps");
    SendMessage(g_hOpusBitrateCombo, CB_ADDSTRING, 0, (LPARAM)L"96 kbps");
    SendMessage(g_hOpusBitrateCombo, CB_ADDSTRING, 0, (LPARAM)L"128 kbps");
    SendMessage(g_hOpusBitrateCombo, CB_ADDSTRING, 0, (LPARAM)L"192 kbps");
    SendMessage(g_hOpusBitrateCombo, CB_ADDSTRING, 0, (LPARAM)L"256 kbps");
    SendMessage(g_hOpusBitrateCombo, CB_SETCURSEL, 2, 0);  // Default to 128 kbps

    // FLAC compression label
    g_hFlacCompressionLabel = CreateWindow(
        L"STATIC", L"FLAC Level:",
        WS_CHILD | SS_LEFT,
        120, 248, 80, 20,
        hwnd, (HMENU)IDC_FLAC_COMPRESSION_LABEL, g_hInst, nullptr
    );

    // FLAC compression combo box
    g_hFlacCompressionCombo = CreateWindow(
        WC_COMBOBOX, L"",
        WS_CHILD | WS_TABSTOP | CBS_DROPDOWNLIST,
        200, 245, 80, 200,
        hwnd, (HMENU)IDC_FLAC_COMPRESSION_COMBO, g_hInst, nullptr
    );
    SendMessage(g_hFlacCompressionCombo, CB_ADDSTRING, 0, (LPARAM)L"0 (Fast)");
    SendMessage(g_hFlacCompressionCombo, CB_ADDSTRING, 0, (LPARAM)L"1");
    SendMessage(g_hFlacCompressionCombo, CB_ADDSTRING, 0, (LPARAM)L"2");
    SendMessage(g_hFlacCompressionCombo, CB_ADDSTRING, 0, (LPARAM)L"3");
    SendMessage(g_hFlacCompressionCombo, CB_ADDSTRING, 0, (LPARAM)L"4");
    SendMessage(g_hFlacCompressionCombo, CB_ADDSTRING, 0, (LPARAM)L"5 (Default)");
    SendMessage(g_hFlacCompressionCombo, CB_ADDSTRING, 0, (LPARAM)L"6");
    SendMessage(g_hFlacCompressionCombo, CB_ADDSTRING, 0, (LPARAM)L"7");
    SendMessage(g_hFlacCompressionCombo, CB_ADDSTRING, 0, (LPARAM)L"8 (Best)");
    SendMessage(g_hFlacCompressionCombo, CB_SETCURSEL, 5, 0);  // Default to level 5

    // Skip silence checkbox
    g_hSkipSilenceCheckbox = CreateWindow(
        L"BUTTON", L"Skip silence",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
        290, 247, 120, 20,
        hwnd, (HMENU)IDC_SKIP_SILENCE_CHECKBOX, g_hInst, nullptr
    );

    // Passthrough checkbox
    g_hPassthroughCheckbox = CreateWindow(
        L"BUTTON", L"Monitor audio",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
        420, 247, 120, 20,
        hwnd, (HMENU)IDC_PASSTHROUGH_CHECKBOX, g_hInst, nullptr
    );

    // Passthrough device label
    g_hPassthroughDeviceLabel = CreateWindow(
        L"STATIC", L"Monitor Device:",
        WS_CHILD | SS_LEFT,
        550, 248, 100, 20,
        hwnd, (HMENU)IDC_PASSTHROUGH_DEVICE_LABEL, g_hInst, nullptr
    );

    // Passthrough device combo box
    g_hPassthroughDeviceCombo = CreateWindow(
        WC_COMBOBOX, L"",
        WS_CHILD | WS_TABSTOP | CBS_DROPDOWNLIST,
        650, 245, 140, 200,
        hwnd, (HMENU)IDC_PASSTHROUGH_DEVICE_COMBO, g_hInst, nullptr
    );

    // Monitor only checkbox
    g_hMonitorOnlyCheckbox = CreateWindow(
        L"BUTTON", L"Monitor only (no recording)",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
        420, 270, 200, 20,
        hwnd, (HMENU)IDC_MONITOR_ONLY_CHECKBOX, g_hInst, nullptr
    );

    // Recording mode label
    g_hRecordingModeLabel = CreateWindow(
        L"STATIC", L"Multi-process recording:",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        10, 330, 140, 20,
        hwnd, (HMENU)IDC_RECORDING_MODE_LABEL, g_hInst, nullptr
    );

    // Recording mode combo box
    g_hRecordingModeCombo = CreateWindow(
        WC_COMBOBOX, L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
        150, 327, 150, 200,
        hwnd, (HMENU)IDC_RECORDING_MODE_COMBO, g_hInst, nullptr
    );
    SendMessage(g_hRecordingModeCombo, CB_ADDSTRING, 0, (LPARAM)L"Separate files");
    SendMessage(g_hRecordingModeCombo, CB_ADDSTRING, 0, (LPARAM)L"Combined file");
    SendMessage(g_hRecordingModeCombo, CB_ADDSTRING, 0, (LPARAM)L"Both");
    SendMessage(g_hRecordingModeCombo, CB_SETCURSEL, 0, 0);  // Default to separate files

    // Microphone capture checkbox
    g_hMicrophoneCheckbox = CreateWindow(
        L"BUTTON", L"Capture microphone",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
        310, 330, 140, 20,
        hwnd, (HMENU)IDC_MICROPHONE_CHECKBOX, g_hInst, nullptr
    );

    // Microphone device label
    g_hMicrophoneDeviceLabel = CreateWindow(
        L"STATIC", L"Microphone:",
        WS_CHILD | SS_LEFT,
        460, 333, 80, 20,
        hwnd, (HMENU)IDC_MICROPHONE_DEVICE_LABEL, g_hInst, nullptr
    );

    // Microphone device combo box
    g_hMicrophoneDeviceCombo = CreateWindow(
        WC_COMBOBOX, L"",
        WS_CHILD | WS_TABSTOP | CBS_DROPDOWNLIST,
        540, 327, 230, 200,
        hwnd, (HMENU)IDC_MICROPHONE_DEVICE_COMBO, g_hInst, nullptr
    );

    // Output path label
    g_hOutputPathLabel = CreateWindow(
        L"STATIC", L"Output Folder:",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        10, 275, 100, 20,
        hwnd, (HMENU)IDC_OUTPUT_PATH_LABEL, g_hInst, nullptr
    );

    // Output path edit
    g_hOutputPath = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        L"EDIT",
        GetDefaultOutputPath().c_str(),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_LEFT | ES_AUTOHSCROLL,
        10, 295, 680, 25,
        hwnd, (HMENU)IDC_OUTPUT_PATH, g_hInst, nullptr
    );

    // Browse button
    g_hBrowseBtn = CreateWindow(
        L"BUTTON", L"Browse...",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        700, 295, 75, 25,
        hwnd, (HMENU)IDC_BROWSE_BTN, g_hInst, nullptr
    );

    // Start button
    g_hStartBtn = CreateWindow(
        L"BUTTON", L"Start Capture",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        10, 360, 100, 30,
        hwnd, (HMENU)IDC_START_BTN, g_hInst, nullptr
    );

    // Stop button
    g_hStopBtn = CreateWindow(
        L"BUTTON", L"Stop Capture",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON | WS_DISABLED,
        120, 360, 100, 30,
        hwnd, (HMENU)IDC_STOP_BTN, g_hInst, nullptr
    );

    // Stop All button (initially hidden)
    g_hStopAllBtn = CreateWindow(
        L"BUTTON", L"Stop All",
        WS_CHILD | WS_TABSTOP | BS_PUSHBUTTON | WS_DISABLED,
        230, 360, 100, 30,
        hwnd, (HMENU)IDC_STOP_ALL_BTN, g_hInst, nullptr
    );

    // Pause All button (initially hidden)
    g_hPauseAllBtn = CreateWindow(
        L"BUTTON", L"Pause All",
        WS_CHILD | WS_TABSTOP | BS_PUSHBUTTON | WS_DISABLED,
        340, 360, 100, 30,
        hwnd, (HMENU)IDC_PAUSE_ALL_BTN, g_hInst, nullptr
    );

    // Resume All button (initially hidden)
    g_hResumeAllBtn = CreateWindow(
        L"BUTTON", L"Resume All",
        WS_CHILD | WS_TABSTOP | BS_PUSHBUTTON | WS_DISABLED,
        450, 360, 100, 30,
        hwnd, (HMENU)IDC_RESUME_ALL_BTN, g_hInst, nullptr
    );

    // Recording list label
    g_hRecordingListLabel = CreateWindow(
        L"STATIC", L"Active Recordings:",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        10, 400, 200, 20,
        hwnd, (HMENU)IDC_RECORDING_LIST_LABEL, g_hInst, nullptr
    );

    // Recording list (ListView)
    g_hRecordingList = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        WC_LISTVIEW,
        L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SINGLESEL,
        10, 420, 760, 100,
        hwnd,
        (HMENU)IDC_RECORDING_LIST,
        g_hInst,
        nullptr
    );

    lvc.mask = LVCF_TEXT | LVCF_WIDTH;
    lvc.cx = 180;
    lvc.pszText = (LPWSTR)L"Process";
    ListView_InsertColumn(g_hRecordingList, 0, &lvc);

    lvc.cx = 80;
    lvc.pszText = (LPWSTR)L"PID";
    ListView_InsertColumn(g_hRecordingList, 1, &lvc);

    lvc.cx = 380;
    lvc.pszText = (LPWSTR)L"Output File";
    ListView_InsertColumn(g_hRecordingList, 2, &lvc);

    lvc.cx = 100;
    lvc.pszText = (LPWSTR)L"Data Written";
    ListView_InsertColumn(g_hRecordingList, 3, &lvc);

    // Status text
    g_hStatusText = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        L"STATIC",
        L"Ready. Select a process and click Start Capture.",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        10, 520, 760, 30,
        hwnd, (HMENU)IDC_STATUS_TEXT, g_hInst, nullptr
    );
}

void RefreshProcessList() {
    // Save checked state before clearing the list
    std::vector<DWORD> checkedPIDs;
    int itemCount = ListView_GetItemCount(g_hProcessList);
    for (int i = 0; i < itemCount; i++) {
        if (ListView_GetCheckState(g_hProcessList, i)) {
            wchar_t pidStr[32];
            ListView_GetItemText(g_hProcessList, i, 1, pidStr, 32);
            DWORD pid = (DWORD)_wtoi(pidStr);
            checkedPIDs.push_back(pid);
        }
    }

    ListView_DeleteAllItems(g_hProcessList);
    g_processes = g_processEnum->GetAllProcesses();

    // Check if we should filter by active audio
    bool showAudioOnly = (SendMessage(g_hShowAudioOnlyCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);

    int displayedCount = 0;

    // Always add system-wide audio option as first item (unless filtering by active audio)
    if (!showAudioOnly) {
        LVITEM lvi = {};
        lvi.mask = LVIF_TEXT;
        lvi.iItem = 0;
        lvi.pszText = (LPWSTR)L"[System Audio - All Processes]";
        int index = ListView_InsertItem(g_hProcessList, &lvi);

        ListView_SetItemText(g_hProcessList, index, 1, (LPWSTR)L"0");
        ListView_SetItemText(g_hProcessList, index, 2, (LPWSTR)L"Capture all system audio");
        ListView_SetItemText(g_hProcessList, index, 3, (LPWSTR)L"System-wide loopback");

        // Check if system audio was previously checked (PID 0)
        for (DWORD checkedPID : checkedPIDs) {
            if (checkedPID == 0) {
                ListView_SetCheckState(g_hProcessList, index, TRUE);
                break;
            }
        }

        displayedCount++;
    }

    for (size_t i = 0; i < g_processes.size(); i++) {
        ProcessInfo& proc = g_processes[i];

        // Fetch window title and audio status on-demand (lazy loading for performance)
        if (proc.windowTitle.empty()) {
            proc.windowTitle = g_processEnum->GetWindowTitle(proc.processId);
        }

        // Only check audio status if we're filtering by it
        if (showAudioOnly) {
            proc.hasActiveAudio = g_processEnum->CheckProcessHasActiveAudio(proc.processId);

            // Skip if filtering and process doesn't have active audio
            if (!proc.hasActiveAudio) {
                continue;
            }
        }

        LVITEM lvi = {};
        lvi.mask = LVIF_TEXT;
        lvi.iItem = displayedCount;

        // Process name (first column)
        lvi.pszText = (LPWSTR)proc.processName.c_str();
        int index = ListView_InsertItem(g_hProcessList, &lvi);

        // PID (second column)
        wchar_t pidStr[32];
        swprintf_s(pidStr, L"%lu", proc.processId);
        ListView_SetItemText(g_hProcessList, index, 1, pidStr);

        // Window Title (third column)
        ListView_SetItemText(g_hProcessList, index, 2, (LPWSTR)proc.windowTitle.c_str());

        // Path (fourth column)
        ListView_SetItemText(g_hProcessList, index, 3, (LPWSTR)proc.executablePath.c_str());

        // Restore checked state if this PID was previously checked
        for (DWORD checkedPID : checkedPIDs) {
            if (proc.processId == checkedPID) {
                ListView_SetCheckState(g_hProcessList, index, TRUE);
                break;
            }
        }

        displayedCount++;
    }

    std::wstring statusMsg = L"Process list refreshed. Showing " + std::to_wstring(displayedCount) + L" process(es)";
    if (showAudioOnly) {
        statusMsg += L" with active audio";
    }
    statusMsg += L".";
    SetWindowText(g_hStatusText, statusMsg.c_str());
}

void StartCapture() {
    // Get all checked processes
    std::vector<int> checkedIndices;
    int itemCount = ListView_GetItemCount(g_hProcessList);

    for (int i = 0; i < itemCount; i++) {
        if (ListView_GetCheckState(g_hProcessList, i)) {
            checkedIndices.push_back(i);
        }
    }

    // If no processes are checked, fall back to the currently focused/selected item
    if (checkedIndices.empty()) {
        int focusedIndex = ListView_GetNextItem(g_hProcessList, -1, LVNI_FOCUSED);
        if (focusedIndex >= 0) {
            checkedIndices.push_back(focusedIndex);
        } else {
            MessageBox(g_hWnd, L"Please check one or more processes, or focus on a process to capture.", L"No Process Selected", MB_OK | MB_ICONWARNING);
            return;
        }
    }

    // Get output path
    wchar_t outputPath[MAX_PATH];
    GetWindowText(g_hOutputPath, outputPath, MAX_PATH);

    // Get format
    int formatIndex = (int)SendMessage(g_hFormatCombo, CB_GETCURSEL, 0, 0);
    AudioFormat format = AudioFormat::WAV;
    const wchar_t* extension = L".wav";
    UINT32 bitrate = 0;

    switch (formatIndex) {
    case 0:
        format = AudioFormat::WAV;
        extension = L".wav";
        break;
    case 1:
        format = AudioFormat::MP3;
        extension = L".mp3";
        // Get MP3 bitrate from combo box
        {
            int bitrateIndex = (int)SendMessage(g_hMp3BitrateCombo, CB_GETCURSEL, 0, 0);
            const UINT32 mp3Bitrates[] = { 128000, 192000, 256000, 320000 };
            bitrate = (bitrateIndex >= 0 && bitrateIndex < 4) ? mp3Bitrates[bitrateIndex] : 192000;
        }
        break;
    case 2:
        format = AudioFormat::OPUS;
        extension = L".opus";
        // Get Opus bitrate from combo box
        {
            int bitrateIndex = (int)SendMessage(g_hOpusBitrateCombo, CB_GETCURSEL, 0, 0);
            const UINT32 opusBitrates[] = { 64000, 96000, 128000, 192000, 256000 };
            bitrate = (bitrateIndex >= 0 && bitrateIndex < 5) ? opusBitrates[bitrateIndex] : 128000;
        }
        break;
    case 3:
        format = AudioFormat::FLAC;
        extension = L".flac";
        // Get FLAC compression level from combo box (0-8)
        {
            int compressionIndex = (int)SendMessage(g_hFlacCompressionCombo, CB_GETCURSEL, 0, 0);
            bitrate = (compressionIndex >= 0 && compressionIndex <= 8) ? compressionIndex : 5;
        }
        break;
    default:
        format = AudioFormat::WAV;
        extension = L".wav";
        break;
    }

    // Get skip silence option
    bool skipSilence = (SendMessage(g_hSkipSilenceCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);

    // Get passthrough option and device ID
    bool enablePassthrough = (SendMessage(g_hPassthroughCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
    std::wstring passthroughDeviceId;
    if (enablePassthrough && g_audioDeviceEnum) {
        int deviceIndex = (int)SendMessage(g_hPassthroughDeviceCombo, CB_GETCURSEL, 0, 0);
        if (deviceIndex >= 0) {
            const auto& devices = g_audioDeviceEnum->GetDevices();
            if (deviceIndex < static_cast<int>(devices.size())) {
                passthroughDeviceId = devices[deviceIndex].deviceId;
            }
        }
    }

    // Get monitor-only option
    bool monitorOnly = (SendMessage(g_hMonitorOnlyCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);

    // Check if microphone is enabled (counts as an additional source)
    bool captureMicrophone = (SendMessage(g_hMicrophoneCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);

    // Calculate total number of audio sources (processes + microphone)
    size_t totalSources = checkedIndices.size();
    if (captureMicrophone && !monitorOnly) {
        totalSources++; // Microphone counts as an additional source
    }

    // Get recording mode (only relevant if multiple sources and not monitor-only)
    int recordingModeIndex = (int)SendMessage(g_hRecordingModeCombo, CB_GETCURSEL, 0, 0);
    // 0 = Separate files, 1 = Combined file, 2 = Both

    bool createSeparateFiles = (totalSources == 1) || (recordingModeIndex == 0) || (recordingModeIndex == 2);
    bool createCombinedFile = (totalSources > 1) && !monitorOnly && ((recordingModeIndex == 1) || (recordingModeIndex == 2));

    int startedCount = 0;
    int alreadyCapturingCount = 0;

    for (int checkedIndex : checkedIndices) {
        // Get process info from ListView (not from g_processes, because system audio isn't in that vector)
        wchar_t processNameBuf[256];
        wchar_t pidStrBuf[32];

        ListView_GetItemText(g_hProcessList, checkedIndex, 0, processNameBuf, 256);
        ListView_GetItemText(g_hProcessList, checkedIndex, 1, pidStrBuf, 32);

        std::wstring processName = processNameBuf;
        DWORD processId = (DWORD)_wtoi(pidStrBuf);

        // Check if already capturing
        if (g_captureManager->IsCapturing(processId)) {
            alreadyCapturingCount++;
            continue;
        }

        // Get current time for timestamp
        SYSTEMTIME st;
        GetLocalTime(&st);
        wchar_t timestamp[64];
        swprintf_s(timestamp, L"%04d_%02d_%02d-%02d_%02d_%02d",
            st.wYear, st.wMonth, st.wDay,
            st.wHour, st.wMinute, st.wSecond);

        // Remove .exe extension from process name if present
        std::wstring cleanProcessName = processName;
        size_t exePos = cleanProcessName.find(L".exe");
        if (exePos != std::wstring::npos) {
            cleanProcessName = cleanProcessName.substr(0, exePos);
        }

        // Build full output path for separate file (if needed)
        std::wstring fullPath;
        bool captureMonitorOnly = monitorOnly;

        if (createSeparateFiles && !monitorOnly) {
            std::wstring basePath = outputPath;
            if (basePath.back() != L'\\') {
                basePath += L'\\';
            }
            fullPath = basePath + cleanProcessName + L"-" + timestamp + extension;
        }
        else {
            // Monitor-only or combined-only mode, no separate file path needed
            fullPath = L"";
            // In combined-only mode, treat individual captures as monitor-only
            if (!createSeparateFiles && createCombinedFile) {
                captureMonitorOnly = true;
            }
        }

        // Start capture with bitrate, skip silence option, passthrough device, and monitor-only mode
        if (g_captureManager->StartCapture(processId, processName, fullPath, format, bitrate, skipSilence, passthroughDeviceId, captureMonitorOnly)) {
            startedCount++;
        }
    }

    // If combined file mode is enabled, start the mixer
    if (createCombinedFile && startedCount > 0) {
        // Build combined file path
        std::wstring combinedPath = outputPath;
        if (combinedPath.back() != L'\\') {
            combinedPath += L'\\';
        }

        SYSTEMTIME st;
        GetLocalTime(&st);
        wchar_t timestamp[64];
        swprintf_s(timestamp, L"%04d_%02d_%02d-%02d_%02d_%02d",
            st.wYear, st.wMonth, st.wDay,
            st.wHour, st.wMinute, st.wSecond);

        combinedPath += L"Combined-" + std::wstring(timestamp) + extension;

        // Enable mixed recording
        if (!g_captureManager->EnableMixedRecording(combinedPath, format, bitrate)) {
            MessageBox(g_hWnd, L"Failed to enable combined recording.", L"Warning", MB_OK | MB_ICONWARNING);
        }
    }

    // Handle microphone capture if enabled (and not in monitor-only mode)
    if (captureMicrophone && !monitorOnly && g_audioDeviceEnum) {
        // Get microphone device ID
        int micDeviceIndex = (int)SendMessage(g_hMicrophoneDeviceCombo, CB_GETCURSEL, 0, 0);
        const auto& inputDevices = g_audioDeviceEnum->GetInputDevices();

        if (micDeviceIndex >= 0 && micDeviceIndex < static_cast<int>(inputDevices.size())) {
            std::wstring micDeviceId = inputDevices[micDeviceIndex].deviceId;
            std::wstring micDeviceName = inputDevices[micDeviceIndex].friendlyName;

            // Use special process ID for microphone (0xFFFFFFFE)
            DWORD micProcessId = 0xFFFFFFFE;

            // Check if already capturing microphone
            if (!g_captureManager->IsCapturing(micProcessId)) {
                // Determine microphone capture mode based on recording mode
                bool createMicFile = false;
                bool micMonitorOnly = false;
                std::wstring micFilePath;

                if (createSeparateFiles) {
                    // Separate files mode: Create mic file
                    createMicFile = true;
                }

                if (createCombinedFile) {
                    // Combined file mode: Send to mixer only (monitor-only mode)
                    if (!createSeparateFiles) {
                        // Combined only - microphone is monitor-only (sends to mixer)
                        micMonitorOnly = true;
                    }
                    // If both createSeparateFiles and createCombinedFile are true (Both mode),
                    // then createMicFile=true and micMonitorOnly=false, which means:
                    // - Mic creates its own file
                    // - Mic also sends to mixer (via OnAudioData callback)
                }

                // Build microphone file path if needed
                if (createMicFile) {
                    std::wstring basePath = outputPath;
                    if (basePath.back() != L'\\') {
                        basePath += L'\\';
                    }

                    // Get current time for timestamp
                    SYSTEMTIME st;
                    GetLocalTime(&st);
                    wchar_t timestamp[64];
                    swprintf_s(timestamp, L"%04d_%02d_%02d-%02d_%02d_%02d",
                        st.wYear, st.wMonth, st.wDay,
                        st.wHour, st.wMinute, st.wSecond);

                    micFilePath = basePath + L"Microphone-" + std::wstring(timestamp) + extension;
                } else {
                    // Monitor-only mode, no file path needed
                    micFilePath = L"";
                }

                // Start microphone capture
                if (g_captureManager->StartCaptureFromDevice(
                    micProcessId,
                    micDeviceName,
                    micDeviceId,
                    true, // isInputDevice
                    micFilePath,
                    format,
                    bitrate,
                    skipSilence,
                    micMonitorOnly)) {
                    startedCount++;
                }
            }
        }
    }

    if (startedCount > 0) {
        EnableWindow(g_hStopBtn, TRUE);
        UpdateRecordingList();
        std::wstring status = L"Started " + std::to_wstring(startedCount) + L" capture(s)";
        if (alreadyCapturingCount > 0) {
            status += L" (" + std::to_wstring(alreadyCapturingCount) + L" already capturing)";
        }
        SetWindowText(g_hStatusText, status.c_str());
    }
    else if (alreadyCapturingCount > 0) {
        MessageBox(g_hWnd, L"All selected processes are already being captured.", L"Already Capturing", MB_OK | MB_ICONINFORMATION);
    }
    else {
        MessageBox(g_hWnd, L"Failed to start any captures.", L"Capture Error", MB_OK | MB_ICONERROR);
    }
}

void StopCapture() {
    // Get selected recording
    int selectedIndex = ListView_GetNextItem(g_hRecordingList, -1, LVNI_SELECTED);
    if (selectedIndex < 0) {
        MessageBox(g_hWnd, L"Please select a recording to stop.", L"No Recording Selected", MB_OK | MB_ICONWARNING);
        return;
    }

    // Get PID from list view (column 1 now)
    wchar_t pidStr[32];
    ListView_GetItemText(g_hRecordingList, selectedIndex, 1, pidStr, 32);
    DWORD processId = (DWORD)wcstoul(pidStr, nullptr, 10);

    if (g_captureManager->StopCapture(processId)) {
        UpdateRecordingList();
        SetWindowText(g_hStatusText, L"Capture stopped.");

        auto sessions = g_captureManager->GetActiveSessions();
        if (sessions.empty()) {
            EnableWindow(g_hStopBtn, FALSE);
            g_captureManager->DisableMixedRecording(); // Stop mixed recording if no more sessions
        }

        // Restore focus to process list to prevent keyboard focus issues
        SetFocus(g_hProcessList);
    }
}

void UpdateRecordingList() {
    // Save currently selected PID before updating
    DWORD selectedPID = 0;
    int selectedIndex = ListView_GetNextItem(g_hRecordingList, -1, LVNI_SELECTED);
    if (selectedIndex >= 0) {
        wchar_t pidStr[32];
        ListView_GetItemText(g_hRecordingList, selectedIndex, 1, pidStr, 32);
        selectedPID = (DWORD)wcstoul(pidStr, nullptr, 10);
    }

    ListView_DeleteAllItems(g_hRecordingList);

    auto sessions = g_captureManager->GetActiveSessions();
    int newSelectedIndex = -1;

    for (size_t i = 0; i < sessions.size(); i++) {
        CaptureSession* session = sessions[i];

        LVITEM lvi = {};
        lvi.mask = LVIF_TEXT;
        lvi.iItem = static_cast<int>(i);

        // Process name (first column)
        lvi.pszText = (LPWSTR)session->processName.c_str();
        int index = ListView_InsertItem(g_hRecordingList, &lvi);

        // PID (second column)
        wchar_t pidStr[32];
        swprintf_s(pidStr, L"%lu", session->processId);
        ListView_SetItemText(g_hRecordingList, index, 1, pidStr);

        // Output file (third column)
        if (session->monitorOnly) {
            ListView_SetItemText(g_hRecordingList, index, 2, (LPWSTR)L"[Monitor Only - No Recording]");
        } else {
            ListView_SetItemText(g_hRecordingList, index, 2, (LPWSTR)session->outputFile.c_str());
        }

        // Data written (fourth column)
        if (session->monitorOnly) {
            ListView_SetItemText(g_hRecordingList, index, 3, (LPWSTR)L"N/A");
        } else {
            std::wstring sizeStr = FormatFileSize(session->bytesWritten);
            ListView_SetItemText(g_hRecordingList, index, 3, (LPWSTR)sizeStr.c_str());
        }

        // Check if this was the previously selected item
        if (selectedIndex >= 0 && session->processId == selectedPID) {
            newSelectedIndex = index;
        }
    }

    // Restore selection if the item still exists
    if (newSelectedIndex >= 0) {
        ListView_SetItemState(g_hRecordingList, newSelectedIndex,
            LVIS_SELECTED | LVIS_FOCUSED,
            LVIS_SELECTED | LVIS_FOCUSED);
    }

    // Show/enable Stop All, Pause All, and Resume All buttons if there are multiple captures
    if (sessions.size() >= 2) {
        ShowWindow(g_hStopAllBtn, SW_SHOW);
        EnableWindow(g_hStopAllBtn, TRUE);
        ShowWindow(g_hPauseAllBtn, SW_SHOW);
        ShowWindow(g_hResumeAllBtn, SW_SHOW);

        // Check pause state of all sessions to enable/disable buttons intelligently
        int pausedCount = 0;
        int resumedCount = 0;
        for (auto* session : sessions) {
            if (session->capture) {
                if (session->capture->IsPaused()) {
                    pausedCount++;
                } else {
                    resumedCount++;
                }
            }
        }

        // Enable/disable buttons based on pause state
        // If all are paused, disable Pause All button
        // If all are resumed, disable Resume All button
        EnableWindow(g_hPauseAllBtn, resumedCount > 0);
        EnableWindow(g_hResumeAllBtn, pausedCount > 0);
    } else {
        ShowWindow(g_hStopAllBtn, SW_HIDE);
        EnableWindow(g_hStopAllBtn, FALSE);
        ShowWindow(g_hPauseAllBtn, SW_HIDE);
        EnableWindow(g_hPauseAllBtn, FALSE);
        ShowWindow(g_hResumeAllBtn, SW_HIDE);
        EnableWindow(g_hResumeAllBtn, FALSE);
    }
}

void BrowseOutputFolder() {
    IFileDialog* pfd = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pfd));
    if (SUCCEEDED(hr)) {
        DWORD dwOptions;
        pfd->GetOptions(&dwOptions);
        pfd->SetOptions(dwOptions | FOS_PICKFOLDERS);

        hr = pfd->Show(g_hWnd);
        if (SUCCEEDED(hr)) {
            IShellItem* psi = nullptr;
            hr = pfd->GetResult(&psi);
            if (SUCCEEDED(hr)) {
                PWSTR pszPath = nullptr;
                hr = psi->GetDisplayName(SIGDN_FILESYSPATH, &pszPath);
                if (SUCCEEDED(hr)) {
                    SetWindowText(g_hOutputPath, pszPath);
                    CoTaskMemFree(pszPath);
                }
                psi->Release();
            }
        }
        pfd->Release();
    }
}

std::wstring GetDefaultOutputPath() {
    wchar_t path[MAX_PATH];
    if (SUCCEEDED(SHGetFolderPath(nullptr, CSIDL_MYDOCUMENTS, nullptr, 0, path))) {
        std::wstring docPath = path;
        docPath += L"\\AudioCaptures";
        CreateDirectory(docPath.c_str(), nullptr);
        return docPath;
    }
    return L"C:\\AudioCaptures";
}

std::wstring FormatFileSize(UINT64 bytes) {
    const wchar_t* units[] = { L"B", L"KB", L"MB", L"GB" };
    int unitIndex = 0;
    double size = static_cast<double>(bytes);

    while (size >= 1024.0 && unitIndex < 3) {
        size /= 1024.0;
        unitIndex++;
    }

    wchar_t buffer[64];
    swprintf_s(buffer, L"%.2f %s", size, units[unitIndex]);
    return buffer;
}

// Helper functions for string conversion
std::string WStringToString(const std::wstring& wstr) {
    if (wstr.empty()) return std::string();
    int size = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string result(size - 1, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &result[0], size, nullptr, nullptr);
    return result;
}

std::wstring StringToWString(const std::string& str) {
    if (str.empty()) return std::wstring();
    int size = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, nullptr, 0);
    std::wstring result(size - 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, &result[0], size);
    return result;
}

std::wstring GetSettingsFilePath() {
    wchar_t path[MAX_PATH];
    if (SUCCEEDED(SHGetFolderPath(nullptr, CSIDL_LOCAL_APPDATA, nullptr, 0, path))) {
        std::wstring settingsDir = path;
        settingsDir += L"\\AudioCapture";
        CreateDirectory(settingsDir.c_str(), nullptr);
        return settingsDir + L"\\settings.json";
    }
    return L"settings.json";
}

void LoadSettings() {
    std::wstring settingsPath = GetSettingsFilePath();
    std::ifstream file(settingsPath);

    if (file.is_open()) {
        try {
            json settings = json::parse(file);

            // Load output path
            if (settings.contains("outputPath") && settings["outputPath"].is_string()) {
                std::string outputPath = settings["outputPath"];
                SetWindowTextW(g_hOutputPath, StringToWString(outputPath).c_str());
            }

            // Load format
            if (settings.contains("format") && settings["format"].is_number_integer()) {
                int formatIndex = settings["format"];
                if (formatIndex >= 0 && formatIndex <= 3) {
                    SendMessage(g_hFormatCombo, CB_SETCURSEL, formatIndex, 0);
                }
            }

            // Load MP3 bitrate
            if (settings.contains("mp3Bitrate") && settings["mp3Bitrate"].is_number_integer()) {
                int bitrateIndex = settings["mp3Bitrate"];
                if (bitrateIndex >= 0 && bitrateIndex <= 3) {
                    SendMessage(g_hMp3BitrateCombo, CB_SETCURSEL, bitrateIndex, 0);
                }
            }

            // Load Opus bitrate
            if (settings.contains("opusBitrate") && settings["opusBitrate"].is_number_integer()) {
                int bitrateIndex = settings["opusBitrate"];
                if (bitrateIndex >= 0 && bitrateIndex <= 4) {
                    SendMessage(g_hOpusBitrateCombo, CB_SETCURSEL, bitrateIndex, 0);
                }
            }

            // Load FLAC compression
            if (settings.contains("flacCompression") && settings["flacCompression"].is_number_integer()) {
                int compressionIndex = settings["flacCompression"];
                if (compressionIndex >= 0 && compressionIndex <= 8) {
                    SendMessage(g_hFlacCompressionCombo, CB_SETCURSEL, compressionIndex, 0);
                }
            }

            // Load skip silence option
            if (settings.contains("skipSilence") && settings["skipSilence"].is_boolean()) {
                bool skipSilence = settings["skipSilence"];
                SendMessage(g_hSkipSilenceCheckbox, BM_SETCHECK, skipSilence ? BST_CHECKED : BST_UNCHECKED, 0);
            }

            // Load passthrough (monitor audio) option
            if (settings.contains("passthrough") && settings["passthrough"].is_boolean()) {
                bool passthrough = settings["passthrough"];
                SendMessage(g_hPassthroughCheckbox, BM_SETCHECK, passthrough ? BST_CHECKED : BST_UNCHECKED, 0);
            }

            // Load passthrough device index
            if (settings.contains("passthroughDeviceIndex") && settings["passthroughDeviceIndex"].is_number_integer()) {
                int deviceIndex = settings["passthroughDeviceIndex"];
                // Will be applied after device enumeration completes
                if (deviceIndex >= 0 && deviceIndex < SendMessage(g_hPassthroughDeviceCombo, CB_GETCOUNT, 0, 0)) {
                    SendMessage(g_hPassthroughDeviceCombo, CB_SETCURSEL, deviceIndex, 0);
                }
            }

            // Load monitor only option
            if (settings.contains("monitorOnly") && settings["monitorOnly"].is_boolean()) {
                bool monitorOnly = settings["monitorOnly"];
                SendMessage(g_hMonitorOnlyCheckbox, BM_SETCHECK, monitorOnly ? BST_CHECKED : BST_UNCHECKED, 0);
            }

            // Load recording mode
            if (settings.contains("recordingMode") && settings["recordingMode"].is_number_integer()) {
                int recordingMode = settings["recordingMode"];
                if (recordingMode >= 0 && recordingMode <= 2) {
                    SendMessage(g_hRecordingModeCombo, CB_SETCURSEL, recordingMode, 0);
                }
            }

            // Load microphone capture option
            if (settings.contains("captureMicrophone") && settings["captureMicrophone"].is_boolean()) {
                bool captureMicrophone = settings["captureMicrophone"];
                SendMessage(g_hMicrophoneCheckbox, BM_SETCHECK, captureMicrophone ? BST_CHECKED : BST_UNCHECKED, 0);
            }

            // Load microphone device index
            if (settings.contains("microphoneDeviceIndex") && settings["microphoneDeviceIndex"].is_number_integer()) {
                int deviceIndex = settings["microphoneDeviceIndex"];
                // Will be applied after device enumeration completes
                if (deviceIndex >= 0 && deviceIndex < SendMessage(g_hMicrophoneDeviceCombo, CB_GETCOUNT, 0, 0)) {
                    SendMessage(g_hMicrophoneDeviceCombo, CB_SETCURSEL, deviceIndex, 0);
                }
            }
        }
        catch (...) {
            // If parsing fails, just use defaults
        }
        file.close();
    }

    // Update visibility of bitrate controls based on selected format
    OnFormatChanged();

    // Update visibility of passthrough controls
    OnPassthroughCheckboxChanged();

    // Update state of recording controls based on monitor-only
    OnMonitorOnlyCheckboxChanged();

    // Update visibility of microphone controls
    OnMicrophoneCheckboxChanged();
}

void SaveSettings() {
    json settings;

    // Save output path
    wchar_t outputPath[MAX_PATH];
    GetWindowTextW(g_hOutputPath, outputPath, MAX_PATH);
    settings["outputPath"] = WStringToString(outputPath);

    // Save format
    int formatIndex = (int)SendMessage(g_hFormatCombo, CB_GETCURSEL, 0, 0);
    settings["format"] = formatIndex;

    // Save bitrates and compression
    int mp3BitrateIndex = (int)SendMessage(g_hMp3BitrateCombo, CB_GETCURSEL, 0, 0);
    settings["mp3Bitrate"] = mp3BitrateIndex;

    int opusBitrateIndex = (int)SendMessage(g_hOpusBitrateCombo, CB_GETCURSEL, 0, 0);
    settings["opusBitrate"] = opusBitrateIndex;

    int flacCompressionIndex = (int)SendMessage(g_hFlacCompressionCombo, CB_GETCURSEL, 0, 0);
    settings["flacCompression"] = flacCompressionIndex;

    // Save skip silence option
    bool skipSilence = (SendMessage(g_hSkipSilenceCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
    settings["skipSilence"] = skipSilence;

    // Save passthrough (monitor audio) option
    bool passthrough = (SendMessage(g_hPassthroughCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
    settings["passthrough"] = passthrough;

    // Save passthrough device index
    int passthroughDeviceIndex = (int)SendMessage(g_hPassthroughDeviceCombo, CB_GETCURSEL, 0, 0);
    settings["passthroughDeviceIndex"] = passthroughDeviceIndex;

    // Save monitor only option
    bool monitorOnly = (SendMessage(g_hMonitorOnlyCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
    settings["monitorOnly"] = monitorOnly;

    // Save recording mode
    int recordingMode = (int)SendMessage(g_hRecordingModeCombo, CB_GETCURSEL, 0, 0);
    settings["recordingMode"] = recordingMode;

    // Save microphone capture option
    bool captureMicrophone = (SendMessage(g_hMicrophoneCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
    settings["captureMicrophone"] = captureMicrophone;

    // Save microphone device index
    int microphoneDeviceIndex = (int)SendMessage(g_hMicrophoneDeviceCombo, CB_GETCURSEL, 0, 0);
    settings["microphoneDeviceIndex"] = microphoneDeviceIndex;

    // Write to file
    std::wstring settingsPath = GetSettingsFilePath();
    std::ofstream file(settingsPath);
    if (file.is_open()) {
        file << settings.dump(4); // Pretty print with 4 spaces
        file.close();
    }
}

void OnFormatChanged() {
    int formatIndex = (int)SendMessage(g_hFormatCombo, CB_GETCURSEL, 0, 0);

    // Hide all bitrate/compression controls first
    ShowWindow(g_hMp3BitrateLabel, SW_HIDE);
    ShowWindow(g_hMp3BitrateCombo, SW_HIDE);
    ShowWindow(g_hOpusBitrateLabel, SW_HIDE);
    ShowWindow(g_hOpusBitrateCombo, SW_HIDE);
    ShowWindow(g_hFlacCompressionLabel, SW_HIDE);
    ShowWindow(g_hFlacCompressionCombo, SW_HIDE);

    // Show appropriate control based on format
    switch (formatIndex) {
    case 1: // MP3
        ShowWindow(g_hMp3BitrateLabel, SW_SHOW);
        ShowWindow(g_hMp3BitrateCombo, SW_SHOW);
        break;
    case 2: // Opus
        ShowWindow(g_hOpusBitrateLabel, SW_SHOW);
        ShowWindow(g_hOpusBitrateCombo, SW_SHOW);
        break;
    case 3: // FLAC
        ShowWindow(g_hFlacCompressionLabel, SW_SHOW);
        ShowWindow(g_hFlacCompressionCombo, SW_SHOW);
        break;
    }
}

void PopulatePassthroughDevices() {
    if (!g_audioDeviceEnum) {
        return;
    }

    // Enumerate audio devices
    if (!g_audioDeviceEnum->EnumerateDevices()) {
        return;
    }

    // Clear existing items
    SendMessage(g_hPassthroughDeviceCombo, CB_RESETCONTENT, 0, 0);

    // Add devices to combo box
    const auto& devices = g_audioDeviceEnum->GetDevices();
    int defaultIndex = -1;

    for (size_t i = 0; i < devices.size(); i++) {
        const AudioDeviceInfo& device = devices[i];

        // Format name with (Default) suffix if it's the default device
        std::wstring displayName = device.friendlyName;
        if (device.isDefault) {
            displayName += L" (Default)";
            defaultIndex = static_cast<int>(i);
        }

        SendMessage(g_hPassthroughDeviceCombo, CB_ADDSTRING, 0, (LPARAM)displayName.c_str());
        // Store device index as item data
        SendMessage(g_hPassthroughDeviceCombo, CB_SETITEMDATA, i, (LPARAM)i);
    }

    // Select default device
    if (defaultIndex >= 0) {
        SendMessage(g_hPassthroughDeviceCombo, CB_SETCURSEL, defaultIndex, 0);
    } else if (devices.size() > 0) {
        SendMessage(g_hPassthroughDeviceCombo, CB_SETCURSEL, 0, 0);
    }

    // Initially hide passthrough controls
    OnPassthroughCheckboxChanged();
}

void OnPassthroughCheckboxChanged() {
    // Show/hide device selector based on checkbox state
    BOOL isChecked = (SendMessage(g_hPassthroughCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);

    ShowWindow(g_hPassthroughDeviceLabel, isChecked ? SW_SHOW : SW_HIDE);
    ShowWindow(g_hPassthroughDeviceCombo, isChecked ? SW_SHOW : SW_HIDE);

    // Enable/disable monitor-only checkbox based on monitoring state
    EnableWindow(g_hMonitorOnlyCheckbox, isChecked);

    // If monitoring is disabled, uncheck monitor-only
    if (!isChecked && SendMessage(g_hMonitorOnlyCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED) {
        SendMessage(g_hMonitorOnlyCheckbox, BM_SETCHECK, BST_UNCHECKED, 0);
        OnMonitorOnlyCheckboxChanged(); // Update dependent controls
    }
}

void OnMonitorOnlyCheckboxChanged() {
    // Enable/disable recording-specific controls based on monitor-only state
    BOOL isMonitorOnly = (SendMessage(g_hMonitorOnlyCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
    BOOL enableRecordingControls = !isMonitorOnly;

    // Disable format selection
    EnableWindow(g_hFormatCombo, enableRecordingControls);

    // Disable bitrate/compression controls
    EnableWindow(g_hMp3BitrateCombo, enableRecordingControls);
    EnableWindow(g_hOpusBitrateCombo, enableRecordingControls);
    EnableWindow(g_hFlacCompressionCombo, enableRecordingControls);

    // Disable output path controls
    EnableWindow(g_hOutputPath, enableRecordingControls);
    EnableWindow(g_hBrowseBtn, enableRecordingControls);

    // Disable skip silence option (only relevant for recording)
    EnableWindow(g_hSkipSilenceCheckbox, enableRecordingControls);

    // Show/hide recording mode controls (only relevant when recording)
    ShowWindow(g_hRecordingModeLabel, enableRecordingControls ? SW_SHOW : SW_HIDE);
    ShowWindow(g_hRecordingModeCombo, enableRecordingControls ? SW_SHOW : SW_HIDE);
}

void PopulateMicrophoneDevices() {
    if (!g_audioDeviceEnum) {
        return;
    }

    // Enumerate input devices
    if (!g_audioDeviceEnum->EnumerateInputDevices()) {
        return;
    }

    // Clear existing items
    SendMessage(g_hMicrophoneDeviceCombo, CB_RESETCONTENT, 0, 0);

    // Add devices to combo box
    const auto& devices = g_audioDeviceEnum->GetInputDevices();
    int defaultIndex = -1;

    for (size_t i = 0; i < devices.size(); i++) {
        const AudioDeviceInfo& device = devices[i];

        // Format name with (Default) suffix if it's the default device
        std::wstring displayName = device.friendlyName;
        if (device.isDefault) {
            displayName += L" (Default)";
            defaultIndex = static_cast<int>(i);
        }

        SendMessage(g_hMicrophoneDeviceCombo, CB_ADDSTRING, 0, (LPARAM)displayName.c_str());
        // Store device index as item data
        SendMessage(g_hMicrophoneDeviceCombo, CB_SETITEMDATA, i, (LPARAM)i);
    }

    // Select default device
    if (defaultIndex >= 0) {
        SendMessage(g_hMicrophoneDeviceCombo, CB_SETCURSEL, defaultIndex, 0);
    } else if (devices.size() > 0) {
        SendMessage(g_hMicrophoneDeviceCombo, CB_SETCURSEL, 0, 0);
    }

    // Initially hide microphone controls
    OnMicrophoneCheckboxChanged();
}

void OnMicrophoneCheckboxChanged() {
    // Show/hide device selector based on checkbox state
    BOOL isChecked = (SendMessage(g_hMicrophoneCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);

    ShowWindow(g_hMicrophoneDeviceLabel, isChecked ? SW_SHOW : SW_HIDE);
    ShowWindow(g_hMicrophoneDeviceCombo, isChecked ? SW_SHOW : SW_HIDE);
}
