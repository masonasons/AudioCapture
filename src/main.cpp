#include <windows.h>
#include <roapi.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <shellapi.h>
#include <string>
#include <vector>
#include <fstream>
#include <algorithm>
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
HWND g_hMicrophoneDeviceList;
HWND g_hMicrophoneDeviceLabel;
HWND g_hFormatLabel;
HWND g_hProcessVolumeSlider;
HWND g_hProcessVolumeLabel;
HWND g_hMicrophoneVolumeSlider;
HWND g_hMicrophoneVolumeLabel;
HWND g_hPresetLabel;
HWND g_hPresetCombo;
HWND g_hSavePresetBtn;
HWND g_hLoadPresetBtn;
HWND g_hDeletePresetBtn;
HWND g_hLastFocusedCtrl = nullptr;
bool g_isAppActive = true;
const UINT WM_APP_RESTORE_FOCUS = WM_APP + 1;
bool g_captureButtonStops = false;
bool g_restoreFocusOnActivate = false;

// Volume settings (0-100%)
float g_processVolume = 100.0f;  // Default to 100%
float g_microphoneVolume = 100.0f;  // Default to 100%

std::unique_ptr<ProcessEnumerator> g_processEnum;
std::unique_ptr<CaptureManager> g_captureManager;
std::unique_ptr<AudioDeviceEnumerator> g_audioDeviceEnum;
std::vector<ProcessInfo> g_processes;
std::vector<std::wstring> g_pendingMicrophoneDeviceIds;
int g_pendingMicrophoneDeviceIndex = -1;
bool g_useWinRT = false;  // Track whether we initialized with WinRT or COM
bool g_supportsProcessCapture = false;  // Track whether OS supports process-specific capture

// Tray icon
NOTIFYICONDATA g_nid = {};
bool g_isMinimizedToTray = false;

// Window class name
const wchar_t CLASS_NAME[] = L"AudioCaptureWindow";

// Forward declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void InitializeControls(HWND hwnd);
void RefreshProcessList();
void StartCapture();
void StopCapture();
void UpdateRecordingList();
void EnsureRecordingListFocusItem();
void BrowseOutputFolder();
void OnFormatChanged();
std::wstring GetDefaultOutputPath();
std::wstring FormatFileSize(UINT64 bytes);
std::wstring NormalizeOutputPath(const std::wstring& path);
bool EnsureDirectoryExists(const std::wstring& path);
std::wstring SanitizeFileName(const std::wstring& name);
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
void AddTrayIcon();
void RemoveTrayIcon();
void ShowTrayContextMenu();
void ShowWindowFromTray();
void HideWindowToTray();
std::wstring GetPresetsDirectory();
std::vector<std::wstring> GetAvailablePresets();
void PopulatePresetCombo();
json GetCurrentSettingsAsJson();
void ApplySettingsFromJson(const json& preset);
std::vector<std::wstring> GetCheckedProcessNames();
std::vector<size_t> GetCheckedMicrophoneDeviceIndices();
std::vector<std::wstring> GetCheckedMicrophoneDeviceIds();
void ApplyMicrophoneSelection(const std::vector<std::wstring>& deviceIds, int fallbackIndex);
void CheckProcessesByNames(const std::vector<std::wstring>& names);
void SavePreset();
void LoadPreset();
void DeletePreset();
LRESULT CALLBACK PresetNameDialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
bool IsValidPresetName(const std::wstring& name);

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

// Helper to detect if process-specific capture is supported (Windows 10 Build 19041+)
bool SupportsProcessCapture() {
    typedef LONG (WINAPI* RtlGetVersionPtr)(PRTL_OSVERSIONINFOW);
    HMODULE hNtdll = GetModuleHandleW(L"ntdll.dll");
    if (!hNtdll) return false;

    RtlGetVersionPtr RtlGetVersion = (RtlGetVersionPtr)GetProcAddress(hNtdll, "RtlGetVersion");
    if (!RtlGetVersion) return false;

    RTL_OSVERSIONINFOW osvi = { 0 };
    osvi.dwOSVersionInfoSize = sizeof(osvi);
    if (RtlGetVersion(&osvi) == 0) {
        // Windows 10 version 2004 (build 19041) and later supports process-specific capture
        // Windows 11 is version 10.0.22000+
        return (osvi.dwMajorVersion == 10 && osvi.dwBuildNumber >= 19041) || (osvi.dwMajorVersion > 10);
    }
    return false;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int nCmdShow) {
    g_hInst = hInstance;

    // Detect OS capabilities
    g_supportsProcessCapture = SupportsProcessCapture();

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
    icex.dwICC = ICC_LISTVIEW_CLASSES | ICC_STANDARD_CLASSES | ICC_BAR_CLASSES;
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

    // Create window with appropriate title
    const wchar_t* windowTitle = g_supportsProcessCapture ?
        L"Audio Capture - Per-Process Recording" :
        L"Audio Capture - System Audio";

    g_hWnd = CreateWindowEx(
        0,
        CLASS_NAME,
        windowTitle,
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
    case WM_ACTIVATE: {
        if (LOWORD(wParam) == WA_INACTIVE) {
            g_hLastFocusedCtrl = GetFocus();
            g_restoreFocusOnActivate = true;
        } else {
            if (g_restoreFocusOnActivate) {
                HWND target = g_hLastFocusedCtrl;
                if (!target || !IsWindow(target)) {
                    target = g_supportsProcessCapture ? g_hProcessList : g_hStartBtn;
                }
                if (target && IsWindow(target)) {
                    PostMessage(hwnd, WM_APP_RESTORE_FOCUS, (WPARAM)target, 0);
                }
                g_restoreFocusOnActivate = false;
            }
        }
        return 0;
    }

    case WM_ACTIVATEAPP:
        g_isAppActive = (wParam != 0);
        if (g_isAppActive) {
            if (g_restoreFocusOnActivate) {
                HWND target = g_hLastFocusedCtrl;
                if (!target || !IsWindow(target)) {
                    target = g_supportsProcessCapture ? g_hProcessList : g_hStartBtn;
                }
                if (target && IsWindow(target)) {
                    PostMessage(hwnd, WM_APP_RESTORE_FOCUS, (WPARAM)target, 0);
                }
                g_restoreFocusOnActivate = false;
            }
        } else {
            g_restoreFocusOnActivate = true;
        }
        return 0;

    case WM_APP_RESTORE_FOCUS: {
        HWND target = (HWND)wParam;
        if (target && IsWindow(target)) {
            SetFocus(target);
            if (target == g_hRecordingList) {
                EnsureRecordingListFocusItem();
            }
        }
        return 0;
    }

    case WM_SETFOCUS: {
        HWND target = g_hLastFocusedCtrl;
        if (!target || !IsWindow(target)) {
            target = g_supportsProcessCapture ? g_hProcessList : g_hStartBtn;
        }
        SetFocus(target);
        if (target == g_hRecordingList) {
            EnsureRecordingListFocusItem();
        }
        return 0;
    }

    case WM_CREATE:
        InitializeControls(hwnd);
        LoadSettings();
        PopulatePresetCombo();  // Load available presets
        RefreshProcessList();
        // Set initial focus based on OS support
        if (g_supportsProcessCapture) {
            SetFocus(g_hProcessList);
        } else {
            SetFocus(g_hStartBtn);
        }
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

    case WM_NOTIFY: {
        LPNMHDR hdr = (LPNMHDR)lParam;
        if (hdr && hdr->hwndFrom == g_hRecordingList && hdr->code == LVN_KEYDOWN) {
            NMLVKEYDOWN* key = (NMLVKEYDOWN*)lParam;
            if (key->wVKey == VK_TAB) {
                bool shiftDown = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
                HWND next = GetNextDlgTabItem(hwnd, g_hRecordingList, shiftDown);
                if (next && IsWindow(next)) {
                    SendMessage(hwnd, WM_NEXTDLGCTL, (WPARAM)next, TRUE);
                    return TRUE;
                }
            }
        }
        break;
    }

    case WM_COMMAND: {
        int wmId = LOWORD(wParam);
        int wmEvent = HIWORD(wParam);

        switch (wmId) {
        case IDC_REFRESH_BTN:
            RefreshProcessList();
            break;

        case IDC_START_BTN:
            SetFocus(g_hStartBtn);
            if (g_captureButtonStops) {
                StopCapture();
            } else {
                StartCapture();
            }
            break;

        case IDC_STOP_ALL_BTN:
            if (MessageBox(g_hWnd, L"Stop all active captures?", L"Confirm", MB_YESNO | MB_ICONQUESTION) == IDYES) {
                // Stop all individual captures FIRST (stops audio callbacks)
                g_captureManager->StopAllCaptures();
                // Then disable mixed recording (mixer thread can exit cleanly)
                g_captureManager->DisableMixedRecording();
                EnableWindow(g_hStopAllBtn, FALSE);
                ShowWindow(g_hStopAllBtn, SW_HIDE);
                UpdateRecordingList();
                SetWindowText(g_hStatusText, L"All captures stopped.");
                if (g_supportsProcessCapture) {
                    SetFocus(g_hProcessList);
                } else {
                    SetFocus(g_hStartBtn);
                }
            }
            break;

        case IDC_PAUSE_ALL_BTN:
            g_captureManager->PauseAllCaptures();
            SetWindowText(g_hStatusText, L"All captures paused.");
            UpdateRecordingList();  // Update button states immediately
            if (g_supportsProcessCapture) {
                SetFocus(g_hProcessList);
            } else {
                SetFocus(g_hStartBtn);
            }
            break;

        case IDC_RESUME_ALL_BTN:
            g_captureManager->ResumeAllCaptures();
            SetWindowText(g_hStatusText, L"All captures resumed.");
            UpdateRecordingList();  // Update button states immediately
            if (g_supportsProcessCapture) {
                SetFocus(g_hProcessList);
            } else {
                SetFocus(g_hStartBtn);
            }
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

        case IDC_SAVE_PRESET_BTN:
            SavePreset();
            break;

        case IDC_LOAD_PRESET_BTN:
            LoadPreset();
            break;

        case IDC_DELETE_PRESET_BTN:
            DeletePreset();
            break;

        case IDM_TRAY_SHOW:
            ShowWindowFromTray();
            break;

        case IDM_TRAY_EXIT:
            RemoveTrayIcon();
            PostQuitMessage(0);
            break;
        }
        return 0;
    }

    case WM_TRAYICON: {
        switch (lParam) {
        case WM_LBUTTONUP:
        case WM_LBUTTONDBLCLK:
            // Left-click or double-click: restore window
            ShowWindowFromTray();
            break;

        case WM_RBUTTONUP:
            // Right-click: show context menu
            ShowTrayContextMenu();
            break;
        }
        return 0;
    }

    case WM_SIZE: {
        // Handle minimize to tray
        if (wParam == SIZE_MINIMIZED) {
            HideWindowToTray();
            return 0;
        }
        // Resize controls when window is resized
        RECT rcClient;
        GetClientRect(hwnd, &rcClient);
        int width = rcClient.right - rcClient.left;
        int height = rcClient.bottom - rcClient.top;

        // Position preset controls (always at top)
        SetWindowPos(g_hPresetLabel, nullptr, 10, 10, 50, 20, SWP_NOZORDER);
        SetWindowPos(g_hPresetCombo, nullptr, 60, 7, 200, 200, SWP_NOZORDER);
        SetWindowPos(g_hSavePresetBtn, nullptr, 270, 7, 100, 25, SWP_NOZORDER);
        SetWindowPos(g_hLoadPresetBtn, nullptr, 375, 7, 100, 25, SWP_NOZORDER);
        SetWindowPos(g_hDeletePresetBtn, nullptr, 480, 7, 110, 25, SWP_NOZORDER);

        // Adjust control positions and sizes based on whether process list is visible
        // Added 30 pixels to account for preset controls
        int topOffset = g_supportsProcessCapture ? 275 : 40;

        if (g_supportsProcessCapture) {
            SetWindowPos(g_hProcessListLabel, nullptr, 10, 40, 200, 20, SWP_NOZORDER);
            SetWindowPos(g_hProcessList, nullptr, 10, 60, width - 20, 180, SWP_NOZORDER);
            SetWindowPos(g_hRefreshBtn, nullptr, 10, 245, 100, 25, SWP_NOZORDER);
            SetWindowPos(g_hShowAudioOnlyCheckbox, nullptr, 120, 248, 280, 20, SWP_NOZORDER);
        } else {
            // Show format label when process list is hidden
            SetWindowPos(g_hFormatLabel, nullptr, 10, topOffset - 20, 60, 20, SWP_NOZORDER);
        }

        SetWindowPos(g_hFormatCombo, nullptr, 10, topOffset, 100, 25, SWP_NOZORDER);
        SetWindowPos(g_hOutputPathLabel, nullptr, 10, topOffset + 30, 100, 20, SWP_NOZORDER);
        SetWindowPos(g_hOutputPath, nullptr, 10, topOffset + 50, width - 100, 25, SWP_NOZORDER);
        SetWindowPos(g_hBrowseBtn, nullptr, width - 85, topOffset + 50, 75, 25, SWP_NOZORDER);
        SetWindowPos(g_hRecordingModeLabel, nullptr, 10, topOffset + 85, 140, 20, SWP_NOZORDER);
        SetWindowPos(g_hRecordingModeCombo, nullptr, 150, topOffset + 82, 150, 25, SWP_NOZORDER);
        SetWindowPos(g_hStartBtn, nullptr, 10, topOffset + 115, 210, 30, SWP_NOZORDER);
        SetWindowPos(g_hStopAllBtn, nullptr, 230, topOffset + 115, 100, 30, SWP_NOZORDER);
        SetWindowPos(g_hPauseAllBtn, nullptr, 340, topOffset + 115, 100, 30, SWP_NOZORDER);
        SetWindowPos(g_hResumeAllBtn, nullptr, 450, topOffset + 115, 100, 30, SWP_NOZORDER);
        SetWindowPos(g_hRecordingListLabel, nullptr, 10, topOffset + 155, 200, 20, SWP_NOZORDER);
        SetWindowPos(g_hRecordingList, nullptr, 10, topOffset + 175, width - 20, height - (topOffset + 225), SWP_NOZORDER);
        SetWindowPos(g_hStatusText, nullptr, 10, height - 40, width - 20, 30, SWP_NOZORDER);
        return 0;
    }

    case WM_DESTROY:
        RemoveTrayIcon();
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

    case WM_HSCROLL: {
        // Handle slider changes
        HWND hSlider = (HWND)lParam;

        if (hSlider == g_hProcessVolumeSlider) {
            int pos = (int)SendMessage(g_hProcessVolumeSlider, TBM_GETPOS, 0, 0);
            g_processVolume = (float)pos;

            wchar_t volumeText[64];
            swprintf_s(volumeText, L"Process Volume: %d%%", pos);
            SetWindowText(g_hProcessVolumeLabel, volumeText);
        }
        else if (hSlider == g_hMicrophoneVolumeSlider) {
            int pos = (int)SendMessage(g_hMicrophoneVolumeSlider, TBM_GETPOS, 0, 0);
            g_microphoneVolume = (float)pos;

            wchar_t volumeText[64];
            swprintf_s(volumeText, L"Microphone Volume: %d%%", pos);
            SetWindowText(g_hMicrophoneVolumeLabel, volumeText);
        }
        return 0;
    }
    }

    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

void InitializeControls(HWND hwnd) {
    // Determine visibility based on OS support
    DWORD processListVisibility = g_supportsProcessCapture ? (WS_CHILD | WS_VISIBLE) : WS_CHILD;

    // ========== Preset Controls (always visible) ==========
    // Preset label
    g_hPresetLabel = CreateWindow(
        L"STATIC", L"&Preset:",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        10, 10, 50, 20,
        hwnd, (HMENU)IDC_PRESET_LABEL, g_hInst, nullptr
    );

    // Preset combo box
    g_hPresetCombo = CreateWindow(
        WC_COMBOBOX, L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
        60, 7, 200, 200,
        hwnd, (HMENU)IDC_PRESET_COMBO, g_hInst, nullptr
    );

    // Save preset button
    g_hSavePresetBtn = CreateWindow(
        L"BUTTON", L"Sa&ve Preset",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        270, 7, 100, 25,
        hwnd, (HMENU)IDC_SAVE_PRESET_BTN, g_hInst, nullptr
    );

    // Load preset button
    g_hLoadPresetBtn = CreateWindow(
        L"BUTTON", L"&Load Preset",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        375, 7, 100, 25,
        hwnd, (HMENU)IDC_LOAD_PRESET_BTN, g_hInst, nullptr
    );

    // Delete preset button
    g_hDeletePresetBtn = CreateWindow(
        L"BUTTON", L"&Delete Preset",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        480, 7, 110, 25,
        hwnd, (HMENU)IDC_DELETE_PRESET_BTN, g_hInst, nullptr
    );

    // Process list label (moved down to make room for presets)
    g_hProcessListLabel = CreateWindow(
        L"STATIC", L"&Available processes:",
        processListVisibility | SS_LEFT,
        10, 40, 200, 20,
        hwnd, (HMENU)IDC_PROCESS_LIST_LABEL, g_hInst, nullptr
    );

    // Process list (ListView) - with checkboxes for multi-select (moved down for presets)
    g_hProcessList = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        WC_LISTVIEW,
        L"",
        processListVisibility | WS_TABSTOP | LVS_REPORT,
        10, 60, 760, 180,
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

    // Refresh button (adjusted position for preset controls)
    g_hRefreshBtn = CreateWindow(
        L"BUTTON", L"Refresh",
        processListVisibility | WS_TABSTOP | BS_PUSHBUTTON,
        10, 245, 100, 25,
        hwnd, (HMENU)IDC_REFRESH_BTN, g_hInst, nullptr
    );

    // Show audio only checkbox (adjusted position for preset controls)
    g_hShowAudioOnlyCheckbox = CreateWindow(
        L"BUTTON", L"Show only processes with active audio",
        processListVisibility | WS_TABSTOP | BS_AUTOCHECKBOX,
        120, 248, 280, 20,
        hwnd, (HMENU)IDC_SHOW_AUDIO_ONLY_CHECKBOX, g_hInst, nullptr
    );

    // Format label (only visible when process list is hidden)
    g_hFormatLabel = CreateWindow(
        L"STATIC", L"&Format:",
        WS_CHILD | (g_supportsProcessCapture ? 0 : WS_VISIBLE) | SS_LEFT,
        10, -15, 60, 20,
        hwnd, nullptr, g_hInst, nullptr
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
        L"BUTTON", L"&Monitor audio",
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
        L"BUTTON", L"Capture &inputs",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
        310, 330, 140, 20,
        hwnd, (HMENU)IDC_MICROPHONE_CHECKBOX, g_hInst, nullptr
    );

    // Microphone device label
    g_hMicrophoneDeviceLabel = CreateWindow(
        L"STATIC", L"Input devices:",
        WS_CHILD | SS_LEFT,
        460, 333, 90, 20,
        hwnd, (HMENU)IDC_MICROPHONE_DEVICE_LABEL, g_hInst, nullptr
    );

    // Microphone device list (multi-select with checkboxes)
    g_hMicrophoneDeviceList = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        WC_LISTVIEW,
        L"",
        WS_CHILD | WS_TABSTOP | LVS_REPORT,
        540, 327, 230, 80,
        hwnd, (HMENU)IDC_MICROPHONE_DEVICE_LIST, g_hInst, nullptr
    );
    ListView_SetExtendedListViewStyle(g_hMicrophoneDeviceList, LVS_EX_CHECKBOXES | LVS_EX_FULLROWSELECT);

    LVCOLUMN lvcMic = {};
    lvcMic.mask = LVCF_TEXT | LVCF_WIDTH;
    lvcMic.cx = 210;
    lvcMic.pszText = (LPWSTR)L"Microphones";
    ListView_InsertColumn(g_hMicrophoneDeviceList, 0, &lvcMic);

    // Process volume label
    g_hProcessVolumeLabel = CreateWindow(
        L"STATIC", L"Process Volume: 100%",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        10, 298, 140, 20,
        hwnd, (HMENU)IDC_PROCESS_VOLUME_LABEL, g_hInst, nullptr
    );

    // Process volume slider (0-100)
    g_hProcessVolumeSlider = CreateWindowEx(
        0, TRACKBAR_CLASS, L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | TBS_HORZ | TBS_AUTOTICKS,
        150, 298, 200, 20,
        hwnd, (HMENU)IDC_PROCESS_VOLUME_SLIDER, g_hInst, nullptr
    );
    SendMessage(g_hProcessVolumeSlider, TBM_SETRANGE, TRUE, MAKELONG(0, 100));
    SendMessage(g_hProcessVolumeSlider, TBM_SETPOS, TRUE, 100);
    SendMessage(g_hProcessVolumeSlider, TBM_SETTICFREQ, 10, 0);

    // Microphone volume label (initially hidden)
    g_hMicrophoneVolumeLabel = CreateWindow(
        L"STATIC", L"Microphone Volume: 100%",
        WS_CHILD | SS_LEFT,
        360, 298, 150, 20,
        hwnd, (HMENU)IDC_MICROPHONE_VOLUME_LABEL, g_hInst, nullptr
    );

    // Microphone volume slider (0-100, initially hidden)
    g_hMicrophoneVolumeSlider = CreateWindowEx(
        0, TRACKBAR_CLASS, L"",
        WS_CHILD | WS_TABSTOP | TBS_HORZ | TBS_AUTOTICKS,
        510, 298, 200, 20,
        hwnd, (HMENU)IDC_MICROPHONE_VOLUME_SLIDER, g_hInst, nullptr
    );
    SendMessage(g_hMicrophoneVolumeSlider, TBM_SETRANGE, TRUE, MAKELONG(0, 100));
    SendMessage(g_hMicrophoneVolumeSlider, TBM_SETPOS, TRUE, 100);
    SendMessage(g_hMicrophoneVolumeSlider, TBM_SETTICFREQ, 10, 0);

    // Output path label
    g_hOutputPathLabel = CreateWindow(
        L"STATIC", L"&Output Folder:",
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
        L"BUTTON", L"&Browse...",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        700, 295, 75, 25,
        hwnd, (HMENU)IDC_BROWSE_BTN, g_hInst, nullptr
    );

    // Start/stop toggle button
    g_hStartBtn = CreateWindow(
        L"BUTTON", L"&Start Capture",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        10, 360, 210, 30,
        hwnd, (HMENU)IDC_START_BTN, g_hInst, nullptr
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
    // If process capture is not supported, don't populate the list
    if (!g_supportsProcessCapture) {
        SetWindowText(g_hStatusText, L"System audio capture mode. Per-process capture not available on this OS version.");
        return;
    }

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

    // Sort processes naturally by name (case-insensitive)
    std::sort(g_processes.begin(), g_processes.end(),
        [](const ProcessInfo& a, const ProcessInfo& b) {
            return StrCmpLogicalW(a.processName.c_str(), b.processName.c_str()) < 0;
        });

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

    // If process capture is not supported, automatically capture system audio (PID 0)
    if (!g_supportsProcessCapture) {
        // Check if already capturing system audio
        if (g_captureManager->IsCapturing(0)) {
            MessageBox(g_hWnd, L"System audio is already being captured.", L"Already Capturing", MB_OK | MB_ICONINFORMATION);
            return;
        }
        // Force system audio capture
        checkedIndices.push_back(-1);  // Marker for system audio
    } else {
        // Normal process selection mode
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
                bool captureInputsOnly = (SendMessage(g_hMicrophoneCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
                if (!captureInputsOnly) {
                    MessageBox(g_hWnd, L"Please check one or more processes, or focus on a process to capture.", L"No Process Selected", MB_OK | MB_ICONWARNING);
                    return;
                }
            }
        }
    }

    // Get output path
    wchar_t outputPath[MAX_PATH];
    GetWindowText(g_hOutputPath, outputPath, MAX_PATH);
    std::wstring normalizedOutputPath = NormalizeOutputPath(outputPath);
    if (normalizedOutputPath != outputPath) {
        SetWindowText(g_hOutputPath, normalizedOutputPath.c_str());
    }

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

    // Check if input capture is enabled (counts as additional sources)
    bool captureMicrophone = (SendMessage(g_hMicrophoneCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);

    std::vector<size_t> micDeviceIndices;
    if (captureMicrophone && !monitorOnly && g_audioDeviceEnum) {
        micDeviceIndices = GetCheckedMicrophoneDeviceIndices();

        if (micDeviceIndices.empty()) {
            int focusedIndex = ListView_GetNextItem(g_hMicrophoneDeviceList, -1, LVNI_FOCUSED);
            if (focusedIndex >= 0) {
                LVITEM item = {};
                item.mask = LVIF_PARAM;
                item.iItem = focusedIndex;
                if (ListView_GetItem(g_hMicrophoneDeviceList, &item)) {
                    micDeviceIndices.push_back(static_cast<size_t>(item.lParam));
                }
            }
        }

        if (micDeviceIndices.empty()) {
            MessageBox(g_hWnd, L"Please check one or more input devices to capture.", L"No Input Device Selected", MB_OK | MB_ICONWARNING);
            return;
        }
    }

    // Calculate total number of audio sources (processes + microphone)
    size_t totalSources = checkedIndices.size();
    if (captureMicrophone && !monitorOnly) {
        totalSources += micDeviceIndices.size(); // Each input device counts as a source
    }

    // Get recording mode (only relevant if multiple sources and not monitor-only)
    int recordingModeIndex = (int)SendMessage(g_hRecordingModeCombo, CB_GETCURSEL, 0, 0);
    // 0 = Separate files, 1 = Combined file, 2 = Both

    bool createSeparateFiles = (totalSources == 1) || (recordingModeIndex == 0) || (recordingModeIndex == 2);
    bool createCombinedFile = (totalSources > 1) && !monitorOnly && ((recordingModeIndex == 1) || (recordingModeIndex == 2));

    if ((createSeparateFiles || createCombinedFile) && normalizedOutputPath.empty()) {
        MessageBox(g_hWnd, L"Please choose a valid output folder.", L"Invalid Output Folder", MB_OK | MB_ICONWARNING);
        return;
    }

    if ((createSeparateFiles || createCombinedFile) && !EnsureDirectoryExists(normalizedOutputPath)) {
        MessageBox(g_hWnd, L"The output folder could not be created or accessed.", L"Invalid Output Folder", MB_OK | MB_ICONWARNING);
        return;
    }

    int startedCount = 0;
    int alreadyCapturingCount = 0;

    for (int checkedIndex : checkedIndices) {
        std::wstring processName;
        DWORD processId;

        // If process capture not supported, use system audio defaults
        if (!g_supportsProcessCapture) {
            processName = L"[System Audio - All Processes]";
            processId = 0;
        } else {
            // Get process info from ListView (not from g_processes, because system audio isn't in that vector)
            wchar_t processNameBuf[256];
            wchar_t pidStrBuf[32];

            ListView_GetItemText(g_hProcessList, checkedIndex, 0, processNameBuf, 256);
            ListView_GetItemText(g_hProcessList, checkedIndex, 1, pidStrBuf, 32);

            processName = processNameBuf;
            processId = (DWORD)_wtoi(pidStrBuf);
        }

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
            std::wstring basePath = normalizedOutputPath;
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
            // Apply process volume setting (convert from 0-100 to 0.0-1.0)
            auto sessions = g_captureManager->GetActiveSessions();
            for (auto* session : sessions) {
                if (session->processId == processId && session->capture) {
                    session->capture->SetVolume(g_processVolume / 100.0f);
                    break;
                }
            }
            startedCount++;
        }
    }

    // If combined file mode is enabled, start the mixer
    if (createCombinedFile && startedCount > 0) {
        // Build combined file path
        std::wstring combinedPath = normalizedOutputPath;
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
        const auto& inputDevices = g_audioDeviceEnum->GetInputDevices();
        const DWORD kMicrophoneSessionBaseId = 0xFFFF0000;

        for (size_t micDeviceIndex : micDeviceIndices) {
            if (micDeviceIndex >= inputDevices.size()) {
                continue;
            }

            const auto& device = inputDevices[micDeviceIndex];
            std::wstring micDeviceId = device.deviceId;
            std::wstring micDeviceName = device.friendlyName;

            DWORD micProcessId = kMicrophoneSessionBaseId + static_cast<DWORD>(micDeviceIndex);

            if (g_captureManager->IsCapturing(micProcessId)) {
                alreadyCapturingCount++;
                continue;
            }

            bool createMicFile = createSeparateFiles;
            bool micMonitorOnly = false;
            std::wstring micFilePath;

            if (createCombinedFile && !createSeparateFiles) {
                micMonitorOnly = true;
            }

            if (createMicFile) {
                std::wstring basePath = normalizedOutputPath;
                if (basePath.back() != L'\\') {
                    basePath += L'\\';
                }

                SYSTEMTIME st;
                GetLocalTime(&st);
                wchar_t timestamp[64];
                swprintf_s(timestamp, L"%04d_%02d_%02d-%02d_%02d_%02d",
                    st.wYear, st.wMonth, st.wDay,
                    st.wHour, st.wMinute, st.wSecond);

                std::wstring deviceLabel = SanitizeFileName(micDeviceName);
                micFilePath = basePath + deviceLabel + L"-" + std::wstring(timestamp) + extension;
            } else {
                micFilePath = L"";
            }

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
                auto sessions = g_captureManager->GetActiveSessions();
                for (auto* session : sessions) {
                    if (session->processId == micProcessId && session->capture) {
                        session->capture->SetVolume(g_microphoneVolume / 100.0f);
                        break;
                    }
                }
                startedCount++;
            }
        }
    }

    if (startedCount > 0) {
        UpdateRecordingList();
        std::wstring status = L"Started " + std::to_wstring(startedCount) + L" capture(s)";
        if (alreadyCapturingCount > 0) {
            status += L" (" + std::to_wstring(alreadyCapturingCount) + L" already capturing)";
        }
        SetWindowText(g_hStatusText, status.c_str());
    }
    else if (alreadyCapturingCount > 0) {
        MessageBox(g_hWnd, L"All selected sources are already being captured.", L"Already Capturing", MB_OK | MB_ICONINFORMATION);
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
            g_captureManager->DisableMixedRecording(); // Stop mixed recording if no more sessions
        }

        // Restore focus to appropriate control to prevent keyboard focus issues
        if (g_supportsProcessCapture) {
            SetFocus(g_hProcessList);
        } else {
            SetFocus(g_hStartBtn);
        }
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
        UINT stateMask = LVIS_SELECTED;
        UINT state = LVIS_SELECTED;
        if (g_isAppActive) {
            stateMask |= LVIS_FOCUSED;
            state |= LVIS_FOCUSED;
        }
        ListView_SetItemState(g_hRecordingList, newSelectedIndex,
            state,
            stateMask);
    }
    EnsureRecordingListFocusItem();
    g_captureButtonStops = ListView_GetNextItem(g_hRecordingList, -1, LVNI_SELECTED) >= 0;
    SetWindowText(g_hStartBtn, g_captureButtonStops ? L"&Stop Capture" : L"&Start Capture");

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

void EnsureRecordingListFocusItem() {
    int focusedIndex = ListView_GetNextItem(g_hRecordingList, -1, LVNI_FOCUSED);
    if (focusedIndex < 0) {
        int selectedIndex = ListView_GetNextItem(g_hRecordingList, -1, LVNI_SELECTED);
        if (selectedIndex >= 0) {
            ListView_SetItemState(g_hRecordingList, selectedIndex,
                LVIS_FOCUSED, LVIS_FOCUSED);
        } else if (ListView_GetItemCount(g_hRecordingList) > 0) {
            ListView_SetItemState(g_hRecordingList, 0,
                LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
        }
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

std::wstring NormalizeOutputPath(const std::wstring& path) {
    std::wstring trimmed = path;

    auto isSpace = [](wchar_t ch) {
        return ch == L' ' || ch == L'\t' || ch == L'\r' || ch == L'\n';
    };

    while (!trimmed.empty() && isSpace(trimmed.front())) {
        trimmed.erase(trimmed.begin());
    }
    while (!trimmed.empty() && isSpace(trimmed.back())) {
        trimmed.pop_back();
    }

    if (trimmed.size() >= 2 && trimmed.front() == L'\"' && trimmed.back() == L'\"') {
        trimmed = trimmed.substr(1, trimmed.size() - 2);
    }

    while (!trimmed.empty() && isSpace(trimmed.front())) {
        trimmed.erase(trimmed.begin());
    }
    while (!trimmed.empty() && isSpace(trimmed.back())) {
        trimmed.pop_back();
    }

    return trimmed;
}

bool EnsureDirectoryExists(const std::wstring& path) {
    if (path.empty()) {
        return false;
    }

    int result = SHCreateDirectoryExW(nullptr, path.c_str(), nullptr);
    if (result == ERROR_SUCCESS || result == ERROR_ALREADY_EXISTS) {
        DWORD attrs = GetFileAttributesW(path.c_str());
        return (attrs != INVALID_FILE_ATTRIBUTES) && (attrs & FILE_ATTRIBUTE_DIRECTORY);
    }

    return false;
}

std::wstring SanitizeFileName(const std::wstring& name) {
    if (name.empty()) {
        return L"Device";
    }

    std::wstring sanitized = name;
    const wchar_t* invalidChars = L"\\/:*?\"<>|";
    for (auto& ch : sanitized) {
        if (wcschr(invalidChars, ch)) {
            ch = L'_';
        }
    }

    while (!sanitized.empty() && (sanitized.back() == L'.' || sanitized.back() == L' ')) {
        sanitized.pop_back();
    }

    return sanitized.empty() ? L"Device" : sanitized;
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

            // Load process volume
            if (settings.contains("processVolume") && settings["processVolume"].is_number()) {
                float volume = settings["processVolume"];
                if (volume >= 0.0f && volume <= 100.0f) {
                    g_processVolume = volume;
                    SendMessage(g_hProcessVolumeSlider, TBM_SETPOS, TRUE, (int)volume);
                    wchar_t volumeText[64];
                    swprintf_s(volumeText, L"Process Volume: %d%%", (int)volume);
                    SetWindowText(g_hProcessVolumeLabel, volumeText);
                }
            }

            // Load microphone volume
            if (settings.contains("microphoneVolume") && settings["microphoneVolume"].is_number()) {
                float volume = settings["microphoneVolume"];
                if (volume >= 0.0f && volume <= 100.0f) {
                    g_microphoneVolume = volume;
                    SendMessage(g_hMicrophoneVolumeSlider, TBM_SETPOS, TRUE, (int)volume);
                    wchar_t volumeText[64];
                    swprintf_s(volumeText, L"Microphone Volume: %d%%", (int)volume);
                    SetWindowText(g_hMicrophoneVolumeLabel, volumeText);
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

    // Save volume settings
    settings["processVolume"] = g_processVolume;
    settings["microphoneVolume"] = g_microphoneVolume;

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

    ListView_DeleteAllItems(g_hMicrophoneDeviceList);

    const auto& devices = g_audioDeviceEnum->GetInputDevices();
    int defaultIndex = -1;

    std::vector<size_t> sortedIndices;
    sortedIndices.reserve(devices.size());
    for (size_t i = 0; i < devices.size(); i++) {
        sortedIndices.push_back(i);
    }

    std::sort(sortedIndices.begin(), sortedIndices.end(),
        [&](size_t a, size_t b) {
            if (devices[a].isDefault != devices[b].isDefault) {
                return devices[a].isDefault;
            }
            return StrCmpLogicalW(devices[a].friendlyName.c_str(),
                                  devices[b].friendlyName.c_str()) < 0;
        });

    for (size_t listIndex = 0; listIndex < sortedIndices.size(); listIndex++) {
        size_t deviceIndex = sortedIndices[listIndex];
        const AudioDeviceInfo& device = devices[deviceIndex];

        std::wstring displayName = device.friendlyName;
        if (device.isDefault) {
            displayName = L"Default: " + displayName;
            defaultIndex = static_cast<int>(listIndex);
        }

        LVITEM lvi = {};
        lvi.mask = LVIF_TEXT | LVIF_PARAM;
        lvi.iItem = static_cast<int>(listIndex);
        lvi.pszText = (LPWSTR)displayName.c_str();
        lvi.lParam = static_cast<LPARAM>(deviceIndex);
        ListView_InsertItem(g_hMicrophoneDeviceList, &lvi);
    }

    if (!g_pendingMicrophoneDeviceIds.empty() || g_pendingMicrophoneDeviceIndex >= 0) {
        ApplyMicrophoneSelection(g_pendingMicrophoneDeviceIds, g_pendingMicrophoneDeviceIndex);
        g_pendingMicrophoneDeviceIds.clear();
        g_pendingMicrophoneDeviceIndex = -1;
    } else if (defaultIndex >= 0) {
        ListView_SetCheckState(g_hMicrophoneDeviceList, defaultIndex, TRUE);
    } else if (!devices.empty()) {
        ListView_SetCheckState(g_hMicrophoneDeviceList, 0, TRUE);
    }

    // Initially hide microphone controls
    OnMicrophoneCheckboxChanged();
}

void OnMicrophoneCheckboxChanged() {
    // Show/hide device selector and volume controls based on checkbox state
    BOOL isChecked = (SendMessage(g_hMicrophoneCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);

    ShowWindow(g_hMicrophoneDeviceLabel, isChecked ? SW_SHOW : SW_HIDE);
    ShowWindow(g_hMicrophoneDeviceList, isChecked ? SW_SHOW : SW_HIDE);
    ShowWindow(g_hMicrophoneVolumeLabel, isChecked ? SW_SHOW : SW_HIDE);
    ShowWindow(g_hMicrophoneVolumeSlider, isChecked ? SW_SHOW : SW_HIDE);
}

void AddTrayIcon() {
    // Initialize the NOTIFYICONDATA structure
    ZeroMemory(&g_nid, sizeof(NOTIFYICONDATA));
    g_nid.cbSize = sizeof(NOTIFYICONDATA);
    g_nid.hWnd = g_hWnd;
    g_nid.uID = IDI_TRAY_ICON;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    g_nid.hIcon = LoadIcon(nullptr, IDI_APPLICATION);

    // Set tooltip text
    const wchar_t* tooltip = g_supportsProcessCapture ?
        L"Audio Capture - Per-Process Recording" :
        L"Audio Capture - System Audio";
    wcscpy_s(g_nid.szTip, sizeof(g_nid.szTip) / sizeof(wchar_t), tooltip);

    // Add the icon to the system tray
    Shell_NotifyIcon(NIM_ADD, &g_nid);
    g_isMinimizedToTray = true;
}

void RemoveTrayIcon() {
    if (g_isMinimizedToTray) {
        Shell_NotifyIcon(NIM_DELETE, &g_nid);
        g_isMinimizedToTray = false;
    }
}

void ShowTrayContextMenu() {
    // Create a popup menu
    HMENU hMenu = CreatePopupMenu();
    if (!hMenu) return;

    // Add menu items
    AppendMenu(hMenu, MF_STRING, IDM_TRAY_SHOW, L"Show Window");
    AppendMenu(hMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenu(hMenu, MF_STRING, IDM_TRAY_EXIT, L"Exit");

    // Set default menu item (bold)
    SetMenuDefaultItem(hMenu, IDM_TRAY_SHOW, FALSE);

    // Get cursor position for menu display
    POINT pt;
    GetCursorPos(&pt);

    // Required for proper menu behavior with taskbar
    SetForegroundWindow(g_hWnd);

    // Show the menu
    TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, g_hWnd, nullptr);

    // Clean up
    DestroyMenu(hMenu);
}

void ShowWindowFromTray() {
    // Remove tray icon
    RemoveTrayIcon();

    // Restore the window
    ShowWindow(g_hWnd, SW_RESTORE);
    SetForegroundWindow(g_hWnd);
}

void HideWindowToTray() {
    // Hide the window
    ShowWindow(g_hWnd, SW_HIDE);

    // Add tray icon
    AddTrayIcon();
}

// ============================================================================
// Preset Management Functions
// ============================================================================

std::wstring GetPresetsDirectory() {
    wchar_t path[MAX_PATH];
    if (SUCCEEDED(SHGetFolderPath(nullptr, CSIDL_LOCAL_APPDATA, nullptr, 0, path))) {
        std::wstring presetsDir = path;
        presetsDir += L"\\AudioCapture\\presets";
        CreateDirectory((std::wstring(path) + L"\\AudioCapture").c_str(), nullptr);
        CreateDirectory(presetsDir.c_str(), nullptr);
        return presetsDir;
    }
    return L"presets";
}

std::vector<std::wstring> GetAvailablePresets() {
    std::vector<std::wstring> presets;
    std::wstring presetsDir = GetPresetsDirectory();
    std::wstring searchPath = presetsDir + L"\\*.json";

    WIN32_FIND_DATA findData;
    HANDLE hFind = FindFirstFile(searchPath.c_str(), &findData);

    if (hFind != INVALID_HANDLE_VALUE) {
        do {
            if (!(findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
                std::wstring filename = findData.cFileName;
                // Remove .json extension
                size_t ext = filename.rfind(L".json");
                if (ext != std::wstring::npos) {
                    presets.push_back(filename.substr(0, ext));
                }
            }
        } while (FindNextFile(hFind, &findData));
        FindClose(hFind);
    }

    return presets;
}

void PopulatePresetCombo() {
    SendMessage(g_hPresetCombo, CB_RESETCONTENT, 0, 0);

    std::vector<std::wstring> presets = GetAvailablePresets();
    for (const auto& preset : presets) {
        SendMessage(g_hPresetCombo, CB_ADDSTRING, 0, (LPARAM)preset.c_str());
    }

    if (!presets.empty()) {
        SendMessage(g_hPresetCombo, CB_SETCURSEL, 0, 0);
    }
}

std::vector<std::wstring> GetCheckedProcessNames() {
    std::vector<std::wstring> processNames;

    if (!g_supportsProcessCapture) {
        return processNames;
    }

    int itemCount = ListView_GetItemCount(g_hProcessList);
    for (int i = 0; i < itemCount; i++) {
        if (ListView_GetCheckState(g_hProcessList, i)) {
            wchar_t processName[256];
            ListView_GetItemText(g_hProcessList, i, 0, processName, 256);
            processNames.push_back(processName);
        }
    }

    return processNames;
}

std::vector<size_t> GetCheckedMicrophoneDeviceIndices() {
    std::vector<size_t> deviceIndices;

    if (!g_hMicrophoneDeviceList) {
        return deviceIndices;
    }

    int itemCount = ListView_GetItemCount(g_hMicrophoneDeviceList);
    for (int i = 0; i < itemCount; i++) {
        if (ListView_GetCheckState(g_hMicrophoneDeviceList, i)) {
            LVITEM item = {};
            item.mask = LVIF_PARAM;
            item.iItem = i;
            if (ListView_GetItem(g_hMicrophoneDeviceList, &item)) {
                deviceIndices.push_back(static_cast<size_t>(item.lParam));
            }
        }
    }

    return deviceIndices;
}

std::vector<std::wstring> GetCheckedMicrophoneDeviceIds() {
    std::vector<std::wstring> deviceIds;

    if (!g_audioDeviceEnum) {
        return deviceIds;
    }

    const auto& devices = g_audioDeviceEnum->GetInputDevices();
    std::vector<size_t> indices = GetCheckedMicrophoneDeviceIndices();
    for (size_t index : indices) {
        if (index < devices.size()) {
            deviceIds.push_back(devices[index].deviceId);
        }
    }

    return deviceIds;
}

void ApplyMicrophoneSelection(const std::vector<std::wstring>& deviceIds, int fallbackIndex) {
    if (!g_audioDeviceEnum || !g_hMicrophoneDeviceList) {
        return;
    }

    const auto& devices = g_audioDeviceEnum->GetInputDevices();
    int itemCount = ListView_GetItemCount(g_hMicrophoneDeviceList);

    for (int i = 0; i < itemCount; i++) {
        ListView_SetCheckState(g_hMicrophoneDeviceList, i, FALSE);
    }

    bool appliedSelection = false;
    if (!deviceIds.empty()) {
        for (int i = 0; i < itemCount; i++) {
            LVITEM item = {};
            item.mask = LVIF_PARAM;
            item.iItem = i;
            if (ListView_GetItem(g_hMicrophoneDeviceList, &item)) {
                size_t deviceIndex = static_cast<size_t>(item.lParam);
                if (deviceIndex < devices.size()) {
                    const auto& deviceId = devices[deviceIndex].deviceId;
                    if (std::find(deviceIds.begin(), deviceIds.end(), deviceId) != deviceIds.end()) {
                        ListView_SetCheckState(g_hMicrophoneDeviceList, i, TRUE);
                        appliedSelection = true;
                    }
                }
            }
        }
    }

    if (!appliedSelection && fallbackIndex >= 0 && fallbackIndex < itemCount) {
        ListView_SetCheckState(g_hMicrophoneDeviceList, fallbackIndex, TRUE);
        appliedSelection = true;
    }

    if (!appliedSelection && itemCount > 0) {
        ListView_SetCheckState(g_hMicrophoneDeviceList, 0, TRUE);
    }
}

void CheckProcessesByNames(const std::vector<std::wstring>& names) {
    if (!g_supportsProcessCapture || names.empty()) {
        return;
    }

    // First, uncheck all processes
    int itemCount = ListView_GetItemCount(g_hProcessList);
    for (int i = 0; i < itemCount; i++) {
        ListView_SetCheckState(g_hProcessList, i, FALSE);
    }

    // Then check processes that match the names
    for (int i = 0; i < itemCount; i++) {
        wchar_t processName[256];
        ListView_GetItemText(g_hProcessList, i, 0, processName, 256);

        for (const auto& name : names) {
            if (_wcsicmp(processName, name.c_str()) == 0) {
                ListView_SetCheckState(g_hProcessList, i, TRUE);
                break;
            }
        }
    }
}

json GetCurrentSettingsAsJson() {
    json settings;

    // Output path
    wchar_t outputPath[MAX_PATH];
    GetWindowTextW(g_hOutputPath, outputPath, MAX_PATH);
    settings["outputPath"] = WStringToString(outputPath);

    // Format and bitrates
    settings["format"] = (int)SendMessage(g_hFormatCombo, CB_GETCURSEL, 0, 0);
    settings["mp3Bitrate"] = (int)SendMessage(g_hMp3BitrateCombo, CB_GETCURSEL, 0, 0);
    settings["opusBitrate"] = (int)SendMessage(g_hOpusBitrateCombo, CB_GETCURSEL, 0, 0);
    settings["flacCompression"] = (int)SendMessage(g_hFlacCompressionCombo, CB_GETCURSEL, 0, 0);

    // Options
    settings["skipSilence"] = (SendMessage(g_hSkipSilenceCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
    settings["passthrough"] = (SendMessage(g_hPassthroughCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
    settings["passthroughDeviceIndex"] = (int)SendMessage(g_hPassthroughDeviceCombo, CB_GETCURSEL, 0, 0);
    settings["monitorOnly"] = (SendMessage(g_hMonitorOnlyCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
    settings["recordingMode"] = (int)SendMessage(g_hRecordingModeCombo, CB_GETCURSEL, 0, 0);
    settings["captureMicrophone"] = (SendMessage(g_hMicrophoneCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
    std::vector<std::wstring> micDeviceIds = GetCheckedMicrophoneDeviceIds();
    json micDeviceIdsJson = json::array();
    for (const auto& deviceId : micDeviceIds) {
        micDeviceIdsJson.push_back(WStringToString(deviceId));
    }
    settings["microphoneDeviceIds"] = micDeviceIdsJson;

    std::vector<size_t> micDeviceIndices = GetCheckedMicrophoneDeviceIndices();
    settings["microphoneDeviceIndex"] = micDeviceIndices.empty() ? -1 : static_cast<int>(micDeviceIndices.front());

    // Volumes
    settings["processVolume"] = g_processVolume;
    settings["microphoneVolume"] = g_microphoneVolume;

    // Process names
    std::vector<std::wstring> processNames = GetCheckedProcessNames();
    json processNamesJson = json::array();
    for (const auto& name : processNames) {
        processNamesJson.push_back(WStringToString(name));
    }
    settings["processNames"] = processNamesJson;

    return settings;
}

void ApplySettingsFromJson(const json& preset) {
    try {
        // Load output path
        if (preset.contains("outputPath") && preset["outputPath"].is_string()) {
            SetWindowTextW(g_hOutputPath, StringToWString(preset["outputPath"]).c_str());
        }

        // Load format
        if (preset.contains("format") && preset["format"].is_number_integer()) {
            int formatIndex = preset["format"];
            if (formatIndex >= 0 && formatIndex <= 3) {
                SendMessage(g_hFormatCombo, CB_SETCURSEL, formatIndex, 0);
            }
        }

        // Load bitrates
        if (preset.contains("mp3Bitrate") && preset["mp3Bitrate"].is_number_integer()) {
            int bitrateIndex = preset["mp3Bitrate"];
            if (bitrateIndex >= 0 && bitrateIndex <= 3) {
                SendMessage(g_hMp3BitrateCombo, CB_SETCURSEL, bitrateIndex, 0);
            }
        }

        if (preset.contains("opusBitrate") && preset["opusBitrate"].is_number_integer()) {
            int bitrateIndex = preset["opusBitrate"];
            if (bitrateIndex >= 0 && bitrateIndex <= 4) {
                SendMessage(g_hOpusBitrateCombo, CB_SETCURSEL, bitrateIndex, 0);
            }
        }

        if (preset.contains("flacCompression") && preset["flacCompression"].is_number_integer()) {
            int compressionIndex = preset["flacCompression"];
            if (compressionIndex >= 0 && compressionIndex <= 8) {
                SendMessage(g_hFlacCompressionCombo, CB_SETCURSEL, compressionIndex, 0);
            }
        }

        // Load options
        if (preset.contains("skipSilence") && preset["skipSilence"].is_boolean()) {
            SendMessage(g_hSkipSilenceCheckbox, BM_SETCHECK, preset["skipSilence"] ? BST_CHECKED : BST_UNCHECKED, 0);
        }

        if (preset.contains("passthrough") && preset["passthrough"].is_boolean()) {
            SendMessage(g_hPassthroughCheckbox, BM_SETCHECK, preset["passthrough"] ? BST_CHECKED : BST_UNCHECKED, 0);
        }

        if (preset.contains("passthroughDeviceIndex") && preset["passthroughDeviceIndex"].is_number_integer()) {
            int deviceIndex = preset["passthroughDeviceIndex"];
            if (deviceIndex >= 0 && deviceIndex < SendMessage(g_hPassthroughDeviceCombo, CB_GETCOUNT, 0, 0)) {
                SendMessage(g_hPassthroughDeviceCombo, CB_SETCURSEL, deviceIndex, 0);
            }
        }

        if (preset.contains("monitorOnly") && preset["monitorOnly"].is_boolean()) {
            SendMessage(g_hMonitorOnlyCheckbox, BM_SETCHECK, preset["monitorOnly"] ? BST_CHECKED : BST_UNCHECKED, 0);
        }

        if (preset.contains("recordingMode") && preset["recordingMode"].is_number_integer()) {
            int recordingMode = preset["recordingMode"];
            if (recordingMode >= 0 && recordingMode <= 2) {
                SendMessage(g_hRecordingModeCombo, CB_SETCURSEL, recordingMode, 0);
            }
        }

        if (preset.contains("captureMicrophone") && preset["captureMicrophone"].is_boolean()) {
            SendMessage(g_hMicrophoneCheckbox, BM_SETCHECK, preset["captureMicrophone"] ? BST_CHECKED : BST_UNCHECKED, 0);
        }

        std::vector<std::wstring> micDeviceIds;
        if (preset.contains("microphoneDeviceIds") && preset["microphoneDeviceIds"].is_array()) {
            for (const auto& id : preset["microphoneDeviceIds"]) {
                if (id.is_string()) {
                    micDeviceIds.push_back(StringToWString(id));
                }
            }
        }

        int micDeviceIndex = -1;
        if (preset.contains("microphoneDeviceIndex") && preset["microphoneDeviceIndex"].is_number_integer()) {
            micDeviceIndex = preset["microphoneDeviceIndex"];
        }

        if (!micDeviceIds.empty() || micDeviceIndex >= 0) {
            if (ListView_GetItemCount(g_hMicrophoneDeviceList) > 0) {
                ApplyMicrophoneSelection(micDeviceIds, micDeviceIndex);
            } else {
                g_pendingMicrophoneDeviceIds = micDeviceIds;
                g_pendingMicrophoneDeviceIndex = micDeviceIndex;
            }
        }

        // Load volumes
        if (preset.contains("processVolume") && preset["processVolume"].is_number()) {
            float volume = preset["processVolume"];
            if (volume >= 0.0f && volume <= 100.0f) {
                g_processVolume = volume;
                SendMessage(g_hProcessVolumeSlider, TBM_SETPOS, TRUE, (int)volume);
                wchar_t volumeText[64];
                swprintf_s(volumeText, L"Process Volume: %d%%", (int)volume);
                SetWindowText(g_hProcessVolumeLabel, volumeText);
            }
        }

        if (preset.contains("microphoneVolume") && preset["microphoneVolume"].is_number()) {
            float volume = preset["microphoneVolume"];
            if (volume >= 0.0f && volume <= 100.0f) {
                g_microphoneVolume = volume;
                SendMessage(g_hMicrophoneVolumeSlider, TBM_SETPOS, TRUE, (int)volume);
                wchar_t volumeText[64];
                swprintf_s(volumeText, L"Microphone Volume: %d%%", (int)volume);
                SetWindowText(g_hMicrophoneVolumeLabel, volumeText);
            }
        }

        // Load process names and check them
        if (preset.contains("processNames") && preset["processNames"].is_array()) {
            std::vector<std::wstring> processNames;
            for (const auto& name : preset["processNames"]) {
                if (name.is_string()) {
                    processNames.push_back(StringToWString(name));
                }
            }
            CheckProcessesByNames(processNames);
        }

        // Update UI visibility based on loaded settings
        OnFormatChanged();
        OnPassthroughCheckboxChanged();
        OnMonitorOnlyCheckboxChanged();
        OnMicrophoneCheckboxChanged();

    } catch (...) {
        MessageBox(g_hWnd, L"Error loading preset settings.", L"Preset Error", MB_OK | MB_ICONERROR);
    }
}

// Validation function for preset names
bool IsValidPresetName(const std::wstring& name) {
    if (name.empty() || name.length() > 100) {
        return false;
    }

    // Check for invalid filename characters
    const wchar_t* invalidChars = L"\\/:*?\"<>|";
    for (wchar_t c : name) {
        if (wcschr(invalidChars, c) != nullptr) {
            return false;
        }
    }

    return true;
}

// Dialog procedure for preset name input
LRESULT CALLBACK PresetNameDialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    static wchar_t* pPresetName = nullptr;
    static HWND hEdit = nullptr;

    switch (uMsg) {
    case WM_CREATE: {
        CREATESTRUCT* pCreate = (CREATESTRUCT*)lParam;
        pPresetName = (wchar_t*)pCreate->lpCreateParams;

        // Create controls
        CreateWindow(L"STATIC", L"Preset Name:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            10, 10, 100, 20,
            hwndDlg, nullptr, g_hInst, nullptr);

        hEdit = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_LEFT | ES_AUTOHSCROLL,
            10, 35, 290, 22,
            hwndDlg, (HMENU)100, g_hInst, nullptr);

        CreateWindow(L"BUTTON", L"OK",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
            100, 70, 80, 25,
            hwndDlg, (HMENU)IDOK, g_hInst, nullptr);

        CreateWindow(L"BUTTON", L"Cancel",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
            190, 70, 80, 25,
            hwndDlg, (HMENU)IDCANCEL, g_hInst, nullptr);

        SetFocus(hEdit);
        return 0;
    }

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDOK: {
            GetWindowText(hEdit, pPresetName, 256);

            // Validate the name
            if (!IsValidPresetName(pPresetName)) {
                MessageBox(hwndDlg,
                    L"Invalid preset name.\n\nPreset names cannot:\n- Be empty\n- Contain: \\ / : * ? \" < > |\n- Be longer than 100 characters",
                    L"Invalid Name", MB_OK | MB_ICONWARNING);
                SetFocus(hEdit);
                return 0;
            }

            // Signal successful completion
            SetWindowLongPtr(hwndDlg, GWLP_USERDATA, 1);
            DestroyWindow(hwndDlg);
            return 0;
        }

        case IDCANCEL:
            SetWindowLongPtr(hwndDlg, GWLP_USERDATA, 0);
            DestroyWindow(hwndDlg);
            return 0;
        }
        break;

    case WM_CLOSE:
        SetWindowLongPtr(hwndDlg, GWLP_USERDATA, 0);
        DestroyWindow(hwndDlg);
        return 0;
    }

    return DefWindowProc(hwndDlg, uMsg, wParam, lParam);
}

void SavePreset() {
    wchar_t presetName[256] = L"";

    // Register dialog window class (if not already registered)
    static bool classRegistered = false;
    const wchar_t* dialogClassName = L"PresetNameDialog";

    if (!classRegistered) {
        WNDCLASS wc = {};
        wc.lpfnWndProc = PresetNameDialogProc;
        wc.hInstance = g_hInst;
        wc.lpszClassName = dialogClassName;
        wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);

        if (RegisterClass(&wc)) {
            classRegistered = true;
        } else {
            MessageBox(g_hWnd, L"Failed to register dialog class.", L"Error", MB_OK | MB_ICONERROR);
            return;
        }
    }

    // Create the dialog window
    HWND hDlg = CreateWindowEx(
        WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
        dialogClassName,
        L"Save Preset",
        WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        (GetSystemMetrics(SM_CXSCREEN) - 320) / 2,
        (GetSystemMetrics(SM_CYSCREEN) - 130) / 2,
        320, 130,
        g_hWnd,
        nullptr,
        g_hInst,
        presetName  // Pass preset name buffer as creation parameter
    );

    if (!hDlg) {
        MessageBox(g_hWnd, L"Failed to create dialog.", L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Disable parent window
    EnableWindow(g_hWnd, FALSE);

    // Message loop for the dialog
    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        // Check if dialog still exists
        if (!IsWindow(hDlg)) {
            break;
        }

        if (!IsDialogMessage(hDlg, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }

        // Check if dialog was closed
        if (!IsWindow(hDlg)) {
            break;
        }
    }

    // Re-enable parent window
    EnableWindow(g_hWnd, TRUE);
    SetForegroundWindow(g_hWnd);

    // If user cancelled or preset name is empty, return
    if (wcslen(presetName) == 0) {
        return;
    }

    // Get current settings as JSON
    json preset = GetCurrentSettingsAsJson();

    // Build preset file path
    std::wstring presetsDir = GetPresetsDirectory();
    std::wstring presetPath = presetsDir + L"\\" + presetName + L".json";

    // Save to file
    std::ofstream file(presetPath);
    if (file.is_open()) {
        file << preset.dump(4);
        file.close();

        // Refresh preset combo
        PopulatePresetCombo();

        // Select the newly saved preset
        int index = (int)SendMessage(g_hPresetCombo, CB_FINDSTRINGEXACT, (WPARAM)-1, (LPARAM)presetName);
        if (index != CB_ERR) {
            SendMessage(g_hPresetCombo, CB_SETCURSEL, index, 0);
        }

        std::wstring statusMsg = L"Preset saved as: " + std::wstring(presetName);
        SetWindowText(g_hStatusText, statusMsg.c_str());
    } else {
        MessageBox(g_hWnd, L"Failed to save preset.", L"Error", MB_OK | MB_ICONERROR);
    }
}

void LoadPreset() {
    // Get selected preset name
    int index = (int)SendMessage(g_hPresetCombo, CB_GETCURSEL, 0, 0);
    if (index == CB_ERR) {
        MessageBox(g_hWnd, L"Please select a preset to load.", L"No Preset Selected", MB_OK | MB_ICONINFORMATION);
        return;
    }

    wchar_t presetName[256];
    SendMessage(g_hPresetCombo, CB_GETLBTEXT, index, (LPARAM)presetName);

    // Build preset file path
    std::wstring presetsDir = GetPresetsDirectory();
    std::wstring presetPath = presetsDir + L"\\" + presetName + L".json";

    // Load from file
    std::ifstream file(presetPath);
    if (file.is_open()) {
        try {
            json preset = json::parse(file);
            file.close();

            // Apply settings
            ApplySettingsFromJson(preset);

            std::wstring msg = L"Preset loaded: " + std::wstring(presetName);
            SetWindowText(g_hStatusText, msg.c_str());
        } catch (...) {
            file.close();
            MessageBox(g_hWnd, L"Failed to parse preset file.", L"Error", MB_OK | MB_ICONERROR);
        }
    } else {
        MessageBox(g_hWnd, L"Failed to load preset file.", L"Error", MB_OK | MB_ICONERROR);
    }
}

void DeletePreset() {
    // Get selected preset name
    int index = (int)SendMessage(g_hPresetCombo, CB_GETCURSEL, 0, 0);
    if (index == CB_ERR) {
        MessageBox(g_hWnd, L"Please select a preset to delete.", L"No Preset Selected", MB_OK | MB_ICONINFORMATION);
        return;
    }

    wchar_t presetName[256];
    SendMessage(g_hPresetCombo, CB_GETLBTEXT, index, (LPARAM)presetName);

    // Confirm deletion
    std::wstring msg = L"Delete preset \"" + std::wstring(presetName) + L"\"?";
    if (MessageBox(g_hWnd, msg.c_str(), L"Confirm Delete", MB_YESNO | MB_ICONQUESTION) != IDYES) {
        return;
    }

    // Build preset file path
    std::wstring presetsDir = GetPresetsDirectory();
    std::wstring presetPath = presetsDir + L"\\" + presetName + L".json";

    // Delete file
    if (DeleteFile(presetPath.c_str())) {
        // Refresh preset combo
        PopulatePresetCombo();

        std::wstring statusMsg = L"Preset deleted: " + std::wstring(presetName);
        SetWindowText(g_hStatusText, statusMsg.c_str());
    } else {
        MessageBox(g_hWnd, L"Failed to delete preset file.", L"Error", MB_OK | MB_ICONERROR);
    }
}
