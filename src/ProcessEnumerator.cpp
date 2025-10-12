#include "ProcessEnumerator.h"
#include <tlhelp32.h>
#include <psapi.h>
#include <algorithm>
#include <mmdeviceapi.h>
#include <audiopolicy.h>
#include <functiondiscoverykeys_devpkey.h>

ProcessEnumerator::ProcessEnumerator() {
}

ProcessEnumerator::~ProcessEnumerator() {
}

std::vector<ProcessInfo> ProcessEnumerator::GetAllProcesses() {
    std::vector<ProcessInfo> processes;

    // Create snapshot of all processes
    HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnapshot == INVALID_HANDLE_VALUE) {
        return processes;
    }

    PROCESSENTRY32W pe32;
    pe32.dwSize = sizeof(PROCESSENTRY32W);

    if (Process32FirstW(hSnapshot, &pe32)) {
        do {
            ProcessInfo info;
            info.processId = pe32.th32ProcessID;
            info.processName = pe32.szExeFile;
            info.executablePath = GetProcessPath(pe32.th32ProcessID);
            // Don't fetch window title and audio status during enumeration - too slow!
            // These will be fetched on-demand when displaying
            info.windowTitle = L"";
            info.hasActiveAudio = false;

            // Skip system processes
            if (info.processId > 0 && !info.processName.empty()) {
                processes.push_back(info);
            }
        } while (Process32NextW(hSnapshot, &pe32));
    }

    CloseHandle(hSnapshot);

    // Sort by process name
    std::sort(processes.begin(), processes.end(),
        [](const ProcessInfo& a, const ProcessInfo& b) {
            return a.processName < b.processName;
        });

    return processes;
}

std::vector<ProcessInfo> ProcessEnumerator::GetProcessesWithAudio() {
    // For simplicity, return all processes
    // A more advanced implementation would query audio sessions
    // using IAudioSessionManager2 to filter only processes with active audio
    return GetAllProcesses();
}

std::wstring ProcessEnumerator::GetProcessName(DWORD processId) {
    HANDLE hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, processId);
    if (hProcess == nullptr) {
        return L"";
    }

    wchar_t processName[MAX_PATH] = { 0 };
    DWORD size = MAX_PATH;
    if (QueryFullProcessImageNameW(hProcess, 0, processName, &size)) {
        // Extract just the filename
        std::wstring fullPath = processName;
        size_t lastSlash = fullPath.find_last_of(L"\\/");
        if (lastSlash != std::wstring::npos) {
            CloseHandle(hProcess);
            return fullPath.substr(lastSlash + 1);
        }
    }

    CloseHandle(hProcess);
    return L"";
}

std::wstring ProcessEnumerator::GetProcessPath(DWORD processId) {
    HANDLE hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, processId);
    if (hProcess == nullptr) {
        return L"";
    }

    wchar_t processPath[MAX_PATH] = { 0 };
    DWORD size = MAX_PATH;
    if (QueryFullProcessImageNameW(hProcess, 0, processPath, &size)) {
        CloseHandle(hProcess);
        return processPath;
    }

    CloseHandle(hProcess);
    return L"";
}

// Callback function for EnumWindows to find main window of a process
struct EnumWindowsCallbackData {
    DWORD processId;
    HWND hWnd;
};

BOOL CALLBACK EnumWindowsCallback(HWND hwnd, LPARAM lParam) {
    EnumWindowsCallbackData* data = reinterpret_cast<EnumWindowsCallbackData*>(lParam);

    DWORD windowProcessId;
    GetWindowThreadProcessId(hwnd, &windowProcessId);

    if (windowProcessId == data->processId) {
        // Check if this is a visible main window
        if (IsWindowVisible(hwnd) && GetWindow(hwnd, GW_OWNER) == nullptr) {
            // Check if window has a title
            int length = GetWindowTextLengthW(hwnd);
            if (length > 0) {
                data->hWnd = hwnd;
                return FALSE; // Stop enumeration
            }
        }
    }

    return TRUE; // Continue enumeration
}

std::wstring ProcessEnumerator::GetWindowTitle(DWORD processId) {
    EnumWindowsCallbackData data = { processId, nullptr };
    EnumWindows(EnumWindowsCallback, reinterpret_cast<LPARAM>(&data));

    if (data.hWnd != nullptr) {
        int length = GetWindowTextLengthW(data.hWnd);
        if (length > 0) {
            std::wstring title(length + 1, L'\0');
            GetWindowTextW(data.hWnd, &title[0], length + 1);
            title.resize(length); // Remove null terminator
            return title;
        }
    }

    return L"";
}

bool ProcessEnumerator::CheckProcessHasActiveAudio(DWORD processId) {
    // Initialize COM for this thread if needed
    CoInitialize(nullptr);

    bool hasAudio = false;
    IMMDeviceEnumerator* deviceEnumerator = nullptr;
    IMMDevice* device = nullptr;
    IAudioSessionManager2* sessionManager = nullptr;
    IAudioSessionEnumerator* sessionEnumerator = nullptr;

    // Create device enumerator
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr,
        CLSCTX_ALL, __uuidof(IMMDeviceEnumerator),
        reinterpret_cast<void**>(&deviceEnumerator));

    if (SUCCEEDED(hr)) {
        // Get default audio endpoint
        hr = deviceEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &device);

        if (SUCCEEDED(hr)) {
            // Get session manager
            hr = device->Activate(__uuidof(IAudioSessionManager2), CLSCTX_ALL,
                nullptr, reinterpret_cast<void**>(&sessionManager));

            if (SUCCEEDED(hr)) {
                // Get session enumerator
                hr = sessionManager->GetSessionEnumerator(&sessionEnumerator);

                if (SUCCEEDED(hr)) {
                    int sessionCount = 0;
                    sessionEnumerator->GetCount(&sessionCount);

                    // Check each session
                    for (int i = 0; i < sessionCount; i++) {
                        IAudioSessionControl* sessionControl = nullptr;
                        IAudioSessionControl2* sessionControl2 = nullptr;

                        if (SUCCEEDED(sessionEnumerator->GetSession(i, &sessionControl))) {
                            if (SUCCEEDED(sessionControl->QueryInterface(__uuidof(IAudioSessionControl2),
                                reinterpret_cast<void**>(&sessionControl2)))) {

                                DWORD sessionProcessId = 0;
                                if (SUCCEEDED(sessionControl2->GetProcessId(&sessionProcessId))) {
                                    if (sessionProcessId == processId) {
                                        // Check if session is active
                                        AudioSessionState state;
                                        if (SUCCEEDED(sessionControl2->GetState(&state))) {
                                            if (state == AudioSessionStateActive) {
                                                hasAudio = true;
                                            }
                                        }
                                    }
                                }

                                sessionControl2->Release();
                            }
                            sessionControl->Release();
                        }

                        if (hasAudio) break;
                    }

                    sessionEnumerator->Release();
                }

                sessionManager->Release();
            }

            device->Release();
        }

        deviceEnumerator->Release();
    }

    CoUninitialize();
    return hasAudio;
}
