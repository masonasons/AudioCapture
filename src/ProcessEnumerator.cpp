#include "ProcessEnumerator.h"
#include <tlhelp32.h>
#include <psapi.h>
#include <algorithm>

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
