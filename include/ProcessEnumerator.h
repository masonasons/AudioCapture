#pragma once

#include <windows.h>
#include <string>
#include <vector>

struct ProcessInfo {
    DWORD processId;
    std::wstring processName;
    std::wstring executablePath;
    std::wstring windowTitle;
    bool hasActiveAudio;
};

class ProcessEnumerator {
public:
    ProcessEnumerator();
    ~ProcessEnumerator();

    // Get list of all running processes with audio sessions
    std::vector<ProcessInfo> GetProcessesWithAudio();

    // Get list of all running processes
    std::vector<ProcessInfo> GetAllProcesses();

    // Get window title for a specific process (lazy loading)
    std::wstring GetWindowTitle(DWORD processId);

    // Check if process has active audio (lazy loading)
    bool CheckProcessHasActiveAudio(DWORD processId);

private:
    std::wstring GetProcessName(DWORD processId);
    std::wstring GetProcessPath(DWORD processId);
};
