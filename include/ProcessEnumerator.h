#pragma once

#include <windows.h>
#include <string>
#include <vector>

struct ProcessInfo {
    DWORD processId;
    std::wstring processName;
    std::wstring executablePath;
};

class ProcessEnumerator {
public:
    ProcessEnumerator();
    ~ProcessEnumerator();

    // Get list of all running processes with audio sessions
    std::vector<ProcessInfo> GetProcessesWithAudio();

    // Get list of all running processes
    std::vector<ProcessInfo> GetAllProcesses();

private:
    std::wstring GetProcessName(DWORD processId);
    std::wstring GetProcessPath(DWORD processId);
};
