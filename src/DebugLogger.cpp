#include "DebugLogger.h"
#include <fstream>
#include <mutex>
#include <ctime>
#include <iomanip>

static std::mutex g_logMutex;

void DebugLog(const std::wstring& message) {
    std::lock_guard<std::mutex> lock(g_logMutex);
    std::wofstream logFile(L"AudioCapture_Debug.log", std::ios::app);
    if (logFile.is_open()) {
        auto now = std::time(nullptr);
        auto tm = *std::localtime(&now);
        logFile << L"[" << std::put_time(&tm, L"%H:%M:%S") << L"] " << message << std::endl;
        logFile.close();
    }
}
