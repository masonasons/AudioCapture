// Command-line audio capture tool
// Reference: https://github.com/bozbez/win-capture-audio
// Usage: AudioCaptureCLI.exe <process_name>

#include <windows.h>
#include <roapi.h>
#include <tlhelp32.h>
#include <iostream>
#include <string>
#include <signal.h>
#include <atomic>
#include <chrono>
#include <io.h>
#include <fcntl.h>

#include "AudioCapture.h"
#include "WavWriter.h"

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "avrt.lib")
#pragma comment(lib, "ksuser.lib")
#pragma comment(lib, "propsys.lib")

// Global flag for graceful shutdown
std::atomic<bool> g_running(true);
std::atomic<uint64_t> g_bytesWritten(0);

// Signal handler for Ctrl+C
void SignalHandler(int signal) {
    if (signal == SIGINT) {
        std::wcout << L"\nStopping..." << std::endl;
        g_running = false;
    }
}

// Find process ID by name (case-insensitive, supports partial match)
DWORD FindProcessByName(const std::wstring& processName, std::wstring& foundName) {
    HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnapshot == INVALID_HANDLE_VALUE) {
        return 0;
    }

    PROCESSENTRY32W pe32;
    pe32.dwSize = sizeof(PROCESSENTRY32W);

    DWORD foundPid = 0;
    std::wstring searchName = processName;
    
    // Convert to lowercase for comparison
    for (auto& c : searchName) {
        c = towlower(c);
    }

    if (Process32FirstW(hSnapshot, &pe32)) {
        do {
            std::wstring currentName = pe32.szExeFile;
            std::wstring currentNameLower = currentName;
            for (auto& c : currentNameLower) {
                c = towlower(c);
            }

            // Exact match (with or without .exe)
            if (currentNameLower == searchName || 
                currentNameLower == searchName + L".exe") {
                foundPid = pe32.th32ProcessID;
                foundName = currentName;
                break;
            }

            // Partial match (contains)
            if (currentNameLower.find(searchName) != std::wstring::npos) {
                foundPid = pe32.th32ProcessID;
                foundName = currentName;
                // Don't break - continue looking for exact match
            }
        } while (Process32NextW(hSnapshot, &pe32));
    }

    CloseHandle(hSnapshot);
    return foundPid;
}

// Get current timestamp string
std::wstring GetTimestamp() {
    SYSTEMTIME st;
    GetLocalTime(&st);
    wchar_t timestamp[64];
    swprintf_s(timestamp, L"%04d%02d%02d_%02d%02d%02d",
        st.wYear, st.wMonth, st.wDay,
        st.wHour, st.wMinute, st.wSecond);
    return timestamp;
}

// Format bytes to human-readable string
std::wstring FormatBytes(uint64_t bytes) {
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

// Format duration
std::wstring FormatDuration(int seconds) {
    int hours = seconds / 3600;
    int minutes = (seconds % 3600) / 60;
    int secs = seconds % 60;
    
    wchar_t buffer[64];
    if (hours > 0) {
        swprintf_s(buffer, L"%02d:%02d:%02d", hours, minutes, secs);
    } else {
        swprintf_s(buffer, L"%02d:%02d", minutes, secs);
    }
    return buffer;
}

int wmain(int argc, wchar_t* argv[]) {
    // Set console to Unicode mode
    _setmode(_fileno(stdout), _O_U16TEXT);
    _setmode(_fileno(stderr), _O_U16TEXT);

    std::wcout << L"========================================\n";
    std::wcout << L"  Process Audio Capture Tool\n";
    std::wcout << L"========================================\n";

    if (argc < 2) {
        std::wcout << L"\nUsage: " << argv[0] << L" <process_name>\n";
        std::wcout << L"Example: " << argv[0] << L" chrome\n";
        std::wcout << L"\nPress Ctrl+C to stop recording\n";
        return 1;
    }

    std::wstring processName = argv[1];

    // Initialize COM/WinRT
    HRESULT hr = RoInitialize(RO_INIT_SINGLETHREADED);
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE && hr != S_FALSE) {
        std::wcerr << L"Error: Failed to initialize COM/WinRT" << std::endl;
        return 1;
    }

    // Find process
    std::wstring foundProcessName;
    DWORD processId = FindProcessByName(processName, foundProcessName);
    
    if (processId == 0) {
        std::wcerr << L"\nError: Process not found '" << processName << L"'\n";
        RoUninitialize();
        return 1;
    }
    
    std::wcout << L"\n[Target] " << foundProcessName << L" (PID: " << processId << L")\n";

    // Setup output path: <process_name>_<timestamp>.wav
    std::wstring cleanName = foundProcessName;
    size_t exePos = cleanName.rfind(L".exe");
    if (exePos != std::wstring::npos) {
        cleanName = cleanName.substr(0, exePos);
    }
    std::wstring outputPath = cleanName + L"_" + GetTimestamp() + L".wav";

    std::wcout << L"[Output] " << outputPath << L"\n";

    // Initialize audio capture (process-specific only, no fallback to system audio)
    auto capture = std::make_unique<AudioCapture>();
    if (!capture->Initialize(processId)) {
        std::wcerr << L"\n";
        std::wcerr << L"*****************************************************\n";
        std::wcerr << L"* ERROR: Failed to capture process audio!           *\n";
        std::wcerr << L"*****************************************************\n\n";
        
        // Show the specific error
        const std::wstring& err = capture->GetLastError();
        if (!err.empty()) {
            std::wcerr << L"[Reason] " << err << L"\n\n";
        }
        
        std::wcerr << L"Requirements:\n";
        std::wcerr << L"  - Windows 10 Build 19041 (version 2004) or later\n";
        std::wcerr << L"  - Run 'winver' to check your Windows version\n";
        std::wcerr << L"  - Target process must be running\n";
        RoUninitialize();
        return 1;
    }

    std::wcout << L"[Mode] Process-specific capture (includes child processes)\n";

    const WAVEFORMATEX* waveFormat = capture->GetFormat();
    if (!waveFormat) {
        std::wcerr << L"Error: Failed to get audio format\n";
        RoUninitialize();
        return 1;
    }

    std::wcout << L"[Audio] " << waveFormat->nSamplesPerSec << L" Hz, " 
               << waveFormat->nChannels << L" channels\n";

    // Create WAV writer
    auto wavWriter = std::make_unique<WavWriter>();
    if (!wavWriter->Open(outputPath, waveFormat)) {
        std::wcerr << L"Error: Failed to create output file\n";
        RoUninitialize();
        return 1;
    }

    // Set up audio data callback
    capture->SetDataCallback([&](const BYTE* data, UINT32 size) {
        if (wavWriter->WriteData(data, size)) {
            g_bytesWritten += size;
        }
    });

    // Set up signal handler
    signal(SIGINT, SignalHandler);

    // Start capture
    std::wcout << L"\nRecording... Press Ctrl+C to stop\n\n";

    if (!capture->Start()) {
        std::wcerr << L"Error: Failed to start audio capture\n";
        RoUninitialize();
        return 1;
    }

    // Recording loop
    auto startTime = std::chrono::steady_clock::now();
    int lastPrintedSecond = -1;

    while (g_running) {
        Sleep(100);

        auto now = std::chrono::steady_clock::now();
        int elapsed = static_cast<int>(std::chrono::duration_cast<std::chrono::seconds>(now - startTime).count());

        if (elapsed != lastPrintedSecond) {
            lastPrintedSecond = elapsed;
            std::wcout << L"\r[Recording] " << FormatDuration(elapsed) 
                      << L"  " << FormatBytes(g_bytesWritten)
                      << L"          " << std::flush;
        }
    }

    // Stop and cleanup
    capture->Stop();
    wavWriter->Close();

    auto endTime = std::chrono::steady_clock::now();
    int totalSeconds = static_cast<int>(std::chrono::duration_cast<std::chrono::seconds>(endTime - startTime).count());

    std::wcout << L"\n\nRecording complete!\n";
    std::wcout << L"[File] " << outputPath << L"\n";
    std::wcout << L"[Duration] " << FormatDuration(totalSeconds) << L"\n";
    std::wcout << L"[Size] " << FormatBytes(g_bytesWritten) << L"\n";

    RoUninitialize();
    return 0;
}
