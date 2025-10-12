#include "AudioCapture.h"
#include <iostream>
#include <fstream>
#include <string>

std::ofstream g_log;

void Log(const std::string& msg) {
    if (!g_log.is_open()) {
        g_log.open("C:\\Dropbox\\projects\\cpp\\AudioCapture\\capture_debug.log", std::ios::app);
    }
    g_log << msg << std::endl;
    g_log.flush();
}

int main(int argc, char* argv[]) {
    Log("=== Diagnostic Capture Test ===");

    DWORD processId = 0;
    if (argc > 1) {
        processId = (DWORD)atoi(argv[1]);
        Log("Testing with PID: " + std::to_string(processId));
    } else {
        Log("ERROR: No process ID provided");
        return 1;
    }

    AudioCapture capture;

    Log("Calling Initialize...");
    if (!capture.Initialize(processId)) {
        Log("ERROR: Initialize failed!");
        return 1;
    }

    Log("Initialize succeeded, checking if process-specific...");
    // We need to add a way to check m_isProcessSpecific

    Log("Starting capture...");
    if (!capture.Start()) {
        Log("ERROR: Start failed!");
        return 1;
    }

    Log("Capture started successfully, running for 3 seconds...");
    Sleep(3000);

    Log("Stopping capture...");
    capture.Stop();
    Log("Test complete");

    g_log.close();
    return 0;
}
