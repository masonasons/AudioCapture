#include "InputSourceManager.h"
#include "ProcessInputSource.h"
#include "SystemAudioInputSource.h"
#include "InputDeviceSource.h"
#include <sstream>
#include <iomanip>
#include <algorithm>

InputSourceManager::InputSourceManager()
    : m_processEnumerator(nullptr)
    , m_deviceEnumerator(nullptr)
{
    // Create enumerators
    m_processEnumerator = std::make_unique<ProcessEnumerator>();
    m_deviceEnumerator = std::make_unique<AudioDeviceEnumerator>();
}

InputSourceManager::~InputSourceManager() {
    // Enumerators are automatically cleaned up by unique_ptr
}

bool InputSourceManager::RefreshAvailableSources(bool includeProcesses,
                                                 bool includeSystemAudio,
                                                 bool includeInputDevices,
                                                 bool includeOutputDevices) {
    std::lock_guard<std::mutex> lock(m_mutex);

    // Clear existing sources
    m_availableSources.clear();

    try {
        // Enumerate system audio source (always available)
        if (includeSystemAudio) {
            EnumerateSystemAudioSource();
        }

        // Enumerate input devices (microphones, line-in)
        if (includeInputDevices) {
            if (m_deviceEnumerator->EnumerateInputDevices()) {
                EnumerateInputDeviceSources();
            }
        }

        // Enumerate output devices (speakers, for loopback capture)
        if (includeOutputDevices) {
            if (m_deviceEnumerator->EnumerateDevices()) {
                EnumerateOutputDeviceSources();
            }
        }

        // Enumerate processes (potentially slow, so do it last)
        if (includeProcesses) {
            EnumerateProcessSources();
        }

        return true;
    }
    catch (...) {
        // If enumeration fails, return what we have so far
        return false;
    }
}

void InputSourceManager::EnumerateSystemAudioSource() {
    // System audio is always available
    AvailableSource source;
    source.metadata.id = L"system:audio";
    source.metadata.displayName = L"System Audio (All Sounds)";
    source.metadata.type = InputSourceType::SystemAudio;
    source.metadata.iconHint = L"speaker";
    source.metadata.processId = 0;
    source.metadata.deviceId.clear();
    source.isAvailable = true;
    source.statusInfo = L"Ready";

    m_availableSources.push_back(source);
}

void InputSourceManager::EnumerateProcessSources() {
    if (!m_processEnumerator) {
        return;
    }

    // Get all running processes
    std::vector<ProcessInfo> processes = m_processEnumerator->GetAllProcesses();

    for (const auto& proc : processes) {
        AvailableSource source;

        // Generate source ID
        std::wostringstream oss;
        oss << L"process:" << proc.processId;
        source.metadata.id = oss.str();

        // Generate display name
        std::wostringstream displayOss;
        if (!proc.processName.empty()) {
            displayOss << proc.processName;
        } else {
            displayOss << L"Process " << proc.processId;
        }

        // Add window title if available
        if (!proc.windowTitle.empty()) {
            displayOss << L" - " << proc.windowTitle;
        }

        source.metadata.displayName = displayOss.str();
        source.metadata.type = InputSourceType::Process;
        source.metadata.iconHint = proc.processName;
        source.metadata.processId = proc.processId;
        source.metadata.deviceId.clear();

        // Process is available if we can enumerate it
        source.isAvailable = true;

        // Status info could include audio activity status, but that's expensive to check
        // For now, just mark as running
        source.statusInfo = L"Running";

        m_availableSources.push_back(source);
    }
}

void InputSourceManager::EnumerateInputDeviceSources() {
    if (!m_deviceEnumerator) {
        return;
    }

    const auto& devices = m_deviceEnumerator->GetInputDevices();

    for (const auto& device : devices) {
        AvailableSource source;

        // Generate source ID by hashing device ID
        std::hash<std::wstring> hasher;
        size_t hashValue = hasher(device.deviceId);
        std::wostringstream idOss;
        idOss << L"device:" << std::hex << std::setw(8) << std::setfill(L'0') << hashValue;
        source.metadata.id = idOss.str();

        // Display name without input indicator (removed for cleaner display)
        std::wstring displayName = device.friendlyName;

        // Remove "[Input]" or "input" from the display name (case-insensitive)
        std::wstring lowerName = displayName;
        std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::towlower);

        // Remove "[Input]" bracketed text (case-insensitive)
        size_t pos = lowerName.find(L"[input]");
        if (pos != std::wstring::npos) {
            displayName.erase(pos, 7);
        }

        // Remove standalone "input" word (with spaces around it)
        lowerName = displayName;
        std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::towlower);
        pos = lowerName.find(L" input ");
        if (pos != std::wstring::npos) {
            displayName.erase(pos, 7);
        }

        // Trim leading/trailing whitespace
        size_t start = displayName.find_first_not_of(L" \t");
        size_t end = displayName.find_last_not_of(L" \t");
        if (start != std::wstring::npos && end != std::wstring::npos) {
            displayName = displayName.substr(start, end - start + 1);
        }

        std::wostringstream nameOss;
        nameOss << displayName;
        if (device.isDefault) {
            nameOss << L" (Default)";
        }
        source.metadata.displayName = nameOss.str();

        source.metadata.type = InputSourceType::InputDevice;
        source.metadata.iconHint = L"microphone";
        source.metadata.processId = 0;
        source.metadata.deviceId = device.deviceId;

        source.isAvailable = true; // If it's in the enumeration, it's available
        source.statusInfo = device.isDefault ? L"Default Device" : L"Ready";

        m_availableSources.push_back(source);
    }
}

void InputSourceManager::EnumerateOutputDeviceSources() {
    if (!m_deviceEnumerator) {
        return;
    }

    const auto& devices = m_deviceEnumerator->GetDevices();

    for (const auto& device : devices) {
        AvailableSource source;

        // Generate source ID by hashing device ID
        std::hash<std::wstring> hasher;
        size_t hashValue = hasher(device.deviceId);
        std::wostringstream idOss;
        idOss << L"device:" << std::hex << std::setw(8) << std::setfill(L'0') << hashValue;
        source.metadata.id = idOss.str();

        // Display name with output indicator
        std::wostringstream nameOss;
        nameOss << L"[Output] " << device.friendlyName;
        if (device.isDefault) {
            nameOss << L" (Default)";
        }
        source.metadata.displayName = nameOss.str();

        source.metadata.type = InputSourceType::InputDevice;
        source.metadata.iconHint = L"speaker";
        source.metadata.processId = 0;
        source.metadata.deviceId = device.deviceId;

        source.isAvailable = true;
        source.statusInfo = device.isDefault ? L"Default Device" : L"Ready";

        m_availableSources.push_back(source);
    }
}

std::vector<AvailableSource> InputSourceManager::GetAvailableSources() const {
    std::lock_guard<std::mutex> lock(m_mutex);
    return m_availableSources; // Return a copy
}

std::vector<AvailableSource> InputSourceManager::GetSourcesByType(InputSourceType type) const {
    std::lock_guard<std::mutex> lock(m_mutex);

    std::vector<AvailableSource> filtered;
    std::copy_if(m_availableSources.begin(), m_availableSources.end(),
                 std::back_inserter(filtered),
                 [type](const AvailableSource& source) {
                     return source.metadata.type == type;
                 });

    return filtered;
}

InputSourcePtr InputSourceManager::CreateProcessSource(DWORD processId,
                                                       const std::wstring& processName,
                                                       const std::wstring& windowTitle) {
    // Look up process info if name not provided
    std::wstring actualProcessName = processName;
    std::wstring actualWindowTitle = windowTitle;

    if (actualProcessName.empty() || actualWindowTitle.empty()) {
        ProcessInfo info = LookupProcessInfo(processId);
        if (actualProcessName.empty()) {
            actualProcessName = info.processName;
        }
        if (actualWindowTitle.empty()) {
            actualWindowTitle = info.windowTitle;
        }
    }

    // Create the source
    try {
        auto source = std::make_shared<ProcessInputSource>(
            processId,
            actualProcessName,
            actualWindowTitle
        );
        return source;
    }
    catch (...) {
        return nullptr;
    }
}

InputSourcePtr InputSourceManager::CreateSystemAudioSource() {
    try {
        auto source = std::make_shared<SystemAudioInputSource>();
        return source;
    }
    catch (...) {
        return nullptr;
    }
}

InputSourcePtr InputSourceManager::CreateDeviceSource(const std::wstring& deviceId,
                                                      const std::wstring& friendlyName,
                                                      bool isInputDevice) {
    // Look up friendly name if not provided
    std::wstring actualFriendlyName = friendlyName;
    if (actualFriendlyName.empty()) {
        actualFriendlyName = LookupDeviceName(deviceId, isInputDevice);
    }

    // Create the source
    try {
        auto source = std::make_shared<InputDeviceSource>(
            deviceId,
            actualFriendlyName,
            isInputDevice
        );
        return source;
    }
    catch (...) {
        return nullptr;
    }
}

InputSourcePtr InputSourceManager::CreateSource(const AvailableSource& source) {
    switch (source.metadata.type) {
    case InputSourceType::Process:
        return CreateProcessSource(
            source.metadata.processId,
            source.metadata.iconHint,  // iconHint contains process name
            L""  // Window title will be looked up
        );

    case InputSourceType::SystemAudio:
        return CreateSystemAudioSource();

    case InputSourceType::InputDevice:
        {
            // Determine if it's an input or output device based on icon hint
            bool isInput = (source.metadata.iconHint == L"microphone");
            return CreateDeviceSource(
                source.metadata.deviceId,
                source.metadata.displayName,
                isInput
            );
        }

    default:
        return nullptr;
    }
}

ProcessInfo InputSourceManager::FindProcessInfo(DWORD processId) {
    return LookupProcessInfo(processId);
}

std::wstring InputSourceManager::LookupDeviceName(const std::wstring& deviceId,
                                                  bool isInputDevice) {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (!m_deviceEnumerator) {
        return L"";
    }

    // Search in the appropriate device list
    const auto& devices = isInputDevice ?
        m_deviceEnumerator->GetInputDevices() :
        m_deviceEnumerator->GetDevices();

    for (const auto& device : devices) {
        if (device.deviceId == deviceId) {
            return device.friendlyName;
        }
    }

    return L"Unknown Device";
}

ProcessInfo InputSourceManager::LookupProcessInfo(DWORD processId) {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (!m_processEnumerator) {
        ProcessInfo empty = {};
        empty.processId = processId;
        return empty;
    }

    // Get all processes and search for the one we want
    // Note: This is not very efficient, but ProcessEnumerator doesn't have
    // a direct lookup method. In a production system, you might want to add one.
    auto processes = m_processEnumerator->GetAllProcesses();

    for (const auto& proc : processes) {
        if (proc.processId == processId) {
            // Check if process has active audio
            ProcessInfo info = proc;
            info.hasActiveAudio = m_processEnumerator->CheckProcessHasActiveAudio(processId);
            return info;
        }
    }

    // Not found - return minimal info
    ProcessInfo info = {};
    info.processId = processId;
    info.processName = L"Unknown Process";
    info.hasActiveAudio = false;
    return info;
}
