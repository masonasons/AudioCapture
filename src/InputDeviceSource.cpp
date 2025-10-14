#include "InputDeviceSource.h"
#include <sstream>
#include <iomanip>
#include <functional>
#include <algorithm>

InputDeviceSource::InputDeviceSource(const std::wstring& deviceId,
                                     const std::wstring& friendlyName,
                                     bool isInputDevice)
    : m_deviceId(deviceId)
    , m_friendlyName(friendlyName)
    , m_isInputDevice(isInputDevice)
    , m_audioCapture(nullptr)
    , m_initialized(false)
{
    m_sourceId = GenerateSourceId();

    // Create AudioCapture instance but don't initialize yet
    // Initialization happens lazily in Initialize() or StartCapture()
    m_audioCapture = std::make_unique<AudioCapture>();
}

InputDeviceSource::~InputDeviceSource() {
    // Stop capture if still running
    StopCapture();

    // AudioCapture is automatically cleaned up by unique_ptr
}

std::wstring InputDeviceSource::GenerateSourceId() const {
    // Create unique ID by hashing the device ID
    // For simplicity, we'll just use a prefix + truncated device ID
    // In production, you might want to use a proper hash function
    std::wostringstream oss;
    oss << L"device:";

    // Use std::hash to generate a numeric hash of the device ID
    std::hash<std::wstring> hasher;
    size_t hashValue = hasher(m_deviceId);
    oss << std::hex << std::setw(8) << std::setfill(L'0') << hashValue;

    return oss.str();
}

std::wstring InputDeviceSource::GenerateDisplayName() const {
    std::wstring displayName;

    // Add friendly name
    if (!m_friendlyName.empty()) {
        displayName = m_friendlyName;
    } else {
        displayName = L"Unknown Device";
    }

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
    } else if (displayName.empty()) {
        displayName = L"Unknown Device";
    }

    return displayName;
}

InputSourceMetadata InputDeviceSource::GetMetadata() const {
    std::lock_guard<std::mutex> lock(m_mutex);

    InputSourceMetadata metadata;
    metadata.id = m_sourceId;
    metadata.displayName = GenerateDisplayName();
    metadata.type = InputSourceType::InputDevice;
    metadata.iconHint = m_isInputDevice ? L"microphone" : L"speaker"; // Hint for UI icon
    metadata.processId = 0; // Not applicable for devices
    metadata.deviceId = m_deviceId;

    return metadata;
}

bool InputDeviceSource::Initialize() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_initialized) {
        return true; // Already initialized
    }

    if (!m_audioCapture) {
        return false;
    }

    // Initialize AudioCapture for this specific device
    // This will use either capture mode (for microphones) or loopback mode (for speakers)
    if (!m_audioCapture->InitializeFromDevice(m_deviceId, m_isInputDevice)) {
        return false;
    }

    m_initialized = true;
    return true;
}

bool InputDeviceSource::StartCapture() {
    // Initialize if not already done
    if (!m_initialized) {
        if (!Initialize()) {
            return false;
        }
    }

    std::lock_guard<std::mutex> lock(m_mutex);

    if (!m_audioCapture) {
        return false;
    }

    // Check if already capturing
    if (m_audioCapture->IsCapturing()) {
        return false; // Already capturing
    }

    // Start the AudioCapture
    return m_audioCapture->Start();
}

void InputDeviceSource::StopCapture() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->Stop();
    }
}

bool InputDeviceSource::IsCapturing() const {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (!m_audioCapture) {
        return false;
    }

    return m_audioCapture->IsCapturing();
}

void InputDeviceSource::SetDataCallback(std::function<void(const BYTE*, UINT32)> callback) {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->SetDataCallback(callback);
    }
}

WAVEFORMATEX* InputDeviceSource::GetFormat() const {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (!m_audioCapture) {
        return nullptr;
    }

    return m_audioCapture->GetFormat();
}

void InputDeviceSource::SetVolume(float volume) {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->SetVolume(volume);
    }
}

void InputDeviceSource::Pause() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->Pause();
    }
}

void InputDeviceSource::Resume() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->Resume();
    }
}

bool InputDeviceSource::IsPaused() const {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (!m_audioCapture) {
        return false;
    }

    return m_audioCapture->IsPaused();
}
