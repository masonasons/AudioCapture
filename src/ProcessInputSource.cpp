#include "ProcessInputSource.h"
#include <sstream>
#include <iomanip>

ProcessInputSource::ProcessInputSource(DWORD processId,
                                       const std::wstring& processName,
                                       const std::wstring& windowTitle)
    : m_processId(processId)
    , m_processName(processName)
    , m_windowTitle(windowTitle)
    , m_audioCapture(nullptr)
    , m_initialized(false)
{
    m_sourceId = GenerateSourceId();

    // Create AudioCapture instance but don't initialize yet
    // Initialization happens lazily in Initialize() or StartCapture()
    m_audioCapture = std::make_unique<AudioCapture>();
}

ProcessInputSource::~ProcessInputSource() {
    // Stop capture if still running
    StopCapture();

    // AudioCapture is automatically cleaned up by unique_ptr
}

std::wstring ProcessInputSource::GenerateSourceId() const {
    // Create unique ID in format "process:1234"
    std::wostringstream oss;
    oss << L"process:" << m_processId;
    return oss.str();
}

std::wstring ProcessInputSource::GenerateDisplayName() const {
    // Combine process name with window title for better identification
    std::wostringstream oss;

    // Start with process name
    if (!m_processName.empty()) {
        oss << m_processName;
    } else {
        oss << L"Process " << m_processId;
    }

    // Add window title if available
    if (!m_windowTitle.empty()) {
        oss << L" - " << m_windowTitle;
    }

    return oss.str();
}

InputSourceMetadata ProcessInputSource::GetMetadata() const {
    std::lock_guard<std::mutex> lock(m_mutex);

    InputSourceMetadata metadata;
    metadata.id = m_sourceId;
    metadata.displayName = GenerateDisplayName();
    metadata.type = InputSourceType::Process;
    metadata.iconHint = m_processName; // Can be used to load process icon
    metadata.processId = m_processId;
    metadata.deviceId.clear(); // Not applicable for process sources

    return metadata;
}

bool ProcessInputSource::Initialize() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_initialized) {
        return true; // Already initialized
    }

    if (!m_audioCapture) {
        return false;
    }

    // Initialize AudioCapture for this specific process
    // This will use the process loopback API internally
    if (!m_audioCapture->Initialize(m_processId)) {
        return false;
    }

    m_initialized = true;
    return true;
}

bool ProcessInputSource::StartCapture() {
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
        return true; // Already capturing - this is OK, return success
    }

    // Start the AudioCapture
    return m_audioCapture->Start();
}

void ProcessInputSource::StopCapture() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->Stop();
    }
}

bool ProcessInputSource::IsCapturing() const {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (!m_audioCapture) {
        return false;
    }

    return m_audioCapture->IsCapturing();
}

void ProcessInputSource::SetDataCallback(std::function<void(const BYTE*, UINT32)> callback) {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->SetDataCallback(callback);
    }
}

WAVEFORMATEX* ProcessInputSource::GetFormat() const {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (!m_audioCapture) {
        return nullptr;
    }

    return m_audioCapture->GetFormat();
}

void ProcessInputSource::SetVolume(float volume) {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->SetVolume(volume);
    }
}

void ProcessInputSource::Pause() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->Pause();
    }
}

void ProcessInputSource::Resume() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->Resume();
    }
}

bool ProcessInputSource::IsPaused() const {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (!m_audioCapture) {
        return false;
    }

    return m_audioCapture->IsPaused();
}
