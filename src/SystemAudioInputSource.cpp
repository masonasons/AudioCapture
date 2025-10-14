#include "SystemAudioInputSource.h"

SystemAudioInputSource::SystemAudioInputSource()
    : m_audioCapture(nullptr)
    , m_initialized(false)
{
    // Create AudioCapture instance but don't initialize yet
    // Initialization happens lazily in Initialize() or StartCapture()
    m_audioCapture = std::make_unique<AudioCapture>();
}

SystemAudioInputSource::~SystemAudioInputSource() {
    // Stop capture if still running
    StopCapture();

    // AudioCapture is automatically cleaned up by unique_ptr
}

InputSourceMetadata SystemAudioInputSource::GetMetadata() const {
    std::lock_guard<std::mutex> lock(m_mutex);

    InputSourceMetadata metadata;
    metadata.id = SOURCE_ID;
    metadata.displayName = DISPLAY_NAME;
    metadata.type = InputSourceType::SystemAudio;
    metadata.iconHint = L"speaker"; // Hint for UI to show speaker icon
    metadata.processId = 0; // Not applicable for system audio
    metadata.deviceId.clear(); // Not applicable for system audio

    return metadata;
}

bool SystemAudioInputSource::Initialize() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_initialized) {
        return true; // Already initialized
    }

    if (!m_audioCapture) {
        return false;
    }

    // Initialize AudioCapture with process ID 0 for system-wide capture
    // This uses the standard WASAPI loopback mode
    if (!m_audioCapture->Initialize(0)) {
        return false;
    }

    m_initialized = true;
    return true;
}

bool SystemAudioInputSource::StartCapture() {
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

void SystemAudioInputSource::StopCapture() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->Stop();
    }
}

bool SystemAudioInputSource::IsCapturing() const {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (!m_audioCapture) {
        return false;
    }

    return m_audioCapture->IsCapturing();
}

void SystemAudioInputSource::SetDataCallback(std::function<void(const BYTE*, UINT32)> callback) {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->SetDataCallback(callback);
    }
}

WAVEFORMATEX* SystemAudioInputSource::GetFormat() const {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (!m_audioCapture) {
        return nullptr;
    }

    return m_audioCapture->GetFormat();
}

void SystemAudioInputSource::SetVolume(float volume) {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->SetVolume(volume);
    }
}

void SystemAudioInputSource::Pause() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->Pause();
    }
}

void SystemAudioInputSource::Resume() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_audioCapture) {
        m_audioCapture->Resume();
    }
}

bool SystemAudioInputSource::IsPaused() const {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (!m_audioCapture) {
        return false;
    }

    return m_audioCapture->IsPaused();
}
