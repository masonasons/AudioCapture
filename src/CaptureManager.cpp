#include "CaptureManager.h"
#include <algorithm>
#include <chrono>

CaptureManager::CaptureManager()
    : m_mixedRecordingEnabled(false), m_mixerThreadRunning(false) {
}

CaptureManager::~CaptureManager() {
    DisableMixedRecording();
    StopAllCaptures();
}

bool CaptureManager::StartCapture(DWORD processId, const std::wstring& processName,
                                  const std::wstring& outputPath, AudioFormat format,
                                  UINT32 bitrate, bool skipSilence,
                                  const std::wstring& passthroughDeviceId,
                                  bool monitorOnly) {
    std::lock_guard<std::mutex> lock(m_mutex);

    // Check if already capturing this process
    if (m_sessions.find(processId) != m_sessions.end()) {
        return false;
    }

    // Create new session
    auto session = std::make_unique<CaptureSession>();
    session->processId = processId;
    session->processName = processName;
    session->outputFile = outputPath;
    session->format = format;
    session->isActive = false;
    session->bytesWritten = 0;
    session->skipSilence = skipSilence;
    session->monitorOnly = monitorOnly;

    // Create audio capture
    session->capture = std::make_unique<AudioCapture>();
    if (!session->capture->Initialize(processId)) {
        return false;
    }

    // Enable passthrough if device ID is provided
    if (!passthroughDeviceId.empty()) {
        if (!session->capture->EnablePassthrough(passthroughDeviceId)) {
            // Passthrough failed, but we can still continue with recording only
            // Could add a warning here if needed
        }
    }

    // Create appropriate encoder (skip if monitor-only mode)
    const WAVEFORMATEX* waveFormat = session->capture->GetFormat();
    bool encoderReady = monitorOnly; // If monitor-only, skip encoder setup

    if (!monitorOnly) {
        switch (format) {
        case AudioFormat::WAV:
            session->wavWriter = std::make_unique<WavWriter>();
            encoderReady = session->wavWriter->Open(outputPath, waveFormat);
            break;

        case AudioFormat::MP3:
            session->mp3Encoder = std::make_unique<Mp3Encoder>();
            // Use provided bitrate or default to 192000 (192 kbps)
            encoderReady = session->mp3Encoder->Open(outputPath, waveFormat,
                                                      bitrate > 0 ? bitrate : 192000);
            break;

        case AudioFormat::OPUS:
            session->opusEncoder = std::make_unique<OpusEncoder>();
            // Use provided bitrate or default to 128000 (128 kbps)
            encoderReady = session->opusEncoder->Open(outputPath, waveFormat,
                                                       bitrate > 0 ? bitrate : 128000);
            break;

        case AudioFormat::FLAC:
            session->flacEncoder = std::make_unique<FlacEncoder>();
            // Use bitrate as compression level (0-8), default to 5
            encoderReady = session->flacEncoder->Open(outputPath, waveFormat,
                                                       bitrate > 0 ? std::min(bitrate, 8u) : 5);
            break;
        }

        if (!encoderReady) {
            return false;
        }
    }

    // Set audio data callback
    session->capture->SetDataCallback([this, processId](const BYTE* data, UINT32 size) {
        OnAudioData(processId, data, size);
    });

    // Start capture
    if (!session->capture->Start()) {
        return false;
    }

    session->isActive = true;
    m_sessions[processId] = std::move(session);

    return true;
}

bool CaptureManager::StartCaptureFromDevice(DWORD sessionId, const std::wstring& deviceName,
                                            const std::wstring& deviceId, bool isInputDevice,
                                            const std::wstring& outputPath, AudioFormat format,
                                            UINT32 bitrate, bool skipSilence, bool monitorOnly) {
    std::lock_guard<std::mutex> lock(m_mutex);

    // Check if already capturing this session
    if (m_sessions.find(sessionId) != m_sessions.end()) {
        return false;
    }

    // Create new session
    auto session = std::make_unique<CaptureSession>();
    session->processId = sessionId;
    session->processName = deviceName;
    session->outputFile = outputPath;
    session->format = format;
    session->isActive = false;
    session->bytesWritten = 0;
    session->skipSilence = skipSilence;
    session->monitorOnly = monitorOnly;

    // Create audio capture for device
    session->capture = std::make_unique<AudioCapture>();
    if (!session->capture->InitializeFromDevice(deviceId, isInputDevice)) {
        return false;
    }

    // Create appropriate encoder (skip if monitor-only mode)
    const WAVEFORMATEX* waveFormat = session->capture->GetFormat();
    bool encoderReady = monitorOnly; // If monitor-only, skip encoder setup

    if (!monitorOnly) {
        switch (format) {
        case AudioFormat::WAV:
            session->wavWriter = std::make_unique<WavWriter>();
            encoderReady = session->wavWriter->Open(outputPath, waveFormat);
            break;

        case AudioFormat::MP3:
            session->mp3Encoder = std::make_unique<Mp3Encoder>();
            encoderReady = session->mp3Encoder->Open(outputPath, waveFormat,
                                                      bitrate > 0 ? bitrate : 192000);
            break;

        case AudioFormat::OPUS:
            session->opusEncoder = std::make_unique<OpusEncoder>();
            encoderReady = session->opusEncoder->Open(outputPath, waveFormat,
                                                       bitrate > 0 ? bitrate : 128000);
            break;

        case AudioFormat::FLAC:
            session->flacEncoder = std::make_unique<FlacEncoder>();
            encoderReady = session->flacEncoder->Open(outputPath, waveFormat,
                                                       bitrate > 0 ? std::min(bitrate, 8u) : 5);
            break;
        }

        if (!encoderReady) {
            return false;
        }
    }

    // Set audio data callback
    session->capture->SetDataCallback([this, sessionId](const BYTE* data, UINT32 size) {
        OnAudioData(sessionId, data, size);
    });

    // Start capture
    if (!session->capture->Start()) {
        return false;
    }

    session->isActive = true;
    m_sessions[sessionId] = std::move(session);

    return true;
}

bool CaptureManager::StopCapture(DWORD processId) {
    // Extract the session from the map (with mutex held)
    std::unique_ptr<CaptureSession> session;
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        auto it = m_sessions.find(processId);
        if (it == m_sessions.end()) {
            return false;
        }
        // Move the session out of the map
        session = std::move(it->second);
        m_sessions.erase(it);
    }

    // Stop capture WITHOUT holding the mutex (avoids deadlock with OnAudioData)
    if (session->capture) {
        session->capture->Stop();
    }

    // Close encoders
    if (session->wavWriter) {
        session->wavWriter->Close();
    }
    if (session->mp3Encoder) {
        session->mp3Encoder->Close();
    }
    if (session->opusEncoder) {
        session->opusEncoder->Close();
    }
    if (session->flacEncoder) {
        session->flacEncoder->Close();
    }

    // Session will be automatically destroyed when it goes out of scope
    return true;
}

void CaptureManager::StopAllCaptures() {
    // Get list of all session IDs first (with mutex held)
    std::vector<DWORD> sessionIds;
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        for (const auto& pair : m_sessions) {
            sessionIds.push_back(pair.first);
        }
    }

    // Stop each session WITHOUT holding the mutex (avoids deadlock)
    for (DWORD processId : sessionIds) {
        StopCapture(processId);
    }
}

std::vector<CaptureSession*> CaptureManager::GetActiveSessions() {
    std::lock_guard<std::mutex> lock(m_mutex);

    std::vector<CaptureSession*> sessions;
    for (auto& pair : m_sessions) {
        sessions.push_back(pair.second.get());
    }

    return sessions;
}

bool CaptureManager::IsCapturing(DWORD processId) const {
    std::lock_guard<std::mutex> lock(const_cast<std::mutex&>(m_mutex));
    return m_sessions.find(processId) != m_sessions.end();
}

void CaptureManager::OnAudioData(DWORD processId, const BYTE* data, UINT32 size) {
    std::lock_guard<std::mutex> lock(m_mutex);

    auto it = m_sessions.find(processId);
    if (it == m_sessions.end()) {
        return;
    }

    CaptureSession* session = it->second.get();

    // Check for silence if skip silence is enabled
    if (session->skipSilence && size > 0) {
        const WAVEFORMATEX* format = session->capture->GetFormat();
        if (format) {
            bool isSilent = true;
            UINT32 bytesPerSample = format->wBitsPerSample / 8;
            UINT32 numSamples = size / bytesPerSample;

            // Define silence threshold (very low amplitude)
            const int16_t SILENCE_THRESHOLD_16 = 50;  // ~0.15% of max amplitude
            const int32_t SILENCE_THRESHOLD_32 = 3276;  // ~0.01% of max amplitude

            if (bytesPerSample == 2) {
                // 16-bit samples
                const int16_t* samples = reinterpret_cast<const int16_t*>(data);
                for (UINT32 i = 0; i < numSamples; i++) {
                    if (abs(samples[i]) > SILENCE_THRESHOLD_16) {
                        isSilent = false;
                        break;
                    }
                }
            } else if (bytesPerSample == 4) {
                // 32-bit samples
                const int32_t* samples = reinterpret_cast<const int32_t*>(data);
                for (UINT32 i = 0; i < numSamples; i++) {
                    if (abs(samples[i]) > SILENCE_THRESHOLD_32) {
                        isSilent = false;
                        break;
                    }
                }
            }

            // Skip writing if silent
            if (isSilent) {
                return;
            }
        }
    }

    // Write data to appropriate encoder (skip if monitor-only mode)
    if (!session->monitorOnly) {
        bool success = false;
        switch (session->format) {
        case AudioFormat::WAV:
            if (session->wavWriter) {
                success = session->wavWriter->WriteData(data, size);
            }
            break;

        case AudioFormat::MP3:
            if (session->mp3Encoder) {
                success = session->mp3Encoder->WriteData(data, size);
            }
            break;

        case AudioFormat::OPUS:
            if (session->opusEncoder) {
                success = session->opusEncoder->WriteData(data, size);
            }
            break;

        case AudioFormat::FLAC:
            if (session->flacEncoder) {
                success = session->flacEncoder->WriteData(data, size);
            }
            break;
        }

        if (success) {
            session->bytesWritten += size;
        }
    }

    // If mixed recording is enabled, also send data to the mixer
    if (m_mixedRecordingEnabled && m_mixer) {
        m_mixer->AddAudioData(processId, data, size);
    }
}

bool CaptureManager::EnableMixedRecording(const std::wstring& outputPath, AudioFormat format, UINT32 bitrate) {
    std::lock_guard<std::mutex> lock(m_mixerMutex);

    if (m_mixedRecordingEnabled) {
        return false;  // Already enabled
    }

    // We need at least one session to get the audio format
    if (m_sessions.empty()) {
        return false;
    }

    // Get format from first session (all sessions should have the same format)
    const WAVEFORMATEX* waveFormat = m_sessions.begin()->second->capture->GetFormat();
    if (!waveFormat) {
        return false;
    }

    // Create mixer
    m_mixer = std::make_unique<AudioMixer>();
    if (!m_mixer->Initialize(waveFormat)) {
        m_mixer.reset();
        return false;
    }

    // Create appropriate encoder
    m_mixedFormat = format;
    bool encoderReady = false;

    switch (format) {
    case AudioFormat::WAV:
        m_mixedWavWriter = std::make_unique<WavWriter>();
        encoderReady = m_mixedWavWriter->Open(outputPath, waveFormat);
        break;

    case AudioFormat::MP3:
        m_mixedMp3Encoder = std::make_unique<Mp3Encoder>();
        encoderReady = m_mixedMp3Encoder->Open(outputPath, waveFormat,
                                                bitrate > 0 ? bitrate : 192000);
        break;

    case AudioFormat::OPUS:
        m_mixedOpusEncoder = std::make_unique<OpusEncoder>();
        encoderReady = m_mixedOpusEncoder->Open(outputPath, waveFormat,
                                                 bitrate > 0 ? bitrate : 128000);
        break;

    case AudioFormat::FLAC:
        m_mixedFlacEncoder = std::make_unique<FlacEncoder>();
        encoderReady = m_mixedFlacEncoder->Open(outputPath, waveFormat,
                                                 bitrate > 0 ? std::min(bitrate, 8u) : 5);
        break;
    }

    if (!encoderReady) {
        m_mixer.reset();
        m_mixedWavWriter.reset();
        m_mixedMp3Encoder.reset();
        m_mixedOpusEncoder.reset();
        m_mixedFlacEncoder.reset();
        return false;
    }

    // Start mixer thread
    m_mixerThreadRunning = true;
    m_mixerThread = std::make_unique<std::thread>(&CaptureManager::MixerThread, this);

    m_mixedRecordingEnabled = true;
    return true;
}

void CaptureManager::DisableMixedRecording() {
    // Check if already disabled and mark as disabled (must do this BEFORE joining thread)
    {
        std::lock_guard<std::mutex> lock(m_mixerMutex);
        if (!m_mixedRecordingEnabled) {
            return;
        }
        // Set to false NOW so OnAudioData stops adding new data to mixer
        m_mixedRecordingEnabled = false;
    }

    // Signal thread to stop (atomic, no lock needed)
    m_mixerThreadRunning = false;

    // Wait for mixer thread to finish
    if (m_mixerThread && m_mixerThread->joinable()) {
        m_mixerThread->join();
    }

    // Now clean up
    std::lock_guard<std::mutex> lock(m_mixerMutex);

    // Close encoders
    if (m_mixedWavWriter) {
        m_mixedWavWriter->Close();
        m_mixedWavWriter.reset();
    }
    if (m_mixedMp3Encoder) {
        m_mixedMp3Encoder->Close();
        m_mixedMp3Encoder.reset();
    }
    if (m_mixedOpusEncoder) {
        m_mixedOpusEncoder->Close();
        m_mixedOpusEncoder.reset();
    }
    if (m_mixedFlacEncoder) {
        m_mixedFlacEncoder->Close();
        m_mixedFlacEncoder.reset();
    }

    m_mixer.reset();
    m_mixerThread.reset();
}

bool CaptureManager::IsMixedRecordingActive() const {
    std::lock_guard<std::mutex> lock(const_cast<std::mutex&>(m_mixerMutex));
    return m_mixedRecordingEnabled;
}

void CaptureManager::MixerThread() {
    std::vector<BYTE> mixedBuffer;

    while (m_mixerThreadRunning) {
        bool hasData = false;
        AudioFormat currentFormat = AudioFormat::WAV;
        WavWriter* wavWriter = nullptr;
        Mp3Encoder* mp3Encoder = nullptr;
        OpusEncoder* opusEncoder = nullptr;
        FlacEncoder* flacEncoder = nullptr;

        // Get mixed audio data and encoder pointers (with lock held)
        {
            std::lock_guard<std::mutex> lock(m_mixerMutex);

            if (m_mixer) {
                hasData = m_mixer->GetMixedAudio(mixedBuffer);
                currentFormat = m_mixedFormat;

                // Get raw pointers to encoders (they're managed by unique_ptrs in CaptureManager)
                if (hasData && !mixedBuffer.empty()) {
                    wavWriter = m_mixedWavWriter.get();
                    mp3Encoder = m_mixedMp3Encoder.get();
                    opusEncoder = m_mixedOpusEncoder.get();
                    flacEncoder = m_mixedFlacEncoder.get();
                }
            }
        }

        // Write data to encoder WITHOUT lock held - encoding can be slow!
        if (hasData && !mixedBuffer.empty() && m_mixerThreadRunning) {
            switch (currentFormat) {
            case AudioFormat::WAV:
                if (wavWriter) {
                    wavWriter->WriteData(mixedBuffer.data(), static_cast<UINT32>(mixedBuffer.size()));
                }
                break;

            case AudioFormat::MP3:
                if (mp3Encoder) {
                    mp3Encoder->WriteData(mixedBuffer.data(), static_cast<UINT32>(mixedBuffer.size()));
                }
                break;

            case AudioFormat::OPUS:
                if (opusEncoder) {
                    opusEncoder->WriteData(mixedBuffer.data(), static_cast<UINT32>(mixedBuffer.size()));
                }
                break;

            case AudioFormat::FLAC:
                if (flacEncoder) {
                    flacEncoder->WriteData(mixedBuffer.data(), static_cast<UINT32>(mixedBuffer.size()));
                }
                break;
            }
        }

        if (!hasData) {
            // Sleep for a short time if no data available
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
        }
    }
}
