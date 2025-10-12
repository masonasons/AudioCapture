#include "CaptureManager.h"
#include <algorithm>

CaptureManager::CaptureManager() {
}

CaptureManager::~CaptureManager() {
    StopAllCaptures();
}

bool CaptureManager::StartCapture(DWORD processId, const std::wstring& processName,
                                  const std::wstring& outputPath, AudioFormat format,
                                  UINT32 bitrate, bool skipSilence) {
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

    // Create audio capture
    session->capture = std::make_unique<AudioCapture>();
    if (!session->capture->Initialize(processId)) {
        return false;
    }

    // Create appropriate encoder
    const WAVEFORMATEX* waveFormat = session->capture->GetFormat();
    bool encoderReady = false;

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

bool CaptureManager::StopCapture(DWORD processId) {
    std::lock_guard<std::mutex> lock(m_mutex);

    auto it = m_sessions.find(processId);
    if (it == m_sessions.end()) {
        return false;
    }

    // Stop capture
    if (it->second->capture) {
        it->second->capture->Stop();
    }

    // Close encoder
    if (it->second->wavWriter) {
        it->second->wavWriter->Close();
    }
    if (it->second->mp3Encoder) {
        it->second->mp3Encoder->Close();
    }
    if (it->second->opusEncoder) {
        it->second->opusEncoder->Close();
    }
    if (it->second->flacEncoder) {
        it->second->flacEncoder->Close();
    }

    // Remove session
    m_sessions.erase(it);

    return true;
}

void CaptureManager::StopAllCaptures() {
    std::lock_guard<std::mutex> lock(m_mutex);

    for (auto& pair : m_sessions) {
        if (pair.second->capture) {
            pair.second->capture->Stop();
        }

        if (pair.second->wavWriter) {
            pair.second->wavWriter->Close();
        }
        if (pair.second->mp3Encoder) {
            pair.second->mp3Encoder->Close();
        }
        if (pair.second->opusEncoder) {
            pair.second->opusEncoder->Close();
        }
        if (pair.second->flacEncoder) {
            pair.second->flacEncoder->Close();
        }
    }

    m_sessions.clear();
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

    // Write data to appropriate encoder
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
