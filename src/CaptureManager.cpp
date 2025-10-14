#include "CaptureManager.h"
#include "ProcessInputSource.h"
#include "InputDeviceSource.h"
#include "SystemAudioInputSource.h"
#include "WavFileDestination.h"
#include "Mp3FileDestination.h"
#include "OpusFileDestination.h"
#include "FlacFileDestination.h"
#include "DeviceOutputDestination.h"
#include "DebugLogger.h"
#include <algorithm>

CaptureManager::CaptureManager()
    : m_nextSessionId(1) {
}

CaptureManager::~CaptureManager() {
    StopAll();
}

// ========================================
// New API Implementation
// ========================================

UINT32 CaptureManager::GenerateSessionId() {
    return m_nextSessionId.fetch_add(1);
}

UINT32 CaptureManager::StartCaptureSession(const CaptureConfig& config) {
    std::lock_guard<std::mutex> lock(m_mutex);

    // Validate configuration
    if (config.sources.empty()) {
        return 0;  // Need at least one source
    }

    // Create new internal session
    auto session = std::make_unique<CaptureSessionInternal>();
    session->sessionId = GenerateSessionId();
    session->sources = config.sources;
    session->destinations = config.destinations;
    session->routingRules = config.routingRules;
    session->isPaused = false;
    session->enableMixedOutput = config.enableMixedOutput;

    UINT32 sessionId = session->sessionId;

    // Set up callbacks and START sources first (needed to get format for mixer)
    for (auto& source : session->sources) {
        InputSourceMetadata metadata = source->GetMetadata();
        std::wstring sourceId = metadata.id;

        // Start capturing from this source FIRST (needed to get format)
        if (!source->StartCapture()) {
            // Failed to start capture from this source
            // Clean up and return failure
            for (auto& src : session->sources) {
                src->StopCapture();
            }
            return 0;
        }

        // Cache format pointer to avoid mutex + linear search in audio callbacks
        // This is THE critical optimization - callbacks run thousands of times per second!
        const WAVEFORMATEX* cachedFormat = source->GetFormat();

        // Set up callback with cached format - no mutex, no search!
        source->SetDataCallback([this, sessionId, sourceId, cachedFormat](const BYTE* data, UINT32 size) {
            if (cachedFormat) {
                OnAudioData(sessionId, sourceId, data, size, cachedFormat);
            }
        });
    }

    // Set up mixer if needed:
    // 1. Explicit mixed output requested (enableMixedOutput = true)
    // 2. Multiple sources with no routing rules (implicit mixing needed)
    bool needsMixer = config.enableMixedOutput ||
                     (config.sources.size() > 1 && config.routingRules.empty());

    if (needsMixer) {
        // Get format from first source (now valid after StartCapture)
        const WAVEFORMATEX* format = nullptr;
        if (!config.sources.empty()) {
            format = config.sources[0]->GetFormat();
        }

        if (format) {
            // Create mixer
            session->mixer = std::make_unique<AudioMixer>();
            if (!session->mixer->Initialize(format)) {
                // Clean up and return failure
                for (auto& src : session->sources) {
                    src->StopCapture();
                }
                return 0;  // Failed to initialize mixer
            }

            // Set the first source as the mixer driver (the one that calls GetMixedAudio)
            if (!session->sources.empty()) {
                session->mixerDriverSourceId = session->sources[0]->GetMetadata().id;
            }

            // If explicit mixed output requested, create dedicated mixed destination
            if (config.enableMixedOutput && !config.mixedOutputPath.empty()) {
                DestinationConfig destConfig;
                destConfig.outputPath = config.mixedOutputPath;
                destConfig.bitrate = config.mixedOutputBitrate;

                switch (config.mixedOutputFormat) {
                case AudioFormat::WAV: {
                    auto wavDest = std::make_unique<WavFileDestination>();
                    if (wavDest->Configure(format, destConfig)) {
                        session->mixedDestination = std::move(wavDest);
                    }
                    break;
                }
                case AudioFormat::MP3: {
                    auto mp3Dest = std::make_unique<Mp3FileDestination>();
                    if (mp3Dest->Configure(format, destConfig)) {
                        session->mixedDestination = std::move(mp3Dest);
                    }
                    break;
                }
                case AudioFormat::OPUS: {
                    auto opusDest = std::make_unique<OpusFileDestination>();
                    if (opusDest->Configure(format, destConfig)) {
                        session->mixedDestination = std::move(opusDest);
                    }
                    break;
                }
                case AudioFormat::FLAC: {
                    auto flacDest = std::make_unique<FlacFileDestination>();
                    if (flacDest->Configure(format, destConfig)) {
                        session->mixedDestination = std::move(flacDest);
                    }
                    break;
                }
                }

                if (!session->mixedDestination) {
                    // Clean up and return failure
                    for (auto& src : session->sources) {
                        src->StopCapture();
                    }
                    return 0;  // Failed to create mixed destination
                }
            }
        }
    }

    // Store session
    m_sessions[sessionId] = std::move(session);

    return sessionId;
}

bool CaptureManager::StopCaptureSession(UINT32 sessionId) {
    // Extract the session from the map (with mutex held)
    std::unique_ptr<CaptureSessionInternal> session;
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        auto it = m_sessions.find(sessionId);
        if (it == m_sessions.end()) {
            return false;
        }
        // Mark session as invalid BEFORE removing from map
        // This allows callbacks to detect session destruction via atomic flag
        it->second->isValid.store(false, std::memory_order_release);

        // Move the session out of the map
        session = std::move(it->second);
        m_sessions.erase(it);
    }

    // Stop all sources WITHOUT holding the mutex (avoids deadlock with callbacks)
    for (auto& source : session->sources) {
        source->StopCapture();
    }

    // Close all destinations
    for (auto& destination : session->destinations) {
        destination->Close();
    }

    // Close mixed destination if present
    if (session->mixedDestination) {
        session->mixedDestination->Close();
    }

    // Session will be automatically destroyed when it goes out of scope
    return true;
}

void CaptureManager::PauseSession(UINT32 sessionId) {
    std::lock_guard<std::mutex> lock(m_mutex);

    auto it = m_sessions.find(sessionId);
    if (it != m_sessions.end()) {
        it->second->isPaused = true;
        for (auto& source : it->second->sources) {
            source->Pause();
        }
    }
}

void CaptureManager::ResumeSession(UINT32 sessionId) {
    std::lock_guard<std::mutex> lock(m_mutex);

    auto it = m_sessions.find(sessionId);
    if (it != m_sessions.end()) {
        it->second->isPaused = false;
        for (auto& source : it->second->sources) {
            source->Resume();
        }
    }
}

bool CaptureManager::AddInputSource(UINT32 sessionId, InputSourcePtr source) {
    if (!source) {
        return false;
    }

    InputSourceMetadata metadata = source->GetMetadata();
    std::wstring sourceId = metadata.id;

    // CRITICAL FIX: Start capturing BEFORE acquiring mutex to avoid blocking audio callbacks
    // StartCapture() can take 10-100ms (WASAPI initialization, buffer allocation, etc.)
    // If we hold the mutex during this, ALL audio callbacks will be blocked causing choppy audio!
    if (!source->StartCapture()) {
        return false;  // Failed to start capture
    }

    // Brief delay to allow WASAPI to initialize format
    Sleep(10);

    // Now add source to session WITH mutex held (brief operation)
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        auto it = m_sessions.find(sessionId);
        if (it == m_sessions.end()) {
            // Session doesn't exist - stop the source and return
            source->StopCapture();
            return false;
        }

        // Add source to session's source list
        it->second->sources.push_back(source);
    }

    // THE FIX: Cache the format pointer in the callback capture to avoid mutex + linear search
    // The old code locked mutex and searched through ALL sources on EVERY audio callback!
    // That's why adding sources made everything choppy - O(n) search in hot path.
    const WAVEFORMATEX* cachedFormat = source->GetFormat();

    // Set up callback with cached format pointer
    source->SetDataCallback([this, sessionId, sourceId, cachedFormat](const BYTE* data, UINT32 size) {
        // No mutex needed! No search needed! Just use the cached format pointer.
        if (cachedFormat) {
            OnAudioData(sessionId, sourceId, data, size, cachedFormat);
        }
    });

    // Post-initialization: validate format and setup mixer if needed
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        auto it = m_sessions.find(sessionId);
        if (it == m_sessions.end()) {
            // Session was stopped - stop the source
            source->StopCapture();
            return false;
        }

        const WAVEFORMATEX* sourceFormat = source->GetFormat();

        // CRITICAL FIX: Create mixer if this is the 2nd source and no mixer exists yet!
        // When starting with 1 source + device output, no mixer is created
        // When adding a 2nd source, we need to create mixer to avoid audio conflicts
        if (!it->second->mixer && it->second->sources.size() >= 2 && it->second->routingRules.empty()) {
            // Get format from first source for mixer initialization
            const WAVEFORMATEX* mixerFormat = nullptr;
            if (!it->second->sources.empty()) {
                mixerFormat = it->second->sources[0]->GetFormat();
            }

            if (mixerFormat) {
                // Create mixer now!
                it->second->mixer = std::make_unique<AudioMixer>();
                if (it->second->mixer->Initialize(mixerFormat)) {
                    // Set the first source as the mixer driver (the one that calls GetMixedAudio)
                    if (!it->second->sources.empty()) {
                        it->second->mixerDriverSourceId = it->second->sources[0]->GetMetadata().id;
                    }

                    // Initialize buffers for ALL existing sources
                    for (const auto& existingSource : it->second->sources) {
                        std::wstring existingId = existingSource->GetMetadata().id;
                        DWORD mixerSourceId = std::hash<std::wstring>{}(existingId);
                        const WAVEFORMATEX* fmt = existingSource->GetFormat();
                        if (fmt) {
                            std::vector<BYTE> silentBuffer(fmt->nBlockAlign, 0);
                            it->second->mixer->AddAudioData(mixerSourceId, silentBuffer.data(),
                                                           fmt->nBlockAlign, fmt);
                        }
                    }
                } else {
                    // Failed to create mixer - remove the source and fail
                    it->second->mixer.reset();
                    auto& sources = it->second->sources;
                    auto sourceIt = std::find_if(sources.begin(), sources.end(),
                        [&sourceId](const InputSourcePtr& src) {
                            return src->GetMetadata().id == sourceId;
                        });
                    if (sourceIt != sources.end()) {
                        sources.erase(sourceIt);
                    }
                    source->StopCapture();
                    return false;
                }
            }
        }

        // Check for format mismatches if mixer exists
        if (sourceFormat && it->second->mixer) {
            // Get mixer format
            const WAVEFORMATEX* mixerFormat = &it->second->mixer->GetFormat();

            // Log format mismatch warning
            if (sourceFormat->nSamplesPerSec != mixerFormat->nSamplesPerSec ||
                sourceFormat->nChannels != mixerFormat->nChannels ||
                sourceFormat->wBitsPerSample != mixerFormat->wBitsPerSample) {

                // FORMAT MISMATCH DETECTED!
                OutputDebugStringW(L"[AudioCapture] WARNING: Format mismatch detected!\n");

                wchar_t buffer[512];
                swprintf_s(buffer, L"  Source: %u Hz, %u ch, %u bits\n  Mixer:  %u Hz, %u ch, %u bits\n",
                    sourceFormat->nSamplesPerSec, sourceFormat->nChannels, sourceFormat->wBitsPerSample,
                    mixerFormat->nSamplesPerSec, mixerFormat->nChannels, mixerFormat->wBitsPerSample);
                OutputDebugStringW(buffer);
            }
        }

        // If this session has a mixer, pre-register the source buffer
        if (it->second->mixer && sourceFormat) {
            DWORD mixerSourceId = std::hash<std::wstring>{}(sourceId);
            // Add a tiny amount of silence to initialize the buffer
            std::vector<BYTE> silentBuffer(sourceFormat->nBlockAlign, 0);
            it->second->mixer->AddAudioData(mixerSourceId, silentBuffer.data(),
                                           sourceFormat->nBlockAlign, sourceFormat);
        }
    }

    return true;
}

bool CaptureManager::RemoveInputSource(UINT32 sessionId, const std::wstring& sourceId) {
    // Extract the source from the session WITHOUT holding the mutex during StopCapture
    // to avoid deadlock with audio callbacks
    InputSourcePtr sourceToRemove;
    AudioMixer* mixer = nullptr;

    {
        std::lock_guard<std::mutex> lock(m_mutex);

        auto it = m_sessions.find(sessionId);
        if (it == m_sessions.end()) {
            return false;
        }

        auto& sources = it->second->sources;
        auto sourceIt = std::find_if(sources.begin(), sources.end(),
            [&sourceId](const InputSourcePtr& source) {
                return source->GetMetadata().id == sourceId;
            });

        if (sourceIt == sources.end()) {
            return false;
        }

        // Get mixer reference if present
        if (it->second->mixer) {
            mixer = it->second->mixer.get();
        }

        // Move the source out of the list (keeps it alive)
        sourceToRemove = *sourceIt;
        sources.erase(sourceIt);
    }

    // Remove from mixer if present (before stopping to drain remaining audio)
    if (mixer && sourceToRemove) {
        DWORD mixerSourceId = std::hash<std::wstring>{}(sourceId);
        mixer->RemoveSource(mixerSourceId);
    }

    // Stop the source WITHOUT holding the mutex
    // This allows audio callbacks to complete gracefully
    if (sourceToRemove) {
        sourceToRemove->StopCapture();
    }

    return true;
}

bool CaptureManager::AddOutputDestination(UINT32 sessionId, OutputDestinationPtr destination) {
    std::lock_guard<std::mutex> lock(m_mutex);

    auto it = m_sessions.find(sessionId);
    if (it == m_sessions.end()) {
        return false;
    }

    it->second->destinations.push_back(std::move(destination));
    return true;
}

bool CaptureManager::RemoveOutputDestination(UINT32 sessionId, const std::wstring& destinationId) {
    std::lock_guard<std::mutex> lock(m_mutex);

    auto it = m_sessions.find(sessionId);
    if (it == m_sessions.end()) {
        return false;
    }

    auto& destinations = it->second->destinations;
    auto destIt = std::find_if(destinations.begin(), destinations.end(),
        [&destinationId](const OutputDestinationPtr& dest) {
            return dest->GetName() == destinationId;
        });

    if (destIt == destinations.end()) {
        return false;
    }

    // Close the destination
    (*destIt)->Close();

    // Remove from list
    destinations.erase(destIt);
    return true;
}

bool CaptureManager::AddRoutingRule(UINT32 sessionId, const RoutingRule& rule) {
    std::lock_guard<std::mutex> lock(m_mutex);

    auto it = m_sessions.find(sessionId);
    if (it == m_sessions.end()) {
        return false;
    }

    it->second->routingRules.push_back(rule);
    return true;
}

void CaptureManager::StopAll() {
    // Get list of all session IDs first (with mutex held)
    std::vector<UINT32> sessionIds;
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        for (const auto& pair : m_sessions) {
            sessionIds.push_back(pair.first);
        }
    }

    // Stop each session WITHOUT holding the mutex (avoids deadlock)
    for (UINT32 sessionId : sessionIds) {
        StopCaptureSession(sessionId);
    }
}

void CaptureManager::PauseAll() {
    std::lock_guard<std::mutex> lock(m_mutex);

    for (auto& pair : m_sessions) {
        pair.second->isPaused = true;
        for (auto& source : pair.second->sources) {
            source->Pause();
        }
    }
}

void CaptureManager::ResumeAll() {
    std::lock_guard<std::mutex> lock(m_mutex);

    for (auto& pair : m_sessions) {
        pair.second->isPaused = false;
        for (auto& source : pair.second->sources) {
            source->Resume();
        }
    }
}

void CaptureManager::PauseFileDestinations() {
    std::lock_guard<std::mutex> lock(m_mutex);

    for (auto& pair : m_sessions) {
        for (auto& dest : pair.second->destinations) {
            // Only pause file destinations, not device destinations
            DestinationType type = dest->GetType();
            if (type == DestinationType::FileWAV || type == DestinationType::FileMP3 ||
                type == DestinationType::FileOpus || type == DestinationType::FileFLAC) {
                dest->Pause();
            }
        }
        // Also pause mixed output if it's a file
        if (pair.second->mixedDestination) {
            DestinationType type = pair.second->mixedDestination->GetType();
            if (type == DestinationType::FileWAV || type == DestinationType::FileMP3 ||
                type == DestinationType::FileOpus || type == DestinationType::FileFLAC) {
                pair.second->mixedDestination->Pause();
            }
        }
    }
}

void CaptureManager::ResumeFileDestinations() {
    std::lock_guard<std::mutex> lock(m_mutex);

    for (auto& pair : m_sessions) {
        for (auto& dest : pair.second->destinations) {
            // Only resume file destinations, not device destinations
            DestinationType type = dest->GetType();
            if (type == DestinationType::FileWAV || type == DestinationType::FileMP3 ||
                type == DestinationType::FileOpus || type == DestinationType::FileFLAC) {
                dest->Resume();
            }
        }
        // Also resume mixed output if it's a file
        if (pair.second->mixedDestination) {
            DestinationType type = pair.second->mixedDestination->GetType();
            if (type == DestinationType::FileWAV || type == DestinationType::FileMP3 ||
                type == DestinationType::FileOpus || type == DestinationType::FileFLAC) {
                pair.second->mixedDestination->Resume();
            }
        }
    }
}

void CaptureManager::OnAudioData(UINT32 sessionId, const std::wstring& sourceId,
                                 const BYTE* data, UINT32 size, const WAVEFORMATEX* format) {
    // DEBUG: Track audio data flow to diagnose 0-byte files
    static std::atomic<int> audioCallCount{0};
    static std::atomic<uint64_t> totalBytesReceived{0};
    int callNum = ++audioCallCount;
    totalBytesReceived += size;

    // Log every 100 calls to avoid spam
    if (callNum == 1 || callNum % 100 == 0) {
        wchar_t debugMsg[256];
        swprintf_s(debugMsg, L"OnAudioData #%d: SessionID=%u, Size=%u bytes, Total=%.2f MB",
                   callNum, sessionId, size, totalBytesReceived.load() / (1024.0 * 1024.0));
        DebugLog(debugMsg);
    }

    // CRITICAL OPTIMIZATION: Minimize mutex hold time to prevent callback blocking
    // We'll acquire mutex briefly just to validate session and copy necessary pointers,
    // then release it immediately before doing any expensive mixer operations

    CaptureSessionInternal* sessionPtr = nullptr;
    bool isPaused = false;
    bool useMixer = false;
    bool enableMixedOutput = false;
    bool hasRoutingRules = false;
    size_t sourceCount = 0;

    // Fast path: Quick validation and data extraction with mutex held
    {
        std::lock_guard<std::mutex> lock(m_mutex);

        auto it = m_sessions.find(sessionId);
        if (it == m_sessions.end()) {
            return;
        }

        sessionPtr = it->second.get();

        // Fast lock-free validity check
        if (!sessionPtr->isValid.load(std::memory_order_acquire)) {
            return;
        }

        isPaused = sessionPtr->isPaused;
        enableMixedOutput = sessionPtr->enableMixedOutput;
        hasRoutingRules = !sessionPtr->routingRules.empty();
        sourceCount = sessionPtr->sources.size();

        // Determine if we need mixer
        useMixer = sessionPtr->mixer &&
                   (enableMixedOutput || (sourceCount > 1 && !hasRoutingRules));
    }
    // CRITICAL: Mutex is now released - other callbacks can proceed immediately

    if (isPaused) {
        return;
    }

    // All mixer operations happen WITHOUT holding m_mutex
    // AudioMixer has its own internal mutex for thread safety
    if (useMixer) {
        // Add audio to mixer (AudioMixer's internal mutex protects this)
        DWORD mixerSourceId = std::hash<std::wstring>{}(sourceId);
        sessionPtr->mixer->AddAudioData(mixerSourceId, data, size, format);

        // CRITICAL FIX: Only call GetMixedAudio from the designated "mixer driver" source
        // If every source calls GetMixedAudio, we get a race condition where mixed audio
        // only succeeds intermittently when all sources happen to have data simultaneously
        bool shouldPullMixedAudio = (sourceId == sessionPtr->mixerDriverSourceId);

        // Get mixed audio and write to destinations (only from mixer driver source)
        std::vector<BYTE> mixedBuffer;
        if (shouldPullMixedAudio && sessionPtr->mixer->GetMixedAudio(mixedBuffer) && !mixedBuffer.empty()) {
            // If explicit mixed output, write to dedicated mixed destination
            if (enableMixedOutput && sessionPtr->mixedDestination && sessionPtr->mixedDestination->IsOpen()) {
                sessionPtr->mixedDestination->WriteAudioData(mixedBuffer.data(),
                                                            static_cast<UINT32>(mixedBuffer.size()));
            }

            // If implicit mixing (multiple sources, no routing), write to all regular destinations
            if (!enableMixedOutput && !hasRoutingRules) {
                // Quick lock to iterate destinations safely
                std::vector<OutputDestinationPtr> destinationsCopy;
                {
                    std::lock_guard<std::mutex> lock(m_mutex);
                    if (sessionPtr->isValid.load(std::memory_order_acquire)) {
                        destinationsCopy = sessionPtr->destinations;
                    }
                }

                for (auto& destination : destinationsCopy) {
                    if (destination && destination->IsOpen()) {
                        destination->WriteAudioData(mixedBuffer.data(), static_cast<UINT32>(mixedBuffer.size()));
                    }
                }
            }
        }
    } else {
        // No mixer needed (single source or routing rules in use)
        // Route audio directly to appropriate destinations
        RouteAudioData(sessionPtr, sourceId, data, size, format);
    }
}

void CaptureManager::RouteAudioData(CaptureSessionInternal* session, const std::wstring& sourceId,
                                   const BYTE* data, UINT32 size, const WAVEFORMATEX* format) {
    // If no routing rules, route to all destinations
    if (session->routingRules.empty()) {
        for (auto& destination : session->destinations) {
            if (destination->IsOpen()) {
                destination->WriteAudioData(data, size);
            }
        }
        return;
    }

    // Apply routing rules
    for (const auto& rule : session->routingRules) {
        // Check if this rule applies to this source
        if (!rule.sourceId.empty() && rule.sourceId != sourceId) {
            continue;  // Rule doesn't apply to this source
        }

        // Skip silence if requested
        if (rule.skipSilence && IsSilent(data, size, format)) {
            continue;
        }

        // Find matching destination
        for (auto& destination : session->destinations) {
            if (!rule.destinationId.empty() && destination->GetName() != rule.destinationId) {
                continue;  // Not the target destination
            }

            if (!destination->IsOpen()) {
                continue;  // Destination not open
            }

            // Apply volume adjustment if needed
            if (rule.volumeMultiplier != 1.0f) {
                // Make a copy of the data to apply volume
                std::vector<BYTE> modifiedData(data, data + size);
                ApplyVolume(modifiedData.data(), size, format, rule.volumeMultiplier);
                destination->WriteAudioData(modifiedData.data(), size);
            } else {
                destination->WriteAudioData(data, size);
            }
        }
    }
}

bool CaptureManager::IsSilent(const BYTE* data, UINT32 size, const WAVEFORMATEX* format) {
    if (!format || size == 0) {
        return true;
    }

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
                return false;
            }
        }
    } else if (bytesPerSample == 4) {
        // 32-bit samples
        const int32_t* samples = reinterpret_cast<const int32_t*>(data);
        for (UINT32 i = 0; i < numSamples; i++) {
            if (abs(samples[i]) > SILENCE_THRESHOLD_32) {
                return false;
            }
        }
    }

    return true;
}

void CaptureManager::ApplyVolume(BYTE* data, UINT32 size, const WAVEFORMATEX* format, float volume) {
    if (!format || volume == 1.0f) {
        return;
    }

    UINT32 bytesPerSample = format->wBitsPerSample / 8;
    UINT32 numSamples = size / bytesPerSample;

    if (bytesPerSample == 2) {
        // 16-bit samples
        int16_t* samples = reinterpret_cast<int16_t*>(data);
        for (UINT32 i = 0; i < numSamples; i++) {
            float sample = static_cast<float>(samples[i]) * volume;
            // Clamp to prevent overflow
            if (sample > 32767.0f) sample = 32767.0f;
            if (sample < -32768.0f) sample = -32768.0f;
            samples[i] = static_cast<int16_t>(sample);
        }
    } else if (bytesPerSample == 4) {
        // 32-bit samples (assume float)
        float* samples = reinterpret_cast<float*>(data);
        for (UINT32 i = 0; i < numSamples; i++) {
            samples[i] *= volume;
            // Clamp to prevent overflow
            if (samples[i] > 1.0f) samples[i] = 1.0f;
            if (samples[i] < -1.0f) samples[i] = -1.0f;
        }
    }
}

size_t CaptureManager::GetActiveSessionCount() const {
    std::lock_guard<std::mutex> lock(const_cast<std::mutex&>(m_mutex));
    return m_sessions.size();
}

bool CaptureManager::IsSessionActive(UINT32 sessionId) const {
    std::lock_guard<std::mutex> lock(const_cast<std::mutex&>(m_mutex));
    return m_sessions.find(sessionId) != m_sessions.end();
}
