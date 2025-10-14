#pragma once

#include "InputSource.h"
#include "OutputDestination.h"
#include "AudioCapture.h"
#include "AudioMixer.h"
#include "WavWriter.h"
#include "Mp3Encoder.h"
#include "OpusEncoder.h"
#include "FlacEncoder.h"
#include <memory>
#include <map>
#include <vector>
#include <mutex>
#include <thread>
#include <atomic>

enum class AudioFormat {
    WAV,
    MP3,
    OPUS,
    FLAC
};

// Configuration structures
struct RoutingRule {
    std::wstring sourceId;          // ID of the input source (empty = all sources)
    std::wstring destinationId;     // ID of the output destination
    float volumeMultiplier;         // Volume adjustment for this route (1.0 = 100%)
    bool skipSilence;               // Skip silent audio frames

    RoutingRule() : volumeMultiplier(1.0f), skipSilence(false) {}
};

struct CaptureConfig {
    std::vector<InputSourcePtr> sources;           // Input sources to capture from
    std::vector<OutputDestinationPtr> destinations; // Output destinations to write to
    std::vector<RoutingRule> routingRules;         // Routing matrix configuration
    bool enableMixedOutput;                        // Enable mixing all sources to one output
    std::wstring mixedOutputPath;                  // Path for mixed output (if enabled)
    AudioFormat mixedOutputFormat;                 // Format for mixed output
    UINT32 mixedOutputBitrate;                     // Bitrate for mixed output

    CaptureConfig()
        : enableMixedOutput(false)
        , mixedOutputFormat(AudioFormat::WAV)
        , mixedOutputBitrate(192000) {}
};

class CaptureManager {
public:
    CaptureManager();
    ~CaptureManager();

    // ========================================
    // New API: Flexible routing architecture
    // ========================================

    /**
     * @brief Start a new capture session with flexible routing
     * @param config Configuration with sources, destinations, and routing rules
     * @return Session ID (non-zero on success, 0 on failure)
     */
    UINT32 StartCaptureSession(const CaptureConfig& config);

    /**
     * @brief Stop a specific capture session
     * @param sessionId The session ID returned from StartCaptureSession
     * @return true if session was stopped, false if not found
     */
    bool StopCaptureSession(UINT32 sessionId);

    /**
     * @brief Pause a specific session (stops data flow but keeps resources)
     * @param sessionId The session ID to pause
     */
    void PauseSession(UINT32 sessionId);

    /**
     * @brief Resume a paused session
     * @param sessionId The session ID to resume
     */
    void ResumeSession(UINT32 sessionId);

    /**
     * @brief Add an input source to an existing session
     * @param sessionId The session ID
     * @param source The input source to add
     * @return true if added successfully
     */
    bool AddInputSource(UINT32 sessionId, InputSourcePtr source);

    /**
     * @brief Remove an input source from a session
     * @param sessionId The session ID
     * @param sourceId The source ID to remove
     * @return true if removed successfully
     */
    bool RemoveInputSource(UINT32 sessionId, const std::wstring& sourceId);

    /**
     * @brief Add an output destination to an existing session
     * @param sessionId The session ID
     * @param destination The output destination to add
     * @return true if added successfully
     */
    bool AddOutputDestination(UINT32 sessionId, OutputDestinationPtr destination);

    /**
     * @brief Remove an output destination from a session
     * @param sessionId The session ID
     * @param destinationId The destination ID to remove
     * @return true if removed successfully
     */
    bool RemoveOutputDestination(UINT32 sessionId, const std::wstring& destinationId);

    /**
     * @brief Add a routing rule to a session
     * @param sessionId The session ID
     * @param rule The routing rule to add
     * @return true if added successfully
     */
    bool AddRoutingRule(UINT32 sessionId, const RoutingRule& rule);

    /**
     * @brief Stop all capture sessions
     */
    void StopAll();

    /**
     * @brief Pause all capture sessions
     */
    void PauseAll();

    /**
     * @brief Resume all capture sessions
     */
    void ResumeAll();

    /**
     * @brief Get count of active sessions
     * @return Number of active capture sessions
     */
    size_t GetActiveSessionCount() const;

    /**
     * @brief Check if a specific session is active
     * @param sessionId The session ID to check
     * @return true if session exists and is active
     */
    bool IsSessionActive(UINT32 sessionId) const;

private:
    // Internal session representation (new architecture)
    struct CaptureSessionInternal {
        UINT32 sessionId;
        std::vector<InputSourcePtr> sources;
        std::vector<OutputDestinationPtr> destinations;
        std::vector<RoutingRule> routingRules;
        bool isPaused;
        bool enableMixedOutput;
        std::unique_ptr<AudioMixer> mixer;
        std::unique_ptr<OutputDestination> mixedDestination;
        std::atomic<bool> isValid;  // Fast lock-free check for session validity
        std::wstring mixerDriverSourceId;  // The source that drives GetMixedAudio calls

        CaptureSessionInternal() : sessionId(0), isPaused(false), enableMixedOutput(false), isValid(true) {}
    };

    // Audio data callback from input sources
    void OnAudioData(UINT32 sessionId, const std::wstring& sourceId,
                     const BYTE* data, UINT32 size, const WAVEFORMATEX* format);

    // Route audio data to destinations based on routing rules
    void RouteAudioData(CaptureSessionInternal* session, const std::wstring& sourceId,
                       const BYTE* data, UINT32 size, const WAVEFORMATEX* format);

    // Check if audio is silent (for skip silence feature)
    bool IsSilent(const BYTE* data, UINT32 size, const WAVEFORMATEX* format);

    // Apply volume adjustment to audio data
    void ApplyVolume(BYTE* data, UINT32 size, const WAVEFORMATEX* format, float volume);

    // Generate unique session ID
    UINT32 GenerateSessionId();

    // Member variables
    std::map<UINT32, std::unique_ptr<CaptureSessionInternal>> m_sessions;
    std::atomic<UINT32> m_nextSessionId;
    std::mutex m_mutex;
};
