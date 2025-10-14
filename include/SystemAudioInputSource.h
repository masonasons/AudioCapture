#pragma once

#include "InputSource.h"
#include "AudioCapture.h"
#include <memory>
#include <mutex>

/**
 * @file SystemAudioInputSource.h
 * @brief Input source for capturing all system audio
 *
 * This class wraps AudioCapture to provide system-wide audio capture through
 * the InputSource interface. It captures all audio playing on the system using
 * the Windows loopback recording API.
 *
 * Design notes:
 * - Uses composition: wraps AudioCapture internally
 * - Thread-safe: Protected by mutex for state changes
 * - RAII: Automatically stops capture and cleans up in destructor
 * - Equivalent to capturing from "process 0" in the old API
 */

class SystemAudioInputSource : public InputSource {
public:
    /**
     * @brief Construct a SystemAudioInputSource
     *
     * Creates a source that captures all system audio output. This uses the
     * Windows WASAPI loopback recording mode to capture the mixed audio stream
     * from the default output device.
     */
    SystemAudioInputSource();

    /**
     * @brief Destructor - stops capture and releases resources
     *
     * Ensures that capture is stopped and AudioCapture is properly cleaned up.
     */
    virtual ~SystemAudioInputSource();

    // InputSource interface implementation
    virtual InputSourceMetadata GetMetadata() const override;
    virtual InputSourceType GetType() const override { return InputSourceType::SystemAudio; }
    virtual bool StartCapture() override;
    virtual void StopCapture() override;
    virtual bool IsCapturing() const override;
    virtual void SetDataCallback(std::function<void(const BYTE*, UINT32)> callback) override;
    virtual WAVEFORMATEX* GetFormat() const override;
    virtual void SetVolume(float volume) override;
    virtual void Pause() override;
    virtual void Resume() override;
    virtual bool IsPaused() const override;

protected:
    virtual bool Initialize() override;

private:
    // Source identifier (constant for system audio)
    static constexpr const wchar_t* SOURCE_ID = L"system:audio";
    static constexpr const wchar_t* DISPLAY_NAME = L"System Audio (All Sounds)";

    // Audio capture engine (composition pattern)
    std::unique_ptr<AudioCapture> m_audioCapture;

    // Thread synchronization
    mutable std::mutex m_mutex;

    // Initialization state
    std::atomic<bool> m_initialized;
};
