#pragma once

#include "InputSource.h"
#include "AudioCapture.h"
#include <memory>
#include <mutex>

/**
 * @file ProcessInputSource.h
 * @brief Input source for capturing audio from a specific process
 *
 * This class wraps AudioCapture to provide process-specific audio capture through
 * the InputSource interface. It uses the Windows process loopback API to isolate
 * audio from a single process.
 *
 * Design notes:
 * - Uses composition: wraps AudioCapture internally
 * - Thread-safe: Protected by mutex for state changes
 * - RAII: Automatically stops capture and cleans up in destructor
 */

class ProcessInputSource : public InputSource {
public:
    /**
     * @brief Construct a ProcessInputSource for a specific process
     * @param processId The Windows process ID to capture audio from
     * @param processName The name of the process (for display purposes)
     * @param windowTitle Optional window title for additional context
     *
     * The source will attempt to capture audio from the specified process using
     * Windows' process loopback API (available on Windows 10 Build 20348+).
     * If the API is not available, initialization will fail.
     */
    ProcessInputSource(DWORD processId,
                      const std::wstring& processName,
                      const std::wstring& windowTitle = L"");

    /**
     * @brief Destructor - stops capture and releases resources
     *
     * Ensures that capture is stopped and AudioCapture is properly cleaned up.
     */
    virtual ~ProcessInputSource();

    // InputSource interface implementation
    virtual InputSourceMetadata GetMetadata() const override;
    virtual InputSourceType GetType() const override { return InputSourceType::Process; }
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
    // Process information
    DWORD m_processId;
    std::wstring m_processName;
    std::wstring m_windowTitle;
    std::wstring m_sourceId;

    // Audio capture engine (composition pattern)
    std::unique_ptr<AudioCapture> m_audioCapture;

    // Thread synchronization
    mutable std::mutex m_mutex;

    // Initialization state
    std::atomic<bool> m_initialized;

    /**
     * @brief Generate a unique ID for this source
     * @return String in format "process:<pid>"
     */
    std::wstring GenerateSourceId() const;

    /**
     * @brief Generate display name for this source
     * @return Display-friendly name combining process name and window title
     */
    std::wstring GenerateDisplayName() const;
};
