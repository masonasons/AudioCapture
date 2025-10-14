#pragma once

#include "InputSource.h"
#include "AudioCapture.h"
#include <memory>
#include <mutex>

/**
 * @file InputDeviceSource.h
 * @brief Input source for capturing audio from input devices (microphones, line-in)
 *
 * This class wraps AudioCapture to provide audio capture from physical input devices
 * through the InputSource interface. It uses the Windows WASAPI to capture from
 * capture endpoints like microphones and line-in devices.
 *
 * Design notes:
 * - Uses composition: wraps AudioCapture internally
 * - Thread-safe: Protected by mutex for state changes
 * - RAII: Automatically stops capture and cleans up in destructor
 * - Supports both input devices (microphones) and output devices (speakers in loopback mode)
 */

class InputDeviceSource : public InputSource {
public:
    /**
     * @brief Construct an InputDeviceSource for a specific device
     * @param deviceId The Windows MMDevice ID string
     * @param friendlyName The user-friendly name of the device
     * @param isInputDevice true for microphones/line-in, false for speakers (loopback)
     *
     * The source will capture audio from the specified device. For input devices
     * (isInputDevice=true), it captures directly from the microphone. For output
     * devices (isInputDevice=false), it uses loopback mode to capture what's playing
     * on that specific output device.
     */
    InputDeviceSource(const std::wstring& deviceId,
                     const std::wstring& friendlyName,
                     bool isInputDevice);

    /**
     * @brief Destructor - stops capture and releases resources
     *
     * Ensures that capture is stopped and AudioCapture is properly cleaned up.
     */
    virtual ~InputDeviceSource();

    // InputSource interface implementation
    virtual InputSourceMetadata GetMetadata() const override;
    virtual InputSourceType GetType() const override { return InputSourceType::InputDevice; }
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
    // Device information
    std::wstring m_deviceId;
    std::wstring m_friendlyName;
    bool m_isInputDevice;
    std::wstring m_sourceId;

    // Audio capture engine (composition pattern)
    std::unique_ptr<AudioCapture> m_audioCapture;

    // Thread synchronization
    mutable std::mutex m_mutex;

    // Initialization state
    std::atomic<bool> m_initialized;

    /**
     * @brief Generate a unique ID for this source
     * @return String in format "device:<device_id_hash>"
     */
    std::wstring GenerateSourceId() const;

    /**
     * @brief Generate display name for this source
     * @return Display-friendly name with device type indicator
     */
    std::wstring GenerateDisplayName() const;
};
