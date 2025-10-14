#pragma once

#include "OutputDestination.h"
#include <objbase.h>
#include <mmdeviceapi.h>
#include <audioclient.h>
#include <string>

/**
 * @brief Audio device output destination for monitoring/passthrough
 *
 * This class provides real-time audio output to a physical or virtual
 * audio device. It's commonly used for monitoring captured audio or
 * implementing passthrough functionality.
 *
 * Features:
 * - Real-time audio playback
 * - Configurable volume adjustment
 * - Uses WASAPI for low-latency output
 * - Supports any Windows audio output device
 * - Handles buffer management automatically
 *
 * Use cases:
 * - Monitor captured audio in real-time
 * - Route audio from one device to another
 * - Preview audio before recording
 * - Implement virtual audio cable functionality
 */
class DeviceOutputDestination : public OutputDestination {
public:
    /**
     * @brief Constructor
     */
    DeviceOutputDestination();

    /**
     * @brief Destructor - ensures proper cleanup
     */
    ~DeviceOutputDestination() override;

    /**
     * @brief Get the name of this destination
     * @return Device friendly name if configured, otherwise "Audio Device"
     */
    std::wstring GetName() const override {
        return m_friendlyName.empty() ? L"Audio Device" : m_friendlyName;
    }

    /**
     * @brief Get the type of this destination
     * @return DestinationType::AudioDevice
     */
    DestinationType GetType() const override {
        return DestinationType::AudioDevice;
    }

    /**
     * @brief Configure and open the audio device for output
     *
     * Initializes the specified audio device for playback with the given format.
     * The device ID should be obtained from AudioDeviceEnumerator.
     *
     * @param format Audio format to output
     * @param config Configuration with device ID and volume multiplier
     * @return true if device opened successfully, false on error
     */
    bool Configure(const WAVEFORMATEX* format, const DestinationConfig& config) override;

    /**
     * @brief Close the audio device and stop playback
     *
     * Stops the audio stream and releases the device. After calling Close(),
     * Configure() must be called again to reuse this object.
     */
    void Close() override;

protected:
    /**
     * @brief Internal write implementation called from async writer thread
     *
     * @param data PCM audio data buffer
     * @param size Size of data in bytes
     * @return true if write succeeded, false on error
     */
    bool WriteAudioDataInternal(const BYTE* data, UINT32 size) override;

    /**
     * @brief Check if the audio device is open and playing
     * @return true if device is active, false otherwise
     */
    bool IsOpen() const override;

    /**
     * @brief Get the current volume multiplier
     * @return Volume multiplier (0.0-2.0, 1.0 = 100%)
     */
    float GetVolumeMultiplier() const {
        return m_volumeMultiplier;
    }

    /**
     * @brief Set the volume multiplier
     *
     * Adjusts the output volume in real-time. Values greater than 1.0
     * will amplify the audio (may cause clipping).
     *
     * @param volume Volume multiplier (0.0-2.0)
     */
    void SetVolumeMultiplier(float volume);

private:
    /**
     * @brief Apply volume adjustment to an audio buffer
     *
     * Multiplies each sample by the volume multiplier. Handles both
     * float and integer PCM formats.
     *
     * @param data Audio data buffer (modified in-place)
     * @param size Size of buffer in bytes
     */
    void ApplyVolumeToBuffer(BYTE* data, UINT32 size);

    /**
     * @brief Validate that volume is in acceptable range
     * @param volume Volume to validate
     * @return Clamped volume value (0.0-2.0)
     */
    float ValidateVolume(float volume);

    IMMDeviceEnumerator* m_deviceEnumerator;  ///< Device enumerator
    IMMDevice* m_device;                       ///< Audio device
    IAudioClient* m_audioClient;               ///< Audio client for playback
    IAudioRenderClient* m_renderClient;        ///< Render client for buffer access
    WAVEFORMATEX* m_format;                    ///< Audio format
    UINT32 m_bufferFrameCount;                 ///< Size of device buffer in frames
    float m_volumeMultiplier;                  ///< Volume adjustment (1.0 = 100%)
    std::wstring m_deviceId;                   ///< Device ID string
    std::wstring m_friendlyName;               ///< Device friendly name for identification
    std::wstring m_lastError;                  ///< Last error message
};
