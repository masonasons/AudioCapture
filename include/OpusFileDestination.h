#pragma once

#include "FileOutputDestination.h"
#include "OpusEncoder.h"
#include <memory>

/**
 * @brief Opus file output destination
 *
 * This class wraps the OpusEncoder class to provide Opus file output
 * through the OutputDestination interface. Opus is a modern lossy codec
 * optimized for low latency and high quality at lower bitrates.
 *
 * Features:
 * - Configurable bitrate (64, 96, 128, 192, 256 kbps)
 * - Uses libopus for encoding
 * - OGG container format
 * - Excellent quality at low bitrates
 * - Supports variable bitrate (VBR)
 *
 * Recommended bitrates:
 * - 64 kbps: Good quality for speech
 * - 96 kbps: Good quality for music
 * - 128 kbps: High quality (recommended default)
 * - 192 kbps: Very high quality
 * - 256 kbps: Maximum quality
 */
class OpusFileDestination : public FileOutputDestination {
public:
    /**
     * @brief Constructor
     */
    OpusFileDestination();

    /**
     * @brief Destructor - ensures proper cleanup
     */
    ~OpusFileDestination() override;

    /**
     * @brief Get the name of this destination (returns configured file path)
     * @return Configured file path, or "Opus File" if not configured
     */
    std::wstring GetName() const override {
        return m_filePath.empty() ? L"Opus File" : m_filePath;
    }

    /**
     * @brief Get the type of this destination
     * @return DestinationType::FileOpus
     */
    DestinationType GetType() const override {
        return DestinationType::FileOpus;
    }

    /**
     * @brief Configure and open the Opus file for writing
     *
     * Creates the output file with the specified format and configuration.
     * The bitrate can be customized via the config parameter.
     * Audio is stored in an OGG container.
     *
     * @param format Audio format (will be converted to Opus)
     * @param config Configuration with output path, timestamp, and bitrate
     * @return true if file opened successfully, false on error
     */
    bool Configure(const WAVEFORMATEX* format, const DestinationConfig& config) override;

    /**
     * @brief Close the Opus file and finalize encoding
     *
     * Flushes any remaining buffered audio, writes end-of-stream marker,
     * and finalizes the OGG container. After calling Close(), Configure()
     * must be called again to reuse this object.
     */
    void Close() override;

protected:
    /**
     * @brief Internal write implementation called from async writer thread
     *
     * Audio data is buffered and encoded in Opus frames (960 samples at 48kHz).
     * The data must match the format specified in Configure().
     * Input audio is automatically resampled to 48kHz for Opus encoding.
     *
     * @param data PCM audio data buffer
     * @param size Size of data in bytes
     * @return true if write succeeded, false on error
     */
    bool WriteAudioDataInternal(const BYTE* data, UINT32 size) override;

    /**
     * @brief Check if the Opus encoder is open
     * @return true if encoder is initialized and ready, false otherwise
     */
    bool IsOpen() const override;

    /**
     * @brief Get the current bitrate setting
     * @return Bitrate in bits per second
     */
    UINT32 GetBitrate() const {
        return m_bitrate;
    }

private:
    /**
     * @brief Validate and clamp bitrate to supported values
     *
     * Opus supports bitrates from 6 kbps to 510 kbps, but for audio
     * capture we limit to practical range (64-256 kbps).
     *
     * @param bitrate Requested bitrate in bps
     * @return Validated bitrate in bps
     */
    UINT32 ValidateBitrate(UINT32 bitrate);

    std::unique_ptr<OpusEncoder> m_encoder;  ///< Opus encoder instance
    std::wstring m_filePath;                 ///< Current file path
    UINT32 m_bitrate;                        ///< Current bitrate in bps
};
