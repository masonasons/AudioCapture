#pragma once

#include "FileOutputDestination.h"
#include "Mp3Encoder.h"
#include <memory>

/**
 * @brief MP3 file output destination
 *
 * This class wraps the Mp3Encoder class to provide MP3 file output
 * through the OutputDestination interface. MP3 uses lossy compression
 * and is widely compatible across devices and software.
 *
 * Features:
 * - Configurable bitrate (128, 192, 256, 320 kbps)
 * - Uses Windows Media Foundation for encoding
 * - Supports PCM and float input formats
 * - Handles frame buffering internally
 *
 * Recommended bitrates:
 * - 128 kbps: Good quality, smaller files
 * - 192 kbps: High quality (recommended default)
 * - 256 kbps: Very high quality
 * - 320 kbps: Maximum quality
 */
class Mp3FileDestination : public FileOutputDestination {
public:
    /**
     * @brief Constructor
     */
    Mp3FileDestination();

    /**
     * @brief Destructor - ensures proper cleanup
     */
    ~Mp3FileDestination() override;

    /**
     * @brief Get the name of this destination (returns configured file path)
     * @return Configured file path, or "MP3 File" if not configured
     */
    std::wstring GetName() const override {
        return m_filePath.empty() ? L"MP3 File" : m_filePath;
    }

    /**
     * @brief Get the type of this destination
     * @return DestinationType::FileMP3
     */
    DestinationType GetType() const override {
        return DestinationType::FileMP3;
    }

    /**
     * @brief Configure and open the MP3 file for writing
     *
     * Creates the output file with the specified format and configuration.
     * The bitrate can be customized via the config parameter.
     *
     * @param format Audio format (will be converted to MP3)
     * @param config Configuration with output path, timestamp, and bitrate
     * @return true if file opened successfully, false on error
     */
    bool Configure(const WAVEFORMATEX* format, const DestinationConfig& config) override;

    /**
     * @brief Close the MP3 file and finalize encoding
     *
     * Flushes any remaining buffered audio and finalizes the MP3 file.
     * After calling Close(), Configure() must be called again to reuse this object.
     */
    void Close() override;

protected:
    /**
     * @brief Internal write implementation called from async writer thread
     *
     * Audio data is buffered and encoded in MP3 frames (1152 samples per frame).
     * The data must match the format specified in Configure().
     *
     * @param data PCM audio data buffer
     * @param size Size of data in bytes
     * @return true if write succeeded, false on error
     */
    bool WriteAudioDataInternal(const BYTE* data, UINT32 size) override;

    /**
     * @brief Check if the MP3 encoder is open
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
     * MP3 supports specific bitrates. This method ensures the requested
     * bitrate is within acceptable range and adjusts if necessary.
     *
     * @param bitrate Requested bitrate in bps
     * @return Validated bitrate in bps
     */
    UINT32 ValidateBitrate(UINT32 bitrate);

    std::unique_ptr<Mp3Encoder> m_encoder;  ///< MP3 encoder instance
    std::wstring m_filePath;                ///< Current file path
    UINT32 m_bitrate;                       ///< Current bitrate in bps
};
