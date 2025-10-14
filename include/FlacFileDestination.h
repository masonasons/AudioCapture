#pragma once

#include "FileOutputDestination.h"
#include "FlacEncoder.h"
#include <memory>

/**
 * @brief FLAC file output destination
 *
 * This class wraps the FlacEncoder class to provide FLAC file output
 * through the OutputDestination interface. FLAC is a lossless compression
 * format that preserves audio quality while reducing file size.
 *
 * Features:
 * - Configurable compression level (0-8)
 * - Lossless compression (perfect quality)
 * - Uses libFLAC for encoding
 * - Typically 40-60% of WAV file size
 * - Supports up to 8 channels
 *
 * Compression levels:
 * - 0: Fastest compression, larger files
 * - 5: Balanced (recommended default)
 * - 8: Maximum compression, slower encoding
 */
class FlacFileDestination : public FileOutputDestination {
public:
    /**
     * @brief Constructor
     */
    FlacFileDestination();

    /**
     * @brief Destructor - ensures proper cleanup
     */
    ~FlacFileDestination() override;

    /**
     * @brief Get the name of this destination (returns configured file path)
     * @return Configured file path, or "FLAC File" if not configured
     */
    std::wstring GetName() const override {
        return m_filePath.empty() ? L"FLAC File" : m_filePath;
    }

    /**
     * @brief Get the type of this destination
     * @return DestinationType::FileFLAC
     */
    DestinationType GetType() const override {
        return DestinationType::FileFLAC;
    }

    /**
     * @brief Configure and open the FLAC file for writing
     *
     * Creates the output file with the specified format and configuration.
     * The compression level can be customized via the config parameter.
     *
     * @param format Audio format (PCM or float, will be encoded as FLAC)
     * @param config Configuration with output path, timestamp, and compression level
     * @return true if file opened successfully, false on error
     */
    bool Configure(const WAVEFORMATEX* format, const DestinationConfig& config) override;

    /**
     * @brief Close the FLAC file and finalize encoding
     *
     * Flushes any remaining buffered audio and finalizes the FLAC file.
     * After calling Close(), Configure() must be called again to reuse this object.
     */
    void Close() override;

    /**
     * @brief Get the current compression level
     * @return Compression level (0-8)
     */
    UINT32 GetCompressionLevel() const {
        return m_compressionLevel;
    }

protected:
    /**
     * @brief Internal write implementation called from async writer thread
     *
     * Audio data is buffered and encoded in FLAC frames (1024 samples per frame).
     * The data must match the format specified in Configure().
     *
     * @param data PCM audio data buffer
     * @param size Size of data in bytes
     * @return true if write succeeded, false on error
     */
    bool WriteAudioDataInternal(const BYTE* data, UINT32 size) override;

    /**
     * @brief Check if the FLAC encoder is open
     * @return true if encoder is initialized and ready, false otherwise
     */
    bool IsOpen() const override;

private:
    /**
     * @brief Validate and clamp compression level to supported range
     *
     * FLAC supports compression levels 0-8.
     *
     * @param level Requested compression level
     * @return Validated compression level (0-8)
     */
    UINT32 ValidateCompressionLevel(UINT32 level);

    std::unique_ptr<FlacEncoder> m_encoder;  ///< FLAC encoder instance
    std::wstring m_filePath;                 ///< Current file path
    UINT32 m_compressionLevel;               ///< Current compression level
};
