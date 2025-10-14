#pragma once

#include "FileOutputDestination.h"
#include "WavWriter.h"
#include <memory>

/**
 * @brief WAV file output destination
 *
 * This class wraps the WavWriter class to provide WAV file output
 * through the OutputDestination interface. WAV files store uncompressed
 * PCM audio data with a standard RIFF header.
 *
 * Features:
 * - Supports all PCM formats (16-bit, 32-bit int, 32-bit float)
 * - Supports WAVEFORMATEXTENSIBLE for multi-channel audio
 * - Handles file creation with optional timestamps
 * - Automatic header finalization on close
 */
class WavFileDestination : public FileOutputDestination {
public:
    /**
     * @brief Constructor
     */
    WavFileDestination();

    /**
     * @brief Destructor - ensures proper cleanup
     */
    ~WavFileDestination() override;

    /**
     * @brief Get the name of this destination (returns configured file path)
     * @return Configured file path, or "WAV File" if not configured
     */
    std::wstring GetName() const override {
        return m_filePath.empty() ? L"WAV File" : m_filePath;
    }

    /**
     * @brief Get the type of this destination
     * @return DestinationType::FileWAV
     */
    DestinationType GetType() const override {
        return DestinationType::FileWAV;
    }

    /**
     * @brief Configure and open the WAV file for writing
     *
     * Creates the output file with the specified format and configuration.
     * The file path can include a timestamp if configured.
     *
     * @param format Audio format (supports PCM and float formats)
     * @param config Configuration with output path and timestamp option
     * @return true if file opened successfully, false on error
     */
    bool Configure(const WAVEFORMATEX* format, const DestinationConfig& config) override;

    /**
     * @brief Close the WAV file and finalize the header
     *
     * Updates the WAV header with the final file size and closes the file.
     * After calling Close(), Configure() must be called again to reuse this object.
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
     * @brief Check if the WAV file is open
     * @return true if file is open and ready to write, false otherwise
     */
    bool IsOpen() const override;

private:
    std::unique_ptr<WavWriter> m_writer;  ///< WAV file writer instance
    std::wstring m_filePath;              ///< Current file path
};
