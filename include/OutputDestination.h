#pragma once

#include <windows.h>
#include <mmreg.h>
#include <combaseapi.h>
#include <string>
#include <memory>
#include <vector>
#include <queue>
#include <mutex>
#include <condition_variable>
#include <thread>
#include <atomic>

/**
 * @brief Output destination types supported by the AudioCapture system
 */
enum class DestinationType {
    FileWAV,      ///< WAV file output (uncompressed PCM)
    FileMP3,      ///< MP3 file output (lossy compression)
    FileOpus,     ///< Opus file output (lossy compression, low latency)
    FileFLAC,     ///< FLAC file output (lossless compression)
    AudioDevice   ///< Audio device output (monitoring/passthrough)
};

/**
 * @brief Configuration structure for output destinations
 *
 * This structure holds format-specific settings for different output types.
 * Not all fields are used by all destination types.
 */
struct DestinationConfig {
    // Common settings
    std::wstring outputPath;           ///< File path or device ID
    bool useTimestamp;                 ///< Append timestamp to filename (file destinations only)
    std::wstring friendlyName;         ///< Friendly name for identification (device destinations)

    // Encoder-specific settings (used by MP3, Opus, FLAC)
    UINT32 bitrate;                    ///< Bitrate in bps (MP3: 128k-320k, Opus: 64k-256k)
    UINT32 compressionLevel;           ///< Compression level 0-8 (FLAC only)

    // Device-specific settings
    float volumeMultiplier;            ///< Volume adjustment (0.0-2.0, 1.0 = 100%)

    /**
     * @brief Default constructor with reasonable defaults
     */
    DestinationConfig()
        : useTimestamp(false)
        , bitrate(192000)              // 192 kbps default
        , compressionLevel(5)          // Medium compression for FLAC
        , volumeMultiplier(1.0f)       // 100% volume
    {}
};

/**
 * @brief Abstract base class for audio output destinations
 *
 * This class defines the interface that all output destinations must implement.
 * Output destinations receive audio data and can be files (WAV, MP3, Opus, FLAC)
 * or audio devices (for monitoring/passthrough).
 *
 * Key features:
 * - Supports multiple simultaneous destinations
 * - Each destination handles its own configuration
 * - Proper resource cleanup via RAII
 * - Error handling for I/O and device failures
 * - Asynchronous write queue to prevent blocking audio callbacks
 */
class OutputDestination {
public:
    OutputDestination();
    virtual ~OutputDestination();

    /**
     * @brief Get the human-readable name of this destination
     * @return Name string (e.g., "WAV File", "MP3 Encoder", "Audio Device")
     */
    virtual std::wstring GetName() const = 0;

    /**
     * @brief Get the type of this destination
     * @return DestinationType enum value
     */
    virtual DestinationType GetType() const = 0;

    /**
     * @brief Configure and open the destination for writing
     *
     * This method must be called before WriteAudioData(). It initializes
     * the destination with the specified audio format and configuration.
     *
     * @param format Pointer to WAVEFORMATEX structure describing audio format
     * @param config Configuration structure with destination-specific settings
     * @return true if configuration succeeded, false on error
     */
    virtual bool Configure(const WAVEFORMATEX* format, const DestinationConfig& config) = 0;

    /**
     * @brief Write audio data to the destination
     *
     * IMPORTANT: This method is NON-BLOCKING. It queues the data for async write
     * and returns immediately. The actual disk I/O happens on a background thread.
     * This ensures audio callbacks are never blocked by file operations.
     *
     * @param data Pointer to audio data buffer (PCM format matching the configured format)
     * @param size Size of the data buffer in bytes
     * @return true if data was queued successfully, false on error
     */
    bool WriteAudioData(const BYTE* data, UINT32 size);

    /**
     * @brief Close the destination and finalize output
     *
     * This method should flush any buffered data, update file headers if needed,
     * and release all resources. After calling Close(), the destination cannot
     * be used again without reconfiguring.
     */
    virtual void Close() = 0;

    /**
     * @brief Check if the destination is currently open and ready to receive data
     * @return true if open, false if closed or not configured
     */
    virtual bool IsOpen() const = 0;

    /**
     * @brief Get the last error message (if any)
     * @return Error message string, or empty string if no error
     */
    virtual std::wstring GetLastError() const { return L""; }

protected:
    /**
     * @brief Subclasses override this to perform the actual write operation
     *
     * This method is called from the background writer thread, NOT from the
     * audio callback thread. It's safe to perform blocking I/O operations here.
     *
     * @param data Pointer to audio data buffer
     * @param size Size of the data buffer in bytes
     * @return true if write succeeded, false on error
     */
    virtual bool WriteAudioDataInternal(const BYTE* data, UINT32 size) = 0;

    /**
     * @brief Start the async write thread (called by Configure)
     */
    void StartAsyncWriter();

    /**
     * @brief Stop the async write thread (called by Close)
     */
    void StopAsyncWriter();

private:
    // Async write queue infrastructure
    struct AudioChunk {
        std::vector<BYTE> data;
    };

    std::queue<AudioChunk> m_writeQueue;
    std::mutex m_queueMutex;
    std::condition_variable m_queueCV;
    std::thread m_writerThread;
    std::atomic<bool> m_writerRunning;
    std::atomic<bool> m_isOpen;

    // Writer thread function
    void WriterThreadFunc();

protected:
    /**
     * @brief Helper method to get the size of a WAVEFORMATEX structure
     *
     * Handles both WAVEFORMATEX and WAVEFORMATEXTENSIBLE structures correctly.
     *
     * @param format Pointer to WAVEFORMATEX structure
     * @return Size in bytes including any extension data
     */
    static UINT32 GetFormatSize(const WAVEFORMATEX* format) {
        if (!format) {
            return 0;
        }

        if (format->wFormatTag == WAVE_FORMAT_EXTENSIBLE && format->cbSize >= 22) {
            return sizeof(WAVEFORMATEX) + format->cbSize;
        }

        return sizeof(WAVEFORMATEX);
    }

    /**
     * @brief Helper method to copy a WAVEFORMATEX structure
     *
     * Allocates memory and creates a copy of the format structure.
     * Caller is responsible for freeing with CoTaskMemFree().
     *
     * @param format Source format structure
     * @return Pointer to copied format, or nullptr on failure
     */
    static WAVEFORMATEX* CopyFormat(const WAVEFORMATEX* format) {
        if (!format) {
            return nullptr;
        }

        UINT32 formatSize = GetFormatSize(format);
        WAVEFORMATEX* copy = reinterpret_cast<WAVEFORMATEX*>(CoTaskMemAlloc(formatSize));

        if (copy) {
            memcpy(copy, format, formatSize);
        }

        return copy;
    }
};

/**
 * @brief Smart pointer type for OutputDestination objects
 */
using OutputDestinationPtr = std::shared_ptr<OutputDestination>;
