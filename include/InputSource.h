#pragma once

#include <windows.h>
#include <mmreg.h>
#include <string>
#include <functional>
#include <memory>
#include <atomic>

/**
 * @file InputSource.h
 * @brief Abstract base interface for audio input sources
 *
 * This abstraction layer provides a unified interface for capturing audio from different
 * sources: specific processes, system-wide audio, and input devices (microphones).
 *
 * Design principles:
 * - Single responsibility: Each source type handles one kind of audio input
 * - Composition over inheritance: Uses existing AudioCapture internally
 * - RAII: Resources are managed automatically via smart pointers and destructors
 * - Thread-safe: Atomic flags and proper synchronization for concurrent access
 */

/**
 * @enum InputSourceType
 * @brief Types of audio input sources supported by the system
 */
enum class InputSourceType {
    Process,        // Captures audio from a specific process
    SystemAudio,    // Captures all system audio (loopback)
    InputDevice     // Captures from a microphone or line-in device
};

/**
 * @struct InputSourceMetadata
 * @brief Metadata describing an input source
 *
 * This structure contains identifying and display information for an input source,
 * useful for UI presentation and source tracking.
 */
struct InputSourceMetadata {
    std::wstring id;            // Unique identifier (e.g., "process:1234" or "device:guid")
    std::wstring displayName;   // Human-readable name for display
    InputSourceType type;       // Type of source
    std::wstring iconHint;      // Icon/indicator hint (e.g., process name, device type)

    // Additional type-specific metadata
    DWORD processId;            // Valid only for Process type
    std::wstring deviceId;      // Valid only for InputDevice type
};

/**
 * @class InputSource
 * @brief Abstract base class for all audio input sources
 *
 * This class defines the interface that all concrete input source implementations must follow.
 * It provides common functionality and ensures consistent behavior across different source types.
 *
 * Usage pattern:
 * 1. Create a concrete InputSource instance (ProcessInputSource, etc.)
 * 2. Set a data callback to receive audio samples
 * 3. Call StartCapture() to begin capturing
 * 4. Call StopCapture() when done
 * 5. Destructor handles cleanup automatically
 *
 * Thread safety:
 * - StartCapture/StopCapture can be called from any thread
 * - Data callback is invoked from an internal capture thread
 * - GetMetadata() is safe to call at any time
 */
class InputSource {
public:
    virtual ~InputSource() = default;

    /**
     * @brief Get metadata describing this input source
     * @return Metadata structure with source information
     *
     * This method is const and thread-safe, can be called at any time.
     */
    virtual InputSourceMetadata GetMetadata() const = 0;

    /**
     * @brief Get the type of this input source
     * @return The source type enum value
     */
    virtual InputSourceType GetType() const = 0;

    /**
     * @brief Start capturing audio from this source
     * @return true if capture started successfully, false otherwise
     *
     * This method initializes the audio capture pipeline and starts the capture thread.
     * If already capturing, this method returns false.
     *
     * Thread-safe: Can be called from any thread
     */
    virtual bool StartCapture() = 0;

    /**
     * @brief Stop capturing audio from this source
     *
     * This method stops the capture thread and releases audio resources.
     * If not currently capturing, this method does nothing.
     *
     * Thread-safe: Can be called from any thread
     * Blocks until capture thread terminates gracefully
     */
    virtual void StopCapture() = 0;

    /**
     * @brief Check if currently capturing audio
     * @return true if actively capturing, false otherwise
     *
     * Thread-safe: Can be called from any thread
     */
    virtual bool IsCapturing() const = 0;

    /**
     * @brief Set callback function for receiving audio data
     * @param callback Function to call when audio data is available
     *                 Parameters: (const BYTE* data, UINT32 size)
     *
     * The callback is invoked from the capture thread whenever new audio samples
     * are available. The data pointer is valid only for the duration of the callback.
     *
     * Important: The callback should process data quickly to avoid buffer overruns.
     * For heavy processing, copy the data and process it on another thread.
     *
     * Thread-safe: Can be called before or during capture
     */
    virtual void SetDataCallback(std::function<void(const BYTE*, UINT32)> callback) = 0;

    /**
     * @brief Get the audio format for this source
     * @return Pointer to WAVEFORMATEX structure, or nullptr if not initialized
     *
     * The returned pointer is valid for the lifetime of the InputSource.
     * Do not free or modify the returned structure.
     *
     * Note: Format is only available after Initialize() has been called successfully.
     */
    virtual WAVEFORMATEX* GetFormat() const = 0;

    /**
     * @brief Set volume multiplier for this source
     * @param volume Volume multiplier (0.0 = silent, 1.0 = original volume)
     *
     * This applies a software volume adjustment to the captured audio.
     * Values > 1.0 may cause clipping.
     *
     * Thread-safe: Can be called at any time
     */
    virtual void SetVolume(float volume) = 0;

    /**
     * @brief Pause audio capture without releasing resources
     *
     * Temporarily stops capturing audio but keeps the audio pipeline initialized.
     * Resume() can be called to continue capture with minimal latency.
     *
     * Thread-safe: Can be called from any thread
     */
    virtual void Pause() = 0;

    /**
     * @brief Resume audio capture after pausing
     *
     * Resumes capture after a Pause() call. Has no effect if not paused.
     *
     * Thread-safe: Can be called from any thread
     */
    virtual void Resume() = 0;

    /**
     * @brief Check if currently paused
     * @return true if paused, false otherwise
     *
     * Thread-safe: Can be called from any thread
     */
    virtual bool IsPaused() const = 0;

protected:
    /**
     * @brief Initialize the input source
     * @return true if initialization successful, false otherwise
     *
     * This is called internally by concrete implementations to set up
     * the audio capture pipeline. It should be called before StartCapture().
     */
    virtual bool Initialize() = 0;
};

/**
 * @typedef InputSourcePtr
 * @brief Smart pointer type for InputSource objects
 *
 * Using shared_ptr allows InputSource objects to be safely shared and
 * automatically cleaned up when no longer needed.
 */
using InputSourcePtr = std::shared_ptr<InputSource>;
