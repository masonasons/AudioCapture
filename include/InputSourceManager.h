#pragma once

#include "InputSource.h"
#include "ProcessEnumerator.h"
#include "AudioDeviceEnumerator.h"
#include <vector>
#include <memory>
#include <mutex>
#include <string>

/**
 * @file InputSourceManager.h
 * @brief Manager for enumerating and creating audio input sources
 *
 * This class provides a centralized interface for discovering available audio sources
 * (processes, system audio, devices) and creating InputSource instances for them.
 *
 * Design principles:
 * - Factory pattern: Creates InputSource instances based on discovery
 * - Lazy initialization: Sources are created on-demand, not during enumeration
 * - Thread-safe: All public methods are protected by mutex
 * - Resource management: Returns shared_ptr for automatic cleanup
 */

/**
 * @struct AvailableSource
 * @brief Descriptor for an available audio source that can be created
 *
 * This lightweight structure describes a source without creating the full InputSource.
 * Useful for UI enumeration before the user selects which source to capture from.
 */
struct AvailableSource {
    InputSourceMetadata metadata;  // Full metadata including ID, name, type
    bool isAvailable;              // Whether the source is currently available (e.g., process still running)
    std::wstring statusInfo;       // Additional status info (e.g., "Running", "No audio", etc.)
};

class InputSourceManager {
public:
    /**
     * @brief Constructor - initializes enumerators
     *
     * Creates ProcessEnumerator and AudioDeviceEnumerator instances for
     * discovering available sources.
     */
    InputSourceManager();

    /**
     * @brief Destructor - cleans up enumerators
     */
    ~InputSourceManager();

    /**
     * @brief Refresh the list of available sources
     * @param includeProcesses Include process sources in enumeration
     * @param includeSystemAudio Include system audio source
     * @param includeInputDevices Include microphone/line-in devices
     * @param includeOutputDevices Include speaker devices (for loopback)
     * @return true if enumeration succeeded, false otherwise
     *
     * This method queries the system for available audio sources and updates
     * the internal list. It's recommended to call this periodically to detect
     * new processes or devices.
     *
     * Thread-safe: Protected by internal mutex
     */
    bool RefreshAvailableSources(bool includeProcesses = true,
                                 bool includeSystemAudio = true,
                                 bool includeInputDevices = true,
                                 bool includeOutputDevices = true);

    /**
     * @brief Get list of all available sources
     * @return Vector of available source descriptors
     *
     * Returns the list from the last RefreshAvailableSources() call.
     * Call RefreshAvailableSources() first to get current sources.
     *
     * Thread-safe: Returns a copy of the list
     */
    std::vector<AvailableSource> GetAvailableSources() const;

    /**
     * @brief Get available sources filtered by type
     * @param type The source type to filter by
     * @return Vector of available sources of the specified type
     *
     * Thread-safe: Returns a copy of the filtered list
     */
    std::vector<AvailableSource> GetSourcesByType(InputSourceType type) const;

    /**
     * @brief Create an InputSource for a specific process
     * @param processId The Windows process ID
     * @param processName Optional process name (will be looked up if empty)
     * @param windowTitle Optional window title (will be looked up if empty)
     * @return Shared pointer to InputSource, or nullptr if creation failed
     *
     * Creates a ProcessInputSource for the specified process. The process must
     * be running and accessible. If processName is empty, it will be looked up
     * using ProcessEnumerator.
     *
     * Thread-safe: Can be called from any thread
     */
    InputSourcePtr CreateProcessSource(DWORD processId,
                                       const std::wstring& processName = L"",
                                       const std::wstring& windowTitle = L"");

    /**
     * @brief Create a system audio input source
     * @return Shared pointer to InputSource for system audio
     *
     * Creates a SystemAudioInputSource that captures all system audio.
     * This always succeeds (barring system-level issues).
     *
     * Thread-safe: Can be called from any thread
     */
    InputSourcePtr CreateSystemAudioSource();

    /**
     * @brief Create an input device source
     * @param deviceId The Windows MMDevice ID
     * @param friendlyName Optional friendly name (will be looked up if empty)
     * @param isInputDevice true for microphones, false for speakers (loopback)
     * @return Shared pointer to InputSource, or nullptr if creation failed
     *
     * Creates an InputDeviceSource for the specified device. The device must
     * exist and be active. If friendlyName is empty, it will be looked up
     * using AudioDeviceEnumerator.
     *
     * Thread-safe: Can be called from any thread
     */
    InputSourcePtr CreateDeviceSource(const std::wstring& deviceId,
                                      const std::wstring& friendlyName = L"",
                                      bool isInputDevice = true);

    /**
     * @brief Create an InputSource from an AvailableSource descriptor
     * @param source The available source descriptor
     * @return Shared pointer to InputSource, or nullptr if creation failed
     *
     * Convenience method that creates the appropriate InputSource based on
     * the descriptor's type and metadata.
     *
     * Thread-safe: Can be called from any thread
     */
    InputSourcePtr CreateSource(const AvailableSource& source);

    /**
     * @brief Find process information by PID
     * @param processId The process ID to look up
     * @return ProcessInfo structure, or empty structure if not found
     *
     * Utility method to get process details. Useful for getting process names
     * and window titles.
     *
     * Thread-safe: Can be called from any thread
     */
    ProcessInfo FindProcessInfo(DWORD processId);

private:
    // Enumerators for discovering sources
    std::unique_ptr<ProcessEnumerator> m_processEnumerator;
    std::unique_ptr<AudioDeviceEnumerator> m_deviceEnumerator;

    // Cached list of available sources
    std::vector<AvailableSource> m_availableSources;

    // Thread synchronization
    mutable std::mutex m_mutex;

    /**
     * @brief Add process sources to the available list
     */
    void EnumerateProcessSources();

    /**
     * @brief Add system audio source to the available list
     */
    void EnumerateSystemAudioSource();

    /**
     * @brief Add input device sources to the available list
     */
    void EnumerateInputDeviceSources();

    /**
     * @brief Add output device sources to the available list (for loopback)
     */
    void EnumerateOutputDeviceSources();

    /**
     * @brief Look up friendly name for a device
     * @param deviceId The device ID to look up
     * @param isInputDevice Whether to search input or output devices
     * @return Friendly name, or empty string if not found
     */
    std::wstring LookupDeviceName(const std::wstring& deviceId, bool isInputDevice);

    /**
     * @brief Look up process information
     * @param processId The process ID to look up
     * @return ProcessInfo structure with details, or empty if not found
     */
    ProcessInfo LookupProcessInfo(DWORD processId);
};
