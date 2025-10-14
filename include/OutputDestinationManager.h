#pragma once

#include "OutputDestination.h"
#include "WavFileDestination.h"
#include "Mp3FileDestination.h"
#include "OpusFileDestination.h"
#include "FlacFileDestination.h"
#include "DeviceOutputDestination.h"
#include <vector>
#include <memory>
#include <string>

/**
 * @brief Manager class for creating and managing output destinations
 *
 * This class provides a factory for creating output destinations and
 * manages multiple active destinations simultaneously. It allows audio
 * to be sent to multiple outputs at once (e.g., save to file while
 * monitoring on speakers).
 *
 * Features:
 * - Factory methods for creating all destination types
 * - Manages multiple simultaneous destinations
 * - Broadcast audio to all active destinations
 * - Automatic error handling and recovery
 * - Resource cleanup on destruction
 *
 * Usage example:
 * @code
 * OutputDestinationManager manager;
 *
 * // Create WAV file destination
 * DestinationConfig wavConfig;
 * wavConfig.outputPath = L"recording.wav";
 * wavConfig.useTimestamp = true;
 * auto wavDest = manager.CreateDestination(DestinationType::FileWAV);
 * if (wavDest->Configure(format, wavConfig)) {
 *     manager.AddDestination(std::move(wavDest));
 * }
 *
 * // Create device destination for monitoring
 * DestinationConfig deviceConfig;
 * deviceConfig.outputPath = deviceId;
 * deviceConfig.volumeMultiplier = 0.8f;
 * auto deviceDest = manager.CreateDestination(DestinationType::AudioDevice);
 * if (deviceDest->Configure(format, deviceConfig)) {
 *     manager.AddDestination(std::move(deviceDest));
 * }
 *
 * // Write audio to all destinations
 * manager.WriteAudioToAll(audioData, audioSize);
 *
 * // Clean up
 * manager.CloseAll();
 * @endcode
 */
class OutputDestinationManager {
public:
    /**
     * @brief Constructor
     */
    OutputDestinationManager();

    /**
     * @brief Destructor - closes all active destinations
     */
    ~OutputDestinationManager();

    /**
     * @brief Create a new output destination of the specified type
     *
     * Creates a new destination instance but does not configure it.
     * The caller must call Configure() on the returned destination
     * before using it.
     *
     * @param type Type of destination to create
     * @return Unique pointer to the destination, or nullptr on error
     */
    OutputDestinationPtr CreateDestination(DestinationType type);

    /**
     * @brief Add a configured destination to the active list
     *
     * Takes ownership of the destination and adds it to the list of
     * active destinations. Audio written via WriteAudioToAll() will
     * be sent to this destination.
     *
     * @param destination Configured destination to add
     * @return true if added successfully, false if destination is invalid
     */
    bool AddDestination(OutputDestinationPtr destination);

    /**
     * @brief Remove a destination by index
     *
     * Closes and removes the destination at the specified index.
     *
     * @param index Index of destination to remove (0-based)
     * @return true if removed successfully, false if index is invalid
     */
    bool RemoveDestination(size_t index);

    /**
     * @brief Remove all destinations of a specific type
     *
     * Closes and removes all destinations matching the specified type.
     *
     * @param type Type of destinations to remove
     * @return Number of destinations removed
     */
    size_t RemoveDestinationsByType(DestinationType type);

    /**
     * @brief Get the number of active destinations
     * @return Number of destinations in the active list
     */
    size_t GetDestinationCount() const {
        return m_destinations.size();
    }

    /**
     * @brief Get a destination by index
     *
     * Returns a raw pointer to the destination for inspection or
     * modification. The manager retains ownership.
     *
     * @param index Index of destination (0-based)
     * @return Pointer to destination, or nullptr if index is invalid
     */
    OutputDestination* GetDestination(size_t index);

    /**
     * @brief Get a const destination by index
     * @param index Index of destination (0-based)
     * @return Const pointer to destination, or nullptr if index is invalid
     */
    const OutputDestination* GetDestination(size_t index) const;

    /**
     * @brief Write audio data to all active destinations
     *
     * Broadcasts the audio data to all destinations in the active list.
     * If a destination fails, it is automatically removed and an error
     * is logged. This ensures one failing destination doesn't stop
     * others from working.
     *
     * @param data Audio data buffer
     * @param size Size of data in bytes
     * @return Number of destinations that successfully wrote the data
     */
    size_t WriteAudioToAll(const BYTE* data, UINT32 size);

    /**
     * @brief Close and remove all destinations
     *
     * Closes all active destinations and clears the list.
     */
    void CloseAll();

    /**
     * @brief Check if any destinations are active
     * @return true if at least one destination is active, false otherwise
     */
    bool HasActiveDestinations() const {
        return !m_destinations.empty();
    }

    /**
     * @brief Get a list of destination names
     *
     * Returns a vector of human-readable names for all active destinations.
     * Useful for displaying in UI.
     *
     * @return Vector of destination names
     */
    std::vector<std::wstring> GetDestinationNames() const;

    /**
     * @brief Get a list of destination types
     *
     * Returns a vector of types for all active destinations.
     *
     * @return Vector of destination types
     */
    std::vector<DestinationType> GetDestinationTypes() const;

    /**
     * @brief Get the last error message from any destination
     *
     * Returns the most recent error from any destination operation.
     *
     * @return Error message, or empty string if no errors
     */
    std::wstring GetLastError() const {
        return m_lastError;
    }

    /**
     * @brief Create a preconfigured destination and add it to the active list
     *
     * Convenience method that combines CreateDestination, Configure, and AddDestination.
     *
     * @param type Type of destination to create
     * @param format Audio format
     * @param config Destination configuration
     * @return true if destination was created and added successfully, false on error
     */
    bool CreateAndAddDestination(DestinationType type, const WAVEFORMATEX* format,
                                  const DestinationConfig& config);

private:
    std::vector<OutputDestinationPtr> m_destinations;  ///< Active destinations
    std::wstring m_lastError;                          ///< Last error message
};
