#pragma once

#include "OutputDestination.h"
#include <string>
#include <chrono>
#include <sstream>
#include <iomanip>

/**
 * @brief Base class for file-based output destinations
 *
 * This class provides common functionality for all file-based outputs:
 * - File path generation with optional timestamps
 * - Common configuration handling
 * - Error message storage
 *
 * Derived classes (WAV, MP3, Opus, FLAC) implement the specific encoding logic.
 */
class FileOutputDestination : public OutputDestination {
public:
    FileOutputDestination() = default;
    virtual ~FileOutputDestination() = default;

    /**
     * @brief Get the last error message
     * @return Error message string, or empty if no error
     */
    std::wstring GetLastError() const override {
        return m_lastError;
    }

protected:
    /**
     * @brief Generate output file path with optional timestamp
     *
     * If useTimestamp is true, inserts a timestamp before the file extension.
     * Example: "recording.wav" becomes "recording_20231215_143022.wav"
     *
     * @param basePath Base file path from configuration
     * @param useTimestamp Whether to add timestamp
     * @return Complete file path
     */
    std::wstring GenerateFilePath(const std::wstring& basePath, bool useTimestamp) {
        if (!useTimestamp) {
            return basePath;
        }

        // Find the last dot (file extension)
        size_t dotPos = basePath.find_last_of(L'.');
        size_t slashPos = basePath.find_last_of(L"\\/");

        // Make sure the dot is after any path separators (not a directory with a dot)
        if (dotPos != std::wstring::npos &&
            (slashPos == std::wstring::npos || dotPos > slashPos)) {

            // Get current time
            auto now = std::chrono::system_clock::now();
            auto time = std::chrono::system_clock::to_time_t(now);

            // Format timestamp as _YYYYMMDD_HHMMSS
            std::wstringstream timestamp;
            struct tm timeinfo;
            localtime_s(&timeinfo, &time);

            timestamp << L"_"
                << std::setfill(L'0')
                << std::setw(4) << (timeinfo.tm_year + 1900)
                << std::setw(2) << (timeinfo.tm_mon + 1)
                << std::setw(2) << timeinfo.tm_mday
                << L"_"
                << std::setw(2) << timeinfo.tm_hour
                << std::setw(2) << timeinfo.tm_min
                << std::setw(2) << timeinfo.tm_sec;

            // Insert timestamp before extension
            std::wstring path = basePath.substr(0, dotPos);
            path += timestamp.str();
            path += basePath.substr(dotPos);

            return path;
        }

        // No extension found or dot is part of directory - just append timestamp
        auto now = std::chrono::system_clock::now();
        auto time = std::chrono::system_clock::to_time_t(now);

        std::wstringstream timestamp;
        struct tm timeinfo;
        localtime_s(&timeinfo, &time);

        timestamp << L"_"
            << std::setfill(L'0')
            << std::setw(4) << (timeinfo.tm_year + 1900)
            << std::setw(2) << (timeinfo.tm_mon + 1)
            << std::setw(2) << timeinfo.tm_mday
            << L"_"
            << std::setw(2) << timeinfo.tm_hour
            << std::setw(2) << timeinfo.tm_min
            << std::setw(2) << timeinfo.tm_sec;

        return basePath + timestamp.str();
    }

    /**
     * @brief Set the last error message
     * @param error Error message to store
     */
    void SetError(const std::wstring& error) {
        m_lastError = error;
    }

    /**
     * @brief Clear the last error message
     */
    void ClearError() {
        m_lastError.clear();
    }

    /**
     * @brief Validate that a file path is not empty
     * @param path File path to validate
     * @return true if valid, false if empty
     */
    bool ValidateFilePath(const std::wstring& path) {
        if (path.empty()) {
            SetError(L"Output path cannot be empty");
            return false;
        }
        return true;
    }

    /**
     * @brief Validate that an audio format is supported
     *
     * Checks for common issues like null pointer, unsupported format tags,
     * invalid channel counts, or invalid sample rates.
     *
     * @param format WAVEFORMATEX structure to validate
     * @return true if valid, false if invalid
     */
    bool ValidateFormat(const WAVEFORMATEX* format) {
        if (!format) {
            SetError(L"Audio format is null");
            return false;
        }

        if (format->nChannels == 0 || format->nChannels > 8) {
            SetError(L"Invalid channel count (must be 1-8)");
            return false;
        }

        if (format->nSamplesPerSec == 0 || format->nSamplesPerSec > 192000) {
            SetError(L"Invalid sample rate (must be 1-192000 Hz)");
            return false;
        }

        if (format->wBitsPerSample == 0 || format->wBitsPerSample > 32) {
            SetError(L"Invalid bits per sample (must be 1-32)");
            return false;
        }

        return true;
    }

private:
    std::wstring m_lastError;  ///< Last error message
};
