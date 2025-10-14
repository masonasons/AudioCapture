#pragma once

#include <windows.h>
#include <mmreg.h>
#include <vector>
#include <mutex>
#include <map>

// Simple audio mixer that combines multiple audio streams by summing samples
class AudioMixer {
public:
    AudioMixer();
    ~AudioMixer();

    // Initialize with the target audio format (mixer output format)
    bool Initialize(const WAVEFORMATEX* format);

    // Add audio data from a specific source (identified by sourceId)
    // The sourceFormat parameter specifies the format of the incoming data
    // Audio will be resampled to match the mixer's target format if needed
    void AddAudioData(DWORD sourceId, const BYTE* data, UINT32 size, const WAVEFORMATEX* sourceFormat);

    // Get the mixed audio buffer (call this periodically to get mixed output)
    // Returns true if there's data available, false otherwise
    bool GetMixedAudio(std::vector<BYTE>& outBuffer);

    // Remove a specific source from the mixer
    void RemoveSource(DWORD sourceId);

    // Clear all pending audio data
    void Clear();

    // Get the mixer's target format
    const WAVEFORMATEX& GetFormat() const { return m_format; }

private:
    struct AudioBuffer {
        std::vector<BYTE> data;
        UINT32 readPosition;
        WAVEFORMATEX sourceFormat;  // Format of this source (may differ from mixer format)
        std::vector<BYTE> resampleBuffer;  // Pre-allocated buffer for resampling (avoids allocation in callback)
    };

    WAVEFORMATEX m_format;  // Target output format
    bool m_initialized;
    std::mutex m_mutex;
    std::map<DWORD, AudioBuffer> m_buffers;  // Per-source audio buffers

    // Mix audio samples based on format
    void MixSamples(const std::vector<const BYTE*>& sources, BYTE* dest, UINT32 frameCount);

    // Resample audio from source format to target format using linear interpolation
    // Writes into pre-allocated destBuffer to avoid heap allocations
    bool ResampleAudioInPlace(const BYTE* data, UINT32 size, const WAVEFORMATEX* sourceFormat,
                              BYTE* destBuffer, UINT32 destSize);
};
