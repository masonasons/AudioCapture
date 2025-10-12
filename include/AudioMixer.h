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

    // Initialize with the audio format (all streams must use the same format)
    bool Initialize(const WAVEFORMATEX* format);

    // Add audio data from a specific source (identified by sourceId)
    void AddAudioData(DWORD sourceId, const BYTE* data, UINT32 size);

    // Get the mixed audio buffer (call this periodically to get mixed output)
    // Returns true if there's data available, false otherwise
    bool GetMixedAudio(std::vector<BYTE>& outBuffer);

    // Clear all pending audio data
    void Clear();

private:
    struct AudioBuffer {
        std::vector<BYTE> data;
        UINT32 readPosition;
    };

    WAVEFORMATEX m_format;
    bool m_initialized;
    std::mutex m_mutex;
    std::map<DWORD, AudioBuffer> m_buffers;  // Per-source audio buffers

    // Mix audio samples based on format
    void MixSamples(const std::vector<const BYTE*>& sources, BYTE* dest, UINT32 frameCount);
};
