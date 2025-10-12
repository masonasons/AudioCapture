#include "AudioMixer.h"
#include <algorithm>
#include <cstring>

AudioMixer::AudioMixer() : m_initialized(false) {
    memset(&m_format, 0, sizeof(m_format));
}

AudioMixer::~AudioMixer() {
    Clear();
}

bool AudioMixer::Initialize(const WAVEFORMATEX* format) {
    if (!format) {
        return false;
    }

    std::lock_guard<std::mutex> lock(m_mutex);
    memcpy(&m_format, format, sizeof(WAVEFORMATEX));
    m_initialized = true;
    return true;
}

void AudioMixer::AddAudioData(DWORD sourceId, const BYTE* data, UINT32 size) {
    if (!m_initialized || !data || size == 0) {
        return;
    }

    std::lock_guard<std::mutex> lock(m_mutex);

    // Get or create buffer for this source
    AudioBuffer& buffer = m_buffers[sourceId];

    // Append new data to the buffer
    buffer.data.insert(buffer.data.end(), data, data + size);
}

bool AudioMixer::GetMixedAudio(std::vector<BYTE>& outBuffer) {
    if (!m_initialized) {
        return false;
    }

    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_buffers.empty()) {
        return false;
    }

    // Find the minimum amount of data available across all sources
    UINT32 minDataAvailable = UINT32_MAX;
    for (const auto& pair : m_buffers) {
        UINT32 available = static_cast<UINT32>(pair.second.data.size()) - pair.second.readPosition;
        if (available < minDataAvailable) {
            minDataAvailable = available;
        }
    }

    if (minDataAvailable == 0 || minDataAvailable == UINT32_MAX) {
        return false;
    }

    // Calculate frame count
    UINT32 bytesPerFrame = m_format.nBlockAlign;
    UINT32 frameCount = minDataAvailable / bytesPerFrame;

    if (frameCount == 0) {
        return false;
    }

    UINT32 bytesToMix = frameCount * bytesPerFrame;

    // Prepare output buffer
    outBuffer.resize(bytesToMix);

    // If only one source, just copy the data
    if (m_buffers.size() == 1) {
        const AudioBuffer& buffer = m_buffers.begin()->second;
        memcpy(outBuffer.data(), buffer.data.data() + buffer.readPosition, bytesToMix);
        m_buffers.begin()->second.readPosition += bytesToMix;
    }
    else {
        // Multiple sources - need to mix
        std::vector<const BYTE*> sources;
        for (auto& pair : m_buffers) {
            sources.push_back(pair.second.data.data() + pair.second.readPosition);
        }

        MixSamples(sources, outBuffer.data(), frameCount);

        // Update read positions for all sources
        for (auto& pair : m_buffers) {
            pair.second.readPosition += bytesToMix;
        }
    }

    // Clean up consumed data
    for (auto it = m_buffers.begin(); it != m_buffers.end(); ) {
        AudioBuffer& buffer = it->second;

        // If we've read everything, erase old data
        if (buffer.readPosition >= buffer.data.size()) {
            buffer.data.clear();
            buffer.readPosition = 0;
        }
        // If we've read a significant amount, compact the buffer
        else if (buffer.readPosition > 48000 * m_format.nBlockAlign) { // Keep last second
            buffer.data.erase(buffer.data.begin(), buffer.data.begin() + buffer.readPosition);
            buffer.readPosition = 0;
        }

        ++it;
    }

    return true;
}

void AudioMixer::MixSamples(const std::vector<const BYTE*>& sources, BYTE* dest, UINT32 frameCount) {
    if (sources.empty() || !dest) {
        return;
    }

    UINT32 channels = m_format.nChannels;
    UINT32 bitsPerSample = m_format.wBitsPerSample;

    if (bitsPerSample == 16) {
        // 16-bit PCM mixing
        int16_t* destSamples = reinterpret_cast<int16_t*>(dest);
        UINT32 sampleCount = frameCount * channels;

        for (UINT32 i = 0; i < sampleCount; i++) {
            int32_t sum = 0;

            for (const BYTE* source : sources) {
                const int16_t* sourceSamples = reinterpret_cast<const int16_t*>(source);
                sum += sourceSamples[i];
            }

            // Clamp to 16-bit range
            if (sum > 32767) sum = 32767;
            if (sum < -32768) sum = -32768;

            destSamples[i] = static_cast<int16_t>(sum);
        }
    }
    else if (bitsPerSample == 32) {
        // 32-bit float mixing
        float* destSamples = reinterpret_cast<float*>(dest);
        UINT32 sampleCount = frameCount * channels;

        for (UINT32 i = 0; i < sampleCount; i++) {
            float sum = 0.0f;

            for (const BYTE* source : sources) {
                const float* sourceSamples = reinterpret_cast<const float*>(source);
                sum += sourceSamples[i];
            }

            // Clamp to float range [-1.0, 1.0]
            if (sum > 1.0f) sum = 1.0f;
            if (sum < -1.0f) sum = -1.0f;

            destSamples[i] = sum;
        }
    }
}

void AudioMixer::Clear() {
    std::lock_guard<std::mutex> lock(m_mutex);
    m_buffers.clear();
}
