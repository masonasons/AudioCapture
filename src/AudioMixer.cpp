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

void AudioMixer::AddAudioData(DWORD sourceId, const BYTE* data, UINT32 size, const WAVEFORMATEX* sourceFormat) {
    if (!m_initialized || !data || size == 0 || !sourceFormat) {
        return;
    }

    std::lock_guard<std::mutex> lock(m_mutex);

    // Get or create buffer for this source
    AudioBuffer& buffer = m_buffers[sourceId];

    // Check if resampling is needed
    if (sourceFormat->nSamplesPerSec != m_format.nSamplesPerSec ||
        sourceFormat->nChannels != m_format.nChannels ||
        sourceFormat->wBitsPerSample != m_format.wBitsPerSample) {
        // Resample the audio to match target format
        std::vector<BYTE> resampledData = ResampleAudio(data, size, sourceFormat);
        buffer.data.insert(buffer.data.end(), resampledData.begin(), resampledData.end());
    } else {
        // No resampling needed, append directly
        buffer.data.insert(buffer.data.end(), data, data + size);
    }
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
        else if (buffer.readPosition > static_cast<UINT32>(48000 * m_format.nBlockAlign)) { // Keep last second
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

std::vector<BYTE> AudioMixer::ResampleAudio(const BYTE* data, UINT32 size, const WAVEFORMATEX* sourceFormat) {
    // Calculate number of frames in source data
    UINT32 sourceBytesPerFrame = sourceFormat->nBlockAlign;
    UINT32 sourceFrameCount = size / sourceBytesPerFrame;

    // Calculate target frame count based on sample rate ratio
    double ratio = (double)m_format.nSamplesPerSec / (double)sourceFormat->nSamplesPerSec;
    UINT32 targetFrameCount = (UINT32)(sourceFrameCount * ratio);

    // Calculate target buffer size
    UINT32 targetBytesPerFrame = m_format.nBlockAlign;
    UINT32 targetSize = targetFrameCount * targetBytesPerFrame;

    std::vector<BYTE> resampledData(targetSize);

    // Only support 32-bit float for now (most common for WASAPI)
    if (sourceFormat->wBitsPerSample == 32 && m_format.wBitsPerSample == 32) {
        const float* sourceSamples = reinterpret_cast<const float*>(data);
        float* targetSamples = reinterpret_cast<float*>(resampledData.data());

        UINT32 sourceChannels = sourceFormat->nChannels;
        UINT32 targetChannels = m_format.nChannels;

        // Linear interpolation resampling
        for (UINT32 targetFrame = 0; targetFrame < targetFrameCount; targetFrame++) {
            // Calculate source position (floating point)
            double sourcePos = (double)targetFrame / ratio;
            UINT32 sourceFrameLow = (UINT32)sourcePos;
            UINT32 sourceFrameHigh = std::min(sourceFrameLow + 1, sourceFrameCount - 1);
            float frac = (float)(sourcePos - sourceFrameLow);

            // Interpolate each channel
            for (UINT32 ch = 0; ch < targetChannels; ch++) {
                // Handle channel count mismatch
                UINT32 sourceCh = (ch < sourceChannels) ? ch : (sourceChannels - 1);

                float sampleLow = sourceSamples[sourceFrameLow * sourceChannels + sourceCh];
                float sampleHigh = sourceSamples[sourceFrameHigh * sourceChannels + sourceCh];
                float interpolated = sampleLow + (sampleHigh - sampleLow) * frac;

                targetSamples[targetFrame * targetChannels + ch] = interpolated;
            }
        }
    }
    else if (sourceFormat->wBitsPerSample == 16 && m_format.wBitsPerSample == 16) {
        // 16-bit PCM resampling
        const int16_t* sourceSamples = reinterpret_cast<const int16_t*>(data);
        int16_t* targetSamples = reinterpret_cast<int16_t*>(resampledData.data());

        UINT32 sourceChannels = sourceFormat->nChannels;
        UINT32 targetChannels = m_format.nChannels;

        for (UINT32 targetFrame = 0; targetFrame < targetFrameCount; targetFrame++) {
            double sourcePos = (double)targetFrame / ratio;
            UINT32 sourceFrameLow = (UINT32)sourcePos;
            UINT32 sourceFrameHigh = std::min(sourceFrameLow + 1, sourceFrameCount - 1);
            float frac = (float)(sourcePos - sourceFrameLow);

            for (UINT32 ch = 0; ch < targetChannels; ch++) {
                UINT32 sourceCh = (ch < sourceChannels) ? ch : (sourceChannels - 1);

                int16_t sampleLow = sourceSamples[sourceFrameLow * sourceChannels + sourceCh];
                int16_t sampleHigh = sourceSamples[sourceFrameHigh * sourceChannels + sourceCh];
                float interpolated = sampleLow + (sampleHigh - sampleLow) * frac;

                targetSamples[targetFrame * targetChannels + ch] = (int16_t)interpolated;
            }
        }
    }
    else {
        // Unsupported format combination - just copy without resampling (fallback)
        resampledData.assign(data, data + size);
    }

    return resampledData;
}
