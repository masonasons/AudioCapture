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

    // Validate format parameters to prevent crashes
    if (sourceFormat->nSamplesPerSec == 0 || sourceFormat->nChannels == 0 ||
        sourceFormat->wBitsPerSample == 0 || sourceFormat->nBlockAlign == 0) {
        return;  // Invalid format - skip this data
    }

    // CRITICAL OPTIMIZATION: Check if resampling is needed WITHOUT holding mutex
    // Resampling is expensive (loops over thousands of samples) and can block audio callbacks
    bool needsResampling = (sourceFormat->nSamplesPerSec != m_format.nSamplesPerSec ||
                           sourceFormat->nChannels != m_format.nChannels ||
                           sourceFormat->wBitsPerSample != m_format.wBitsPerSample);

    // Acquire mutex to get buffer reference (brief lock)
    std::lock_guard<std::mutex> lock(m_mutex);

    // Get or create buffer for this source
    AudioBuffer& buffer = m_buffers[sourceId];

    // Store source format on first call for this source (for format change detection)
    bool isNewSource = (buffer.sourceFormat.nSamplesPerSec == 0);
    if (isNewSource) {
        memcpy(&buffer.sourceFormat, sourceFormat, sizeof(WAVEFORMATEX));
    }

    // CRITICAL: Pre-reserve capacity to avoid reallocation during insert()
    // Reallocation copies ALL existing data, blocking the mutex and causing choppy audio!
    // Reserve in 1-second chunks (at 48kHz stereo 32-bit = 384KB per second)
    const UINT32 RESERVE_CHUNK_SIZE = m_format.nSamplesPerSec * m_format.nBlockAlign;
    if (buffer.data.capacity() < buffer.data.size() + size + RESERVE_CHUNK_SIZE) {
        buffer.data.reserve(buffer.data.size() + RESERVE_CHUNK_SIZE);
    }

    // Perform resampling if needed using PRE-ALLOCATED buffer
    // This prevents expensive heap allocations on every audio callback
    if (needsResampling) {
        // Calculate required buffer size
        UINT32 sourceBytesPerFrame = sourceFormat->nBlockAlign;
        UINT32 sourceFrameCount = size / sourceBytesPerFrame;
        double ratio = (double)m_format.nSamplesPerSec / (double)sourceFormat->nSamplesPerSec;
        UINT32 targetFrameCount = (UINT32)(sourceFrameCount * ratio);
        UINT32 targetBytesPerFrame = m_format.nBlockAlign;
        UINT32 targetSize = targetFrameCount * targetBytesPerFrame;

        // Resize ONCE if needed (amortized cost, not per-callback)
        if (buffer.resampleBuffer.size() < targetSize) {
            buffer.resampleBuffer.resize(targetSize);
        }

        // Resample into the pre-allocated buffer
        if (!ResampleAudioInPlace(data, size, sourceFormat, buffer.resampleBuffer.data(), targetSize)) {
            return;  // Resampling failed - skip this data
        }

        // Append resampled data - now guaranteed not to reallocate!
        buffer.data.insert(buffer.data.end(), buffer.resampleBuffer.data(), buffer.resampleBuffer.data() + targetSize);
    }
    else {
        // No resampling needed - append original data directly - no reallocation!
        buffer.data.insert(buffer.data.end(), data, data + size);
    }
}

bool AudioMixer::GetMixedAudio(std::vector<BYTE>& outBuffer) {
    if (!m_initialized) {
        return false;
    }

    // CRITICAL OPTIMIZATION: Minimize mutex hold time
    // We'll acquire mutex ONLY to snapshot buffer state, then release it
    // before doing expensive operations (mixing, memory allocation, etc.)

    // Snapshot of buffer data for mixing
    struct BufferSnapshot {
        DWORD sourceId;
        const BYTE* dataPtr;
        UINT32 availableBytes;
    };
    std::vector<BufferSnapshot> snapshots;
    UINT32 bytesToMix = 0;
    UINT32 frameCount = 0;
    UINT32 bytesPerFrame = 0;

    // Phase 1: BRIEF mutex lock to snapshot buffer state
    {
        std::lock_guard<std::mutex> lock(m_mutex);

        if (m_buffers.empty()) {
            return false;
        }

        // Validate format before proceeding
        if (m_format.nBlockAlign == 0) {
            return false;
        }

        bytesPerFrame = m_format.nBlockAlign;

        // STRICT SYNC with minimum buffer requirement
        // Wait for ALL sources to have at least some minimum amount of data
        // This ensures synchronization while preventing buffer underruns
        UINT32 minDataAvailable = UINT32_MAX;

        // Find the minimum available data across all sources
        // We'll mix whatever is available from ALL sources to maintain sync
        for (const auto& pair : m_buffers) {
            UINT32 available = 0;

            // Calculate available data for this source
            if (pair.second.data.size() > pair.second.readPosition) {
                available = static_cast<UINT32>(pair.second.data.size()) - pair.second.readPosition;
            }

            // Track minimum across all sources
            if (available < minDataAvailable) {
                minDataAvailable = available;
            }
        }

        // If ANY source has no data, return false - wait for sync
        if (minDataAvailable == 0) {
            return false;  // At least one source has no data - wait
        }

        // Calculate frame count
        frameCount = minDataAvailable / bytesPerFrame;
        if (frameCount == 0) {
            return false;
        }

        bytesToMix = frameCount * bytesPerFrame;

        // Snapshot buffer pointers for sources that have enough data
        // Sources without enough data will be padded with silence during mixing
        snapshots.reserve(m_buffers.size());
        for (auto& pair : m_buffers) {
            UINT32 available = 0;
            if (pair.second.data.size() > pair.second.readPosition) {
                available = static_cast<UINT32>(pair.second.data.size()) - pair.second.readPosition;
            }

            if (available >= bytesToMix) {
                // This source has enough data
                BufferSnapshot snapshot;
                snapshot.sourceId = pair.first;
                snapshot.dataPtr = pair.second.data.data() + pair.second.readPosition;
                snapshot.availableBytes = available;
                snapshots.push_back(snapshot);
            } else {
                // This source doesn't have enough data yet - will be padded with silence
                BufferSnapshot snapshot;
                snapshot.sourceId = pair.first;
                snapshot.dataPtr = nullptr;  // nullptr indicates silence padding needed
                snapshot.availableBytes = 0;
                snapshots.push_back(snapshot);
            }
        }

        if (snapshots.empty()) {
            return false;
        }
    }
    // CRITICAL: Mutex is now released - audio callbacks can proceed!

    // Phase 2: Perform expensive operations WITHOUT holding mutex

    // OPTIMIZATION: Only resize if buffer is too small (reuse existing capacity)
    // This avoids repeated allocations on every GetMixedAudio() call
    if (outBuffer.capacity() < bytesToMix) {
        outBuffer.reserve(bytesToMix + (m_format.nSamplesPerSec * m_format.nBlockAlign));  // Reserve extra
    }
    outBuffer.resize(bytesToMix);

    // Mix the audio data
    if (snapshots.size() == 1) {
        // Single source
        if (snapshots[0].dataPtr) {
            // Has data - just copy
            memcpy(outBuffer.data(), snapshots[0].dataPtr, bytesToMix);
        } else {
            // No data - fill with silence
            memset(outBuffer.data(), 0, bytesToMix);
        }
    }
    else {
        // Multiple sources - mix them (MixSamples handles nullptr for silence)
        std::vector<const BYTE*> sourcePtrs;
        sourcePtrs.reserve(snapshots.size());
        for (const auto& snapshot : snapshots) {
            sourcePtrs.push_back(snapshot.dataPtr);  // nullptr means silence
        }
        MixSamples(sourcePtrs, outBuffer.data(), frameCount);
    }

    // Phase 3: BRIEF mutex lock to update read positions and cleanup
    {
        std::lock_guard<std::mutex> lock(m_mutex);

        // Update read positions ONLY for sources that had data (dataPtr != nullptr)
        for (const auto& snapshot : snapshots) {
            if (snapshot.dataPtr) {  // Only update if we actually read data
                auto it = m_buffers.find(snapshot.sourceId);
                if (it != m_buffers.end()) {
                    it->second.readPosition += bytesToMix;
                }
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
            else if (buffer.readPosition > static_cast<UINT32>(48000 * m_format.nBlockAlign)) {
                buffer.data.erase(buffer.data.begin(), buffer.data.begin() + buffer.readPosition);
                buffer.readPosition = 0;
            }

            ++it;
        }
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
                if (source) {  // Skip nullptr sources (silence padding)
                    const int16_t* sourceSamples = reinterpret_cast<const int16_t*>(source);
                    sum += sourceSamples[i];
                }
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
                if (source) {  // Skip nullptr sources (silence padding)
                    const float* sourceSamples = reinterpret_cast<const float*>(source);
                    sum += sourceSamples[i];
                }
            }

            // Clamp to float range [-1.0, 1.0]
            if (sum > 1.0f) sum = 1.0f;
            if (sum < -1.0f) sum = -1.0f;

            destSamples[i] = sum;
        }
    }
}

void AudioMixer::RemoveSource(DWORD sourceId) {
    std::lock_guard<std::mutex> lock(m_mutex);
    m_buffers.erase(sourceId);
}

void AudioMixer::Clear() {
    std::lock_guard<std::mutex> lock(m_mutex);
    m_buffers.clear();
}

bool AudioMixer::ResampleAudioInPlace(const BYTE* data, UINT32 size, const WAVEFORMATEX* sourceFormat,
                                      BYTE* destBuffer, UINT32 destSize) {
    if (!data || !destBuffer || !sourceFormat) {
        return false;
    }

    // Calculate number of frames in source data
    UINT32 sourceBytesPerFrame = sourceFormat->nBlockAlign;
    UINT32 sourceFrameCount = size / sourceBytesPerFrame;

    // Calculate target frame count based on sample rate ratio
    double ratio = (double)m_format.nSamplesPerSec / (double)sourceFormat->nSamplesPerSec;
    UINT32 targetFrameCount = (UINT32)(sourceFrameCount * ratio);

    // Calculate target buffer size
    UINT32 targetBytesPerFrame = m_format.nBlockAlign;
    UINT32 targetSize = targetFrameCount * targetBytesPerFrame;

    // Verify destination buffer is large enough
    if (destSize < targetSize) {
        return false;
    }

    // Only support 32-bit float for now (most common for WASAPI)
    if (sourceFormat->wBitsPerSample == 32 && m_format.wBitsPerSample == 32) {
        const float* sourceSamples = reinterpret_cast<const float*>(data);
        float* targetSamples = reinterpret_cast<float*>(destBuffer);

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
        int16_t* targetSamples = reinterpret_cast<int16_t*>(destBuffer);

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
        // Unsupported format combination - return false to indicate failure
        return false;
    }

    return true;
}
