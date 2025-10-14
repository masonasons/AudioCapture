#include "OutputDestination.h"
#include <ks.h>
#include <ksmedia.h>
#include <cmath>

OutputDestination::OutputDestination()
    : m_writerRunning(false)
    , m_isOpen(false)
    , m_isPaused(false)
    , m_skipSilence(false)
    , m_silenceThreshold(0.01f)
    , m_silenceDurationMs(1000)
    , m_silenceDurationSamples(0)
    , m_consecutiveSilentSamples(0)
    , m_format(nullptr)
{
}

OutputDestination::~OutputDestination() {
    StopAsyncWriter();
    if (m_format) {
        CoTaskMemFree(m_format);
        m_format = nullptr;
    }
}

bool OutputDestination::WriteAudioData(const BYTE* data, UINT32 size) {
    if (!m_isOpen.load(std::memory_order_acquire) || !data || size == 0) {
        return false;
    }

    // If paused, don't queue any data
    if (m_isPaused.load(std::memory_order_acquire)) {
        return true;  // Return success but don't queue
    }

    // Check for silence if skip silence is enabled
    if (m_skipSilence && m_format) {
        if (IsSilent(data, size)) {
            // Count consecutive silent samples
            UINT32 numSamples = size / m_format->nBlockAlign;
            m_consecutiveSilentSamples += numSamples;

            // If we've accumulated enough silence, skip this buffer
            if (m_consecutiveSilentSamples >= m_silenceDurationSamples) {
                return true;  // Skip writing
            }
        } else {
            // Reset silence counter on non-silent audio
            m_consecutiveSilentSamples = 0;
        }
    }

    // CRITICAL: This is NON-BLOCKING. We just copy the data and queue it.
    // The actual disk I/O happens on the writer thread.
    AudioChunk chunk;
    chunk.data.assign(data, data + size);

    {
        std::lock_guard<std::mutex> lock(m_queueMutex);
        m_writeQueue.push(std::move(chunk));
    }

    // Wake up the writer thread
    m_queueCV.notify_one();

    return true;
}

void OutputDestination::StartAsyncWriter() {
    // Stop any existing writer first
    StopAsyncWriter();

    m_isOpen.store(true, std::memory_order_release);
    m_writerRunning.store(true, std::memory_order_release);
    m_writerThread = std::thread(&OutputDestination::WriterThreadFunc, this);
}

void OutputDestination::StopAsyncWriter() {
    // Signal the writer thread to stop
    m_writerRunning.store(false, std::memory_order_release);
    m_isOpen.store(false, std::memory_order_release);
    m_queueCV.notify_one();

    // Wait for writer thread to finish
    if (m_writerThread.joinable()) {
        m_writerThread.join();
    }

    // Clear any remaining queued data
    {
        std::lock_guard<std::mutex> lock(m_queueMutex);
        while (!m_writeQueue.empty()) {
            m_writeQueue.pop();
        }
    }
}

void OutputDestination::WriterThreadFunc() {
    while (m_writerRunning.load(std::memory_order_acquire)) {
        AudioChunk chunk;
        bool hasData = false;

        // Wait for data or stop signal
        {
            std::unique_lock<std::mutex> lock(m_queueMutex);
            m_queueCV.wait(lock, [this] {
                return !m_writeQueue.empty() || !m_writerRunning.load(std::memory_order_acquire);
            });

            if (!m_writeQueue.empty()) {
                chunk = std::move(m_writeQueue.front());
                m_writeQueue.pop();
                hasData = true;
            }
        }

        // Perform the actual write OUTSIDE the mutex
        // This allows audio callbacks to continue queuing data
        if (hasData) {
            WriteAudioDataInternal(chunk.data.data(), static_cast<UINT32>(chunk.data.size()));
        }
    }

    // Flush any remaining data before exiting
    while (true) {
        AudioChunk chunk;
        bool hasData = false;

        {
            std::lock_guard<std::mutex> lock(m_queueMutex);
            if (!m_writeQueue.empty()) {
                chunk = std::move(m_writeQueue.front());
                m_writeQueue.pop();
                hasData = true;
            }
        }

        if (!hasData) {
            break;
        }

        WriteAudioDataInternal(chunk.data.data(), static_cast<UINT32>(chunk.data.size()));
    }
}

void OutputDestination::Pause() {
    m_isPaused.store(true, std::memory_order_release);
}

void OutputDestination::Resume() {
    m_isPaused.store(false, std::memory_order_release);
}

bool OutputDestination::IsSilent(const BYTE* data, UINT32 size) {
    if (!m_format || !data || size == 0) {
        return false;
    }

    // Determine format type
    bool isFloat = false;
    WORD bitsPerSample = m_format->wBitsPerSample;

    if (m_format->wFormatTag == WAVE_FORMAT_EXTENSIBLE && m_format->cbSize >= 22) {
        WAVEFORMATEXTENSIBLE* wfex = reinterpret_cast<WAVEFORMATEXTENSIBLE*>(m_format);
        isFloat = (wfex->SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT);
    } else {
        isFloat = (m_format->wFormatTag == WAVE_FORMAT_IEEE_FLOAT);
    }

    // Calculate maximum amplitude in this buffer
    float maxAmplitude = 0.0f;

    if (isFloat && bitsPerSample == 32) {
        // 32-bit float PCM
        const float* samples = reinterpret_cast<const float*>(data);
        UINT32 numSamples = size / sizeof(float);
        for (UINT32 i = 0; i < numSamples; i++) {
            float amplitude = std::abs(samples[i]);
            if (amplitude > maxAmplitude) {
                maxAmplitude = amplitude;
            }
        }
    } else if (bitsPerSample == 16) {
        // 16-bit PCM
        const int16_t* samples = reinterpret_cast<const int16_t*>(data);
        UINT32 numSamples = size / sizeof(int16_t);
        for (UINT32 i = 0; i < numSamples; i++) {
            float amplitude = std::abs(samples[i]) / 32768.0f;  // Normalize to 0.0-1.0
            if (amplitude > maxAmplitude) {
                maxAmplitude = amplitude;
            }
        }
    } else {
        // Unsupported format - don't skip
        return false;
    }

    // Check if this buffer is below the silence threshold
    return (maxAmplitude < m_silenceThreshold);
}
