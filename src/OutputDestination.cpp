#include "OutputDestination.h"

OutputDestination::OutputDestination()
    : m_writerRunning(false)
    , m_isOpen(false)
{
}

OutputDestination::~OutputDestination() {
    StopAsyncWriter();
}

bool OutputDestination::WriteAudioData(const BYTE* data, UINT32 size) {
    if (!m_isOpen.load(std::memory_order_acquire) || !data || size == 0) {
        return false;
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
