#include "WavFileDestination.h"
#include "DebugLogger.h"

WavFileDestination::WavFileDestination()
    : m_writer(std::make_unique<WavWriter>())
{
}

WavFileDestination::~WavFileDestination() {
    Close();
}

bool WavFileDestination::Configure(const WAVEFORMATEX* format, const DestinationConfig& config) {
    // Clear any previous errors
    ClearError();

    // Validate inputs
    if (!ValidateFormat(format)) {
        return false;
    }

    if (!ValidateFilePath(config.outputPath)) {
        return false;
    }

    // Close any previously opened file
    if (IsOpen()) {
        Close();
    }

    // Generate file path with optional timestamp
    m_filePath = GenerateFilePath(config.outputPath, config.useTimestamp);

    // Open the WAV file
    if (!m_writer->Open(m_filePath, format)) {
        SetError(L"Failed to open WAV file: " + m_filePath);
        return false;
    }

    // Start the async writer thread
    StartAsyncWriter();

    InitializeSilenceDetection(format, config);

    return true;
}

bool WavFileDestination::WriteAudioDataInternal(const BYTE* data, UINT32 size) {
    // DEBUG: Track writes to diagnose 0-byte files
    static std::atomic<int> writeCount{0};
    static std::atomic<uint64_t> totalWritten{0};
    int writeNum = ++writeCount;
    totalWritten += size;

    if (writeNum == 1 || writeNum % 100 == 0) {
        wchar_t debugMsg[512];
        swprintf_s(debugMsg, L"[WAV] Write #%d: %u bytes, Total=%.2f MB, File=%s",
                   writeNum, size, totalWritten.load() / (1024.0 * 1024.0), m_filePath.c_str());
        DebugLog(debugMsg);
    }

    if (!IsOpen()) {
        DebugLog(L"[WAV] ERROR: WriteAudioDataInternal called but file is NOT OPEN!");
        SetError(L"Cannot write - WAV file is not open");
        return false;
    }

    if (!data || size == 0) {
        // Silently ignore empty writes
        return true;
    }

    if (!m_writer->WriteData(data, size)) {
        DebugLog(L"[WAV] ERROR: m_writer->WriteData() FAILED!");
        SetError(L"Failed to write data to WAV file");
        return false;
    }

    return true;
}

void WavFileDestination::Close() {
    // Stop the async writer thread (flushes pending writes)
    StopAsyncWriter();

    if (m_writer) {
        m_writer->Close();
    }
    m_filePath.clear();
}

bool WavFileDestination::IsOpen() const {
    return m_writer && m_writer->IsOpen();
}
