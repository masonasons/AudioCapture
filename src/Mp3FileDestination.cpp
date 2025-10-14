#include "Mp3FileDestination.h"

Mp3FileDestination::Mp3FileDestination()
    : m_encoder(std::make_unique<Mp3Encoder>())
    , m_bitrate(192000)  // Default to 192 kbps
{
}

Mp3FileDestination::~Mp3FileDestination() {
    Close();
}

UINT32 Mp3FileDestination::ValidateBitrate(UINT32 bitrate) {
    // MP3 typically supports bitrates from 32 kbps to 320 kbps
    // Clamp to reasonable range
    if (bitrate < 32000) {
        return 32000;  // 32 kbps minimum
    }
    if (bitrate > 320000) {
        return 320000;  // 320 kbps maximum
    }

    // Round to common bitrates if close
    const UINT32 commonBitrates[] = {
        32000, 40000, 48000, 56000, 64000, 80000, 96000, 112000,
        128000, 160000, 192000, 224000, 256000, 320000
    };

    // Find closest common bitrate
    UINT32 closest = commonBitrates[0];
    UINT32 minDiff = abs((int)bitrate - (int)closest);

    for (UINT32 rate : commonBitrates) {
        UINT32 diff = abs((int)bitrate - (int)rate);
        if (diff < minDiff) {
            minDiff = diff;
            closest = rate;
        }
    }

    // If within 10% of a common bitrate, use the common one
    if (minDiff < bitrate / 10) {
        return closest;
    }

    return bitrate;
}

bool Mp3FileDestination::Configure(const WAVEFORMATEX* format, const DestinationConfig& config) {
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

    // Validate and store bitrate
    m_bitrate = ValidateBitrate(config.bitrate);

    // Generate file path with optional timestamp
    m_filePath = GenerateFilePath(config.outputPath, config.useTimestamp);

    // Ensure file has .mp3 extension
    if (m_filePath.length() < 4 ||
        m_filePath.substr(m_filePath.length() - 4) != L".mp3") {
        // If no extension or wrong extension, we still proceed
        // The user might have specified the full path intentionally
    }

    // Open the MP3 encoder
    if (!m_encoder->Open(m_filePath, format, m_bitrate)) {
        SetError(L"Failed to open MP3 encoder for file: " + m_filePath);
        return false;
    }

    // Start the async writer thread
    StartAsyncWriter();

    return true;
}

bool Mp3FileDestination::WriteAudioDataInternal(const BYTE* data, UINT32 size) {
    if (!IsOpen()) {
        SetError(L"Cannot write - MP3 encoder is not open");
        return false;
    }

    if (!data || size == 0) {
        // Silently ignore empty writes
        return true;
    }

    if (!m_encoder->WriteData(data, size)) {
        SetError(L"Failed to encode data to MP3 file");
        return false;
    }

    return true;
}

void Mp3FileDestination::Close() {
    // Stop the async writer thread (flushes pending writes)
    StopAsyncWriter();

    if (m_encoder) {
        m_encoder->Close();
    }
    m_filePath.clear();
}

bool Mp3FileDestination::IsOpen() const {
    return m_encoder && m_encoder->IsOpen();
}
