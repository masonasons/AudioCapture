#include "OpusFileDestination.h"

OpusFileDestination::OpusFileDestination()
    : m_encoder(std::make_unique<OpusEncoder>())
    , m_bitrate(128000)  // Default to 128 kbps
{
}

OpusFileDestination::~OpusFileDestination() {
    Close();
}

UINT32 OpusFileDestination::ValidateBitrate(UINT32 bitrate) {
    // Opus supports 6 kbps to 510 kbps, but for audio capture
    // we use a more practical range
    if (bitrate < 64000) {
        return 64000;  // 64 kbps minimum for quality audio
    }
    if (bitrate > 256000) {
        return 256000;  // 256 kbps maximum (diminishing returns above this)
    }

    // Round to common bitrates if close
    const UINT32 commonBitrates[] = {
        64000, 96000, 128000, 160000, 192000, 256000
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

bool OpusFileDestination::Configure(const WAVEFORMATEX* format, const DestinationConfig& config) {
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

    // Ensure file has .opus extension (Opus in OGG container)
    if (m_filePath.length() < 5 ||
        m_filePath.substr(m_filePath.length() - 5) != L".opus") {
        // If no extension or wrong extension, we still proceed
        // The user might have specified the full path intentionally
    }

    // Open the Opus encoder
    if (!m_encoder->Open(m_filePath, format, m_bitrate)) {
        SetError(L"Failed to open Opus encoder for file: " + m_filePath);
        return false;
    }

    // Start the async writer thread
    StartAsyncWriter();

    return true;
}

bool OpusFileDestination::WriteAudioDataInternal(const BYTE* data, UINT32 size) {
    if (!IsOpen()) {
        SetError(L"Cannot write - Opus encoder is not open");
        return false;
    }

    if (!data || size == 0) {
        // Silently ignore empty writes
        return true;
    }

    if (!m_encoder->WriteData(data, size)) {
        SetError(L"Failed to encode data to Opus file");
        return false;
    }

    return true;
}

void OpusFileDestination::Close() {
    // Stop the async writer thread (flushes pending writes)
    StopAsyncWriter();

    if (m_encoder) {
        m_encoder->Close();
    }
    m_filePath.clear();
}

bool OpusFileDestination::IsOpen() const {
    return m_encoder && m_encoder->IsOpen();
}
