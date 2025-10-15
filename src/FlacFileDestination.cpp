#include "FlacFileDestination.h"

FlacFileDestination::FlacFileDestination()
    : m_encoder(std::make_unique<FlacEncoder>())
    , m_compressionLevel(5)  // Default to balanced compression
{
}

FlacFileDestination::~FlacFileDestination() {
    Close();
}

UINT32 FlacFileDestination::ValidateCompressionLevel(UINT32 level) {
    // FLAC supports compression levels 0-8
    if (level > 8) {
        return 8;
    }
    return level;
}

bool FlacFileDestination::Configure(const WAVEFORMATEX* format, const DestinationConfig& config) {
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

    // Validate and store compression level
    m_compressionLevel = ValidateCompressionLevel(config.compressionLevel);

    // Generate file path with optional timestamp
    m_filePath = GenerateFilePath(config.outputPath, config.useTimestamp);

    // Ensure directory exists
    if (!EnsureDirectoryExists(m_filePath)) {
        return false;  // Error already set by EnsureDirectoryExists
    }

    // Ensure file has .flac extension
    if (m_filePath.length() < 5 ||
        m_filePath.substr(m_filePath.length() - 5) != L".flac") {
        // If no extension or wrong extension, we still proceed
        // The user might have specified the full path intentionally
    }

    // Open the FLAC encoder
    if (!m_encoder->Open(m_filePath, format, m_compressionLevel)) {
        SetError(L"Failed to open FLAC encoder for file: " + m_filePath);
        return false;
    }

    // Start the async writer thread
    StartAsyncWriter();

    InitializeSilenceDetection(format, config);

    return true;
}

bool FlacFileDestination::WriteAudioDataInternal(const BYTE* data, UINT32 size) {
    if (!IsOpen()) {
        SetError(L"Cannot write - FLAC encoder is not open");
        return false;
    }

    if (!data || size == 0) {
        // Silently ignore empty writes
        return true;
    }

    if (!m_encoder->WriteData(data, size)) {
        SetError(L"Failed to encode data to FLAC file");
        return false;
    }

    return true;
}

void FlacFileDestination::Close() {
    // Stop the async writer thread (flushes pending writes)
    StopAsyncWriter();

    if (m_encoder) {
        m_encoder->Close();
    }
    m_filePath.clear();
}

bool FlacFileDestination::IsOpen() const {
    return m_encoder && m_encoder->IsOpen();
}
