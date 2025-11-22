#include "WavWriter.h"
#include <cstring>

WavWriter::WavWriter()
    : m_dataSize(0)
    , m_totalDataSize(0)
    , m_filePartNumber(1)
    , m_dataStartPos(0)
{
}

WavWriter::~WavWriter() {
    Close();
}

bool WavWriter::Open(const std::wstring& filename, const WAVEFORMATEX* format) {
    if (m_file.is_open()) {
        return false;
    }

    m_filename = filename;
    m_dataSize = 0;

    // Extract base filename (remove extension) for multi-part file naming
    size_t extPos = filename.rfind(L'.');
    if (extPos != std::wstring::npos) {
        m_baseFilename = filename.substr(0, extPos);
    } else {
        m_baseFilename = filename;
    }

    // Reset file part number for new recording
    m_filePartNumber = 1;
    m_totalDataSize = 0;

    // Calculate format size
    UINT32 formatSize = sizeof(WAVEFORMATEX);
    if (format->wFormatTag == WAVE_FORMAT_EXTENSIBLE && format->cbSize >= 22) {
        formatSize = sizeof(WAVEFORMATEX) + format->cbSize;
    }

    // Store the format data
    m_formatData.resize(formatSize);
    std::memcpy(m_formatData.data(), format, formatSize);

    // Open file
    m_file.open(filename, std::ios::binary | std::ios::out);
    if (!m_file.is_open()) {
        return false;
    }

    // Write initial header (will be updated when closing)
    WriteWavHeader();
    m_dataStartPos = m_file.tellp();

    return true;
}

bool WavWriter::WriteData(const BYTE* data, UINT32 size) {
    if (!m_file.is_open()) {
        return false;
    }

    // Calculate current file size (header + data)
    // WAV header structure: RIFF(4) + size(4) + WAVE(4) + fmt(4) + fmtsize(4) + fmtdata + data(4) + datasize(4) + audiodata
    UINT64 currentFileSize = 12 + 8 + static_cast<UINT64>(m_formatData.size()) + 8 + m_dataSize;

    // Check if writing this data would exceed the 4GB limit
    if (currentFileSize + size > MAX_FILE_SIZE) {
        // Split to next file part before writing
        if (!SplitToNextFile()) {
            return false;
        }
    }

    m_file.write(reinterpret_cast<const char*>(data), size);
    m_dataSize += size;
    m_totalDataSize += size;

    return m_file.good();
}

bool WavWriter::SplitToNextFile() {
    // Update current file's header with final sizes
    UpdateWavHeader();

    // Close current file
    m_file.close();

    // Increment part number
    m_filePartNumber++;

    // Generate new filename: baseFilename_part2.wav, baseFilename_part3.wav, etc.
    wchar_t partSuffix[32];
    swprintf_s(partSuffix, L"_part%u.wav", m_filePartNumber);
    std::wstring newFilename = m_baseFilename + partSuffix;

    // Reset data size for new file
    m_dataSize = 0;

    // Open new file
    m_file.open(newFilename, std::ios::binary | std::ios::out);
    if (!m_file.is_open()) {
        return false;
    }

    // Write header for new file
    WriteWavHeader();
    m_dataStartPos = m_file.tellp();

    // Update current filename
    m_filename = newFilename;

    return true;
}

void WavWriter::Close() {
    if (!m_file.is_open()) {
        return;
    }

    // Update header with final size
    UpdateWavHeader();

    m_file.close();
    m_dataSize = 0;
}

void WavWriter::WriteWavHeader() {
    if (m_formatData.empty()) {
        return;
    }

    // Write RIFF header
    m_file.write("RIFF", 4);
    UINT32 riffSize = 0;  // Will be updated later
    m_file.write(reinterpret_cast<const char*>(&riffSize), 4);
    m_file.write("WAVE", 4);

    // Write fmt chunk
    m_file.write("fmt ", 4);
    UINT32 fmtSize = static_cast<UINT32>(m_formatData.size());
    m_file.write(reinterpret_cast<const char*>(&fmtSize), 4);
    m_file.write(reinterpret_cast<const char*>(m_formatData.data()), fmtSize);

    // Write data chunk header
    m_file.write("data", 4);
    UINT32 dataSize = 0;  // Will be updated later
    m_file.write(reinterpret_cast<const char*>(&dataSize), 4);
}

void WavWriter::UpdateWavHeader() {
    if (!m_file.is_open()) {
        return;
    }

    // Save current position
    auto currentPos = m_file.tellp();

    // Update RIFF size (offset 4)
    m_file.seekp(4);
    UINT32 riffSize = static_cast<UINT32>(currentPos) - 8;
    m_file.write(reinterpret_cast<const char*>(&riffSize), 4);

    // Update data size (offset = 12 + 4 + 4 + fmtSize + 4)
    // = 12 (RIFF header) + 8 (fmt chunk header) + fmtSize + 4 (data chunk ID)
    UINT32 dataSizeOffset = 12 + 8 + static_cast<UINT32>(m_formatData.size()) + 4;
    m_file.seekp(dataSizeOffset);
    m_file.write(reinterpret_cast<const char*>(&m_dataSize), 4);

    // Restore position
    m_file.seekp(currentPos);
}
