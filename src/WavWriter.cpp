#include "WavWriter.h"
#include <cstring>

WavWriter::WavWriter()
    : m_dataSize(0)
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

    m_file.write(reinterpret_cast<const char*>(data), size);
    m_dataSize += size;

    return m_file.good();
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
