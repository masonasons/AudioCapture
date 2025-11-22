#pragma once

#include <windows.h>
#include <mmreg.h>
#include <ks.h>
#include <ksmedia.h>
#include <string>
#include <fstream>
#include <vector>

class WavWriter {
public:
    WavWriter();
    ~WavWriter();

    // Open WAV file for writing
    bool Open(const std::wstring& filename, const WAVEFORMATEX* format);

    // Write audio data
    bool WriteData(const BYTE* data, UINT32 size);

    // Close file and finalize WAV header
    void Close();

    // Check if file is open
    bool IsOpen() const { return m_file.is_open(); }

private:
    void WriteWavHeader();
    void UpdateWavHeader();
    bool SplitToNextFile();  // Open next file part seamlessly

    static constexpr UINT64 MAX_FILE_SIZE = 4000000000ULL; // ~3.7GB safety limit

    std::ofstream m_file;
    std::wstring m_filename;
    std::wstring m_baseFilename;     // Base filename without extension
    std::vector<BYTE> m_formatData;  // Store full format (WAVEFORMATEX or WAVEFORMATEXTENSIBLE)
    UINT32 m_dataSize;               // Data size in current file
    UINT64 m_totalDataSize;          // Total data written across all parts
    UINT32 m_filePartNumber;         // Current file part (1, 2, 3...)
    std::streampos m_dataStartPos;
};
