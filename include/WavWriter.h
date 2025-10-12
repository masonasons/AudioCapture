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

    std::ofstream m_file;
    std::wstring m_filename;
    std::vector<BYTE> m_formatData;  // Store full format (WAVEFORMATEX or WAVEFORMATEXTENSIBLE)
    UINT32 m_dataSize;
    std::streampos m_dataStartPos;
};
