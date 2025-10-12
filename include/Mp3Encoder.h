#pragma once

#include <windows.h>
#include <string>
#include <mfapi.h>
#include <mfidl.h>
#include <mfreadwrite.h>
#include <vector>

class Mp3Encoder {
public:
    Mp3Encoder();
    ~Mp3Encoder();

    // Open MP3 file for writing
    bool Open(const std::wstring& filename, const WAVEFORMATEX* format, UINT32 bitrate = 192000);

    // Write audio data (PCM format)
    bool WriteData(const BYTE* data, UINT32 size);

    // Close file and finalize
    void Close();

    // Check if file is open
    bool IsOpen() const { return m_sinkWriter != nullptr; }

private:
    IMFSinkWriter* m_sinkWriter;
    DWORD m_streamIndex;
    WAVEFORMATEX m_inputFormat;
    LONGLONG m_sampleDuration;
    UINT64 m_rtStart;
    std::vector<BYTE> m_buffer;
    UINT32 m_samplesPerFrame;
};
