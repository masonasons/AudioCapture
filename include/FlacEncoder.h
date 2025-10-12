#pragma once

#include <windows.h>
#include <mmreg.h>
#include <string>
#include <fstream>
#include <vector>
#include <FLAC/stream_encoder.h>

class FlacEncoder {
public:
    FlacEncoder();
    ~FlacEncoder();

    // Open FLAC file for writing
    bool Open(const std::wstring& filename, const WAVEFORMATEX* format, UINT32 compressionLevel = 5);

    // Write audio data (PCM format)
    bool WriteData(const BYTE* data, UINT32 size);

    // Close file and finalize
    void Close();

    // Check if file is open
    bool IsOpen() const { return m_encoder != nullptr; }

private:
    static FLAC__StreamEncoderWriteStatus WriteCallback(
        const FLAC__StreamEncoder* encoder,
        const FLAC__byte buffer[],
        size_t bytes,
        unsigned samples,
        unsigned current_frame,
        void* client_data);

    static FLAC__StreamEncoderSeekStatus SeekCallback(
        const FLAC__StreamEncoder* encoder,
        FLAC__uint64 absolute_byte_offset,
        void* client_data);

    static FLAC__StreamEncoderTellStatus TellCallback(
        const FLAC__StreamEncoder* encoder,
        FLAC__uint64* absolute_byte_offset,
        void* client_data);

    bool ProcessBuffer();

    std::ofstream m_file;
    std::wstring m_filename;
    WAVEFORMATEX m_format;

    FLAC__StreamEncoder* m_encoder;
    std::vector<BYTE> m_buffer;
    UINT32 m_samplesPerFrame;
    UINT32 m_compressionLevel;
    UINT64 m_totalSamples;
};
