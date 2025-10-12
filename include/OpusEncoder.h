#pragma once

#include <windows.h>
#include <mmreg.h>
#include <string>
#include <fstream>
#include <vector>
#include <opus/opus.h>
#include <ogg/ogg.h>

class OpusEncoder {
public:
    OpusEncoder();
    ~OpusEncoder();

    // Open Opus file for writing (OGG container)
    bool Open(const std::wstring& filename, const WAVEFORMATEX* format, UINT32 bitrate = 128000);

    // Write audio data (PCM format)
    bool WriteData(const BYTE* data, UINT32 size);

    // Close file and finalize
    void Close();

    // Check if file is open
    bool IsOpen() const { return m_file.is_open(); }

private:
    bool InitializeOggStream();
    bool WriteOggHeaders();
    bool WriteOggPage(bool flush = false);
    bool EncodeBuffer();
    void WriteInt32LE(std::vector<unsigned char>& data, int32_t value);

    std::ofstream m_file;
    std::wstring m_filename;
    WAVEFORMATEX m_format;

    // Opus encoder and OGG stream
    OpusEncoder* m_opusEncoder;
    ogg_stream_state m_oggStream;
    ogg_page m_oggPage;
    ogg_packet m_oggPacket;

    std::vector<BYTE> m_buffer;
    UINT32 m_samplesPerFrame;
    UINT32 m_bitrate;
    UINT64 m_totalSamples;
    int m_serialno;
    int64_t m_granulePos;
    int64_t m_packetCount;
};
