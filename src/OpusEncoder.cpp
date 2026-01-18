#include "OpusEncoder.h"
#include <cstring>
#include <ctime>
#include <algorithm>

OpusOggEncoder::OpusOggEncoder()
    : m_opusEncoder(nullptr)
    , m_samplesPerFrame(960) // 20ms at 48kHz
    , m_bitrate(128000)
    , m_totalSamples(0)
    , m_serialno(0)
    , m_granulePos(0)
    , m_packetCount(0)
{
    std::memset(&m_format, 0, sizeof(WAVEFORMATEX));
    std::memset(&m_oggStream, 0, sizeof(ogg_stream_state));
    std::memset(&m_oggPage, 0, sizeof(ogg_page));
    std::memset(&m_oggPacket, 0, sizeof(ogg_packet));
}

OpusOggEncoder::~OpusOggEncoder() {
    Close();
}

void OpusOggEncoder::WriteInt32LE(std::vector<unsigned char>& data, int32_t value) {
    data.push_back(value & 0xFF);
    data.push_back((value >> 8) & 0xFF);
    data.push_back((value >> 16) & 0xFF);
    data.push_back((value >> 24) & 0xFF);
}

bool OpusOggEncoder::Open(const std::wstring& filename, const WAVEFORMATEX* format, UINT32 bitrate) {
    if (m_file.is_open()) {
        return false;
    }

    m_filename = filename;
    m_format = *format;
    m_bitrate = bitrate;
    m_totalSamples = 0;
    m_granulePos = 0;
    m_packetCount = 0;

    // Opus only supports 48kHz for encoding (or 24, 16, 12, 8 kHz)
    // We'll use 48kHz as it's the highest quality
    int opusSampleRate = 48000;
    int opusChannels = std::min((int)format->nChannels, 2); // Stereo max

    // Create Opus encoder
    int error = 0;
    m_opusEncoder = (::OpusEncoder*)opus_encoder_create(opusSampleRate, opusChannels, OPUS_APPLICATION_AUDIO, &error);
    if (error != OPUS_OK || !m_opusEncoder) {
        return false;
    }

    // Configure encoder
    opus_encoder_ctl(m_opusEncoder, OPUS_SET_BITRATE(bitrate));
    opus_encoder_ctl(m_opusEncoder, OPUS_SET_VBR(1)); // Variable bitrate
    opus_encoder_ctl(m_opusEncoder, OPUS_SET_COMPLEXITY(10)); // Max quality

    // Frame size is 20ms at 48kHz = 960 samples
    m_samplesPerFrame = 960;

    // Initialize OGG stream with random serial number
    m_serialno = static_cast<int>(static_cast<int64_t>(std::time(nullptr)) & 0x7fffffff);
    if (ogg_stream_init(&m_oggStream, m_serialno) != 0) {
        opus_encoder_destroy(m_opusEncoder);
        m_opusEncoder = nullptr;
        return false;
    }

    // Open output file
    m_file.open(filename, std::ios::binary | std::ios::out);
    if (!m_file.is_open()) {
        opus_encoder_destroy(m_opusEncoder);
        ogg_stream_clear(&m_oggStream);
        m_opusEncoder = nullptr;
        return false;
    }

    // Write Opus headers
    if (!WriteOggHeaders()) {
        Close();
        return false;
    }

    return true;
}

bool OpusOggEncoder::WriteOggHeaders() {
    // Create OpusHead header
    std::vector<unsigned char> opusHead;
    opusHead.push_back('O');
    opusHead.push_back('p');
    opusHead.push_back('u');
    opusHead.push_back('s');
    opusHead.push_back('H');
    opusHead.push_back('e');
    opusHead.push_back('a');
    opusHead.push_back('d');
    opusHead.push_back(1); // Version
    opusHead.push_back(static_cast<unsigned char>(m_format.nChannels)); // Channel count
    opusHead.push_back(0); // Pre-skip LSB
    opusHead.push_back(0); // Pre-skip MSB

    // Original sample rate (little endian)
    WriteInt32LE(opusHead, static_cast<int32_t>(m_format.nSamplesPerSec));

    // Output gain (0 dB)
    opusHead.push_back(0);
    opusHead.push_back(0);

    // Channel mapping family (0 = mono/stereo)
    opusHead.push_back(0);

    // Create OGG packet for OpusHead
    m_oggPacket.packet = opusHead.data();
    m_oggPacket.bytes = static_cast<long>(opusHead.size());
    m_oggPacket.b_o_s = 1; // Beginning of stream
    m_oggPacket.e_o_s = 0;
    m_oggPacket.granulepos = 0;
    m_oggPacket.packetno = m_packetCount++;

    // Submit packet to OGG stream
    if (ogg_stream_packetin(&m_oggStream, &m_oggPacket) != 0) {
        return false;
    }

    // Write OGG page
    while (ogg_stream_flush(&m_oggStream, &m_oggPage) != 0) {
        m_file.write(reinterpret_cast<const char*>(m_oggPage.header), m_oggPage.header_len);
        m_file.write(reinterpret_cast<const char*>(m_oggPage.body), m_oggPage.body_len);
    }

    // Create OpusTags header
    std::vector<unsigned char> opusTags;
    opusTags.push_back('O');
    opusTags.push_back('p');
    opusTags.push_back('u');
    opusTags.push_back('s');
    opusTags.push_back('T');
    opusTags.push_back('a');
    opusTags.push_back('g');
    opusTags.push_back('s');

    // Vendor string
    std::string vendor = "AudioCapture 1.0";
    WriteInt32LE(opusTags, static_cast<int32_t>(vendor.size()));
    for (char c : vendor) {
        opusTags.push_back(c);
    }

    // User comment list length (0 comments)
    WriteInt32LE(opusTags, 0);

    // Create OGG packet for OpusTags
    m_oggPacket.packet = opusTags.data();
    m_oggPacket.bytes = static_cast<long>(opusTags.size());
    m_oggPacket.b_o_s = 0;
    m_oggPacket.e_o_s = 0;
    m_oggPacket.granulepos = 0;
    m_oggPacket.packetno = m_packetCount++;

    // Submit packet to OGG stream
    if (ogg_stream_packetin(&m_oggStream, &m_oggPacket) != 0) {
        return false;
    }

    // Write OGG page
    while (ogg_stream_flush(&m_oggStream, &m_oggPage) != 0) {
        m_file.write(reinterpret_cast<const char*>(m_oggPage.header), m_oggPage.header_len);
        m_file.write(reinterpret_cast<const char*>(m_oggPage.body), m_oggPage.body_len);
    }

    return true;
}

bool OpusOggEncoder::WriteData(const BYTE* data, UINT32 size) {
    if (!m_file.is_open() || !m_opusEncoder) {
        return false;
    }

    // Add data to buffer
    m_buffer.insert(m_buffer.end(), data, data + size);

    // Process complete frames
    UINT32 frameSize = m_samplesPerFrame * m_format.nBlockAlign;

    while (m_buffer.size() >= frameSize) {
        if (!EncodeBuffer()) {
            return false;
        }

        // Remove processed data from buffer
        m_buffer.erase(m_buffer.begin(), m_buffer.begin() + frameSize);
    }

    return true;
}

bool OpusOggEncoder::EncodeBuffer() {
    // Prepare PCM samples for encoding
    int frameSamples = m_samplesPerFrame;
    int channels = m_format.nChannels;

    // Convert PCM bytes to float samples for Opus
    std::vector<float> pcmFloat(frameSamples * channels);

    if (m_format.wBitsPerSample == 16) {
        // Convert 16-bit PCM to float
        const int16_t* pcm16 = reinterpret_cast<const int16_t*>(m_buffer.data());
        for (int i = 0; i < frameSamples * channels; i++) {
            pcmFloat[i] = pcm16[i] / 32768.0f;
        }
    }
    else if (m_format.wBitsPerSample == 32) {
        // Assume float PCM
        std::memcpy(pcmFloat.data(), m_buffer.data(), frameSamples * channels * sizeof(float));
    }
    else {
        return false; // Unsupported format
    }

    // Encode frame
    std::vector<unsigned char> opusData(4000); // Max Opus packet size
    int encodedBytes = opus_encode_float(m_opusEncoder, pcmFloat.data(), frameSamples,
                                         opusData.data(), static_cast<opus_int32>(opusData.size()));

    if (encodedBytes < 0) {
        return false; // Encoding error
    }

    // Update granule position
    m_granulePos += frameSamples;

    // Create OGG packet
    m_oggPacket.packet = opusData.data();
    m_oggPacket.bytes = static_cast<long>(encodedBytes);
    m_oggPacket.b_o_s = 0;
    m_oggPacket.e_o_s = 0;
    m_oggPacket.granulepos = m_granulePos;
    m_oggPacket.packetno = m_packetCount++;

    // Submit packet to OGG stream
    if (ogg_stream_packetin(&m_oggStream, &m_oggPacket) != 0) {
        return false;
    }

    // Write OGG pages
    while (ogg_stream_pageout(&m_oggStream, &m_oggPage) != 0) {
        m_file.write(reinterpret_cast<const char*>(m_oggPage.header), m_oggPage.header_len);
        m_file.write(reinterpret_cast<const char*>(m_oggPage.body), m_oggPage.body_len);
    }

    m_totalSamples += frameSamples;
    return true;
}

void OpusOggEncoder::Close() {
    if (!m_file.is_open()) {
        return;
    }

    // Encode any remaining samples
    if (!m_buffer.empty()) {
        // Pad buffer to frame size
        UINT32 frameSize = m_samplesPerFrame * m_format.nBlockAlign;
        if (m_buffer.size() < frameSize) {
            m_buffer.resize(frameSize, 0);
        }
        EncodeBuffer();
    }

    // Write final OGG packet with e_o_s flag
    if (m_opusEncoder) {
        std::vector<unsigned char> emptyPacket(4000);
        int encodedBytes = opus_encode_float(m_opusEncoder, nullptr, 0,
                                             emptyPacket.data(), static_cast<opus_int32>(emptyPacket.size()));

        if (encodedBytes > 0) {
            m_oggPacket.packet = emptyPacket.data();
            m_oggPacket.bytes = static_cast<long>(encodedBytes);
            m_oggPacket.b_o_s = 0;
            m_oggPacket.e_o_s = 1; // End of stream
            m_oggPacket.granulepos = m_granulePos;
            m_oggPacket.packetno = m_packetCount++;

            ogg_stream_packetin(&m_oggStream, &m_oggPacket);
        }
    }

    // Flush remaining OGG pages
    while (ogg_stream_flush(&m_oggStream, &m_oggPage) != 0) {
        m_file.write(reinterpret_cast<const char*>(m_oggPage.header), m_oggPage.header_len);
        m_file.write(reinterpret_cast<const char*>(m_oggPage.body), m_oggPage.body_len);
    }

    // Clean up
    if (m_opusEncoder) {
        opus_encoder_destroy(m_opusEncoder);
        m_opusEncoder = nullptr;
    }

    ogg_stream_clear(&m_oggStream);
    m_file.close();
    m_buffer.clear();
}

bool OpusOggEncoder::InitializeOggStream() {
    // Already done in Open()
    return true;
}

bool OpusOggEncoder::WriteOggPage(bool /*flush*/) {
    // Already handled in EncodeBuffer()
    return true;
}
