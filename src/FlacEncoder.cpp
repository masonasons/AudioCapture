#include "FlacEncoder.h"
#include <algorithm>
#include <cstring>

FlacEncoder::FlacEncoder()
    : m_encoder(nullptr)
    , m_samplesPerFrame(0)
    , m_compressionLevel(5)
    , m_totalSamples(0) {
    memset(&m_format, 0, sizeof(m_format));
}

FlacEncoder::~FlacEncoder() {
    Close();
}

bool FlacEncoder::Open(const std::wstring& filename, const WAVEFORMATEX* format, UINT32 compressionLevel) {
    if (!format || m_encoder != nullptr) {
        return false;
    }

    m_filename = filename;
    memcpy(&m_format, format, sizeof(WAVEFORMATEX));
    m_compressionLevel = std::min(compressionLevel, 8u);

    // Open output file
    m_file.open(filename, std::ios::binary);
    if (!m_file.is_open()) {
        return false;
    }

    // Create FLAC encoder
    m_encoder = FLAC__stream_encoder_new();
    if (!m_encoder) {
        m_file.close();
        return false;
    }

    // Configure encoder
    FLAC__stream_encoder_set_channels(m_encoder, m_format.nChannels);
    // Use 24-bit for optimal quality with float input (32-bit float -> 24-bit FLAC)
    // FLAC doesn't support more than 24 bits per sample effectively
    UINT32 flacBitsPerSample = (m_format.wBitsPerSample == 32) ? 24 : m_format.wBitsPerSample;
    FLAC__stream_encoder_set_bits_per_sample(m_encoder, flacBitsPerSample);
    FLAC__stream_encoder_set_sample_rate(m_encoder, m_format.nSamplesPerSec);
    FLAC__stream_encoder_set_compression_level(m_encoder, m_compressionLevel);

    // Disable verification for better performance during live capture
    FLAC__stream_encoder_set_verify(m_encoder, false);

    // Initialize encoder with callbacks
    FLAC__StreamEncoderInitStatus init_status = FLAC__stream_encoder_init_stream(
        m_encoder,
        WriteCallback,
        SeekCallback,
        TellCallback,
        nullptr,  // metadata callback
        this      // client data
    );

    if (init_status != FLAC__STREAM_ENCODER_INIT_STATUS_OK) {
        FLAC__stream_encoder_delete(m_encoder);
        m_encoder = nullptr;
        m_file.close();
        return false;
    }

    // Calculate frame size (1024 samples is a good default for FLAC)
    m_samplesPerFrame = 1024;
    m_totalSamples = 0;

    return true;
}

FLAC__StreamEncoderWriteStatus FlacEncoder::WriteCallback(
    const FLAC__StreamEncoder* encoder,
    const FLAC__byte buffer[],
    size_t bytes,
    unsigned samples,
    unsigned current_frame,
    void* client_data) {

    (void)encoder;
    (void)samples;
    (void)current_frame;

    FlacEncoder* self = static_cast<FlacEncoder*>(client_data);
    if (!self || !self->m_file.is_open()) {
        return FLAC__STREAM_ENCODER_WRITE_STATUS_FATAL_ERROR;
    }

    self->m_file.write(reinterpret_cast<const char*>(buffer), bytes);
    if (self->m_file.fail()) {
        return FLAC__STREAM_ENCODER_WRITE_STATUS_FATAL_ERROR;
    }

    return FLAC__STREAM_ENCODER_WRITE_STATUS_OK;
}

FLAC__StreamEncoderSeekStatus FlacEncoder::SeekCallback(
    const FLAC__StreamEncoder* encoder,
    FLAC__uint64 absolute_byte_offset,
    void* client_data) {

    (void)encoder;

    FlacEncoder* self = static_cast<FlacEncoder*>(client_data);
    if (!self || !self->m_file.is_open()) {
        return FLAC__STREAM_ENCODER_SEEK_STATUS_ERROR;
    }

    self->m_file.seekp(static_cast<std::streamoff>(absolute_byte_offset), std::ios::beg);
    if (self->m_file.fail()) {
        return FLAC__STREAM_ENCODER_SEEK_STATUS_ERROR;
    }

    return FLAC__STREAM_ENCODER_SEEK_STATUS_OK;
}

FLAC__StreamEncoderTellStatus FlacEncoder::TellCallback(
    const FLAC__StreamEncoder* encoder,
    FLAC__uint64* absolute_byte_offset,
    void* client_data) {

    (void)encoder;

    FlacEncoder* self = static_cast<FlacEncoder*>(client_data);
    if (!self || !self->m_file.is_open()) {
        return FLAC__STREAM_ENCODER_TELL_STATUS_ERROR;
    }

    *absolute_byte_offset = static_cast<FLAC__uint64>(self->m_file.tellp());
    if (self->m_file.fail()) {
        return FLAC__STREAM_ENCODER_TELL_STATUS_ERROR;
    }

    return FLAC__STREAM_ENCODER_TELL_STATUS_OK;
}

bool FlacEncoder::WriteData(const BYTE* data, UINT32 size) {
    if (!m_encoder || size == 0) {
        return false;
    }

    // Add data to buffer
    m_buffer.insert(m_buffer.end(), data, data + size);

    // Process complete frames
    return ProcessBuffer();
}

bool FlacEncoder::ProcessBuffer() {
    if (!m_encoder) {
        return false;
    }

    UINT32 bytesPerSample = m_format.wBitsPerSample / 8;
    UINT32 bytesPerFrame = m_samplesPerFrame * m_format.nChannels * bytesPerSample;

    while (m_buffer.size() >= bytesPerFrame) {
        // Convert PCM data to FLAC format (32-bit integers)
        std::vector<FLAC__int32> flacBuffer(m_samplesPerFrame * m_format.nChannels);

        for (UINT32 i = 0; i < m_samplesPerFrame * m_format.nChannels; i++) {
            FLAC__int32 sample = 0;

            if (bytesPerSample == 2) {
                // 16-bit PCM
                int16_t s = *reinterpret_cast<const int16_t*>(&m_buffer[i * 2]);
                sample = static_cast<FLAC__int32>(s);
            } else if (bytesPerSample == 4) {
                // Check if this is floating point audio (common in Windows)
                // Windows often uses 32-bit float format
                float f = *reinterpret_cast<const float*>(&m_buffer[i * 4]);

                // Check if it looks like a float (typical range -1.0 to 1.0)
                if (f >= -1.0f && f <= 1.0f) {
                    // Convert float [-1.0, 1.0] to 24-bit integer for FLAC
                    // Using 24-bit gives better quality without unnecessary precision
                    sample = static_cast<FLAC__int32>(f * 8388607.0f); // 2^23 - 1
                } else {
                    // Assume it's 32-bit integer PCM
                    sample = *reinterpret_cast<const int32_t*>(&m_buffer[i * 4]);
                }
            }

            flacBuffer[i] = sample;
        }

        // Prepare channel buffers
        std::vector<FLAC__int32*> channelBuffers(m_format.nChannels);
        std::vector<std::vector<FLAC__int32>> tempBuffers(m_format.nChannels);

        for (UINT32 ch = 0; ch < m_format.nChannels; ch++) {
            tempBuffers[ch].resize(m_samplesPerFrame);
            channelBuffers[ch] = tempBuffers[ch].data();

            // Deinterleave samples
            for (UINT32 i = 0; i < m_samplesPerFrame; i++) {
                tempBuffers[ch][i] = flacBuffer[i * m_format.nChannels + ch];
            }
        }

        // Encode frame
        if (!FLAC__stream_encoder_process(m_encoder, channelBuffers.data(), m_samplesPerFrame)) {
            return false;
        }

        m_totalSamples += m_samplesPerFrame;

        // Remove processed data from buffer
        m_buffer.erase(m_buffer.begin(), m_buffer.begin() + bytesPerFrame);
    }

    return true;
}

void FlacEncoder::Close() {
    if (m_encoder) {
        // Process any remaining buffered data
        if (!m_buffer.empty()) {
            UINT32 bytesPerSample = m_format.wBitsPerSample / 8;
            UINT32 remainingSamples = static_cast<UINT32>(m_buffer.size()) / (m_format.nChannels * bytesPerSample);

            if (remainingSamples > 0) {
                std::vector<FLAC__int32> flacBuffer(remainingSamples * m_format.nChannels);

                for (UINT32 i = 0; i < remainingSamples * m_format.nChannels; i++) {
                    FLAC__int32 sample = 0;

                    if (bytesPerSample == 2) {
                        int16_t s = *reinterpret_cast<const int16_t*>(&m_buffer[i * 2]);
                        sample = static_cast<FLAC__int32>(s);
                    } else if (bytesPerSample == 4) {
                        // Check if this is floating point audio
                        float f = *reinterpret_cast<const float*>(&m_buffer[i * 4]);

                        // Check if it looks like a float (typical range -1.0 to 1.0)
                        if (f >= -1.0f && f <= 1.0f) {
                            // Convert float [-1.0, 1.0] to 24-bit integer for FLAC
                            sample = static_cast<FLAC__int32>(f * 8388607.0f); // 2^23 - 1
                        } else {
                            // Assume it's 32-bit integer PCM
                            sample = *reinterpret_cast<const int32_t*>(&m_buffer[i * 4]);
                        }
                    }

                    flacBuffer[i] = sample;
                }

                std::vector<FLAC__int32*> channelBuffers(m_format.nChannels);
                std::vector<std::vector<FLAC__int32>> tempBuffers(m_format.nChannels);

                for (UINT32 ch = 0; ch < m_format.nChannels; ch++) {
                    tempBuffers[ch].resize(remainingSamples);
                    channelBuffers[ch] = tempBuffers[ch].data();

                    for (UINT32 i = 0; i < remainingSamples; i++) {
                        tempBuffers[ch][i] = flacBuffer[i * m_format.nChannels + ch];
                    }
                }

                FLAC__stream_encoder_process(m_encoder, channelBuffers.data(), remainingSamples);
            }
        }

        // Finish encoding
        FLAC__stream_encoder_finish(m_encoder);
        FLAC__stream_encoder_delete(m_encoder);
        m_encoder = nullptr;
    }

    if (m_file.is_open()) {
        m_file.close();
    }

    m_buffer.clear();
}
