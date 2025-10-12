#include "Mp3Encoder.h"
#include <mftransform.h>
#include <wmcodecdsp.h>
#include <cstring>
#include <ks.h>
#include <ksmedia.h>

Mp3Encoder::Mp3Encoder()
    : m_sinkWriter(nullptr)
    , m_streamIndex(0)
    , m_sampleDuration(0)
    , m_rtStart(0)
    , m_samplesPerFrame(0)
{
    std::memset(&m_inputFormat, 0, sizeof(WAVEFORMATEX));
    MFStartup(MF_VERSION);
}

Mp3Encoder::~Mp3Encoder() {
    Close();
    MFShutdown();
}

bool Mp3Encoder::Open(const std::wstring& filename, const WAVEFORMATEX* format, UINT32 bitrate) {
    if (m_sinkWriter != nullptr) {
        return false;
    }

    m_inputFormat = *format;

    // Create sink writer
    HRESULT hr = MFCreateSinkWriterFromURL(filename.c_str(), nullptr, nullptr, &m_sinkWriter);
    if (FAILED(hr)) {
        return false;
    }

    // Configure output media type (MP3)
    IMFMediaType* pOutputType = nullptr;
    hr = MFCreateMediaType(&pOutputType);
    if (FAILED(hr)) {
        return false;
    }

    pOutputType->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Audio);
    pOutputType->SetGUID(MF_MT_SUBTYPE, MFAudioFormat_MP3);
    pOutputType->SetUINT32(MF_MT_AUDIO_NUM_CHANNELS, format->nChannels);
    pOutputType->SetUINT32(MF_MT_AUDIO_SAMPLES_PER_SECOND, format->nSamplesPerSec);
    pOutputType->SetUINT32(MF_MT_AUDIO_AVG_BYTES_PER_SECOND, bitrate / 8);

    hr = m_sinkWriter->AddStream(pOutputType, &m_streamIndex);
    pOutputType->Release();
    if (FAILED(hr)) {
        m_sinkWriter->Release();
        m_sinkWriter = nullptr;
        return false;
    }

    // Configure input media type (PCM or Float)
    IMFMediaType* pInputType = nullptr;
    hr = MFCreateMediaType(&pInputType);
    if (FAILED(hr)) {
        m_sinkWriter->Release();
        m_sinkWriter = nullptr;
        return false;
    }

    // Determine if this is float or PCM
    GUID subtype = MFAudioFormat_PCM;
    if (format->wFormatTag == WAVE_FORMAT_IEEE_FLOAT) {
        subtype = MFAudioFormat_Float;
    }
    else if (format->wFormatTag == WAVE_FORMAT_EXTENSIBLE && format->cbSize >= 22) {
        // Check the SubFormat GUID in WAVEFORMATEXTENSIBLE
        const WAVEFORMATEXTENSIBLE* wfex = reinterpret_cast<const WAVEFORMATEXTENSIBLE*>(format);
        if (wfex->SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT) {
            subtype = MFAudioFormat_Float;
        }
    }

    pInputType->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Audio);
    pInputType->SetGUID(MF_MT_SUBTYPE, subtype);
    pInputType->SetUINT32(MF_MT_AUDIO_NUM_CHANNELS, format->nChannels);
    pInputType->SetUINT32(MF_MT_AUDIO_SAMPLES_PER_SECOND, format->nSamplesPerSec);
    pInputType->SetUINT32(MF_MT_AUDIO_BITS_PER_SAMPLE, format->wBitsPerSample);
    pInputType->SetUINT32(MF_MT_AUDIO_BLOCK_ALIGNMENT, format->nBlockAlign);
    pInputType->SetUINT32(MF_MT_AUDIO_AVG_BYTES_PER_SECOND, format->nAvgBytesPerSec);

    hr = m_sinkWriter->SetInputMediaType(m_streamIndex, pInputType, nullptr);
    pInputType->Release();
    if (FAILED(hr)) {
        m_sinkWriter->Release();
        m_sinkWriter = nullptr;
        return false;
    }

    // Begin writing
    hr = m_sinkWriter->BeginWriting();
    if (FAILED(hr)) {
        m_sinkWriter->Release();
        m_sinkWriter = nullptr;
        return false;
    }

    // Calculate sample duration
    m_sampleDuration = 10000000LL * 1152 / format->nSamplesPerSec; // MP3 frame = 1152 samples
    m_samplesPerFrame = 1152;
    m_rtStart = 0;

    return true;
}

bool Mp3Encoder::WriteData(const BYTE* data, UINT32 size) {
    if (!m_sinkWriter) {
        return false;
    }

    // Add data to buffer
    m_buffer.insert(m_buffer.end(), data, data + size);

    // Process complete frames
    UINT32 frameSize = m_samplesPerFrame * m_inputFormat.nBlockAlign;
    while (m_buffer.size() >= frameSize) {
        // Create media buffer
        IMFMediaBuffer* pBuffer = nullptr;
        HRESULT hr = MFCreateMemoryBuffer(frameSize, &pBuffer);
        if (FAILED(hr)) {
            return false;
        }

        // Copy data to buffer
        BYTE* pData = nullptr;
        hr = pBuffer->Lock(&pData, nullptr, nullptr);
        if (SUCCEEDED(hr)) {
            std::memcpy(pData, m_buffer.data(), frameSize);
            pBuffer->Unlock();
            pBuffer->SetCurrentLength(frameSize);
        }

        // Create sample
        IMFSample* pSample = nullptr;
        hr = MFCreateSample(&pSample);
        if (SUCCEEDED(hr)) {
            pSample->AddBuffer(pBuffer);
            pSample->SetSampleTime(m_rtStart);
            pSample->SetSampleDuration(m_sampleDuration);

            // Write sample
            hr = m_sinkWriter->WriteSample(m_streamIndex, pSample);
            pSample->Release();

            m_rtStart += m_sampleDuration;
        }

        pBuffer->Release();

        // Remove processed data from buffer
        m_buffer.erase(m_buffer.begin(), m_buffer.begin() + frameSize);
    }

    return true;
}

void Mp3Encoder::Close() {
    if (!m_sinkWriter) {
        return;
    }

    // Finalize
    m_sinkWriter->Finalize();
    m_sinkWriter->Release();
    m_sinkWriter = nullptr;

    m_buffer.clear();
}
