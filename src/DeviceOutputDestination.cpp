#include "DeviceOutputDestination.h"
#include <ks.h>
#include <ksmedia.h>
#include <cstring>
#include <algorithm>

DeviceOutputDestination::DeviceOutputDestination()
    : m_deviceEnumerator(nullptr)
    , m_device(nullptr)
    , m_audioClient(nullptr)
    , m_renderClient(nullptr)
    , m_format(nullptr)
    , m_bufferFrameCount(0)
    , m_volumeMultiplier(1.0f)
{
}

DeviceOutputDestination::~DeviceOutputDestination() {
    Close();
}

float DeviceOutputDestination::ValidateVolume(float volume) {
    // Clamp to reasonable range (0.0 to 2.0)
    // 0.0 = mute, 1.0 = normal, 2.0 = double amplitude
    if (volume < 0.0f) {
        return 0.0f;
    }
    if (volume > 2.0f) {
        return 2.0f;
    }
    return volume;
}

void DeviceOutputDestination::SetVolumeMultiplier(float volume) {
    m_volumeMultiplier = ValidateVolume(volume);
}

void DeviceOutputDestination::ApplyVolumeToBuffer(BYTE* data, UINT32 size) {
    if (!data || !m_format || size == 0) {
        return;
    }

    if (m_volumeMultiplier >= 0.99f && m_volumeMultiplier <= 1.01f) {
        return; // No adjustment needed (within 1% of unity)
    }

    // Check if this is WAVEFORMATEXTENSIBLE
    bool isExtensible = (m_format->wFormatTag == WAVE_FORMAT_EXTENSIBLE &&
                         m_format->cbSize >= 22);

    // Determine the actual format
    bool isFloat = false;
    if (isExtensible) {
        WAVEFORMATEXTENSIBLE* wfex = reinterpret_cast<WAVEFORMATEXTENSIBLE*>(m_format);
        isFloat = (wfex->SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT);
    } else {
        isFloat = (m_format->wFormatTag == WAVE_FORMAT_IEEE_FLOAT);
    }

    if (isFloat && m_format->wBitsPerSample == 32) {
        // 32-bit float PCM
        float* samples = reinterpret_cast<float*>(data);
        UINT32 numSamples = size / sizeof(float);
        for (UINT32 i = 0; i < numSamples; i++) {
            samples[i] = samples[i] * m_volumeMultiplier;
            // Optionally clamp to prevent clipping
            if (samples[i] > 1.0f) samples[i] = 1.0f;
            if (samples[i] < -1.0f) samples[i] = -1.0f;
        }
    }
    else if (m_format->wBitsPerSample == 16) {
        // 16-bit PCM
        int16_t* samples = reinterpret_cast<int16_t*>(data);
        UINT32 numSamples = size / sizeof(int16_t);
        for (UINT32 i = 0; i < numSamples; i++) {
            int32_t adjusted = static_cast<int32_t>(samples[i] * m_volumeMultiplier);
            // Clamp to prevent overflow
            if (adjusted > 32767) adjusted = 32767;
            if (adjusted < -32768) adjusted = -32768;
            samples[i] = static_cast<int16_t>(adjusted);
        }
    }
}

bool DeviceOutputDestination::Configure(const WAVEFORMATEX* format, const DestinationConfig& config) {
    // Clear any previous errors
    m_lastError.clear();

    // Validate inputs
    if (!format) {
        m_lastError = L"Audio format is null";
        return false;
    }

    if (config.outputPath.empty()) {
        m_lastError = L"Device ID cannot be empty";
        return false;
    }

    // Close any previously opened device
    if (IsOpen()) {
        Close();
    }

    // Store configuration
    m_deviceId = config.outputPath;  // outputPath contains device ID for devices
    m_friendlyName = config.friendlyName;  // Store friendly name for identification
    m_volumeMultiplier = ValidateVolume(config.volumeMultiplier);

    // Create device enumerator
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr,
        CLSCTX_ALL, __uuidof(IMMDeviceEnumerator),
        (void**)&m_deviceEnumerator);

    if (FAILED(hr)) {
        m_lastError = L"Failed to create device enumerator";
        return false;
    }

    // Get the specified device
    hr = m_deviceEnumerator->GetDevice(m_deviceId.c_str(), &m_device);
    if (FAILED(hr)) {
        m_lastError = L"Failed to get audio device: " + m_deviceId;
        m_deviceEnumerator->Release();
        m_deviceEnumerator = nullptr;
        return false;
    }

    // Activate audio client
    hr = m_device->Activate(__uuidof(IAudioClient), CLSCTX_ALL,
        nullptr, (void**)&m_audioClient);
    if (FAILED(hr)) {
        m_lastError = L"Failed to activate audio client";
        m_device->Release();
        m_device = nullptr;
        m_deviceEnumerator->Release();
        m_deviceEnumerator = nullptr;
        return false;
    }

    // Copy the format
    UINT32 formatSize = GetFormatSize(format);
    m_format = reinterpret_cast<WAVEFORMATEX*>(CoTaskMemAlloc(formatSize));
    if (!m_format) {
        m_lastError = L"Failed to allocate format structure";
        Close();
        return false;
    }
    memcpy(m_format, format, formatSize);

    // Initialize audio client for rendering
    // Use smaller buffer for lower latency (100ms)
    const REFERENCE_TIME hnsRequestedDuration = 1000000; // 100ms
    hr = m_audioClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        0,  // No special flags for rendering
        hnsRequestedDuration,
        0,
        m_format,
        nullptr);

    if (FAILED(hr)) {
        m_lastError = L"Failed to initialize audio client";
        Close();
        return false;
    }

    // Get buffer size
    hr = m_audioClient->GetBufferSize(&m_bufferFrameCount);
    if (FAILED(hr)) {
        m_lastError = L"Failed to get buffer size";
        Close();
        return false;
    }

    // Get render client
    hr = m_audioClient->GetService(__uuidof(IAudioRenderClient),
        (void**)&m_renderClient);

    if (FAILED(hr)) {
        m_lastError = L"Failed to get render client";
        Close();
        return false;
    }

    // Pre-fill half the buffer with silence to reduce latency while avoiding underruns
    UINT32 prefillFrames = m_bufferFrameCount / 2;
    BYTE* renderBuffer = nullptr;
    hr = m_renderClient->GetBuffer(prefillFrames, &renderBuffer);
    if (SUCCEEDED(hr) && renderBuffer) {
        // Fill with silence
        memset(renderBuffer, 0, prefillFrames * m_format->nBlockAlign);
        m_renderClient->ReleaseBuffer(prefillFrames, 0);
    }

    // Start the audio client
    hr = m_audioClient->Start();
    if (FAILED(hr)) {
        m_lastError = L"Failed to start audio client";
        Close();
        return false;
    }

    // Start the async writer thread
    StartAsyncWriter();

    return true;
}

bool DeviceOutputDestination::WriteAudioDataInternal(const BYTE* data, UINT32 size) {
    if (!IsOpen()) {
        m_lastError = L"Cannot write - audio device is not open";
        return false;
    }

    if (!data || size == 0) {
        // Silently ignore empty writes
        return true;
    }

    // Calculate number of frames to write
    UINT32 framesToWrite = size / m_format->nBlockAlign;
    if (framesToWrite == 0) {
        return true;  // Not enough data for even one frame
    }

    // Get current padding (how much is already in the buffer)
    UINT32 numFramesPadding = 0;
    HRESULT hr = m_audioClient->GetCurrentPadding(&numFramesPadding);
    if (FAILED(hr)) {
        m_lastError = L"Failed to get current padding";
        return false;
    }

    // Calculate how many frames we can write
    UINT32 framesAvailable = m_bufferFrameCount - numFramesPadding;

    // Only write as many frames as we have available space for
    UINT32 framesToWriteNow = std::min(framesAvailable, framesToWrite);

    if (framesToWriteNow == 0) {
        // Buffer is full - could drop data or block
        // For real-time monitoring, we'll just drop the data
        return true;
    }

    // Get buffer from render client
    BYTE* renderBuffer = nullptr;
    hr = m_renderClient->GetBuffer(framesToWriteNow, &renderBuffer);
    if (FAILED(hr)) {
        m_lastError = L"Failed to get render buffer";
        return false;
    }

    // Copy data to render buffer with volume adjustment
    UINT32 bytesToCopy = framesToWriteNow * m_format->nBlockAlign;
    memcpy(renderBuffer, data, bytesToCopy);

    // Apply volume adjustment in-place
    ApplyVolumeToBuffer(renderBuffer, bytesToCopy);

    // Release the buffer
    hr = m_renderClient->ReleaseBuffer(framesToWriteNow, 0);
    if (FAILED(hr)) {
        m_lastError = L"Failed to release render buffer";
        return false;
    }

    return true;
}

void DeviceOutputDestination::Close() {
    // Stop the async writer thread (flushes pending writes)
    StopAsyncWriter();

    // Stop the audio client
    if (m_audioClient) {
        m_audioClient->Stop();
    }

    // Release all COM interfaces
    if (m_renderClient) {
        m_renderClient->Release();
        m_renderClient = nullptr;
    }

    if (m_audioClient) {
        m_audioClient->Release();
        m_audioClient = nullptr;
    }

    if (m_device) {
        m_device->Release();
        m_device = nullptr;
    }

    if (m_deviceEnumerator) {
        m_deviceEnumerator->Release();
        m_deviceEnumerator = nullptr;
    }

    if (m_format) {
        CoTaskMemFree(m_format);
        m_format = nullptr;
    }

    m_bufferFrameCount = 0;
    m_deviceId.clear();
}

bool DeviceOutputDestination::IsOpen() const {
    return m_audioClient != nullptr && m_renderClient != nullptr;
}
