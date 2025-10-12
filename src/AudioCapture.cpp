#include "AudioCapture.h"
#include <avrt.h>
#include <functiondiscoverykeys_devpkey.h>
#include <audioclientactivationparams.h>
#include <cstring>
#include <propkey.h>
#include <propvarutil.h>
#include <mmreg.h>
#include <ks.h>
#include <ksmedia.h>

//=============================================================================
// AudioClientActivationHandler Implementation
//=============================================================================

AudioClientActivationHandler::AudioClientActivationHandler()
    : m_refCount(1)
    , m_completionEvent(nullptr)
    , m_audioClient(nullptr)
    , m_activationResult(E_FAIL)
{
    // Use manual-reset event (TRUE) like app2clap, not auto-reset
    m_completionEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
    if (!m_completionEvent) {
        // Event creation failed - set error state
        m_activationResult = E_OUTOFMEMORY;
    }
}

AudioClientActivationHandler::~AudioClientActivationHandler() {
    if (m_audioClient) {
        m_audioClient->Release();
    }
    if (m_completionEvent) {
        CloseHandle(m_completionEvent);
    }
}

STDMETHODIMP AudioClientActivationHandler::QueryInterface(REFIID riid, void** ppvObject) {
    if (!ppvObject) {
        return E_POINTER;
    }

    if (riid == __uuidof(IUnknown) ||
        riid == __uuidof(IActivateAudioInterfaceCompletionHandler)) {
        *ppvObject = static_cast<IActivateAudioInterfaceCompletionHandler*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IAgileObject)) {
        *ppvObject = static_cast<IAgileObject*>(this);
        AddRef();
        return S_OK;
    }

    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) AudioClientActivationHandler::AddRef() {
    return InterlockedIncrement(&m_refCount);
}

STDMETHODIMP_(ULONG) AudioClientActivationHandler::Release() {
    LONG refCount = InterlockedDecrement(&m_refCount);
    if (refCount == 0) {
        delete this;
    }
    return refCount;
}

STDMETHODIMP AudioClientActivationHandler::ActivateCompleted(IActivateAudioInterfaceAsyncOperation* operation) {
    if (!operation) {
        m_activationResult = E_INVALIDARG;
        SetEvent(m_completionEvent);
        return E_INVALIDARG;
    }

    // Get the activation result
    IUnknown* audioInterface = nullptr;
    HRESULT hr = operation->GetActivateResult(&m_activationResult, &audioInterface);

    if (SUCCEEDED(hr) && SUCCEEDED(m_activationResult) && audioInterface) {
        // Get the IAudioClient interface
        hr = audioInterface->QueryInterface(__uuidof(IAudioClient), (void**)&m_audioClient);
        audioInterface->Release();

        if (FAILED(hr)) {
            m_activationResult = hr;
        }
    }

    // Signal completion
    SetEvent(m_completionEvent);
    return S_OK;
}

bool AudioClientActivationHandler::WaitForCompletion(DWORD timeoutMs) {
    if (!m_completionEvent) {
        return false;
    }

    // Use CoWaitForMultipleHandles to pump messages while waiting (required for STA)
    HANDLE handles[1] = { m_completionEvent };
    DWORD index = 0;
    HRESULT hr = CoWaitForMultipleHandles(
        COWAIT_DISPATCH_CALLS | COWAIT_DISPATCH_WINDOW_MESSAGES,
        timeoutMs,
        1,
        handles,
        &index);

    return (hr == S_OK) && SUCCEEDED(m_activationResult);
}

IAudioClient* AudioClientActivationHandler::GetAudioClient() {
    if (m_audioClient) {
        m_audioClient->AddRef();
    }
    return m_audioClient;
}

//=============================================================================
// AudioCapture Implementation
//=============================================================================

AudioCapture::AudioCapture()
    : m_deviceEnumerator(nullptr)
    , m_device(nullptr)
    , m_audioClient(nullptr)
    , m_captureClient(nullptr)
    , m_waveFormat(nullptr)
    , m_isCapturing(false)
    , m_targetProcessId(0)
    , m_volumeMultiplier(0.5f)  // Default to 50% volume
    , m_isProcessSpecific(false)
{
    // COM should already be initialized by the main thread
}

AudioCapture::~AudioCapture() {
    Stop();

    if (m_waveFormat) {
        CoTaskMemFree(m_waveFormat);
    }
    if (m_captureClient) {
        m_captureClient->Release();
    }
    if (m_audioClient) {
        m_audioClient->Release();
    }
    if (m_device) {
        m_device->Release();
    }
    if (m_deviceEnumerator) {
        m_deviceEnumerator->Release();
    }

    // Don't call CoUninitialize - main thread will handle it
}

bool AudioCapture::Initialize(DWORD processId) {
    m_targetProcessId = processId;

    // Clean up any existing state from previous initialization
    if (m_captureClient) {
        m_captureClient->Release();
        m_captureClient = nullptr;
    }
    if (m_audioClient) {
        m_audioClient->Release();
        m_audioClient = nullptr;
    }
    if (m_waveFormat) {
        CoTaskMemFree(m_waveFormat);
        m_waveFormat = nullptr;
    }
    if (m_device) {
        m_device->Release();
        m_device = nullptr;
    }
    if (m_deviceEnumerator) {
        m_deviceEnumerator->Release();
        m_deviceEnumerator = nullptr;
    }

    // Create device enumerator
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr,
        CLSCTX_ALL, __uuidof(IMMDeviceEnumerator),
        (void**)&m_deviceEnumerator);

    if (FAILED(hr)) {
        return false;
    }

    // Get default audio endpoint
    hr = m_deviceEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &m_device);

    if (FAILED(hr)) {
        return false;
    }

    // Try process-specific capture for non-zero process IDs
    if (processId != 0) {
        if (InitializeProcessSpecificCapture(processId)) {
            m_isProcessSpecific = true;
            return true;
        }
    }

    // Fall back to system-wide capture
    m_isProcessSpecific = false;
    return InitializeSystemWideCapture();
}

bool AudioCapture::InitializeProcessSpecificCapture(DWORD processId) {
    // Check Windows version using RtlGetVersion
    typedef LONG (WINAPI* RtlGetVersionPtr)(PRTL_OSVERSIONINFOW);
    HMODULE hNtdll = GetModuleHandleW(L"ntdll.dll");
    RtlGetVersionPtr RtlGetVersion = nullptr;

    DWORD dwMajorVersion = 0;
    DWORD dwBuild = 0;

    if (hNtdll) {
        RtlGetVersion = (RtlGetVersionPtr)GetProcAddress(hNtdll, "RtlGetVersion");
        if (RtlGetVersion) {
            RTL_OSVERSIONINFOW osvi = { 0 };
            osvi.dwOSVersionInfoSize = sizeof(osvi);
            if (RtlGetVersion(&osvi) == 0) {
                dwMajorVersion = osvi.dwMajorVersion;
                dwBuild = osvi.dwBuildNumber;
            }
        }
    }

    // Use the virtual device for process loopback (not a physical device)
    LPCWSTR deviceId = VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK;

    // Setup activation parameters for process loopback
    AUDIOCLIENT_ACTIVATION_PARAMS activationParams = {};
    activationParams.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;
    activationParams.ProcessLoopbackParams.ProcessLoopbackMode = PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE;
    activationParams.ProcessLoopbackParams.TargetProcessId = processId;

    PROPVARIANT activateParams = {};
    activateParams.vt = VT_BLOB;
    activateParams.blob.cbSize = sizeof(activationParams);
    activateParams.blob.pBlobData = (BYTE*)&activationParams;

    // Create completion handler (starts with refcount = 1)
    AudioClientActivationHandler* handler = new (std::nothrow) AudioClientActivationHandler();
    if (!handler || !handler->IsValid()) {
        if (handler) {
            handler->Release();
        }
        return false;
    }

    // Start async activation
    IActivateAudioInterfaceAsyncOperation* asyncOp = nullptr;
    HRESULT hr = ActivateAudioInterfaceAsync(
        deviceId,
        __uuidof(IAudioClient),
        &activateParams,
        handler,
        &asyncOp);

    if (FAILED(hr) || !asyncOp) {
        handler->Release();
        return false;
    }

    // Wait for activation to complete (blocks with timeout)
    bool success = handler->WaitForCompletion(5000);  // 5 second timeout

    // Release the async operation
    asyncOp->Release();

    if (!success) {
        handler->Release();
        return false;
    }

    // Get the audio client (this AddRefs it for us)
    m_audioClient = handler->GetAudioClient();

    if (!m_audioClient) {
        handler->Release();
        return false;
    }

    // Release ownership so handler doesn't Release our audio client in its destructor
    handler->ReleaseOwnership();

    // Release our reference to the handler (may delete it if refcount reaches 0)
    handler->Release();

    // Get mix format - for process loopback, GetMixFormat returns E_NOTIMPL
    // so we need to use the default device's format as a template
    hr = m_audioClient->GetMixFormat(&m_waveFormat);

    if (FAILED(hr)) {
        // GetMixFormat not supported on process loopback device
        // Get format from default device and use it as template
        if (m_device) {
            IAudioClient* tempClient = nullptr;
            hr = m_device->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, (void**)&tempClient);

            if (SUCCEEDED(hr) && tempClient) {
                WAVEFORMATEX* defaultFormat = nullptr;
                hr = tempClient->GetMixFormat(&defaultFormat);

                if (SUCCEEDED(hr) && defaultFormat) {
                    // Ask the process loopback client what format it will actually use
                    WAVEFORMATEX* closestMatch = nullptr;
                    hr = m_audioClient->IsFormatSupported(AUDCLNT_SHAREMODE_SHARED, defaultFormat, &closestMatch);

                    if (hr == S_OK && defaultFormat) {
                        // Exact match - use the default format
                        m_waveFormat = defaultFormat;
                        defaultFormat = nullptr; // Don't free it
                    } else if (hr == S_FALSE && closestMatch) {
                        // Close match provided
                        m_waveFormat = closestMatch;
                        closestMatch = nullptr; // Don't free it
                    } else {
                        // No match - use default anyway
                        m_waveFormat = defaultFormat;
                        defaultFormat = nullptr;
                    }

                    if (closestMatch) CoTaskMemFree(closestMatch);
                    if (defaultFormat) CoTaskMemFree(defaultFormat);
                }
                tempClient->Release();
            }
        }

        if (!m_waveFormat) {
            m_audioClient->Release();
            m_audioClient = nullptr;
            return false;
        }
    }

    // Validate before Initialize
    if (!m_audioClient || !m_waveFormat) {
        if (m_audioClient) {
            m_audioClient->Release();
            m_audioClient = nullptr;
        }
        return false;
    }

    // Initialize audio client for loopback capture
    const REFERENCE_TIME hnsRequestedDuration = 10000000; // 1 second
    hr = m_audioClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM | AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY,
        hnsRequestedDuration,
        0,
        m_waveFormat,
        nullptr);

    if (FAILED(hr)) {
        CoTaskMemFree(m_waveFormat);
        m_waveFormat = nullptr;
        m_audioClient->Release();
        m_audioClient = nullptr;
        return false;
    }

    // Get capture client
    hr = m_audioClient->GetService(__uuidof(IAudioCaptureClient),
        (void**)&m_captureClient);

    if (FAILED(hr) || !m_captureClient) {
        CoTaskMemFree(m_waveFormat);
        m_waveFormat = nullptr;
        m_audioClient->Release();
        m_audioClient = nullptr;
        return false;
    }

    return true;
}

bool AudioCapture::InitializeSystemWideCapture() {
    // Activate audio client (legacy method - captures all system audio)
    HRESULT hr = m_device->Activate(__uuidof(IAudioClient), CLSCTX_ALL,
        nullptr, (void**)&m_audioClient);
    if (FAILED(hr)) {
        return false;
    }

    // Get mix format
    hr = m_audioClient->GetMixFormat(&m_waveFormat);
    if (FAILED(hr)) {
        return false;
    }

    // Initialize audio client for loopback capture
    const REFERENCE_TIME hnsRequestedDuration = 10000000; // 1 second
    hr = m_audioClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        AUDCLNT_STREAMFLAGS_LOOPBACK,
        hnsRequestedDuration,
        0,
        m_waveFormat,
        nullptr);
    if (FAILED(hr)) {
        return false;
    }

    // Get capture client
    hr = m_audioClient->GetService(__uuidof(IAudioCaptureClient),
        (void**)&m_captureClient);
    if (FAILED(hr)) {
        return false;
    }

    return true;
}

bool AudioCapture::Start() {
    if (m_isCapturing) {
        return false;
    }

    if (!m_audioClient || !m_captureClient) {
        return false;
    }

    // Start audio client
    HRESULT hr = m_audioClient->Start();
    if (FAILED(hr)) {
        return false;
    }

    m_isCapturing = true;
    m_captureThread = std::thread(&AudioCapture::CaptureThread, this);

    return true;
}

void AudioCapture::Stop() {
    if (!m_isCapturing) {
        return;
    }

    // Stop the audio client FIRST, before stopping the thread
    if (m_audioClient) {
        m_audioClient->Stop();
    }

    // Then signal the thread to stop
    m_isCapturing = false;

    // Wait for thread to finish
    if (m_captureThread.joinable()) {
        m_captureThread.join();
    }
}

void AudioCapture::ApplyVolumeToBuffer(BYTE* data, UINT32 size) {
    if (!data || !m_waveFormat || size == 0) {
        return;
    }

    if (m_volumeMultiplier >= 1.0f) {
        return; // No adjustment needed
    }

    // Check if this is WAVEFORMATEXTENSIBLE
    bool isExtensible = (m_waveFormat->wFormatTag == WAVE_FORMAT_EXTENSIBLE &&
                         m_waveFormat->cbSize >= 22);

    // Determine the actual format
    bool isFloat = false;
    if (isExtensible) {
        WAVEFORMATEXTENSIBLE* wfex = reinterpret_cast<WAVEFORMATEXTENSIBLE*>(m_waveFormat);
        isFloat = (wfex->SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT);
    } else {
        isFloat = (m_waveFormat->wFormatTag == WAVE_FORMAT_IEEE_FLOAT);
    }

    if (isFloat && m_waveFormat->wBitsPerSample == 32) {
        // 32-bit float PCM
        float* samples = reinterpret_cast<float*>(data);
        UINT32 numSamples = size / sizeof(float);
        for (UINT32 i = 0; i < numSamples; i++) {
            samples[i] = samples[i] * m_volumeMultiplier;
        }
    }
    else if (m_waveFormat->wBitsPerSample == 16) {
        // 16-bit PCM
        int16_t* samples = reinterpret_cast<int16_t*>(data);
        UINT32 numSamples = size / sizeof(int16_t);
        for (UINT32 i = 0; i < numSamples; i++) {
            samples[i] = static_cast<int16_t>(samples[i] * m_volumeMultiplier);
        }
    }
}

void AudioCapture::CaptureThread() {
    // Validate required members
    if (!m_captureClient || !m_waveFormat) {
        return;
    }

    // Set thread priority for audio
    DWORD taskIndex = 0;
    HANDLE hTask = AvSetMmThreadCharacteristics(L"Audio", &taskIndex);

    // Poll every 10ms
    const DWORD sleepMs = 10;

    while (m_isCapturing) {
        Sleep(sleepMs);

        // Check if we should stop
        if (!m_isCapturing) {
            break;
        }

        UINT32 packetLength = 0;
        HRESULT hr = m_captureClient->GetNextPacketSize(&packetLength);
        if (FAILED(hr)) {
            break;
        }

        while (packetLength > 0 && m_isCapturing) {
            BYTE* data = nullptr;
            UINT32 numFramesAvailable = 0;
            DWORD flags = 0;

            hr = m_captureClient->GetBuffer(&data, &numFramesAvailable, &flags, nullptr, nullptr);
            if (FAILED(hr)) {
                break;
            }

            // Calculate buffer size
            UINT32 bufferSize = numFramesAvailable * m_waveFormat->nBlockAlign;

            // Send data to callback - even if silent, send zeros to keep stream continuous
            if (m_dataCallback && bufferSize > 0) {
                if (flags & AUDCLNT_BUFFERFLAGS_SILENT) {
                    // For silent buffers, send zeros
                    std::vector<BYTE> silence(bufferSize, 0);
                    m_dataCallback(silence.data(), bufferSize);
                }
                else if (data) {
                    // Apply volume adjustment
                    std::vector<BYTE> buffer(data, data + bufferSize);
                    ApplyVolumeToBuffer(buffer.data(), bufferSize);
                    m_dataCallback(buffer.data(), bufferSize);
                }
            }

            hr = m_captureClient->ReleaseBuffer(numFramesAvailable);
            if (FAILED(hr)) {
                break;
            }

            // Check if we should stop before getting next packet
            if (!m_isCapturing) {
                break;
            }

            hr = m_captureClient->GetNextPacketSize(&packetLength);
            if (FAILED(hr)) {
                break;
            }
        }
    }

    if (hTask) {
        AvRevertMmThreadCharacteristics(hTask);
    }
}
