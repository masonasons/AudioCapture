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
#include <algorithm>
#include <mutex>

// Define VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK if not already defined by SDK
#ifndef VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK
#define VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK L"VAD\\Process_Loopback"
#endif

// CRITICAL FIX: Global mutex to serialize WASAPI Initialize/Start operations
// Multiple sources calling IAudioClient::Initialize() simultaneously causes audio glitches
// This mutex ensures WASAPI operations happen one at a time to avoid driver conflicts
static std::mutex g_wasapiMutex;

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
    , m_isPaused(false)
    , m_targetProcessId(0)
    , m_volumeMultiplier(1.0f)  // Default to 100% volume
    , m_isProcessSpecific(false)
    , m_isInputDevice(false)
    , m_passthroughEnabled(false)
    , m_renderDevice(nullptr)
    , m_renderClient(nullptr)
    , m_audioRenderClient(nullptr)
    , m_renderBufferFrameCount(0)
{
    // COM should already be initialized by the main thread
}

AudioCapture::~AudioCapture() {
    Stop();
    DisablePassthrough();

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

    // Check if Windows version supports process loopback (Build 19041+)
    if (dwMajorVersion < 10 || (dwMajorVersion == 10 && dwBuild < 19041)) {
        wchar_t msg[512];
        swprintf_s(msg, L"Process capture requires Windows 10 Build 19041 or later.\n\n"
                       L"Your system: Windows %lu Build %lu\n\n"
                       L"Process capture will not be available. The application will fall back to system-wide audio capture.",
                       dwMajorVersion, dwBuild);
        MessageBox(nullptr, msg, L"Process Capture Not Supported", MB_OK | MB_ICONINFORMATION);
        return false;
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

    // Dynamically load ActivateAudioInterfaceAsync to support older Windows versions
    // This function only exists on Windows 8+ with WinRT
    typedef HRESULT (STDAPICALLTYPE *ActivateAudioInterfaceAsyncPtr)(
        LPCWSTR, REFIID, PROPVARIANT*, IActivateAudioInterfaceCompletionHandler*, IActivateAudioInterfaceAsyncOperation**);

    HMODULE hMmDevApi = LoadLibraryW(L"Mmdevapi.dll");
    if (!hMmDevApi) {
        handler->Release();
        return false;
    }

    ActivateAudioInterfaceAsyncPtr pActivateAudioInterfaceAsync =
        (ActivateAudioInterfaceAsyncPtr)GetProcAddress(hMmDevApi, "ActivateAudioInterfaceAsync");

    if (!pActivateAudioInterfaceAsync) {
        // Function not available (Windows 7 or earlier) - fail gracefully
        FreeLibrary(hMmDevApi);
        handler->Release();
        return false;
    }

    // Start async activation
    IActivateAudioInterfaceAsyncOperation* asyncOp = nullptr;
    HRESULT hr = pActivateAudioInterfaceAsync(
        deviceId,
        __uuidof(IAudioClient),
        &activateParams,
        handler,
        &asyncOp);

    FreeLibrary(hMmDevApi);

    if (FAILED(hr) || !asyncOp) {
        handler->Release();
        return false;
    }

    // Wait for activation to complete (blocks with timeout)
    bool success = handler->WaitForCompletion(5000);  // 5 second timeout

    // Release the async operation
    asyncOp->Release();

    if (!success) {
        // Get the activation result to show detailed error
        HRESULT activationHr = handler->GetActivationResult();
        handler->Release();

        // Show detailed error message to help user diagnose
        wchar_t msg[1024];
        const wchar_t* errorDesc = L"Unknown error";

        // Common error codes and what they mean
        if (activationHr == E_NOTIMPL || activationHr == 0x88890008) {  // AUDCLNT_E_UNSUPPORTED_FORMAT
            errorDesc = L"Process loopback audio not supported by your audio driver.\n\n"
                       L"This is usually caused by:\n"
                       L"- Outdated or incompatible audio drivers\n"
                       L"- Missing Windows updates\n"
                       L"- Audio driver not fully supporting Windows 10 loopback features\n\n"
                       L"Try updating your audio drivers and Windows.";
        } else if (activationHr == E_ACCESSDENIED) {
            errorDesc = L"Access denied to audio device.\n\n"
                       L"Try running the application as administrator.";
        } else if (activationHr == 0x887A0002) {  // DXGI_ERROR_NOT_FOUND
            errorDesc = L"Audio device not found or not available.";
        } else if (activationHr == 0x88890001) {  // AUDCLNT_E_NOT_INITIALIZED
            errorDesc = L"Audio client could not be initialized.\n\n"
                       L"The audio driver may not support this feature.";
        }

        swprintf_s(msg,
            L"Process capture failed for this application.\n\n"
            L"Windows Build: %lu (Requires 19041+)\n"
            L"Error Code: 0x%08X\n\n"
            L"%s\n\n"
            L"The application will fall back to system-wide audio capture.",
            dwBuild, activationHr, errorDesc);

        MessageBox(nullptr, msg, L"Process Capture Failed", MB_OK | MB_ICONWARNING);
        return false;
    }

    // Get the audio client (this AddRefs it for us)
    m_audioClient = handler->GetAudioClient();

    if (!m_audioClient) {
        HRESULT activationHr = handler->GetActivationResult();
        handler->Release();

        wchar_t msg[512];
        swprintf_s(msg,
            L"Failed to get audio client for process capture.\n\n"
            L"Error Code: 0x%08X\n\n"
            L"The application will fall back to system-wide audio capture.",
            activationHr);
        MessageBox(nullptr, msg, L"Process Capture Failed", MB_OK | MB_ICONWARNING);
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

        // Show detailed error message for E_UNEXPECTED on older Windows 10 builds
        // Build 19044 has a severe bug where process loopback is never released
        if (hr == E_UNEXPECTED && dwBuild <= 19044) {
            // Only show this message once per session to avoid spam
            static bool shownBugMessage = false;
            if (!shownBugMessage) {
                shownBugMessage = true;
                wchar_t msg[1536];
                swprintf_s(msg,
                    L"WINDOWS 10 BUILD %lu BUG DETECTED\n\n"
                    L"Process audio capture has failed because Windows 10 Build 19044 has a\n"
                    L"severe bug where process loopback audio interfaces are permanently locked\n"
                    L"after first use and cannot be reused.\n\n"
                    L"WORKAROUNDS:\n"
                    L"1. RESTART AudioCapture - Close and reopen this application\n"
                    L"2. RESTART target process - Close and reopen the app being captured (PID %lu)\n"
                    L"3. DON'T STOP - Once started, let capture run continuously\n"
                    L"4. UPGRADE WINDOWS - Update to Windows 11 or newer Windows 10 builds\n\n"
                    L"The application will now fall back to SYSTEM-WIDE audio capture,\n"
                    L"which captures ALL system audio, not just the target process.\n\n"
                    L"This is a known Microsoft bug in older Windows 10 builds and\n"
                    L"cannot be fixed in AudioCapture.",
                    dwBuild, processId);
                MessageBox(nullptr, msg, L"Critical Windows Bug - Process Audio Locked", MB_OK | MB_ICONERROR);
            }
        }

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

bool AudioCapture::InitializeFromDevice(const std::wstring& deviceId, bool isInputDevice) {
    m_isInputDevice = isInputDevice;
    m_isProcessSpecific = false;

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

    // Get the specified device by ID
    hr = m_deviceEnumerator->GetDevice(deviceId.c_str(), &m_device);
    if (FAILED(hr)) {
        return false;
    }

    // Activate audio client
    hr = m_device->Activate(__uuidof(IAudioClient), CLSCTX_ALL,
        nullptr, (void**)&m_audioClient);
    if (FAILED(hr)) {
        return false;
    }

    // Get mix format
    hr = m_audioClient->GetMixFormat(&m_waveFormat);
    if (FAILED(hr)) {
        return false;
    }

    // Initialize audio client
    const REFERENCE_TIME hnsRequestedDuration = 10000000; // 1 second
    DWORD streamFlags = 0;

    // For input devices (microphones), don't use loopback mode
    // For output devices, use loopback mode to capture what's playing
    if (!isInputDevice) {
        streamFlags = AUDCLNT_STREAMFLAGS_LOOPBACK;
    }

    hr = m_audioClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        streamFlags,
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

    // CRITICAL FIX: Serialize WASAPI Start() calls to avoid audio driver conflicts
    // When multiple sources start simultaneously, they compete for audio hardware resources
    // causing glitches in ALL active streams
    {
        std::lock_guard<std::mutex> lock(g_wasapiMutex);

        // Start audio client
        HRESULT hr = m_audioClient->Start();

        if (FAILED(hr)) {
            return false;
        }

        // Small delay to let WASAPI stabilize before starting capture thread
        Sleep(50);
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
    m_isPaused = false;

    // Wait for thread to finish
    if (m_captureThread.joinable()) {
        m_captureThread.join();
    }
}

void AudioCapture::Pause() {
    if (!m_isCapturing || m_isPaused) {
        return;
    }

    m_isPaused = true;

    // Pause the audio client (stops reading data but keeps the stream open)
    if (m_audioClient) {
        m_audioClient->Stop();
    }

    // Pause passthrough if enabled
    if (m_passthroughEnabled && m_renderClient) {
        m_renderClient->Stop();
    }
}

void AudioCapture::Resume() {
    if (!m_isCapturing || !m_isPaused) {
        return;
    }

    m_isPaused = false;

    // Resume the audio client
    if (m_audioClient) {
        m_audioClient->Start();
    }

    // Resume passthrough if enabled
    if (m_passthroughEnabled && m_renderClient) {
        m_renderClient->Start();
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
                    // For silent buffers, send zeros using pre-allocated buffer
                    if (m_silenceBuffer.size() < bufferSize) {
                        m_silenceBuffer.resize(bufferSize, 0);
                    }
                    // Fill with zeros (memset is faster than resize with initialization)
                    memset(m_silenceBuffer.data(), 0, bufferSize);
                    m_dataCallback(m_silenceBuffer.data(), bufferSize);
                }
                else if (data) {
                    // Use pre-allocated buffer to avoid allocation in audio callback
                    if (m_captureBuffer.size() < bufferSize) {
                        m_captureBuffer.resize(bufferSize);
                    }
                    // Copy data to our buffer
                    memcpy(m_captureBuffer.data(), data, bufferSize);
                    // Apply volume adjustment in-place
                    ApplyVolumeToBuffer(m_captureBuffer.data(), bufferSize);
                    m_dataCallback(m_captureBuffer.data(), bufferSize);

                    // If passthrough is enabled, also send to render device
                    if (m_passthroughEnabled && m_audioRenderClient && m_renderClient) {
                        // Get padding (how much is already in the buffer)
                        UINT32 numFramesPadding = 0;
                        if (SUCCEEDED(m_renderClient->GetCurrentPadding(&numFramesPadding))) {
                            // Calculate how many frames we can write
                            UINT32 renderFramesAvailable = m_renderBufferFrameCount - numFramesPadding;

                            // Calculate how many frames we have from capture
                            UINT32 captureFrames = numFramesAvailable;

                            // Only write as many frames as we have available, and don't exceed buffer space
                            UINT32 framesToWrite = std::min(renderFramesAvailable, captureFrames);

                            if (framesToWrite > 0) {
                                BYTE* renderBuffer = nullptr;
                                if (SUCCEEDED(m_audioRenderClient->GetBuffer(framesToWrite, &renderBuffer))) {
                                    // Copy audio data to render buffer from our pre-allocated buffer
                                    UINT32 bytesToCopy = framesToWrite * m_waveFormat->nBlockAlign;
                                    memcpy(renderBuffer, m_captureBuffer.data(), bytesToCopy);

                                    m_audioRenderClient->ReleaseBuffer(framesToWrite, 0);
                                }
                            }
                        }
                    }
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

bool AudioCapture::EnablePassthrough(const std::wstring& deviceId) {
    // Disable any existing passthrough
    DisablePassthrough();

    if (!m_waveFormat || !m_deviceEnumerator) {
        return false;
    }

    // Get the render device
    HRESULT hr = m_deviceEnumerator->GetDevice(deviceId.c_str(), &m_renderDevice);
    if (FAILED(hr) || !m_renderDevice) {
        return false;
    }

    // Activate audio client for rendering
    hr = m_renderDevice->Activate(__uuidof(IAudioClient), CLSCTX_ALL,
        nullptr, (void**)&m_renderClient);
    if (FAILED(hr)) {
        m_renderDevice->Release();
        m_renderDevice = nullptr;
        return false;
    }

    // Initialize the render client with the same format as capture
    // Use smaller buffer for lower latency (100ms instead of 1 second)
    const REFERENCE_TIME hnsRequestedDuration = 1000000; // 100ms
    hr = m_renderClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        0,
        hnsRequestedDuration,
        0,
        m_waveFormat,
        nullptr);

    if (FAILED(hr)) {
        m_renderClient->Release();
        m_renderClient = nullptr;
        m_renderDevice->Release();
        m_renderDevice = nullptr;
        return false;
    }

    // Get buffer size
    hr = m_renderClient->GetBufferSize(&m_renderBufferFrameCount);
    if (FAILED(hr)) {
        m_renderClient->Release();
        m_renderClient = nullptr;
        m_renderDevice->Release();
        m_renderDevice = nullptr;
        return false;
    }

    // Get render client service
    hr = m_renderClient->GetService(__uuidof(IAudioRenderClient),
        (void**)&m_audioRenderClient);

    if (FAILED(hr) || !m_audioRenderClient) {
        m_renderClient->Release();
        m_renderClient = nullptr;
        m_renderDevice->Release();
        m_renderDevice = nullptr;
        return false;
    }

    // Pre-fill half the buffer with silence to reduce latency while avoiding underruns
    UINT32 prefillFrames = m_renderBufferFrameCount / 2;
    BYTE* renderBuffer = nullptr;
    hr = m_audioRenderClient->GetBuffer(prefillFrames, &renderBuffer);
    if (SUCCEEDED(hr) && renderBuffer) {
        // Fill with silence
        memset(renderBuffer, 0, prefillFrames * m_waveFormat->nBlockAlign);
        m_audioRenderClient->ReleaseBuffer(prefillFrames, 0);
    }

    // Start the render client
    hr = m_renderClient->Start();
    if (FAILED(hr)) {
        m_audioRenderClient->Release();
        m_audioRenderClient = nullptr;
        m_renderClient->Release();
        m_renderClient = nullptr;
        m_renderDevice->Release();
        m_renderDevice = nullptr;
        return false;
    }

    m_passthroughEnabled = true;
    return true;
}

void AudioCapture::DisablePassthrough() {
    if (!m_passthroughEnabled) {
        return;
    }

    if (m_renderClient) {
        m_renderClient->Stop();
    }

    if (m_audioRenderClient) {
        m_audioRenderClient->Release();
        m_audioRenderClient = nullptr;
    }

    if (m_renderClient) {
        m_renderClient->Release();
        m_renderClient = nullptr;
    }

    if (m_renderDevice) {
        m_renderDevice->Release();
        m_renderDevice = nullptr;
    }

    m_renderBufferFrameCount = 0;
    m_passthroughEnabled = false;
}
