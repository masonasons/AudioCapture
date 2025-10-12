#pragma once

#include <windows.h>
#include <mmdeviceapi.h>
#include <audioclient.h>
#include <audiopolicy.h>
#include <thread>
#include <atomic>
#include <functional>
#include <vector>
#include <string>

// For process-specific audio capture (Windows 10 Build 20348+)
#include <audioclientactivationparams.h>

// Forward declaration
class AudioClientActivationHandler;

class AudioCapture {
public:
    AudioCapture();
    ~AudioCapture();

    // Initialize capture for a specific process (0 for system-wide)
    bool Initialize(DWORD processId);

    // Initialize capture from a specific device (for microphone/line-in)
    bool InitializeFromDevice(const std::wstring& deviceId, bool isInputDevice);

    // Start capturing audio
    bool Start();

    // Stop capturing audio
    void Stop();

    // Check if currently capturing
    bool IsCapturing() const { return m_isCapturing; }

    // Get audio format information
    WAVEFORMATEX* GetFormat() const { return m_waveFormat; }

    // Set callback for audio data (called when new audio data is available)
    void SetDataCallback(std::function<void(const BYTE*, UINT32)> callback) {
        m_dataCallback = callback;
    }

    // Set volume multiplier (0.0 to 1.0)
    void SetVolume(float volume) { m_volumeMultiplier = volume; }

    // Enable/disable audio passthrough to a render device
    bool EnablePassthrough(const std::wstring& deviceId);
    void DisablePassthrough();
    bool IsPassthroughEnabled() const { return m_passthroughEnabled; }

private:
    void CaptureThread();
    bool InitializeProcessSpecificCapture(DWORD processId);
    bool InitializeSystemWideCapture();
    void ApplyVolumeToBuffer(BYTE* data, UINT32 size);

    IMMDeviceEnumerator* m_deviceEnumerator;
    IMMDevice* m_device;
    IAudioClient* m_audioClient;
    IAudioCaptureClient* m_captureClient;
    WAVEFORMATEX* m_waveFormat;

    std::atomic<bool> m_isCapturing;
    std::thread m_captureThread;
    std::function<void(const BYTE*, UINT32)> m_dataCallback;

    DWORD m_targetProcessId;
    float m_volumeMultiplier;
    bool m_isProcessSpecific;
    bool m_isInputDevice;  // True if capturing from input device (mic), false if loopback

    // Passthrough/monitor members
    bool m_passthroughEnabled;
    IMMDevice* m_renderDevice;
    IAudioClient* m_renderClient;
    IAudioRenderClient* m_audioRenderClient;
    UINT32 m_renderBufferFrameCount;
};

// COM completion handler for async audio interface activation
// Must implement IAgileObject for free-threaded marshaling to avoid E_ILLEGAL_METHOD_CALL
class AudioClientActivationHandler : public IActivateAudioInterfaceCompletionHandler, public IAgileObject {
public:
    AudioClientActivationHandler();
    virtual ~AudioClientActivationHandler();

    // IUnknown methods
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;

    // IActivateAudioInterfaceCompletionHandler method
    STDMETHODIMP ActivateCompleted(IActivateAudioInterfaceAsyncOperation* operation) override;

    // Wait for activation to complete
    bool WaitForCompletion(DWORD timeoutMs = 5000);

    // Get the activated audio client
    IAudioClient* GetAudioClient();

    // Release ownership of the audio client (so destructor won't Release it)
    void ReleaseOwnership() { m_audioClient = nullptr; }

    // Check if the handler is valid (event created successfully)
    bool IsValid() const { return m_completionEvent != nullptr; }

private:
    LONG m_refCount;
    HANDLE m_completionEvent;
    IAudioClient* m_audioClient;
    HRESULT m_activationResult;
};
