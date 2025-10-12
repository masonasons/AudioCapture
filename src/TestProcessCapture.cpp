#include <windows.h>
#include <mmdeviceapi.h>
#include <audioclient.h>
#include <audioclientactivationparams.h>
#include <roapi.h>
#include <iostream>
#include <fstream>
#include <string>
#include <comdef.h>

#pragma comment(lib, "RuntimeObject.lib")

std::ofstream g_logFile;

void Log(const std::string& message) {
    std::cout << message << std::endl;
    if (g_logFile.is_open()) {
        g_logFile << message << std::endl;
        g_logFile.flush();
    }
}

int main(int argc, char* argv[]) {
    // Open log file
    g_logFile.open("test_output.txt", std::ios::out | std::ios::trunc);

    if (argc < 2) {
        Log("Usage: TestProcessCapture.exe <process_id>");
        return 1;
    }

    DWORD processId = (DWORD)atoi(argv[1]);
    char buf[256];
    sprintf_s(buf, "Testing process-specific capture for PID: %lu", processId);
    Log(buf);

    // Initialize Windows Runtime
    Log("Initializing Windows Runtime with RO_INIT_SINGLETHREADED...");
    HRESULT hr = RoInitialize(RO_INIT_SINGLETHREADED);

    sprintf_s(buf, "RoInitialize result: 0x%08X", hr);
    Log(buf);

    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE && hr != S_FALSE) {
        Log("Failed to initialize Windows Runtime!");
        return 1;
    }

    // Check apartment type
    APTTYPE aptType;
    APTTYPEQUALIFIER aptQualifier;
    hr = CoGetApartmentType(&aptType, &aptQualifier);
    sprintf_s(buf, "Apartment type: %d (0=STA, 1=MTA, 2=NA, 3=MAINSTA)", (int)aptType);
    Log(buf);

    // Check Windows version
    typedef LONG (WINAPI* RtlGetVersionPtr)(PRTL_OSVERSIONINFOW);
    HMODULE hNtdll = GetModuleHandleW(L"ntdll.dll");
    if (hNtdll) {
        RtlGetVersionPtr RtlGetVersion = (RtlGetVersionPtr)GetProcAddress(hNtdll, "RtlGetVersion");
        if (RtlGetVersion) {
            RTL_OSVERSIONINFOW osvi = { 0 };
            osvi.dwOSVersionInfoSize = sizeof(osvi);
            if (RtlGetVersion(&osvi) == 0) {
                sprintf_s(buf, "Windows version: %lu.%lu Build %lu",
                    osvi.dwMajorVersion, osvi.dwMinorVersion, osvi.dwBuildNumber);
                Log(buf);
            }
        }
    }

    // Setup device ID
    LPCWSTR deviceId = VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK;
    Log("Device ID: VAD\\Process_Loopback");

    // Setup activation parameters
    AUDIOCLIENT_ACTIVATION_PARAMS activationParams = {};
    activationParams.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;
    activationParams.ProcessLoopbackParams.ProcessLoopbackMode = PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE;
    activationParams.ProcessLoopbackParams.TargetProcessId = processId;

    sprintf_s(buf, "Activation Type: %d", activationParams.ActivationType);
    Log(buf);
    sprintf_s(buf, "Loopback Mode: %d", activationParams.ProcessLoopbackParams.ProcessLoopbackMode);
    Log(buf);
    sprintf_s(buf, "Target PID: %lu", activationParams.ProcessLoopbackParams.TargetProcessId);
    Log(buf);

    PROPVARIANT activateParams = {};
    activateParams.vt = VT_BLOB;
    activateParams.blob.cbSize = sizeof(activationParams);
    activateParams.blob.pBlobData = (BYTE*)&activationParams;

    sprintf_s(buf, "PROPVARIANT vt: %d (VT_BLOB=%d)", activateParams.vt, VT_BLOB);
    Log(buf);
    sprintf_s(buf, "PROPVARIANT blob size: %lu", activateParams.blob.cbSize);
    Log(buf);

    // Try to activate
    Log("\nCalling ActivateAudioInterfaceAsync...");

    // Simple completion handler - must implement IAgileObject to avoid E_ILLEGAL_METHOD_CALL
    class SimpleHandler : public IActivateAudioInterfaceCompletionHandler, public IAgileObject {
        LONG m_refCount = 1;
        HANDLE m_event;
    public:
        HRESULT m_result = E_FAIL;

        SimpleHandler() {
            m_event = CreateEvent(nullptr, TRUE, FALSE, nullptr);
        }
        ~SimpleHandler() {
            if (m_event) CloseHandle(m_event);
        }

        STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject) {
            if (!ppvObject) return E_POINTER;
            if (riid == __uuidof(IUnknown) || riid == __uuidof(IActivateAudioInterfaceCompletionHandler)) {
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

        STDMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&m_refCount); }
        STDMETHODIMP_(ULONG) Release() {
            LONG ref = InterlockedDecrement(&m_refCount);
            if (ref == 0) delete this;
            return ref;
        }

        STDMETHODIMP ActivateCompleted(IActivateAudioInterfaceAsyncOperation* operation) {
            if (operation) {
                IUnknown* audioInterface = nullptr;
                operation->GetActivateResult(&m_result, &audioInterface);
                if (audioInterface) audioInterface->Release();
            }
            SetEvent(m_event);
            return S_OK;
        }

        bool Wait(DWORD timeout) {
            return WaitForSingleObject(m_event, timeout) == WAIT_OBJECT_0;
        }
    };

    SimpleHandler* handler = new SimpleHandler();
    IActivateAudioInterfaceAsyncOperation* asyncOp = nullptr;

    hr = ActivateAudioInterfaceAsync(
        deviceId,
        __uuidof(IAudioClient),
        &activateParams,
        handler,
        &asyncOp);

    sprintf_s(buf, "ActivateAudioInterfaceAsync result: 0x%08X", hr);
    Log(buf);

    if (SUCCEEDED(hr) && asyncOp) {
        Log("Async operation created successfully");
        Log("Waiting for completion (5 seconds)...");

        if (handler->Wait(5000)) {
            Log("Activation completed!");
            sprintf_s(buf, "Activation result: 0x%08X", handler->m_result);
            Log(buf);

            if (SUCCEEDED(handler->m_result)) {
                Log("SUCCESS: Per-process audio client activated!");
            } else {
                Log("FAILED: Activation completed but with error");
                _com_error err(handler->m_result);
                sprintf_s(buf, "Error: %S", err.ErrorMessage());
                Log(buf);
            }
        } else {
            Log("Timeout waiting for activation");
        }

        asyncOp->Release();
    } else {
        Log("FAILED: ActivateAudioInterfaceAsync failed");
        _com_error err(hr);
        sprintf_s(buf, "Error: %S", err.ErrorMessage());
        Log(buf);
    }

    handler->Release();
    RoUninitialize();

    Log("\nTest complete - check test_output.txt for results");
    g_logFile.close();

    return 0;
}
