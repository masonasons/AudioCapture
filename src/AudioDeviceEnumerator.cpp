#include "AudioDeviceEnumerator.h"
#include <functiondiscoverykeys_devpkey.h>

AudioDeviceEnumerator::AudioDeviceEnumerator()
    : m_deviceEnumerator(nullptr)
{
}

AudioDeviceEnumerator::~AudioDeviceEnumerator() {
    if (m_deviceEnumerator) {
        m_deviceEnumerator->Release();
    }
}

bool AudioDeviceEnumerator::EnumerateDevices() {
    m_devices.clear();

    // Release existing enumerator if any
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

    // Get default device
    IMMDevice* defaultDevice = nullptr;
    std::wstring defaultDeviceId;
    hr = m_deviceEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &defaultDevice);
    if (SUCCEEDED(hr) && defaultDevice) {
        LPWSTR deviceId = nullptr;
        if (SUCCEEDED(defaultDevice->GetId(&deviceId))) {
            defaultDeviceId = deviceId;
            CoTaskMemFree(deviceId);
        }
        defaultDevice->Release();
    }

    // Enumerate all active render devices
    IMMDeviceCollection* deviceCollection = nullptr;
    hr = m_deviceEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &deviceCollection);

    if (FAILED(hr)) {
        return false;
    }

    UINT deviceCount = 0;
    hr = deviceCollection->GetCount(&deviceCount);

    if (SUCCEEDED(hr)) {
        for (UINT i = 0; i < deviceCount; i++) {
            IMMDevice* device = nullptr;
            hr = deviceCollection->Item(i, &device);

            if (SUCCEEDED(hr) && device) {
                AudioDeviceInfo info;

                // Get device ID
                LPWSTR deviceId = nullptr;
                if (SUCCEEDED(device->GetId(&deviceId))) {
                    info.deviceId = deviceId;
                    info.isDefault = (info.deviceId == defaultDeviceId);
                    CoTaskMemFree(deviceId);
                }

                // Get friendly name
                IPropertyStore* propertyStore = nullptr;
                hr = device->OpenPropertyStore(STGM_READ, &propertyStore);
                if (SUCCEEDED(hr) && propertyStore) {
                    PROPVARIANT friendlyName;
                    PropVariantInit(&friendlyName);

                    hr = propertyStore->GetValue(PKEY_Device_FriendlyName, &friendlyName);
                    if (SUCCEEDED(hr)) {
                        if (friendlyName.vt == VT_LPWSTR && friendlyName.pwszVal) {
                            info.friendlyName = friendlyName.pwszVal;
                        }
                    }

                    PropVariantClear(&friendlyName);
                    propertyStore->Release();
                }

                device->Release();

                // Add to list if we have a name
                if (!info.friendlyName.empty()) {
                    m_devices.push_back(info);
                }
            }
        }
    }

    deviceCollection->Release();
    return true;
}

int AudioDeviceEnumerator::GetDefaultDeviceIndex() const {
    for (size_t i = 0; i < m_devices.size(); i++) {
        if (m_devices[i].isDefault) {
            return static_cast<int>(i);
        }
    }
    return 0; // Return first device if no default found
}
