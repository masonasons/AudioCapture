#pragma once

#include <windows.h>
#include <mmdeviceapi.h>
#include <string>
#include <vector>

struct AudioDeviceInfo {
    std::wstring deviceId;
    std::wstring friendlyName;
    bool isDefault;
};

class AudioDeviceEnumerator {
public:
    AudioDeviceEnumerator();
    ~AudioDeviceEnumerator();

    // Get all available audio render (output) devices
    bool EnumerateDevices();

    // Get the list of devices
    const std::vector<AudioDeviceInfo>& GetDevices() const { return m_devices; }

    // Get default device index
    int GetDefaultDeviceIndex() const;

private:
    IMMDeviceEnumerator* m_deviceEnumerator;
    std::vector<AudioDeviceInfo> m_devices;
};
