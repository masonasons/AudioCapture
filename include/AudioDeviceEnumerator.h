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

    // Get all available audio capture (input/microphone) devices
    bool EnumerateInputDevices();

    // Get the list of output devices
    const std::vector<AudioDeviceInfo>& GetDevices() const { return m_devices; }

    // Get the list of input devices
    const std::vector<AudioDeviceInfo>& GetInputDevices() const { return m_inputDevices; }

    // Get default device index
    int GetDefaultDeviceIndex() const;

    // Get default input device index
    int GetDefaultInputDeviceIndex() const;

private:
    IMMDeviceEnumerator* m_deviceEnumerator;
    std::vector<AudioDeviceInfo> m_devices;
    std::vector<AudioDeviceInfo> m_inputDevices;
};
