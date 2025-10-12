#pragma once

#include "AudioCapture.h"
#include "AudioMixer.h"
#include "WavWriter.h"
#include "Mp3Encoder.h"
#include "OpusEncoder.h"
#include "FlacEncoder.h"
#include <memory>
#include <map>
#include <mutex>
#include <thread>
#include <atomic>

enum class AudioFormat {
    WAV,
    MP3,
    OPUS,
    FLAC
};

struct CaptureSession {
    DWORD processId;
    std::wstring processName;
    std::wstring outputFile;
    AudioFormat format;
    std::unique_ptr<AudioCapture> capture;
    std::unique_ptr<WavWriter> wavWriter;
    std::unique_ptr<Mp3Encoder> mp3Encoder;
    std::unique_ptr<OpusEncoder> opusEncoder;
    std::unique_ptr<FlacEncoder> flacEncoder;
    bool isActive;
    UINT64 bytesWritten;
    bool skipSilence;
    bool monitorOnly;
};

class CaptureManager {
public:
    CaptureManager();
    ~CaptureManager();

    // Start capturing from a process
    bool StartCapture(DWORD processId, const std::wstring& processName,
                     const std::wstring& outputPath, AudioFormat format,
                     UINT32 bitrate = 0, bool skipSilence = false,
                     const std::wstring& passthroughDeviceId = L"",
                     bool monitorOnly = false);

    // Start capturing from an audio device (microphone/line-in)
    bool StartCaptureFromDevice(DWORD sessionId, const std::wstring& deviceName,
                                const std::wstring& deviceId, bool isInputDevice,
                                const std::wstring& outputPath, AudioFormat format,
                                UINT32 bitrate = 0, bool skipSilence = false,
                                bool monitorOnly = false);

    // Enable mixed recording (all processes will be mixed into one file)
    bool EnableMixedRecording(const std::wstring& outputPath, AudioFormat format, UINT32 bitrate = 0);

    // Disable mixed recording
    void DisableMixedRecording();

    // Check if mixed recording is active
    bool IsMixedRecordingActive() const;

    // Stop capturing from a specific process
    bool StopCapture(DWORD processId);

    // Stop all captures
    void StopAllCaptures();

    // Pause all captures
    void PauseAllCaptures();

    // Resume all captures
    void ResumeAllCaptures();

    // Get active capture sessions
    std::vector<CaptureSession*> GetActiveSessions();

    // Check if a process is being captured
    bool IsCapturing(DWORD processId) const;

private:
    void OnAudioData(DWORD processId, const BYTE* data, UINT32 size);
    void MixerThread();

    std::map<DWORD, std::unique_ptr<CaptureSession>> m_sessions;
    std::mutex m_mutex;

    // Mixed recording members
    bool m_mixedRecordingEnabled;
    std::unique_ptr<AudioMixer> m_mixer;
    std::unique_ptr<WavWriter> m_mixedWavWriter;
    std::unique_ptr<Mp3Encoder> m_mixedMp3Encoder;
    std::unique_ptr<OpusEncoder> m_mixedOpusEncoder;
    std::unique_ptr<FlacEncoder> m_mixedFlacEncoder;
    AudioFormat m_mixedFormat;
    std::unique_ptr<std::thread> m_mixerThread;
    std::atomic<bool> m_mixerThreadRunning;
    std::mutex m_mixerMutex;
};
