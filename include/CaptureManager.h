#pragma once

#include "AudioCapture.h"
#include "WavWriter.h"
#include "Mp3Encoder.h"
#include "OpusEncoder.h"
#include "FlacEncoder.h"
#include <memory>
#include <map>
#include <mutex>

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
};

class CaptureManager {
public:
    CaptureManager();
    ~CaptureManager();

    // Start capturing from a process
    bool StartCapture(DWORD processId, const std::wstring& processName,
                     const std::wstring& outputPath, AudioFormat format,
                     UINT32 bitrate = 0, bool skipSilence = false);

    // Stop capturing from a specific process
    bool StopCapture(DWORD processId);

    // Stop all captures
    void StopAllCaptures();

    // Get active capture sessions
    std::vector<CaptureSession*> GetActiveSessions();

    // Check if a process is being captured
    bool IsCapturing(DWORD processId) const;

private:
    void OnAudioData(DWORD processId, const BYTE* data, UINT32 size);

    std::map<DWORD, std::unique_ptr<CaptureSession>> m_sessions;
    std::mutex m_mutex;
};
