cmake_minimum_required(VERSION 3.15)
project(AudioCapture VERSION 1.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Output directories
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)

# Find packages
find_package(Opus CONFIG REQUIRED)
find_package(Ogg CONFIG REQUIRED)
find_package(FLAC CONFIG REQUIRED)
find_package(nlohmann_json CONFIG REQUIRED)

# Source files
set(SOURCES
    src/main.cpp
    src/AudioCapture.cpp
    src/ProcessEnumerator.cpp
    src/WavWriter.cpp
    src/Mp3Encoder.cpp
    src/OpusEncoder.cpp
    src/FlacEncoder.cpp
    src/CaptureManager.cpp
    src/AudioDeviceEnumerator.cpp
    src/AudioMixer.cpp
)

set(HEADERS
    include/AudioCapture.h
    include/ProcessEnumerator.h
    include/WavWriter.h
    include/Mp3Encoder.h
    include/OpusEncoder.h
    include/FlacEncoder.h
    include/CaptureManager.h
    include/AudioDeviceEnumerator.h
    include/AudioMixer.h
    include/resource.h
)

# Create executable
add_executable(AudioCapture WIN32 ${SOURCES} ${HEADERS})

# Include directories
target_include_directories(AudioCapture PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/include
)

# Link Windows libraries and third-party libraries
target_link_libraries(AudioCapture PRIVATE
    Avrt.lib
    Mmdevapi.lib
    Ksuser.lib
    Mfplat.lib
    Mfreadwrite.lib
    Mfuuid.lib
    Ole32.lib
    Shell32.lib
    Comctl32.lib
    Shlwapi.lib
    Propsys.lib
    RuntimeObject.lib
    Delayimp.lib
    Opus::opus
    Ogg::ogg
    FLAC::FLAC FLAC::FLAC++
    nlohmann_json::nlohmann_json
)

# Enable Windows subsystem and static runtime
set_target_properties(AudioCapture PROPERTIES
    WIN32_EXECUTABLE TRUE
    LINK_FLAGS "/DELAYLOAD:api-ms-win-core-winrt-l1-1-0.dll"
    MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>"
)

# Add compile definitions
target_compile_definitions(AudioCapture PRIVATE
    UNICODE
    _UNICODE
    WIN32_LEAN_AND_MEAN
    NOMINMAX
)

# Enable warnings
if(MSVC)
    target_compile_options(AudioCapture PRIVATE /W4)
else()
    target_compile_options(AudioCapture PRIVATE -Wall -Wextra)
endif()
