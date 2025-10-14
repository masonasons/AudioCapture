#include "OutputDestinationManager.h"

OutputDestinationManager::OutputDestinationManager() {
}

OutputDestinationManager::~OutputDestinationManager() {
    CloseAll();
}

OutputDestinationPtr OutputDestinationManager::CreateDestination(DestinationType type) {
    m_lastError.clear();

    switch (type) {
        case DestinationType::FileWAV:
            return std::make_shared<WavFileDestination>();

        case DestinationType::FileMP3:
            return std::make_shared<Mp3FileDestination>();

        case DestinationType::FileOpus:
            return std::make_shared<OpusFileDestination>();

        case DestinationType::FileFLAC:
            return std::make_shared<FlacFileDestination>();

        case DestinationType::AudioDevice:
            return std::make_shared<DeviceOutputDestination>();

        default:
            m_lastError = L"Unknown destination type";
            return nullptr;
    }
}

bool OutputDestinationManager::AddDestination(OutputDestinationPtr destination) {
    if (!destination) {
        m_lastError = L"Cannot add null destination";
        return false;
    }

    if (!destination->IsOpen()) {
        m_lastError = L"Cannot add destination that is not configured/open";
        return false;
    }

    m_destinations.push_back(std::move(destination));
    return true;
}

bool OutputDestinationManager::RemoveDestination(size_t index) {
    if (index >= m_destinations.size()) {
        m_lastError = L"Invalid destination index";
        return false;
    }

    // Close the destination before removing it
    m_destinations[index]->Close();

    // Remove from the vector
    m_destinations.erase(m_destinations.begin() + index);

    return true;
}

size_t OutputDestinationManager::RemoveDestinationsByType(DestinationType type) {
    size_t removedCount = 0;

    // Iterate backwards so removal doesn't affect indices
    for (int i = static_cast<int>(m_destinations.size()) - 1; i >= 0; i--) {
        if (m_destinations[i]->GetType() == type) {
            m_destinations[i]->Close();
            m_destinations.erase(m_destinations.begin() + i);
            removedCount++;
        }
    }

    return removedCount;
}

OutputDestination* OutputDestinationManager::GetDestination(size_t index) {
    if (index >= m_destinations.size()) {
        return nullptr;
    }
    return m_destinations[index].get();
}

const OutputDestination* OutputDestinationManager::GetDestination(size_t index) const {
    if (index >= m_destinations.size()) {
        return nullptr;
    }
    return m_destinations[index].get();
}

size_t OutputDestinationManager::WriteAudioToAll(const BYTE* data, UINT32 size) {
    if (!data || size == 0) {
        return 0;
    }

    size_t successCount = 0;
    std::vector<size_t> failedIndices;

    // Write to all destinations
    for (size_t i = 0; i < m_destinations.size(); i++) {
        if (m_destinations[i]->WriteAudioData(data, size)) {
            successCount++;
        } else {
            // Mark this destination for removal
            failedIndices.push_back(i);

            // Log the error
            std::wstring error = m_destinations[i]->GetLastError();
            if (!error.empty()) {
                m_lastError = m_destinations[i]->GetName() + L": " + error;
            }
        }
    }

    // Remove failed destinations (iterate backwards to maintain indices)
    for (int i = static_cast<int>(failedIndices.size()) - 1; i >= 0; i--) {
        size_t index = failedIndices[i];
        m_destinations[index]->Close();
        m_destinations.erase(m_destinations.begin() + index);
    }

    return successCount;
}

void OutputDestinationManager::CloseAll() {
    for (auto& destination : m_destinations) {
        if (destination) {
            destination->Close();
        }
    }
    m_destinations.clear();
}

std::vector<std::wstring> OutputDestinationManager::GetDestinationNames() const {
    std::vector<std::wstring> names;
    names.reserve(m_destinations.size());

    for (const auto& destination : m_destinations) {
        if (destination) {
            names.push_back(destination->GetName());
        }
    }

    return names;
}

std::vector<DestinationType> OutputDestinationManager::GetDestinationTypes() const {
    std::vector<DestinationType> types;
    types.reserve(m_destinations.size());

    for (const auto& destination : m_destinations) {
        if (destination) {
            types.push_back(destination->GetType());
        }
    }

    return types;
}

bool OutputDestinationManager::CreateAndAddDestination(DestinationType type,
                                                        const WAVEFORMATEX* format,
                                                        const DestinationConfig& config) {
    m_lastError.clear();

    // Validate inputs
    if (!format) {
        m_lastError = L"Audio format is null";
        return false;
    }

    // Create the destination
    auto destination = CreateDestination(type);
    if (!destination) {
        // m_lastError is already set by CreateDestination
        return false;
    }

    // Configure the destination
    if (!destination->Configure(format, config)) {
        m_lastError = destination->GetLastError();
        return false;
    }

    // Add to active list
    if (!AddDestination(std::move(destination))) {
        // m_lastError is already set by AddDestination
        return false;
    }

    return true;
}
