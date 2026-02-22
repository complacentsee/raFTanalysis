#pragma once
#include "RSLinxHook_fwd.h"

// ============================================================
// Configuration types and parsing
// ============================================================

enum class HookMode { Inject, Monitor };

struct DriverEntry {
    std::wstring name;
    std::vector<std::wstring> ipAddresses;
    bool newDriver = false;
};

struct HookConfig
{
    std::vector<DriverEntry> drivers;
    HookMode mode = HookMode::Inject;
    std::wstring logDir = L"C:\\temp";
    bool debugXml = false;
    bool probeDispids = false;

    // Backward compat helpers
    const std::wstring& driverName() const { return drivers[0].name; }
    const std::vector<std::wstring>& ipAddresses() const { return drivers[0].ipAddresses; }
    bool newDriver() const {
        for (auto& d : drivers) if (d.newDriver) return true;
        return false;
    }
    // Aggregate all IPs across all drivers
    std::vector<std::wstring> allIPs() const {
        std::vector<std::wstring> all;
        for (auto& d : drivers)
            for (auto& ip : d.ipAddresses)
                all.push_back(ip);
        return all;
    }
};

struct BusInfo {
    std::wstring driverName;
    IDispatch* pBusDisp;
    IUnknown* pBusUnk;
};

std::wstring Utf8ToWide(const char* utf8);
bool ReadConfigFromPipe(HookConfig& config);
