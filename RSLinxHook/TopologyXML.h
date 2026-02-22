#pragma once
#include "RSLinxHook_fwd.h"
#include "ComInterfaces.h"

// ============================================================
// XML topology handling — save, parse, and analyze
// Globals defined in TopologyXML.cpp
// ============================================================

struct TopologyCounts {
    int totalDevices;
    int identifiedDevices;
};

struct DeviceInfo {
    std::wstring ip;           // "10.39.31.200" (from topology XML address mapping)
    std::wstring productName;  // DISPID 1 = Name (e.g. "5069-L310ER LOGIX310ER")
    std::wstring objectId;     // DISPID 2 = topology objectid GUID
};

extern std::map<std::wstring, DeviceInfo> g_deviceDetails;

bool SaveTopologyXML(IRSTopologyGlobals* pGlobals, const wchar_t* filename);
TopologyCounts CountDevicesInXML(const wchar_t* filename);
bool IsTargetIdentifiedInXML(const wchar_t* filename, const std::vector<std::wstring>& targetIPs);
void UpdateDeviceIPsFromXML(const wchar_t* filename);
