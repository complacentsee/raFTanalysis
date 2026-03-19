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

// Path query result: ip-only (portName="", slot=-1) or backplane slot
struct QueryResult {
    bool found = false;
    std::wstring classname;
    std::wstring deviceName;
    std::wstring ip;
    std::wstring portName;
    int slot = -1;
};

// In-memory query cache: key = "ip" or "ip\PortName\slot"
// Populated once after each browse phase; queried without any file I/O.
extern std::map<std::wstring, QueryResult> g_queryCache;

bool SaveTopologyXML(IRSTopologyGlobals* pGlobals, const wchar_t* filename);
TopologyCounts CountDevicesInXML(const wchar_t* filename);
int CountTargetsIdentifiedInXML(const wchar_t* filename, const std::vector<std::wstring>& targetIPs);
bool IsTargetIdentifiedInXML(const wchar_t* filename, const std::vector<std::wstring>& targetIPs);
void UpdateDeviceIPsFromXML(const wchar_t* filename);
void PopulateQueryCache(const wchar_t* xmlFile);
QueryResult QueryXMLForPath(const wchar_t* xmlFile,
                             const std::wstring& ip,
                             const std::wstring& portName,
                             int slot);

// Cache-based counting (fallback when SaveTopologyXML fails with DISP_E_EXCEPTION)
TopologyCounts CountDevicesFromCache();
int CountTargetsFromCache(const std::vector<std::wstring>& targetIPs);

// N| topology pipe: ordered Ethernet device names per driver (populated by DoBusBrowse).
// Used by WalkTopologyTree so no COM calls are needed from the worker thread.
extern std::map<std::wstring, std::vector<std::wstring>> g_driverDeviceNames;

// Emit the topology as N| messages (N|BEGIN...N|END) on the pipe.
// Must be called AFTER UpdateDeviceIPsFromXML + PopulateQueryCache for the current snapshot.
void WalkTopologyTree(IRSTopologyGlobals* pGlobals);
