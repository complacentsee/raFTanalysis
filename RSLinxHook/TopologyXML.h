#pragma once
#include "RSLinxHook_fwd.h"
#include "ComInterfaces.h"

// ============================================================
// Topology data structures and handling
// Globals defined in TopologyXML.cpp
// ============================================================

struct TopologyCounts {
    int totalDevices;
    int identifiedDevices;
};

struct DeviceInfo {
    std::wstring name;         // DISPID 1 (IADs) = display name
    std::wstring classname;    // ITopologyObject DISPID 3 = device class/model
    std::wstring objectId;     // DISPID 2 (IADs) = topology objectid GUID
    std::wstring ip;           // From event sink (Ethernet) or ""
    int slot;                  // Collection index (backplane) or -1
};

// Hierarchical topology node — mirrors XML schema structure
struct TopoNode {
    enum NodeType { Workstation, Bus, Device, Port };
    NodeType type;
    std::wstring name;
    std::wstring classname;
    std::wstring objectId;
    std::wstring address;      // IP (Ethernet) or slot number string (backplane)
    std::wstring addressType;  // "String", "Short", or ""
    std::vector<TopoNode> children;
};

extern std::map<std::wstring, DeviceInfo> g_deviceDetails;

// XML-based functions (legacy, kept for debugXml mode)
bool SaveTopologyXML(IRSTopologyGlobals* pGlobals, const wchar_t* filename);
TopologyCounts CountDevicesInXML(const wchar_t* filename);
bool IsTargetIdentifiedInXML(const wchar_t* filename, const std::vector<std::wstring>& targetIPs);
void UpdateDeviceIPsFromXML(const wchar_t* filename);

// Tree-based functions (direct COM walk, no XML)
TopologyCounts CountDevicesInTree(const TopoNode& root);
bool IsTargetIdentifiedInTree(const TopoNode& root, const std::vector<std::wstring>& targetIPs);
