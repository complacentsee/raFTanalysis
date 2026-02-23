#pragma once
#include "RSLinxHook_fwd.h"
#include "ComInterfaces.h"
#include "Config.h"

// ============================================================
// Browse operations — core browse logic and monitor loop
// Globals defined in BrowseOperations.cpp
// ============================================================

struct ConnectionPointInfo { IConnectionPoint* pCP; DWORD cookie; };
struct EnumeratorInfo { void* pEnumInterface; DualEventSink* pSink; };

extern std::vector<ConnectionPointInfo> g_connectionPoints;
extern std::vector<EnumeratorInfo> g_enumerators;
extern DualEventSink* g_pMainSink;
extern IUnknown* g_pMainEnumUnk;

bool EnumeratorsCycledSince(int baseline);
void GetEnumeratorStatusSince(int baseline, int& completed, int& total);

IDispatch* GetBusDispatch(const wchar_t* driverName);
HRESULT DoMainSTABrowse();
HRESULT DoBusBrowse();
HRESULT DoBackplaneBrowse();
HRESULT DoCleanupOnMainSTA();
void RunMonitorLoop(const HookConfig& config, IRSTopologyGlobals* pGlobals, const std::vector<BusInfo>& buses);

// Direct COM topology tree walk (no XML dependency)
struct TopoNode;  // Forward declaration (defined in TopologyXML.h)
TopoNode DoTopologyTreeWalk();
