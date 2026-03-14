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

// Browse-state tracking: which driver names / IPs have completed each phase
extern std::set<std::wstring> g_browsedDrivers;    // driver names with Phase 2+3 complete
extern std::set<std::wstring> g_browsedBackplanes; // IPs with Phase 4+4b complete

bool EnumeratorsCycledSince(int baseline);
void GetEnumeratorStatusSince(int baseline, int& completed, int& total);

IDispatch* GetBusDispatch(const wchar_t* driverName);
HRESULT DoMainSTABrowse();
HRESULT DoBusBrowse();
HRESULT DoBackplaneBrowse();
HRESULT DoCleanupOnMainSTA();
void RunMonitorLoop(const HookConfig& config, IRSTopologyGlobals* pGlobals, const std::vector<BusInfo>& buses);
