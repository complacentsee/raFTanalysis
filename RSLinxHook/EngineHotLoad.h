#pragma once
#include "RSLinxHook_fwd.h"

// ============================================================
// ENGINE.DLL hot-load — registry-based driver discovery and loading
// Globals defined in EngineHotLoad.cpp
// ============================================================

extern std::wstring g_engineDriverName;

const char* GetDrvFileForDriverID(DWORD driverID);
bool FindDriverTypeAndDrv(const char* driverName, char* outDriverType, size_t typeLen,
                           char* outDrvFile, size_t drvLen);
void TryEngineHotLoad(const wchar_t* driverNameW);
HRESULT DoEngineHotLoadOnMainSTA();
