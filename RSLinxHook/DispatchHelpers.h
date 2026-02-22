#pragma once
#include "RSLinxHook_fwd.h"

// ============================================================
// IDispatch helper functions for topology object enumeration
// ============================================================

int DispatchGetInt(IDispatch* pDisp, DISPID dispid);
std::wstring DispatchGetString(IDispatch* pDisp, DISPID dispid);
IDispatch* DispatchGetCollection(IDispatch* pDisp, DISPID dispid);
std::vector<IDispatch*> EnumerateCollection(IDispatch* pCollection);
IUnknown* DispatchGetPath(IDispatch* pDisp);

// DISPID discovery — probe device/bus objects for available properties
void ProbeDeviceDISPIDs(IDispatch* pDisp, const wchar_t* label);
void ProbeBusDISPIDs(IDispatch* pDisp, const wchar_t* label);
