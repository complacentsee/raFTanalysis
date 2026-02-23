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

// ITopologyObject property extraction — gets classname, path, parent via dual interface
struct TopoObjectProps {
    std::wstring name;
    std::wstring classname;
    std::wstring objectId;
    std::wstring path;
    std::wstring parent;
    std::wstring schemaPath;
};
TopoObjectProps GetTopoObjectProps(IDispatch* pDefaultDisp);

// DISPID discovery — probe device/bus objects for available properties
void ProbeDeviceDISPIDs(IDispatch* pDisp, const wchar_t* label);
void ProbeBusDISPIDs(IDispatch* pDisp, const wchar_t* label);
