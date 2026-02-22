#pragma once
#include "RSLinxHook_fwd.h"

// ============================================================
// SEH-safe utility functions
// These use __try/__except and must be in their own translation
// unit to avoid MSVC C2712 (cannot use C++ objects with
// destructors in functions using SEH).
// ============================================================

bool SafeVariantToString(VARIANT* pAddr, wchar_t* outBuf, int bufLen);
bool SafeReadMemory(void* pAddr, BYTE* outBuf, int bytes);
HRESULT TryStartAtSlot(void* pInterface, IUnknown* pPath, int slot);
HRESULT TryVtableGetObject(void* pInterface, int slot, IUnknown** ppResult);
HRESULT TryVtableGetLabel(IUnknown* pObj, int slot, std::wstring& outLabel);
HRESULT TryVtableAddPort(void* pDevice, int slot, GUID* pClsid,
                          const wchar_t* label, IID* pIID, IUnknown** ppResult);
void SafeRelease(IUnknown* pUnk, const wchar_t* label);
