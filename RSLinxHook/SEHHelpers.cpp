#include "SEHHelpers.h"
#include "Logging.h"

// ============================================================
// SEH-safe utility functions
// ============================================================

// SEH-safe VARIANT to wstring conversion (standalone C function, no C++ objects)
bool SafeVariantToString(VARIANT* pAddr, wchar_t* outBuf, int bufLen)
{
    __try
    {
        if (pAddr->vt == VT_BSTR && pAddr->bstrVal)
        {
            UINT len = SysStringLen(pAddr->bstrVal);
            if (len > 0 && len < (UINT)bufLen)
            {
                wcsncpy(outBuf, pAddr->bstrVal, bufLen - 1);
                outBuf[bufLen - 1] = 0;
                return true;
            }
        }
        else if (pAddr->vt != VT_EMPTY && pAddr->vt != VT_NULL)
        {
            VARIANT conv;
            VariantInit(&conv);
            if (SUCCEEDED(VariantChangeType(&conv, pAddr, 0, VT_BSTR)) && conv.bstrVal)
            {
                UINT len = SysStringLen(conv.bstrVal);
                if (len > 0 && len < (UINT)bufLen)
                {
                    wcsncpy(outBuf, conv.bstrVal, bufLen - 1);
                    outBuf[bufLen - 1] = 0;
                    VariantClear(&conv);
                    return true;
                }
            }
            VariantClear(&conv);
        }
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        // Bad pointer or corrupted VARIANT
    }
    return false;
}

// SEH-safe memory dump (no C++ objects allowed)
bool SafeReadMemory(void* pAddr, BYTE* outBuf, int bytes)
{
    __try
    {
        BYTE* pSrc = (BYTE*)pAddr;
        for (int i = 0; i < bytes; i++)
            outBuf[i] = pSrc[i];
        return true;
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        return false;
    }
}

HRESULT TryStartAtSlot(void* pInterface, IUnknown* pPath, int slot)
{
    typedef HRESULT (__stdcall *StartFunc)(void* pThis, IUnknown* pPath);
    HRESULT hr = E_FAIL;
    __try
    {
        void** vtable = *(void***)pInterface;
        StartFunc pfn = (StartFunc)vtable[slot];
        hr = pfn(pInterface, pPath);
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        hr = E_UNEXPECTED;
    }
    return hr;
}

// SEH-safe vtable call: pInterface->vtable[slot](&ppResult)
// For methods like GetBackplanePort, GetBus that return a single IUnknown** out param
HRESULT TryVtableGetObject(void* pInterface, int slot, IUnknown** ppResult)
{
    typedef HRESULT (__stdcall *GetObjFunc)(void* pThis, IUnknown** ppResult);
    *ppResult = nullptr;
    HRESULT hr = E_FAIL;
    __try
    {
        void** vtable = *(void***)pInterface;
        GetObjFunc pfn = (GetObjFunc)vtable[slot];
        hr = pfn(pInterface, ppResult);
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        hr = E_UNEXPECTED;
    }
    return hr;
}

// SEH-safe vtable call: pInterface->vtable[slot](buffer, bufLen)
// For methods like IRSTopologyPort::GetLabel that write to a WCHAR buffer
HRESULT TryVtableGetLabel(IUnknown* pObj, int slot, std::wstring& outLabel)
{
    outLabel.clear();
    if (!pObj) return E_POINTER;
    typedef HRESULT (__stdcall *GetLabelFunc)(void*, WCHAR*, int);
    HRESULT hr = E_FAIL;
    __try
    {
        void** vtable = *(void***)pObj;
        GetLabelFunc pfn = (GetLabelFunc)vtable[slot];
        wchar_t buf[256] = {0};
        hr = pfn(pObj, buf, 256);
        if (SUCCEEDED(hr))
            outLabel = buf;
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        hr = E_UNEXPECTED;
    }
    return hr;
}

// SEH-safe vtable call: pDevice->vtable[slot](clsid, label, region, hwnd, iid, ppResult)
// IRSTopologyDevice::AddPort — creates a new port on the device
HRESULT TryVtableAddPort(void* pDevice, int slot, GUID* pClsid,
                          const wchar_t* label, IID* pIID, IUnknown** ppResult)
{
    typedef HRESULT (__stdcall *AddPortFunc)(void*, GUID*, const wchar_t*, RECT*, HWND, IID*, void**);
    *ppResult = nullptr;
    HRESULT hr = E_FAIL;
    __try
    {
        void** vtable = *(void***)pDevice;
        AddPortFunc pfn = (AddPortFunc)vtable[slot];
        hr = pfn(pDevice, pClsid, label, nullptr, nullptr, pIID, (void**)ppResult);
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        hr = E_UNEXPECTED;
    }
    return hr;
}

// SEH-safe COM Release (standalone C function — no C++ objects)
void SafeRelease(IUnknown* pUnk, const wchar_t* label)
{
    __try
    {
        if (pUnk) pUnk->Release();
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        Log(L"[WARN] %s->Release() crashed (SEH caught)", label);
    }
}
