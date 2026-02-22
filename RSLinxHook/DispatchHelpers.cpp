#include "DispatchHelpers.h"
#include "Logging.h"
#include "ComInterfaces.h"

// ============================================================
// IDispatch helper implementations
// ============================================================

// Get int property via DISPID (no args)
int DispatchGetInt(IDispatch* pDisp, DISPID dispid)
{
    if (!pDisp) return -1;
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    VARIANT result;
    VariantInit(&result);
    HRESULT hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);
    if (FAILED(hr))
    {
        Log(L"[DISP] GetInt(DISPID %d): FAILED hr=0x%08x", dispid, hr);
        return -1;
    }
    int val = -1;
    if (result.vt == VT_I4) val = result.lVal;
    else if (result.vt == VT_I2) val = result.iVal;
    else Log(L"[DISP] GetInt(DISPID %d): unexpected vt=%d", dispid, result.vt);
    VariantClear(&result);
    return val;
}

// Get string property via DISPID (no args)
std::wstring DispatchGetString(IDispatch* pDisp, DISPID dispid)
{
    if (!pDisp) return L"";
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    VARIANT result;
    VariantInit(&result);
    HRESULT hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);
    if (FAILED(hr)) { VariantClear(&result); return L""; }
    std::wstring s;
    if (result.vt == VT_BSTR && result.bstrVal)
        s = result.bstrVal;
    VariantClear(&result);
    return s;
}

// Get collection property via DISPID (no args) — returns IDispatch*
// Tries to QI returned object for ITopologyCollection dispatch to get correct DISPIDs
IDispatch* DispatchGetCollection(IDispatch* pDisp, DISPID dispid)
{
    if (!pDisp) return nullptr;
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    VARIANT result;
    VariantInit(&result);
    HRESULT hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);
    if (FAILED(hr))
    {
        Log(L"[DISP] GetCollection(DISPID %d): FAILED hr=0x%08x", dispid, hr);
        VariantClear(&result);
        return nullptr;
    }
    Log(L"[DISP] GetCollection(DISPID %d): OK vt=%d", dispid, result.vt);

    // Get underlying IUnknown first
    IUnknown* pUnk = nullptr;
    if (result.vt == VT_DISPATCH && result.pdispVal)
        result.pdispVal->QueryInterface(IID_IUnknown, (void**)&pUnk);
    else if (result.vt == VT_UNKNOWN && result.punkVal)
        result.punkVal->QueryInterface(IID_IUnknown, (void**)&pUnk);
    VariantClear(&result);

    if (!pUnk) return nullptr;

    // Try QI for ITopologyCollection dispatch — this gives correct DISPIDs (0=GetObject, 1=Count)
    // The default IDispatch is often the IADs wrapper which doesn't have topology DISPIDs
    IDispatch* pResult = nullptr;
    hr = pUnk->QueryInterface(IID_ITopologyCollection, (void**)&pResult);
    if (SUCCEEDED(hr) && pResult)
    {
        Log(L"[DISP] QI for ITopologyCollection: OK");
    }
    else
    {
        // Fallback: use default IDispatch
        hr = pUnk->QueryInterface(IID_IDispatch, (void**)&pResult);
        Log(L"[DISP] QI for ITopologyCollection FAILED, using default IDispatch");
    }
    pUnk->Release();
    return pResult;
}

// Enumerate collection using _NewEnum → IEnumVARIANT
// Returns all items as a vector of IDispatch* (caller must Release each)
std::vector<IDispatch*> EnumerateCollection(IDispatch* pCollection)
{
    std::vector<IDispatch*> items;
    if (!pCollection) return items;

    // Get _NewEnum (DISPID -4)
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    VARIANT result;
    VariantInit(&result);
    HRESULT hr = pCollection->Invoke(-4, IID_NULL, LOCALE_USER_DEFAULT,
                                      DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);
    if (FAILED(hr))
    {
        Log(L"[DISP] _NewEnum (DISPID -4): FAILED hr=0x%08x", hr);
        return items;
    }

    IEnumVARIANT* pEnum = nullptr;
    if (result.vt == VT_UNKNOWN && result.punkVal)
    {
        result.punkVal->QueryInterface(IID_IEnumVARIANT, (void**)&pEnum);
    }
    else if (result.vt == VT_DISPATCH && result.pdispVal)
    {
        result.pdispVal->QueryInterface(IID_IEnumVARIANT, (void**)&pEnum);
    }
    VariantClear(&result);

    if (!pEnum)
    {
        Log(L"[DISP] _NewEnum: could not get IEnumVARIANT (vt was %d)", result.vt);
        return items;
    }

    // Iterate
    VARIANT item;
    ULONG fetched = 0;
    while (pEnum->Next(1, &item, &fetched) == S_OK && fetched > 0)
    {
        IUnknown* pUnk = nullptr;
        if (item.vt == VT_DISPATCH && item.pdispVal)
            item.pdispVal->QueryInterface(IID_IUnknown, (void**)&pUnk);
        else if (item.vt == VT_UNKNOWN && item.punkVal)
            item.punkVal->QueryInterface(IID_IUnknown, (void**)&pUnk);
        VariantClear(&item);

        if (pUnk)
        {
            // QI for ITopologyObject dispatch (correct DISPIDs: 1=Name, 4=path)
            // Then try bus/chassis dispatch, fallback to default IDispatch
            IDispatch* pDisp = nullptr;
            HRESULT hrQI = pUnk->QueryInterface(IID_ITopologyObject, (void**)&pDisp);
            if (FAILED(hrQI))
                hrQI = pUnk->QueryInterface(IID_ITopologyBus, (void**)&pDisp);
            if (FAILED(hrQI))
                hrQI = pUnk->QueryInterface(IID_ITopologyChassis, (void**)&pDisp);
            if (FAILED(hrQI))
                pUnk->QueryInterface(IID_IDispatch, (void**)&pDisp);
            pUnk->Release();

            if (pDisp)
                items.push_back(pDisp);
        }
        fetched = 0;
    }
    pEnum->Release();

    Log(L"[DISP] _NewEnum: enumerated %d items", (int)items.size());
    return items;
}

// ============================================================
// DISPID Discovery — probe objects for available properties
// ============================================================

// Helper: convert VARIANT to a loggable string representation
static std::wstring VariantToLogString(const VARIANT& v)
{
    wchar_t buf[512];
    switch (v.vt)
    {
    case VT_EMPTY:  return L"(empty)";
    case VT_NULL:   return L"(null)";
    case VT_I2:     swprintf(buf, 512, L"%d", v.iVal); return buf;
    case VT_I4:     swprintf(buf, 512, L"%d", v.lVal); return buf;
    case VT_R4:     swprintf(buf, 512, L"%.3f", v.fltVal); return buf;
    case VT_R8:     swprintf(buf, 512, L"%.3f", v.dblVal); return buf;
    case VT_BOOL:   return v.boolVal ? L"TRUE" : L"FALSE";
    case VT_BSTR:   return v.bstrVal ? v.bstrVal : L"(null bstr)";
    case VT_DISPATCH: swprintf(buf, 512, L"IDispatch*=0x%p", v.pdispVal); return buf;
    case VT_UNKNOWN:  swprintf(buf, 512, L"IUnknown*=0x%p", v.punkVal); return buf;
    default:
        swprintf(buf, 512, L"vt=%d", v.vt);
        return buf;
    }
}

// Probe a single IDispatch for named properties and DISPID range
static void ProbeDispatch(IDispatch* pDisp, const wchar_t* ifaceName, const wchar_t* objLabel,
                           const wchar_t* const* propNames, int numNames, int dispidMin, int dispidMax)
{
    if (!pDisp) return;

    Log(L"[PROBE:%s] === Probing via %s ===", objLabel, ifaceName);

    // Phase 1: GetIDsOfNames for candidate properties
    for (int i = 0; i < numNames; i++)
    {
        LPOLESTR name = const_cast<LPOLESTR>(propNames[i]);
        DISPID dispid = -999;
        HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (SUCCEEDED(hr))
            Log(L"[PROBE:%s]   GetIDsOfNames(\"%s\") = DISPID %d", objLabel, propNames[i], dispid);
    }

    // Phase 2: Range scan DISPIDs with PROPERTYGET (no args)
    for (DISPID d = dispidMin; d <= dispidMax; d++)
    {
        DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
        VARIANT result;
        VariantInit(&result);
        HRESULT hr = pDisp->Invoke(d, IID_NULL, LOCALE_USER_DEFAULT,
                                    DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);
        if (hr != (HRESULT)0x80020003)  // DISP_E_MEMBERNOTFOUND
        {
            std::wstring val = VariantToLogString(result);
            Log(L"[PROBE:%s]   DISPID %d: hr=0x%08x vt=%d val=%s",
                objLabel, d, hr, result.vt, val.c_str());
        }
        VariantClear(&result);
    }
}

void ProbeDeviceDISPIDs(IDispatch* pDisp, const wchar_t* label)
{
    if (!pDisp) return;

    static const wchar_t* devicePropNames[] = {
        // Known
        L"Name", L"ObjectID", L"Path", L"Devices", L"Busses",
        // Candidates — identity
        L"Class", L"ClassName", L"classname", L"ClassGuid", L"ClassID",
        L"Vendor", L"VendorId", L"VendorID",
        L"ProductType", L"ProductCode", L"ProductName",
        L"Serial", L"SerialNumber",
        L"Revision", L"MajorRevision", L"MinorRevision",
        L"Description", L"Label",
        // Candidates — state
        L"Online", L"OnlineState", L"State", L"Status",
        // Candidates — structure
        L"Address", L"Addresses", L"Ports", L"Parent", L"Children",
        L"Port", L"Bus", L"Chassis",
    };
    static const int numDeviceProps = sizeof(devicePropNames) / sizeof(devicePropNames[0]);

    // Probe via the IDispatch we already have
    ProbeDispatch(pDisp, L"IDispatch", label, devicePropNames, numDeviceProps, 0, 100);

    // Probe via alternate interfaces (may expose different DISPIDs)
    IUnknown* pUnk = nullptr;
    pDisp->QueryInterface(IID_IUnknown, (void**)&pUnk);
    if (!pUnk) return;

    struct { const GUID& iid; const wchar_t* name; } interfaces[] = {
        { IID_ITopologyObject,       L"ITopologyObject" },
        { IID_ITopologyDevice_Dual,  L"ITopologyDevice_Dual" },
        { IID_ITopologyBus,          L"ITopologyBus" },
        { IID_IRSTopologyObject,     L"IRSTopologyObject" },
        { IID_IRSTopologyDevice,     L"IRSTopologyDevice" },
    };

    for (auto& iface : interfaces)
    {
        IDispatch* pAlt = nullptr;
        HRESULT hr = pUnk->QueryInterface(iface.iid, (void**)&pAlt);
        if (SUCCEEDED(hr) && pAlt && pAlt != pDisp)
        {
            ProbeDispatch(pAlt, iface.name, label, devicePropNames, numDeviceProps, 0, 100);
            pAlt->Release();
        }
        else if (pAlt == pDisp)
        {
            Log(L"[PROBE:%s] %s == same IDispatch, skipping", label, iface.name);
            pAlt->Release();
        }
    }
    pUnk->Release();
}

void ProbeBusDISPIDs(IDispatch* pDisp, const wchar_t* label)
{
    if (!pDisp) return;

    static const wchar_t* busPropNames[] = {
        L"Name", L"ObjectID", L"Path",
        L"Devices", L"Addresses", L"Ports",
        L"Class", L"ClassName", L"classname",
        L"Count", L"Item", L"GetObject",
        L"Parent", L"Children",
        L"Online", L"State", L"Status",
        L"BusType", L"AddressType",
    };
    static const int numBusProps = sizeof(busPropNames) / sizeof(busPropNames[0]);

    ProbeDispatch(pDisp, L"IDispatch", label, busPropNames, numBusProps, 0, 60);

    // Also try ITopologyBus interface
    IUnknown* pUnk = nullptr;
    pDisp->QueryInterface(IID_IUnknown, (void**)&pUnk);
    if (!pUnk) return;

    IDispatch* pBusAlt = nullptr;
    HRESULT hr = pUnk->QueryInterface(IID_ITopologyBus, (void**)&pBusAlt);
    if (SUCCEEDED(hr) && pBusAlt && pBusAlt != pDisp)
    {
        ProbeDispatch(pBusAlt, L"ITopologyBus", label, busPropNames, numBusProps, 0, 60);
        pBusAlt->Release();
    }
    else if (pBusAlt)
        pBusAlt->Release();

    pUnk->Release();
}

// Get path object from topology object (DISPID 4, flags=0)
IUnknown* DispatchGetPath(IDispatch* pDisp)
{
    if (!pDisp) return nullptr;
    VARIANT argFlags;
    VariantInit(&argFlags);
    argFlags.vt = VT_I4;
    argFlags.lVal = 0;
    DISPPARAMS dp = { &argFlags, nullptr, 1, 0 };
    VARIANT result;
    VariantInit(&result);
    HRESULT hr = pDisp->Invoke(4, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);
    if (FAILED(hr)) { VariantClear(&result); return nullptr; }
    IUnknown* pResult = nullptr;
    if (result.vt == VT_DISPATCH && result.pdispVal)
    {
        pResult = (IUnknown*)result.pdispVal;
        pResult->AddRef();
    }
    else if (result.vt == VT_UNKNOWN && result.punkVal)
    {
        pResult = result.punkVal;
        pResult->AddRef();
    }
    VariantClear(&result);
    return pResult;
}
