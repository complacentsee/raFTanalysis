#include "RSLinxHook_fwd.h"
#include "Logging.h"
#include "ComInterfaces.h"

// ============================================================
// Dead code — diagnostic functions not called in production
// Kept for reference / future debugging use
// ============================================================

// Probe collection for available DISPIDs (diagnostic)
void ProbeCollectionDISPIDs(IDispatch* pCollection)
{
    if (!pCollection) return;

    // Try GetIDsOfNames for common collection method names
    const wchar_t* names[] = { L"Item", L"GetObject", L"Object", L"Count", L"_NewEnum" };
    for (int i = 0; i < 5; i++)
    {
        LPOLESTR name = const_cast<LPOLESTR>(names[i]);
        DISPID dispid = -999;
        HRESULT hr = pCollection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (SUCCEEDED(hr))
            Log(L"[DISP] Collection has \"%s\" = DISPID %d", names[i], dispid);
        else
            Log(L"[DISP] Collection missing \"%s\" (hr=0x%08x)", names[i], hr);
    }

    // Also try direct Invoke on a few DISPIDs with no args to see what exists
    for (DISPID d = -4; d <= 5; d++)
    {
        DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
        VARIANT result;
        VariantInit(&result);
        HRESULT hr = pCollection->Invoke(d, IID_NULL, LOCALE_USER_DEFAULT,
                                          DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);
        if (hr != 0x80020003)  // Only log if not MEMBERNOTFOUND
            Log(L"[DISP] Collection Invoke(DISPID %d, 0 args): hr=0x%08x vt=%d", d, hr, result.vt);
        VariantClear(&result);
    }
}
