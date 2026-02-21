/**
 * TopologyBrowser.cpp
 *
 * Implementation of RSLinx topology browser with event sink.
 */

#include <winsock2.h>  // Must come before windows.h (via TopologyBrowser.h)
#include "TopologyBrowser.h"
#include <iostream>
#include <iomanip>
#include <chrono>
#include <thread>
#include <atlbase.h>
#include <fstream>
#include <sstream>

// Helper: Try calling IOnlineEnumerator::Start() at a specific vtable slot with SEH protection
// Must be a standalone C-style function (no C++ objects with destructors) for __try/__except
static HRESULT TryStartAtVtableSlot(void* pInterface, IUnknown* pPath, int slot)
{
    typedef HRESULT (__stdcall *StartFunc)(void* pThis, IUnknown* pPath);
    HRESULT hr = E_FAIL;

    __try
    {
        void** vtable = *(void***)pInterface;
        StartFunc pfnStart = (StartFunc)vtable[slot];
        hr = pfnStart(pInterface, pPath);
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        hr = E_UNEXPECTED;
    }

    return hr;
}

// Helper: Try calling ConnectNewDevice at vtable slot 19 with SEH protection
// ConnectNewDevice(flags, Type, &port, Name, portLabel, address)
static HRESULT TryConnectNewDevice(void* pBusDisp, long flags, BSTR bstrType,
                                     VARIANT* pPort, VARIANT* pName, VARIANT* pPortLabel, VARIANT* pAddress)
{
    typedef HRESULT (__stdcall *ConnectNewDeviceFunc)(
        void* pThis, long flags, BSTR Type,
        VARIANT* port, VARIANT Name, VARIANT portLabel, VARIANT address);

    HRESULT hr = E_FAIL;

    __try
    {
        void** vtable = *(void***)pBusDisp;
        ConnectNewDeviceFunc pfnConnect = (ConnectNewDeviceFunc)vtable[19];
        hr = pfnConnect(pBusDisp, flags, bstrType, pPort, *pName, *pPortLabel, *pAddress);
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        hr = E_UNEXPECTED;
    }

    return hr;
}

TopologyBrowser::TopologyBrowser()
    : m_initialized(false)
    , m_pHarmonyConnector(nullptr)
    , m_pTopologyGlobals(nullptr)
    , m_pProjectGlobal(nullptr)
    , m_pProject(nullptr)
    , m_pWorkstation(nullptr)
{
}

TopologyBrowser::~TopologyBrowser()
{
    Uninitialize();
}

bool TopologyBrowser::Initialize()
{
    if (m_initialized)
        return true;

    // Initialize COM
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE)
    {
        SetLastError(L"CoInitializeEx failed");
        return false;
    }

    if (!InitializeHarmonyServices())
        return false;

    if (!InitializeTopologyGlobals())
        return false;

    if (!GetWorkstation())
        return false;

    m_initialized = true;
    return true;
}

void TopologyBrowser::Uninitialize()
{
    ReleaseAll();
    CoUninitialize();
    m_initialized = false;
}

bool TopologyBrowser::BrowseDriver(const std::wstring& driverName, std::vector<std::wstring>& discoveredDevices, DWORD timeoutMs)
{
    if (!m_initialized)
    {
        SetLastError(L"Not initialized");
        return false;
    }

    std::wcout << L"========================================" << std::endl;
    std::wcout << L"Browsing driver: " << driverName << std::endl;
    std::wcout << L"========================================" << std::endl;

    HRESULT hr = S_OK;

    // Step 1: Get the bus object via vtable navigation
    std::wcout << L"[INFO] Getting bus object for driver: " << driverName << std::endl;
    IUnknown* pBus = GetBusObjectByName(driverName);
    if (!pBus)
    {
        SetLastError(L"Failed to get bus object");
        return false;
    }
    std::wcout << L"[OK] Got bus object" << std::endl;

    // Step 2: Connect event sink BEFORE starting browse
    BrowseEventSink* pSink = new BrowseEventSink();
    DWORD cookie = 0;
    bool sinkConnected = false;

    if (ConnectEventSink(pBus, pSink, &cookie))
    {
        std::wcout << L"[OK] Event sink connected to bus (cookie: " << cookie << L")" << std::endl;
        sinkConnected = true;
    }
    else
    {
        std::wcout << L"[WARN] Event sink connection to bus failed" << std::endl;
    }

    // Step 3: Get the bus's PATH OBJECT
    // TypeLib: ITopologyBus DISPID=4 PROPGET path(flags) -> IUnknown*
    // This is what IOnlineEnumerator::Start() needs - a path OBJECT, not a string!
    //
    // KEY INSIGHT: MFC COM objects expose different DISPIDs on different dual interfaces.
    // The bus's *default* IDispatch might not be ITopologyBus. We must QI for the
    // specific dual interface IID to get the correct IDispatch with the right DISPIDs.
    std::wcout << L"[INFO] Getting bus path object..." << std::endl;

    IUnknown* pPathObject = nullptr;

    // ---- Attempt 1: QI bus for ITopologyBus (dual interface) and use its IDispatch ----
    {
        IDispatch* pBusDisp = nullptr;
        // QI for ITopologyBus IID - since it inherits from IDispatch, this gives us
        // the ITopologyBus-specific IDispatch with its DISPIDs
        hr = pBus->QueryInterface(IID_ITopologyBus, (void**)&pBusDisp);
        if (SUCCEEDED(hr) && pBusDisp)
        {
            std::wcout << L"[OK] Got ITopologyBus dual interface" << std::endl;

            // First try: DISPID 4 with NO args (PROPGET may not need flags parameter)
            DISPPARAMS noArgs = { nullptr, nullptr, 0, 0 };
            VARIANT varPath;
            VariantInit(&varPath);
            EXCEPINFO excep = {};
            UINT argErr = 0;

            hr = pBusDisp->Invoke(4, IID_NULL, LOCALE_USER_DEFAULT,
                                   DISPATCH_PROPERTYGET, &noArgs, &varPath, &excep, &argErr);
            std::wcout << L"  ITopologyBus DISPID 4 (no args): hr=0x" << std::hex << hr
                       << L" vt=" << std::dec << varPath.vt << std::endl;

            if (FAILED(hr))
            {
                VariantClear(&varPath);

                // Second try: DISPID 4 with flags=0 argument
                VARIANT argFlags;
                VariantInit(&argFlags);
                argFlags.vt = VT_I4;
                argFlags.lVal = 0;
                DISPPARAMS dpFlags = { &argFlags, nullptr, 1, 0 };
                VariantInit(&varPath);

                hr = pBusDisp->Invoke(4, IID_NULL, LOCALE_USER_DEFAULT,
                                       DISPATCH_PROPERTYGET, &dpFlags, &varPath, &excep, &argErr);
                std::wcout << L"  ITopologyBus DISPID 4 (flags=0): hr=0x" << std::hex << hr
                           << L" vt=" << std::dec << varPath.vt << std::endl;
            }

            if (SUCCEEDED(hr))
            {
                std::wcout << L"[OK] Bus path returned vt=" << varPath.vt << std::endl;
                if (varPath.vt == VT_UNKNOWN && varPath.punkVal)
                {
                    pPathObject = varPath.punkVal;
                    pPathObject->AddRef();
                }
                else if (varPath.vt == VT_DISPATCH && varPath.pdispVal)
                {
                    pPathObject = varPath.pdispVal;
                    pPathObject->AddRef();
                }
                if (pPathObject)
                    std::wcout << L"[OK] Got path object: 0x" << std::hex << (void*)pPathObject << std::dec << std::endl;
            }

            VariantClear(&varPath);

            // Also try GetIDsOfNames to discover what DISPIDs are available
            if (!pPathObject)
            {
                LPOLESTR names[] = { const_cast<LPOLESTR>(L"path"), const_cast<LPOLESTR>(L"Path"),
                                     const_cast<LPOLESTR>(L"name"), const_cast<LPOLESTR>(L"Name") };
                for (int n = 0; n < 4; n++)
                {
                    DISPID did;
                    hr = pBusDisp->GetIDsOfNames(IID_NULL, &names[n], 1, LOCALE_USER_DEFAULT, &did);
                    if (SUCCEEDED(hr))
                        std::wcout << L"  GetIDsOfNames(\"" << names[n] << L"\") = DISPID " << did << std::endl;
                }
            }

            pBusDisp->Release();
        }
        else
        {
            std::wcout << L"[WARN] QI for ITopologyBus failed: hr=0x" << std::hex << hr << std::dec << std::endl;
        }
    }

    // ---- Attempt 2: QI for IRSTopologyObject and use its path property ----
    if (!pPathObject)
    {
        IDispatch* pObjDisp = nullptr;
        hr = pBus->QueryInterface(IID_IRSTopologyObject, (void**)&pObjDisp);
        if (SUCCEEDED(hr) && pObjDisp)
        {
            std::wcout << L"[OK] Got IRSTopologyObject dual interface" << std::endl;

            // IRSTopologyObject also has path at DISPID 4
            // Try with no args first
            DISPPARAMS noArgs = { nullptr, nullptr, 0, 0 };
            VARIANT varPath;
            VariantInit(&varPath);

            hr = pObjDisp->Invoke(4, IID_NULL, LOCALE_USER_DEFAULT,
                                   DISPATCH_PROPERTYGET, &noArgs, &varPath, nullptr, nullptr);
            std::wcout << L"  IRSTopologyObject DISPID 4 (no args): hr=0x" << std::hex << hr
                       << L" vt=" << std::dec << varPath.vt << std::endl;

            if (FAILED(hr))
            {
                VariantClear(&varPath);

                // Try with flags=0
                VARIANT argFlags;
                VariantInit(&argFlags);
                argFlags.vt = VT_I4;
                argFlags.lVal = 0;
                DISPPARAMS dpFlags = { &argFlags, nullptr, 1, 0 };
                VariantInit(&varPath);

                hr = pObjDisp->Invoke(4, IID_NULL, LOCALE_USER_DEFAULT,
                                       DISPATCH_PROPERTYGET, &dpFlags, &varPath, nullptr, nullptr);
                std::wcout << L"  IRSTopologyObject DISPID 4 (flags=0): hr=0x" << std::hex << hr
                           << L" vt=" << std::dec << varPath.vt << std::endl;
            }

            if (SUCCEEDED(hr))
            {
                if (varPath.vt == VT_UNKNOWN && varPath.punkVal)
                {
                    pPathObject = varPath.punkVal;
                    pPathObject->AddRef();
                }
                else if (varPath.vt == VT_DISPATCH && varPath.pdispVal)
                {
                    pPathObject = varPath.pdispVal;
                    pPathObject->AddRef();
                }
                if (pPathObject)
                    std::wcout << L"[OK] Got path from IRSTopologyObject: 0x" << std::hex << (void*)pPathObject << std::dec << std::endl;
            }

            VariantClear(&varPath);
            pObjDisp->Release();
        }
        else
        {
            std::wcout << L"[WARN] QI for IRSTopologyObject failed: hr=0x" << std::hex << hr << std::dec << std::endl;
        }
    }

    // ---- Attempt 3: Generic IDispatch with no args ----
    if (!pPathObject)
    {
        IDispatch* pBusDisp = nullptr;
        hr = pBus->QueryInterface(IID_IDispatch, (void**)&pBusDisp);
        if (SUCCEEDED(hr) && pBusDisp)
        {
            std::wcout << L"[INFO] Trying default IDispatch..." << std::endl;

            // Try DISPID 4 with no args
            DISPPARAMS noArgs = { nullptr, nullptr, 0, 0 };
            VARIANT varPath;
            VariantInit(&varPath);
            hr = pBusDisp->Invoke(4, IID_NULL, LOCALE_USER_DEFAULT,
                                   DISPATCH_PROPERTYGET, &noArgs, &varPath, nullptr, nullptr);
            std::wcout << L"  Default IDispatch DISPID 4 (no args): hr=0x" << std::hex << hr
                       << L" vt=" << std::dec << varPath.vt << std::endl;

            if (SUCCEEDED(hr) && (varPath.vt == VT_UNKNOWN || varPath.vt == VT_DISPATCH))
            {
                pPathObject = (varPath.vt == VT_UNKNOWN) ? varPath.punkVal : (IUnknown*)varPath.pdispVal;
                if (pPathObject) pPathObject->AddRef();
            }
            VariantClear(&varPath);

            // Also enumerate DISPIDs by name on the default IDispatch
            if (!pPathObject)
            {
                LPOLESTR tryNames[] = {
                    const_cast<LPOLESTR>(L"path"), const_cast<LPOLESTR>(L"Path"),
                    const_cast<LPOLESTR>(L"GetPath"), const_cast<LPOLESTR>(L"name"),
                    const_cast<LPOLESTR>(L"Name"), const_cast<LPOLESTR>(L"Busses"),
                    const_cast<LPOLESTR>(L"Ports"), const_cast<LPOLESTR>(L"Devices")
                };
                for (int n = 0; n < 8; n++)
                {
                    DISPID did;
                    hr = pBusDisp->GetIDsOfNames(IID_NULL, &tryNames[n], 1, LOCALE_USER_DEFAULT, &did);
                    if (SUCCEEDED(hr))
                    {
                        std::wcout << L"  Default.GetIDsOfNames(\"" << tryNames[n] << L"\") = DISPID " << did << std::endl;

                        // If we found "path" or "Path", try calling it
                        if (_wcsicmp(tryNames[n], L"path") == 0 || _wcsicmp(tryNames[n], L"GetPath") == 0)
                        {
                            DISPPARAMS noA = { nullptr, nullptr, 0, 0 };
                            VARIANT vp;
                            VariantInit(&vp);
                            HRESULT hr2 = pBusDisp->Invoke(did, IID_NULL, LOCALE_USER_DEFAULT,
                                                            DISPATCH_PROPERTYGET, &noA, &vp, nullptr, nullptr);
                            std::wcout << L"    Invoke DISPID " << did << L": hr=0x" << std::hex << hr2
                                       << L" vt=" << std::dec << vp.vt << std::endl;
                            if (SUCCEEDED(hr2) && (vp.vt == VT_UNKNOWN || vp.vt == VT_DISPATCH))
                            {
                                pPathObject = (vp.vt == VT_UNKNOWN) ? vp.punkVal : (IUnknown*)vp.pdispVal;
                                if (pPathObject) pPathObject->AddRef();
                            }
                            VariantClear(&vp);
                            if (pPathObject) break;
                        }
                    }
                }
            }

            pBusDisp->Release();
        }
    }

    // ---- Attempt 4: Create path via ITopologyPathComposer (QI for specific dual interface) ----
    if (!pPathObject)
    {
        std::wcout << L"[INFO] Fallback: creating path via ITopologyPathComposer..." << std::endl;

        IUnknown* pPathComposerUnk = nullptr;
        hr = CoCreateInstance(CLSID_RSPath, NULL, CLSCTX_ALL, IID_IUnknown, (void**)&pPathComposerUnk);
        if (SUCCEEDED(hr) && pPathComposerUnk)
        {
            // CRITICAL: QI for ITopologyPathComposer IID specifically, NOT generic IDispatch.
            // MFC COM objects expose different DISPIDs on different dual interfaces.
            // The default IDispatch may be ITopologyPath which doesn't have ParseString.
            IDispatch* pPathComposerDisp = nullptr;
            hr = pPathComposerUnk->QueryInterface(IID_ITopologyPathComposer, (void**)&pPathComposerDisp);
            std::wcout << L"  QI for ITopologyPathComposer: hr=0x" << std::hex << hr << std::dec << std::endl;

            if (FAILED(hr) || !pPathComposerDisp)
            {
                // Fallback to generic IDispatch
                hr = pPathComposerUnk->QueryInterface(IID_IDispatch, (void**)&pPathComposerDisp);
                std::wcout << L"  Fallback to generic IDispatch: hr=0x" << std::hex << hr << std::dec << std::endl;
            }

            if (SUCCEEDED(hr) && pPathComposerDisp)
            {
                // Discover available methods
                LPOLESTR pcNames[] = {
                    const_cast<LPOLESTR>(L"ParseString"), const_cast<LPOLESTR>(L"Start"),
                    const_cast<LPOLESTR>(L"path"), const_cast<LPOLESTR>(L"Path"),
                    const_cast<LPOLESTR>(L"Compose"), const_cast<LPOLESTR>(L"SetPath")
                };
                for (int n = 0; n < 6; n++)
                {
                    DISPID did;
                    HRESULT hrName = pPathComposerDisp->GetIDsOfNames(IID_NULL, &pcNames[n], 1, LOCALE_USER_DEFAULT, &did);
                    if (SUCCEEDED(hrName))
                        std::wcout << L"  PathComposer.GetIDsOfNames(\"" << pcNames[n] << L"\") = DISPID " << did << std::endl;
                }

                // Set Start property to workstation (DISPID 3 PROPPUT)
                DISPID namedArgPut = DISPID_PROPERTYPUT;
                VARIANT argWs;
                VariantInit(&argWs);
                argWs.vt = VT_UNKNOWN;
                argWs.punkVal = m_pWorkstation;
                DISPPARAMS dpPut = { &argWs, &namedArgPut, 1, 1 };

                hr = pPathComposerDisp->Invoke(3, IID_NULL, LOCALE_USER_DEFAULT,
                                                DISPATCH_PROPERTYPUT, &dpPut, nullptr, nullptr, nullptr);
                std::wcout << L"  Set Start to workstation: hr=0x" << std::hex << hr << std::dec << std::endl;

                // Call ParseString (DISPID 17) with path string
                std::wstring pathStr = std::wstring(L"CORRAEWS01/") + driverName;
                VARIANT argPathStr;
                VariantInit(&argPathStr);
                argPathStr.vt = VT_BSTR;
                argPathStr.bstrVal = SysAllocString(pathStr.c_str());
                DISPPARAMS dpParse = { &argPathStr, nullptr, 1, 0 };

                hr = pPathComposerDisp->Invoke(17, IID_NULL, LOCALE_USER_DEFAULT,
                                                DISPATCH_METHOD, &dpParse, nullptr, nullptr, nullptr);
                std::wcout << L"  ParseString(\"" << pathStr << L"\"): hr=0x" << std::hex << hr << std::dec << std::endl;

                if (FAILED(hr))
                {
                    // Try with just the driver name
                    VariantClear(&argPathStr);
                    argPathStr.vt = VT_BSTR;
                    argPathStr.bstrVal = SysAllocString(driverName.c_str());
                    DISPPARAMS dpParse2 = { &argPathStr, nullptr, 1, 0 };

                    hr = pPathComposerDisp->Invoke(17, IID_NULL, LOCALE_USER_DEFAULT,
                                                    DISPATCH_METHOD, &dpParse2, nullptr, nullptr, nullptr);
                    std::wcout << L"  ParseString(\"" << driverName << L"\"): hr=0x" << std::hex << hr << std::dec << std::endl;
                }

                VariantClear(&argPathStr);

                if (SUCCEEDED(hr))
                {
                    // The path composer itself IS the path object (implements ITopologyPath + IMoniker)
                    pPathObject = pPathComposerUnk;
                    pPathObject->AddRef();
                    std::wcout << L"[OK] Created path object via ParseString" << std::endl;
                }

                pPathComposerDisp->Release();
            }
            pPathComposerUnk->Release();
        }
    }

    // ---- Attempt 5: Just pass the bus object itself as the path ----
    // Some COM objects accept the bus IUnknown directly as the "path"
    if (!pPathObject)
    {
        std::wcout << L"[INFO] Last resort: using bus object itself as path argument..." << std::endl;
        pPathObject = pBus;
        pPathObject->AddRef();
    }

    if (!pPathObject)
    {
        std::wcout << L"[FAIL] Could not obtain path object for bus" << std::endl;
        if (sinkConnected) DisconnectEventSink(pBus, cookie);
        pSink->Release();
        pBus->Release();
        SetLastError(L"Failed to get bus path object");
        return false;
    }

    // Step 4: Create IOnlineEnumerator and call Start() with the path OBJECT
    // CRITICAL: Start() takes IUnknown* path, NOT LPCWSTR!
    // CRITICAL: IOnlineEnumerator inherits from IDispatch, not IUnknown!
    //
    // IDispatch::Invoke fails with TYPE_E_LIBNOTREGISTERED (0x8002801d) because
    // LINXCOMM.DLL's type library isn't registered. Fix: register RSTOP.DLL TypeLib first.
    // Register type libraries so IDispatch::Invoke can find interface definitions
    std::wcout << L"[INFO] Registering type libraries..." << std::endl;
    ITypeLib* pRstopTypeLib = nullptr;  // Keep reference for ITypeInfo::Invoke fallback
    {
        const wchar_t* tlbPaths[] = {
            L"C:\\Program Files (x86)\\Rockwell Software\\RSCommon\\RSTOP.DLL",
            L"C:\\Program Files (x86)\\Rockwell Software\\RSLinx\\LINXCOMM.DLL",
            L"C:\\Program Files (x86)\\Rockwell Software\\RSCommon\\RSWHO.OCX",
        };
        for (int i = 0; i < 3; i++)
        {
            ITypeLib* pTL = nullptr;
            hr = LoadTypeLibEx(tlbPaths[i], REGKIND_REGISTER, &pTL);
            std::wcout << L"  " << tlbPaths[i] << L": hr=0x" << std::hex << hr << std::dec << std::endl;
            if (SUCCEEDED(hr) && pTL)
            {
                if (i == 0) { pRstopTypeLib = pTL; pRstopTypeLib->AddRef(); }
                pTL->Release();
            }
        }
    }

    std::wcout << L"[INFO] Getting IOnlineEnumerator..." << std::endl;

    IUnknown* pEnumUnk = nullptr;
    bool browseStarted = false;

    // Try creating enumerator in RSLinx's process (LOCAL_SERVER) so it connects to the engine
    hr = CoCreateInstance(CLSID_OnlineBusExt, NULL, CLSCTX_LOCAL_SERVER, IID_IUnknown, (void**)&pEnumUnk);
    if (SUCCEEDED(hr) && pEnumUnk)
    {
        std::wcout << L"[OK] Created OnlineBusExt via LOCAL_SERVER (in RSLinx process)" << std::endl;
    }
    else
    {
        std::wcout << L"[INFO] LOCAL_SERVER failed: hr=0x" << std::hex << hr << std::dec
                  << L", falling back to CLSCTX_ALL" << std::endl;
        hr = CoCreateInstance(CLSID_OnlineBusExt, NULL, CLSCTX_ALL, IID_IUnknown, (void**)&pEnumUnk);
        if (SUCCEEDED(hr))
            std::wcout << L"[OK] Created standalone OnlineBusExt enumerator" << std::endl;
    }
    if (SUCCEEDED(hr) && pEnumUnk)
    {
        std::wcout << L"[OK] Created OnlineBusExt enumerator" << std::endl;

        // Connect event sink to enumerator's connection points
        IConnectionPointContainer* pEnumCPC = nullptr;
        hr = pEnumUnk->QueryInterface(IID_IConnectionPointContainer, (void**)&pEnumCPC);
        if (SUCCEEDED(hr) && pEnumCPC)
        {
            IEnumConnectionPoints* pEnumCPEnum = nullptr;
            hr = pEnumCPC->EnumConnectionPoints(&pEnumCPEnum);
            if (SUCCEEDED(hr) && pEnumCPEnum)
            {
                IConnectionPoint* pCP = nullptr;
                ULONG fetched = 0;
                int cpIndex = 0;
                while (pEnumCPEnum->Next(1, &pCP, &fetched) == S_OK && fetched > 0)
                {
                    DWORD enumCookie = 0;
                    HRESULT hrAdv = pCP->Advise(pSink, &enumCookie);
                    GUID cpIID;
                    pCP->GetConnectionInterface(&cpIID);
                    wchar_t guidStr[64];
                    StringFromGUID2(cpIID, guidStr, 64);
                    std::wcout << L"  Enum CP[" << cpIndex << L"] " << guidStr
                              << (SUCCEEDED(hrAdv) ? L" connected" : L" failed") << std::endl;
                    pCP->Release();
                    cpIndex++;
                }
                pEnumCPEnum->Release();
            }
            pEnumCPC->Release();
        }

        // Probe all IIDs to see what pointers we get back
        void* pPtrUnk = nullptr;
        void* pPtrDisp = nullptr;
        void* pPtrTypeLib = nullptr;
        void* pPtrAPImon = nullptr;

        pEnumUnk->QueryInterface(IID_IUnknown, &pPtrUnk);
        pEnumUnk->QueryInterface(IID_IDispatch, &pPtrDisp);
        pEnumUnk->QueryInterface(IID_IOnlineEnumeratorTypeLib, &pPtrTypeLib);
        pEnumUnk->QueryInterface(IID_IOnlineEnumerator, &pPtrAPImon);

        std::wcout << L"  IUnknown:     0x" << std::hex << pPtrUnk << std::dec << std::endl;
        std::wcout << L"  IDispatch:    0x" << std::hex << pPtrDisp << std::dec << std::endl;
        std::wcout << L"  TypeLib IID:  0x" << std::hex << pPtrTypeLib << std::dec << std::endl;
        std::wcout << L"  APImon IID:   0x" << std::hex << pPtrAPImon << std::dec << std::endl;

        // Dump vtable pointers for each non-null interface
        void** interfaces[] = { (void**)pPtrDisp, (void**)pPtrTypeLib, (void**)pPtrAPImon };
        const wchar_t* ifaceNames[] = { L"IDispatch", L"TypeLib", L"APImon" };
        for (int i = 0; i < 3; i++)
        {
            if (interfaces[i])
            {
                void** vt = *(void***)interfaces[i];
                std::wcout << L"  " << ifaceNames[i] << L" vtable: ";
                for (int s = 0; s < 10; s++)
                    std::wcout << L"[" << s << L"]=0x" << std::hex << vt[s] << L" ";
                std::wcout << std::dec << std::endl;
            }
        }

        if (pPtrUnk) ((IUnknown*)pPtrUnk)->Release();
        if (pPtrDisp) ((IUnknown*)pPtrDisp)->Release();
        if (pPtrTypeLib) ((IUnknown*)pPtrTypeLib)->Release();
        if (pPtrAPImon) ((IUnknown*)pPtrAPImon)->Release();

        // QI for the IOnlineEnumerator dual interface
        IDispatch* pEnumDisp = nullptr;
        hr = pEnumUnk->QueryInterface(IID_IOnlineEnumeratorTypeLib, (void**)&pEnumDisp);
        std::wcout << L"  QI for IOnlineEnumerator (TypeLib): hr=0x" << std::hex << hr << std::dec << std::endl;

        if (FAILED(hr) || !pEnumDisp)
        {
            hr = pEnumUnk->QueryInterface(IID_IOnlineEnumerator, (void**)&pEnumDisp);
            std::wcout << L"  QI for IOnlineEnumerator (APImon): hr=0x" << std::hex << hr << std::dec << std::endl;
        }

        if (FAILED(hr) || !pEnumDisp)
        {
            hr = pEnumUnk->QueryInterface(IID_IDispatch, (void**)&pEnumDisp);
            std::wcout << L"  QI for generic IDispatch: hr=0x" << std::hex << hr << std::dec << std::endl;
        }

        if (SUCCEEDED(hr) && pEnumDisp)
        {
            // Try IDispatch::Invoke first (should work now that TypeLib is registered)
            VARIANT argPath;
            VariantInit(&argPath);
            argPath.vt = VT_UNKNOWN;
            argPath.punkVal = pPathObject;
            DISPPARAMS dpStart = { &argPath, nullptr, 1, 0 };
            EXCEPINFO excep = {};
            UINT argErr = 0;

            // Discover DISPIDs
            LPOLESTR enumNames[] = {
                const_cast<LPOLESTR>(L"Start"), const_cast<LPOLESTR>(L"Stop"),
                const_cast<LPOLESTR>(L"Browse"), const_cast<LPOLESTR>(L"Run")
            };
            for (int n = 0; n < 4; n++)
            {
                DISPID did;
                HRESULT hrN = pEnumDisp->GetIDsOfNames(IID_NULL, &enumNames[n], 1, LOCALE_USER_DEFAULT, &did);
                if (SUCCEEDED(hrN))
                    std::wcout << L"  Enum.GetIDsOfNames(\"" << enumNames[n] << L"\") = DISPID " << did << std::endl;
            }

            std::wcout << L"[INFO] Calling IOnlineEnumerator.Start(pathObject) DISPID=1..." << std::endl;
            hr = pEnumDisp->Invoke(1, IID_NULL, LOCALE_USER_DEFAULT,
                                    DISPATCH_METHOD, &dpStart, nullptr, &excep, &argErr);
            std::wcout << L"  Start result: hr=0x" << std::hex << hr << std::dec << std::endl;

            if (SUCCEEDED(hr))
            {
                browseStarted = true;
                std::wcout << L"[OK] Browse started via IDispatch!" << std::endl;
            }
            else
            {
                if (excep.bstrDescription)
                {
                    std::wcout << L"  Exception: " << excep.bstrDescription << std::endl;
                    SysFreeString(excep.bstrDescription);
                }
                SysFreeString(excep.bstrSource);
                SysFreeString(excep.bstrHelpFile);

                // Try VT_DISPATCH instead of VT_UNKNOWN
                IDispatch* pPathDisp = nullptr;
                HRESULT hrQI = pPathObject->QueryInterface(IID_IDispatch, (void**)&pPathDisp);
                if (SUCCEEDED(hrQI) && pPathDisp)
                {
                    VariantInit(&argPath);
                    argPath.vt = VT_DISPATCH;
                    argPath.pdispVal = pPathDisp;
                    DISPPARAMS dpStart2 = { &argPath, nullptr, 1, 0 };
                    memset(&excep, 0, sizeof(excep));

                    std::wcout << L"[INFO] Retrying Start with VT_DISPATCH..." << std::endl;
                    hr = pEnumDisp->Invoke(1, IID_NULL, LOCALE_USER_DEFAULT,
                                            DISPATCH_METHOD, &dpStart2, nullptr, &excep, &argErr);
                    std::wcout << L"  Start (VT_DISPATCH): hr=0x" << std::hex << hr << std::dec << std::endl;

                    if (SUCCEEDED(hr))
                    {
                        browseStarted = true;
                        std::wcout << L"[OK] Browse started!" << std::endl;
                    }
                    else if (excep.bstrDescription)
                    {
                        std::wcout << L"  Exception: " << excep.bstrDescription << std::endl;
                        SysFreeString(excep.bstrDescription);
                    }
                    SysFreeString(excep.bstrSource);
                    SysFreeString(excep.bstrHelpFile);
                    pPathDisp->Release();
                }

                // Call Start(pathObject) via vtable[7] on the real IOnlineEnumerator
                // (APImon IID {91748520} which has a distinct vtable from IDispatch)
                if (!browseStarted)
                {
                    void* pRealEnum = nullptr;
                    hr = pEnumUnk->QueryInterface(IID_IOnlineEnumerator, &pRealEnum);
                    if (SUCCEEDED(hr) && pRealEnum)
                    {
                        std::wcout << L"[INFO] Calling Start(pathObject) vtable[7] on real IOnlineEnumerator 0x"
                                  << std::hex << pRealEnum << std::dec << std::endl;

                        hr = TryStartAtVtableSlot(pRealEnum, pPathObject, 7);
                        std::wcout << L"  Result: hr=0x" << std::hex << hr << std::dec << std::endl;

                        if (SUCCEEDED(hr) && hr != E_UNEXPECTED)
                        {
                            browseStarted = true;
                            std::wcout << L"[OK] Browse started!" << std::endl;
                        }
                        else if (hr == E_UNEXPECTED)
                        {
                            std::wcout << L"  SEH exception caught" << std::endl;
                        }

                        ((IUnknown*)pRealEnum)->Release();
                    }
                }
            }

            pEnumDisp->Release();
        }
    }
    else
    {
        std::wcout << L"[FAIL] Cannot create OnlineBusExt: hr=0x" << std::hex << hr << std::dec << std::endl;
    }

    if (!browseStarted)
    {
        std::wcout << L"[FAIL] Could not start browse" << std::endl;
        if (pEnumUnk) pEnumUnk->Release();
        pPathObject->Release();
        if (sinkConnected) DisconnectEventSink(pBus, cookie);
        pSink->Release();
        pBus->Release();
        SetLastError(L"Failed to start browse");
        return false;
    }

    // Step 5: Wait for devices - pump COM messages and poll XML
    std::wcout << L"Waiting for devices (" << timeoutMs / 1000 << L" seconds)..." << std::endl;

    auto startTime = std::chrono::steady_clock::now();
    DWORD elapsed = 0;
    DWORD lastPollTime = 0;
    const DWORD pollInterval = 5000;

    while (elapsed < timeoutMs)
    {
        MSG msg;
        while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }

        if (sinkConnected && pSink->IsBrowseEnded())
        {
            std::wcout << L"[OK] Browse ended naturally via event" << std::endl;
            break;
        }

        if (elapsed - lastPollTime >= pollInterval)
        {
            std::wstring pollFile = L"C:\\temp\\topology_poll.xml";
            if (SaveTopologyXML(pollFile))
            {
                std::vector<std::wstring> xmlDevices = ParseDevicesFromXML(pollFile, driverName);
                std::wcout << L"  [POLL " << elapsed / 1000 << L"s] "
                          << xmlDevices.size() << L" device(s) in topology XML";
                if (sinkConnected)
                    std::wcout << L", " << pSink->GetDiscoveredDevices().size() << L" via events";
                std::wcout << std::endl;
                for (const auto& dev : xmlDevices)
                    std::wcout << L"    XML device: " << dev << std::endl;
            }
            lastPollTime = elapsed;
        }

        Sleep(100);

        auto now = std::chrono::steady_clock::now();
        elapsed = static_cast<DWORD>(std::chrono::duration_cast<std::chrono::milliseconds>(now - startTime).count());
    }

    std::wcout << L"[OK] Browse wait complete" << std::endl;

    // Step 6: Collect results
    discoveredDevices = pSink->GetDiscoveredDevices();

    std::wstring finalXml = L"C:\\temp\\topology_final.xml";
    if (SaveTopologyXML(finalXml))
    {
        std::vector<std::wstring> xmlDevices = ParseDevicesFromXML(finalXml, driverName);
        for (const auto& xmlDev : xmlDevices)
        {
            bool found = false;
            for (const auto& existing : discoveredDevices)
            {
                if (existing == xmlDev) { found = true; break; }
            }
            if (!found) discoveredDevices.push_back(xmlDev);
        }
    }

    std::wcout << L"========================================" << std::endl;
    std::wcout << L"Discovered " << discoveredDevices.size() << L" device(s):" << std::endl;
    for (size_t i = 0; i < discoveredDevices.size(); i++)
        std::wcout << L"  " << i + 1 << L". " << discoveredDevices[i] << std::endl;
    std::wcout << L"========================================" << std::endl;

    // Cleanup
    if (pRstopTypeLib) pRstopTypeLib->Release();
    if (pEnumUnk) pEnumUnk->Release();
    pPathObject->Release();
    if (sinkConnected) DisconnectEventSink(pBus, cookie);
    pSink->Release();
    pBus->Release();

    return true;
}

std::vector<std::wstring> TopologyBrowser::ParseDevicesFromXML(const std::wstring& xmlPath, const std::wstring& busName)
{
    std::vector<std::wstring> devices;

    // Read XML file
    std::wifstream file(xmlPath);
    if (!file.is_open())
        return devices;

    std::wstring xml((std::istreambuf_iterator<wchar_t>(file)),
                      std::istreambuf_iterator<wchar_t>());
    file.close();

    // Find the bus section matching our driver name
    // XML format: <bus name="Test" ...> <device name="..." ...> <address ...>10.39.29.132</address>
    // Look for devices within the target bus

    // Find the bus tag
    std::wstring busTag = L"<bus name=\"" + busName + L"\"";
    size_t busPos = xml.find(busTag);
    if (busPos == std::wstring::npos)
        return devices;

    // Find the closing </bus> tag
    size_t busEnd = xml.find(L"</bus>", busPos);
    if (busEnd == std::wstring::npos)
        busEnd = xml.length();

    std::wstring busSection = xml.substr(busPos, busEnd - busPos);

    // Extract device names from within this bus section
    // Pattern: <device name="..." ...>
    size_t searchPos = 0;
    while (true)
    {
        size_t devPos = busSection.find(L"<device ", searchPos);
        if (devPos == std::wstring::npos)
            break;

        // Extract device name
        size_t namePos = busSection.find(L"name=\"", devPos);
        if (namePos == std::wstring::npos || namePos > devPos + 200)
        {
            searchPos = devPos + 1;
            continue;
        }
        namePos += 6;  // Skip past name="
        size_t nameEnd = busSection.find(L"\"", namePos);
        if (nameEnd == std::wstring::npos)
            break;

        std::wstring deviceName = busSection.substr(namePos, nameEnd - namePos);

        // Also look for address within this device's section
        size_t devEnd = busSection.find(L"</device>", devPos);
        if (devEnd == std::wstring::npos)
            devEnd = busSection.find(L"/>", devPos);

        // Look for address type="IP" or just <address> tags
        std::wstring devSection = busSection.substr(devPos, (devEnd != std::wstring::npos ? devEnd - devPos : 200));

        // Try to find address value
        size_t addrPos = devSection.find(L"<address");
        std::wstring addrInfo;
        if (addrPos != std::wstring::npos)
        {
            // Find the address value - could be attribute or text content
            size_t valPos = devSection.find(L"value=\"", addrPos);
            if (valPos != std::wstring::npos && valPos < addrPos + 100)
            {
                valPos += 7;
                size_t valEnd = devSection.find(L"\"", valPos);
                if (valEnd != std::wstring::npos)
                    addrInfo = devSection.substr(valPos, valEnd - valPos);
            }
        }

        // Build device entry
        std::wstring entry = deviceName;
        if (!addrInfo.empty())
            entry += L" (" + addrInfo + L")";
        devices.push_back(entry);

        searchPos = (devEnd != std::wstring::npos) ? devEnd + 1 : devPos + 1;
    }

    return devices;
}

bool TopologyBrowser::SaveTopologyXML(const std::wstring& filename)
{
    if (!m_pTopologyGlobals)
    {
        SetLastError(L"Topology globals not initialized");
        return false;
    }

    // Get IDispatch for SaveAsXML
    IDispatch* pDisp = nullptr;
    HRESULT hr = m_pTopologyGlobals->QueryInterface(IID_IDispatch, (void**)&pDisp);
    if (FAILED(hr))
    {
        SetLastError(L"Failed to get IDispatch");
        return false;
    }

    // Try to get DISPID for SaveAsXML
    LPOLESTR methodName = const_cast<LPOLESTR>(L"SaveAsXML");
    DISPID dispid;
    hr = pDisp->GetIDsOfNames(IID_NULL, &methodName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr))
    {
        std::wcout << L"[DEBUG] GetIDsOfNames for 'SaveAsXML' failed, hr=0x" << std::hex << hr << std::dec << std::endl;
        std::wcout << L"[DEBUG] Using hardcoded DISPID 1610743808" << std::endl;
        dispid = 1610743808;
    }
    else
    {
        std::wcout << L"[DEBUG] Found 'SaveAsXML' at DISPID " << dispid << std::endl;
    }
    // COM IDispatch parameters are RIGHT-TO-LEFT in rgvarg array
    // Method signature: SaveAsXML(path, depth, filename)
    // So: args[2]=path (first), args[1]=depth (middle), args[0]=filename (last)
    DISPPARAMS params = { 0 };
    VARIANT args[3];

    // args[0] = LAST parameter: filename
    VariantInit(&args[0]);
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(filename.c_str());

    // args[1] = MIDDLE parameter: depth
    VariantInit(&args[1]);
    args[1].vt = VT_I4;
    args[1].lVal = 100;

    // args[2] = FIRST parameter: path (empty = root)
    VariantInit(&args[2]);
    args[2].vt = VT_BSTR;
    args[2].bstrVal = SysAllocString(L"");

    params.rgvarg = args;
    params.cArgs = 3;

    VARIANT result;
    VariantInit(&result);

    EXCEPINFO excepInfo;
    memset(&excepInfo, 0, sizeof(excepInfo));
    UINT argErr = 0;

    std::wcout << L"[DEBUG] Calling SaveAsXML with DISPID " << dispid << std::endl;
    hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &result, &excepInfo, &argErr);
    std::wcout << L"[DEBUG] SaveAsXML Invoke result: hr=0x" << std::hex << hr << std::dec << std::endl;

    if (FAILED(hr) && hr == DISP_E_EXCEPTION)
    {
        std::wcerr << L"[ERROR] Exception details:" << std::endl;
        if (excepInfo.bstrDescription)
            std::wcerr << L"  Description: " << excepInfo.bstrDescription << std::endl;
        if (excepInfo.bstrSource)
            std::wcerr << L"  Source: " << excepInfo.bstrSource << std::endl;
        std::wcerr << L"  Error code: 0x" << std::hex << excepInfo.scode << std::dec << std::endl;
        SysFreeString(excepInfo.bstrDescription);
        SysFreeString(excepInfo.bstrSource);
        SysFreeString(excepInfo.bstrHelpFile);
    }
    else if (FAILED(hr) && hr == DISP_E_PARAMNOTFOUND)
    {
        std::wcerr << L"[ERROR] Parameter error at argument index: " << argErr << std::endl;
    }

    VariantClear(&args[0]);
    VariantClear(&args[1]);
    VariantClear(&args[2]);
    VariantClear(&result);
    pDisp->Release();

    if (FAILED(hr))
    {
        std::wcerr << L"[ERROR] SaveAsXML failed with hr=0x" << std::hex << hr << std::dec << std::endl;
        SetLastError(L"SaveAsXML failed");
        return false;
    }

    std::wcout << L"[OK] SaveAsXML succeeded" << std::endl;
    return true;
}

bool TopologyBrowser::AddDeviceManually(const std::wstring& driverName, const std::wstring& ipAddress, const std::wstring& deviceName)
{
    std::wcout << L"[INFO] ConnectNewDevice: " << ipAddress << L" on driver '" << driverName << L"'" << std::endl;

    if (!m_pWorkstation)
    {
        SetLastError(L"Workstation not initialized");
        return false;
    }

    HRESULT hr;

    // Step 1: Get the bus object via IDispatch navigation (not vtable)
    // Using ITopologyDevice::Bus("Test") DISPID 38 which may return a more
    // complete COM object than the vtable GetPort/GetBus path
    IDispatch* pBusDisp = nullptr;
    IUnknown* pBus = nullptr;

    // Try ITopologyDevice dual interface DISPID 38 = Bus(Name)
    {
        // QI workstation for ITopologyDevice IID (B2A20A5E)
        static const GUID IID_ITopologyDevice_Dual =
            {0xB2A20A5E, 0xF7B9, 0x11D0, {0xAD, 0x73, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9}};

        IDispatch* pWsDisp = nullptr;
        hr = m_pWorkstation->QueryInterface(IID_ITopologyDevice_Dual, (void**)&pWsDisp);
        if (SUCCEEDED(hr) && pWsDisp)
        {
            std::wcout << L"  [OK] Got ITopologyDevice dual interface" << std::endl;

            // Call Bus(Name) DISPID 38 with driver name
            VARIANT argName;
            VariantInit(&argName);
            argName.vt = VT_BSTR;
            argName.bstrVal = SysAllocString(driverName.c_str());
            DISPPARAMS dpBus = { &argName, nullptr, 1, 0 };
            VARIANT varBus;
            VariantInit(&varBus);

            hr = pWsDisp->Invoke(38, IID_NULL, LOCALE_USER_DEFAULT,
                                  DISPATCH_PROPERTYGET, &dpBus, &varBus, nullptr, nullptr);
            std::wcout << L"  Bus(\"" << driverName << L"\"): hr=0x" << std::hex << hr
                       << L" vt=" << std::dec << varBus.vt << std::endl;

            if (SUCCEEDED(hr) && (varBus.vt == VT_DISPATCH || varBus.vt == VT_UNKNOWN))
            {
                IUnknown* pBusUnk = (varBus.vt == VT_DISPATCH) ? (IUnknown*)varBus.pdispVal : varBus.punkVal;
                if (pBusUnk)
                {
                    // QI for ITopologyBus dual interface
                    hr = pBusUnk->QueryInterface(IID_ITopologyBus, (void**)&pBusDisp);
                    if (SUCCEEDED(hr) && pBusDisp)
                    {
                        pBus = pBusUnk;
                        pBus->AddRef();
                        std::wcout << L"  [OK] Got ITopologyBus from Bus(Name)" << std::endl;

                        // Check if this object supports IRSTopologyNetwork (the vtable version didn't)
                        IUnknown* pNetTest = nullptr;
                        hr = pBus->QueryInterface(IID_IRSTopologyNetwork, (void**)&pNetTest);
                        std::wcout << L"  IRSTopologyNetwork: " << (SUCCEEDED(hr) ? L"YES" : L"NO") << std::endl;
                        if (pNetTest) pNetTest->Release();
                    }
                }
            }

            VariantClear(&argName);
            VariantClear(&varBus);
            pWsDisp->Release();
        }
    }

    // Fallback: use vtable navigation
    if (!pBusDisp)
    {
        std::wcout << L"  Falling back to vtable bus navigation..." << std::endl;
        pBus = GetBusObjectByName(driverName);
        if (!pBus)
        {
            SetLastError(L"Failed to get bus object for driver");
            return false;
        }
        hr = pBus->QueryInterface(IID_ITopologyBus, (void**)&pBusDisp);
        if (FAILED(hr) || !pBusDisp)
        {
            pBus->Release();
            SetLastError(L"Bus does not support ITopologyBus");
            return false;
        }
    }

    // Step 3: Call ConnectNewDevice via IDispatch::Invoke DISPID 54
    //
    // TypeLib signature:
    //   ConnectNewDevice(
    //     [in] USERDEFINED flags,      -> VT_I4, value 0
    //     [in] BSTR Type,              -> Device class GUID string
    //     [out] VARIANT* port,         -> VT_VARIANT|VT_BYREF (output)
    //     [in] VARIANT Name,           -> Device name
    //     [in] VARIANT portLabel,      -> Port label (empty)
    //     [in] VARIANT address         -> IP address
    //   )
    //
    // COM rgvarg is RIGHT-TO-LEFT:
    //   rgvarg[0] = address (last param)
    //   rgvarg[1] = portLabel
    //   rgvarg[2] = Name
    //   rgvarg[3] = port (out)
    //   rgvarg[4] = Type
    //   rgvarg[5] = flags (first param)

    // RSTopGenericDevice class GUID from TypeLib dump
    // NOT the bus class! The bus class {00010010-...} is Ethernet.
    // Device class {6550834C-...} implements IRSTopologyDevice which ConnectNewDevice needs.
    BSTR bstrType = SysAllocString(L"{6550834C-A8C6-11CF-AC15-A07C03C10E27}");

    std::wstring name = deviceName.empty() ? ipAddress : deviceName;

    // Convert IP address string to proper VARIANT format using DisplayStringToAddress (DISPID 62)
    VARIANT varAddress;
    VariantInit(&varAddress);
    {
        VARIANT argAddr;
        VariantInit(&argAddr);
        argAddr.vt = VT_BSTR;
        argAddr.bstrVal = SysAllocString(ipAddress.c_str());
        DISPPARAMS dpAddr = { &argAddr, nullptr, 1, 0 };
        VARIANT rAddr;
        VariantInit(&rAddr);

        HRESULT hrAddr = pBusDisp->Invoke(62, IID_NULL, LOCALE_USER_DEFAULT,
                                            DISPATCH_METHOD, &dpAddr, &rAddr, nullptr, nullptr);
        std::wcout << L"  DisplayStringToAddress: hr=0x" << std::hex << hrAddr
                   << L" vt=" << std::dec << rAddr.vt << std::endl;

        if (SUCCEEDED(hrAddr) && rAddr.vt != VT_EMPTY)
        {
            VariantCopy(&varAddress, &rAddr);
            std::wcout << L"  [OK] Converted address to vt=" << varAddress.vt << std::endl;
        }
        else
        {
            // Fallback: use BSTR directly
            varAddress.vt = VT_BSTR;
            varAddress.bstrVal = SysAllocString(ipAddress.c_str());
        }
        VariantClear(&argAddr);
        VariantClear(&rAddr);
    }

    VARIANT args[6];
    for (int i = 0; i < 6; i++) VariantInit(&args[i]);

    // args[0] = address (properly formatted VARIANT)
    VariantCopy(&args[0], &varAddress);

    // args[1] = portLabel (VARIANT, empty string)
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(L"");

    // args[2] = Name (VARIANT containing BSTR)
    args[2].vt = VT_BSTR;
    args[2].bstrVal = SysAllocString(name.c_str());

    // args[3] = port ([out] VARIANT* - pass by reference)
    VARIANT portOut;
    VariantInit(&portOut);
    args[3].vt = VT_VARIANT | VT_BYREF;
    args[3].pvarVal = &portOut;

    // args[4] = Type (BSTR)
    args[4].vt = VT_BSTR;
    args[4].bstrVal = bstrType;

    // args[5] = flags (USERDEFINED -> I4)
    args[5].vt = VT_I4;
    args[5].lVal = 0;

    DISPPARAMS dp = {};
    dp.rgvarg = args;
    dp.cArgs = 6;
    dp.cNamedArgs = 0;
    dp.rgdispidNamedArgs = nullptr;

    VARIANT result;
    VariantInit(&result);
    EXCEPINFO excep = {};
    UINT argErr = 0;

    std::wcout << L"[DEBUG] Calling ConnectNewDevice DISPID 54 with 6 args..." << std::endl;
    hr = pBusDisp->Invoke(54, IID_NULL, LOCALE_USER_DEFAULT,
                          DISPATCH_METHOD, &dp, &result, &excep, &argErr);

    // Helper lambda to print EXCEPINFO
    auto printExcep = [](const wchar_t* prefix, HRESULT hr, EXCEPINFO& ei, UINT argErr) {
        std::wcout << prefix << L" hr=0x" << std::hex << hr << std::dec << std::endl;
        if (hr == DISP_E_EXCEPTION)
        {
            // Call deferred fill if available
            if (ei.pfnDeferredFillIn)
                ei.pfnDeferredFillIn(&ei);
            if (ei.bstrDescription)
                std::wcout << L"  Description: " << ei.bstrDescription << std::endl;
            if (ei.bstrSource)
                std::wcout << L"  Source: " << ei.bstrSource << std::endl;
            std::wcout << L"  scode: 0x" << std::hex << ei.scode << std::dec << std::endl;
        }
        if (hr == DISP_E_TYPEMISMATCH || hr == DISP_E_PARAMNOTFOUND)
            std::wcout << L"  Arg error at index: " << argErr << std::endl;
    };

    if (FAILED(hr))
    {
        printExcep(L"[FAIL] 6-arg attempt:", hr, excep, argErr);

        // Try with 5 args (port as retval, not in rgvarg)
        std::wcout << L"[INFO] Retrying with 5 args (port as retval)..." << std::endl;
        {
            VARIANT a5[5];
            for (int i = 0; i < 5; i++) VariantInit(&a5[i]);
            VariantCopy(&a5[0], &varAddress);  // Use converted address
            a5[1].vt = VT_BSTR; a5[1].bstrVal = SysAllocString(L"");
            a5[2].vt = VT_BSTR; a5[2].bstrVal = SysAllocString(name.c_str());
            a5[3].vt = VT_BSTR; a5[3].bstrVal = SysAllocString(L"{6550834C-A8C6-11CF-AC15-A07C03C10E27}");
            a5[4].vt = VT_I4;   a5[4].lVal = 0;

            DISPPARAMS dp5 = { a5, nullptr, 5, 0 };
            VARIANT r5; VariantInit(&r5);
            EXCEPINFO e5 = {}; UINT ae5 = 0;

            hr = pBusDisp->Invoke(54, IID_NULL, LOCALE_USER_DEFAULT,
                                  DISPATCH_METHOD, &dp5, &r5, &e5, &ae5);
            if (FAILED(hr)) printExcep(L"  5-arg:", hr, e5, ae5);
            else std::wcout << L"  [OK] 5-arg succeeded!" << std::endl;

            for (int i = 0; i < 5; i++) VariantClear(&a5[i]);
            VariantClear(&r5);
            SysFreeString(e5.bstrDescription); SysFreeString(e5.bstrSource); SysFreeString(e5.bstrHelpFile);
        }

        // Try with Type as plain class name "Ethernet" instead of GUID
        if (FAILED(hr))
        {
            std::wcout << L"[INFO] Retrying with Type='Ethernet'..." << std::endl;
            VARIANT a5[5];
            for (int i = 0; i < 5; i++) VariantInit(&a5[i]);
            VariantCopy(&a5[0], &varAddress);
            a5[1].vt = VT_BSTR; a5[1].bstrVal = SysAllocString(L"");
            a5[2].vt = VT_BSTR; a5[2].bstrVal = SysAllocString(name.c_str());
            a5[3].vt = VT_BSTR; a5[3].bstrVal = SysAllocString(L"Ethernet");
            a5[4].vt = VT_I4;   a5[4].lVal = 0;

            DISPPARAMS dp5 = { a5, nullptr, 5, 0 };
            VARIANT r5; VariantInit(&r5);
            EXCEPINFO e5 = {}; UINT ae5 = 0;

            hr = pBusDisp->Invoke(54, IID_NULL, LOCALE_USER_DEFAULT,
                                  DISPATCH_METHOD, &dp5, &r5, &e5, &ae5);
            if (FAILED(hr)) printExcep(L"  Ethernet:", hr, e5, ae5);
            else std::wcout << L"  [OK] Ethernet type succeeded!" << std::endl;

            for (int i = 0; i < 5; i++) VariantClear(&a5[i]);
            VariantClear(&r5);
            SysFreeString(e5.bstrDescription); SysFreeString(e5.bstrSource); SysFreeString(e5.bstrHelpFile);
        }

        // Try with minimal params: just flags + Type + address (3 args)
        if (FAILED(hr))
        {
            std::wcout << L"[INFO] Retrying with 3 args (flags, Type, address)..." << std::endl;
            VARIANT a3[3];
            for (int i = 0; i < 3; i++) VariantInit(&a3[i]);
            VariantCopy(&a3[0], &varAddress);
            a3[1].vt = VT_BSTR; a3[1].bstrVal = SysAllocString(L"{6550834C-A8C6-11CF-AC15-A07C03C10E27}");
            a3[2].vt = VT_I4;   a3[2].lVal = 0;

            DISPPARAMS dp3 = { a3, nullptr, 3, 0 };
            VARIANT r3; VariantInit(&r3);
            EXCEPINFO e3 = {}; UINT ae3 = 0;

            hr = pBusDisp->Invoke(54, IID_NULL, LOCALE_USER_DEFAULT,
                                  DISPATCH_METHOD, &dp3, &r3, &e3, &ae3);
            if (FAILED(hr)) printExcep(L"  3-arg:", hr, e3, ae3);
            else std::wcout << L"  [OK] 3-arg succeeded!" << std::endl;

            for (int i = 0; i < 3; i++) VariantClear(&a3[i]);
            VariantClear(&r3);
            SysFreeString(e3.bstrDescription); SysFreeString(e3.bstrSource); SysFreeString(e3.bstrHelpFile);
        }

        // Probe all available methods on ITopologyBus
        std::wcout << L"[DEBUG] Probing available methods on ITopologyBus..." << std::endl;
        LPOLESTR probeNames[] = {
            const_cast<LPOLESTR>(L"ConnectNewDevice"),
            const_cast<LPOLESTR>(L"AddDevice"),
            const_cast<LPOLESTR>(L"Connect"),
            const_cast<LPOLESTR>(L"Devices"),
            const_cast<LPOLESTR>(L"Browse"),
            const_cast<LPOLESTR>(L"StartBrowse"),
            const_cast<LPOLESTR>(L"StopBrowse"),
            const_cast<LPOLESTR>(L"Ports"),
            const_cast<LPOLESTR>(L"Name"),
            const_cast<LPOLESTR>(L"path"),
            const_cast<LPOLESTR>(L"Class"),
            const_cast<LPOLESTR>(L"ClassName"),
            const_cast<LPOLESTR>(L"Parent"),
            const_cast<LPOLESTR>(L"Port"),
            const_cast<LPOLESTR>(L"OnlineBrowse"),
            const_cast<LPOLESTR>(L"OnlineEnumerator"),
            const_cast<LPOLESTR>(L"BrowseState"),
            const_cast<LPOLESTR>(L"RefreshDevices"),
            const_cast<LPOLESTR>(L"Update"),
            const_cast<LPOLESTR>(L"Insert"),
            const_cast<LPOLESTR>(L"Add"),
            const_cast<LPOLESTR>(L"Item"),
            const_cast<LPOLESTR>(L"Count"),
            const_cast<LPOLESTR>(L"Remove"),
        };
        for (int n = 0; n < 24; n++)
        {
            DISPID did;
            HRESULT hrN = pBusDisp->GetIDsOfNames(IID_NULL, &probeNames[n], 1, LOCALE_USER_DEFAULT, &did);
            if (SUCCEEDED(hrN))
                std::wcout << L"  \"" << probeNames[n] << L"\" = DISPID " << did << std::endl;
        }

        // Try using Devices collection Add method instead
        std::wcout << L"[DEBUG] Trying Devices collection Add method..." << std::endl;
        {
            DISPPARAMS noArgs = { nullptr, nullptr, 0, 0 };
            VARIANT varDevices;
            VariantInit(&varDevices);

            HRESULT hrDev = pBusDisp->Invoke(50, IID_NULL, LOCALE_USER_DEFAULT,
                                              DISPATCH_PROPERTYGET, &noArgs, &varDevices, nullptr, nullptr);
            std::wcout << L"  Devices DISPID 50: hr=0x" << std::hex << hrDev
                       << L" vt=" << std::dec << varDevices.vt << std::endl;

            if (SUCCEEDED(hrDev) && varDevices.vt == VT_DISPATCH && varDevices.pdispVal)
            {
                IDispatch* pDevices = varDevices.pdispVal;

                // Probe collection methods
                LPOLESTR colNames[] = {
                    const_cast<LPOLESTR>(L"Count"), const_cast<LPOLESTR>(L"Item"),
                    const_cast<LPOLESTR>(L"Add"), const_cast<LPOLESTR>(L"Insert"),
                    const_cast<LPOLESTR>(L"Remove"), const_cast<LPOLESTR>(L"_NewEnum"),
                };
                for (int n = 0; n < 6; n++)
                {
                    DISPID did;
                    HRESULT hrN = pDevices->GetIDsOfNames(IID_NULL, &colNames[n], 1, LOCALE_USER_DEFAULT, &did);
                    if (SUCCEEDED(hrN))
                        std::wcout << L"  Collection.\"" << colNames[n] << L"\" = DISPID " << did << std::endl;
                }

                // Get initial count
                VARIANT varCount;
                VariantInit(&varCount);
                pDevices->Invoke(1, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_PROPERTYGET, &noArgs, &varCount, nullptr, nullptr);
                std::wcout << L"  Initial Devices.Count = " << varCount.lVal << std::endl;
                VariantClear(&varCount);

                // Get Add method signature via ITypeInfo
                ITypeInfo* pTI = nullptr;
                HRESULT hrTI = pDevices->GetTypeInfo(0, LOCALE_USER_DEFAULT, &pTI);
                if (SUCCEEDED(hrTI) && pTI)
                {
                    std::wcout << L"  [OK] Got TypeInfo for collection" << std::endl;

                    // Get FUNCDESC for DISPID 2 (Add)
                    TYPEATTR* pAttr = nullptr;
                    hrTI = pTI->GetTypeAttr(&pAttr);
                    if (SUCCEEDED(hrTI) && pAttr)
                    {
                        std::wcout << L"  TypeInfo: " << pAttr->cFuncs << L" functions, "
                                   << pAttr->cVars << L" vars" << std::endl;

                        for (WORD f = 0; f < pAttr->cFuncs; f++)
                        {
                            FUNCDESC* pFunc = nullptr;
                            hrTI = pTI->GetFuncDesc(f, &pFunc);
                            if (SUCCEEDED(hrTI) && pFunc)
                            {
                                // Get function name
                                BSTR funcName = nullptr;
                                UINT nameCount = 0;
                                pTI->GetNames(pFunc->memid, &funcName, 1, &nameCount);

                                std::wcout << L"  func[" << f << L"] DISPID="
                                           << pFunc->memid << L" \"" << (funcName ? funcName : L"?")
                                           << L"\" params=" << pFunc->cParams
                                           << L" invkind=" << pFunc->invkind << std::endl;

                                // Print parameter types for Add method
                                if (pFunc->memid == 2)
                                {
                                    BSTR* paramNames = new BSTR[pFunc->cParams + 1];
                                    UINT pnCount = 0;
                                    pTI->GetNames(pFunc->memid, paramNames, pFunc->cParams + 1, &pnCount);

                                    for (SHORT p = 0; p < pFunc->cParams; p++)
                                    {
                                        std::wcout << L"    param[" << p << L"] vt=0x"
                                                   << std::hex << pFunc->lprgelemdescParam[p].tdesc.vt << std::dec;
                                        if (p + 1 < (SHORT)pnCount)
                                            std::wcout << L" name=\"" << paramNames[p + 1] << L"\"";
                                        USHORT flags = pFunc->lprgelemdescParam[p].paramdesc.wParamFlags;
                                        if (flags & PARAMFLAG_FIN) std::wcout << L" [in]";
                                        if (flags & PARAMFLAG_FOUT) std::wcout << L" [out]";
                                        if (flags & PARAMFLAG_FRETVAL) std::wcout << L" [retval]";
                                        std::wcout << std::endl;
                                    }

                                    for (UINT i = 0; i < pnCount; i++) SysFreeString(paramNames[i]);
                                    delete[] paramNames;
                                }

                                SysFreeString(funcName);
                                pTI->ReleaseFuncDesc(pFunc);
                            }
                        }

                        pTI->ReleaseTypeAttr(pAttr);
                    }
                    pTI->Release();
                }

                // Try Add with various parameter patterns
                // Try creating a device via CoCreateInstance and adding it
                // RSTopGenericDevice {6550834C-A8C6-11CF-AC15-A07C03C10E27}
                static const GUID CLSID_RSTopGenericDevice =
                    {0x6550834C, 0xA8C6, 0x11CF, {0xAC, 0x15, 0xA0, 0x7C, 0x03, 0xC1, 0x0E, 0x27}};
                static const GUID CLSID_RSTopUnknownDevice =
                    {0x21E93224, 0x90F5, 0x11D0, {0xAD, 0x56, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9}};

                // Try CoCreateInstance for each device CLSID
                const GUID* deviceCLSIDs[] = { &CLSID_RSTopGenericDevice, &CLSID_RSTopUnknownDevice };
                const wchar_t* deviceNames[] = { L"GenericDevice", L"UnknownDevice" };

                for (int ci = 0; ci < 2 && FAILED(hr); ci++)
                {
                    IUnknown* pNewDevice = nullptr;
                    HRESULT hrCreate = CoCreateInstance(*deviceCLSIDs[ci], NULL, CLSCTX_ALL,
                                                        IID_IUnknown, (void**)&pNewDevice);
                    std::wcout << L"  CoCreate " << deviceNames[ci] << L": hr=0x"
                               << std::hex << hrCreate << std::dec << std::endl;

                    if (SUCCEEDED(hrCreate) && pNewDevice)
                    {
                        // Add(Name=ipAddress, item=VT_DISPATCH device)
                        IDispatch* pDevDisp = nullptr;
                        pNewDevice->QueryInterface(IID_IDispatch, (void**)&pDevDisp);

                        VARIANT aAdd[2];
                        VariantInit(&aAdd[0]); VariantInit(&aAdd[1]);
                        if (pDevDisp)
                        {
                            aAdd[0].vt = VT_DISPATCH;
                            aAdd[0].pdispVal = pDevDisp;
                        }
                        else
                        {
                            aAdd[0].vt = VT_UNKNOWN;
                            aAdd[0].punkVal = pNewDevice;
                        }
                        aAdd[1].vt = VT_BSTR;
                        aAdd[1].bstrVal = SysAllocString(ipAddress.c_str());

                        DISPPARAMS dpAdd = { aAdd, nullptr, 2, 0 };
                        VARIANT rAdd; VariantInit(&rAdd);
                        EXCEPINFO eAdd = {}; UINT aeAdd = 0;

                        HRESULT hrAdd = pDevices->Invoke(2, IID_NULL, LOCALE_USER_DEFAULT,
                                                          DISPATCH_METHOD, &dpAdd, &rAdd, &eAdd, &aeAdd);
                        std::wcout << L"  Add(name, " << deviceNames[ci] << L"): hr=0x"
                                   << std::hex << hrAdd << std::dec;
                        if (hrAdd == DISP_E_EXCEPTION)
                        {
                            if (eAdd.pfnDeferredFillIn) eAdd.pfnDeferredFillIn(&eAdd);
                            std::wcout << L" scode=0x" << std::hex << eAdd.scode << std::dec;
                        }
                        if (hrAdd == DISP_E_TYPEMISMATCH || hrAdd == DISP_E_PARAMNOTFOUND)
                            std::wcout << L" argErr=" << aeAdd;
                        std::wcout << std::endl;

                        if (SUCCEEDED(hrAdd))
                        {
                            hr = hrAdd;
                            std::wcout << L"  [OK] Collection.Add succeeded!" << std::endl;
                        }

                        VariantClear(&rAdd);
                        SysFreeString(eAdd.bstrDescription); SysFreeString(eAdd.bstrSource); SysFreeString(eAdd.bstrHelpFile);
                        if (pDevDisp) pDevDisp->Release();
                        pNewDevice->Release();
                    }
                }

                // Also try: Add(Name=ipAddress, item=VT_EMPTY) - maybe it creates a blank device
                if (FAILED(hr))
                {
                    VARIANT aAdd[2];
                    VariantInit(&aAdd[0]); VariantInit(&aAdd[1]);
                    // rgvarg[0] = item (VT_EMPTY), rgvarg[1] = Name
                    aAdd[1].vt = VT_BSTR; aAdd[1].bstrVal = SysAllocString(ipAddress.c_str());
                    DISPPARAMS dpAdd = { aAdd, nullptr, 2, 0 };
                    VARIANT rAdd; VariantInit(&rAdd);
                    EXCEPINFO eAdd = {}; UINT aeAdd = 0;

                    HRESULT hrAdd = pDevices->Invoke(2, IID_NULL, LOCALE_USER_DEFAULT,
                                                      DISPATCH_METHOD, &dpAdd, &rAdd, &eAdd, &aeAdd);
                    std::wcout << L"  Add(name, VT_EMPTY): hr=0x" << std::hex << hrAdd << std::dec;
                    if (hrAdd == DISP_E_EXCEPTION)
                    {
                        if (eAdd.pfnDeferredFillIn) eAdd.pfnDeferredFillIn(&eAdd);
                        std::wcout << L" scode=0x" << std::hex << eAdd.scode << std::dec;
                    }
                    std::wcout << std::endl;
                    if (SUCCEEDED(hrAdd)) { hr = hrAdd; std::wcout << L"  [OK] succeeded!" << std::endl; }
                    VariantClear(&aAdd[0]); VariantClear(&aAdd[1]); VariantClear(&rAdd);
                    SysFreeString(eAdd.bstrDescription); SysFreeString(eAdd.bstrSource); SysFreeString(eAdd.bstrHelpFile);
                }

                // Check count after adds
                VariantInit(&varCount);
                pDevices->Invoke(1, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_PROPERTYGET, &noArgs, &varCount, nullptr, nullptr);
                std::wcout << L"  Final Devices.Count = " << varCount.lVal << std::endl;
                VariantClear(&varCount);
            }

            VariantClear(&varDevices);
        }

        // Try DIRECT VTABLE CALL to ConnectNewDevice at vtable slot 19
        // (bypasses MFC IDispatch type checking which may cause E_NOINTERFACE on USERDEFINED params)
        std::wcout << L"[DEBUG] Trying ConnectNewDevice via direct vtable[19]..." << std::endl;
        {
            BSTR bstrTypeVt = SysAllocString(L"{6550834C-A8C6-11CF-AC15-A07C03C10E27}");

            VARIANT vtPort, vtName, vtPortLabel, vtAddr;
            VariantInit(&vtPort);
            VariantInit(&vtName);
            vtName.vt = VT_BSTR;
            vtName.bstrVal = SysAllocString(name.c_str());
            VariantInit(&vtPortLabel);
            vtPortLabel.vt = VT_BSTR;
            vtPortLabel.bstrVal = SysAllocString(L"");
            VariantCopy(&vtAddr, &varAddress);

            HRESULT hrVt = TryConnectNewDevice(pBusDisp, 0, bstrTypeVt,
                                                &vtPort, &vtName, &vtPortLabel, &vtAddr);
            std::wcout << L"  vtable[19] result: hr=0x" << std::hex << hrVt << std::dec << std::endl;

            if (hrVt == E_UNEXPECTED)
                std::wcout << L"  SEH exception in vtable[19]" << std::endl;

            if (SUCCEEDED(hrVt))
            {
                hr = hrVt;
                std::wcout << L"  [OK] ConnectNewDevice via vtable succeeded!" << std::endl;
                if (vtPort.vt != VT_EMPTY)
                    std::wcout << L"  Port output vt=" << vtPort.vt << std::endl;
            }
            else if (hrVt != E_UNEXPECTED)
            {
                // Try with RSTopUnknownDevice CLSID instead
                SysFreeString(bstrTypeVt);
                bstrTypeVt = SysAllocString(L"{21E93224-90F5-11D0-AD56-00C04FD915B9}");
                VariantInit(&vtPort);

                hrVt = TryConnectNewDevice(pBusDisp, 0, bstrTypeVt,
                                            &vtPort, &vtName, &vtPortLabel, &vtAddr);
                std::wcout << L"  UnknownDevice type: hr=0x" << std::hex << hrVt << std::dec << std::endl;

                if (SUCCEEDED(hrVt))
                {
                    hr = hrVt;
                    std::wcout << L"  [OK] ConnectNewDevice via vtable (UnknownDevice) succeeded!" << std::endl;
                }
            }

            SysFreeString(bstrTypeVt);
            VariantClear(&vtPort);
            VariantClear(&vtName);
            VariantClear(&vtPortLabel);
            VariantClear(&vtAddr);
        }
    }

    if (SUCCEEDED(hr))
    {
        std::wcout << L"[OK] ConnectNewDevice succeeded for " << ipAddress << std::endl;
        if (portOut.vt != VT_EMPTY)
            std::wcout << L"  Port output vt=" << portOut.vt << std::endl;
    }

    // Cleanup
    VariantClear(&varAddress);
    SysFreeString(excep.bstrDescription);
    SysFreeString(excep.bstrSource);
    SysFreeString(excep.bstrHelpFile);
    for (int i = 0; i < 6; i++)
    {
        if (i != 3) // Don't clear the byref arg
            VariantClear(&args[i]);
    }
    VariantClear(&portOut);
    VariantClear(&result);
    pBusDisp->Release();
    pBus->Release();

    return SUCCEEDED(hr);
}

// ========================================
// Private Helper Methods
// ========================================

bool TopologyBrowser::InitializeHarmonyServices()
{
    HRESULT hr = CoCreateInstance(CLSID_HarmonyServices, NULL, CLSCTX_ALL, IID_IUnknown, (void**)&m_pHarmonyConnector);
    if (FAILED(hr))
    {
        SetLastError(L"Failed to create HarmonyServices");
        return false;
    }

    hr = m_pHarmonyConnector->SetServerOptions(0, L"");
    if (FAILED(hr))
    {
        SetLastError(L"SetServerOptions failed");
        return false;
    }

    std::wcout << L"[OK] HarmonyServices initialized" << std::endl;
    return true;
}

bool TopologyBrowser::InitializeTopologyGlobals()
{
    IUnknown* pUnk = nullptr;
    HRESULT hr = m_pHarmonyConnector->GetSpecialObject(&CLSID_RSTopologyGlobals, &IID_IRSTopologyGlobals, &pUnk);
    if (FAILED(hr))
    {
        SetLastError(L"GetSpecialObject failed for RSTopologyGlobals");
        return false;
    }

    hr = pUnk->QueryInterface(IID_IRSTopologyGlobals, (void**)&m_pTopologyGlobals);
    pUnk->Release();

    if (FAILED(hr))
    {
        SetLastError(L"QueryInterface for IRSTopologyGlobals failed");
        return false;
    }

    std::wcout << L"[OK] RSTopologyGlobals obtained" << std::endl;
    return true;
}

bool TopologyBrowser::GetWorkstation()
{
    // Create project global
    HRESULT hr = CoCreateInstance(CLSID_RSProjectGlobal, NULL, CLSCTX_ALL, IID_IRSProjectGlobal, (void**)&m_pProjectGlobal);
    if (FAILED(hr))
    {
        SetLastError(L"Failed to create RSProjectGlobal");
        return false;
    }

    // Open project
    hr = m_pProjectGlobal->OpenProject(L"", 0, NULL, NULL, &IID_IRSProject, (void**)&m_pProject);
    if (FAILED(hr))
    {
        SetLastError(L"OpenProject failed");
        return false;
    }

    // Get workstation
    hr = m_pTopologyGlobals->GetThisWorkstationObject(m_pProject, &m_pWorkstation);
    if (FAILED(hr))
    {
        SetLastError(L"GetThisWorkstationObject failed");
        return false;
    }

    std::wcout << L"[OK] Workstation object obtained" << std::endl;
    return true;
}

IUnknown* TopologyBrowser::GetBusObjectByName(const std::wstring& driverName)
{
    if (!m_pWorkstation)
    {
        SetLastError(L"Workstation not initialized");
        return nullptr;
    }

    std::wcout << L"========================================" << std::endl;
    std::wcout << L"Bus Navigation via SaveAsXML + GetPort()" << std::endl;
    std::wcout << L"========================================" << std::endl;

    // Step 1: Save topology to XML to discover port names
    std::wstring xmlPath = L"C:\\temp\\topology_port_discovery.xml";
    if (!SaveTopologyXML(xmlPath))
    {
        std::wcout << L"[FAIL] Cannot save topology XML for port discovery" << std::endl;
        SetLastError(L"SaveTopologyXML failed");
        return nullptr;
    }

    // Step 2: Parse XML to find the port that contains the bus with driverName
    std::wcout << L"[INFO] Parsing XML to find port for bus: " << driverName << std::endl;

    // TODO: Implement XML parsing to extract port name
    // For now, use "Test" as requested
    std::wstring portName = L"Test";
    std::wcout << L"[INFO] Using port name: " << portName << std::endl;

    // Step 2.5: QI for IRSTopologyDevice (required to access GetPort method)
    // Define IRSTopologyDevice GUID: {DCEAD8E1-2E7A-11CF-B4B5-C46F03C10000}
    const GUID IID_IRSTopologyDevice =
        {0xDCEAD8E1, 0x2E7A, 0x11CF, {0xB4, 0xB5, 0xC4, 0x6F, 0x03, 0xC1, 0x00, 0x00}};

    IUnknown* pDevice = nullptr;
    HRESULT hr = m_pWorkstation->QueryInterface(IID_IRSTopologyDevice, (void**)&pDevice);

    if (FAILED(hr) || !pDevice)
    {
        std::wcout << L"[FAIL] Cannot QI workstation for IRSTopologyDevice: hr=0x" << std::hex << hr << std::dec << std::endl;
        SetLastError(L"Workstation does not support IRSTopologyDevice");
        return nullptr;
    }

    std::wcout << L"[OK] Got workstation as IRSTopologyDevice" << std::endl;

    // Step 3: Call IRSTopologyDevice::GetPort() via vtable (slot 18)
    // Signature: HRESULT GetPort(void* this, WCHAR* pwszPortName, void** ppPort)

    typedef HRESULT (__stdcall *GetPortFunc)(void* pThis, WCHAR* pwszPortName, void** ppPort);

    void** vtable = *(void***)pDevice;
    GetPortFunc GetPort = (GetPortFunc)vtable[18];  // Slot 18 from Ghidra analysis

    void* pPort = nullptr;
    hr = GetPort(pDevice, const_cast<WCHAR*>(portName.c_str()), &pPort);

    // Release IRSTopologyDevice - we don't need it anymore
    pDevice->Release();

    if (FAILED(hr) || !pPort)
    {
        std::wcout << L"[FAIL] GetPort() failed: hr=0x" << std::hex << hr << std::dec << std::endl;
        SetLastError(L"GetPort() failed");
        return nullptr;
    }

    std::wcout << L"[OK] Got port object via vtable GetPort()" << std::endl;

    // Step 4: Get the bus from the port using IRSTopologyPort::GetBus()
    // Signature: HRESULT GetBus(void* this, ITopologyBus** ppBus)
    // Vtable slot: 10 (confirmed by Ghidra analysis and Python testing)

    typedef HRESULT (__stdcall *GetBusFunc)(void* pThis, void** ppBus);

    void** portVtable = *(void***)pPort;
    GetBusFunc GetBus = (GetBusFunc)portVtable[10];  // Slot 10 from Ghidra analysis

    void* pBus = nullptr;
    hr = GetBus(pPort, &pBus);

    // Release port object - we don't need it anymore
    ((IUnknown*)pPort)->Release();

    if (FAILED(hr) || !pBus)
    {
        std::wcout << L"[FAIL] GetBus() failed: hr=0x" << std::hex << hr << std::dec << std::endl;
        SetLastError(L"GetBus() failed");
        return nullptr;
    }

    std::wcout << L"[OK] Got bus object via vtable GetBus()" << std::endl;

    // Return the bus object as IUnknown*
    return (IUnknown*)pBus;
}

bool TopologyBrowser::ConnectEventSink(IUnknown* pBus, BrowseEventSink* pSink, DWORD* pCookie)
{
    std::wcout << L"[DEBUG] ConnectEventSink: pBus=0x" << std::hex << (void*)pBus
               << L", pSink=0x" << (void*)pSink << std::dec << std::endl;

    // Get connection point container
    IConnectionPointContainer* pContainer = nullptr;
    HRESULT hr = pBus->QueryInterface(IID_IConnectionPointContainer, (void**)&pContainer);
    if (FAILED(hr))
    {
        std::wcout << L"[FAIL] QI for IConnectionPointContainer failed: hr=0x" << std::hex << hr << std::dec << std::endl;
        SetLastError(L"Bus does not support IConnectionPointContainer");
        return false;
    }
    std::wcout << L"[DEBUG] Got IConnectionPointContainer: 0x" << std::hex << (void*)pContainer << std::dec << std::endl;

    // Connect to ITopologyBusEvents
    IConnectionPoint* pConnPoint = nullptr;
    std::wcout << L"[DEBUG] Finding connection point for ITopologyBusEvents..." << std::endl;
    hr = pContainer->FindConnectionPoint(IID_ITopologyBusEvents, &pConnPoint);

    if (FAILED(hr))
    {
        std::wcout << L"[FAIL] FindConnectionPoint failed: hr=0x" << std::hex << hr << std::dec << std::endl;
        pContainer->Release();
        SetLastError(L"FindConnectionPoint failed");
        return false;
    }
    std::wcout << L"[DEBUG] Found connection point: 0x" << std::hex << (void*)pConnPoint << std::dec << std::endl;

    // Connect sink
    std::wcout << L"[DEBUG] Calling Advise..." << std::endl;
    hr = pConnPoint->Advise(pSink, pCookie);
    pConnPoint->Release();

    if (FAILED(hr))
    {
        std::wcout << L"[FAIL] Advise failed: hr=0x" << std::hex << hr << std::dec << std::endl;
        pContainer->Release();
        SetLastError(L"Advise failed");
        return false;
    }
    std::wcout << L"[OK] Connected to ITopologyBusEvents (cookie: " << *pCookie << L")" << std::endl;

    pContainer->Release();
    return true;
}

bool TopologyBrowser::DisconnectEventSink(IUnknown* pBus, DWORD cookie)
{
    IConnectionPointContainer* pContainer = nullptr;
    HRESULT hr = pBus->QueryInterface(IID_IConnectionPointContainer, (void**)&pContainer);
    if (FAILED(hr))
        return false;

    IConnectionPoint* pConnPoint = nullptr;
    hr = pContainer->FindConnectionPoint(IID_ITopologyBusEvents, &pConnPoint);
    pContainer->Release();

    if (FAILED(hr))
        return false;

    hr = pConnPoint->Unadvise(cookie);
    pConnPoint->Release();

    return SUCCEEDED(hr);
}

void TopologyBrowser::SetLastError(const std::wstring& error)
{
    m_lastError = error;
    std::wcerr << L"[ERROR] " << error << std::endl;
}

void TopologyBrowser::ReleaseAll()
{
    if (m_pWorkstation)
    {
        m_pWorkstation->Release();
        m_pWorkstation = nullptr;
    }
    if (m_pProject)
    {
        m_pProject->Release();
        m_pProject = nullptr;
    }
    if (m_pProjectGlobal)
    {
        m_pProjectGlobal->Release();
        m_pProjectGlobal = nullptr;
    }
    if (m_pTopologyGlobals)
    {
        m_pTopologyGlobals->Release();
        m_pTopologyGlobals = nullptr;
    }
    if (m_pHarmonyConnector)
    {
        m_pHarmonyConnector->Release();
        m_pHarmonyConnector = nullptr;
    }
}
