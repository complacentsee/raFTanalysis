#include "BrowseOperations.h"
#include "Logging.h"
#include "SEHHelpers.h"
#include "EventSink.h"
#include "DispatchHelpers.h"
#include "TopologyXML.h"
#include "STAHook.h"

// ============================================================
// BrowseOperations globals
// ============================================================

std::vector<ConnectionPointInfo> g_connectionPoints;
std::vector<EnumeratorInfo> g_enumerators;
DualEventSink* g_pMainSink = nullptr;
IUnknown* g_pMainEnumUnk = nullptr;

// ============================================================
// Enumerator tracking
// ============================================================

// Check if all enumerators added since 'baseline' index have completed their browse cycle
bool EnumeratorsCycledSince(int baseline)
{
    if ((int)g_enumerators.size() <= baseline) return true;
    for (int i = baseline; i < (int)g_enumerators.size(); i++)
    {
        if (g_enumerators[i].pSink && !g_enumerators[i].pSink->m_cycleComplete)
            return false;
    }
    return true;
}

// Count cycle-completed vs total enumerators since baseline
void GetEnumeratorStatusSince(int baseline, int& completed, int& total)
{
    completed = 0;
    total = (int)g_enumerators.size() - baseline;
    if (total < 0) total = 0;
    for (int i = baseline; i < (int)g_enumerators.size(); i++)
    {
        if (g_enumerators[i].pSink && g_enumerators[i].pSink->m_cycleComplete)
            completed++;
    }
}

// ============================================================
// GetBusDispatch — fresh bus IDispatch from COM objects on current STA
// ============================================================

IDispatch* GetBusDispatch(const wchar_t* driverName)
{
    IHarmonyConnector* pHarmony = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_HarmonyServices, NULL, CLSCTX_ALL,
                                   IID_IHarmonyConnector, (void**)&pHarmony);
    if (FAILED(hr)) return nullptr;
    pHarmony->SetServerOptions(0, L"");

    IRSTopologyGlobals* pGlobals = nullptr;
    {
        IUnknown* pUnk = nullptr;
        hr = pHarmony->GetSpecialObject(&CLSID_RSTopologyGlobals, &IID_IRSTopologyGlobals, &pUnk);
        if (FAILED(hr)) { pHarmony->Release(); return nullptr; }
        pUnk->QueryInterface(IID_IRSTopologyGlobals, (void**)&pGlobals);
        pUnk->Release();
    }

    IRSProject* pProject = nullptr;
    {
        IRSProjectGlobal* pPG = nullptr;
        hr = CoCreateInstance(CLSID_RSProjectGlobal, NULL, CLSCTX_ALL,
                              IID_IRSProjectGlobal, (void**)&pPG);
        if (FAILED(hr)) { pGlobals->Release(); pHarmony->Release(); return nullptr; }
        hr = pPG->OpenProject(L"", 0, NULL, NULL, &IID_IRSProject, (void**)&pProject);
        pPG->Release();
        if (FAILED(hr)) { pGlobals->Release(); pHarmony->Release(); return nullptr; }
    }

    IUnknown* pWorkstation = nullptr;
    hr = pGlobals->GetThisWorkstationObject(pProject, &pWorkstation);
    pProject->Release(); pGlobals->Release(); pHarmony->Release();
    if (FAILED(hr)) return nullptr;

    IDispatch* pBusDisp = nullptr;
    {
        IDispatch* pWsDisp = nullptr;
        hr = pWorkstation->QueryInterface(IID_ITopologyDevice_Dual, (void**)&pWsDisp);
        if (SUCCEEDED(hr) && pWsDisp)
        {
            VARIANT argName;
            VariantInit(&argName);
            argName.vt = VT_BSTR;
            argName.bstrVal = SysAllocString(driverName);
            DISPPARAMS dp = { &argName, nullptr, 1, 0 };
            VARIANT varBus;
            VariantInit(&varBus);
            hr = pWsDisp->Invoke(38, IID_NULL, LOCALE_USER_DEFAULT,
                                  DISPATCH_PROPERTYGET, &dp, &varBus, nullptr, nullptr);
            if (SUCCEEDED(hr) && (varBus.vt == VT_DISPATCH || varBus.vt == VT_UNKNOWN))
            {
                IUnknown* pBusUnk = (varBus.vt == VT_DISPATCH)
                    ? (IUnknown*)varBus.pdispVal : varBus.punkVal;
                if (pBusUnk)
                {
                    pBusUnk->AddRef();
                    pBusUnk->QueryInterface(IID_ITopologyBus, (void**)&pBusDisp);
                    pBusUnk->Release();
                }
            }
            VariantClear(&argName);
            VariantClear(&varBus);
            pWsDisp->Release();
        }
        pWorkstation->Release();
    }
    return pBusDisp;
}

// ============================================================
// DoBusBrowse — runs on MAIN STA thread
// ============================================================

HRESULT DoBusBrowse()
{
    Log(L"[BUS] DoBusBrowse starting on TID=%d", GetCurrentThreadId());

    if (!g_pSharedConfig)
    {
        Log(L"[BUS] ERROR: no shared config");
        return E_INVALIDARG;
    }

    int startedCount = 0;
    int backplanesFound = 0;

    for (auto& drv : g_pSharedConfig->drivers)
    {
    Log(L"[BUS] === Driver: %s ===", drv.name.c_str());
    IDispatch* pBusDisp = GetBusDispatch(drv.name.c_str());
    if (!pBusDisp)
    {
        Log(L"[BUS] FAIL: Could not get bus '%s'", drv.name.c_str());
        continue;
    }
    Log(L"[BUS] Bus OK: 0x%p", pBusDisp);

    IDispatch* pDevices = DispatchGetCollection(pBusDisp, 50);
    if (!pDevices)
    {
        Log(L"[BUS] FAIL: bus.Devices() returned null");
        pBusDisp->Release();
        continue;
    }

    int deviceCount = DispatchGetInt(pDevices, 1);
    Log(L"[BUS] Bus has %d devices", deviceCount);

    std::vector<IDispatch*> devices = EnumerateCollection(pDevices);
    Log(L"[BUS] Enumerated %d devices", (int)devices.size());

    for (int i = 0; i < (int)devices.size(); i++)
    {
        IDispatch* pDevice = devices[i];
        if (!pDevice) continue;

        std::wstring devName = DispatchGetString(pDevice, 1);
        Log(L"[BUS] Device %d: \"%s\"", i, devName.c_str());

        {
            std::wstring objectId = DispatchGetString(pDevice, 2);
            Log(L"[BUS] Device %d: objectId='%s'", i, objectId.c_str());

            DeviceInfo info;
            info.productName = devName;
            info.objectId = objectId;
            if (!devName.empty())
                g_deviceDetails[devName] = info;
        }

        // Probe DISPIDs if requested (Phase A discovery)
        if (g_pSharedConfig->probeDispids)
        {
            std::wstring probeLabel = L"BUS/" + devName;
            ProbeDeviceDISPIDs(pDevice, probeLabel.c_str());

            // Also probe the bus itself (first device only to avoid log spam)
            if (i == 0)
                ProbeBusDISPIDs(pBusDisp, drv.name.c_str());
        }

        IUnknown* pDevVtable = nullptr;
        pDevice->QueryInterface(IID_IRSTopologyDevice, (void**)&pDevVtable);
        if (!pDevVtable)
        {
            Log(L"[BUS]   QI for IRSTopologyDevice FAILED, skipping");
            pDevice->Release();
            continue;
        }

        IUnknown* pBackplanePort = nullptr;
        HRESULT hrBP = TryVtableGetObject(pDevVtable, 19, &pBackplanePort);
        Log(L"[BUS]   GetBackplanePort[19]: hr=0x%08x port=0x%p", hrBP, pBackplanePort);
        pDevVtable->Release();

        if (FAILED(hrBP) || !pBackplanePort)
        {
            Log(L"[BUS]   No backplane port (device may not have backplane)");
            pDevice->Release();
            continue;
        }
        backplanesFound++;

        pBackplanePort->Release();

        IUnknown* pDevUnk2 = nullptr;
        pDevice->QueryInterface(IID_IUnknown, (void**)&pDevUnk2);
        void* pDevEnum = nullptr;
        HRESULT hrEnum = pDevUnk2->QueryInterface(IID_IOnlineEnumeratorTypeLib, &pDevEnum);
        Log(L"[BUS]   QI device for enumerator: hr=0x%08x ptr=0x%p", hrEnum, pDevEnum);
        pDevUnk2->Release();

        if (FAILED(hrEnum) || !pDevEnum)
        {
            Log(L"[BUS]   Device does not support enumerator QI");
            pDevice->Release();
            continue;
        }

        IDispatch* pDevTopoObj = nullptr;
        pDevice->QueryInterface(IID_ITopologyObject, (void**)&pDevTopoObj);
        if (!pDevTopoObj)
            pDevice->QueryInterface(IID_IDispatch, (void**)&pDevTopoObj);

        IUnknown* pDevPath = nullptr;
        if (pDevTopoObj)
        {
            pDevPath = DispatchGetPath(pDevTopoObj);
            Log(L"[BUS]   Device path: 0x%p", pDevPath);
            pDevTopoObj->Release();
        }

        if (!pDevPath)
        {
            Log(L"[BUS]   Could not get device path, skipping");
            pDevice->Release();
            continue;
        }

        DualEventSink* pSink = new DualEventSink(devName.c_str());
        {
            IConnectionPointContainer* pDevCPC = nullptr;
            ((IUnknown*)pDevEnum)->QueryInterface(IID_IConnectionPointContainer, (void**)&pDevCPC);
            if (pDevCPC)
            {
                IEnumConnectionPoints* pEnumCP = nullptr;
                pDevCPC->EnumConnectionPoints(&pEnumCP);
                if (pEnumCP)
                {
                    IConnectionPoint* pCP = nullptr;
                    ULONG fetched = 0;
                    int cpIdx = 0;
                    while (pEnumCP->Next(1, &pCP, &fetched) == S_OK && fetched > 0)
                    {
                        DWORD c = 0;
                        HRESULT hrAdv = pCP->Advise(static_cast<IRSTopologyOnlineNotify*>(pSink), &c);
                        if (SUCCEEDED(hrAdv))
                            g_connectionPoints.push_back({pCP, c});
                        else
                            pCP->Release();
                        cpIdx++;
                    }
                    Log(L"[BUS]   Connected %d device enum CPs", cpIdx);
                    pEnumCP->Release();
                }
                pDevCPC->Release();
            }
        }

        HRESULT hrStart = TryStartAtSlot(pDevEnum, pDevPath, 7);
        Log(L"[BUS]   Start(device path) via device enum: hr=0x%08x", hrStart);
        if (SUCCEEDED(hrStart))
        {
            startedCount++;
            Log(L"[BUS]   >> Backplane browse started for \"%s\"", devName.c_str());
        }

        g_enumerators.push_back({pDevEnum, pSink});

        SafeRelease(pDevPath, L"pDevPath");
        pDevice->Release();
    }

    pDevices->Release();
    pBusDisp->Release();
    } // end per-driver loop

    Log(L"[BUS] DoBusBrowse done: %d backplanes found, %d started across %d drivers",
        backplanesFound, startedCount, (int)g_pSharedConfig->drivers.size());
    return (startedCount > 0) ? S_OK : S_FALSE;
}

// ============================================================
// DoBackplaneBrowse — runs on MAIN STA thread (Phase 4b)
// ============================================================

HRESULT DoBackplaneBrowse()
{
    Log(L"[BP] DoBackplaneBrowse starting on TID=%d", GetCurrentThreadId());

    if (!g_pSharedConfig)
    {
        Log(L"[BP] ERROR: no shared config");
        return E_INVALIDARG;
    }

    int startedCount = 0;

    for (auto& drv : g_pSharedConfig->drivers)
    {
    Log(L"[BP] === Driver: %s ===", drv.name.c_str());
    IDispatch* pEthBusDisp = GetBusDispatch(drv.name.c_str());
    if (!pEthBusDisp)
    {
        Log(L"[BP] FAIL: Could not get Ethernet bus '%s'", drv.name.c_str());
        continue;
    }

    IDispatch* pDevices = DispatchGetCollection(pEthBusDisp, 50);
    if (!pDevices) { pEthBusDisp->Release(); continue; }

    std::vector<IDispatch*> devices = EnumerateCollection(pDevices);
    Log(L"[BP] Enumerated %d Ethernet devices", (int)devices.size());

    for (int i = 0; i < (int)devices.size(); i++)
    {
        IDispatch* pDevice = devices[i];
        if (!pDevice) continue;

        std::wstring devName = DispatchGetString(pDevice, 1);
        Log(L"[BP] Device %d: \"%s\"", i, devName.c_str());

        {
            std::wstring objectId = DispatchGetString(pDevice, 2);
            Log(L"[BP] Device %d: objectId='%s'", i, objectId.c_str());

            DeviceInfo info;
            info.productName = devName;
            info.objectId = objectId;
            if (!devName.empty())
                g_deviceDetails[devName] = info;
        }

        // Probe DISPIDs if requested (Phase A discovery)
        if (g_pSharedConfig->probeDispids)
        {
            std::wstring probeLabel = L"BP/" + devName;
            ProbeDeviceDISPIDs(pDevice, probeLabel.c_str());
        }

        IUnknown* pDevVtable = nullptr;
        pDevice->QueryInterface(IID_IRSTopologyDevice, (void**)&pDevVtable);
        if (!pDevVtable) { pDevice->Release(); continue; }

        IUnknown* pBackplanePort = nullptr;
        HRESULT hrBP = TryVtableGetObject(pDevVtable, 19, &pBackplanePort);
        pDevVtable->Release();

        if (FAILED(hrBP) || !pBackplanePort)
        {
            Log(L"[BP]   No backplane port");
            pDevice->Release();
            continue;
        }

        Log(L"[BP]   Backplane port: 0x%p", pBackplanePort);

        IUnknown* pBackplaneBus = nullptr;
        std::wstring busLabel;

        IUnknown* pPortVtable = nullptr;
        HRESULT hrPortQI = pBackplanePort->QueryInterface(IID_IRSTopologyPort, (void**)&pPortVtable);
        Log(L"[BP]   QI IRSTopologyPort: hr=0x%08x", hrPortQI);

        if (pPortVtable)
        {
            IUnknown* pBusRaw = nullptr;
            HRESULT hrGetBus = TryVtableGetObject(pPortVtable, 10, &pBusRaw);
            Log(L"[BP]   GetBus[10]: hr=0x%08x bus=0x%p", hrGetBus, pBusRaw);

            if (SUCCEEDED(hrGetBus) && pBusRaw)
            {
                pBackplaneBus = pBusRaw;

                IUnknown* pBusRSObj = nullptr;
                if (SUCCEEDED(pBusRaw->QueryInterface(IID_IRSObject, (void**)&pBusRSObj)) && pBusRSObj)
                {
                    TryVtableGetLabel(pBusRSObj, 7, busLabel);
                    Log(L"[BP]   Bus IRSObject::GetName[7]: \"%s\"", busLabel.c_str());
                    pBusRSObj->Release();
                }
            }
            pPortVtable->Release();
        }

        if (!pBackplaneBus)
        {
            Log(L"[BP]   GetBus failed, falling back to DISPID 38");
            IDispatch* pDevDualDisp = nullptr;
            pDevice->QueryInterface(IID_ITopologyDevice_Dual, (void**)&pDevDualDisp);
            if (pDevDualDisp)
            {
                const wchar_t* fallbackNames[] = {
                    L"Backplane", L"CompactBus", L"PointBus", L"Chassis", L"BP"
                };
                const int numFallbacks = sizeof(fallbackNames) / sizeof(fallbackNames[0]);

                auto tryDispid38 = [&](const wchar_t* name) -> bool
                {
                    VARIANT argName;
                    VariantInit(&argName);
                    argName.vt = VT_BSTR;
                    argName.bstrVal = SysAllocString(name);
                    DISPPARAMS dp = { &argName, nullptr, 1, 0 };
                    VARIANT varBus;
                    VariantInit(&varBus);
                    HRESULT hr38 = pDevDualDisp->Invoke(38, IID_NULL, LOCALE_USER_DEFAULT,
                        DISPATCH_PROPERTYGET, &dp, &varBus, nullptr, nullptr);

                    bool found = false;
                    if (SUCCEEDED(hr38) && (varBus.vt == VT_DISPATCH || varBus.vt == VT_UNKNOWN) && varBus.punkVal)
                    {
                        pBackplaneBus = varBus.punkVal;
                        pBackplaneBus->AddRef();
                        busLabel = name;
                        Log(L"[BP]   >> Found backplane bus via DISPID 38(\"%s\")", name);
                        found = true;
                    }
                    VariantClear(&varBus);
                    SysFreeString(argName.bstrVal);
                    return found;
                };

                for (int pn = 0; pn < numFallbacks && !pBackplaneBus; pn++)
                    tryDispid38(fallbackNames[pn]);

                pDevDualDisp->Release();
            }
        }
        pBackplanePort->Release();

        if (!pBackplaneBus)
        {
            Log(L"[BP]   Could not find backplane bus");
            pDevice->Release();
            continue;
        }

        std::wstring sinkLabel = devName;
        if (!busLabel.empty())
            sinkLabel += L"/" + busLabel;

        void* pBPEnum = nullptr;
        HRESULT hrEnum = pBackplaneBus->QueryInterface(IID_IOnlineEnumeratorTypeLib, &pBPEnum);
        Log(L"[BP]   QI bus for enumerator: hr=0x%08x", hrEnum);

        if (FAILED(hrEnum) || !pBPEnum)
        {
            Log(L"[BP]   Bus doesn't support enumerator");
            pBackplaneBus->Release();
            pDevice->Release();
            continue;
        }

        IDispatch* pBusDisp = nullptr;
        pBackplaneBus->QueryInterface(IID_ITopologyBus, (void**)&pBusDisp);
        if (!pBusDisp)
            pBackplaneBus->QueryInterface(IID_ITopologyObject, (void**)&pBusDisp);

        IUnknown* pBPPath = nullptr;
        if (pBusDisp)
        {
            pBPPath = DispatchGetPath(pBusDisp);
            Log(L"[BP]   Bus path: 0x%p", pBPPath);
            pBusDisp->Release();
        }

        if (!pBPPath)
        {
            Log(L"[BP]   Could not get bus path");
            pBackplaneBus->Release();
            pDevice->Release();
            continue;
        }

        DualEventSink* pSink = new DualEventSink(sinkLabel.c_str());
        int cpCount = 0;
        {
            IConnectionPointContainer* pCPC = nullptr;
            pBackplaneBus->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
            if (pCPC)
            {
                IConnectionPoint* pCP = nullptr;
                if (SUCCEEDED(pCPC->FindConnectionPoint(IID_ITopologyBusEvents, &pCP)) && pCP)
                {
                    DWORD cookie = 0;
                    HRESULT hrAdv = pCP->Advise(static_cast<ITopologyBusEvents*>(pSink), &cookie);
                    cpCount++;
                    if (SUCCEEDED(hrAdv))
                        g_connectionPoints.push_back({pCP, cookie});
                    else
                        pCP->Release();
                }
                pCPC->Release();
            }
        }
        Log(L"[BP]   Connected %d CPs for \"%s\"", cpCount, sinkLabel.c_str());

        HRESULT hrStart = TryStartAtSlot(pBPEnum, pBPPath, 7);
        Log(L"[BP]   Start(bus path): hr=0x%08x", hrStart);
        if (SUCCEEDED(hrStart))
        {
            startedCount++;
            Log(L"[BP]   >> Backplane bus browse STARTED for \"%s\"", sinkLabel.c_str());
        }

        g_enumerators.push_back({pBPEnum, pSink});

        SafeRelease(pBPPath, L"pBPPath");
        pBackplaneBus->Release();
        pDevice->Release();
    }

    pDevices->Release();
    pEthBusDisp->Release();
    } // end per-driver loop

    Log(L"[BP] DoBackplaneBrowse done: %d started across %d drivers",
        startedCount, (int)g_pSharedConfig->drivers.size());
    return (startedCount > 0) ? S_OK : S_FALSE;
}

// ============================================================
// DoCleanupOnMainSTA — runs on MAIN STA thread
// ============================================================

HRESULT DoCleanupOnMainSTA()
{
    Log(L"[CLEANUP] DoCleanupOnMainSTA starting on TID=%d", GetCurrentThreadId());

    // 1. Stop all enumerators via vtable[8]
    int stopCount = 0;
    for (auto& ei : g_enumerators)
    {
        if (!ei.pEnumInterface) continue;
        __try
        {
            typedef HRESULT (__stdcall *StopFunc)(void* pThis);
            void** vtable = *(void***)ei.pEnumInterface;
            StopFunc pfn = (StopFunc)vtable[8];
            HRESULT hr = pfn(ei.pEnumInterface);
            Log(L"[CLEANUP] Stop enumerator 0x%p: hr=0x%08x", ei.pEnumInterface, hr);
            stopCount++;
        }
        __except (EXCEPTION_EXECUTE_HANDLER)
        {
            Log(L"[CLEANUP] Stop enumerator 0x%p: SEH exception (ignored)", ei.pEnumInterface);
        }
    }
    Log(L"[CLEANUP] Stopped %d enumerators", stopCount);

    // 2. Unadvise all connection points
    int unadviseCount = 0;
    for (auto& cpi : g_connectionPoints)
    {
        if (!cpi.pCP) continue;
        __try
        {
            HRESULT hr = cpi.pCP->Unadvise(cpi.cookie);
            Log(L"[CLEANUP] Unadvise CP 0x%p cookie=%d: hr=0x%08x", cpi.pCP, cpi.cookie, hr);
            unadviseCount++;
        }
        __except (EXCEPTION_EXECUTE_HANDLER)
        {
            Log(L"[CLEANUP] Unadvise CP 0x%p: SEH exception (ignored)", cpi.pCP);
        }
    }
    Log(L"[CLEANUP] Unadvised %d connection points", unadviseCount);

    // 3. Release all CPs
    for (auto& cpi : g_connectionPoints)
    {
        SafeRelease(cpi.pCP, L"cleanup-CP");
        cpi.pCP = nullptr;
    }
    g_connectionPoints.clear();

    // 4. Release all enumerators and sinks
    for (auto& ei : g_enumerators)
    {
        if (ei.pEnumInterface)
        {
            SafeRelease((IUnknown*)ei.pEnumInterface, L"cleanup-Enum");
            ei.pEnumInterface = nullptr;
        }
        if (ei.pSink)
        {
            ei.pSink->Release();
            ei.pSink = nullptr;
        }
    }
    g_enumerators.clear();

    // 5. Clear captured buses
    for (auto* pBus : g_capturedBuses)
        SafeRelease(pBus, L"cleanup-Bus");
    g_capturedBuses.clear();

    // 6. Clear global persistent pointers
    g_pMainSink = nullptr;
    g_pMainEnumUnk = nullptr;

    Log(L"[CLEANUP] DoCleanupOnMainSTA complete");
    return S_OK;
}

// ============================================================
// DoMainSTABrowse — runs on the MAIN STA thread
// ============================================================

HRESULT DoMainSTABrowse()
{
    Log(L"[MAIN-STA] DoMainSTABrowse starting on TID=%d", GetCurrentThreadId());

    if (!g_pSharedConfig)
    {
        Log(L"[MAIN-STA] ERROR: no shared config");
        return E_INVALIDARG;
    }

    HRESULT hr;

    IHarmonyConnector* pHarmony = nullptr;
    hr = CoCreateInstance(CLSID_HarmonyServices, NULL, CLSCTX_ALL,
                          IID_IHarmonyConnector, (void**)&pHarmony);
    if (FAILED(hr))
    {
        Log(L"[MAIN-STA] FAIL HarmonyServices: 0x%08x", hr);
        return hr;
    }
    pHarmony->SetServerOptions(0, L"");
    Log(L"[MAIN-STA] HarmonyServices OK");

    IRSTopologyGlobals* pGlobals = nullptr;
    {
        IUnknown* pUnk = nullptr;
        hr = pHarmony->GetSpecialObject(&CLSID_RSTopologyGlobals, &IID_IRSTopologyGlobals, &pUnk);
        if (FAILED(hr))
        {
            Log(L"[MAIN-STA] FAIL TopologyGlobals: 0x%08x", hr);
            pHarmony->Release();
            return hr;
        }
        pUnk->QueryInterface(IID_IRSTopologyGlobals, (void**)&pGlobals);
        pUnk->Release();
    }
    Log(L"[MAIN-STA] TopologyGlobals OK");

    IRSProject* pProject = nullptr;
    {
        IRSProjectGlobal* pPG = nullptr;
        hr = CoCreateInstance(CLSID_RSProjectGlobal, NULL, CLSCTX_ALL,
                              IID_IRSProjectGlobal, (void**)&pPG);
        if (FAILED(hr))
        {
            Log(L"[MAIN-STA] FAIL ProjectGlobal: 0x%08x", hr);
            pGlobals->Release(); pHarmony->Release();
            return hr;
        }
        hr = pPG->OpenProject(L"", 0, NULL, NULL, &IID_IRSProject, (void**)&pProject);
        pPG->Release();
        if (FAILED(hr))
        {
            Log(L"[MAIN-STA] FAIL OpenProject: 0x%08x", hr);
            pGlobals->Release(); pHarmony->Release();
            return hr;
        }
    }
    Log(L"[MAIN-STA] Project OK");

    IUnknown* pWorkstation = nullptr;
    hr = pGlobals->GetThisWorkstationObject(pProject, &pWorkstation);
    if (FAILED(hr))
    {
        Log(L"[MAIN-STA] FAIL Workstation: 0x%08x", hr);
        pProject->Release(); pGlobals->Release(); pHarmony->Release();
        return hr;
    }
    Log(L"[MAIN-STA] Workstation OK");

    HRESULT hrStart = E_FAIL;
    int busesStarted = 0;

    for (auto& drv : g_pSharedConfig->drivers)
    {
    Log(L"[MAIN-STA] === Driver: %s ===", drv.name.c_str());

    IDispatch* pBusDisp = nullptr;
    IUnknown* pBusUnk = nullptr;
    {
        IDispatch* pWsDisp = nullptr;
        hr = pWorkstation->QueryInterface(IID_ITopologyDevice_Dual, (void**)&pWsDisp);
        if (SUCCEEDED(hr) && pWsDisp)
        {
            VARIANT argName;
            VariantInit(&argName);
            argName.vt = VT_BSTR;
            argName.bstrVal = SysAllocString(drv.name.c_str());
            DISPPARAMS dp = { &argName, nullptr, 1, 0 };
            VARIANT varBus;
            VariantInit(&varBus);
            hr = pWsDisp->Invoke(38, IID_NULL, LOCALE_USER_DEFAULT,
                                  DISPATCH_PROPERTYGET, &dp, &varBus, nullptr, nullptr);
            Log(L"[MAIN-STA] Bus(\"%s\"): hr=0x%08x vt=%d",
                drv.name.c_str(), hr, varBus.vt);
            if (SUCCEEDED(hr) && (varBus.vt == VT_DISPATCH || varBus.vt == VT_UNKNOWN))
            {
                pBusUnk = (varBus.vt == VT_DISPATCH) ? (IUnknown*)varBus.pdispVal : varBus.punkVal;
                if (pBusUnk) pBusUnk->AddRef();
                pBusUnk->QueryInterface(IID_ITopologyBus, (void**)&pBusDisp);
            }
            VariantClear(&argName);
            VariantClear(&varBus);
            pWsDisp->Release();
        }
    }

    if (!pBusDisp || !pBusUnk)
    {
        Log(L"[MAIN-STA] FAIL Could not get bus '%s' — skipping", drv.name.c_str());
        if (pBusDisp) pBusDisp->Release();
        if (pBusUnk) pBusUnk->Release();
        continue;
    }
    Log(L"[MAIN-STA] Bus OK: 0x%p", pBusDisp);

    std::wstring ethLabel = drv.name + L"/Ethernet";
    DualEventSink* pSink = new DualEventSink(ethLabel.c_str());
    {
        IConnectionPointContainer* pCPC = nullptr;
        hr = pBusUnk->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
        if (SUCCEEDED(hr) && pCPC)
        {
            IConnectionPoint* pCP = nullptr;
            hr = pCPC->FindConnectionPoint(IID_ITopologyBusEvents, &pCP);
            if (SUCCEEDED(hr) && pCP)
            {
                DWORD cookie = 0;
                hr = pCP->Advise(static_cast<IRSTopologyOnlineNotify*>(pSink), &cookie);
                Log(L"[MAIN-STA] Bus CP Advise: hr=0x%08x cookie=%d", hr, cookie);
                if (SUCCEEDED(hr))
                    g_connectionPoints.push_back({pCP, cookie});
                else
                    pCP->Release();
            }
            pCPC->Release();
        }
    }

    void* pBusEnum = nullptr;
    bool enumFromBus = false;
    IUnknown* pEnumUnk = nullptr;
    hr = pBusUnk->QueryInterface(IID_IOnlineEnumeratorTypeLib, &pBusEnum);
    Log(L"[MAIN-STA] QI bus for IID_IOnlineEnumeratorTypeLib: hr=0x%08x ptr=0x%p", hr, pBusEnum);

    if (FAILED(hr) || !pBusEnum)
    {
        hr = pBusUnk->QueryInterface(IID_IOnlineEnumerator, &pBusEnum);
        Log(L"[MAIN-STA] QI bus for IID_IOnlineEnumerator: hr=0x%08x ptr=0x%p", hr, pBusEnum);
    }

    if (SUCCEEDED(hr) && pBusEnum)
    {
        Log(L"[MAIN-STA] Got enumerator from bus via QI!");
        enumFromBus = true;
        pEnumUnk = (IUnknown*)pBusEnum;
    }
    else
    {
        Log(L"[MAIN-STA] Bus QI failed, falling back to CoCreateInstance(OnlineBusExt)");
        hr = CoCreateInstance(CLSID_OnlineBusExt, NULL, CLSCTX_INPROC_SERVER,
                              IID_IUnknown, (void**)&pEnumUnk);
        if (FAILED(hr))
            hr = CoCreateInstance(CLSID_OnlineBusExt, NULL, CLSCTX_ALL,
                                  IID_IUnknown, (void**)&pEnumUnk);
        Log(L"[MAIN-STA] OnlineBusExt fallback: hr=0x%08x ptr=0x%p", hr, pEnumUnk);
    }

    if (!pEnumUnk)
    {
        Log(L"[MAIN-STA] FAIL Could not get enumerator for '%s'", drv.name.c_str());
        pBusDisp->Release(); pBusUnk->Release();
        continue;
    }

    // Connect sink to enumerator CPs
    {
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
                int cpIdx = 0;
                while (pEnumCPEnum->Next(1, &pCP, &fetched) == S_OK && fetched > 0)
                {
                    DWORD enumCookie = 0;
                    hr = pCP->Advise(static_cast<IRSTopologyOnlineNotify*>(pSink), &enumCookie);
                    if (SUCCEEDED(hr))
                        g_connectionPoints.push_back({pCP, enumCookie});
                    else
                        pCP->Release();
                    cpIdx++;
                }
                Log(L"[MAIN-STA] Connected %d enumerator CPs", cpIdx);
                pEnumCPEnum->Release();
            }
            pEnumCPC->Release();
        }
    }

    IUnknown* pPathObject = nullptr;
    {
        VARIANT argFlags;
        VariantInit(&argFlags);
        argFlags.vt = VT_I4;
        argFlags.lVal = 0;
        DISPPARAMS dpPath = { &argFlags, nullptr, 1, 0 };
        VARIANT varPath;
        VariantInit(&varPath);
        hr = pBusDisp->Invoke(4, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_PROPERTYGET, &dpPath, &varPath, nullptr, nullptr);
        Log(L"[MAIN-STA] bus.path(flags=0): hr=0x%08x vt=%d", hr, varPath.vt);
        if (SUCCEEDED(hr) && (varPath.vt == VT_DISPATCH || varPath.vt == VT_UNKNOWN))
        {
            pPathObject = (varPath.vt == VT_DISPATCH) ? (IUnknown*)varPath.pdispVal : varPath.punkVal;
            if (pPathObject) pPathObject->AddRef();
        }
        VariantClear(&varPath);
    }

    if (!pPathObject)
    {
        Log(L"[MAIN-STA] FAIL Could not get path object for '%s'", drv.name.c_str());
        pBusDisp->Release(); pBusUnk->Release();
        pEnumUnk->Release(); delete pSink;
        continue;
    }
    Log(L"[MAIN-STA] Path object OK: 0x%p", pPathObject);

    HRESULT hrDrvStart = E_FAIL;
    if (enumFromBus)
    {
        pSink->DumpCounters(L"main-before-start");
        hrDrvStart = TryStartAtSlot(pBusEnum, pPathObject, 7);
        Log(L"[MAIN-STA] Start(bus.path) via bus-QI enum: hr=0x%08x", hrDrvStart);
        pSink->DumpCounters(L"main-after-start");

        if (SUCCEEDED(hrDrvStart))
            Log(L"[MAIN-STA] Browse started for '%s' (bus-QI path)!", drv.name.c_str());
        else
            Log(L"[MAIN-STA] Start via bus-QI FAILED: 0x%08x", hrDrvStart);
    }
    else
    {
        void* pRealEnum = nullptr;
        hr = pEnumUnk->QueryInterface(IID_IOnlineEnumerator, &pRealEnum);
        Log(L"[MAIN-STA] QI IOnlineEnumerator: hr=0x%08x", hr);

        if (SUCCEEDED(hr) && pRealEnum)
        {
            pSink->DumpCounters(L"main-before-start");
            hrDrvStart = TryStartAtSlot(pRealEnum, pPathObject, 7);
            Log(L"[MAIN-STA] Start(bus.path) via standalone: hr=0x%08x", hrDrvStart);
            pSink->DumpCounters(L"main-after-start");

            if (SUCCEEDED(hrDrvStart))
                Log(L"[MAIN-STA] Browse started for '%s' (standalone fallback)!", drv.name.c_str());
            else
                Log(L"[MAIN-STA] Start standalone FAILED: 0x%08x", hrDrvStart);

            if (FAILED(hrDrvStart))
                ((IUnknown*)pRealEnum)->Release();
        }
    }

    SafeRelease(pPathObject, L"pPathObject");

    g_enumerators.push_back({(void*)pEnumUnk, pSink});

    SafeRelease(pBusDisp, L"pBusDisp");
    SafeRelease(pBusUnk, L"pBusUnk");

    if (SUCCEEDED(hrDrvStart))
    {
        busesStarted++;
        hrStart = S_OK;
    }

    } // end per-driver loop

    SafeRelease(pWorkstation, L"pWorkstation");
    SafeRelease(pProject, L"pProject");
    SafeRelease(pGlobals, L"pGlobals");
    SafeRelease(pHarmony, L"pHarmony");

    Log(L"[MAIN-STA] DoMainSTABrowse done, %d/%d drivers started",
        busesStarted, (int)g_pSharedConfig->drivers.size());
    return (busesStarted > 0) ? S_OK : E_FAIL;
}

// ============================================================
// RunMonitorLoop — continuous browse mode
// ============================================================

void RunMonitorLoop(const HookConfig& config, IRSTopologyGlobals* pGlobals, const std::vector<BusInfo>& buses)
{
    Log(L"");
    Log(L"=== Monitor Mode: Continuous browse ===");

    bool busBrowseDone = false;
    bool backplaneBrowseDone = false;
    int snapshotNum = 0;
    int lastIdentified = 0;
    DWORD startTick = GetTickCount();

    while (true)
    {
        if (g_shouldStop)
        {
            Log(L"[MONITOR] DLL unload stop signal received");
            break;
        }
        if (g_pipeConnected) {
            if (PipeCheckStop()) {
                Log(L"[PIPE] Stop signal received");
                break;
            }
        } else {
            std::wstring stopPath = LogPath(config.logDir, L"hook_stop.txt");
            DWORD attrs = GetFileAttributesW(stopPath.c_str());
            if (attrs != INVALID_FILE_ATTRIBUTES) {
                Log(L"[MONITOR] Stop signal received (file)");
                break;
            }
        }

        MSG msg;
        while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
        { TranslateMessage(&msg); DispatchMessage(&msg); }

        Sleep(100);

        DWORD elapsed = GetTickCount() - startTick;

        if (elapsed > 0 && elapsed % 10000 < 100)
        {
            snapshotNum++;
            std::wstring snapFile;
            if (config.debugXml) {
                wchar_t fname[64];
                swprintf(fname, 64, L"hook_topo_monitor_%d.xml", snapshotNum);
                snapFile = LogPath(config.logDir, fname);
            } else {
                snapFile = LogPath(config.logDir, L"hook_topo_poll.xml");
            }
            if (SaveTopologyXML(pGlobals, snapFile.c_str()))
            {
                TopologyCounts c = CountDevicesInXML(snapFile.c_str());
                PipeSendTopology(snapFile.c_str());
                PipeSendStatus(c.totalDevices, c.identifiedDevices, (int)g_discoveredDevices.size());
                Log(L"[MONITOR] Snapshot %d @ %ds: %d devices, %d identified, %d events",
                    snapshotNum, elapsed / 1000, c.totalDevices, c.identifiedDevices,
                    (int)g_discoveredDevices.size());

                if (!busBrowseDone && c.identifiedDevices > 0)
                {
                    Log(L"[MONITOR] Devices identified — triggering bus browse");
                    g_capturedBuses.clear();
                    g_captureBuses = true;
                    HRESULT hrBus = ExecuteOnMainSTA(DoBusBrowse);
                    Log(L"[MONITOR] Bus browse: hr=0x%08x", hrBus);
                    busBrowseDone = true;
                }

                if (busBrowseDone && !backplaneBrowseDone && !g_capturedBuses.empty())
                {
                    g_captureBuses = false;
                    Log(L"[MONITOR] Captured %d buses — triggering backplane browse", (int)g_capturedBuses.size());
                    HRESULT hrBP = ExecuteOnMainSTA(DoBackplaneBrowse);
                    Log(L"[MONITOR] Backplane browse: hr=0x%08x", hrBP);
                    backplaneBrowseDone = true;
                }

                UpdateDeviceIPsFromXML(snapFile.c_str());

                std::wstring resultsPath = LogPath(config.logDir, L"hook_results.txt");
                FILE* rf = _wfopen(resultsPath.c_str(), L"w, ccs=UTF-8");
                if (rf)
                {
                    fwprintf(rf, L"MODE: monitor\n");
                    fwprintf(rf, L"SNAPSHOT: %d\n", snapshotNum);
                    fwprintf(rf, L"DEVICES_IDENTIFIED: %d\n", c.identifiedDevices);
                    fwprintf(rf, L"DEVICES_TOTAL: %d\n", c.totalDevices);
                    fwprintf(rf, L"EVENTS: %d\n", (int)g_discoveredDevices.size());
                    fwprintf(rf, L"ELAPSED: %d\n", elapsed / 1000);
                    for (const auto& kv : g_deviceDetails)
                    {
                        const DeviceInfo& d = kv.second;
                        fwprintf(rf, L"DEVICE: %s | %s\n",
                            d.ip.empty() ? L"(no IP)" : d.ip.c_str(),
                            d.productName.c_str());
                    }
                    fclose(rf);
                }

                lastIdentified = c.identifiedDevices;
            }
        }
    }

    Log(L"[MONITOR] Stopping — cleaning up");
    Log(L"Tracked: %d connection points, %d enumerators",
        (int)g_connectionPoints.size(), (int)g_enumerators.size());
    HRESULT hrClean = ExecuteOnMainSTA(DoCleanupOnMainSTA);
    Log(L"[MONITOR] Cleanup result: 0x%08x", hrClean);

    DWORD totalElapsed = GetTickCount() - startTick;
    {
        std::wstring finalSnap = LogPath(config.logDir, config.debugXml ? L"hook_topo_monitor_final.xml" : L"hook_topo_poll.xml");
        SaveTopologyXML(pGlobals, finalSnap.c_str());
        TopologyCounts fc = CountDevicesInXML(finalSnap.c_str());
        UpdateDeviceIPsFromXML(finalSnap.c_str());

        std::wstring resultsPath = LogPath(config.logDir, L"hook_results.txt");
        FILE* rf = _wfopen(resultsPath.c_str(), L"w, ccs=UTF-8");
        if (rf)
        {
            fwprintf(rf, L"MODE: monitor\n");
            fwprintf(rf, L"SNAPSHOT: %d (final)\n", snapshotNum);
            fwprintf(rf, L"DEVICES_IDENTIFIED: %d\n", fc.identifiedDevices);
            fwprintf(rf, L"DEVICES_TOTAL: %d\n", fc.totalDevices);
            fwprintf(rf, L"EVENTS: %d\n", (int)g_discoveredDevices.size());
            fwprintf(rf, L"ELAPSED: %d\n", totalElapsed / 1000);
            for (const auto& kv : g_deviceDetails)
            {
                const DeviceInfo& d = kv.second;
                fwprintf(rf, L"DEVICE: %s | %s\n",
                    d.ip.empty() ? L"(no IP)" : d.ip.c_str(),
                    d.productName.c_str());
            }
            fclose(rf);
        }

        Log(L"[MONITOR] Final: %d devices, %d identified, %d events, %ds elapsed",
            fc.totalDevices, fc.identifiedDevices, (int)g_discoveredDevices.size(), totalElapsed / 1000);
    }

    if (!config.debugXml)
    {
        std::wstring pollPath = LogPath(config.logDir, L"hook_topo_poll.xml");
        DeleteFileW(pollPath.c_str());
    }

    Log(L"[MONITOR] Monitor loop complete");
}
