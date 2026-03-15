#include "RSLinxHook_fwd.h"
#include "Logging.h"
#include "ComInterfaces.h"
#include "Config.h"
#include "SEHHelpers.h"
#include "EventSink.h"
#include "DispatchHelpers.h"
#include "TopologyXML.h"
#include "EngineHotLoad.h"
#include "STAHook.h"
#include "BrowseOperations.h"

// ============================================================
// Globals owned by DllMain.cpp
// ============================================================

HANDLE g_hWorkerThread = NULL;
volatile bool g_shouldStop = false;

// ============================================================
// AcquireNewBuses
// Acquire ITopologyBus objects for any driver in config that
// does not already have an entry in buses[].
// Uses fresh pProject/pWorkstation per call.
// ============================================================

static void AcquireNewBuses(HookConfig& config, IRSTopologyGlobals* pGlobals,
                             std::vector<BusInfo>& buses)
{
    IRSProject* pProject = nullptr;
    {
        IRSProjectGlobal* pPG = nullptr;
        HRESULT hr = CoCreateInstance(CLSID_RSProjectGlobal, NULL, CLSCTX_ALL,
                                      IID_IRSProjectGlobal, (void**)&pPG);
        if (FAILED(hr)) { Log(L"[FAIL] ProjectGlobal: 0x%08x", hr); return; }
        hr = pPG->OpenProject(L"", 0, NULL, NULL, &IID_IRSProject, (void**)&pProject);
        pPG->Release();
        if (FAILED(hr)) { Log(L"[FAIL] OpenProject: 0x%08x", hr); return; }
    }

    IUnknown* pWorkstation = nullptr;
    HRESULT hr = pGlobals->GetThisWorkstationObject(pProject, &pWorkstation);
    if (FAILED(hr)) { pProject->Release(); Log(L"[FAIL] GetWorkstation: 0x%08x", hr); return; }

    // Per-driver engine hot-load for new drivers
    for (auto& drv : config.drivers)
    {
        if (drv.newDriver)
        {
            g_engineDriverName = drv.name;
            Log(L"[ENGINE] Hot-loading driver '%s'...", drv.name.c_str());
            HRESULT hrEngine = ExecuteOnMainSTA(DoEngineHotLoadOnMainSTA);
            Log(L"[ENGINE] Hot-load '%s': 0x%08x", drv.name.c_str(), hrEngine);
        }
    }

    // Re-acquire workstation after any hot-loads
    if (config.newDriver())
    {
        pWorkstation->Release();
        pWorkstation = nullptr;
        hr = pGlobals->GetThisWorkstationObject(pProject, &pWorkstation);
        if (FAILED(hr) || !pWorkstation)
        {
            Log(L"[FAIL] Re-acquire workstation: 0x%08x", hr);
            pProject->Release();
            return;
        }
        Log(L"[OK] Re-acquired workstation after hot-load(s)");

        std::wstring verifyPath = LogPath(config.logDir, L"hook_topo_after_hotload.xml");
        SaveTopologyXML(pGlobals, verifyPath.c_str());
    }

    const int maxRetries = 12;
    const int retryDelayMs = 5000;

    for (auto& drv : config.drivers)
    {
        // Skip if already acquired
        bool alreadyHave = false;
        for (auto& b : buses) if (b.driverName == drv.name) { alreadyHave = true; break; }
        if (alreadyHave) continue;

        IDispatch* pBusDisp = nullptr;
        IUnknown* pBusUnk = nullptr;

        Log(L"[BUS] Acquiring bus for driver '%s'...", drv.name.c_str());
        bool addPortAttempted = false;

        for (int attempt = 0; attempt < maxRetries && !pBusDisp; attempt++)
        {
            if (attempt == 1 && !drv.newDriver)
            {
                g_engineDriverName = drv.name;
                Log(L"[ENGINE] Fallback hot-load for '%s'...", drv.name.c_str());
                ExecuteOnMainSTA(DoEngineHotLoadOnMainSTA);
                pWorkstation->Release();
                hr = pGlobals->GetThisWorkstationObject(pProject, &pWorkstation);
                if (FAILED(hr)) break;
            }

            // Strategy 1: DISPID 38 (name-based bus lookup)
            {
                IDispatch* pWsDisp = nullptr;
                pWorkstation->QueryInterface(IID_ITopologyDevice_Dual, (void**)&pWsDisp);
                if (pWsDisp)
                {
                    VARIANT argName; VariantInit(&argName);
                    argName.vt = VT_BSTR;
                    argName.bstrVal = SysAllocString(drv.name.c_str());
                    DISPPARAMS dp = { &argName, nullptr, 1, 0 };
                    VARIANT varBus; VariantInit(&varBus);
                    EXCEPINFO excep = {};
                    hr = pWsDisp->Invoke(38, IID_NULL, LOCALE_USER_DEFAULT,
                                          DISPATCH_PROPERTYGET, &dp, &varBus, &excep, nullptr);
                    if (SUCCEEDED(hr) && (varBus.vt == VT_DISPATCH || varBus.vt == VT_UNKNOWN))
                    {
                        pBusUnk = (varBus.vt == VT_DISPATCH)
                            ? (IUnknown*)varBus.pdispVal : varBus.punkVal;
                        if (pBusUnk) pBusUnk->AddRef();
                        if (pBusUnk) pBusUnk->QueryInterface(IID_ITopologyBus, (void**)&pBusDisp);
                    }
                    if (pBusDisp) Log(L"[OK] Got bus via DISPID 38(\"%s\")", drv.name.c_str());
                    else
                    {
                        Log(L"[BUS] DISPID 38(\"%s\"): failed hr=0x%08x vt=%d",
                            drv.name.c_str(), hr, varBus.vt);
                        if (excep.bstrDescription) SysFreeString(excep.bstrDescription);
                        if (excep.bstrSource) SysFreeString(excep.bstrSource);
                        if (excep.bstrHelpFile) SysFreeString(excep.bstrHelpFile);
                    }
                    VariantClear(&argName); VariantClear(&varBus);
                    pWsDisp->Release();
                }
                if (pBusDisp) break;
            }

            // Strategy 2: Enumerate workstation Busses() collection (DISPID 51)
            {
                IDispatch* pWsDisp = nullptr;
                pWorkstation->QueryInterface(IID_ITopologyDevice_Dual, (void**)&pWsDisp);
                if (pWsDisp)
                {
                    IDispatch* pBusColl = DispatchGetCollection(pWsDisp, 51);
                    if (pBusColl)
                    {
                        auto busList = EnumerateCollection(pBusColl);
                        for (size_t i = 0; i < busList.size(); i++)
                        {
                            std::wstring name = DispatchGetString(busList[i], 1);
                            if (_wcsicmp(name.c_str(), drv.name.c_str()) == 0)
                            {
                                IUnknown* pUnk = nullptr;
                                busList[i]->QueryInterface(IID_IUnknown, (void**)&pUnk);
                                if (pUnk)
                                {
                                    pUnk->QueryInterface(IID_ITopologyBus, (void**)&pBusDisp);
                                    pBusUnk = pUnk;
                                }
                                Log(L"[OK] Got bus via Busses() enumeration: \"%s\"", name.c_str());
                            }
                        }
                        for (auto* p : busList) p->Release();
                        pBusColl->Release();
                    }
                    pWsDisp->Release();
                }
                if (pBusDisp) break;
            }

            // Strategy 3: BindToObject on IRSProject
            {
                IUnknown* pBound = nullptr;
                hr = pProject->BindToObject(drv.name.c_str(), &pBound);
                if (SUCCEEDED(hr) && pBound)
                {
                    pBound->QueryInterface(IID_ITopologyBus, (void**)&pBusDisp);
                    if (pBusDisp)
                    {
                        pBusUnk = pBound;
                        Log(L"[OK] Got bus via BindToObject(\"%s\")", drv.name.c_str());
                        break;
                    }
                    pBound->Release();
                }
                if (pBusDisp) break;
            }

            // Strategy 4: AddPort
            if (!addPortAttempted)
            {
                addPortAttempted = true;
                Log(L"[BUS] Attempting AddPort(\"%s\") via vtable[14]...", drv.name.c_str());

                IUnknown* pDevVtable = nullptr;
                hr = pWorkstation->QueryInterface(IID_IRSTopologyDevice, (void**)&pDevVtable);
                if (SUCCEEDED(hr) && pDevVtable)
                {
                    struct { GUID clsid; const wchar_t* name; } clsids[] = {
                        { CLSID_EthernetBus,  L"EthernetBus" },
                        { CLSID_EthernetPort, L"EthernetPort" },
                    };

                    for (int ci = 0; ci < 2 && !pBusDisp; ci++)
                    {
                        IUnknown* pNewObj = nullptr;
                        IID iidUnk = IID_IUnknown;
                        hr = TryVtableAddPort(pDevVtable, 14, &clsids[ci].clsid,
                                               drv.name.c_str(), &iidUnk, &pNewObj);
                        Log(L"[BUS] AddPort[14] with %s: hr=0x%08x obj=0x%p",
                            clsids[ci].name, hr, pNewObj);

                        if (FAILED(hr) || !pNewObj) continue;

                        IDispatch* pDisp = nullptr;
                        HRESULT hrQI;

                        hrQI = pNewObj->QueryInterface(IID_ITopologyBus, (void**)&pBusDisp);
                        Log(L"[BUS]   QI ITopologyBus: hr=0x%08x", hrQI);
                        if (pBusDisp) { pBusUnk = pNewObj; pBusUnk->AddRef(); Log(L"[OK] AddPort returned bus directly!"); }

                        if (!pBusDisp)
                        {
                            IUnknown* pPortVtable = nullptr;
                            hrQI = pNewObj->QueryInterface(IID_IRSTopologyPort, (void**)&pPortVtable);
                            Log(L"[BUS]   QI IRSTopologyPort: hr=0x%08x", hrQI);
                            if (pPortVtable)
                            {
                                IUnknown* pBusRaw = nullptr;
                                HRESULT hrGB = TryVtableGetObject(pPortVtable, 10, &pBusRaw);
                                Log(L"[BUS]   GetBus[10]: hr=0x%08x bus=0x%p", hrGB, pBusRaw);
                                if (SUCCEEDED(hrGB) && pBusRaw)
                                {
                                    pBusRaw->QueryInterface(IID_ITopologyBus, (void**)&pBusDisp);
                                    if (pBusDisp) { pBusUnk = pBusRaw; Log(L"[OK] Created bus via AddPort+GetBus!"); }
                                    else pBusRaw->Release();
                                }
                                pPortVtable->Release();
                            }
                        }

                        if (!pBusDisp)
                        {
                            hrQI = pNewObj->QueryInterface(IID_IDispatch, (void**)&pDisp);
                            if (pDisp)
                            {
                                std::wstring name = DispatchGetString(pDisp, 1);
                                Log(L"[BUS]   Object Name (DISPID 1): \"%s\"", name.c_str());
                                pDisp->Release();
                            }
                        }

                        pNewObj->Release();
                    }
                    pDevVtable->Release();
                }
                else
                {
                    Log(L"[BUS] QI IRSTopologyDevice on workstation: hr=0x%08x", hr);
                }

                if (pBusDisp) break;
            }

            if (!pBusDisp && attempt < maxRetries - 1)
            {
                Log(L"[WAIT] Bus \"%s\" not found (attempt %d/%d), retrying in %ds...",
                    drv.name.c_str(), attempt + 1, maxRetries, retryDelayMs / 1000);
                Sleep(retryDelayMs);
            }
        }

        if (pBusDisp && pBusUnk)
        {
            buses.push_back({drv.name, pBusDisp, pBusUnk});
            Log(L"[OK] Acquired bus '%s': IDispatch=0x%p", drv.name.c_str(), pBusDisp);
        }
        else
        {
            Log(L"[FAIL] Could not get/create bus \"%s\" after %d attempts",
                drv.name.c_str(), maxRetries);
        }
    }

    pWorkstation->Release();
    pProject->Release();

    if (buses.empty())
        Log(L"[FAIL] No buses acquired  - cannot browse");
    else
        Log(L"[OK] Have %d bus(es) total", (int)buses.size());
}

// ============================================================
// RunBrowsePhases
// Execute phases 1-4b on the current config/buses and write
// the results file. Populates g_browsedDrivers /
// g_browsedBackplanes after each phase.
// ============================================================

static void RunBrowsePhases(HookConfig& config, IRSTopologyGlobals* pGlobals,
                             std::vector<BusInfo>& buses)
{
    g_discoveredDevices.clear();

    // =============================================================
    // Phase 1: ConnectNewDevice
    // =============================================================
    {
        Log(L"");
        Log(L"=== Phase 1: ConnectNewDevice (smart, %d drivers) ===", (int)buses.size());

        static const wchar_t* UNRECOGNIZED_DEVICE_GUID = L"{00000004-5D68-11CF-B4B9-C46F03C10000}";
        int totalAdded = 0, totalExisting = 0, totalFailed = 0;

        for (auto& bus : buses)
        {
            const DriverEntry* pDrv = nullptr;
            for (auto& d : config.drivers)
                if (d.name == bus.driverName) { pDrv = &d; break; }
            if (!pDrv || pDrv->ipAddresses.empty())
            {
                Log(L"  [%s] No IPs to add  - using Node Table", bus.driverName.c_str());
                continue;
            }

            Log(L"  [%s] Adding %d IPs...", bus.driverName.c_str(), (int)pDrv->ipAddresses.size());
            int addedCount = 0, existingCount = 0, failedCount = 0;

            for (size_t i = 0; i < pDrv->ipAddresses.size(); i++)
            {
                VARIANT args[6];
                VariantInit(&args[5]); args[5].vt = VT_I4; args[5].lVal = 0;
                VariantInit(&args[4]); args[4].vt = VT_BSTR;
                args[4].bstrVal = SysAllocString(UNRECOGNIZED_DEVICE_GUID);
                VariantInit(&args[3]);
                VariantInit(&args[2]); args[2].vt = VT_BSTR;
                args[2].bstrVal = SysAllocString(L"Device");
                VariantInit(&args[1]); args[1].vt = VT_BSTR;
                args[1].bstrVal = SysAllocString(L"A");
                VariantInit(&args[0]); args[0].vt = VT_BSTR;
                args[0].bstrVal = SysAllocString(pDrv->ipAddresses[i].c_str());

                DISPPARAMS dp = { args, nullptr, 6, 0 };
                EXCEPINFO excep = {};
                UINT argErr = 0;
                VARIANT result;
                VariantInit(&result);

                HRESULT hr = bus.pBusDisp->Invoke(54, IID_NULL, LOCALE_USER_DEFAULT,
                                                   DISPATCH_METHOD, &dp, &result, &excep, &argErr);

                if (SUCCEEDED(hr))
                {
                    Log(L"    [OK] %s added", pDrv->ipAddresses[i].c_str());
                    addedCount++;
                }
                else if (hr == DISP_E_EXCEPTION)
                {
                    Log(L"    [SKIP] %s already exists", pDrv->ipAddresses[i].c_str());
                    existingCount++;
                    if (excep.bstrDescription) SysFreeString(excep.bstrDescription);
                    if (excep.bstrSource) SysFreeString(excep.bstrSource);
                    if (excep.bstrHelpFile) SysFreeString(excep.bstrHelpFile);
                }
                else
                {
                    Log(L"    [FAIL] %s: hr=0x%08x", pDrv->ipAddresses[i].c_str(), hr);
                    failedCount++;
                    if (excep.bstrDescription) SysFreeString(excep.bstrDescription);
                    if (excep.bstrSource) SysFreeString(excep.bstrSource);
                    if (excep.bstrHelpFile) SysFreeString(excep.bstrHelpFile);
                }

                for (int a = 0; a < 6; a++) VariantClear(&args[a]);
                VariantClear(&result);
            }

            Log(L"  [%s] Added: %d, Existing: %d, Failed: %d",
                bus.driverName.c_str(), addedCount, existingCount, failedCount);
            totalAdded += addedCount;
            totalExisting += existingCount;
            totalFailed += failedCount;
        }

        Log(L"  Total: Added: %d, Existing: %d, Failed: %d", totalAdded, totalExisting, totalFailed);
    }

    // Initial topology snapshot
    Log(L"");
    {
        std::wstring beforePath = LogPath(config.logDir, config.debugXml ? L"hook_topo_before.xml" : L"hook_topo_poll.xml");
        SaveTopologyXML(pGlobals, beforePath.c_str());
        TopologyCounts before = CountDevicesInXML(beforePath.c_str());
        PipeSendTopology(beforePath.c_str());
        PipeSendStatus(before.totalDevices, before.identifiedDevices, (int)g_discoveredDevices.size());
        Log(L"Topology BEFORE browse: %d devices, %d identified",
            before.totalDevices, before.identifiedDevices);
    }

    // =============================================================
    // Phase 2: Main-STA browse
    // =============================================================
    {
        Log(L"");
        Log(L"=== Phase 2: Main-STA browse via thread hook ===");

        HRESULT hrBrowse = ExecuteOnMainSTA(DoMainSTABrowse);
        Log(L"Phase 2 result: 0x%08x", hrBrowse);

        if (FAILED(hrBrowse))
            Log(L"[WARN] Main-STA browse failed  - identification may not trigger");
    }

    // Monitor mode: enter continuous loop then return
    if (config.mode == HookMode::Monitor)
    {
        Log(L"=== Entering Monitor Mode ===");
        RunMonitorLoop(config, pGlobals, buses);
        return;
    }

    // =============================================================
    // Phase 3: Topology polling
    // =============================================================
    {
        Log(L"");
        Log(L"=== Phase 3: Topology polling (2s interval, early exit on target) ===");
        bool targetIdentified = false;
        int enumBaseline = (int)g_enumerators.size();

        DWORD startTick = GetTickCount();
        while (!g_shouldStop && GetTickCount() - startTick < 30000)
        {
            DWORD elapsed = GetTickCount() - startTick;

            if (elapsed >= 3000 && EnumeratorsCycledSince(enumBaseline))
            {
                Log(L"  >> All Phase 2 enumerators cycled at %dms  - advancing", elapsed);
                break;
            }

            if (elapsed > 0 && elapsed % 2000 < 100)
            {
                std::wstring pollFile;
                if (config.debugXml) {
                    wchar_t fname[64];
                    swprintf(fname, 64, L"hook_topo_%ds.xml", elapsed / 1000);
                    pollFile = LogPath(config.logDir, fname);
                } else {
                    pollFile = LogPath(config.logDir, L"hook_topo_poll.xml");
                }
                if (SaveTopologyXML(pGlobals, pollFile.c_str()))
                {
                    TopologyCounts c = CountDevicesInXML(pollFile.c_str());
                    PipeSendTopology(pollFile.c_str());
                    PipeSendStatus(c.totalDevices, c.identifiedDevices, (int)g_discoveredDevices.size());
                    int cycled, total;
                    GetEnumeratorStatusSince(enumBaseline, cycled, total);
                    Log(L"  [%ds] %d devices, %d identified, %d events, %d/%d enumerators cycled",
                        elapsed / 1000, c.totalDevices, c.identifiedDevices,
                        (int)g_discoveredDevices.size(), cycled, total);

                    std::vector<std::wstring> allIPs = config.allIPs();
                    if (elapsed >= 3000 && !targetIdentified && !allIPs.empty() &&
                        IsTargetIdentifiedInXML(pollFile.c_str(), allIPs))
                    {
                        targetIdentified = true;
                        Log(L"  >> Target IPs identified at %ds  - exiting Phase 3 early", elapsed / 1000);
                        break;
                    }
                }
            }

            MSG msg;
            while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
            { TranslateMessage(&msg); DispatchMessage(&msg); }

            Sleep(100);
        }

        if (!targetIdentified)
            Log(L"  Target not identified after 30s  - proceeding anyway");
    }

    // Mark Phase 2+3 complete for all drivers
    for (auto& drv : config.drivers)
        g_browsedDrivers.insert(drv.name);

    // =============================================================
    // Phase 4: Bus browse + Phase 5 + Phase 4b + Phase 5b
    // =============================================================
    {
        std::wstring midPath = LogPath(config.logDir, config.debugXml ? L"hook_topo_mid.xml" : L"hook_topo_poll.xml");
        SaveTopologyXML(pGlobals, midPath.c_str());
        TopologyCounts pre = CountDevicesInXML(midPath.c_str());
        Log(L"Topology after bus browse: %d devices, %d identified",
            pre.totalDevices, pre.identifiedDevices);

        if (pre.identifiedDevices > 0)
        {
            Log(L"");
            Log(L"=== Phase 4: Bus browse (backplanes) ===");

            g_capturedBuses.clear();
            g_captureBuses = true;

            int phase4Baseline = (int)g_enumerators.size();
            HRESULT hrBus = ExecuteOnMainSTA(DoBusBrowse);
            Log(L"Phase 4 result: 0x%08x", hrBus);

            if (SUCCEEDED(hrBus))
            {
                Log(L"");
                Log(L"=== Phase 5: Bus browse polling (event-driven, max 30s) ===");

                DWORD busStart = GetTickCount();
                while (!g_shouldStop && GetTickCount() - busStart < 30000)
                {
                    DWORD elapsed = GetTickCount() - busStart;

                    if (elapsed >= 2000 && EnumeratorsCycledSince(phase4Baseline))
                    {
                        Log(L"  >> All bus enumerators cycled at %dms  - advancing", elapsed);
                        break;
                    }

                    if (elapsed > 0 && elapsed % 2000 < 100)
                    {
                        std::wstring pollFile;
                        if (config.debugXml) {
                            wchar_t fname[64];
                            swprintf(fname, 64, L"hook_topo_bus_%ds.xml", elapsed / 1000);
                            pollFile = LogPath(config.logDir, fname);
                        } else {
                            pollFile = LogPath(config.logDir, L"hook_topo_poll.xml");
                        }
                        if (SaveTopologyXML(pGlobals, pollFile.c_str()))
                        {
                            TopologyCounts c = CountDevicesInXML(pollFile.c_str());
                            PipeSendTopology(pollFile.c_str());
                            PipeSendStatus(c.totalDevices, c.identifiedDevices, (int)g_discoveredDevices.size());
                            int cycled, total;
                            GetEnumeratorStatusSince(phase4Baseline, cycled, total);
                            Log(L"  [%ds] %d devices, %d identified, %d events, %d/%d enumerators cycled",
                                elapsed / 1000, c.totalDevices, c.identifiedDevices,
                                (int)g_discoveredDevices.size(), cycled, total);
                        }
                    }

                    MSG msg;
                    while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
                    { TranslateMessage(&msg); DispatchMessage(&msg); }

                    Sleep(100);
                }
            }
            else
            {
                Log(L"[WARN] Bus browse failed  - skipping Phase 5");
            }

            g_captureBuses = false;
            Log(L"");
            Log(L"=== Phase 4b: Backplane bus browse ===");
            Log(L"Captured %d backplane buses from events", (int)g_capturedBuses.size());
            int phase4bBaseline = (int)g_enumerators.size();
            HRESULT hrBP = ExecuteOnMainSTA(DoBackplaneBrowse);
            Log(L"Phase 4b result: 0x%08x", hrBP);

            if (SUCCEEDED(hrBP))
            {
                Log(L"");
                Log(L"=== Phase 5b: Backplane module polling (event-driven, max 30s) ===");

                DWORD bpStart = GetTickCount();
                while (!g_shouldStop && GetTickCount() - bpStart < 30000)
                {
                    DWORD elapsed = GetTickCount() - bpStart;

                    if (elapsed >= 2000 && EnumeratorsCycledSince(phase4bBaseline))
                    {
                        Log(L"  >> All backplane enumerators cycled at %dms  - advancing", elapsed);
                        break;
                    }

                    if (elapsed > 0 && elapsed % 2000 < 100)
                    {
                        std::wstring pollFile;
                        if (config.debugXml) {
                            wchar_t fname[64];
                            swprintf(fname, 64, L"hook_topo_bp_%ds.xml", elapsed / 1000);
                            pollFile = LogPath(config.logDir, fname);
                        } else {
                            pollFile = LogPath(config.logDir, L"hook_topo_poll.xml");
                        }
                        if (SaveTopologyXML(pGlobals, pollFile.c_str()))
                        {
                            TopologyCounts c = CountDevicesInXML(pollFile.c_str());
                            PipeSendTopology(pollFile.c_str());
                            PipeSendStatus(c.totalDevices, c.identifiedDevices, (int)g_discoveredDevices.size());
                            int cycled, total;
                            GetEnumeratorStatusSince(phase4bBaseline, cycled, total);
                            Log(L"  [%ds] %d devices, %d identified, %d events, %d/%d enumerators cycled",
                                elapsed / 1000, c.totalDevices, c.identifiedDevices,
                                (int)g_discoveredDevices.size(), cycled, total);
                        }
                    }

                    MSG msg;
                    while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
                    { TranslateMessage(&msg); DispatchMessage(&msg); }

                    Sleep(100);
                }
            }

            // Mark Phase 4+4b complete for all IPs
            for (auto& drv : config.drivers)
                for (auto& ip : drv.ipAddresses)
                    g_browsedBackplanes.insert(ip);
        }
        else
        {
            Log(L"");
            Log(L"=== Phase 4: Skipped (no identified devices for bus browse) ===");
        }
    }

    // Final results
    Log(L"");
    Log(L"=== Final Results ===");

    {
        std::wstring afterPath = LogPath(config.logDir, config.debugXml ? L"hook_topo_after.xml" : L"hook_topo_poll.xml");
        SaveTopologyXML(pGlobals, afterPath.c_str());

        TopologyCounts fc = CountDevicesInXML(afterPath.c_str());
        // Populate caches before tree walk so WalkTopologyTree has fresh IP/classname/slot data
        UpdateDeviceIPsFromXML(afterPath.c_str());
        PopulateQueryCache(afterPath.c_str());
        WalkTopologyTree(pGlobals);
        if (config.debugXml)
            PipeSendTopology(afterPath.c_str());
        PipeSendStatus(fc.totalDevices, fc.identifiedDevices, (int)g_discoveredDevices.size());
        std::vector<std::wstring> allIPs = config.allIPs();
        bool targetFound = !allIPs.empty() && IsTargetIdentifiedInXML(afterPath.c_str(), allIPs);
        Log(L"Final topology: %d devices, %d identified",
            fc.totalDevices, fc.identifiedDevices);
        if (!allIPs.empty())
            Log(L"Target IPs identified: %s", targetFound ? L"YES" : L"NO");
        Log(L"Events received: %d", (int)g_discoveredDevices.size());

        std::wstring resultsPath = LogPath(config.logDir, L"hook_results.txt");
        FILE* resultFile = _wfopen(resultsPath.c_str(), L"w, ccs=UTF-8");
        if (resultFile)
        {
            fwprintf(resultFile, L"DRIVERS: %d\n", (int)config.drivers.size());
            fwprintf(resultFile, L"DEVICES_IDENTIFIED: %d\n", fc.identifiedDevices);
            fwprintf(resultFile, L"DEVICES_TOTAL: %d\n", fc.totalDevices);
            fwprintf(resultFile, L"EVENTS: %d\n", (int)g_discoveredDevices.size());
            if (!allIPs.empty())
                fwprintf(resultFile, L"TARGET: %s\n", allIPs[0].c_str());
            fwprintf(resultFile, L"TARGET_STATUS: %s\n", targetFound ? L"IDENTIFIED" : L"NOT_FOUND");
            for (const auto& kv : g_deviceDetails)
            {
                const DeviceInfo& d = kv.second;
                fwprintf(resultFile, L"DEVICE: %s | %s\n",
                    d.ip.empty() ? L"(no IP)" : d.ip.c_str(),
                    d.productName.c_str());
            }
            fclose(resultFile);
        }
        Log(L"[OK] Results file written");

        if (!config.debugXml)
            DeleteFileW(afterPath.c_str());
    }
}

// ============================================================
// HandleQuery
// Respond to a Q|path command: look up the device in topology
// XML, triggering browse if the path hasn't been browsed yet.
// ============================================================

// Build the cache key used in g_queryCache
static std::wstring MakeCacheKey(const std::wstring& ip, const std::wstring& portName, int slot)
{
    if (portName.empty()) return ip;
    wchar_t buf[512];
    swprintf(buf, 512, L"%s\\%s\\%d", ip.c_str(), portName.c_str(), slot);
    return buf;
}

// Wait for enumerators added since 'baseline' to complete (max 30s)
static void WaitForEnumerators(int baseline)
{
    DWORD t0 = GetTickCount();
    while (!g_shouldStop && GetTickCount() - t0 < 30000)
    {
        if (GetTickCount() - t0 >= 2000 && EnumeratorsCycledSince(baseline)) break;
        MSG msg;
        while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
        { TranslateMessage(&msg); DispatchMessage(&msg); }
        Sleep(100);
    }
}

static void HandleQuery(const char* path, IRSTopologyGlobals* pGlobals,
                        HookConfig& config, std::vector<BusInfo>& buses)
{
    // Parse path: ip[\portName[\slot]]
    std::string pathStr(path);
    if (!pathStr.empty() && pathStr.back() == '\r') pathStr.pop_back();

    std::string ipA, portNameA;
    int slot = -1;

    size_t first = pathStr.find('\\');
    if (first == std::string::npos)
    {
        ipA = pathStr;
    }
    else
    {
        ipA = pathStr.substr(0, first);
        std::string rest = pathStr.substr(first + 1);
        size_t second = rest.find('\\');
        if (second == std::string::npos)
        {
            portNameA = rest;
        }
        else
        {
            portNameA = rest.substr(0, second);
            slot = atoi(rest.substr(second + 1).c_str());
        }
    }

    std::wstring ip = Utf8ToWide(ipA.c_str());
    std::wstring portName = Utf8ToWide(portNameA.c_str());
    std::wstring cacheKey = MakeCacheKey(ip, portName, slot);

    Log(L"[QUERY] path='%hs' ip='%s' portName='%s' slot=%d",
        path, ip.c_str(), portName.c_str(), slot);

    // --- Cache-first lookup (no file I/O) ---
    auto it = g_queryCache.find(cacheKey);
    bool cacheHit = (it != g_queryCache.end() &&
                     it->second.found &&
                     it->second.classname != L"Unrecognized Device");

    if (!cacheHit)
    {
        // Determine what browse is needed
        bool needsDriverBrowse = (g_browsedDrivers.find(ip) == g_browsedDrivers.end());
        bool needsBackplane = !portName.empty() &&
                              (g_browsedBackplanes.find(ip) == g_browsedBackplanes.end());

        if (needsDriverBrowse)
        {
            Log(L"[QUERY] Running Phase 2+3 for %s", ip.c_str());
            ExecuteOnMainSTA(DoMainSTABrowse);
            WaitForEnumerators((int)g_enumerators.size());
            for (auto& drv : config.drivers) g_browsedDrivers.insert(drv.name);
        }

        if (needsBackplane)
        {
            Log(L"[QUERY] Running Phase 4+4b for %s", ip.c_str());
            g_capturedBuses.clear();
            g_captureBuses = true;
            HRESULT hrBus = ExecuteOnMainSTA(DoBusBrowse);
            if (SUCCEEDED(hrBus))
            {
                WaitForEnumerators((int)g_enumerators.size());
                g_captureBuses = false;
                HRESULT hrBP = ExecuteOnMainSTA(DoBackplaneBrowse);
                if (SUCCEEDED(hrBP))
                    WaitForEnumerators((int)g_enumerators.size());
            }
            g_captureBuses = false;
            g_browsedBackplanes.insert(ip);
        }

        if (needsDriverBrowse || needsBackplane)
        {
            // Refresh cache from topology  - ONE XML write, only when browse ran
            std::wstring xmlFile = LogPath(config.logDir, L"hook_topo_live.xml");
            if (SaveTopologyXML(pGlobals, xmlFile.c_str()))
            {
                UpdateDeviceIPsFromXML(xmlFile.c_str());
                PopulateQueryCache(xmlFile.c_str());
                if (!config.debugXml) DeleteFileW(xmlFile.c_str());
            }
            it = g_queryCache.find(cacheKey);
            cacheHit = (it != g_queryCache.end() &&
                        it->second.found &&
                        it->second.classname != L"Unrecognized Device");
        }
    }

    char resBuf[512];
    if (cacheHit)
    {
        char classA[128] = {}, nameA[256] = {}, ipABuf[64] = {};
        WideCharToMultiByte(CP_UTF8, 0, it->second.classname.c_str(), -1, classA, sizeof(classA), NULL, NULL);
        WideCharToMultiByte(CP_UTF8, 0, it->second.deviceName.c_str(), -1, nameA, sizeof(nameA), NULL, NULL);
        WideCharToMultiByte(CP_UTF8, 0, ip.c_str(), -1, ipABuf, sizeof(ipABuf), NULL, NULL);
        snprintf(resBuf, sizeof(resBuf), "R|FOUND|%s|%s|%s|%d", classA, nameA, ipABuf, slot);
        Log(L"[QUERY] %hs", resBuf);
    }
    else
    {
        snprintf(resBuf, sizeof(resBuf), "R|NOTFOUND|%s", path);
        Log(L"[QUERY] Not found: %hs", path);
    }
    PipeSendLine(resBuf);
    PipeSend("D|\n", 3);
}

// ============================================================
// SafeRunBrowsePhases
// Thin SEH wrapper — must have no local C++ objects with dtors.
// Returns true on success, false if an SEH exception was caught.
// ============================================================

static bool SafeRunBrowsePhases(HookConfig& config, IRSTopologyGlobals* pGlobals,
                                 std::vector<BusInfo>& buses)
{
    __try
    {
        RunBrowsePhases(config, pGlobals, buses);
        return true;
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        Log(L"[PIPE] RunBrowsePhases SEH exception: 0x%08x", GetExceptionCode());
        return false;
    }
}

// ============================================================
// HandleSession
// Manage one connected CLI client: read config, acquire buses,
// run browse if needed, enter command loop.
// ============================================================

static void HandleSession(HookConfig& config, IRSTopologyGlobals* pGlobals,
                          std::vector<BusInfo>& buses)
{
    static bool logDirSet = false;

    // Read config from pipe
    HookConfig newConfig;
    if (!ReadConfigFromPipe(newConfig))
    {
        Log(L"[FAIL] Cannot read config from pipe");
        return;
    }
    Log(L"[PIPE] Config received via pipe");

    // First session: switch log to config logDir
    if (!logDirSet && newConfig.logDir != L"C:\\temp")
    {
        std::wstring newLogPath = LogPath(newConfig.logDir, L"hook_log.txt");
        Log(L"[INFO] Switching log to: %s", newLogPath.c_str());
        FILE* newLog = _wfopen(newLogPath.c_str(), L"w, ccs=UTF-8");
        if (newLog)
        {
            fclose(g_logFile);
            g_logFile = newLog;
            logDirSet = true;
            Log(L"=== RSLinxHook v7 Worker Thread (logdir: %s) ===", newConfig.logDir.c_str());
        }
    }

    // Merge new drivers/IPs into persistent config
    bool hasNewWork = false;
    config.mode = newConfig.mode;
    config.debugXml = newConfig.debugXml;
    config.probeDispids = newConfig.probeDispids;
    if (!newConfig.logDir.empty()) config.logDir = newConfig.logDir;

    for (auto& newDrv : newConfig.drivers)
    {
        bool found = false;
        for (auto& existDrv : config.drivers)
        {
            if (_wcsicmp(existDrv.name.c_str(), newDrv.name.c_str()) == 0)
            {
                for (auto& ip : newDrv.ipAddresses)
                {
                    bool ipFound = false;
                    for (auto& existIp : existDrv.ipAddresses)
                        if (existIp == ip) { ipFound = true; break; }
                    if (!ipFound) { existDrv.ipAddresses.push_back(ip); hasNewWork = true; }
                }
                if (newDrv.newDriver) existDrv.newDriver = true;
                found = true;
                break;
            }
        }
        if (!found) { config.drivers.push_back(newDrv); hasNewWork = true; }
    }

    g_pSharedConfig = &config;

    Log(L"Drivers: %d, Mode: %s, LogDir: %s, NewWork: %s",
        (int)config.drivers.size(),
        config.mode == HookMode::Monitor ? L"monitor" : L"inject",
        config.logDir.c_str(),
        hasNewWork ? L"yes" : L"no");

    // Acquire buses for any new drivers
    if (hasNewWork || buses.empty())
    {
        AcquireNewBuses(config, pGlobals, buses);
    }

    // Run browse phases if this is the first session or there's new work
    bool shouldBrowse = buses.empty() ? false : (hasNewWork || g_browsedDrivers.empty());
    if (shouldBrowse)
    {
        // Always unadvise stale sinks before registering new ones.
        // Handles the case where the hook is reused across viewer sessions
        // without being unloaded — old CPs would otherwise accumulate and
        // fire every event twice (or more).
        ExecuteOnMainSTA(DoCleanupOnMainSTA);
        RunBrowsePhases(config, pGlobals, buses);
    }
    else
    {
        Log(L"[INFO] Already browsed  - replaying cached topology to new client");
        std::wstring pollFile = LogPath(config.logDir, L"hook_topo_poll.xml");
        DWORD attr = GetFileAttributesW(pollFile.c_str());
        if (attr != INVALID_FILE_ATTRIBUTES && !(attr & FILE_ATTRIBUTE_DIRECTORY))
        {
            TopologyCounts cached = CountDevicesInXML(pollFile.c_str());
            WalkTopologyTree(pGlobals);
            if (config.debugXml)
                PipeSendTopology(pollFile.c_str());
            PipeSendStatus(cached.totalDevices, cached.identifiedDevices, (int)g_discoveredDevices.size());
            Log(L"[INFO] Replayed cached topology: %d devices, %d identified",
                cached.totalDevices, cached.identifiedDevices);
        }
        else
        {
            Log(L"[WARN] No cached topology at %s  - client will get empty result", pollFile.c_str());
        }
    }

    // Signal end of initial browse to CLI
    PipeSend("D|\n", 3);

    // Command loop: handle Q| queries, B| re-browse, and STOP
    char line[512];
    while (g_pipeConnected && !g_shouldStop)
    {
        if (!PipeReadLine(line, sizeof(line))) break;
        if (strcmp(line, "STOP") == 0) break;
        if (line[0] == 'Q' && line[1] == '|')
        {
            HandleQuery(line + 2, pGlobals, config, buses);
        }
        else if (strcmp(line, "B|") == 0)
        {
            Log(L"[PIPE] Re-browse requested by client");
            ExecuteOnMainSTA(DoCleanupOnMainSTA); // unadvise stale sinks before re-registering
            bool ok = SafeRunBrowsePhases(config, pGlobals, buses);
            PipeSend("D|\n", 3);
            if (!ok) break; // COM state likely corrupt — disconnect client
        }
    }
}

// ============================================================
// Worker Thread
// ============================================================

static DWORD WINAPI WorkerThread(LPVOID lpParam)
{
    InitializeCriticalSection(&g_logCS);
    g_logFile = _wfopen(L"C:\\temp\\hook_log.txt", L"w, ccs=UTF-8");
    Log(L"=== RSLinxHook v7 Worker Thread Started ===");
    Log(L"PID: %d, TID: %d", GetCurrentProcessId(), GetCurrentThreadId());

    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    Log(L"[INFO] CoInitEx(STA): hr=0x%08x", hr);
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE && hr != S_FALSE)
    {
        Log(L"[FAIL] CoInitializeEx failed: 0x%08x", hr);
        if (g_logFile) fclose(g_logFile);
        DeleteCriticalSection(&g_logCS);
        return 1;
    }

    IHarmonyConnector* pHarmony = nullptr;
    IRSTopologyGlobals* pGlobals = nullptr;

    hr = CoCreateInstance(CLSID_HarmonyServices, NULL, CLSCTX_ALL,
                          IID_IHarmonyConnector, (void**)&pHarmony);
    if (FAILED(hr)) { Log(L"[FAIL] HarmonyServices: 0x%08x", hr); goto cleanup; }
    pHarmony->SetServerOptions(0, L"");
    Log(L"[OK] HarmonyServices");

    {
        IUnknown* pUnk = nullptr;
        hr = pHarmony->GetSpecialObject(&CLSID_RSTopologyGlobals, &IID_IRSTopologyGlobals, &pUnk);
        if (FAILED(hr)) { Log(L"[FAIL] GetSpecialObject: 0x%08x", hr); goto cleanup; }
        pUnk->QueryInterface(IID_IRSTopologyGlobals, (void**)&pGlobals);
        pUnk->Release();
        Log(L"[OK] TopologyGlobals");
    }

    // Register type libraries
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
            if (pTL) pTL->Release();
        }
    }

    // Create named pipe server and enter accept loop
    if (!PipeCreateServer())
    {
        Log(L"[FAIL] PipeCreateServer: %d", GetLastError());
        goto cleanup;
    }
    Log(L"[PIPE] Server created, waiting for clients...");

    {
        HookConfig config;
        std::vector<BusInfo> buses;

        while (!g_shouldStop)
        {
            Log(L"[PIPE] Waiting for client connection...");
            if (!PipeAcceptClient())
            {
                if (!g_shouldStop)
                    Log(L"[PIPE] AcceptClient failed: %d", GetLastError());
                break;
            }
            Log(L"[PIPE] Client connected");

            HandleSession(config, pGlobals, buses);

            PipeDisconnectClient();
            Log(L"[PIPE] Client disconnected, looping");
        }

        // Release buses accumulated across sessions
        for (auto& bus : buses)
        {
            if (bus.pBusDisp) bus.pBusDisp->Release();
            if (bus.pBusUnk) bus.pBusUnk->Release();
        }
    }

    // Cleanup: stop enumerators, unadvise connection points
    {
        Log(L"");
        Log(L"=== Cleanup: stopping enumerators ===");
        Log(L"Tracked: %d connection points, %d enumerators",
            (int)g_connectionPoints.size(), (int)g_enumerators.size());
        HRESULT hrClean = ExecuteOnMainSTA(DoCleanupOnMainSTA);
        Log(L"Cleanup result: 0x%08x", hrClean);
    }

cleanup:
    if (pGlobals) pGlobals->Release();
    if (pHarmony) pHarmony->Release();

    CoUninitialize();

    g_pSharedConfig = nullptr;
    Log(L"=== RSLinxHook v7 Worker Thread Done ===");

    PipeDestroyServer();

    if (g_logFile) fclose(g_logFile);
    g_logFile = nullptr;
    DeleteCriticalSection(&g_logCS);

    return 0;
}

// ============================================================
// DLL Entry Point
// ============================================================

BOOL WINAPI DllMain(HINSTANCE hDLL, DWORD dwReason, LPVOID lpReserved)
{
    if (dwReason == DLL_PROCESS_ATTACH)
    {
        DisableThreadLibraryCalls(hDLL);
        g_hWorkerThread = CreateThread(NULL, 0, WorkerThread, (LPVOID)hDLL, 0, NULL);
    }
    else if (dwReason == DLL_PROCESS_DETACH && lpReserved == NULL)
    {
        // FreeLibrary-initiated unload: signal stop and close pipe handle to
        // unblock any blocking ConnectNamedPipe / ReadFile in the worker thread.
        g_shouldStop = true;
        if (g_hPipe != INVALID_HANDLE_VALUE)
        {
            CloseHandle(g_hPipe);
            g_hPipe = INVALID_HANDLE_VALUE;
        }
        if (g_hWorkerThread != NULL)
        {
            CloseHandle(g_hWorkerThread);
            g_hWorkerThread = NULL;
        }
    }
    return TRUE;
}
