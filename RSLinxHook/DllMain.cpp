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
// Worker Thread
// ============================================================

static DWORD WINAPI WorkerThread(LPVOID lpParam)
{
    InitializeCriticalSection(&g_logCS);
    g_logFile = _wfopen(L"C:\\temp\\hook_log.txt", L"w, ccs=UTF-8");
    Log(L"=== RSLinxHook v7 Worker Thread Started ===");
    Log(L"PID: %d, TID: %d", GetCurrentProcessId(), GetCurrentThreadId());

    // Try connecting to pipe (bidirectional — read config, write log/status/topology)
    g_hPipe = CreateFileW(L"\\\\.\\pipe\\RSLinxHook", GENERIC_READ | GENERIC_WRITE,
                          0, NULL, OPEN_EXISTING, 0, NULL);
    if (g_hPipe != INVALID_HANDLE_VALUE) {
        g_pipeConnected = true;
        Log(L"[PIPE] Connected to viewer/CLI pipe");
    } else {
        Log(L"[PIPE] No pipe (%d) — file-only mode", GetLastError());
    }

    // Read config from pipe
    HookConfig config;
    if (!g_pipeConnected || !ReadConfigFromPipe(config))
    {
        Log(L"[FAIL] Cannot read config from pipe");
        if (g_logFile) fclose(g_logFile);
        DeleteCriticalSection(&g_logCS);
        return 1;
    }
    Log(L"[PIPE] Config received via pipe");

    // If logDir differs from default, reopen log at correct location
    if (config.logDir != L"C:\\temp")
    {
        std::wstring newLogPath = LogPath(config.logDir, L"hook_log.txt");
        Log(L"[INFO] Switching log to: %s", newLogPath.c_str());
        FILE* newLog = _wfopen(newLogPath.c_str(), L"w, ccs=UTF-8");
        if (newLog)
        {
            fclose(g_logFile);
            g_logFile = newLog;
            Log(L"=== RSLinxHook v7 Worker Thread (logdir: %s) ===", config.logDir.c_str());
        }
    }

    g_pSharedConfig = &config;

    // Log pipe status (after config read, so log may have moved to correct dir)
    if (!g_pipeConnected) {
        Log(L"[PIPE] Running in file-only mode (no pipe)");
    }

    Log(L"Drivers: %d, Mode: %s, LogDir: %s", (int)config.drivers.size(),
        config.mode == HookMode::Monitor ? L"monitor" : L"inject", config.logDir.c_str());
    for (size_t di = 0; di < config.drivers.size(); di++)
    {
        Log(L"  Driver[%d]: %s, IPs: %d%s", (int)di, config.drivers[di].name.c_str(),
            (int)config.drivers[di].ipAddresses.size(),
            config.drivers[di].newDriver ? L", NewDriver: YES" : L"");
    }

    g_discoveredDevices.clear();

    // Worker STA for topology polling via IDispatch
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    Log(L"[INFO] CoInitEx(STA): hr=0x%08x", hr);
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE && hr != S_FALSE)
    {
        Log(L"[FAIL] CoInitializeEx failed: 0x%08x", hr);
        if (g_logFile) fclose(g_logFile);
        g_pSharedConfig = nullptr;
        DeleteCriticalSection(&g_logCS);
        return 1;
    }

    FILE* resultFile = nullptr;

    // Worker thread COM objects (for ConnectNewDevice + topology polling)
    IHarmonyConnector* pHarmony = nullptr;
    IRSTopologyGlobals* pGlobals = nullptr;

    // Per-driver bus objects
    std::vector<BusInfo> buses;

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

    // Get bus objects for each driver — multi-strategy with creation fallback
    {
        IRSProject* pProject = nullptr;
        {
            IRSProjectGlobal* pPG = nullptr;
            hr = CoCreateInstance(CLSID_RSProjectGlobal, NULL, CLSCTX_ALL,
                                  IID_IRSProjectGlobal, (void**)&pPG);
            if (FAILED(hr)) { Log(L"[FAIL] ProjectGlobal: 0x%08x", hr); goto cleanup; }
            hr = pPG->OpenProject(L"", 0, NULL, NULL, &IID_IRSProject, (void**)&pProject);
            pPG->Release();
            if (FAILED(hr)) { Log(L"[FAIL] OpenProject: 0x%08x", hr); goto cleanup; }
        }

        IUnknown* pWorkstation = nullptr;
        hr = pGlobals->GetThisWorkstationObject(pProject, &pWorkstation);
        if (FAILED(hr)) { pProject->Release(); Log(L"[FAIL] GetWorkstation: 0x%08x", hr); goto cleanup; }

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
                goto cleanup;
            }
            Log(L"[OK] Re-acquired workstation after hot-load(s)");

            if (pGlobals)
            {
                std::wstring verifyPath = LogPath(config.logDir, L"hook_topo_after_hotload.xml");
                SaveTopologyXML(pGlobals, verifyPath.c_str());
            }
        }

        // Acquire bus for each driver
        const int maxRetries = 12;
        const int retryDelayMs = 5000;

        for (auto& drv : config.drivers)
        {
            IDispatch* pBusDisp = nullptr;
            IUnknown* pBusUnk = nullptr;

            Log(L"[BUS] Acquiring bus for driver '%s'...", drv.name.c_str());
            bool addPortAttempted = false;

            for (int attempt = 0; attempt < maxRetries && !pBusDisp; attempt++)
            {
                // Engine hot-load fallback on attempt 1+ (if not already done)
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

                // Strategy 4: AddPort — create the port/bus on workstation (one-shot per driver)
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
                                Log(L"[BUS]   QI IDispatch: hr=0x%08x", hrQI);
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

                // All strategies failed this attempt — wait and retry
                if (!pBusDisp && attempt < maxRetries - 1)
                {
                    Log(L"[WAIT] Bus \"%s\" not found (attempt %d/%d), retrying in %ds...",
                        drv.name.c_str(), attempt + 1, maxRetries, retryDelayMs / 1000);
                    Sleep(retryDelayMs);
                }
            } // end retry loop

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
        } // end per-driver loop

        pWorkstation->Release();
        pProject->Release();

        if (buses.empty())
        {
            Log(L"[FAIL] No buses acquired — cannot browse");
            goto cleanup;
        }
        Log(L"[OK] Acquired %d of %d buses", (int)buses.size(), (int)config.drivers.size());
    }

    // Register type libraries (needed for IDispatch on bus)
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

    // =============================================================
    // Phase 1: ConnectNewDevice — smart add with existence check (per driver)
    // =============================================================
    {
        Log(L"");
        Log(L"=== Phase 1: ConnectNewDevice (smart, %d drivers) ===", (int)buses.size());

        static const wchar_t* UNRECOGNIZED_DEVICE_GUID = L"{00000004-5D68-11CF-B4B9-C46F03C10000}";
        int totalAdded = 0, totalExisting = 0, totalFailed = 0;

        for (auto& bus : buses)
        {
            // Find driver config for this bus
            const DriverEntry* pDrv = nullptr;
            for (auto& d : config.drivers)
                if (d.name == bus.driverName) { pDrv = &d; break; }
            if (!pDrv || pDrv->ipAddresses.empty())
            {
                Log(L"  [%s] No IPs to add — using Node Table", bus.driverName.c_str());
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

                hr = bus.pBusDisp->Invoke(54, IID_NULL, LOCALE_USER_DEFAULT,
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

    // Save and send initial topology snapshot
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
    // Phase 2: Main-STA browse via thread hook
    // ENGINE.DLL uses WSAAsyncSelect to bind CIP I/O to main thread.
    // Start(path) must run on the main STA for callbacks to fire.
    // =============================================================
    {
        Log(L"");
        Log(L"=== Phase 2: Main-STA browse via thread hook ===");

        HRESULT hrBrowse = ExecuteOnMainSTA(DoMainSTABrowse);
        Log(L"Phase 2 result: 0x%08x", hrBrowse);

        if (FAILED(hrBrowse))
            Log(L"[WARN] Main-STA browse failed — identification may not trigger");
    }

    // =============================================================
    // Mode branch: monitor mode enters continuous loop
    // =============================================================
    if (config.mode == HookMode::Monitor)
    {
        Log(L"=== Entering Monitor Mode ===");
        RunMonitorLoop(config, pGlobals, buses);

        for (auto& bus : buses)
        {
            if (bus.pBusDisp) bus.pBusDisp->Release();
            if (bus.pBusUnk) bus.pBusUnk->Release();
        }
        goto cleanup;
    }

    // =============================================================
    // Phase 3: Topology polling — wait for target identification
    // Early exit: break once target IP is identified (min 3s, max 30s)
    // Polls every 2s for faster response.
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

            // Check enumerator cycle completion every 100ms (after 3s min)
            if (elapsed >= 3000 && EnumeratorsCycledSince(enumBaseline))
            {
                Log(L"  >> All Phase 2 enumerators cycled at %dms — advancing", elapsed);
                break;
            }

            // Poll every 2s
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

                    // Check if all target IPs are identified (early exit after min 3s)
                    std::vector<std::wstring> allIPs = config.allIPs();
                    if (elapsed >= 3000 && !targetIdentified && !allIPs.empty() &&
                        IsTargetIdentifiedInXML(pollFile.c_str(), allIPs))
                    {
                        targetIdentified = true;
                        Log(L"  >> Target IPs identified at %ds — exiting Phase 3 early", elapsed / 1000);
                        break;
                    }
                }
            }

            // Pump messages on worker thread (for COM marshaling)
            MSG msg;
            while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
            { TranslateMessage(&msg); DispatchMessage(&msg); }

            Sleep(100);
        }

        if (!targetIdentified)
            Log(L"  Target not identified after 30s — proceeding anyway");
    }

    // =============================================================
    // Phase 4: Bus browse — backplane modules
    // After bus-level identification, browse into device backplanes
    // to discover modules (slots on the backplane bus).
    // =============================================================
    {
        std::wstring midPath = LogPath(config.logDir, config.debugXml ? L"hook_topo_mid.xml" : L"hook_topo_poll.xml");
        SaveTopologyXML(pGlobals, midPath.c_str());
        TopologyCounts pre = CountDevicesInXML(midPath.c_str());
        Log(L"Topology after bus browse: %d devices, %d identified",
            pre.totalDevices, pre.identifiedDevices);
        // Only try bus browse if we have identified devices
        if (pre.identifiedDevices > 0)
        {
            Log(L"");
            Log(L"=== Phase 4: Bus browse (backplanes) ===");

            // Enable bus capture — OnBrowseStarted will save backplane bus refs
            g_capturedBuses.clear();
            g_captureBuses = true;

            int phase4Baseline = (int)g_enumerators.size();
            HRESULT hrBus = ExecuteOnMainSTA(DoBusBrowse);
            Log(L"Phase 4 result: 0x%08x", hrBus);

            if (SUCCEEDED(hrBus))
            {
                // Phase 5: Event-driven poll — exit when all bus enumerators cycle
                Log(L"");
                Log(L"=== Phase 5: Bus browse polling (event-driven, max 30s) ===");

                DWORD busStart = GetTickCount();
                while (!g_shouldStop && GetTickCount() - busStart < 30000)
                {
                    DWORD elapsed = GetTickCount() - busStart;

                    // Primary exit: all Phase 4 enumerators have cycled (after 2s min)
                    if (elapsed >= 2000 && EnumeratorsCycledSince(phase4Baseline))
                    {
                        Log(L"  >> All bus enumerators cycled at %dms — advancing", elapsed);
                        break;
                    }

                    // XML status logging every 2s
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
                Log(L"[WARN] Bus browse failed — skipping Phase 5");
            }

            // Phase 4b: Browse backplane buses (now that they exist)
            g_captureBuses = false;
            Log(L"");
            Log(L"=== Phase 4b: Backplane bus browse ===");
            Log(L"Captured %d backplane buses from events", (int)g_capturedBuses.size());
            int phase4bBaseline = (int)g_enumerators.size();
            HRESULT hrBP = ExecuteOnMainSTA(DoBackplaneBrowse);
            Log(L"Phase 4b result: 0x%08x", hrBP);

            if (SUCCEEDED(hrBP))
            {
                // Phase 5b: Event-driven poll — exit when all backplane enumerators cycle
                Log(L"");
                Log(L"=== Phase 5b: Backplane module polling (event-driven, max 30s) ===");

                DWORD bpStart = GetTickCount();
                while (!g_shouldStop && GetTickCount() - bpStart < 30000)
                {
                    DWORD elapsed = GetTickCount() - bpStart;

                    // Primary exit: all Phase 4b enumerators have cycled (after 2s min)
                    if (elapsed >= 2000 && EnumeratorsCycledSince(phase4bBaseline))
                    {
                        Log(L"  >> All backplane enumerators cycled at %dms — advancing", elapsed);
                        break;
                    }

                    // XML status logging every 2s
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
        }
        else
        {
            Log(L"");
            Log(L"=== Phase 4: Skipped (no identified devices for bus browse) ===");
        }
    }

    // =============================================================
    // Phase 6: Cleanup — stop enumerators, unadvise CPs, release
    // =============================================================
    {
        Log(L"");
        Log(L"=== Phase 6: Cleanup ===");
        Log(L"Tracked: %d connection points, %d enumerators",
            (int)g_connectionPoints.size(), (int)g_enumerators.size());
        HRESULT hrClean = ExecuteOnMainSTA(DoCleanupOnMainSTA);
        Log(L"Cleanup result: 0x%08x", hrClean);
    }

    // =============================================================
    // Final results
    // =============================================================
    Log(L"");
    Log(L"=== Final Results ===");

    {
        std::wstring afterPath = LogPath(config.logDir, config.debugXml ? L"hook_topo_after.xml" : L"hook_topo_poll.xml");
        SaveTopologyXML(pGlobals, afterPath.c_str());
        PipeSendTopology(afterPath.c_str());
    }
    {
        std::wstring afterPath = LogPath(config.logDir, config.debugXml ? L"hook_topo_after.xml" : L"hook_topo_poll.xml");
        TopologyCounts fc = CountDevicesInXML(afterPath.c_str());
        PipeSendStatus(fc.totalDevices, fc.identifiedDevices, (int)g_discoveredDevices.size());
        UpdateDeviceIPsFromXML(afterPath.c_str());
        std::vector<std::wstring> allIPs = config.allIPs();
        bool targetFound = !allIPs.empty() && IsTargetIdentifiedInXML(afterPath.c_str(), allIPs);
        Log(L"Final topology: %d devices, %d identified",
            fc.totalDevices, fc.identifiedDevices);
        if (!allIPs.empty())
            Log(L"Target IPs identified: %s", targetFound ? L"YES" : L"NO");
        Log(L"Events received: %d", (int)g_discoveredDevices.size());

        // Write results file (external program watches for this)
        std::wstring resultsPath = LogPath(config.logDir, L"hook_results.txt");
        resultFile = _wfopen(resultsPath.c_str(), L"w, ccs=UTF-8");
        if (resultFile)
        {
            fwprintf(resultFile, L"DRIVERS: %d\n", (int)config.drivers.size());
            fwprintf(resultFile, L"DEVICES_IDENTIFIED: %d\n", fc.identifiedDevices);
            fwprintf(resultFile, L"DEVICES_TOTAL: %d\n", fc.totalDevices);
            fwprintf(resultFile, L"EVENTS: %d\n", (int)g_discoveredDevices.size());
            if (!allIPs.empty())
                fwprintf(resultFile, L"TARGET: %s\n", allIPs[0].c_str());
            fwprintf(resultFile, L"TARGET_STATUS: %s\n", targetFound ? L"IDENTIFIED" : L"NOT_FOUND");
            // Write DEVICE: lines with probed info
            for (const auto& kv : g_deviceDetails)
            {
                const DeviceInfo& d = kv.second;
                fwprintf(resultFile, L"DEVICE: %s | %s\n",
                    d.ip.empty() ? L"(no IP)" : d.ip.c_str(),
                    d.productName.c_str());
            }
            fclose(resultFile);
            resultFile = nullptr;
        }
        Log(L"[OK] Results file written");

        // Clean up temporary poll XML when not in debug-xml mode
        if (!config.debugXml)
            DeleteFileW(afterPath.c_str());
    }

    // Release all bus objects
    for (auto& bus : buses)
    {
        if (bus.pBusDisp) bus.pBusDisp->Release();
        if (bus.pBusUnk) bus.pBusUnk->Release();
    }

cleanup:
    if (pGlobals) pGlobals->Release();
    if (pHarmony) pHarmony->Release();
    if (resultFile) fclose(resultFile);

    CoUninitialize();

    g_pSharedConfig = nullptr;
    Log(L"=== RSLinxHook v7 Worker Thread Done ===");

    // Signal TUI pipe that we're done, then close
    PipeSend("D|\n", 3);
    if (g_hPipe != INVALID_HANDLE_VALUE) {
        CloseHandle(g_hPipe);
        g_hPipe = INVALID_HANDLE_VALUE;
    }
    g_pipeConnected = false;

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
        // FreeLibrary-initiated unload: just signal and close handle.
        // Do NOT WaitForSingleObject here — it deadlocks under the loader lock
        // when the worker thread does COM cleanup (CoUninitialize, Release).
        // The viewer/CLI waits for pipe disconnect before calling FreeLibrary,
        // so the worker thread should already be finished by this point.
        g_shouldStop = true;
        if (g_hWorkerThread != NULL)
        {
            CloseHandle(g_hWorkerThread);
            g_hWorkerThread = NULL;
        }
    }
    return TRUE;
}
