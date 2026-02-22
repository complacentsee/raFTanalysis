#include "EngineHotLoad.h"
#include "Logging.h"

// ============================================================
// EngineHotLoad globals
// ============================================================

std::wstring g_engineDriverName;

// ============================================================
// EngineHotLoad implementations
// ============================================================

// DriverID → DRV filename mapping (from Ghidra analysis of ENGINE.DLL static table)
const char* GetDrvFileForDriverID(DWORD driverID)
{
    struct DriverMap { DWORD id; const char* drv; };
    static const DriverMap map[] = {
        { 0x62,   "ABTCP.DRV"         },  // AB_ETH (Ethernet devices)
        { 0x16,   "ABTCP.DRV"         },  // TCP
        { 0x121,  "ABTCP.DRV"         },  // AB_ETHIP (EtherNet/IP)
        { 0x03,   "AB_DF1.drv"        },  // AB_DF1
        { 0x9313, "AB_DF1.drv"        },  // AB_DF1_AUTOCONFIG
        { 0x9311, "AB_DF1.drv"        },
        { 0x9312, "AB_DF1.drv"        },
        { 0x13,   "ABEMU5.DRV"        },  // EMU500
        { 0x0F,   "ABEMU5.DRV"        },
        { 0x0B,   "AB_PIC.DRV"        },
        { 0x00,   "AB_KT.DRV"         },
        { 0x11,   "AB_MASTR.DRV"      },  // AB_MASTR
        { 0x12,   "AB_SLAVE.DRV"      },  // AB_SLAVE
        { 0x110,  "AB_VBP.DRV"        },  // AB_VBP
        { 0x111,  "AB_VBP.DRV"        },
        { 0x303A, "SmartGuardUSB.drv"  },  // SAFEGUARD
        { 0x3448, "AB_KTVista.DRV"    },  // AB_KTVista
        { 0x3039, "ABDNET.DRV"        },  // DNET
        { 0xFE62, "AB_PCC.DRV"        },
    };
    for (int i = 0; i < _countof(map); i++)
        if (map[i].id == driverID)
            return map[i].drv;
    return nullptr;
}

// Find the driver type key name and DRV filename for a given driver Name
bool FindDriverTypeAndDrv(const char* driverName, char* outDriverType, size_t typeLen,
                           char* outDrvFile, size_t drvLen)
{
    outDriverType[0] = '\0';
    outDrvFile[0] = '\0';

    // Enumerate all driver type keys under Drivers
    HKEY hDriversKey;
    if (RegOpenKeyExA(HKEY_LOCAL_MACHINE,
        "SOFTWARE\\WOW6432Node\\Rockwell Software\\RSLinx\\Drivers",
        0, KEY_READ, &hDriversKey) != ERROR_SUCCESS)
        return false;

    char typeName[128];
    DWORD typeNameLen;
    bool found = false;

    for (DWORD typeIdx = 0; !found; typeIdx++)
    {
        typeNameLen = sizeof(typeName);
        if (RegEnumKeyExA(hDriversKey, typeIdx, typeName, &typeNameLen,
            NULL, NULL, NULL, NULL) != ERROR_SUCCESS) break;

        HKEY hTypeKey;
        if (RegOpenKeyExA(hDriversKey, typeName, 0, KEY_READ, &hTypeKey) != ERROR_SUCCESS)
            continue;

        // Enumerate instances
        char instName[128];
        DWORD instNameLen;
        for (DWORD instIdx = 0; !found; instIdx++)
        {
            instNameLen = sizeof(instName);
            if (RegEnumKeyExA(hTypeKey, instIdx, instName, &instNameLen,
                NULL, NULL, NULL, NULL) != ERROR_SUCCESS) break;

            HKEY hInstKey;
            if (RegOpenKeyExA(hTypeKey, instName, 0, KEY_READ, &hInstKey) != ERROR_SUCCESS)
                continue;

            char nameVal[256] = {};
            DWORD nameValLen = sizeof(nameVal);
            DWORD nameType;
            if (RegQueryValueExA(hInstKey, "Name", NULL, &nameType,
                (BYTE*)nameVal, &nameValLen) == ERROR_SUCCESS
                && nameType == REG_SZ && _stricmp(nameVal, driverName) == 0)
            {
                strcpy_s(outDriverType, typeLen, typeName);

                // Look up DriverID from Loadable Drivers key
                char loadablePath[256];
                sprintf_s(loadablePath, "SOFTWARE\\WOW6432Node\\Rockwell Software\\RSLinx\\Loadable Drivers\\%s", typeName);
                HKEY hLoadable;
                if (RegOpenKeyExA(HKEY_LOCAL_MACHINE, loadablePath, 0, KEY_READ, &hLoadable) == ERROR_SUCCESS)
                {
                    DWORD driverID = 0;
                    DWORD idSize = sizeof(driverID);
                    if (RegQueryValueExA(hLoadable, "DriverID", NULL, NULL,
                        (BYTE*)&driverID, &idSize) == ERROR_SUCCESS)
                    {
                        const char* drv = GetDrvFileForDriverID(driverID);
                        if (drv)
                            strcpy_s(outDrvFile, drvLen, drv);
                    }
                    RegCloseKey(hLoadable);
                }
                found = true;
            }
            RegCloseKey(hInstKey);
        }
        RegCloseKey(hTypeKey);
    }
    RegCloseKey(hDriversKey);
    return found;
}

void TryEngineHotLoad(const wchar_t* driverNameW)
{
    HMODULE hEngine = GetModuleHandleA("ENGINE.DLL");
    if (!hEngine) hEngine = GetModuleHandleA("engine.dll");
    if (!hEngine) hEngine = GetModuleHandleA("Engine.dll");
    if (!hEngine)
    {
        Log(L"[ENGINE] ENGINE.DLL not found in process");
        return;
    }
    Log(L"[ENGINE] ENGINE.DLL found at 0x%p", hEngine);

    // Resolve ENGINE.DLL exports (mangled C++ names)
    typedef void* (__cdecl *pfnFindCDriverByName)(const char*);
    pfnFindCDriverByName pFindDriverByName = (pfnFindCDriverByName)GetProcAddress(hEngine,
        "?Engine_FindCDriver@@YAPAVCDriver@@PBD@Z");

    typedef void* (__cdecl *pfnFindCDriverByID)(int);
    pfnFindCDriverByID pFindDriverByID = (pfnFindCDriverByID)GetProcAddress(hEngine,
        "?Engine_FindCDriver@@YAPAVCDriver@@H@Z");

    typedef void (__cdecl *pfnRefreshPorts)();
    pfnRefreshPorts pRefreshPorts = (pfnRefreshPorts)GetProcAddress(hEngine,
        "?Engine_RefreshWorkstationPorts@@YAXXZ");

    typedef void (__cdecl *pfnRefreshDevices)();
    pfnRefreshDevices pRefreshDevices = (pfnRefreshDevices)GetProcAddress(hEngine,
        "?Engine_RefreshDeviceList@@YAXXZ");

    typedef int (__cdecl *pfnGetDriverCount)();
    pfnGetDriverCount pGetDriverCount = (pfnGetDriverCount)GetProcAddress(hEngine,
        "?Engine_GetDriverCount@@YAHXZ");

    typedef long (__cdecl *pfnAddDevice)(void*);
    pfnAddDevice pAddDevice = (pfnAddDevice)GetProcAddress(hEngine,
        "?Engine_AddDevice@@YAJPAVCDevice@@@Z");

    Log(L"[ENGINE] FindByName=0x%p RefreshDevices=0x%p RefreshPorts=0x%p AddDevice=0x%p DriverCount=0x%p",
        pFindDriverByName, pRefreshDevices, pRefreshPorts, pAddDevice, pGetDriverCount);

    // Convert driver name to narrow
    char narrowName[256] = {};
    WideCharToMultiByte(CP_ACP, 0, driverNameW, -1, narrowName, sizeof(narrowName), NULL, NULL);

    // Find driver type and DRV filename from registry
    char driverType[128] = {};
    char drvFile[128] = {};
    FindDriverTypeAndDrv(narrowName, driverType, sizeof(driverType), drvFile, sizeof(drvFile));
    Log(L"[ENGINE] Driver \"%S\" -> type=\"%S\" drv=\"%S\"", narrowName, driverType, drvFile);

    // === PHASE 1: Find existing CDriver for this driver type ===
    void* pCDriver = nullptr;

    if (pFindDriverByName && driverType[0])
    {
        __try { pCDriver = pFindDriverByName(driverType); }
        __except(EXCEPTION_EXECUTE_HANDLER) { pCDriver = nullptr; }
        Log(L"[ENGINE] FindCDriver(\"%S\") = 0x%p", driverType, pCDriver);
    }

    if (!pCDriver && pFindDriverByID)
    {
        char loadablePath[256];
        sprintf_s(loadablePath, "SOFTWARE\\WOW6432Node\\Rockwell Software\\RSLinx\\Loadable Drivers\\%s", driverType);
        HKEY hLoadable;
        DWORD driverID = 0;
        if (RegOpenKeyExA(HKEY_LOCAL_MACHINE, loadablePath, 0, KEY_READ, &hLoadable) == ERROR_SUCCESS)
        {
            DWORD idSize = sizeof(driverID);
            RegQueryValueExA(hLoadable, "DriverID", NULL, NULL, (BYTE*)&driverID, &idSize);
            RegCloseKey(hLoadable);
        }
        if (driverID)
        {
            __try { pCDriver = pFindDriverByID((int)driverID); }
            __except(EXCEPTION_EXECUTE_HANDLER) { pCDriver = nullptr; }
            Log(L"[ENGINE] FindCDriver(DriverID=0x%X) = 0x%p", driverID, pCDriver);
        }
    }

    if (!pCDriver)
    {
        Log(L"[ENGINE] No existing CDriver found");
        return;
    }

    // === PHASE 2: Check device count and determine if Stop/Start needed ===
    typedef void  (__thiscall *pfnDriverMethod)(void* thisPtr);
    typedef int   (__thiscall *pfnGetDeviceCount)(void* thisPtr);
    typedef void* (__thiscall *pfnGetDevice)(void* thisPtr, int index);

    char* pBase = (char*)pCDriver;
    void** vtable = *(void***)pBase;

    int deviceCountBefore = -1;
    __try {
        pfnGetDeviceCount fnGetCount = (pfnGetDeviceCount)vtable[7];
        deviceCountBefore = fnGetCount(pCDriver);
    } __except(EXCEPTION_EXECUTE_HANDLER) { deviceCountBefore = -99; }
    Log(L"[ENGINE] CDriver GetDeviceCount() BEFORE = %d", deviceCountBefore);

    // === PHASE 3: Stop + Start cycle to force CDriver to re-read registry ===
    {
        Log(L"[ENGINE] Cycling CDriver: Stop() + Start() to pick up new registry entries...");

        pfnDriverMethod fnStop = (pfnDriverMethod)vtable[2];
        Log(L"[ENGINE] Calling CDriver::Stop() [vtable[2]=0x%p]...", vtable[2]);
        __try { fnStop(pCDriver); }
        __except(EXCEPTION_EXECUTE_HANDLER) { Log(L"[ENGINE] CDriver::Stop() CRASHED"); }

        int countAfterStop = -1;
        __try {
            pfnGetDeviceCount fnGetCount = (pfnGetDeviceCount)vtable[7];
            countAfterStop = fnGetCount(pCDriver);
        } __except(EXCEPTION_EXECUTE_HANDLER) { countAfterStop = -99; }
        Log(L"[ENGINE] CDriver GetDeviceCount() after Stop = %d", countAfterStop);

        pfnDriverMethod fnStart = (pfnDriverMethod)vtable[1];
        Log(L"[ENGINE] Calling CDriver::Start() [vtable[1]=0x%p]...", vtable[1]);
        __try { fnStart(pCDriver); }
        __except(EXCEPTION_EXECUTE_HANDLER) { Log(L"[ENGINE] CDriver::Start() CRASHED"); }

        int countAfterStart = -1;
        __try {
            pfnGetDeviceCount fnGetCount = (pfnGetDeviceCount)vtable[7];
            countAfterStart = fnGetCount(pCDriver);
        } __except(EXCEPTION_EXECUTE_HANDLER) { countAfterStart = -99; }
        Log(L"[ENGINE] CDriver GetDeviceCount() after Start = %d (was %d)", countAfterStart, deviceCountBefore);
    }

    // === PHASE 4: Register all devices with ENGINE topology and refresh ===
    {
        pfnGetDeviceCount fnGetCount = (pfnGetDeviceCount)vtable[7];
        pfnGetDevice fnGetDevice = (pfnGetDevice)vtable[8];
        int deviceCount = 0;
        __try { deviceCount = fnGetCount(pCDriver); }
        __except(EXCEPTION_EXECUTE_HANDLER) { deviceCount = 0; }

        if (pAddDevice && deviceCount > 0)
        {
            for (int i = 0; i < deviceCount && i < 10; i++)
            {
                void* pDevice = nullptr;
                __try { pDevice = fnGetDevice(pCDriver, i); }
                __except(EXCEPTION_EXECUTE_HANDLER) { pDevice = nullptr; }
                if (pDevice)
                {
                    long addResult = -1;
                    __try { addResult = pAddDevice(pDevice); }
                    __except(EXCEPTION_EXECUTE_HANDLER) { addResult = -99; }
                    Log(L"[ENGINE] Engine_AddDevice(Device[%d]=0x%p): hr=0x%08x", i, pDevice, addResult);
                }
            }
        }
    }

    if (pRefreshDevices)
    {
        Log(L"[ENGINE] Calling Engine_RefreshDeviceList...");
        __try { pRefreshDevices(); }
        __except(EXCEPTION_EXECUTE_HANDLER) { Log(L"[ENGINE] RefreshDeviceList CRASHED"); }
        Log(L"[ENGINE] RefreshDeviceList done");
    }

    if (pRefreshPorts)
    {
        Log(L"[ENGINE] Calling Engine_RefreshWorkstationPorts...");
        __try { pRefreshPorts(); }
        __except(EXCEPTION_EXECUTE_HANDLER) { Log(L"[ENGINE] RefreshWorkstationPorts CRASHED"); }
        Log(L"[ENGINE] RefreshWorkstationPorts done");
    }

    Log(L"[ENGINE] Hot-load complete, waiting for topology propagation...");
    Sleep(5000);
}

// Wrapper to call TryEngineHotLoad on the main STA thread
HRESULT DoEngineHotLoadOnMainSTA()
{
    Log(L"[ENGINE-STA] Running TryEngineHotLoad on main STA thread (TID=%d)", GetCurrentThreadId());
    TryEngineHotLoad(g_engineDriverName.c_str());
    return S_OK;
}
