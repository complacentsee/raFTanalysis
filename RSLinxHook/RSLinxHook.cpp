/**
 * RSLinxHook.dll - In-process DLL for RSLinx device discovery
 *
 * When injected into rslinx.exe via CreateRemoteThread+LoadLibrary,
 * this DLL runs inside RSLinx's process where ENGINE.DLL is loaded.
 *
 * Strategy v6: Main-STA Browse via Thread Hook
 *
 * Phase 1 (PROVEN): ConnectNewDevice (DISPID 54) adds devices to topology
 *   - Works from worker STA via COM marshaling
 *   - Smart: DISP_E_EXCEPTION (0x80020009) = device already exists, skip
 *
 * Phase 2 (NEW): Main-STA browse via SetWindowsHookEx
 *   - ENGINE.DLL uses WSAAsyncSelect to bind CIP I/O to main thread's msg pump
 *   - IOnlineEnumerator::Start(path) must run on main STA for callbacks to fire
 *   - SetWindowsHookEx(WH_GETMESSAGE) hooks main thread - no window needed
 *   - Fallback: window subclass via EnumThreadWindows + SetWindowLongPtrW
 *
 * Phase 3: Topology polling to track identification progress
 *
 * Config: Reads C:\temp\hook_config.txt
 *   Line 1: Driver name (e.g., "Test")
 *   Lines 2+: IP addresses to add
 *
 * Results: Writes C:\temp\hook_results.txt and C:\temp\hook_log.txt
 */

#include <windows.h>
#include <commctrl.h>
#include <oleauto.h>
#include <ocidl.h>
#include <tlhelp32.h>
#include <stdio.h>
#include <string>
#include <vector>
#include <fstream>
#include <sstream>
#include <map>
#include <set>

// ============================================================
// Logging
// ============================================================

static FILE* g_logFile = nullptr;
static CRITICAL_SECTION g_logCS;

// Named pipe to TUI viewer (\\.\pipe\RSLinxHook)
static HANDLE g_hPipe = INVALID_HANDLE_VALUE;
static bool g_pipeConnected = false;

// Safe unload: worker thread handle + stop flag
static HANDLE g_hWorkerThread = NULL;
static volatile bool g_shouldStop = false;

static void PipeSend(const char* data, int len)
{
    if (!g_pipeConnected) return;
    DWORD written;
    if (!WriteFile(g_hPipe, data, len, &written, NULL))
    {
        g_pipeConnected = false;
        CloseHandle(g_hPipe);
        g_hPipe = INVALID_HANDLE_VALUE;
    }
}

static void Log(const wchar_t* fmt, ...)
{
    EnterCriticalSection(&g_logCS);

    // Write to log file
    va_list args;
    va_start(args, fmt);
    if (g_logFile) {
        vfwprintf(g_logFile, fmt, args);
        fwprintf(g_logFile, L"\n");
        fflush(g_logFile);
    }
    va_end(args);

    // Stream to TUI pipe as L|<text>\n
    if (g_pipeConnected) {
        wchar_t buf[2048];
        va_start(args, fmt);
        int len = _vsnwprintf(buf, 2046, fmt, args);
        va_end(args);
        if (len > 0) {
            char utf8[4096];
            utf8[0] = 'L'; utf8[1] = '|';
            int n = WideCharToMultiByte(CP_UTF8, 0, buf, len, utf8 + 2, 4090, NULL, NULL);
            if (n > 0) {
                utf8[n + 2] = '\n';
                PipeSend(utf8, n + 3);
            }
        }
    }

    LeaveCriticalSection(&g_logCS);
}

static void PipeSendTopology(const wchar_t* xmlPath)
{
    if (!g_pipeConnected) return;
    FILE* f = _wfopen(xmlPath, L"r");
    if (!f) return;

    EnterCriticalSection(&g_logCS);
    PipeSend("X|BEGIN\n", 8);
    char line[4096];
    while (fgets(line, sizeof(line), f))
        PipeSend(line, (int)strlen(line));
    PipeSend("X|END\n", 6);
    LeaveCriticalSection(&g_logCS);

    fclose(f);
}

static void PipeSendStatus(int total, int identified, int events)
{
    if (!g_pipeConnected) return;
    char buf[128];
    int n = snprintf(buf, sizeof(buf), "S|%d|%d|%d\n", total, identified, events);
    EnterCriticalSection(&g_logCS);
    PipeSend(buf, n);
    LeaveCriticalSection(&g_logCS);
}

// ============================================================
// COM GUIDs
// ============================================================

static const GUID CLSID_HarmonyServices =
    { 0x043D7910, 0x38FB, 0x11D2, { 0x90, 0x3D, 0x00, 0xC0, 0x4F, 0xA3, 0x63, 0xC1 } };
static const GUID CLSID_RSTopologyGlobals =
    { 0x38593054, 0x38E4, 0x11D0, { 0xBE, 0x25, 0x00, 0xC0, 0x4F, 0xC2, 0xAA, 0x48 } };
static const GUID CLSID_RSProjectGlobal =
    { 0xC92DFEA6, 0x1D29, 0x11D0, { 0xAD, 0x3F, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };
static const GUID CLSID_OnlineBusExt =
    { 0x2EC6B980, 0xC629, 0x11D0, { 0xBD, 0xCF, 0x08, 0x00, 0x09, 0xDC, 0x75, 0xC8 } };
static const GUID CLSID_RSPath =
    { 0x6DBDFEB8, 0xF703, 0x11D0, { 0xAD, 0x73, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };
// Port class for EtherNet/IP driver (from topology XML: port class="{00000001-0000-11D0-BDB8-080009DC75C8}")
static const GUID CLSID_EthernetPort =
    { 0x00000001, 0x0000, 0x11D0, { 0xBD, 0xB8, 0x08, 0x00, 0x09, 0xDC, 0x75, 0xC8 } };
// Bus class for Ethernet network (from topology XML: bus class="{00010010-5DFF-11CF-B4B9-C46F03C10000}")
static const GUID CLSID_EthernetBus =
    { 0x00010010, 0x5DFF, 0x11CF, { 0xB4, 0xB9, 0xC4, 0x6F, 0x03, 0xC1, 0x00, 0x00 } };

static const GUID IID_IHarmonyConnector =
    { 0x19EECB80, 0x3868, 0x11D2, { 0x90, 0x3D, 0x00, 0xC0, 0x4F, 0xA3, 0x63, 0xC1 } };
static const GUID IID_IRSTopologyGlobals =
    { 0x640DAC76, 0x38E3, 0x11D0, { 0xBE, 0x25, 0x00, 0xC0, 0x4F, 0xC2, 0xAA, 0x48 } };
static const GUID IID_IRSProjectGlobal =
    { 0xB286B85E, 0x211C, 0x11D0, { 0xAD, 0x42, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };
static const GUID IID_IRSProject =
    { 0xD61BFDA0, 0xEEAB, 0x11CE, { 0xB4, 0xB5, 0xC4, 0x6F, 0x03, 0xC1, 0x00, 0x00 } };
static const GUID IID_ITopologyBus =
    { 0x25C81D16, 0xF7BA, 0x11D0, { 0xAD, 0x73, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };
static const GUID IID_ITopologyBusEvents =
    { 0xAFF2BF80, 0x8D86, 0x11D0, { 0xB7, 0x7F, 0xF8, 0x72, 0x05, 0xC1, 0x00, 0x00 } };
static const GUID IID_ITopologyPathComposer =
    { 0xEF3A0832, 0x940C, 0x11D1, { 0xAD, 0xAC, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };
static const GUID IID_ITopologyDevice_Dual =
    { 0xB2A20A5E, 0xF7B9, 0x11D0, { 0xAD, 0x73, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };
static const GUID IID_IOnlineEnumeratorTypeLib =
    { 0xFC357A88, 0x0A98, 0x11D1, { 0xAD, 0x78, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };
static const GUID IID_IOnlineEnumerator =
    { 0x91748520, 0xA51B, 0x11D0, { 0xBD, 0xC9, 0x08, 0x00, 0x09, 0xDC, 0x75, 0xC8 } };
static const GUID IID_IRSTopologyNetwork =
    { 0x46DAD8E4, 0x4048, 0x11D0, { 0xBE, 0x26, 0x00, 0xC0, 0x4F, 0xC2, 0xAA, 0x48 } };
static const GUID IID_IRSTopologyObject =
    { 0xDCEAD8E0, 0x2E7A, 0x11CF, { 0xB4, 0xB5, 0xC4, 0x6F, 0x03, 0xC1, 0x00, 0x00 } };
static const GUID IID_IRSTopologyDevice =
    { 0xDCEAD8E1, 0x2E7A, 0x11CF, { 0xB4, 0xB5, 0xC4, 0x6F, 0x03, 0xC1, 0x00, 0x00 } };
// {2D76DE6C-94A0-11D0-AD56-00C04FD915B9} - ITopologyCollection (dispatch)
static const GUID IID_ITopologyCollection =
    { 0x2D76DE6C, 0x94A0, 0x11D0, { 0xAD, 0x56, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };
// {6C39E001-8775-11D0-B77F-F87205C10000} - ITopologyChassis (dispatch for backplane bus)
static const GUID IID_ITopologyChassis =
    { 0x6C39E001, 0x8775, 0x11D0, { 0xB7, 0x7F, 0xF8, 0x72, 0x05, 0xC1, 0x00, 0x00 } };
// {BB55A38E-8502-11D0-AD54-00C04FD915B9} - ITopologyObject (dispatch) — correct dispatch with path()
static const GUID IID_ITopologyObject =
    { 0xBB55A38E, 0x8502, 0x11D0, { 0xAD, 0x54, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };
// {98E549B2-B27E-11D0-AD5E-00C04FD915B9} - IRSTopologyPort (vtable)
static const GUID IID_IRSTopologyPort =
    { 0x98E549B2, 0xB27E, 0x11D0, { 0xAD, 0x5E, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };
// {94CB2140-450F-11CF-B4B5-C46F03C10000} - IRSObject (base Rockwell object, has GetName at vtable[7])
// Works on bus and device objects. Port objects do NOT implement this (E_NOINTERFACE).
static const GUID IID_IRSObject =
    { 0x94CB2140, 0x450F, 0x11CF, { 0xB4, 0xB5, 0xC4, 0x6F, 0x03, 0xC1, 0x00, 0x00 } };
// ============================================================
// Interface declarations (minimal)
// ============================================================

#undef INTERFACE
#define INTERFACE IHarmonyConnector
DECLARE_INTERFACE_(IHarmonyConnector, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(SetServerOptions)(THIS_ DWORD options, LPCWSTR servername) PURE;
    STDMETHOD(GetSpecialObject)(THIS_ const GUID* clsID, const GUID* iid, IUnknown** ppObject) PURE;
};

#undef INTERFACE
#define INTERFACE IRSTopologyGlobals
DECLARE_INTERFACE_(IRSTopologyGlobals, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(GetClassFromPCCCReply)(THIS_ void* pReply, GUID* pClassID) PURE;
    STDMETHOD(GetClassFromASAKey)(THIS_ void* pASAKey, GUID* pClassID) PURE;
    STDMETHOD(FindSinglePath)(THIS_ void* a, void* b, void** c) PURE;
    STDMETHOD(FindPartialPath)(THIS_ void* a, void* b, void** c) PURE;
    STDMETHOD(AddToPath)(THIS_ void* a, LPCWSTR b, void* c, void** d) PURE;
    STDMETHOD(ReleasePATHPART)(THIS_ void* a) PURE;
    STDMETHOD(GetThisWorkstationObject)(THIS_ void* container, IUnknown** workstation) PURE;
};

#undef INTERFACE
#define INTERFACE IRSProjectGlobal
DECLARE_INTERFACE_(IRSProjectGlobal, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(OpenProject)(THIS_ LPCWSTR name, DWORD flags, HWND hWnd, void* pApp, const GUID* iid, void** pp) PURE;
};

#undef INTERFACE
#define INTERFACE IRSProject
DECLARE_INTERFACE_(IRSProject, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(BindToObject)(THIS_ LPCWSTR path, IUnknown** ppObject) PURE;
};

// ============================================================
// ITopologyBusEvents interface (must be declared before SimpleEventSink)
// ============================================================

#undef INTERFACE
#define INTERFACE ITopologyBusEvents
DECLARE_INTERFACE_(ITopologyBusEvents, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(OnPortConnect)(THIS_ IUnknown*, IUnknown*, VARIANT) PURE;
    STDMETHOD(OnPortDisconnect)(THIS_ IUnknown*, IUnknown*, VARIANT) PURE;
    STDMETHOD(OnPortChangeAddress)(THIS_ IUnknown*, IUnknown*, VARIANT, VARIANT) PURE;
    STDMETHOD(OnPortChangeState)(THIS_ IUnknown*, long) PURE;
    STDMETHOD(OnBrowseStarted)(THIS_ IUnknown*) PURE;
    STDMETHOD(OnBrowseCycled)(THIS_ IUnknown*) PURE;
    STDMETHOD(OnBrowseEnded)(THIS_ IUnknown*) PURE;
    STDMETHOD(OnBrowseAddressFound)(THIS_ IUnknown*, VARIANT) PURE;
    STDMETHOD(OnBrowseAddressNotFound)(THIS_ IUnknown*, VARIANT) PURE;
};

// ============================================================
// IRSTopologyOnlineNotify interface (browse event callback)
// ============================================================
// Vtable layout (from analysis):
//   [0] QI  [1] AddRef  [2] Release
//   [3] BrowseStarted(IUnknown* pBus)
//   [4] BrowseCycled(IUnknown* pBus)
//   [5] BrowseEnded(IUnknown* pBus)
//   [6] Found(IUnknown* pBus, VARIANT addr)
//   [7] NothingAtAddress(IUnknown* pBus, VARIANT addr)

static const GUID IID_IRSTopologyOnlineNotify =
    { 0xFA5D9CF0, 0xA259, 0x11D1, { 0xBE, 0x10, 0x08, 0x00, 0x09, 0xDC, 0x75, 0xC8 } };

#undef INTERFACE
#define INTERFACE IRSTopologyOnlineNotify
DECLARE_INTERFACE_(IRSTopologyOnlineNotify, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(BrowseStarted)(THIS_ IUnknown* pBus) PURE;
    STDMETHOD(BrowseCycled)(THIS_ IUnknown* pBus) PURE;
    STDMETHOD(BrowseEnded)(THIS_ IUnknown* pBus) PURE;
    STDMETHOD(Found)(THIS_ IUnknown* pBus, VARIANT addr) PURE;
    STDMETHOD(NothingAtAddress)(THIS_ IUnknown* pBus, VARIANT addr) PURE;
};

// ============================================================
// Event Sink implementation - uses compiler-generated vtable
// ============================================================

static std::vector<std::wstring> g_discoveredDevices;

// SEH-safe VARIANT to wstring conversion (standalone C function, no C++ objects)
static bool SafeVariantToString(VARIANT* pAddr, wchar_t* outBuf, int bufLen)
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

// ============================================================
// Dual-interface event sink with FTM + padding.
//
// Key design points:
// 1. Dual inheritance gives correct vtables for both CPs
// 2. Accept-all QI lets CP[1] and CP[2] connect (required for browse)
// 3. FTM (Free Threaded Marshaler) enables cross-apartment event
//    delivery without proxy/stubs
// 4. Padding at +8 accommodates Start()'s direct memory writes
//    (Start writes up to offset ~68 from object start)
//
// Object layout:
//   +0:  vtable_1 (IRSTopologyOnlineNotify*)  [4 bytes]
//   +4:  vtable_2 (ITopologyBusEvents*)       [4 bytes]
//   +8:  m_pad[128] - absorbs Start()'s direct writes
//   +136: m_refCount (safe from Start)
//   +140: m_pFTM (IMarshal* for FTM)
//   +144: m_magic
// ============================================================

// Captured backplane bus references from OnBrowseStarted events (Phase 4a→4b)
static std::vector<IUnknown*> g_capturedBuses;
static volatile bool g_captureBuses = false;

class DualEventSink : public IRSTopologyOnlineNotify, public ITopologyBusEvents
{
public:
    BYTE m_pad[2048];           // +8: absorbs Start()'s writes (observed up to ~+528)
    LONG m_refCount;            // +136: safe from Start
    IUnknown* m_pFTM;           // +140: Free Threaded Marshaler
    DWORD m_magic;              // +144
    std::wstring m_label;       // identity for log messages (e.g., "Test/Ethernet", "5069-L320ER/Backplane")

    // Cycle detection: track seen addresses to detect when browse repeats
    CRITICAL_SECTION m_cs;
    std::set<std::wstring> m_seenAddresses;
    volatile bool m_cycleComplete;  // true when repeat address, BrowseCycled, or BrowseEnded
    volatile bool m_browseEnded;    // true when BrowseEnded fires
    int m_addressCount;             // total Found() calls

    DualEventSink(const wchar_t* label = L"") : m_refCount(1), m_pFTM(nullptr),
        m_magic(0xDEADBEEF), m_label(label ? label : L""),
        m_cycleComplete(false), m_browseEnded(false), m_addressCount(0)
    {
        memset(m_pad, 0, sizeof(m_pad));
        // +8 must be non-zero for Start to succeed
        DWORD* pDW = (DWORD*)m_pad;
        pDW[0] = 1;  // fake initial value at +8

        InitializeCriticalSection(&m_cs);

        // Create Free Threaded Marshaler for cross-apartment calls
        HRESULT hr = CoCreateFreeThreadedMarshaler(
            static_cast<IRSTopologyOnlineNotify*>(this), &m_pFTM);
        if (FAILED(hr))
            m_pFTM = nullptr;
    }

    ~DualEventSink()
    {
        DeleteCriticalSection(&m_cs);
        if (m_pFTM) m_pFTM->Release();
    }

    void DumpCounters(const wchar_t* label)
    {
        DWORD* pDW = (DWORD*)m_pad;
        Log(L"[COUNTERS@%s] magic=0x%08x ref=%d pad+0=0x%08x pad+4=0x%08x pad+8=0x%08x pad+16=0x%08x",
            label, m_magic, m_refCount, pDW[0], pDW[1], pDW[2], pDW[4]);
    }

    void DumpDWords(const wchar_t* label, int count = 32)
    {
        Log(L"=== DWORD DUMP @ %s (from obj+8) ===", label);
        DWORD* pBase = (DWORD*)m_pad;
        for (int i = 0; i < count; i++)
        {
            DWORD val = pBase[i];
            if (val != 0)
                Log(L"  +%03d: 0x%08X (%u / %d)", i * 4 + 8, val, val, (int)val);
        }
    }

    // --- IUnknown ---

    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_INVALIDARG;
        wchar_t gs[64];
        StringFromGUID2(riid, gs, 64);

        if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IRSTopologyOnlineNotify))
        {
            Log(L"[SINK] QI for %s - accepting (notify vtable)", gs);
            *ppv = static_cast<IRSTopologyOnlineNotify*>(this);
            AddRef();
            return S_OK;
        }
        if (IsEqualIID(riid, IID_ITopologyBusEvents))
        {
            Log(L"[SINK] QI for %s - accepting (bus vtable)", gs);
            *ppv = static_cast<ITopologyBusEvents*>(this);
            AddRef();
            return S_OK;
        }
        // FTM: enables cross-apartment marshaling without proxy/stubs
        if (IsEqualIID(riid, IID_IMarshal) && m_pFTM)
        {
            Log(L"[SINK] QI for IMarshal - delegating to FTM");
            return m_pFTM->QueryInterface(riid, ppv);
        }

        // Accept ALL unknown IIDs - return the notify vtable.
        // CP[1] {F0B077A1-...} and CP[2] {E548068D-...} MUST connect.
        Log(L"[SINK] QI for %s - accepting (unknown, returning notify vtable)", gs);
        *ppv = static_cast<IRSTopologyOnlineNotify*>(this);
        AddRef();
        return S_OK;
    }

    STDMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&m_refCount); }
    STDMETHODIMP_(ULONG) Release() override
    {
        ULONG c = InterlockedDecrement(&m_refCount);
        return c;  // Don't self-delete
    }

    // --- IRSTopologyOnlineNotify methods ---

    STDMETHODIMP BrowseStarted(IUnknown* pBus) override
    {
        Log(L"[ENUM:%s] Browse started (pBus=0x%p capture=%d)",
            m_label.c_str(), pBus, (int)g_captureBuses);
        // Capture backplane buses during Phase 4a
        if (g_captureBuses && pBus && (uintptr_t)pBus > 0x10000)
        {
            IDispatch* pDisp = nullptr;
            if (SUCCEEDED(pBus->QueryInterface(IID_IDispatch, (void**)&pDisp)) && pDisp)
            {
                pDisp->Release();
                bool dup = false;
                for (size_t i = 0; i < g_capturedBuses.size(); i++)
                {
                    if (g_capturedBuses[i] == pBus) { dup = true; break; }
                }
                if (!dup)
                {
                    pBus->AddRef();
                    g_capturedBuses.push_back(pBus);
                    Log(L"[ENUM:%s] Captured object #%d (0x%p)",
                        m_label.c_str(), (int)g_capturedBuses.size(), pBus);
                }
            }
        }
        return S_OK;
    }
    STDMETHODIMP BrowseCycled(IUnknown* pBus) override
    {
        m_cycleComplete = true;
        Log(L"[ENUM:%s] BrowseCycled (explicit)", m_label.c_str());
        return S_OK;
    }
    STDMETHODIMP BrowseEnded(IUnknown* pBus) override
    {
        m_cycleComplete = true;
        m_browseEnded = true;
        Log(L"[ENUM:%s] BrowseEnded (%d addresses seen)", m_label.c_str(), m_addressCount);
        return S_OK;
    }
    STDMETHODIMP Found(IUnknown* pBus, VARIANT addr) override
    {
        wchar_t addrBuf[256] = L"<unknown>";
        SafeVariantToString(&addr, addrBuf, 256);
        if (addr.vt == VT_I2 || addr.vt == VT_I4)
            Log(L"[ENUM:%s] Slot %s found", m_label.c_str(), addrBuf);
        else
            Log(L"[ENUM:%s] Address %s found", m_label.c_str(), addrBuf);

        EnterCriticalSection(&m_cs);
        m_addressCount++;
        if (m_seenAddresses.count(addrBuf) > 0)
        {
            if (!m_cycleComplete)
            {
                m_cycleComplete = true;
                Log(L"[ENUM:%s] Cycle complete -- repeat address %s (after %d addresses)",
                    m_label.c_str(), addrBuf, m_addressCount);
            }
        }
        else
            m_seenAddresses.insert(addrBuf);
        LeaveCriticalSection(&m_cs);

        g_discoveredDevices.push_back(addrBuf);
        return S_OK;
    }
    STDMETHODIMP NothingAtAddress(IUnknown* pBus, VARIANT addr) override
    {
        wchar_t addrBuf[256] = L"<unknown>";
        SafeVariantToString(&addr, addrBuf, 256);
        Log(L"[ENUM:%s] Nothing at %s", m_label.c_str(), addrBuf);
        return S_OK;
    }

    // --- ITopologyBusEvents methods ---

    STDMETHODIMP OnPortConnect(IUnknown*, IUnknown*, VARIANT) override { return S_OK; }
    STDMETHODIMP OnPortDisconnect(IUnknown*, IUnknown*, VARIANT) override { return S_OK; }
    STDMETHODIMP OnPortChangeAddress(IUnknown*, IUnknown*, VARIANT, VARIANT) override { return S_OK; }
    STDMETHODIMP OnPortChangeState(IUnknown*, long) override { return S_OK; }
    STDMETHODIMP OnBrowseStarted(IUnknown* pBus) override
    {
        Log(L"[BUS:%s] Browse started", m_label.c_str());
        return S_OK;
    }
    STDMETHODIMP OnBrowseCycled(IUnknown*) override
    {
        m_cycleComplete = true;
        Log(L"[BUS:%s] OnBrowseCycled (explicit)", m_label.c_str());
        return S_OK;
    }
    STDMETHODIMP OnBrowseEnded(IUnknown*) override
    {
        m_cycleComplete = true;
        m_browseEnded = true;
        Log(L"[BUS:%s] BrowseEnded (%d addresses seen)", m_label.c_str(), m_addressCount);
        return S_OK;
    }
    STDMETHODIMP OnBrowseAddressFound(IUnknown*, VARIANT addr) override
    {
        wchar_t addrBuf[256] = L"<unknown>";
        SafeVariantToString(&addr, addrBuf, 256);
        if (addr.vt == VT_I2 || addr.vt == VT_I4)
            Log(L"[BUS:%s] Slot %s found", m_label.c_str(), addrBuf);
        else
            Log(L"[BUS:%s] Address %s found", m_label.c_str(), addrBuf);

        EnterCriticalSection(&m_cs);
        m_addressCount++;
        if (m_seenAddresses.count(addrBuf) > 0)
        {
            if (!m_cycleComplete)
            {
                m_cycleComplete = true;
                Log(L"[BUS:%s] Cycle complete -- repeat address %s (after %d addresses)",
                    m_label.c_str(), addrBuf, m_addressCount);
            }
        }
        else
            m_seenAddresses.insert(addrBuf);
        LeaveCriticalSection(&m_cs);

        g_discoveredDevices.push_back(addrBuf);
        return S_OK;
    }
    STDMETHODIMP OnBrowseAddressNotFound(IUnknown*, VARIANT) override { return S_OK; }
};

// ============================================================
// Config
// ============================================================

enum class HookMode { Inject, Monitor };

struct DriverEntry {
    std::wstring name;
    std::vector<std::wstring> ipAddresses;
    bool newDriver = false;
};

struct HookConfig
{
    std::vector<DriverEntry> drivers;
    HookMode mode = HookMode::Inject;
    std::wstring logDir = L"C:\\temp";
    bool debugXml = false;

    // Backward compat helpers
    const std::wstring& driverName() const { return drivers[0].name; }
    const std::vector<std::wstring>& ipAddresses() const { return drivers[0].ipAddresses; }
    bool newDriver() const {
        for (auto& d : drivers) if (d.newDriver) return true;
        return false;
    }
    // Aggregate all IPs across all drivers
    std::vector<std::wstring> allIPs() const {
        std::vector<std::wstring> all;
        for (auto& d : drivers)
            for (auto& ip : d.ipAddresses)
                all.push_back(ip);
        return all;
    }
};

struct BusInfo {
    std::wstring driverName;
    IDispatch* pBusDisp;
    IUnknown* pBusUnk;
};

static bool ReadConfig(HookConfig& config)
{
    std::wifstream file(L"C:\\temp\\hook_config.txt");
    if (!file.is_open()) return false;

    std::wstring line;
    // Peek at first line to detect format
    if (!std::getline(file, line)) return false;

    if (line.substr(0, 5) == L"MODE=" || line.substr(0, 7) == L"DRIVER=" ||
        line.substr(0, 7) == L"LOGDIR=" || line == L"DEBUGXML=1")
    {
        // New multi-driver format: global directives + DRIVER= sections
        // Process the first line we already read, then continue
        file.seekg(0);
        while (std::getline(file, line))
        {
            if (line.empty()) continue;
            if (line == L"MODE=inject")
                config.mode = HookMode::Inject;
            else if (line == L"MODE=monitor")
                config.mode = HookMode::Monitor;
            else if (line.length() >= 7 && line.substr(0, 7) == L"LOGDIR=")
                config.logDir = line.substr(7);
            else if (line == L"DEBUGXML=1")
                config.debugXml = true;
            else if (line.length() >= 7 && line.substr(0, 7) == L"DRIVER=")
                config.drivers.push_back({line.substr(7), {}, false});
            else if (line == L"NEWDRIVER=1" && !config.drivers.empty())
                config.drivers.back().newDriver = true;
            else if (line.length() >= 3 && line.substr(0, 3) == L"IP=" && !config.drivers.empty())
                config.drivers.back().ipAddresses.push_back(line.substr(3));
        }
    }
    else
    {
        // Old single-driver format: line 1 = driver name
        DriverEntry drv;
        drv.name = line;
        if (std::getline(file, line))
        {
            if (line == L"MODE=monitor")
                config.mode = HookMode::Monitor;
            else if (line == L"MODE=inject")
                config.mode = HookMode::Inject;
            else if (!line.empty())
                drv.ipAddresses.push_back(line);
        }
        while (std::getline(file, line))
        {
            if (line.length() >= 7 && line.substr(0, 7) == L"LOGDIR=")
                config.logDir = line.substr(7);
            else if (line == L"DEBUGXML=1")
                config.debugXml = true;
            else if (line == L"NEWDRIVER=1")
                drv.newDriver = true;
            else if (!line.empty())
                drv.ipAddresses.push_back(line);
        }
        config.drivers.push_back(drv);
    }

    return !config.drivers.empty();
}

// Convert UTF-8 string to wide string
static std::wstring Utf8ToWide(const char* utf8)
{
    if (!utf8 || !*utf8) return L"";
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
    if (wlen <= 0) return L"";
    std::wstring result(wlen - 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, &result[0], wlen);
    return result;
}

// Read config from bidirectional pipe (C|KEY=VALUE lines, terminated by C|END)
static bool ReadConfigFromPipe(HookConfig& config)
{
    if (!g_pipeConnected) return false;
    char buf[4096];
    std::string accumulated;
    while (true) {
        DWORD bytesRead = 0;
        if (!ReadFile(g_hPipe, buf, sizeof(buf) - 1, &bytesRead, NULL) || bytesRead == 0)
            return false;
        buf[bytesRead] = '\0';
        accumulated += buf;
        // Process complete lines
        size_t pos;
        while ((pos = accumulated.find('\n')) != std::string::npos) {
            std::string line = accumulated.substr(0, pos);
            accumulated.erase(0, pos + 1);
            if (!line.empty() && line.back() == '\r') line.pop_back();
            if (line == "C|END") return !config.drivers.empty();
            if (line.length() >= 2 && line[0] == 'C' && line[1] == '|') {
                std::wstring wval = Utf8ToWide(line.c_str() + 2);
                if (wval == L"MODE=inject") config.mode = HookMode::Inject;
                else if (wval == L"MODE=monitor") config.mode = HookMode::Monitor;
                else if (wval.length() >= 7 && wval.substr(0, 7) == L"LOGDIR=") config.logDir = wval.substr(7);
                else if (wval == L"DEBUGXML=1") config.debugXml = true;
                else if (wval.length() >= 7 && wval.substr(0, 7) == L"DRIVER=") config.drivers.push_back({wval.substr(7), {}, false});
                else if (wval == L"NEWDRIVER=1" && !config.drivers.empty()) config.drivers.back().newDriver = true;
                else if (wval.length() >= 3 && wval.substr(0, 3) == L"IP=" && !config.drivers.empty()) config.drivers.back().ipAddresses.push_back(wval.substr(3));
            }
        }
    }
}

// Check for STOP signal from pipe (non-blocking)
static bool PipeCheckStop()
{
    if (!g_pipeConnected) return false;
    DWORD bytesAvail = 0;
    if (PeekNamedPipe(g_hPipe, NULL, 0, NULL, &bytesAvail, NULL) && bytesAvail > 0) {
        char stopBuf[64];
        DWORD bytesRead = 0;
        DWORD toRead = bytesAvail < 63 ? bytesAvail : 63;
        if (ReadFile(g_hPipe, stopBuf, toRead, &bytesRead, NULL) && bytesRead > 0) {
            stopBuf[bytesRead] = '\0';
            if (strstr(stopBuf, "STOP")) {
                return true;
            }
        }
    }
    return false;
}

// Build full path from log directory + filename
static std::wstring LogPath(const std::wstring& logDir, const wchar_t* filename)
{
    std::wstring path = logDir;
    if (!path.empty() && path.back() != L'\\')
        path += L'\\';
    path += filename;
    return path;
}

// ============================================================
// SEH helpers (standalone C functions)
// ============================================================

// SEH-safe memory dump (no C++ objects allowed)
static bool SafeReadMemory(void* pAddr, BYTE* outBuf, int bytes)
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

static HRESULT TryStartAtSlot(void* pInterface, IUnknown* pPath, int slot)
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
static HRESULT TryVtableGetObject(void* pInterface, int slot, IUnknown** ppResult)
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
static HRESULT TryVtableGetLabel(IUnknown* pObj, int slot, std::wstring& outLabel)
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
static HRESULT TryVtableAddPort(void* pDevice, int slot, GUID* pClsid,
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

// ============================================================
// ENGINE.DLL hot-load helper (SEH-safe)
// Calls Engine_LoadDriverModule + Engine_RefreshWorkstationPorts
// to trigger RSLinx to load a new driver from registry without restart.
// ============================================================

// DriverID → DRV filename mapping (from Ghidra analysis of ENGINE.DLL static table)
static const char* GetDrvFileForDriverID(DWORD driverID)
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

// Find the driver type key name (e.g. "AB_ETH") and DRV filename for a given driver Name
static bool FindDriverTypeAndDrv(const char* driverName, char* outDriverType, size_t typeLen,
                                  char* outDrvFile, size_t drvLen)
{
    outDriverType[0] = '\0';
    outDrvFile[0] = '\0';

    // Enumerate all driver type keys under Drivers (e.g., AB_ETH, AB_DF1, TCP)
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

        // Open this driver type key (e.g., Drivers\AB_ETH)
        HKEY hTypeKey;
        if (RegOpenKeyExA(hDriversKey, typeName, 0, KEY_READ, &hTypeKey) != ERROR_SUCCESS)
            continue;

        // Enumerate instances (e.g., AB_ETH-1, AB_ETH-2)
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

static void TryEngineHotLoad(const wchar_t* driverNameW)
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

    // Engine_AddDevice(CDevice*) — adds a single device to the topology port list
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
    // CMiniportDriver vtable layout (from Ghidra analysis):
    //   [1] = Start()   — re-reads registry, creates all CDevices
    //   [2] = Stop()    — stops all devices, clears collection
    //   [7] = GetDeviceCount()
    //   [8] = GetDevice(int)
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
    // The CDriver only reads registry at initialization (Start).
    // When we add a new device to the registry at runtime, we must
    // cycle the driver: Stop() clears devices, Start() re-reads registry.
    {
        Log(L"[ENGINE] Cycling CDriver: Stop() + Start() to pick up new registry entries...");

        // Stop: clears device collection, stops all devices
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

        // Start: re-reads registry, creates CDevices, calls SetRunning(1)
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

    // Engine_RefreshDeviceList rebuilds the full port list, diffs, fires events, refreshes COM
    if (pRefreshDevices)
    {
        Log(L"[ENGINE] Calling Engine_RefreshDeviceList...");
        __try { pRefreshDevices(); }
        __except(EXCEPTION_EXECUTE_HANDLER) { Log(L"[ENGINE] RefreshDeviceList CRASHED"); }
        Log(L"[ENGINE] RefreshDeviceList done");
    }

    // Explicit COM workstation ports refresh
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

// ============================================================
// SaveAsXML helper
// ============================================================

static bool SaveTopologyXML(IRSTopologyGlobals* pGlobals, const wchar_t* filename)
{
    IDispatch* pDisp = nullptr;
    HRESULT hr = pGlobals->QueryInterface(IID_IDispatch, (void**)&pDisp);
    if (FAILED(hr)) return false;

    VARIANT args[3];
    VariantInit(&args[0]);
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(filename);
    VariantInit(&args[1]);
    args[1].vt = VT_I4;
    args[1].lVal = 100;
    VariantInit(&args[2]);
    args[2].vt = VT_BSTR;
    args[2].bstrVal = SysAllocString(L"");

    DISPPARAMS params = {};
    params.rgvarg = args;
    params.cArgs = 3;

    VARIANT result;
    VariantInit(&result);

    hr = pDisp->Invoke(1610743808, IID_NULL, LOCALE_USER_DEFAULT,
                        DISPATCH_METHOD, &params, &result, nullptr, nullptr);

    VariantClear(&args[0]);
    VariantClear(&args[1]);
    VariantClear(&args[2]);
    VariantClear(&result);
    pDisp->Release();

    return SUCCEEDED(hr);
}

// ============================================================
// CountDevicesInXML helper
// ============================================================

struct TopologyCounts {
    int totalDevices;
    int identifiedDevices;
};

static TopologyCounts CountDevicesInXML(const wchar_t* filename)
{
    TopologyCounts counts = { 0, 0 };
    FILE* f = _wfopen(filename, L"r");
    if (!f) return counts;
    char buf[65536];
    size_t n = fread(buf, 1, sizeof(buf) - 1, f);
    buf[n] = 0;
    fclose(f);
    char* p = buf;
    while ((p = strstr(p, "<device ")) != nullptr)
    {
        counts.totalDevices++;
        char* cn = strstr(p, "classname=\"");
        if (cn && cn < p + 300)
        {
            cn += 11;
            if (strncmp(cn, "Unrecognized Device", 19) != 0 &&
                strncmp(cn, "Workstation", 11) != 0)
                counts.identifiedDevices++;
        }
        p++;
    }
    return counts;
}

// Check if a specific IP address has been identified (non-Unrecognized) in topology XML
// XML format: <address type="String" value="IP"> ... <device classname="..."> ... </address>
static bool IsTargetIdentifiedInXML(const wchar_t* filename, const std::vector<std::wstring>& targetIPs)
{
    FILE* f = _wfopen(filename, L"r");
    if (!f) return false;
    char buf[65536];
    size_t n = fread(buf, 1, sizeof(buf) - 1, f);
    buf[n] = 0;
    fclose(f);

    for (auto& wip : targetIPs)
    {
        // Convert wide IP to narrow
        char ipNarrow[64];
        WideCharToMultiByte(CP_ACP, 0, wip.c_str(), -1, ipNarrow, 64, NULL, NULL);

        // Build search pattern: value="IP"
        char pattern[128];
        snprintf(pattern, sizeof(pattern), "value=\"%s\"", ipNarrow);

        // Find the address element containing this IP
        char* addrPos = strstr(buf, pattern);
        if (!addrPos) continue;

        // Look for a <device> with a non-Unrecognized classname after this address
        // (within a reasonable range — device is nested in the address block)
        char* p = addrPos;
        char* searchEnd = addrPos + 2000;  // look within 2KB
        if (searchEnd > buf + n) searchEnd = buf + n;

        while (p < searchEnd && (p = strstr(p, "<device ")) != nullptr && p < searchEnd)
        {
            char* cn = strstr(p, "classname=\"");
            if (cn && cn < p + 500)
            {
                cn += 11;
                if (strncmp(cn, "Unrecognized Device", 19) != 0 &&
                    strncmp(cn, "Workstation", 11) != 0 &&
                    strncmp(cn, "\"", 1) != 0)  // empty classname
                    return true;
            }
            p++;
        }
    }
    return false;
}

// ============================================================
// Device Info — collected from COM object DISPIDs during browse
// ============================================================
struct DeviceInfo {
    std::wstring ip;           // "10.39.31.200" (from topology XML address mapping)
    std::wstring productName;  // DISPID 1 = Name (e.g. "5069-L310ER LOGIX310ER")
    std::wstring objectId;     // DISPID 2 = topology objectid GUID
};
static std::map<std::wstring, DeviceInfo> g_deviceDetails;  // keyed by productName

// Update g_deviceDetails with IP addresses extracted from topology XML
// Looks for: <address type="String" value="IP"> ... <device name="ProductName" ...>
static void UpdateDeviceIPsFromXML(const wchar_t* filename)
{
    FILE* f = _wfopen(filename, L"r");
    if (!f) return;
    char buf[65536];
    size_t n = fread(buf, 1, sizeof(buf) - 1, f);
    buf[n] = 0;
    fclose(f);

    // Find each <address type="String" value="IP">
    char* p = buf;
    while ((p = strstr(p, "<address type=\"String\" value=\"")) != nullptr)
    {
        p += 30;  // skip past '<address type="String" value="' to start of IP
        char* ipEnd = strchr(p, '"');
        if (!ipEnd) break;
        std::string ipA(p, ipEnd);
        p = ipEnd;

        // Look for <device name="..." within this address block (next 2KB)
        char* searchEnd = p + 2000;
        if (searchEnd > buf + n) searchEnd = buf + n;
        char* dev = strstr(p, "<device ");
        if (dev && dev < searchEnd)
        {
            // Skip <device reference="..."> entries (no name attribute)
            if (strncmp(dev + 8, "reference", 9) == 0) continue;

            char* nameAttr = strstr(dev, "name=\"");
            if (nameAttr && nameAttr < dev + 200)
            {
                nameAttr += 6;
                char* nameEnd = strchr(nameAttr, '"');
                if (nameEnd)
                {
                    std::string nameA(nameAttr, nameEnd);
                    std::wstring nameW(nameA.begin(), nameA.end());
                    std::wstring ipW(ipA.begin(), ipA.end());

                    auto it = g_deviceDetails.find(nameW);
                    if (it != g_deviceDetails.end())
                        it->second.ip = ipW;
                    else
                    {
                        DeviceInfo info;
                        info.productName = nameW;
                        info.ip = ipW;
                        g_deviceDetails[nameW] = info;
                    }
                }
            }
        }
    }
}

// ============================================================
// Main-STA Thread Hook Mechanism
// ============================================================

// Function pointer for what to execute on the main STA thread
typedef HRESULT (*MainSTAFunc)();
static MainSTAFunc g_pMainSTAFunc = nullptr;

// Shared state between worker thread and main STA
static HHOOK g_hHook = NULL;
static DWORD g_mainThreadId = 0;
static volatile LONG g_browseRequested = 0;
static volatile LONG g_browseResult = (LONG)E_PENDING;
static HookConfig* g_pSharedConfig = nullptr;

// Persistent objects — must survive after DoMainSTABrowse returns
// (events arrive later via main thread's message pump)
static DualEventSink* g_pMainSink = nullptr;
static IUnknown* g_pMainEnumUnk = nullptr;

// Global tracking for cleanup — stores all CPs we Advise'd and all enumerators we Start'd
struct ConnectionPointInfo { IConnectionPoint* pCP; DWORD cookie; };
struct EnumeratorInfo { void* pEnumInterface; DualEventSink* pSink; };
static std::vector<ConnectionPointInfo> g_connectionPoints;
static std::vector<EnumeratorInfo> g_enumerators;

// Check if all enumerators added since 'baseline' index have completed their browse cycle
static bool EnumeratorsCycledSince(int baseline)
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
static void GetEnumeratorStatusSince(int baseline, int& completed, int& total)
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

// Driver name for engine hot-load on main STA thread
static std::wstring g_engineDriverName;

// Wrapper to call TryEngineHotLoad on the main STA thread
static HRESULT DoEngineHotLoadOnMainSTA()
{
    Log(L"[ENGINE-STA] Running TryEngineHotLoad on main STA thread (TID=%d)", GetCurrentThreadId());
    TryEngineHotLoad(g_engineDriverName.c_str());
    return S_OK;
}

#define HOOK_MAGIC_WPARAM 0xDEAD7F00
#define SUBCLASS_MSG (WM_USER + 0x7F00)

// SEH-safe COM Release (standalone C function — no C++ objects)
static void SafeRelease(IUnknown* pUnk, const wchar_t* label)
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

// ============================================================
// IDispatch helpers for topology object enumeration
// ============================================================

// Get int property via DISPID (no args)
static int DispatchGetInt(IDispatch* pDisp, DISPID dispid)
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
static std::wstring DispatchGetString(IDispatch* pDisp, DISPID dispid)
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
static IDispatch* DispatchGetCollection(IDispatch* pDisp, DISPID dispid)
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

// Probe collection for available DISPIDs (diagnostic)
static void ProbeCollectionDISPIDs(IDispatch* pCollection)
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

// Enumerate collection using _NewEnum → IEnumVARIANT
// Returns all items as a vector of IDispatch* (caller must Release each)
static std::vector<IDispatch*> EnumerateCollection(IDispatch* pCollection)
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

// Get path object from topology object (DISPID 4, flags=0)
static IUnknown* DispatchGetPath(IDispatch* pDisp)
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

// Find the main (oldest) thread of current process
static DWORD FindMainThreadId()
{
    DWORD pid = GetCurrentProcessId();
    DWORD mainTid = 0;
    ULONGLONG oldestTime = (ULONGLONG)(-1);

    HANDLE snap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0);
    if (snap == INVALID_HANDLE_VALUE) return 0;

    THREADENTRY32 te = {};
    te.dwSize = sizeof(te);
    if (Thread32First(snap, &te))
    {
        do
        {
            if (te.th32OwnerProcessID == pid)
            {
                HANDLE hThread = OpenThread(THREAD_QUERY_LIMITED_INFORMATION, FALSE, te.th32ThreadID);
                if (hThread)
                {
                    FILETIME create, exitFT, kernel, user;
                    if (GetThreadTimes(hThread, &create, &exitFT, &kernel, &user))
                    {
                        ULONGLONG t = ((ULONGLONG)create.dwHighDateTime << 32) | create.dwLowDateTime;
                        if (t < oldestTime)
                        {
                            oldestTime = t;
                            mainTid = te.th32ThreadID;
                        }
                    }
                    CloseHandle(hThread);
                }
            }
        } while (Thread32Next(snap, &te));
    }
    CloseHandle(snap);
    return mainTid;
}

// Forward declarations
static HRESULT DoMainSTABrowse();
static HRESULT DoBusBrowse();
static HRESULT DoBackplaneBrowse();

// Helper: Get bus IDispatch from fresh COM objects on current STA
static IDispatch* GetBusDispatch(const wchar_t* driverName)
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
// For each device on the bus:
//   1. GetBackplanePort (vtable[19]) → port
//   2. port.GetBus (vtable[10]) → backplane bus
//   3. QI backplane bus for IOnlineEnumeratorTypeLib → enumerator
//   4. Get bus path, connect events, call Start(busPath)
// This matches RSWho's CDevice::StartBrowse approach.
// ============================================================
static HRESULT DoBusBrowse()
{
    Log(L"[BUS] DoBusBrowse starting on TID=%d", GetCurrentThreadId());

    if (!g_pSharedConfig)
    {
        Log(L"[BUS] ERROR: no shared config");
        return E_INVALIDARG;
    }

    // Iterate all drivers — get fresh bus IDispatch per driver on this STA
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

    // Get devices collection: bus.Devices() [DISPID 50]
    IDispatch* pDevices = DispatchGetCollection(pBusDisp, 50);
    if (!pDevices)
    {
        Log(L"[BUS] FAIL: bus.Devices() returned null");
        pBusDisp->Release();
        continue;
    }

    int deviceCount = DispatchGetInt(pDevices, 1);  // DISPID 1 = Count
    Log(L"[BUS] Bus has %d devices", deviceCount);

    // Enumerate using _NewEnum
    std::vector<IDispatch*> devices = EnumerateCollection(pDevices);
    Log(L"[BUS] Enumerated %d devices", (int)devices.size());

    for (int i = 0; i < (int)devices.size(); i++)
    {
        IDispatch* pDevice = devices[i];
        if (!pDevice) continue;

        std::wstring devName = DispatchGetString(pDevice, 1);  // DISPID 1 = Name
        Log(L"[BUS] Device %d: \"%s\"", i, devName.c_str());

        // Collect device properties via IDispatch
        // DISPID 1 = product name (already read), DISPID 2 = topology objectid GUID
        {
            std::wstring objectId = DispatchGetString(pDevice, 2);
            Log(L"[BUS] Device %d: objectId='%s'", i, objectId.c_str());

            DeviceInfo info;
            info.productName = devName;
            info.objectId = objectId;
            if (!devName.empty())
                g_deviceDetails[devName] = info;
        }

        // Step 1: QI for IRSTopologyDevice vtable interface
        IUnknown* pDevVtable = nullptr;
        pDevice->QueryInterface(IID_IRSTopologyDevice, (void**)&pDevVtable);
        if (!pDevVtable)
        {
            Log(L"[BUS]   QI for IRSTopologyDevice FAILED, skipping");
            pDevice->Release();
            continue;
        }

        // Step 2: Call GetBackplanePort at vtable[19]
        // IRSTopologyDevice vtable layout (after IRSTopologyObject[3-13]):
        //   [14]=AddPort [15]=RemovePort [16]=GetPortCount [17]=GetPortList
        //   [18]=GetPort [19]=GetBackplanePort [20]=GetModuleWidth
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

        // Step 3: QI the DEVICE for IOnlineEnumeratorTypeLib
        // Devices support their own enumerator that browses their internals
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

        // Step 4: Get device path via ITopologyObject dispatch DISPID 4
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

        // Step 5: Connect event sink to device's enumerator CPs
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

        // Step 6: Call Start(devicePath) on device's own enumerator
        HRESULT hrStart = TryStartAtSlot(pDevEnum, pDevPath, 7);
        Log(L"[BUS]   Start(device path) via device enum: hr=0x%08x", hrStart);
        if (SUCCEEDED(hrStart))
        {
            startedCount++;
            Log(L"[BUS]   >> Backplane browse started for \"%s\"", devName.c_str());
        }

        // Track enumerator+sink for cleanup (events arrive later)
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
// Navigate topology: device → backplane port → enumerate children to find bus.
// For each backplane bus: QI for enumerator, get path, call Start.
// ============================================================
static HRESULT DoBackplaneBrowse()
{
    Log(L"[BP] DoBackplaneBrowse starting on TID=%d", GetCurrentThreadId());

    if (!g_pSharedConfig)
    {
        Log(L"[BP] ERROR: no shared config");
        return E_INVALIDARG;
    }

    // Iterate all drivers
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

    // Get devices from Ethernet bus
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

        // Collect device properties
        {
            std::wstring objectId = DispatchGetString(pDevice, 2);
            Log(L"[BP] Device %d: objectId='%s'", i, objectId.c_str());

            DeviceInfo info;
            info.productName = devName;
            info.objectId = objectId;
            if (!devName.empty())
                g_deviceDetails[devName] = info;
        }

        // Get IRSTopologyDevice → GetBackplanePort[19]
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

        // === Primary approach: navigate port → bus directly ===
        // IRSTopologyPort::GetBus returns the bus without needing the port name.
        // Ghidra analysis of RSWho.ocx confirmed port objects don't implement IRSObject;
        // bus objects DO — so we get the bus first, then its name for labeling.
        IUnknown* pBackplaneBus = nullptr;
        std::wstring busLabel;

        IUnknown* pPortVtable = nullptr;
        HRESULT hrPortQI = pBackplanePort->QueryInterface(IID_IRSTopologyPort, (void**)&pPortVtable);
        Log(L"[BP]   QI IRSTopologyPort: hr=0x%08x", hrPortQI);

        if (pPortVtable)
        {
            // GetBus at slot 10 (inherited from IRSTopologyObject)
            // Slot 21 was documented as GetBus but returns E_ABORT; slot 10 confirmed working.
            IUnknown* pBusRaw = nullptr;
            HRESULT hrGetBus = TryVtableGetObject(pPortVtable, 10, &pBusRaw);
            Log(L"[BP]   GetBus[10]: hr=0x%08x bus=0x%p", hrGetBus, pBusRaw);

            if (SUCCEEDED(hrGetBus) && pBusRaw)
            {
                pBackplaneBus = pBusRaw;  // caller will Release

                // Get bus name via IRSObject::GetName[7] for labeling
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

        // === Fallback: DISPID 38 with hardcoded names (if GetBus failed) ===
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

        // Build sink label: "DeviceName/BusName"
        std::wstring sinkLabel = devName;
        if (!busLabel.empty())
            sinkLabel += L"/" + busLabel;

        // QI backplane bus for enumerator
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

        // Get bus path via ITopologyBus or ITopologyObject dispatch DISPID 4
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

        // Connect event sink with descriptive label
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

        // Call Start(busPath) on backplane bus enumerator
        HRESULT hrStart = TryStartAtSlot(pBPEnum, pBPPath, 7);
        Log(L"[BP]   Start(bus path): hr=0x%08x", hrStart);
        if (SUCCEEDED(hrStart))
        {
            startedCount++;
            Log(L"[BP]   >> Backplane bus browse STARTED for \"%s\"", sinkLabel.c_str());
        }

        // Track enumerator+sink for cleanup
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
// Stops all enumerators, unadvises all CPs, releases everything.
// ============================================================
static HRESULT DoCleanupOnMainSTA()
{
    Log(L"[CLEANUP] DoCleanupOnMainSTA starting on TID=%d", GetCurrentThreadId());

    // 1. Stop all enumerators via vtable[8] (SEH-wrapped, no-arg call)
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

    // 6. Clear global persistent pointers (already released via g_enumerators above)
    g_pMainSink = nullptr;
    g_pMainEnumUnk = nullptr;

    Log(L"[CLEANUP] DoCleanupOnMainSTA complete");
    return S_OK;
}

// DoMainSTABrowse — runs on the MAIN STA thread
// Creates COM objects fresh, sets up event sink, calls Start(path)
static HRESULT DoMainSTABrowse()
{
    Log(L"[MAIN-STA] DoMainSTABrowse starting on TID=%d", GetCurrentThreadId());

    if (!g_pSharedConfig)
    {
        Log(L"[MAIN-STA] ERROR: no shared config");
        return E_INVALIDARG;
    }

    HRESULT hr;

    // 1. Create HarmonyServices
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

    // 2. TopologyGlobals
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

    // 3. Project -> Workstation
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

    // 4-8. Per-driver: get bus, event sink, enumerator, path, Start
    HRESULT hrStart = E_FAIL;
    int busesStarted = 0;

    for (auto& drv : g_pSharedConfig->drivers)
    {
    Log(L"[MAIN-STA] === Driver: %s ===", drv.name.c_str());

    // 4. Get bus object
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

    // 5. Create event sink + connect to bus CP
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

    // 6. Get enumerator — QI bus for IID_IOnlineEnumeratorTypeLib
    //    RSWHO.OCX decompilation (Ghidra) shows CDevice::StartBrowse QI's the bus
    //    object for IID {FC357A88-0A98-11D1-AD78-00C04FD915B9} at DAT_10039cd0.
    //    This returns the enumerator that's connected to the bus's CIP I/O.
    //    A standalone CoCreateInstance(OnlineBusExt) is NOT connected to the bus.
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

    // 7. Get bus.path(flags=0)
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

    // 8. Call Start(path) via vtable[7]
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

    // Track enumerator+sink for cleanup (events arrive later)
    g_enumerators.push_back({(void*)pEnumUnk, pSink});

    // Release bus objects for this driver (enumerator/sink stay alive for events)
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

// WH_GETMESSAGE hook callback — runs on MAIN STA thread
static LRESULT CALLBACK MainSTAHookProc(int code, WPARAM wp, LPARAM lp)
{
    if (code >= 0)
    {
        MSG* pMsg = (MSG*)lp;
        if (pMsg->message == WM_NULL && pMsg->wParam == HOOK_MAGIC_WPARAM)
        {
            if (InterlockedCompareExchange(&g_browseRequested, 0, 1) == 1)
            {
                Log(L"[HOOK] Processing request on TID=%d", GetCurrentThreadId());
                HRESULT hr = g_pMainSTAFunc ? g_pMainSTAFunc() : E_POINTER;
                InterlockedExchange(&g_browseResult, (LONG)hr);
                Log(L"[HOOK] Result: 0x%08x", hr);
            }
        }
    }
    return CallNextHookEx(g_hHook, code, wp, lp);
}

// Window subclass fallback WndProc
static WNDPROC g_origWndProc = nullptr;

static LRESULT CALLBACK HookWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    if (msg == SUBCLASS_MSG)
    {
        if (InterlockedCompareExchange(&g_browseRequested, 0, 1) == 1)
        {
            Log(L"[SUBCLASS] Processing request on TID=%d", GetCurrentThreadId());
            HRESULT hr = g_pMainSTAFunc ? g_pMainSTAFunc() : E_POINTER;
            InterlockedExchange(&g_browseResult, (LONG)hr);
            Log(L"[SUBCLASS] Result: 0x%08x", hr);
        }
        return 0;
    }
    return CallWindowProcW(g_origWndProc, hwnd, msg, wp, lp);
}

// EnumThreadWindows callback — find any window on the main thread
static BOOL CALLBACK FindThreadWindowProc(HWND hwnd, LPARAM lParam)
{
    HWND* pResult = (HWND*)lParam;
    *pResult = hwnd;
    return FALSE;  // Stop after first window
}

// Execute a function on the main STA thread.
// Tries SetWindowsHookEx first, falls back to window subclass.
static HRESULT ExecuteOnMainSTA(MainSTAFunc func = nullptr)
{
    Log(L"=== ExecuteOnMainSTA ===");

    // Set the function to execute (default to DoMainSTABrowse)
    if (func)
        g_pMainSTAFunc = func;
    else
        g_pMainSTAFunc = DoMainSTABrowse;

    // Find main thread
    g_mainThreadId = FindMainThreadId();
    Log(L"  Main thread TID: %d (our TID: %d)", g_mainThreadId, GetCurrentThreadId());

    if (g_mainThreadId == 0 || g_mainThreadId == GetCurrentThreadId())
    {
        Log(L"  FAIL: Could not find main thread (or we ARE the main thread)");
        return E_FAIL;
    }

    // --- Try 1: SetWindowsHookEx(WH_GETMESSAGE) ---
    Log(L"  Trying SetWindowsHookEx(WH_GETMESSAGE)...");
    InterlockedExchange(&g_browseRequested, 1);
    InterlockedExchange(&g_browseResult, (LONG)E_PENDING);

    g_hHook = SetWindowsHookExW(WH_GETMESSAGE, MainSTAHookProc,
                                 GetModuleHandle(NULL), g_mainThreadId);
    if (g_hHook)
    {
        Log(L"  Hook installed: 0x%p", g_hHook);

        // Post WM_NULL with magic wParam to trigger the hook
        BOOL posted = PostThreadMessageW(g_mainThreadId, WM_NULL, HOOK_MAGIC_WPARAM, 0);
        Log(L"  PostThreadMessage: %s", posted ? L"OK" : L"FAILED");

        if (posted)
        {
            // Wait for result (up to 30s)
            DWORD t0 = GetTickCount();
            while (!g_shouldStop && InterlockedCompareExchange(&g_browseResult, 0, 0) == (LONG)E_PENDING)
            {
                if (GetTickCount() - t0 > 30000)
                {
                    Log(L"  TIMEOUT waiting for hook result after 30s");
                    break;
                }
                Sleep(100);
            }
        }
        else
        {
            Log(L"  PostThreadMessage failed (err=%d), will try fallback", GetLastError());
            InterlockedExchange(&g_browseRequested, 1);
        }

        UnhookWindowsHookEx(g_hHook);
        g_hHook = NULL;

        HRESULT result = (HRESULT)InterlockedCompareExchange(&g_browseResult, 0, 0);
        if (result != (LONG)E_PENDING)
        {
            Log(L"  Hook approach result: 0x%08x", result);
            return result;
        }
    }
    else
    {
        Log(L"  SetWindowsHookEx failed: %d", GetLastError());
    }

    // --- Try 2: Window subclass fallback ---
    Log(L"  Trying window subclass fallback...");
    HWND hTargetWnd = NULL;
    EnumThreadWindows(g_mainThreadId, FindThreadWindowProc, (LPARAM)&hTargetWnd);

    if (!hTargetWnd)
    {
        Log(L"  FAIL: No windows found on main thread TID=%d", g_mainThreadId);
        return E_FAIL;
    }

    wchar_t cls[256] = {};
    GetClassNameW(hTargetWnd, cls, 256);
    Log(L"  Found window: HWND=0x%p class=\"%s\"", hTargetWnd, cls);

    // Subclass the window
    InterlockedExchange(&g_browseRequested, 1);
    InterlockedExchange(&g_browseResult, (LONG)E_PENDING);

    g_origWndProc = (WNDPROC)SetWindowLongPtrW(hTargetWnd, GWLP_WNDPROC, (LONG_PTR)HookWndProc);
    if (g_origWndProc)
    {
        Log(L"  Subclassed window, sending SUBCLASS_MSG...");
        // SendMessage is synchronous — blocks until HookWndProc processes it
        SendMessageW(hTargetWnd, SUBCLASS_MSG, 0, 0);

        // Restore original wndproc
        SetWindowLongPtrW(hTargetWnd, GWLP_WNDPROC, (LONG_PTR)g_origWndProc);
        g_origWndProc = nullptr;

        HRESULT result = (HRESULT)InterlockedCompareExchange(&g_browseResult, 0, 0);
        Log(L"  Subclass approach result: 0x%08x", result);
        return result;
    }
    else
    {
        Log(L"  SetWindowLongPtrW failed: %d", GetLastError());
        return E_FAIL;
    }
}

// ============================================================
// RunMonitorLoop — continuous browse mode
// Runs after Phase 2 (browse started). Polls topology, triggers
// bus/backplane browse as devices appear, writes periodic results.
// Exits on stop signal (pipe STOP or DLL unload) or timeout.
// ============================================================
static void RunMonitorLoop(const HookConfig& config, IRSTopologyGlobals* pGlobals, const std::vector<BusInfo>& buses)
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
        // Check stop signal (DLL unload flag)
        if (g_shouldStop)
        {
            Log(L"[MONITOR] DLL unload stop signal received");
            break;
        }
        // Check stop signal (pipe or file fallback)
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

        // Pump messages (needed for COM event delivery)
        MSG msg;
        while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
        { TranslateMessage(&msg); DispatchMessage(&msg); }

        Sleep(100);

        DWORD elapsed = GetTickCount() - startTick;

        // Every 10s: save topology snapshot + update results
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

                // Trigger bus browse once devices are identified
                if (!busBrowseDone && c.identifiedDevices > 0)
                {
                    Log(L"[MONITOR] Devices identified — triggering bus browse");
                    g_capturedBuses.clear();
                    g_captureBuses = true;
                    HRESULT hrBus = ExecuteOnMainSTA(DoBusBrowse);
                    Log(L"[MONITOR] Bus browse: hr=0x%08x", hrBus);
                    busBrowseDone = true;
                }

                // Trigger backplane browse once bus browse has captured buses
                if (busBrowseDone && !backplaneBrowseDone && !g_capturedBuses.empty())
                {
                    g_captureBuses = false;
                    Log(L"[MONITOR] Captured %d buses — triggering backplane browse", (int)g_capturedBuses.size());
                    HRESULT hrBP = ExecuteOnMainSTA(DoBackplaneBrowse);
                    Log(L"[MONITOR] Backplane browse: hr=0x%08x", hrBP);
                    backplaneBrowseDone = true;
                }

                // Update device IPs from topology XML
                UpdateDeviceIPsFromXML(snapFile.c_str());

                // Write periodic results file
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

    // Cleanup
    Log(L"[MONITOR] Stopping — cleaning up");
    Log(L"Tracked: %d connection points, %d enumerators",
        (int)g_connectionPoints.size(), (int)g_enumerators.size());
    HRESULT hrClean = ExecuteOnMainSTA(DoCleanupOnMainSTA);
    Log(L"[MONITOR] Cleanup result: 0x%08x", hrClean);

    // Write final results
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

    // Clean up temporary poll XML when not in debug-xml mode
    if (!config.debugXml)
    {
        std::wstring pollPath = LogPath(config.logDir, L"hook_topo_poll.xml");
        DeleteFileW(pollPath.c_str());
    }

    // Signal pipe that monitor mode is done (D| + close handled in WorkerThread cleanup)
    Log(L"[MONITOR] Monitor loop complete");
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

    // Try connecting to pipe (bidirectional — read config, write log/status/topology)
    g_hPipe = CreateFileW(L"\\\\.\\pipe\\RSLinxHook", GENERIC_READ | GENERIC_WRITE,
                          0, NULL, OPEN_EXISTING, 0, NULL);
    if (g_hPipe != INVALID_HANDLE_VALUE) {
        g_pipeConnected = true;
        Log(L"[PIPE] Connected to viewer/CLI pipe");
    } else {
        Log(L"[PIPE] No pipe (%d) — file-only mode", GetLastError());
    }

    // Read config: try pipe first, fall back to file
    HookConfig config;
    bool configFromPipe = false;
    if (g_pipeConnected) {
        configFromPipe = ReadConfigFromPipe(config);
        if (configFromPipe)
            Log(L"[PIPE] Config received via pipe");
        else
            Log(L"[PIPE] Pipe config failed, trying file fallback");
    }
    if (!configFromPipe) {
        if (!ReadConfig(config))
        {
            Log(L"[FAIL] Cannot read config from pipe or file");
            if (g_logFile) fclose(g_logFile);
            DeleteCriticalSection(&g_logCS);
            return 1;
        }
    }

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

        HRESULT hrBrowse = ExecuteOnMainSTA();
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
