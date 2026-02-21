# RSLinx COM Architecture — Device Discovery & Browse

Comprehensive documentation of the COM interfaces, DISPIDs, vtable layouts, and browse patterns used by RSLinxHook.dll to programmatically discover and identify devices in RSLinx Classic's topology.

---

## Table of Contents

1. [COM Object Chain](#1-com-object-chain)
2. [Driver Acquisition](#2-driver-acquisition)
3. [Device Addition (ConnectNewDevice)](#3-device-addition-connectnewdevice)
4. [Browse Triggering](#4-browse-triggering)
5. [Event Sink Pattern](#5-event-sink-pattern)
6. [Backplane Browsing](#6-backplane-browsing)
7. [Topology Polling](#7-topology-polling)
8. [Phase Architecture](#8-phase-architecture)
9. [Threading Model](#9-threading-model)
10. [COM GUID Reference](#10-com-guid-reference)
11. [DISPID Reference](#11-dispid-reference)
12. [Vtable Layouts](#12-vtable-layouts)
13. [What Works / What Doesn't](#13-what-works--what-doesnt)

---

## 1. COM Object Chain

RSLinx's topology is accessed through a chain of COM objects starting from HarmonyServices:

```
CoCreateInstance(CLSID_HarmonyServices)
    → IHarmonyConnector
    → SetServerOptions(0, "")
    → GetSpecialObject(CLSID_RSTopologyGlobals, IID_IRSTopologyGlobals)
        → IRSTopologyGlobals

CoCreateInstance(CLSID_RSProjectGlobal)
    → IRSProjectGlobal
    → OpenProject("") → IRSProject

IRSTopologyGlobals::GetThisWorkstationObject(pProject)
    → IUnknown (Workstation)
    → QueryInterface(IID_ITopologyDevice_Dual)
        → ITopologyDevice_Dual (IDispatch)

Workstation.Invoke(DISPID 38, "DriverName")
    → ITopologyBus (IDispatch)
```

Each COM object must be created on an STA thread. The workstation device is the root of the topology tree — all drivers and their buses are accessible through it.

---

## 2. Driver Acquisition

Drivers are accessed via `DISPID 38` on the workstation's `ITopologyDevice_Dual` dispatch interface. This is a universal bus-by-port-name accessor:

```cpp
VARIANT arg;
arg.vt = VT_BSTR;
arg.bstrVal = SysAllocString(L"DriverName");

DISPPARAMS dp = { &arg, nullptr, 1, 0 };
VARIANT result;
VariantInit(&result);

hr = pWorkstation->Invoke(38, IID_NULL, LOCALE_USER_DEFAULT,
                          DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);
// result.vt == VT_DISPATCH → ITopologyBus
```

**DISPID 38** works universally:
- On the **workstation**: `ws.Invoke(38, "Test")` → Ethernet bus for driver "Test"
- On a **device**: `dev.Invoke(38, "Backplane")` → backplane bus
- On a **CompactLogix device**: `dev.Invoke(38, "CompactBus")` → backplane bus

The port names match the `<port name="...">` attribute in topology XML.

---

## 3. Device Addition (ConnectNewDevice)

Devices are added to the topology via `DISPID 54` on the bus's IDispatch. This takes 6 VARIANT arguments in reverse order (COM convention):

```cpp
VARIANT args[6];
VariantInit(args); // for all 6

args[5].vt = VT_I4;    args[5].lVal = 0;                    // flags
args[4].vt = VT_BSTR;  args[4].bstrVal = SysAllocString(    // device class GUID
    L"{00000004-5D68-11CF-B4B9-C46F03C10000}");              // = Unrecognized Device
args[3].vt = VT_EMPTY;                                       // reserved
args[2].vt = VT_BSTR;  args[2].bstrVal = SysAllocString(L"Device");  // type label
args[1].vt = VT_BSTR;  args[1].bstrVal = SysAllocString(L"A");       // port name
args[0].vt = VT_BSTR;  args[0].bstrVal = SysAllocString(L"10.13.30.68"); // IP

DISPPARAMS dp = { args, nullptr, 6, 0 };
hr = pBus->Invoke(54, IID_NULL, LOCALE_USER_DEFAULT,
                   DISPATCH_METHOD, &dp, &result, &excepInfo, &argErr);
```

**Result handling:**
- `S_OK` → device newly added to topology
- `DISP_E_EXCEPTION (0x80020009)` → device already exists (skip silently)
- Other HRESULT → failure, log and continue

Devices must exist in the topology before browse can identify them. The "Unrecognized Device" class GUID is a placeholder — browse replaces it with the actual device identity.

---

## 4. Browse Triggering

### The Core Problem

`IOnlineEnumerator::Start(path)` triggers CIP identity requests over the network. But `ENGINE.DLL` uses `WSAAsyncSelect()` to bind CIP socket I/O to the **main thread's message pump**. If `Start()` is called from a different apartment, the enumerator's internal callbacks are bound to the wrong message pump and never fire.

### The Solution: Main-STA Thread Hook

```
Worker Thread (STA)                     Main Thread (STA)
       |                                      |
  SetWindowsHookEx(WH_GETMESSAGE) ---------> |
  PostThreadMessage(WM_NULL, MAGIC) -------> |
       |                                  HookProc fires
       |                                  DoMainSTABrowse()
       |                                    CoCreateInstance(HarmonyServices)
       |                                    → TopologyGlobals → Project → WS → Bus
       |                                    Bus QI → IOnlineEnumeratorTypeLib
       |                                    Connect DualEventSink to CPs
       |                                    Start(bus.path) → CIP browse begins
       |                                  return hr
  g_browseResult = hr <------------------ |
  UnhookWindowsHookEx                      |
       |                              Message pump processes CIP responses
       |                              Events fire: BrowseStarted, Found, Cycled
```

**Fallback**: If `SetWindowsHookEx` fails, uses window subclassing via `SetWindowLongPtrW(GWLP_WNDPROC)` on a hidden COM window found via `EnumThreadWindows`.

### Why Bus QI, Not CoCreateInstance?

**The critical discovery.** Ghidra decompilation of `RSWHO.OCX` (`CDevice::StartBrowse` at `0x100128d0`) shows that RSWho QI's the bus for `IID_IOnlineEnumeratorTypeLib`:

```
IID at DAT_10039cd0 = {FC357A88-0A98-11D1-AD78-00C04FD915B9}
```

The bus returns its **internal** enumerator — one connected to the bus's CIP I/O machinery via `ENGINE.DLL`. A standalone `CoCreateInstance(CLSID_OnlineBusExt)` creates a disconnected enumerator: `Start()` succeeds but triggers zero CIP traffic.

```cpp
// CORRECT: Bus QI returns connected enumerator
IUnknown* pBusEnum = nullptr;
pBusUnk->QueryInterface(IID_IOnlineEnumeratorTypeLib, (void**)&pBusEnum);

// WRONG: Standalone creates disconnected enumerator
CoCreateInstance(CLSID_OnlineBusExt, nullptr, CLSCTX_INPROC_SERVER,
                 IID_IOnlineEnumerator, (void**)&pEnum);
```

The bus-QI'd enumerator exposes **6 connection points** (vs 3 for standalone):

| # | IID | Purpose |
|---|-----|---------|
| 1 | `IID_IRSTopologyOnlineNotify` | Browse events (Started/Found/Ended) |
| 2 | `{F0B077A1-7483-4BCE-A33C-5E27B6A5FEA1}` | Undocumented |
| 3 | `{E548068D-20F5-48EC-8519-12BA9AC0002D}` | Undocumented |
| 4 | `IID_ITopologyBusEvents` | Bus events (OnBrowseAddressFound) |
| 5 | `{DC9A007C-BA7F-11D0-AD5F-00C04FD915B9}` | Undocumented |
| 6 | `{595E68A0-61FB-11D1-8922-444553540000}` | Undocumented |

### Getting the Path Object

The enumerator's `Start()` method requires a path object:

```cpp
// bus.path(flags=0) via DISPID 4 with VT_I4 argument
VARIANT flagArg;
flagArg.vt = VT_I4;
flagArg.lVal = 0;

DISPPARAMS dp = { &flagArg, nullptr, 1, 0 };
VARIANT pathResult;
hr = pBusDisp->Invoke(4, IID_NULL, LOCALE_USER_DEFAULT,
                       DISPATCH_PROPERTYGET, &dp, &pathResult, nullptr, nullptr);
// pathResult.vt == VT_DISPATCH → path object
```

**Important:** `DISPID 4` with no arguments routes to `IADs.GUID` and returns `TYPE_E_BADMODULEKIND`. The `VT_I4 flags=0` argument is required to reach the topology bus's `path` property.

### Starting the Browse

```cpp
// Start via vtable[7] on the enumerator
typedef HRESULT (__stdcall *StartFn)(void* pThis, IUnknown* pPath);
void** vtable = *(void***)pBusEnum;
StartFn pfnStart = (StartFn)vtable[7];
hr = pfnStart(pBusEnum, pPathUnk);
```

After `Start()` returns, the main thread's existing message loop processes CIP responses. `BrowseStarted` events fire immediately; `BrowseAddressFound` events arrive within seconds as devices respond.

---

## 5. Event Sink Pattern

### DualEventSink Class

Implements dual-dispatch sink with connection points for both bus events and online notify events:

```cpp
class DualEventSink : public IRSTopologyOnlineNotify,  // vtable at +0
                      public ITopologyBusEvents         // vtable at +4
{
    BYTE m_pad[2048];       // +8: absorbs Start()'s direct memory writes
    LONG m_refCount;        // behind padding — safe
    IUnknown* m_pFTM;      // Free Threaded Marshaler
    DWORD m_magic;          // 0xDEADBEEF integrity check
    std::wstring m_label;   // identity for logging
    CRITICAL_SECTION m_cs;  // lock for cycle detection
    std::set<std::wstring> m_seenAddresses;
    volatile bool m_cycleComplete;
    volatile bool m_browseEnded;
    int m_addressCount;
};
```

**Design decisions:**

- **2048-byte padding** at offset +8 absorbs `Start()`'s direct memory writes to the sink object (observed up to ~528 bytes written past the vtable pointers).
- **Accept-all QueryInterface** returns `S_OK` for any unknown IID — the enumerator's undocumented connection points must connect successfully.
- **Free Threaded Marshaler** (`CoCreateFreeThreadedMarshaler`) enables event delivery across apartment boundaries.
- **Thread-safe logging** via `CRITICAL_SECTION` since events fire on the main STA while the worker thread also logs.

### Connection Point Setup

```cpp
// 1. Get IConnectionPointContainer from bus
IConnectionPointContainer* pCPC = nullptr;
pBusUnk->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);

// 2. Enumerate all connection points
IEnumConnectionPoints* pEnumCP = nullptr;
pCPC->EnumConnectionPoints(&pEnumCP);

// 3. Advise sink on each connection point
IConnectionPoint* pCP = nullptr;
while (pEnumCP->Next(1, &pCP, nullptr) == S_OK) {
    DWORD cookie;
    pCP->Advise(static_cast<IRSTopologyOnlineNotify*>(pSink), &cookie);
    // Store pCP + cookie for later Unadvise
}
```

### Event Callbacks

**IRSTopologyOnlineNotify** (online browse events):
- `BrowseStarted(IUnknown* pBus)` — browse initiated
- `Found(IUnknown* pBus, VARIANT address)` — device/slot found
- `NothingAtAddress(IUnknown* pBus, VARIANT addr)` — no response at address
- `BrowseCycled(IUnknown* pBus)` — one browse cycle complete
- `BrowseEnded(IUnknown* pBus)` — browse finished

**ITopologyBusEvents** (bus-level events):
- `OnBrowseStarted(IUnknown* pBus)` — bus browse started
- `OnBrowseAddressFound(IUnknown* pBus, VARIANT addr)` — address discovered
- `OnPortConnect / OnPortDisconnect` — port state changes
- `OnBrowseCycled / OnBrowseEnded` — bus browse lifecycle

### Cycle Detection

The sink tracks browse completion through multiple signals:

1. **Repeat address** — if `Found()` delivers an address already in `m_seenAddresses`, the browse has cycled
2. **BrowseCycled event** — explicit cycle notification
3. **BrowseEnded event** — browse terminated

The worker thread polls `m_cycleComplete` on all active sinks to determine when to advance to the next phase.

---

## 6. Backplane Browsing

After Ethernet-level browse identifies devices, a second pass discovers backplane modules (individual slots within controllers/chassis).

### Step 1: Find Backplane Port

Navigate from device to its backplane port via vtable:

```cpp
// QI device for IRSTopologyDevice
IUnknown* pDevUnk = nullptr;
pDevDisp->QueryInterface(IID_IRSTopologyDevice, (void**)&pDevUnk);

// Call GetBackplanePort at vtable[19]
typedef HRESULT (__stdcall *GetBackplanePortFn)(void*, IUnknown**);
void** vtable = *(void***)pDevUnk;
GetBackplanePortFn pfn = (GetBackplanePortFn)vtable[19];

IUnknown* pPort = nullptr;
hr = pfn(pDevUnk, &pPort);
// S_OK = has backplane, E_FAIL = no backplane (e.g., switches)
```

### Step 2: Get Backplane Bus

**Primary approach — IRSTopologyPort::GetBus at vtable[10]:**

```cpp
IUnknown* pPortUnk = nullptr;
pPort->QueryInterface(IID_IRSTopologyPort, (void**)&pPortUnk);

typedef HRESULT (__stdcall *GetBusFn)(void*, IUnknown**);
void** portVtable = *(void***)pPortUnk;
GetBusFn pfnGetBus = (GetBusFn)portVtable[10];

IUnknown* pBusUnk = nullptr;
hr = pfnGetBus(pPortUnk, &pBusUnk);
```

**Fallback approach — DISPID 38 with port name:**

If `GetBus[10]` returns `E_POINTER` (common — the port's internal bus pointer may be null even when the bus exists in topology), use the universal DISPID 38:

```cpp
// Try known port names in order
const wchar_t* portNames[] = {L"Backplane", L"CompactBus", L"PointBus", L"Chassis", L"BP"};
for (auto name : portNames) {
    VARIANT arg;
    arg.vt = VT_BSTR;
    arg.bstrVal = SysAllocString(name);
    DISPPARAMS dp = { &arg, nullptr, 1, 0 };
    VARIANT result;
    hr = pDevDisp->Invoke(38, ..., &dp, &result, ...);
    if (SUCCEEDED(hr) && result.vt == VT_DISPATCH) {
        pBackplaneBus = result.pdispVal;
        break;
    }
}
```

### Step 3: Browse Backplane Bus

Same pattern as Ethernet browse: QI backplane bus for `IOnlineEnumeratorTypeLib`, create `DualEventSink`, advise connection points, call `Start(busPath)`.

Backplane browses discover individual modules at slot addresses (e.g., "0", "1", "2"... "12").

---

## 7. Topology Polling

### SaveTopologyXML

Uses `IRSTopologyGlobals` to save the current topology state as XML:

```cpp
// TopologyGlobals exposes SaveAsXML through IDispatch
// The resulting XML has structure:
// <topology>
//   <tree>
//     <device name="Workstation" classname="Workstation">
//       <port name="Test">
//         <bus name="Ethernet" classname="AB_ETH">
//           <address type="String" value="10.13.30.68">
//             <device name="1756-EN2T (3)" classname="1756-EN2T/B" online="true">
//               <port name="Backplane">
//                 <bus name="Backplane (6)">
//                   <address type="Short" value="0">
//                     <device name="1756-L85E" classname="1756-L85E/B">
```

### Device Counting

From topology XML:
- **Total devices**: count of `<device>` elements (excluding `reference` attributes)
- **Identified devices**: devices with `classname` not equal to `"Unrecognized Device"` or `"Workstation"`
- **Target identified**: target IP's device has a non-"Unrecognized" classname

### Polling Strategy

- **Phase 3** (Ethernet): Every 2 seconds, max 30s. Exit early when target IP identified or all enumerators cycled.
- **Phase 5/5b** (backplane): Event-driven with 2s check intervals, max 30s. Exit when all backplane enumerators have cycled.
- **Monitor mode**: Every 10 seconds. Auto-triggers bus/backplane browse when new devices appear.

---

## 8. Phase Architecture

```
DLL_PROCESS_ATTACH → CreateThread(WorkerThread)

WorkerThread:
│
├─ Connect pipe, read config
├─ CoInitializeEx(STA)
├─ Create HarmonyServices → TopologyGlobals → Buses
│
├─ Phase 1: ConnectNewDevice ─────────────── DISPID 54 per IP per driver
│   Runs on: Worker STA (COM marshals to main STA internally)
│
├─ Phase 2: Main-STA Browse ─────────────── SetWindowsHookEx → DoMainSTABrowse
│   Runs on: Main STA thread (via hook)
│   Actions: Bus QI → Enumerator → Advise Sink → Start(path)
│
├─ [Monitor mode branches to RunMonitorLoop here]
│
├─ Phase 3: Topology Polling ─────────────── 2s interval, max 30s
│   Runs on: Worker STA
│   Exit: Target identified OR all enumerators cycled
│
├─ Phase 4: Bus Browse (Backplanes) ──────── ExecuteOnMainSTA(DoBusBrowse)
│   Runs on: Main STA thread (via hook)
│   Actions: Enumerate devices → GetBackplanePort[19] → QI enumerator → Start
│
├─ Phase 5: Bus Browse Polling ───────────── Event-driven, max 30s
│   Runs on: Worker STA
│   Exit: All bus enumerators cycled
│
├─ Phase 4b: Backplane Bus Browse ────────── ExecuteOnMainSTA(DoBackplaneBrowse)
│   Runs on: Main STA thread (via hook)
│   Actions: Port → GetBus[10] / DISPID 38 → QI enumerator → Start
│
├─ Phase 5b: Backplane Module Polling ────── Event-driven, max 30s
│   Runs on: Worker STA
│   Exit: All backplane enumerators cycled
│
├─ Phase 6: Cleanup ──────────────────────── ExecuteOnMainSTA(DoCleanupOnMainSTA)
│   Runs on: Main STA thread (via hook)
│   Actions: Stop enumerators[8] → Unadvise CPs → Release all
│
└─ Write results, close pipe, CoUninitialize
```

### Monitor Mode Flow

```
Phases 1-2 (same as inject)
    │
    └─ RunMonitorLoop:
        while no STOP signal:
            pump messages (100ms)
            every 10s:
                save topology snapshot
                send status via pipe (S|total|ident|events)
                send topology via pipe (X|BEGIN...X|END)
                if identified > 0 and !busBrowseDone:
                    ExecuteOnMainSTA(DoBusBrowse)
                if capturedBuses not empty and !bpBrowseDone:
                    ExecuteOnMainSTA(DoBackplaneBrowse)
                write hook_results.txt
        on STOP:
            ExecuteOnMainSTA(DoCleanupOnMainSTA)
            send D| via pipe
            exit
```

---

## 9. Threading Model

### Why STA?

RSLinx's COM objects are apartment-threaded. All topology interfaces must be accessed from an STA.

The DLL creates two STA contexts:
1. **Worker STA** — `CoInitializeEx(COINIT_APARTMENTTHREADED)` on the worker thread. Used for ConnectNewDevice (marshaled), topology polling, and message pumping.
2. **Main STA** — RSLinx's main thread already has an STA. Browse operations execute here via thread hook because ENGINE.DLL's `WSAAsyncSelect` binds CIP socket I/O to this thread's message pump.

### ExecuteOnMainSTA Mechanism

```
Worker Thread                        Main Thread
     |                                    |
  FindMainThreadId()                      |
     |                                    |
  SetWindowsHookEx(WH_GETMESSAGE, ----> installs hook
      MainSTAHookProc, mainTID)           |
     |                                    |
  g_browseRequested = 1                   |
  PostThreadMessage(WM_NULL, MAGIC) ---> message arrives
     |                                 MainSTAHookProc:
     |                                   if magic match && requested:
     |                                     hr = g_pMainSTAFunc()
     |                                     g_browseResult = hr
  poll g_browseResult <--------------- done
  UnhookWindowsHookEx                     |
```

**Fallback** (if PostThreadMessage fails):
```
Worker Thread                        Main Thread
     |                                    |
  EnumThreadWindows → find hwnd           |
  SetWindowLongPtrW(GWLP_WNDPROC, ---> subclass installed
      HookWndProc)                        |
  SendMessage(hwnd, SUBCLASS_MSG) ----> HookWndProc:
     |  (blocks)                          hr = g_pMainSTAFunc()
     |                                    restore original WndProc
  returns with hr <-------------------- done
```

### DllMain Safety

`DLL_PROCESS_DETACH` does **not** call `WaitForSingleObject` on the worker thread — this would deadlock under the loader lock while the thread performs COM cleanup. Instead:

```cpp
case DLL_PROCESS_DETACH:
    g_shouldStop = true;           // signal worker to exit
    CloseHandle(g_hWorkerThread);  // release handle (no wait)
    break;
```

The worker thread exits naturally after finishing its current phase or when it checks `g_shouldStop`.

---

## 10. COM GUID Reference

### CLSIDs

| Name | GUID | Purpose |
|------|------|---------|
| `CLSID_HarmonyServices` | `{043D7910-38FB-11D2-903D-00C04FA363C1}` | COM entry point for RSLinx |
| `CLSID_RSTopologyGlobals` | `{38593054-38E4-11D0-BE25-00C04FC2AA48}` | Global topology access |
| `CLSID_RSProjectGlobal` | `{C92DFEA6-1D29-11D0-AD3F-00C04FD915B9}` | Project factory |
| `CLSID_OnlineBusExt` | `{2EC6B980-C629-11D0-BDCF-080009DC75C8}` | Standalone enumerator (do NOT use for browse) |
| `CLSID_RSPath` | `{6DBDFEB8-F703-11D0-AD73-00C04FD915B9}` | Path composer |
| `CLSID_EthernetPort` | `{00000001-0000-11D0-BDB8-080009DC75C8}` | Ethernet port type |
| `CLSID_EthernetBus` | `{00010010-5DFF-11CF-B4B9-C46F03C10000}` | Ethernet bus type |
| Unrecognized Device | `{00000004-5D68-11CF-B4B9-C46F03C10000}` | Placeholder for ConnectNewDevice |

### IIDs

| Name | GUID | Purpose |
|------|------|---------|
| `IID_IHarmonyConnector` | `{19EECB80-3868-11D2-903D-00C04FA363C1}` | Harmony services entry |
| `IID_IRSTopologyGlobals` | `{640DAC76-38E3-11D0-BE25-00C04FC2AA48}` | Topology globals interface |
| `IID_IRSProjectGlobal` | `{B286B85E-211C-11D0-AD42-00C04FD915B9}` | Project global interface |
| `IID_IRSProject` | `{D61BFDA0-EEAB-11CE-B4B5-C46F03C10000}` | Project interface |
| `IID_ITopologyBus` | `{25C81D16-F7BA-11D0-AD73-00C04FD915B9}` | Bus interface |
| `IID_ITopologyBusEvents` | `{AFF2BF80-8D86-11D0-B77F-F87205C10000}` | Bus event sink |
| `IID_ITopologyPathComposer` | `{EF3A0832-940C-11D1-ADAC-00C04FD915B9}` | Path composer interface |
| `IID_ITopologyDevice_Dual` | `{B2A20A5E-F7B9-11D0-AD73-00C04FD915B9}` | Device dispatch (DISPID 38) |
| `IID_IOnlineEnumeratorTypeLib` | `{FC357A88-0A98-11D1-AD78-00C04FD915B9}` | Enumerator via bus QI |
| `IID_IOnlineEnumerator` | `{91748520-A51B-11D0-BDC9-080009DC75C8}` | Online enumerator interface |
| `IID_IRSTopologyNetwork` | `{46DAD8E4-4048-11D0-BE26-00C04FC2AA48}` | Network interface |
| `IID_IRSTopologyObject` | `{DCEAD8E0-2E7A-11CF-B4B5-C46F03C10000}` | Base topology object |
| `IID_IRSTopologyDevice` | `{DCEAD8E1-2E7A-11CF-B4B5-C46F03C10000}` | Device vtable (GetBackplanePort) |
| `IID_ITopologyCollection` | `{2D76DE6C-94A0-11D0-AD56-00C04FD915B9}` | Device collection |
| `IID_IRSTopologyPort` | `{98E549B2-B27E-11D0-AD5E-00C04FD915B9}` | Port vtable (GetBus) |
| `IID_IRSObject` | `{94CB2140-450F-11CF-B4B5-C46F03C10000}` | Base object (GetName) |
| `IID_IRSTopologyOnlineNotify` | `{FA5D9CF0-A259-11D1-BE10-080009DC75C8}` | Online browse notify events |

---

## 11. DISPID Reference

| DISPID | Interface | Method | Args | Returns | Purpose |
|--------|-----------|--------|------|---------|---------|
| -4 | ITopologyCollection | `_NewEnum` | none | IEnumVARIANT | Enumerate collection items |
| 1 | ITopologyCollection | `Count` | none | VT_I4 | Number of items |
| 1 | ITopologyObject | `Name` | none | VT_BSTR | Object name |
| 2 | ITopologyObject | `objectid` | none | VT_BSTR | Topology GUID string |
| 4 | ITopologyBus | `path(flags)` | VT_I4 flags=0 | VT_DISPATCH | Get path composer object |
| 38 | ITopologyDevice_Dual | `GetBusEx(name)` | VT_BSTR | VT_DISPATCH | Bus by port name |
| 50 | ITopologyBus | `Devices` | none | VT_DISPATCH | Get devices collection |
| 54 | ITopologyBus | `ConnectNewDevice` | 6 args | VT_DISPATCH | Add device to topology |

**DISPID 4 note:** With no arguments, routes to `IADs.GUID` and returns `TYPE_E_BADMODULEKIND`. Must pass `VT_I4 flags=0` to get the topology path.

**DISPID 54 arguments** (in reverse order per COM convention):

| Index | Type | Value | Purpose |
|-------|------|-------|---------|
| args[5] | VT_I4 | 0 | Flags |
| args[4] | VT_BSTR | `{00000004-...}` | Device class GUID |
| args[3] | VT_EMPTY | - | Reserved |
| args[2] | VT_BSTR | `"Device"` | Type label |
| args[1] | VT_BSTR | `"A"` | Port name |
| args[0] | VT_BSTR | IP address | Network address |

---

## 12. Vtable Layouts

### IRSTopologyDevice (IID_IRSTopologyDevice)

After IUnknown (slots 0-2) and IRSTopologyObject base (slots 3-13):

| Slot | Method | Signature | Notes |
|------|--------|-----------|-------|
| 14 | AddPort | `(GUID*, label, rect*, hwnd, IID*, ppResult)` | |
| 15 | RemovePort | | |
| 16 | GetPortCount | | |
| 17 | GetPortList | | |
| 18 | GetPort | | |
| **19** | **GetBackplanePort** | `(IUnknown** ppPort)` | Key method for backplane discovery |
| 20 | GetModuleWidth | | |

### IRSTopologyPort (IID_IRSTopologyPort)

After IUnknown (slots 0-2) and IRSTopologyObject base (slots 3-13):

| Slot | Method | Signature | Notes |
|------|--------|-----------|-------|
| 14 | GetDevice | `(ITopologyDevice**)` | Returns owning device |
| 15 | GetLabel | `(WCHAR*, int)` | Port label |
| 16 | GetID | `(GUID*)` | Port identifier |
| 17 | GetType | `(CLSID*, WCHAR*, int)` | Port type info |
| 18 | IsType | `(CLSID byval!)` | **DANGER: 16-byte value param corrupts stack** |
| 19 | IsBackplane | `(ITopologyPort**)` | Returns backplane port or E_FAIL |
| 20 | IsConnected | `(BOOL*)` | Connection status |
| **21** | **GetBus** | `(ITopologyBus**)` | Often returns E_POINTER — use DISPID 38 instead |

**Warning:** Slot 18 (`IsType`) takes a `CLSID` by value (16 bytes on stack). Calling with a pointer argument causes `__stdcall` stack corruption — the callee pops 20 bytes but only 8 were pushed.

### IRSObject (IID_IRSObject)

| Slot | Method | Signature | Notes |
|------|--------|-----------|-------|
| **7** | **GetName** | `(WCHAR* buffer, int bufLen)` | Object name string |

### IOnlineEnumerator

| Slot | Method | Signature | Notes |
|------|--------|-----------|-------|
| **7** | **Start** | `(IUnknown* pPath)` | Initiates browse on current STA |
| **8** | **Stop** | `()` | Stops active browse |

---

## 13. What Works / What Doesn't

### Proven Techniques

| Technique | Result |
|-----------|--------|
| STA threading (`COINIT_APARTMENTTHREADED`) | Required for all COM calls |
| ConnectNewDevice (DISPID 54, 6 args) | Devices added; smart skip if exists |
| bus.path(flags=0) via DISPID 4 with VT_I4 | Returns valid path object |
| SetWindowsHookEx(WH_GETMESSAGE) on main thread | Executes functions on correct STA |
| Window subclass fallback | Alternative if hook fails |
| **Bus QI for IOnlineEnumeratorTypeLib** | Returns connected enumerator |
| Start(path) via bus-QI'd enumerator on main STA | Devices identified via CIP |
| Device QI for IOnlineEnumeratorTypeLib | Device enumerator for backplane structure |
| GetBackplanePort at IRSTopologyDevice vtable[19] | Detects which devices have backplanes |
| **DISPID 38 on ITopologyDevice_Dual** | Universal bus-by-port-name accessor |
| IRSTopologyPort::GetBus at vtable[10] | Returns backplane bus (when available) |
| DISPID 38 with "Backplane" / "CompactBus" | Fallback for backplane bus access |
| Backplane bus QI + Start(busPath) | Discovers individual slot modules |
| DualEventSink with 2048-byte padding | Survives Start()'s direct memory writes |
| Free Threaded Marshaler on sink | Cross-apartment event delivery |
| Accept-all QI on sink | Required for undocumented connection points |

### Failed Approaches (Do Not Retry)

| Technique | Error | Reason |
|-----------|-------|--------|
| IRSTopologyNetwork::Browse(0) | E_NOINTERFACE | Bus doesn't implement this |
| Start(path) from worker STA | S_OK, 0 events | Wrong apartment for ENGINE.DLL callbacks |
| CoCreateInstance(OnlineBusExt) standalone | S_OK, 0 events | Not connected to bus CIP I/O |
| bus.path with no arguments | TYPE_E_BADMODULEKIND | Routes to IADs.GUID, not topology path |
| IDispatch::Invoke for Start | TYPE_E_LIBNOTREGISTERED | No TypeLib for IOnlineEnumerator |
| MTA threading | Various failures | RSLinx COM objects require STA |
| IRSTopologyPort::GetBus at vtable[21] | E_POINTER | Internal pointer often null |
| IRSTopologyPort vtable[18] (IsType) | Stack corruption | CLSID by value (16 bytes) with __stdcall |
| Port _NewEnum / DISPID 50 | DISP_E_MEMBERNOTFOUND | Port dispatch doesn't support enumeration |
| BrowseStarted capture for buses | Captures devices | IRSTopologyOnlineNotify passes device objects, not buses |
| PipeOptions.Asynchronous (.NET) | Pipe broken / hang | Client must open with FILE_FLAG_OVERLAPPED to match |
| FILE_FLAG_OVERLAPPED with sync ReadFile | Undefined behavior | Must use OVERLAPPED structures consistently |
| std::wcout with Unicode (em dash) | Silent output loss | Use WriteConsoleW instead |
