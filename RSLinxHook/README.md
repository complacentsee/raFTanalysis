# RSLinxHook — In-Process RSLinx Topology Discovery DLL

Injected DLL that runs inside `rslinx.exe` to perform COM-based device discovery, browse triggering, and backplane enumeration against the RSLinx topology engine.

## Why a DLL?

RSLinx Classic's COM objects are in-process STA only. Cross-process COM calls return `E_NOINTERFACE`. The DLL runs inside RSLinx's address space where `ENGINE.DLL` is loaded, giving direct access to the topology COM interfaces and the main thread's message pump (required for CIP I/O callbacks).

## Build

**Requirements**: Visual Studio 2019 (v142 toolset), Win32/x86 only

```
MSBuild.exe RSLinxHook.vcxproj -p:Configuration=Release -p:Platform=Win32
```

Output: `Release/RSLinxHook.dll`

## How It Works

The DLL is loaded via `CreateRemoteThread(LoadLibraryW)` by either RSLinxBrowse.exe (CLI) or RSLinxViewer.exe (TUI). On `DLL_PROCESS_ATTACH`, it spawns a worker thread that executes a six-phase discovery sequence:

### Phase 1: ConnectNewDevice
Adds IP addresses to the topology as unrecognized devices via `IDispatch::Invoke(DISPID 54)`. Smart: skips devices that already exist (`DISP_E_EXCEPTION`).

### Phase 2: Main-STA Browse
Installs a `WH_GETMESSAGE` hook on RSLinx's main thread to execute `IOnlineEnumerator::Start(path)` on the correct STA. This is required because `ENGINE.DLL` uses `WSAAsyncSelect` to bind CIP socket I/O to the main thread's message pump.

### Phase 3: Topology Polling
Polls topology XML every 2 seconds. Exits early when the target device is identified or all enumerators have cycled.

### Phase 4: Bus Browse (Backplanes)
Navigates each Ethernet device's backplane port via `IRSTopologyDevice::GetBackplanePort()` (vtable[19]). Starts device-level enumerators to discover backplane bus structure.

### Phase 4b: Backplane Bus Browse
Gets backplane bus objects via `DISPID 38` (universal bus-by-port-name accessor) and starts enumerators on each backplane bus to discover individual slot modules.

### Phase 5/5b: Event-Driven Polling
Waits for BrowseCycled/BrowseEnded events from enumerators. Exits when all active enumerators have completed their cycles.

### Phase 6: Cleanup
Stops all enumerators, unadvises all connection points, releases all COM objects.

## Modes

- **Inject** (default): Runs Phases 1-6, writes results, exits.
- **Monitor**: Runs Phases 1-2, then enters a continuous polling loop with periodic topology snapshots. Triggers bus/backplane browse when new devices appear. Exits on STOP signal.

## IPC: Named Pipe Protocol

Communicates with the launcher (RSLinxBrowse or RSLinxViewer) via bidirectional named pipe `\\.\pipe\RSLinxHook`:

**Launcher -> Hook** (config, then stop signal):
```
C|MODE=inject          config line
C|DRIVER=SUFF2         driver name
C|IP=10.13.30.68       IP address
C|DEBUGXML=1           enable debug XML
C|END                  config complete — hook proceeds
STOP                   stop signal (Ctrl+C in launcher)
```

**Hook -> Launcher** (log, status, topology, done):
```
L|<text>               log line (UTF-8)
S|109|107|22           status: total|identified|events
X|BEGIN ... X|END      topology XML block
D|                     done — hook is finished
```

Falls back to file-based config (`C:\temp\hook_config.txt`) if no pipe is available.

## Output Files

| File | Content |
|------|---------|
| `hook_log.txt` | Detailed execution log |
| `hook_results.txt` | Summary (DEVICES_IDENTIFIED, TARGET_STATUS, etc.) |
| `hook_topo_before.xml` | Topology snapshot before browse |
| `hook_topo_after.xml` | Final topology snapshot |

## Source

Single file: `RSLinxHook.cpp` (~3800 lines)

For detailed COM interface documentation, see [COM_ARCHITECTURE.md](COM_ARCHITECTURE.md).
