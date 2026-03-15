# RSLinxHook — In-Process RSLinx Topology Discovery DLL

Injected DLL that runs inside `rslinx.exe` to perform COM-based device discovery, browse triggering, and backplane enumeration against the RSLinx topology engine.

## Why a DLL?

RSLinx Classic's COM objects are in-process STA only. Cross-process COM calls return `E_NOINTERFACE`. The DLL runs inside RSLinx's address space where `ENGINE.DLL` is loaded, giving direct access to the topology COM interfaces and the main thread's message pump (required for CIP I/O callbacks).

## Build

**Requirements**: VS Build Tools 2026 (v145 toolset), Win32/x86 only

```
MSBuild.exe RSLinxHook.vcxproj -p:Configuration=Release -p:Platform=Win32
```

Output: `Release/RSLinxHook.dll`

## How It Works

The DLL is loaded via `CreateRemoteThread(LoadLibraryW)` by either RSLinxBrowse.exe (CLI) or RSLinxViewer.exe (TUI). On `DLL_PROCESS_ATTACH`, it spawns a worker thread that:

1. Creates the named pipe server `\\.\pipe\RSLinxHook`
2. Enters an accept loop — waits for a client to connect, handles the session, then loops back to accept the next client
3. The hook remains alive (and the pipe server remains open) until the RSLinx process exits

### Browse Phases

Each client session that sends a full config triggers a six-phase discovery sequence:

**Phase 1: ConnectNewDevice**
Adds IP addresses to the topology as unrecognized devices via `IDispatch::Invoke(DISPID 54)`. Skips devices that already exist (`DISP_E_EXCEPTION`).

**Phase 2: Main-STA Browse**
Installs a `WH_GETMESSAGE` hook on RSLinx's main thread to execute `IOnlineEnumerator::Start(path)` on the correct STA. Required because `ENGINE.DLL` uses `WSAAsyncSelect` to bind CIP socket I/O to the main thread's message pump.

**Phase 3: Topology Polling**
Polls topology XML every 2 seconds. Exits early when the target device is identified or all enumerators have cycled.

**Phase 4: Bus Browse (Backplanes)**
Navigates each Ethernet device's backplane port via `IRSTopologyDevice::GetBackplanePort()` (vtable[19]). Starts device-level enumerators to discover backplane bus structure.

**Phase 4b: Backplane Bus Browse**
Gets backplane bus objects via `DISPID 38` and starts enumerators on each backplane bus to discover individual slot modules.

**Phase 5/5b: Event-Driven Polling**
Waits for BrowseCycled/BrowseEnded events from enumerators. Exits when all active enumerators have completed their cycles.

**Phase 6: Cleanup**
Stops all enumerators, unadvises all connection points, releases all COM objects.

## Modes

- **Inject** (default): Runs Phases 1–6, caches topology, sends `D|`, enters command loop waiting for `Q|`/`B|`/`STOP`.
- **Monitor**: Runs Phases 1–2, then enters a continuous polling loop with periodic topology snapshots. Triggers bus/backplane browse when new devices appear. Exits on STOP signal.

## IPC: Named Pipe Protocol

The hook is the **pipe server** (`\\.\pipe\RSLinxHook`, `nMaxInstances=1`). Clients (RSLinxBrowse, RSLinxViewer) connect as pipe clients. Only one client may be connected at a time; subsequent clients block until the current session sends `STOP`.

**Client → Hook:**

```
C|MODE=inject          browse mode (or monitor)
C|DRIVER=Test          driver name
C|IP=192.168.1.55      IP address (repeatable)
C|NEWDRIVER=1          hot-load new driver into RSLinx
C|DEBUGXML=1           enable debug XML snapshots
C|END                  config complete — hook proceeds with browse
Q|192.168.1.55\Backplane\1   query cached topology for path
B|                     trigger re-browse on existing connection
STOP                   end session (hook loops back to accept-wait)
```

**Hook → Client:**

```
L|<text>               log line (UTF-8)
S|109|107|22           status: total|identified|events
X|BEGIN ... X|END      topology XML block
D|                     browse complete — command loop open for Q|/B|/STOP
R|FOUND|...            query result: path found, pipe-delimited fields
R|NOTFOUND|path        query result: path not in cached topology
```

`D|` does **not** end the session. After `D|`, the hook waits in a command loop for `Q|` queries, `B|` re-browse requests, or `STOP`. The pipe stays open until `STOP` is received or the client disconnects.

Falls back to file-based config (`C:\temp\hook_config.txt`) if no pipe client connects within the startup window.

## Output Files

| File | Content |
|------|---------|
| `hook_log.txt` | Detailed execution log |
| `hook_results.txt` | Summary (DEVICES_IDENTIFIED, TARGET_STATUS, etc.) |
| `hook_topo_before.xml` | Topology snapshot before browse |
| `hook_topo_after.xml` | Final topology snapshot |

## Source

Main file: `DllMain.cpp`

For detailed COM interface documentation, see [COM_ARCHITECTURE.md](COM_ARCHITECTURE.md).
