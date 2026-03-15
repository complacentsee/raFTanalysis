# RSLinxBrowse — CLI RSLinx Topology Browser

Command-line tool that injects RSLinxHook.dll into rslinx.exe to discover and identify EtherNet/IP devices on RSLinx Classic drivers. Supports full browse, monitor mode, and fast path queries against a cached topology.

## Build

**Requirements**: VS Build Tools 2026 (v145 toolset), Win32/x86 only

```
MSBuild.exe RSLinxBrowse.vcxproj -p:Configuration=Release -p:Platform=Win32
```

Output: `Release/RSLinxBrowse.exe`

RSLinxHook.dll must also be built and placed in the same `Release/` directory as the executable.

## Usage

```
RSLinxBrowse.exe [--driver NAME [--ip IP]...] [--query PATH] [--monitor] [--debug-xml] [--logdir DIR]
```

### Examples

```bash
# Browse "Test" driver, inject hook (or reuse existing), discover all devices
RSLinxBrowse.exe --driver Test --ip 192.168.1.55

# Multiple IPs on one driver
RSLinxBrowse.exe --driver Test --ip 10.39.31.200 --ip 10.39.33.87

# Query cached topology — fast (~1s), no re-browse
RSLinxBrowse.exe --query 192.168.1.55\Backplane\2

# Query with explicit backplane slot check (slot 99 = NOTFOUND)
RSLinxBrowse.exe --query 192.168.1.55\Backplane\99

# Monitor mode (browse existing driver, no registry changes)
RSLinxBrowse.exe --driver SUFF2 --monitor

# Positional args (backward compat): RSLinxBrowse.exe <driver> <ip>
RSLinxBrowse.exe SUFF2 10.13.30.68
```

### Flags

| Flag | Description |
|------|-------------|
| `--driver NAME` | Driver name (default: `Test`) |
| `--ip IP` | IP address to add to driver (repeatable) |
| `--query PATH` | Query cached topology for a path (e.g. `192.168.1.55\Backplane\1`) |
| `--monitor` | Browse existing driver topology without creating/modifying drivers |
| `--inject` | Default mode (accepted for backward compat) |
| `--debug-xml` | Write topology XML snapshots at each polling interval |
| `--logdir DIR` | Log directory (default: `C:\temp`) |

## How It Works

### Browse Mode (default)

1. **Auto-create driver** — If the named driver doesn't exist in the RSLinx registry (`HKLM\...\AB_ETH`), creates it with the target IPs in the Node Table and restarts RSLinx
2. **Find RSLinx** — Locates `rslinx.exe` PID via `CreateToolhelp32Snapshot`
3. **Smart injection** — Probes `\\.\pipe\RSLinxHook` (500 ms timeout). If the hook is already running, connects directly without re-injecting. If not found, injects RSLinxHook.dll via `CreateRemoteThread(LoadLibraryW)` and waits up to 10 s for the pipe server to appear.
4. **Send config over pipe** — Driver names, IPs, and mode (`C|MODE`, `C|DRIVER`, `C|IP`, `C|END`)
5. **Stream log output** — Reads `L|` log lines from pipe and prints to console in real-time
6. **Wait for completion** — `D|` signal means hook finished all browse phases; pipe stays open
7. **Display results** — Reads `hook_results.txt` for final summary
8. **Clean exit** — Sends `STOP`, closes pipe handle (hook remains injected for future sessions)

### Query Mode (`--query PATH`)

Skips all browse phases. Connects to the already-running hook (or injects if needed), sends `C|END` to complete config handshake, then sends `Q|PATH`. Receives `R|FOUND|...` or `R|NOTFOUND|...` and exits. Total time ~1 s.

If NOTFOUND, the cache may be empty (hook freshly injected) or the device is genuinely absent. Run a full browse first if uncertain.

### Monitor Mode

Same as browse except:
- Does not create or modify driver registry entries
- Driver must already exist in RSLinx
- Press Ctrl+C to send `STOP` signal and exit cleanly

## Persistent Hook

The hook DLL stays injected in RSLinx after RSLinxBrowse exits. Each subsequent run reconnects to the running hook without re-injecting. The topology cache (browsed device tree) persists between sessions.

The hook is unloaded only when the RSLinx service stops or restarts.

## IPC

All communication with RSLinxHook.dll uses a bidirectional named pipe. The **hook is the pipe server**; RSLinxBrowse is the client. No intermediate files are used for config delivery, stop signaling, or completion detection.

The hook still writes `hook_log.txt` and `hook_results.txt` to the log directory for standalone inspection.

## Project Structure

```
RSLinxBrowse/
├── main.cpp                — CLI entry point, DLL injection, pipe client
├── DriverConfig.h/cpp      — RSLinx driver registry reader, Node Table management
├── TopologyBrowser.h/cpp   — External COM topology browser (post-inject verification)
├── BrowseEventSink.h/cpp   — COM event sink (external browse mode)
├── RSLinxInterfaces.h      — COM interface GUIDs and declarations
└── RSLinxBrowse.vcxproj    — Build config (Win32/x86)
```

## Output

After a successful run, `C:\temp\` contains:

| File | Content |
|------|---------|
| `hook_log.txt` | Detailed execution log from the injected DLL |
| `hook_results.txt` | Summary: device counts, target status, elapsed time |
| `hook_topo_before.xml` | Topology snapshot before browse |
| `hook_topo_after.xml` | Final topology snapshot after all phases |
