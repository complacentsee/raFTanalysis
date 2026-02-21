# RSLinxBrowse — CLI RSLinx Topology Browser

Command-line tool that injects RSLinxHook.dll into rslinx.exe to discover and identify EtherNet/IP devices on RSLinx Classic drivers.

## Build

**Requirements**: Visual Studio 2019 (v142 toolset), Win32/x86 only

```
MSBuild.exe RSLinxBrowse.vcxproj -p:Configuration=Release -p:Platform=Win32
```

Output: `Release/RSLinxBrowse.exe`

RSLinxHook.dll must also be built and placed where the executable can find it (same directory or `Release/` subfolder).

## Usage

```
RSLinxBrowse.exe [--driver NAME] [--ip IP] [--monitor] [--debug-xml] [--logdir DIR]
```

### Examples

```bash
# Browse default "Test" driver, auto-create if missing
RSLinxBrowse.exe

# Browse specific driver with target IP
RSLinxBrowse.exe --driver SUFF2 --ip 10.13.30.68

# Multiple IPs
RSLinxBrowse.exe --driver Test --ip 10.39.31.200 --ip 10.39.33.87

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
| `--monitor` | Browse existing driver topology without creating/modifying drivers |
| `--inject` | Default mode (accepted for backward compat) |
| `--debug-xml` | Write topology XML snapshots at each polling interval |
| `--logdir DIR` | Log directory (default: `C:\temp`) |

## How It Works

### Default Mode (Browse)

1. **Auto-create driver** — If the named driver doesn't exist in the RSLinx registry (`HKLM\...\AB_ETH`), creates it with the target IPs in the Node Table and restarts RSLinx
2. **Find RSLinx** — Locates `rslinx.exe` PID via `CreateToolhelp32Snapshot`
3. **Create named pipe** — `\\.\pipe\RSLinxHook` (bidirectional, synchronous)
4. **Eject old DLL** — If RSLinxHook.dll is already loaded from a prior run, ejects via `FreeLibrary`
5. **Inject DLL** — `LoadLibraryW` via `CreateRemoteThread` into RSLinx
6. **Send config over pipe** — Driver names, IPs, and mode
7. **Stream log output** — Reads `L|` log lines from pipe and prints to console in real-time
8. **Wait for completion** — `D|` signal means hook finished all phases
9. **Display results** — Reads `hook_results.txt` for final summary

### Monitor Mode

Same as default except:
- Does not create or modify driver registry entries
- Driver must already exist in RSLinx
- Hook enters continuous polling loop instead of six-phase sequence
- Press Ctrl+C to send `STOP` signal and exit cleanly

## IPC

All communication with RSLinxHook.dll uses a bidirectional named pipe. No intermediate files are used for config delivery, stop signaling, or completion detection.

The hook still writes `hook_log.txt` and `hook_results.txt` to the log directory for standalone inspection.

## Project Structure

```
RSLinxBrowse/
├── main.cpp                — CLI entry point, DLL injection, pipe server
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
