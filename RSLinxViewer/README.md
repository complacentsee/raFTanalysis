# RSLinxViewer — TUI Topology Viewer

.NET 8 terminal UI application that injects RSLinxHook.dll into rslinx.exe and displays the topology tree in real-time using Spectre.Console.

## Build

**Requirements**: .NET 8 SDK, x86 platform target (RSLinx is 32-bit)

```bash
# Development build
dotnet build -c Release

# Self-contained publish (for machines without .NET SDK)
dotnet publish -r win-x86 --self-contained -c Release
```

Output: `bin/Release/net8.0/RSLinxViewer.dll` (or self-contained in `publish/`)

RSLinxHook.dll must be placed next to the executable, or use `--dll` to specify its path.

## Usage

```
RSLinxViewer.exe --driver NAME [IP...] [--driver NAME2 [IP...]] [--dll PATH] [--debug-xml] [--monitor] [--log-dir DIR]
```

### Examples

```bash
# Browse driver SUFF2 with one IP
RSLinxViewer.exe --driver SUFF2 10.13.30.68

# Monitor mode (no driver changes)
RSLinxViewer.exe --driver SUFF2 --monitor

# Multiple drivers
RSLinxViewer.exe --driver Test 10.39.31.200 --driver SUFF2 10.13.30.68

# Custom DLL path
RSLinxViewer.exe --driver Test --dll C:\path\to\RSLinxHook.dll
```

### Flags

| Flag | Description |
|------|-------------|
| `--driver NAME [IP...]` | Driver name, optionally followed by IP addresses |
| `--dll PATH` | Path to RSLinxHook.dll (default: next to exe) |
| `--monitor` | Browse existing topology without modifying drivers |
| `--debug-xml` | Enable debug XML topology snapshots in hook |
| `--log-dir DIR` | Log directory (default: `C:\temp`) |

## TUI Controls

| Key | Action |
|-----|--------|
| Up/Down arrows | Scroll topology tree |
| PgUp/PgDn | Scroll 10 lines |
| Home | Scroll to top |
| Ctrl+C | Stop hook and exit |

## How It Works

1. **Inject** — Finds RSLinx PID, ejects any old DLL, injects RSLinxHook.dll via `CreateRemoteThread(LoadLibraryW)`
2. **Pipe** — Creates `\\.\pipe\RSLinxHook` named pipe server, waits for hook to connect
3. **Config** — Sends driver names, IPs, and mode over the pipe
4. **Display** — Reads pipe messages in background:
   - `L|` log lines shown in the log panel
   - `S|` status updates shown in the footer
   - `X|BEGIN...X|END` topology XML parsed into a scrollable tree
   - `D|` signals completion
5. **Cleanup** — Sends `STOP` over pipe, waits for hook to finish, ejects DLL

## Architecture

```
RSLinxViewer/
├── Program.cs          — CLI parsing, DLL injection, TUI live display loop
├── Injector.cs         — P/Invoke DLL injection/ejection (LoadLibraryW/FreeLibrary)
├── PipeProtocol.cs     — Bidirectional named pipe server (SendConfig, SendStop, ReadLoop)
├── TopologyParser.cs   — XML topology → Spectre.Console Tree widget
├── Viewport.cs         — Vertical scroll viewport for tree rendering
└── RSLinxViewer.csproj — .NET 8, x86, Spectre.Console dependency
```

### PipeProtocol

Synchronous byte-mode pipe (`PipeOptions.None`) with thread-pool async wrappers. Key lesson: `PipeOptions.Asynchronous` causes `Write()` / `WriteAsync()` to fail or hang when the client opened without `FILE_FLAG_OVERLAPPED`.

### TopologyParser

Parses Rockwell topology XML (`<topology><tree><device>...`) into a Spectre.Console `Tree`. Ethernet devices show as IP addresses, backplane modules show as slot numbers in a compact 2-column grid.

### Injector

Direct port of `RSLinxBrowse/main.cpp` injection logic to C# P/Invoke. Handles `SeDebugPrivilege` escalation for service-mode RSLinx processes.

## Dependencies

- [Spectre.Console](https://spectreconsole.net/) 0.49.x — TUI rendering
- .NET 8.0 — Runtime
