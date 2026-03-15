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
# Browse driver Test with one IP
RSLinxViewer.exe --driver Test --ip 192.168.1.55

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
| B | Trigger re-browse (in-band, no reconnect) |
| Ctrl+C | Send STOP and exit |

## How It Works

1. **Smart injection** — Probes `\\.\pipe\RSLinxHook` (500 ms timeout). If the hook is already running, connects directly without re-injecting. If not found, injects RSLinxHook.dll via `CreateRemoteThread(LoadLibraryW)` and waits up to 10 s for the pipe server to appear.
2. **Pipe** — Connects to `\\.\pipe\RSLinxHook` as a named pipe **client** (hook is the server)
3. **Config** — Sends driver names, IPs, and mode over the pipe (`C|MODE`, `C|DRIVER`, `C|IP`, `C|END`)
4. **Display** — Reads pipe messages in background:
   - `L|` log lines shown in the log panel
   - `S|` status updates shown in the footer
   - `X|BEGIN...X|END` topology XML parsed into a scrollable tree
   - `D|` signals browse complete; pipe stays open for re-browse and queries
5. **Re-browse** — Pressing `B` sends `B|` over the existing connection; hook re-runs all browse phases in-place and sends a new `X|BEGIN...X|END` block followed by `D|`. No disconnect/reconnect.
6. **Cleanup** — Ctrl+C sends `STOP` over pipe and closes the handle. The hook remains injected and loops back to accept the next client.

## Architecture

```
RSLinxViewer/
├── Program.cs          — CLI parsing, DLL injection, TUI live display loop
├── Injector.cs         — P/Invoke DLL injection (LoadLibraryW via CreateRemoteThread)
├── PipeProtocol.cs     — Bidirectional named pipe client (SendConfig, SendBrowse, SendStop, ReadLoop)
├── TopologyParser.cs   — XML topology → Spectre.Console Tree widget
├── Viewport.cs         — Vertical scroll viewport for tree rendering
└── RSLinxViewer.csproj — .NET 8, x86, Spectre.Console dependency
```

### PipeProtocol

`NamedPipeClientStream` with `PipeOptions.Asynchronous` (`FILE_FLAG_OVERLAPPED`). This is required because the background `ReadLoopAsync` thread and the main thread both perform I/O on the same handle concurrently. Without `PipeOptions.Asynchronous`, Windows serializes synchronous `ReadFile`/`WriteFile` calls on the same handle, causing a deadlock when pressing 'B' to send `B|` while the read loop is blocked on `ReadLine`.

### TopologyParser

Parses Rockwell topology XML (`<topology><tree><device>...`) into a Spectre.Console `Tree`. Ethernet devices show as IP addresses, backplane modules show as slot numbers in a compact 2-column grid.

### Injector

Direct port of `RSLinxBrowse/main.cpp` injection logic to C# P/Invoke. Handles `SeDebugPrivilege` escalation for service-mode RSLinx processes. EjectDLL is defined but not called — the hook is intentionally left resident after viewer exit.

## Dependencies

- [Spectre.Console](https://spectreconsole.net/) 0.49.x — TUI rendering
- .NET 8.0 — Runtime
