# raFTanalysis

A collection of Windows tools for Rockwell Automation reverse engineering and analysis. Includes EtherNet/IP device discovery via RSLinx Classic, MER file signature utilities, and COM interface exploration tools.

## Projects

### RSLinx Topology Discovery

Three projects work together to discover and display EtherNet/IP devices through RSLinx Classic's internal COM interfaces.

| Project | Description |
|---------|-------------|
| **RSLinxHook** | C++ DLL injected into `rslinx.exe` for in-process COM topology access. Performs a six-phase discovery sequence including device connection, STA-thread browse triggering, topology polling, and backplane enumeration. |
| **RSLinxBrowse** | C++ CLI that injects RSLinxHook into RSLinx, sends configuration over a named pipe, and writes topology XML snapshots and device discovery results to disk. |
| **RSLinxViewer** | .NET 8 TUI application that injects the same RSLinxHook DLL and displays the topology tree in real-time using Spectre.Console. Supports interactive scrolling and continuous monitoring. |

#### Architecture

```
┌─────────────────┐      ┌─────────────────┐
│  RSLinxBrowse    │      │  RSLinxViewer    │
│  (C++ CLI)      │      │  (.NET 8 TUI)   │
└────────┬────────┘      └────────┬────────┘
         │  CreateRemoteThread     │  CreateRemoteThread
         │  (LoadLibraryW)         │  (LoadLibraryW)
         └───────────┬────────────┘
                     ▼
         ┌───────────────────────┐
         │      rslinx.exe       │
         │  ┌─────────────────┐  │
         │  │  RSLinxHook.dll  │  │
         │  │  (injected)     │  │
         │  └────────┬────────┘  │
         │           │ COM/STA   │
         │  ┌────────▼────────┐  │
         │  │  ENGINE.DLL     │  │
         │  │  (topology)     │  │
         │  └─────────────────┘  │
         └───────────────────────┘
                     ▲
                     │ Named Pipe
                     │ \\.\pipe\RSLinxHook
                     ▼
              Launcher process
              (Browse or Viewer)
```

**Why DLL injection?** RSLinx Classic's COM objects are STA-only. Cross-process COM calls fail with `E_NOINTERFACE`. Running inside `rslinx.exe` gives direct access to the topology engine and the main thread's message pump required for CIP I/O callbacks.

### MER File Utilities

| Project | Description |
|---------|-------------|
| **RockwellRsvCRCCreateSigniture** | C++ utility that instantiates the `IRsvCRC` COM interface to sign MER files via the `SignMerFile` vtable method. |
| **RockwellRsvCRCValidateSigniture** | C++ utility that uses the same `IRsvCRC` COM interface to validate existing MER file signatures. |
| **RockwellFTArchiveDirTesting** | C++ utility that instantiates the `IFTArchiveDir` COM interface to extract and inspect MER archive directory structures. |

## Build Requirements

| Requirement | Version |
|-------------|---------|
| Visual Studio | 2019+ (v142 platform toolset) |
| .NET SDK | 8.0 (RSLinxViewer only) |
| Platform | Windows x86 (32-bit) — RSLinx Classic is 32-bit |
| C++ Standard | C++17 |

### Building

**C++ projects** (RSLinxBrowse, RSLinxHook, and the MER utilities):

```
MSBuild.exe raFTMEanalysis.sln -p:Configuration=Release -p:Platform=Win32
```

**RSLinxViewer** (.NET 8):

```
cd RSLinxViewer
dotnet publish -r win-x86 --self-contained -c Release
```

> RSLinxHook.dll must be placed alongside whichever launcher executable you use (RSLinxBrowse.exe or RSLinxViewer.exe).

## Technology Stack

- **COM Automation** — All projects interact with Rockwell's COM interfaces (STA threading, IDispatch, vtable inspection)
- **DLL Injection** — `CreateRemoteThread` + `LoadLibraryW` for in-process access to RSLinx
- **Named Pipes** — Bidirectional IPC between launcher and injected DLL (`\\.\pipe\RSLinxHook`)
- **WH_GETMESSAGE Hook** — RSLinxHook installs a Windows message hook to execute COM calls on RSLinx's main STA thread
- **Spectre.Console** — Terminal UI rendering for RSLinxViewer
