#pragma once

// ============================================================
// RSLinxHook shared includes, forward declarations, and types
// ============================================================

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

// Forward declarations
enum class HookMode;
struct DriverEntry;
struct HookConfig;
struct BusInfo;
struct DeviceInfo;
struct TopologyCounts;
struct ConnectionPointInfo;
struct EnumeratorInfo;
class DualEventSink;

// Function pointer type for main-STA execution
typedef HRESULT (*MainSTAFunc)();

// Constants
#define HOOK_MAGIC_WPARAM 0xDEAD7F00
#define SUBCLASS_MSG (WM_USER + 0x7F00)

// Global stop flag and worker thread handle (defined in DllMain.cpp)
extern HANDLE g_hWorkerThread;
extern volatile bool g_shouldStop;
