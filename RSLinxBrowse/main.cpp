/**
 * RSLinxBrowse - Programmatic RSLinx Device Discovery
 *
 * Injects RSLinxHook.dll into rslinx.exe to perform in-process COM operations
 * against the RSLinx topology engine. External COM is not viable because RSLinx
 * COM objects are in-process STA only (E_NOINTERFACE cross-process).
 *
 * Modes:
 *   Default:    Create/update driver, inject DLL, trigger CDriver Stop/Start
 *               hot-load cycle, browse full topology (equivalent to --inject)
 *   --monitor:  Inject DLL and browse existing driver topology without
 *               creating or modifying drivers
 *
 * REQUIREMENTS:
 * - Must compile for Win32 (x86) - RSLinx is 32-bit only
 * - RSLinx Classic must be installed and running
 * - For --monitor mode, the driver must already exist in RSLinx
 */

#include "DriverConfig.h"
#include "TopologyBrowser.h"
#include <iostream>
#include <iomanip>
#include <fstream>
#include <tlhelp32.h>
#include <psapi.h>
#pragma comment(lib, "psapi.lib")

// Per-driver specification from CLI
struct DriverSpec {
    std::wstring name;
    std::vector<std::wstring> ips;
};

// ============================================================
// Named Pipe Server — bidirectional IPC with RSLinxHook.dll
// ============================================================

static HANDLE g_hPipeServer = INVALID_HANDLE_VALUE;

static bool CreatePipeServer()
{
    g_hPipeServer = CreateNamedPipeW(
        L"\\\\.\\pipe\\RSLinxHook",
        PIPE_ACCESS_DUPLEX,
        PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT,
        1, 4096, 4096, 0, NULL);
    return g_hPipeServer != INVALID_HANDLE_VALUE;
}

static DWORD WINAPI ConnectPipeThread(LPVOID)
{
    ConnectNamedPipe(g_hPipeServer, NULL);
    return 0;
}

static bool WaitForPipeConnection(DWORD timeoutMs)
{
    HANDLE hThread = CreateThread(NULL, 0, ConnectPipeThread, NULL, 0, NULL);
    if (!hThread) return false;
    DWORD result = WaitForSingleObject(hThread, timeoutMs);
    CloseHandle(hThread);
    if (result == WAIT_OBJECT_0) return true;
    // Timeout — cancel by closing pipe (unblocks ConnectNamedPipe)
    CloseHandle(g_hPipeServer);
    g_hPipeServer = INVALID_HANDLE_VALUE;
    return false;
}

static void PipeSendLine(const std::string& line)
{
    std::string data = line + "\n";
    DWORD written;
    WriteFile(g_hPipeServer, data.c_str(), (DWORD)data.size(), &written, NULL);
}

static void PipeSendConfig(const std::vector<DriverSpec>& drivers,
    const std::vector<bool>& needsHotLoad,
    const std::wstring& logDir, bool debugXml, bool monitorMode, bool probeDispids = false)
{
    PipeSendLine(monitorMode ? "C|MODE=monitor" : "C|MODE=inject");
    if (logDir != L"C:\\temp") {
        char utf8[512];
        WideCharToMultiByte(CP_UTF8, 0, logDir.c_str(), -1, utf8, sizeof(utf8), NULL, NULL);
        PipeSendLine(std::string("C|LOGDIR=") + utf8);
    }
    if (debugXml) PipeSendLine("C|DEBUGXML=1");
    if (probeDispids) PipeSendLine("C|PROBE=1");

    for (size_t di = 0; di < drivers.size(); di++) {
        char utf8[256];
        WideCharToMultiByte(CP_UTF8, 0, drivers[di].name.c_str(), -1, utf8, sizeof(utf8), NULL, NULL);
        PipeSendLine(std::string("C|DRIVER=") + utf8);
        if (di < needsHotLoad.size() && needsHotLoad[di])
            PipeSendLine("C|NEWDRIVER=1");
        for (const auto& ip : drivers[di].ips) {
            char ipUtf8[256];
            WideCharToMultiByte(CP_UTF8, 0, ip.c_str(), -1, ipUtf8, sizeof(ipUtf8), NULL, NULL);
            PipeSendLine(std::string("C|IP=") + ipUtf8);
        }
    }
    PipeSendLine("C|END");
}

static void PipeSendStop()
{
    if (g_hPipeServer == INVALID_HANDLE_VALUE) return;
    DWORD written;
    WriteFile(g_hPipeServer, "STOP\n", 5, &written, NULL);
}

static void PipeClose()
{
    if (g_hPipeServer != INVALID_HANDLE_VALUE) {
        DisconnectNamedPipe(g_hPipeServer);
        CloseHandle(g_hPipeServer);
        g_hPipeServer = INVALID_HANDLE_VALUE;
    }
}

// Read pipe messages, print L| log lines to console in real-time.
// Returns true when D| received or pipe disconnects, false on timeout.
static bool PipeReadLoop(DWORD timeoutMs)
{
    // Use WriteConsoleW for Unicode-safe output (wcout breaks on em dashes etc.)
    HANDLE hOut = GetStdHandle(STD_OUTPUT_HANDLE);

    std::string accumulated;
    char buf[4096];
    DWORD startTick = GetTickCount();

    while (true) {
        DWORD elapsed = GetTickCount() - startTick;
        if (elapsed >= timeoutMs) {
            std::wcerr << L"[FAIL] Timed out waiting for hook (" << (timeoutMs / 1000) << L"s)" << std::endl;
            return false;
        }

        DWORD bytesAvail = 0;
        if (!PeekNamedPipe(g_hPipeServer, NULL, 0, NULL, &bytesAvail, NULL)) {
            return true; // pipe broken = hook disconnected = done
        }

        if (bytesAvail == 0) {
            Sleep(100);
            continue;
        }

        DWORD toRead = bytesAvail < (DWORD)(sizeof(buf) - 1) ? bytesAvail : (DWORD)(sizeof(buf) - 1);
        DWORD bytesRead = 0;
        if (!ReadFile(g_hPipeServer, buf, toRead, &bytesRead, NULL) || bytesRead == 0)
            return true; // pipe closed

        buf[bytesRead] = '\0';
        accumulated += buf;

        size_t pos;
        while ((pos = accumulated.find('\n')) != std::string::npos) {
            std::string line = accumulated.substr(0, pos);
            accumulated.erase(0, pos + 1);
            if (!line.empty() && line.back() == '\r') line.pop_back();

            if (line.length() >= 2 && line[0] == 'L' && line[1] == '|') {
                std::string msg = line.substr(2);
                // Convert UTF-8 to wide and output via WriteConsoleW (handles all Unicode)
                int wlen = MultiByteToWideChar(CP_UTF8, 0, msg.c_str(), -1, NULL, 0);
                if (wlen > 0) {
                    std::wstring wmsg(wlen - 1, 0);
                    MultiByteToWideChar(CP_UTF8, 0, msg.c_str(), -1, &wmsg[0], wlen);
                    DWORD written;
                    WriteConsoleW(hOut, L"  ", 2, &written, NULL);
                    WriteConsoleW(hOut, wmsg.c_str(), (DWORD)wmsg.size(), &written, NULL);
                    WriteConsoleW(hOut, L"\n", 1, &written, NULL);
                }
            }
            else if (line.length() >= 2 && line[0] == 'S' && line[1] == '|') {
                // Status update — could print periodically if desired
            }
            else if (line.length() >= 2 && line[0] == 'D' && line[1] == '|') {
                return true; // done
            }
            // X|BEGIN..X|END blocks — skip in CLI mode
        }
    }
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

void PrintHeader(const std::wstring& title)
{
    std::wcout << std::endl;
    std::wcout << L"============================================================" << std::endl;
    std::wcout << title << std::endl;
    std::wcout << L"============================================================" << std::endl;
}

// ============================================================
// DLL Injection into rslinx.exe
// ============================================================

static DWORD FindProcessByName(const wchar_t* processName)
{
    HANDLE hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnap == INVALID_HANDLE_VALUE)
        return 0;

    PROCESSENTRY32W pe;
    pe.dwSize = sizeof(pe);

    DWORD pid = 0;
    if (Process32FirstW(hSnap, &pe))
    {
        do
        {
            if (_wcsicmp(pe.szExeFile, processName) == 0)
            {
                pid = pe.th32ProcessID;
                break;
            }
        } while (Process32NextW(hSnap, &pe));
    }

    CloseHandle(hSnap);
    return pid;
}

// Enable SeDebugPrivilege in current process token.
// Required for opening service processes (Session 0 / SYSTEM).
static bool EnableDebugPrivilege()
{
    HANDLE hToken = NULL;
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken))
        return false;

    TOKEN_PRIVILEGES tp;
    tp.PrivilegeCount = 1;
    tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;

    if (!LookupPrivilegeValueW(NULL, SE_DEBUG_NAME, &tp.Privileges[0].Luid))
    {
        CloseHandle(hToken);
        return false;
    }

    BOOL ok = AdjustTokenPrivileges(hToken, FALSE, &tp, sizeof(tp), NULL, NULL);
    DWORD err = GetLastError();
    CloseHandle(hToken);

    // AdjustTokenPrivileges returns TRUE even if privilege not assigned,
    // but sets ERROR_NOT_ALL_ASSIGNED in that case.
    return ok && err != ERROR_NOT_ALL_ASSIGNED;
}

static bool InjectDLL(DWORD pid, const std::wstring& dllPath)
{
    std::wcout << L"[INFO] Opening process PID " << pid << std::endl;

    HANDLE hProcess = OpenProcess(
        PROCESS_CREATE_THREAD | PROCESS_VM_OPERATION | PROCESS_VM_WRITE | PROCESS_VM_READ | PROCESS_QUERY_INFORMATION,
        FALSE, pid);
    if (!hProcess && GetLastError() == ERROR_ACCESS_DENIED)
    {
        std::wcout << L"[INFO] Access denied — enabling SeDebugPrivilege (service mode)..." << std::endl;
        if (EnableDebugPrivilege())
        {
            std::wcout << L"[OK] SeDebugPrivilege enabled" << std::endl;
            hProcess = OpenProcess(
                PROCESS_CREATE_THREAD | PROCESS_VM_OPERATION | PROCESS_VM_WRITE | PROCESS_VM_READ | PROCESS_QUERY_INFORMATION,
                FALSE, pid);
        }
        else
        {
            std::wcerr << L"[FAIL] Could not enable SeDebugPrivilege — run as Administrator" << std::endl;
            return false;
        }
    }
    if (!hProcess)
    {
        std::wcerr << L"[FAIL] OpenProcess failed: " << GetLastError() << std::endl;
        return false;
    }

    // Allocate memory in target process for DLL path
    size_t pathBytes = (dllPath.length() + 1) * sizeof(wchar_t);
    void* pRemotePath = VirtualAllocEx(hProcess, NULL, pathBytes, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
    if (!pRemotePath)
    {
        std::wcerr << L"[FAIL] VirtualAllocEx failed: " << GetLastError() << std::endl;
        CloseHandle(hProcess);
        return false;
    }

    // Write DLL path to target process
    if (!WriteProcessMemory(hProcess, pRemotePath, dllPath.c_str(), pathBytes, NULL))
    {
        std::wcerr << L"[FAIL] WriteProcessMemory failed: " << GetLastError() << std::endl;
        VirtualFreeEx(hProcess, pRemotePath, 0, MEM_RELEASE);
        CloseHandle(hProcess);
        return false;
    }

    // Get LoadLibraryW address (same in all processes due to kernel32 base address)
    HMODULE hKernel32 = GetModuleHandleW(L"kernel32.dll");
    FARPROC pLoadLibraryW = GetProcAddress(hKernel32, "LoadLibraryW");
    if (!pLoadLibraryW)
    {
        std::wcerr << L"[FAIL] GetProcAddress(LoadLibraryW) failed" << std::endl;
        VirtualFreeEx(hProcess, pRemotePath, 0, MEM_RELEASE);
        CloseHandle(hProcess);
        return false;
    }

    // Create remote thread to call LoadLibraryW(dllPath)
    std::wcout << L"[INFO] Creating remote thread in rslinx.exe..." << std::endl;
    HANDLE hThread = CreateRemoteThread(hProcess, NULL, 0,
        (LPTHREAD_START_ROUTINE)pLoadLibraryW, pRemotePath, 0, NULL);
    if (!hThread)
    {
        std::wcerr << L"[FAIL] CreateRemoteThread failed: " << GetLastError() << std::endl;
        VirtualFreeEx(hProcess, pRemotePath, 0, MEM_RELEASE);
        CloseHandle(hProcess);
        return false;
    }

    std::wcout << L"[OK] DLL injected, waiting for completion..." << std::endl;

    // Wait for the LoadLibrary call to complete (DLL loaded, DllMain called)
    WaitForSingleObject(hThread, 10000);

    // Get exit code (HMODULE of loaded DLL, or 0 on failure)
    DWORD exitCode = 0;
    GetExitCodeThread(hThread, &exitCode);
    std::wcout << L"[INFO] LoadLibrary returned: 0x" << std::hex << exitCode << std::dec << std::endl;

    CloseHandle(hThread);
    VirtualFreeEx(hProcess, pRemotePath, 0, MEM_RELEASE);
    CloseHandle(hProcess);

    return (exitCode != 0);
}

// Eject a previously injected DLL by calling FreeLibrary in the remote process
static bool EjectDLL(DWORD pid, const std::wstring& dllPath)
{
    HANDLE hProcess = OpenProcess(
        PROCESS_CREATE_THREAD | PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_QUERY_INFORMATION,
        FALSE, pid);
    if (!hProcess && GetLastError() == ERROR_ACCESS_DENIED)
    {
        EnableDebugPrivilege();
        hProcess = OpenProcess(
            PROCESS_CREATE_THREAD | PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_QUERY_INFORMATION,
            FALSE, pid);
    }
    if (!hProcess) return false;

    // Find the DLL's HMODULE in the remote process
    HMODULE hMods[1024];
    DWORD cbNeeded = 0;
    if (!EnumProcessModules(hProcess, hMods, sizeof(hMods), &cbNeeded))
    {
        CloseHandle(hProcess);
        return false;
    }

    HMODULE hTargetDLL = NULL;
    for (DWORD i = 0; i < cbNeeded / sizeof(HMODULE); i++)
    {
        wchar_t modName[MAX_PATH];
        if (GetModuleFileNameExW(hProcess, hMods[i], modName, MAX_PATH))
        {
            if (_wcsicmp(modName, dllPath.c_str()) == 0)
            {
                hTargetDLL = hMods[i];
                break;
            }
        }
    }

    if (!hTargetDLL)
    {
        CloseHandle(hProcess);
        return false;  // DLL not loaded
    }

    // Call FreeLibrary in the remote process
    HMODULE hKernel32 = GetModuleHandleW(L"kernel32.dll");
    FARPROC pFreeLibrary = GetProcAddress(hKernel32, "FreeLibrary");
    HANDLE hThread = CreateRemoteThread(hProcess, NULL, 0,
        (LPTHREAD_START_ROUTINE)pFreeLibrary, hTargetDLL, 0, NULL);
    if (hThread)
    {
        WaitForSingleObject(hThread, 5000);
        CloseHandle(hThread);
    }

    CloseHandle(hProcess);
    std::wcout << L"[OK] Ejected previous DLL instance" << std::endl;
    return true;
}

// ============================================================
// Auto-create driver registry entry if missing
// ============================================================

static bool DriverExistsInRegistry(const std::wstring& driverName)
{
    std::wstring basePath = L"SOFTWARE\\WOW6432Node\\Rockwell Software\\RSLinx\\Drivers\\AB_ETH";
    HKEY hBaseKey;
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, basePath.c_str(), 0, KEY_READ, &hBaseKey) != ERROR_SUCCESS)
        return false;

    wchar_t subKeyName[256];
    DWORD subKeyNameLen;
    bool found = false;

    for (DWORD i = 0; ; i++)
    {
        subKeyNameLen = 256;
        if (RegEnumKeyExW(hBaseKey, i, subKeyName, &subKeyNameLen, NULL, NULL, NULL, NULL) != ERROR_SUCCESS)
            break;

        HKEY hDriverKey;
        if (RegOpenKeyExW(hBaseKey, subKeyName, 0, KEY_READ, &hDriverKey) != ERROR_SUCCESS)
            continue;

        wchar_t nameValue[256] = {};
        DWORD nameValueLen = sizeof(nameValue);
        DWORD nameType;
        if (RegQueryValueExW(hDriverKey, L"Name", NULL, &nameType, (BYTE*)nameValue, &nameValueLen) == ERROR_SUCCESS
            && nameType == REG_SZ && _wcsicmp(nameValue, driverName.c_str()) == 0)
        {
            found = true;
        }
        RegCloseKey(hDriverKey);
        if (found) break;
    }

    RegCloseKey(hBaseKey);
    return found;
}

// Returns: 0=failure, 1=existed no changes, 2=new driver created, 3=existing driver IPs added
static int CreateDriverRegistry(const std::wstring& driverName, const std::vector<std::wstring>& ips)
{
    std::wcout << L"[INFO] Checking if driver '" << driverName << L"' exists in registry..." << std::endl;

    if (DriverExistsInRegistry(driverName))
    {
        // Driver exists — check if we need to add new IPs to its Node Table
        std::vector<std::wstring> existingIPs = DriverConfig::ReadNodeTable(driverName);

        // Find IPs that are not yet in the Node Table
        const auto& desiredIPs = ips;

        // Find IPs that are not yet in the Node Table
        std::vector<std::wstring> newIPs;
        for (const auto& ip : desiredIPs)
        {
            bool exists = false;
            for (const auto& e : existingIPs) if (e == ip) { exists = true; break; }
            if (!exists) newIPs.push_back(ip);
        }

        if (newIPs.empty())
        {
            std::wcout << L"[OK] Driver '" << driverName << L"' already exists, all IPs present" << std::endl;
            return 1;
        }

        // Find the driver's registry path so we can open its Node Table for writing
        std::wstring basePath = L"SOFTWARE\\WOW6432Node\\Rockwell Software\\RSLinx\\Drivers\\AB_ETH";
        HKEY hBaseKey;
        if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, basePath.c_str(), 0, KEY_READ, &hBaseKey) != ERROR_SUCCESS)
            return 0;

        std::wstring foundDriverPath;
        wchar_t subKeyName[256];
        DWORD subKeyNameLen;
        for (DWORD i = 0; ; i++)
        {
            subKeyNameLen = 256;
            if (RegEnumKeyExW(hBaseKey, i, subKeyName, &subKeyNameLen, NULL, NULL, NULL, NULL) != ERROR_SUCCESS)
                break;
            HKEY hDriverKey;
            if (RegOpenKeyExW(hBaseKey, subKeyName, 0, KEY_READ, &hDriverKey) != ERROR_SUCCESS)
                continue;
            wchar_t nameValue[256] = {};
            DWORD nameValueLen = sizeof(nameValue);
            DWORD nameType;
            if (RegQueryValueExW(hDriverKey, L"Name", NULL, &nameType, (BYTE*)nameValue, &nameValueLen) == ERROR_SUCCESS
                && nameType == REG_SZ && _wcsicmp(nameValue, driverName.c_str()) == 0)
            {
                foundDriverPath = basePath + L"\\" + subKeyName;
            }
            RegCloseKey(hDriverKey);
            if (!foundDriverPath.empty()) break;
        }
        RegCloseKey(hBaseKey);

        if (foundDriverPath.empty()) return 0;

        // Open Node Table for writing and append new IPs
        std::wstring nodeTablePath = foundDriverPath + L"\\Node Table";
        HKEY hNodeTable;
        LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, nodeTablePath.c_str(), 0, KEY_ALL_ACCESS, &hNodeTable);
        if (result != ERROR_SUCCESS)
        {
            result = RegCreateKeyExW(HKEY_LOCAL_MACHINE, nodeTablePath.c_str(), 0, NULL,
                                      REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hNodeTable, NULL);
            if (result != ERROR_SUCCESS) return 0;
        }

        // Find next available index (existing count)
        int nextIndex = (int)existingIPs.size();
        for (const auto& ip : newIPs)
        {
            wchar_t indexStr[16];
            swprintf_s(indexStr, L"%d", nextIndex++);
            RegSetValueExW(hNodeTable, indexStr, 0, REG_SZ,
                           (const BYTE*)ip.c_str(), (DWORD)((ip.length() + 1) * sizeof(wchar_t)));
            std::wcout << L"[OK] Added IP " << ip << L" to Node Table (index " << (nextIndex - 1) << L")" << std::endl;
        }
        RegCloseKey(hNodeTable);

        std::wcout << L"[OK] Driver '" << driverName << L"' updated: added " << newIPs.size() << L" new IP(s)" << std::endl;
        return 3;
    }

    std::wcout << L"[INFO] Driver '" << driverName << L"' not found, creating registry entry..." << std::endl;

    std::wstring basePath = L"SOFTWARE\\WOW6432Node\\Rockwell Software\\RSLinx\\Drivers\\AB_ETH";

    // Find next available AB_ETH-N index
    HKEY hBaseKey;
    LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, basePath.c_str(), 0, KEY_READ, &hBaseKey);
    if (result != ERROR_SUCCESS)
    {
        // AB_ETH key doesn't exist — create it
        result = RegCreateKeyExW(HKEY_LOCAL_MACHINE, basePath.c_str(), 0, NULL,
                                  REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hBaseKey, NULL);
        if (result != ERROR_SUCCESS)
        {
            std::wcerr << L"[FAIL] Cannot create AB_ETH registry key: " << result << std::endl;
            return 0;
        }
    }

    int nextIndex = 1;
    wchar_t subKeyName[256];
    DWORD subKeyNameLen;
    for (DWORD i = 0; ; i++)
    {
        subKeyNameLen = 256;
        if (RegEnumKeyExW(hBaseKey, i, subKeyName, &subKeyNameLen, NULL, NULL, NULL, NULL) != ERROR_SUCCESS)
            break;
        // Parse "AB_ETH-N" to find highest N
        wchar_t* dash = wcsrchr(subKeyName, L'-');
        if (dash)
        {
            int n = _wtoi(dash + 1);
            if (n >= nextIndex)
                nextIndex = n + 1;
        }
    }
    RegCloseKey(hBaseKey);

    // Create AB_ETH-N subkey
    wchar_t driverSubKey[64];
    swprintf_s(driverSubKey, L"AB_ETH-%d", nextIndex);
    std::wstring driverPath = basePath + L"\\" + driverSubKey;

    std::wcout << L"[INFO] Creating registry key: " << driverPath << std::endl;

    HKEY hDriverKey;
    result = RegCreateKeyExW(HKEY_LOCAL_MACHINE, driverPath.c_str(), 0, NULL,
                              REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hDriverKey, NULL);
    if (result != ERROR_SUCCESS)
    {
        std::wcerr << L"[FAIL] Cannot create driver registry key: " << result << std::endl;
        return 0;
    }

    // Write driver configuration values (matching existing "Test" driver)
    RegSetValueExW(hDriverKey, L"Name", 0, REG_SZ,
                   (const BYTE*)driverName.c_str(), (DWORD)((driverName.length() + 1) * sizeof(wchar_t)));

    DWORD startup = 0;
    RegSetValueExW(hDriverKey, L"Startup", 0, REG_DWORD, (const BYTE*)&startup, sizeof(DWORD));

    DWORD station = 0x3f;
    RegSetValueExW(hDriverKey, L"Station", 0, REG_DWORD, (const BYTE*)&station, sizeof(DWORD));

    DWORD inactivityTimeout = 0x1e;
    RegSetValueExW(hDriverKey, L"Inactivity Timeout", 0, REG_DWORD, (const BYTE*)&inactivityTimeout, sizeof(DWORD));

    DWORD pingTimeout = 0x6;
    RegSetValueExW(hDriverKey, L"Ping Timeout", 0, REG_DWORD, (const BYTE*)&pingTimeout, sizeof(DWORD));

    RegCloseKey(hDriverKey);

    // Create Node Table subkey with IPs
    std::wstring nodeTablePath = driverPath + L"\\Node Table";
    HKEY hNodeTable;
    result = RegCreateKeyExW(HKEY_LOCAL_MACHINE, nodeTablePath.c_str(), 0, NULL,
                              REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hNodeTable, NULL);
    if (result != ERROR_SUCCESS)
    {
        std::wcerr << L"[FAIL] Cannot create Node Table key: " << result << std::endl;
        return 0;
    }

    // Write all IPs to Node Table
    int ipIndex = 0;
    for (const auto& ip : ips)
    {
        wchar_t indexStr[16];
        swprintf_s(indexStr, L"%d", ipIndex++);
        RegSetValueExW(hNodeTable, indexStr, 0, REG_SZ,
                       (const BYTE*)ip.c_str(), (DWORD)((ip.length() + 1) * sizeof(wchar_t)));
    }

    RegCloseKey(hNodeTable);

    std::wcout << L"[OK] Created driver '" << driverName << L"' as " << driverSubKey
               << L" with " << ipIndex << L" IP(s) in Node Table" << std::endl;
    return 2;
}

static int RunInjectMode(const std::vector<DriverSpec>& drivers, const std::wstring& logDir, bool debugXml, bool probeDispids)
{
    PrintHeader(L"RSLinx Hook Injection Mode");

    // Ensure log directory exists
    CreateDirectoryW(logDir.c_str(), NULL);

    // Step 0: Ensure each driver exists in registry with all IPs
    std::vector<bool> driverNeedsHotLoad(drivers.size(), false);
    bool anyNeedsHotLoad = false;
    for (size_t di = 0; di < drivers.size(); di++)
    {
        int createResult = CreateDriverRegistry(drivers[di].name, drivers[di].ips);
        if (createResult == 0)
        {
            std::wcerr << L"[WARN] Could not create driver '" << drivers[di].name
                       << L"' registry entry (may need admin privileges)" << std::endl;
        }
        if (createResult == 2 || createResult == 3)
        {
            driverNeedsHotLoad[di] = true;
            anyNeedsHotLoad = true;
        }
    }
    if (anyNeedsHotLoad)
        std::wcout << L"[INFO] Registry modified -- Hook DLL will hot-load via ENGINE.DLL (CDriver Stop/Start)" << std::endl;

    // Step 1: Find rslinx.exe
    DWORD rslinxPid = FindProcessByName(L"RSLinx.exe");
    if (rslinxPid == 0) rslinxPid = FindProcessByName(L"RSLINX.EXE");
    if (rslinxPid == 0) rslinxPid = FindProcessByName(L"rslinx.exe");
    if (rslinxPid == 0)
    {
        std::wcerr << L"[FAIL] RSLinx.exe not found. Is it running?" << std::endl;
        return 1;
    }
    std::wcout << L"[OK] Found RSLinx.exe PID: " << rslinxPid << std::endl;

    // Step 2: Create pipe server BEFORE injection
    std::wcout << L"[INFO] Creating named pipe server..." << std::endl;
    if (!CreatePipeServer())
    {
        std::wcerr << L"[FAIL] Cannot create named pipe: " << GetLastError() << std::endl;
        return 1;
    }

    // Step 3: Delete old results/log files
    DeleteFileW(LogPath(logDir, L"hook_results.txt").c_str());
    DeleteFileW(LogPath(logDir, L"hook_log.txt").c_str());

    // Step 4: Find RSLinxHook.dll path (same directory as our exe)
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    std::wstring dllPath(exePath);
    size_t lastSlash = dllPath.rfind(L'\\');
    if (lastSlash != std::wstring::npos)
        dllPath = dllPath.substr(0, lastSlash + 1);
    dllPath += L"RSLinxHook.dll";

    DWORD attrs = GetFileAttributesW(dllPath.c_str());
    if (attrs == INVALID_FILE_ATTRIBUTES)
    {
        std::wcerr << L"[FAIL] RSLinxHook.dll not found at: " << dllPath << std::endl;
        PipeClose();
        return 1;
    }
    std::wcout << L"[OK] DLL path: " << dllPath << std::endl;

    // Log driver info
    for (size_t di = 0; di < drivers.size(); di++)
    {
        std::wcout << L"  Driver '" << drivers[di].name << L"': " << drivers[di].ips.size() << L" IP(s)";
        if (driverNeedsHotLoad[di]) std::wcout << L" [NEWDRIVER]";
        std::wcout << std::endl;
        for (const auto& ip : drivers[di].ips)
            std::wcout << L"    " << ip << std::endl;
    }

    // Step 5: Eject any previously loaded instance, then inject DLL
    EjectDLL(rslinxPid, dllPath);
    Sleep(500);

    if (!InjectDLL(rslinxPid, dllPath))
    {
        std::wcerr << L"[FAIL] DLL injection failed" << std::endl;
        PipeClose();
        return 1;
    }

    // Step 6: Wait for hook to connect to pipe
    std::wcout << L"[INFO] Waiting for hook to connect to pipe..." << std::endl;
    if (!WaitForPipeConnection(10000))
    {
        std::wcerr << L"[FAIL] Hook did not connect to pipe within 10s" << std::endl;
        PipeClose();
        EjectDLL(rslinxPid, dllPath);
        return 1;
    }
    std::wcout << L"[OK] Hook connected via pipe" << std::endl;

    // Step 7: Send config over pipe
    PipeSendConfig(drivers, driverNeedsHotLoad, logDir, debugXml, false, probeDispids);
    std::wcout << L"[OK] Config sent via pipe" << std::endl;

    // Step 8: Read pipe messages (log lines streamed to console in real-time)
    std::wcout << std::endl << L"--- Hook Log (live) ---" << std::endl;
    PipeReadLoop(210000);

    // Step 9: Cleanup pipe and eject DLL
    PipeSendStop();
    PipeClose();
    Sleep(500);
    EjectDLL(rslinxPid, dllPath);

    // Step 10: Read and display results (hook_results.txt still written by hook)
    PrintHeader(L"Hook Results");

    std::wcout << std::endl;

    int totalIdentified = 0;
    int totalDevices = 0;
    std::wifstream resultFile(LogPath(logDir, L"hook_results.txt").c_str());
    if (resultFile.is_open())
    {
        std::wcout << L"--- Results ---" << std::endl;
        std::wstring line;
        while (std::getline(resultFile, line))
        {
            std::wcout << L"  " << line << std::endl;
            if (line.find(L"DEVICES_IDENTIFIED:") != std::wstring::npos)
                totalIdentified += _wtoi(line.substr(line.find(L':') + 1).c_str());
            if (line.find(L"DEVICES_TOTAL:") != std::wstring::npos)
                totalDevices += _wtoi(line.substr(line.find(L':') + 1).c_str());
        }
        resultFile.close();
    }
    else
    {
        std::wcerr << L"[FAIL] No results file - hook may have failed" << std::endl;
    }

    // Step 8: Save topology after injection (debug-xml only)
    if (debugXml)
    {
        PrintHeader(L"Topology AFTER Injection");
        TopologyBrowser browser;
        if (browser.Initialize())
        {
            std::wstring topoAfterInject = LogPath(logDir, L"topology_after_inject.xml");
            if (browser.SaveTopologyXML(topoAfterInject.c_str()))
            {
                std::wcout << L"[OK] Saved topology" << std::endl;
                for (const auto& drv : drivers)
                {
                    auto devs = browser.ParseDevicesFromXML(topoAfterInject.c_str(), drv.name);
                    std::wcout << L"  Driver '" << drv.name << L"': " << devs.size() << L" devices" << std::endl;
                    for (const auto& d : devs)
                        std::wcout << L"    - " << d << std::endl;
                }
            }
            browser.Uninitialize();
        }
    }

    PrintHeader(L"Final Result");
    std::wcout << L"[OK] Browse complete -- " << drivers.size() << L" driver(s), "
               << totalIdentified << L" devices identified" << std::endl;

    return totalIdentified > 0 ? 0 : 1;
}

// ============================================================
// Original external COM mode
// ============================================================

static int RunExternalMode(const std::wstring& driverName, const std::wstring& targetIP, DWORD browseTimeout, const std::wstring& logDir, bool debugXml)
{
    // Initialize
    TopologyBrowser browser;

    PrintHeader(L"Initializing RSLinx COM");
    if (!browser.Initialize())
    {
        std::wcerr << L"[FAIL] " << browser.GetLastError() << std::endl;
        return 1;
    }
    std::wcout << L"[OK] Initialization complete" << std::endl;

    // Ensure log directory exists
    CreateDirectoryW(logDir.c_str(), NULL);

    // Save topology BEFORE (debug-xml only — informational)
    if (debugXml)
    {
        PrintHeader(L"Topology BEFORE");
        std::wstring topoBefore = LogPath(logDir, L"topology_before_cpp.xml");
        if (browser.SaveTopologyXML(topoBefore.c_str()))
        {
            std::wcout << L"[OK] Saved " << topoBefore << std::endl;
            auto devsBefore = browser.ParseDevicesFromXML(topoBefore.c_str(), driverName);
            std::wcout << L"Devices in '" << driverName << L"' bus: " << devsBefore.size() << std::endl;
            for (const auto& d : devsBefore) std::wcout << L"  - " << d << std::endl;
        }
    }

    // Step 1: Read Node Table IPs from registry
    PrintHeader(L"Reading Node Table");
    std::vector<std::wstring> nodeTableIPs = DriverConfig::ReadNodeTable(driverName);

    if (nodeTableIPs.empty())
    {
        std::wcout << L"[WARN] No IPs in Node Table, using target IP only" << std::endl;
        nodeTableIPs.push_back(targetIP);
    }

    // Step 2: Add each device via ConnectNewDevice
    PrintHeader(L"Adding devices via ConnectNewDevice");
    int addedCount = 0;
    int failedCount = 0;

    for (const auto& ip : nodeTableIPs)
    {
        if (browser.AddDeviceManually(driverName, ip))
        {
            addedCount++;
        }
        else
        {
            failedCount++;
            if (addedCount == 0 && failedCount >= 2)
            {
                std::wcout << L"[WARN] Multiple failures, stopping device addition" << std::endl;
                break;
            }
        }
    }

    std::wcout << std::endl;
    std::wcout << L"Added: " << addedCount << L"  Failed: " << failedCount << std::endl;

    // Step 3: Save topology AFTER adding devices
    PrintHeader(L"Topology AFTER ConnectNewDevice");
    std::wstring topoAfter = LogPath(logDir, L"topology_after_cpp.xml");
    if (browser.SaveTopologyXML(topoAfter.c_str()))
    {
        std::wcout << L"[OK] Saved " << topoAfter << std::endl;
        auto devsAfter = browser.ParseDevicesFromXML(topoAfter.c_str(), driverName);
        std::wcout << L"Devices in '" << driverName << L"' bus: " << devsAfter.size() << std::endl;
        for (const auto& d : devsAfter) std::wcout << L"  - " << d << std::endl;
    }

    // Check for target
    PrintHeader(L"Results");
    bool found = false;

    auto finalDevs = browser.ParseDevicesFromXML(topoAfter.c_str(), driverName);
    for (const auto& dev : finalDevs)
    {
        if (dev.find(targetIP) != std::wstring::npos)
        {
            found = true;
            std::wcout << L"[SUCCESS] Target device " << targetIP << L" found in topology!" << std::endl;
            break;
        }
    }

    if (!found)
    {
        std::wcout << L"[INFO] Target " << targetIP << L" not in topology." << std::endl;
    }

    std::wcout << L"Total in topology: " << finalDevs.size() << L" device(s)" << std::endl;
    std::wcout << L"ConnectNewDevice results: " << addedCount << L" added, " << failedCount << L" failed" << std::endl;

    // Clean up XML files when not in debug-xml mode
    if (!debugXml)
        DeleteFileW(topoAfter.c_str());

    browser.Uninitialize();
    return found ? 0 : 1;
}

// ============================================================
// Monitor Mode — browse existing driver topology (no creation)
// ============================================================

static int RunMonitorMode(const std::vector<DriverSpec>& drivers, const std::wstring& logDir, bool debugXml, bool probeDispids)
{
    PrintHeader(L"RSLinx Monitor Mode (browse existing)");

    // Ensure log directory exists
    CreateDirectoryW(logDir.c_str(), NULL);

    // Step 0: Verify all drivers exist (no creation, no modification)
    for (const auto& drv : drivers)
    {
        if (!DriverExistsInRegistry(drv.name))
        {
            std::wcerr << L"[FAIL] Driver '" << drv.name << L"' not found in registry" << std::endl;
            std::wcerr << L"[INFO] Use default mode (without --monitor) to create a new driver" << std::endl;
            return 1;
        }
        std::wcout << L"[OK] Driver '" << drv.name << L"' exists in registry" << std::endl;
    }

    // Step 1: Find RSLinx
    DWORD rslinxPid = FindProcessByName(L"RSLinx.exe");
    if (rslinxPid == 0) rslinxPid = FindProcessByName(L"RSLINX.EXE");
    if (rslinxPid == 0) rslinxPid = FindProcessByName(L"rslinx.exe");
    if (rslinxPid == 0)
    {
        std::wcerr << L"[FAIL] RSLinx.exe not found. Is it running?" << std::endl;
        return 1;
    }
    std::wcout << L"[OK] Found RSLinx.exe PID: " << rslinxPid << std::endl;

    // Step 2: Create pipe server BEFORE injection
    std::wcout << L"[INFO] Creating named pipe server..." << std::endl;
    if (!CreatePipeServer())
    {
        std::wcerr << L"[FAIL] Cannot create named pipe: " << GetLastError() << std::endl;
        return 1;
    }

    // Step 3: Delete old results/log files
    DeleteFileW(LogPath(logDir, L"hook_results.txt").c_str());
    DeleteFileW(LogPath(logDir, L"hook_log.txt").c_str());

    // Step 4: Find RSLinxHook.dll path (same directory as our exe)
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    std::wstring dllPath(exePath);
    size_t lastSlash = dllPath.rfind(L'\\');
    if (lastSlash != std::wstring::npos) dllPath = dllPath.substr(0, lastSlash + 1);
    dllPath += L"RSLinxHook.dll";
    if (GetFileAttributesW(dllPath.c_str()) == INVALID_FILE_ATTRIBUTES)
    {
        std::wcerr << L"[FAIL] DLL not found: " << dllPath << std::endl;
        PipeClose();
        return 1;
    }
    std::wcout << L"[OK] DLL path: " << dllPath << std::endl;

    // Log driver info
    for (const auto& drv : drivers)
    {
        std::wcout << L"  Driver '" << drv.name << L"': " << drv.ips.size() << L" IP(s)" << std::endl;
        for (const auto& ip : drv.ips)
            std::wcout << L"    " << ip << std::endl;
    }

    // Step 5: Eject any previous instance, then inject
    EjectDLL(rslinxPid, dllPath);
    Sleep(500);

    if (!InjectDLL(rslinxPid, dllPath))
    {
        std::wcerr << L"[FAIL] DLL injection failed" << std::endl;
        PipeClose();
        return 1;
    }

    // Step 6: Wait for hook to connect to pipe
    std::wcout << L"[INFO] Waiting for hook to connect to pipe..." << std::endl;
    if (!WaitForPipeConnection(10000))
    {
        std::wcerr << L"[FAIL] Hook did not connect to pipe within 10s" << std::endl;
        PipeClose();
        EjectDLL(rslinxPid, dllPath);
        return 1;
    }
    std::wcout << L"[OK] Hook connected via pipe" << std::endl;

    // Step 7: Send config over pipe (MODE=inject preserved — monitor flag only skips driver creation)
    std::vector<bool> noHotLoad(drivers.size(), false);
    PipeSendConfig(drivers, noHotLoad, logDir, debugXml, false, probeDispids);
    std::wcout << L"[OK] Config sent via pipe" << std::endl;

    // Step 8: Read pipe messages (log lines streamed to console in real-time)
    std::wcout << std::endl << L"--- Hook Log (live) ---" << std::endl;
    PipeReadLoop(120000);

    // Step 9: Cleanup pipe and eject DLL
    PipeSendStop();
    PipeClose();
    Sleep(500);
    EjectDLL(rslinxPid, dllPath);

    // Step 10: Read results file (hook_results.txt still written by hook)
    PrintHeader(L"Hook Results");

    std::wcout << L"--- Results ---" << std::endl;
    std::wifstream results(LogPath(logDir, L"hook_results.txt").c_str());
    int totalIdentified = 0;
    int totalDeviceCount = 0;
    if (results.is_open())
    {
        std::wstring line;
        while (std::getline(results, line))
        {
            std::wcout << L"  " << line << std::endl;
            if (line.find(L"DEVICES_IDENTIFIED:") != std::wstring::npos)
                totalIdentified += _wtoi(line.substr(line.find(L':') + 1).c_str());
            if (line.find(L"DEVICES_TOTAL:") != std::wstring::npos)
                totalDeviceCount += _wtoi(line.substr(line.find(L':') + 1).c_str());
        }
        results.close();
    }

    // Step 11: Save post-browse topology XML
    if (debugXml)
    {
        PrintHeader(L"Topology AFTER Browse");
        TopologyBrowser browser;
        if (browser.Initialize())
        {
            std::wstring topoPath = LogPath(logDir, L"topology_after_monitor.xml");
            if (browser.SaveTopologyXML(topoPath.c_str()))
            {
                std::wcout << L"[OK] Saved topology" << std::endl;
                for (const auto& drv : drivers)
                {
                    auto devs = browser.ParseDevicesFromXML(topoPath.c_str(), drv.name);
                    std::wcout << L"  Driver '" << drv.name << L"': " << devs.size() << L" devices" << std::endl;
                    for (const auto& d : devs) std::wcout << L"    - " << d << std::endl;
                }
            }
            browser.Uninitialize();
        }
    }

    PrintHeader(L"Final Result");
    std::wcout << L"[OK] Full browse complete -- " << drivers.size() << L" driver(s), "
               << totalIdentified << L" devices identified" << std::endl;

    return totalIdentified > 0 ? 0 : 1;
}

// ============================================================
// Entry Point
// ============================================================

int wmain(int argc, wchar_t* argv[])
{
    SetConsoleOutputCP(CP_UTF8);

    std::vector<DriverSpec> drivers;
    bool monitorMode = false;
    std::wstring logDir = L"C:\\temp";
    bool debugXml = false;
    bool probeDispids = false;

    // Parse arguments — --driver pushes a new entry, --ip appends to the last driver
    int posArg = 0;
    for (int i = 1; i < argc; i++)
    {
        if (_wcsicmp(argv[i], L"--inject") == 0 || _wcsicmp(argv[i], L"-inject") == 0)
        {
            // Accepted for backwards compat, inject is now the default
        }
        else if (_wcsicmp(argv[i], L"--monitor") == 0 || _wcsicmp(argv[i], L"-monitor") == 0)
        {
            monitorMode = true;
        }
        else if (_wcsicmp(argv[i], L"--debug-xml") == 0 || _wcsicmp(argv[i], L"-debug-xml") == 0)
        {
            debugXml = true;
        }
        else if (_wcsicmp(argv[i], L"--probe") == 0 || _wcsicmp(argv[i], L"-probe") == 0)
        {
            probeDispids = true;
        }
        else if ((_wcsicmp(argv[i], L"--ip") == 0 || _wcsicmp(argv[i], L"-ip") == 0) && i + 1 < argc)
        {
            if (drivers.empty()) drivers.push_back({L"Test", {}});
            drivers.back().ips.push_back(argv[++i]);
        }
        else if ((_wcsicmp(argv[i], L"--logdir") == 0 || _wcsicmp(argv[i], L"-logdir") == 0) && i + 1 < argc)
        {
            logDir = argv[++i];
        }
        else if ((_wcsicmp(argv[i], L"--driver") == 0 || _wcsicmp(argv[i], L"-driver") == 0) && i + 1 < argc)
        {
            drivers.push_back({argv[++i], {}});
        }
        else if ((_wcsicmp(argv[i], L"--timeout") == 0 || _wcsicmp(argv[i], L"-timeout") == 0) && i + 1 < argc)
        {
            ++i; // consumed but no longer used (hook manages its own timeouts)
        }
        else if (argv[i][0] != L'-')
        {
            // Positional args: driver, ip (backward compat)
            posArg++;
            if (posArg == 1)
            {
                drivers.push_back({argv[i], {}});
            }
            else if (posArg == 2)
            {
                if (drivers.empty()) drivers.push_back({L"Test", {}});
                drivers.back().ips.push_back(argv[i]);
            }
        }
    }

    // Default driver if none specified
    if (drivers.empty())
        drivers.push_back({L"Test", {}});

    PrintHeader(L"RSLinx Topology Browser - COM Automation");

    const wchar_t* modeStr = monitorMode ? L"Monitor (browse existing)" : L"Browse (create/update driver)";
    std::wcout << L"Mode: " << modeStr << std::endl;
    std::wcout << L"Drivers: " << drivers.size() << std::endl;
    for (const auto& drv : drivers)
    {
        std::wcout << L"  " << drv.name;
        if (!drv.ips.empty())
            std::wcout << L" (" << drv.ips.size() << L" IPs)";
        std::wcout << std::endl;
    }
    if (logDir != L"C:\\temp")
        std::wcout << L"Log directory: " << logDir << std::endl;
    if (debugXml)
        std::wcout << L"Debug XML: enabled" << std::endl;
    if (probeDispids)
        std::wcout << L"DISPID probing: enabled" << std::endl;

    if (monitorMode)
        return RunMonitorMode(drivers, logDir, debugXml, probeDispids);
    else
        return RunInjectMode(drivers, logDir, debugXml, probeDispids);
}
