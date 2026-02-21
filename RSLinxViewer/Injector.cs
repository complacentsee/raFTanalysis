using System.Runtime.InteropServices;
using System.Text;

namespace RSLinxViewer;

/// <summary>
/// P/Invoke-based DLL injection/ejection into RSLinx.exe.
/// Direct port of raFTMEanalysis/RSLinxBrowse/main.cpp (lines 84-257).
/// </summary>
static class Injector
{
    #region P/Invoke Constants

    const uint PROCESS_CREATE_THREAD = 0x0002;
    const uint PROCESS_VM_OPERATION = 0x0008;
    const uint PROCESS_VM_READ = 0x0010;
    const uint PROCESS_VM_WRITE = 0x0020;
    const uint PROCESS_QUERY_INFORMATION = 0x0400;

    const uint MEM_COMMIT = 0x1000;
    const uint MEM_RESERVE = 0x2000;
    const uint MEM_RELEASE = 0x8000;
    const uint PAGE_READWRITE = 0x04;

    const uint INFINITE = 0xFFFFFFFF;
    const uint WAIT_OBJECT_0 = 0;
    const uint WAIT_TIMEOUT = 0x102;

    const uint TH32CS_SNAPPROCESS = 0x02;
    const uint TH32CS_SNAPMODULE = 0x08;

    const uint TOKEN_ADJUST_PRIVILEGES = 0x0020;
    const uint TOKEN_QUERY = 0x0008;
    const uint SE_PRIVILEGE_ENABLED = 0x00000002;
    const uint ERROR_NOT_ALL_ASSIGNED = 1300;
    const uint ERROR_ACCESS_DENIED = 5;

    const string SE_DEBUG_NAME = "SeDebugPrivilege";

    #endregion

    #region P/Invoke Structs

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct PROCESSENTRY32W
    {
        public uint dwSize;
        public uint cntUsage;
        public uint th32ProcessID;
        public IntPtr th32DefaultHeapID;
        public uint th32ModuleID;
        public uint cntThreads;
        public uint th32ParentProcessID;
        public int pcPriClassBase;
        public uint dwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szExeFile;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct LUID
    {
        public uint LowPart;
        public int HighPart;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct LUID_AND_ATTRIBUTES
    {
        public LUID Luid;
        public uint Attributes;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct TOKEN_PRIVILEGES
    {
        public uint PrivilegeCount;
        public LUID_AND_ATTRIBUTES Privileges;
    }

    #endregion

    #region P/Invoke Declarations

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool CloseHandle(IntPtr hObject);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, uint flAllocationType, uint flProtect);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, uint dwFreeType);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] lpBuffer, uint nSize, out uint lpNumberOfBytesWritten);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr CreateRemoteThread(IntPtr hProcess, IntPtr lpThreadAttributes, uint dwStackSize, IntPtr lpStartAddress, IntPtr lpParameter, uint dwCreationFlags, out uint lpThreadId);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern uint WaitForSingleObject(IntPtr hHandle, uint dwMilliseconds);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool GetExitCodeThread(IntPtr hThread, out uint lpExitCode);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern IntPtr GetModuleHandleW(string lpModuleName);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi)]
    static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr CreateToolhelp32Snapshot(uint dwFlags, uint th32ProcessID);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern bool Process32FirstW(IntPtr hSnapshot, ref PROCESSENTRY32W lppe);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern bool Process32NextW(IntPtr hSnapshot, ref PROCESSENTRY32W lppe);

    [DllImport("advapi32.dll", SetLastError = true)]
    static extern bool OpenProcessToken(IntPtr ProcessHandle, uint DesiredAccess, out IntPtr TokenHandle);

    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern bool LookupPrivilegeValueW(string? lpSystemName, string lpName, out LUID lpLuid);

    [DllImport("advapi32.dll", SetLastError = true)]
    static extern bool AdjustTokenPrivileges(IntPtr TokenHandle, bool DisableAllPrivileges, ref TOKEN_PRIVILEGES NewState, uint BufferLength, IntPtr PreviousState, IntPtr ReturnLength);

    [DllImport("psapi.dll", SetLastError = true)]
    static extern bool EnumProcessModules(IntPtr hProcess, IntPtr[] lphModule, uint cb, out uint lpcbNeeded);

    [DllImport("psapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern uint GetModuleFileNameExW(IntPtr hProcess, IntPtr hModule, StringBuilder lpFilename, uint nSize);

    // For getting current process handle
    [DllImport("kernel32.dll")]
    static extern IntPtr GetCurrentProcess();

    #endregion

    /// <summary>Find RSLinx.exe PID using toolhelp snapshot.</summary>
    public static uint FindRSLinxPid()
    {
        IntPtr hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
        if (hSnap == IntPtr.Zero || hSnap == new IntPtr(-1))
            return 0;

        var pe = new PROCESSENTRY32W();
        pe.dwSize = (uint)Marshal.SizeOf<PROCESSENTRY32W>();

        uint pid = 0;
        if (Process32FirstW(hSnap, ref pe))
        {
            do
            {
                if (pe.szExeFile.Equals("RSLinx.exe", StringComparison.OrdinalIgnoreCase))
                {
                    pid = pe.th32ProcessID;
                    break;
                }
            } while (Process32NextW(hSnap, ref pe));
        }

        CloseHandle(hSnap);
        return pid;
    }

    /// <summary>
    /// Enable SeDebugPrivilege in the current process token.
    /// Required for opening service processes (Session 0 / SYSTEM).
    /// </summary>
    public static bool EnableDebugPrivilege(Action<string>? log = null)
    {
        if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, out IntPtr hToken))
        {
            log?.Invoke($"  OpenProcessToken failed: {Marshal.GetLastWin32Error()}");
            return false;
        }

        if (!LookupPrivilegeValueW(null, SE_DEBUG_NAME, out LUID luid))
        {
            log?.Invoke($"  LookupPrivilegeValue failed: {Marshal.GetLastWin32Error()}");
            CloseHandle(hToken);
            return false;
        }
        log?.Invoke($"  LUID: {luid.LowPart}:{luid.HighPart}");

        var tp = new TOKEN_PRIVILEGES
        {
            PrivilegeCount = 1,
            Privileges = new LUID_AND_ATTRIBUTES
            {
                Luid = luid,
                Attributes = SE_PRIVILEGE_ENABLED
            }
        };

        bool ok = AdjustTokenPrivileges(hToken, false, ref tp, (uint)Marshal.SizeOf<TOKEN_PRIVILEGES>(), IntPtr.Zero, IntPtr.Zero);
        uint err = (uint)Marshal.GetLastWin32Error();
        CloseHandle(hToken);
        log?.Invoke($"  AdjustTokenPrivileges: ok={ok} err={err}");

        return ok && err != ERROR_NOT_ALL_ASSIGNED;
    }

    /// <summary>
    /// Open a process with injection rights, with SeDebugPrivilege retry on access denied.
    /// </summary>
    static IntPtr OpenProcessForInjection(uint pid, uint access, Action<string> log)
    {
        IntPtr hProcess = OpenProcess(access, false, pid);
        if (hProcess == IntPtr.Zero && (uint)Marshal.GetLastWin32Error() == ERROR_ACCESS_DENIED)
        {
            log("[INFO] Access denied - enabling SeDebugPrivilege (service mode)...");
            if (EnableDebugPrivilege(log))
            {
                log("[OK] SeDebugPrivilege enabled");
                hProcess = OpenProcess(access, false, pid);
            }
            else
            {
                log("[FAIL] Could not enable SeDebugPrivilege - run as Administrator");
            }
        }
        return hProcess;
    }

    /// <summary>
    /// Inject a DLL into the target process via CreateRemoteThread + LoadLibraryW.
    /// </summary>
    public static bool InjectDLL(uint pid, string dllPath, Action<string> log)
    {
        log($"[INFO] Opening process PID {pid}");

        uint access = PROCESS_CREATE_THREAD | PROCESS_VM_OPERATION | PROCESS_VM_WRITE | PROCESS_VM_READ | PROCESS_QUERY_INFORMATION;
        IntPtr hProcess = OpenProcessForInjection(pid, access, log);
        if (hProcess == IntPtr.Zero)
        {
            log($"[FAIL] OpenProcess failed: {Marshal.GetLastWin32Error()}");
            return false;
        }

        // Allocate memory in target for DLL path (wide string)
        byte[] pathBytes = Encoding.Unicode.GetBytes(dllPath + '\0');
        uint pathSize = (uint)pathBytes.Length;

        IntPtr pRemotePath = VirtualAllocEx(hProcess, IntPtr.Zero, pathSize, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
        if (pRemotePath == IntPtr.Zero)
        {
            log($"[FAIL] VirtualAllocEx failed: {Marshal.GetLastWin32Error()}");
            CloseHandle(hProcess);
            return false;
        }

        if (!WriteProcessMemory(hProcess, pRemotePath, pathBytes, pathSize, out _))
        {
            log($"[FAIL] WriteProcessMemory failed: {Marshal.GetLastWin32Error()}");
            VirtualFreeEx(hProcess, pRemotePath, 0, MEM_RELEASE);
            CloseHandle(hProcess);
            return false;
        }

        // Get LoadLibraryW address (same in all processes due to kernel32 base)
        IntPtr hKernel32 = GetModuleHandleW("kernel32.dll");
        IntPtr pLoadLibraryW = GetProcAddress(hKernel32, "LoadLibraryW");
        if (pLoadLibraryW == IntPtr.Zero)
        {
            log("[FAIL] GetProcAddress(LoadLibraryW) failed");
            VirtualFreeEx(hProcess, pRemotePath, 0, MEM_RELEASE);
            CloseHandle(hProcess);
            return false;
        }

        log("[INFO] Creating remote thread in rslinx.exe...");
        IntPtr hThread = CreateRemoteThread(hProcess, IntPtr.Zero, 0, pLoadLibraryW, pRemotePath, 0, out _);
        if (hThread == IntPtr.Zero)
        {
            log($"[FAIL] CreateRemoteThread failed: {Marshal.GetLastWin32Error()}");
            VirtualFreeEx(hProcess, pRemotePath, 0, MEM_RELEASE);
            CloseHandle(hProcess);
            return false;
        }

        WaitForSingleObject(hThread, 10000);

        GetExitCodeThread(hThread, out uint exitCode);
        log($"[INFO] LoadLibrary returned: 0x{exitCode:X}");

        CloseHandle(hThread);
        VirtualFreeEx(hProcess, pRemotePath, 0, MEM_RELEASE);
        CloseHandle(hProcess);

        if (exitCode == 0)
        {
            log("[FAIL] LoadLibrary returned NULL - DLL load failed");
            return false;
        }

        log("[OK] DLL injected successfully");
        return true;
    }

    /// <summary>
    /// Eject a previously injected DLL by calling FreeLibrary in the remote process.
    /// </summary>
    public static bool EjectDLL(uint pid, string dllPath, Action<string> log)
    {
        uint access = PROCESS_CREATE_THREAD | PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_QUERY_INFORMATION;
        IntPtr hProcess = OpenProcessForInjection(pid, access, log);
        if (hProcess == IntPtr.Zero)
            return false;

        // Find the DLL's module handle in the remote process
        var hMods = new IntPtr[1024];
        if (!EnumProcessModules(hProcess, hMods, (uint)(hMods.Length * IntPtr.Size), out uint cbNeeded))
        {
            CloseHandle(hProcess);
            return false;
        }

        IntPtr hTargetDLL = IntPtr.Zero;
        int modCount = (int)(cbNeeded / (uint)IntPtr.Size);
        var sb = new StringBuilder(260);

        for (int i = 0; i < modCount; i++)
        {
            sb.Clear();
            if (GetModuleFileNameExW(hProcess, hMods[i], sb, 260) > 0)
            {
                if (sb.ToString().Equals(dllPath, StringComparison.OrdinalIgnoreCase))
                {
                    hTargetDLL = hMods[i];
                    break;
                }
            }
        }

        if (hTargetDLL == IntPtr.Zero)
        {
            CloseHandle(hProcess);
            return false; // DLL not loaded
        }

        IntPtr hKernel32 = GetModuleHandleW("kernel32.dll");
        IntPtr pFreeLibrary = GetProcAddress(hKernel32, "FreeLibrary");

        IntPtr hThread = CreateRemoteThread(hProcess, IntPtr.Zero, 0, pFreeLibrary, hTargetDLL, 0, out _);
        if (hThread != IntPtr.Zero)
        {
            WaitForSingleObject(hThread, 5000);
            CloseHandle(hThread);
        }

        CloseHandle(hProcess);
        log("[OK] Ejected previous DLL instance");
        return true;
    }
}
