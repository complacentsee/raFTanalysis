#include "STAHook.h"
#include "Logging.h"

// ============================================================
// STAHook globals
// ============================================================

MainSTAFunc g_pMainSTAFunc = nullptr;
HHOOK g_hHook = NULL;
DWORD g_mainThreadId = 0;
volatile LONG g_browseRequested = 0;
volatile LONG g_browseResult = (LONG)E_PENDING;
HookConfig* g_pSharedConfig = nullptr;
WNDPROC g_origWndProc = nullptr;

// ============================================================
// STAHook implementations
// ============================================================

// Find the main (oldest) thread of current process
DWORD FindMainThreadId()
{
    DWORD pid = GetCurrentProcessId();
    DWORD mainTid = 0;
    ULONGLONG oldestTime = (ULONGLONG)(-1);

    HANDLE snap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0);
    if (snap == INVALID_HANDLE_VALUE) return 0;

    THREADENTRY32 te = {};
    te.dwSize = sizeof(te);
    if (Thread32First(snap, &te))
    {
        do
        {
            if (te.th32OwnerProcessID == pid)
            {
                HANDLE hThread = OpenThread(THREAD_QUERY_LIMITED_INFORMATION, FALSE, te.th32ThreadID);
                if (hThread)
                {
                    FILETIME create, exitFT, kernel, user;
                    if (GetThreadTimes(hThread, &create, &exitFT, &kernel, &user))
                    {
                        ULONGLONG t = ((ULONGLONG)create.dwHighDateTime << 32) | create.dwLowDateTime;
                        if (t < oldestTime)
                        {
                            oldestTime = t;
                            mainTid = te.th32ThreadID;
                        }
                    }
                    CloseHandle(hThread);
                }
            }
        } while (Thread32Next(snap, &te));
    }
    CloseHandle(snap);
    return mainTid;
}

// WH_GETMESSAGE hook callback — runs on MAIN STA thread
LRESULT CALLBACK MainSTAHookProc(int code, WPARAM wp, LPARAM lp)
{
    if (code >= 0)
    {
        MSG* pMsg = (MSG*)lp;
        if (pMsg->message == WM_NULL && pMsg->wParam == HOOK_MAGIC_WPARAM)
        {
            if (InterlockedCompareExchange(&g_browseRequested, 0, 1) == 1)
            {
                Log(L"[HOOK] Processing request on TID=%d", GetCurrentThreadId());
                HRESULT hr = g_pMainSTAFunc ? g_pMainSTAFunc() : E_POINTER;
                InterlockedExchange(&g_browseResult, (LONG)hr);
                Log(L"[HOOK] Result: 0x%08x", hr);
            }
        }
    }
    return CallNextHookEx(g_hHook, code, wp, lp);
}

// Window subclass fallback WndProc
LRESULT CALLBACK HookWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    if (msg == SUBCLASS_MSG)
    {
        if (InterlockedCompareExchange(&g_browseRequested, 0, 1) == 1)
        {
            Log(L"[SUBCLASS] Processing request on TID=%d", GetCurrentThreadId());
            HRESULT hr = g_pMainSTAFunc ? g_pMainSTAFunc() : E_POINTER;
            InterlockedExchange(&g_browseResult, (LONG)hr);
            Log(L"[SUBCLASS] Result: 0x%08x", hr);
        }
        return 0;
    }
    return CallWindowProcW(g_origWndProc, hwnd, msg, wp, lp);
}

// EnumThreadWindows callback — find any window on the main thread
BOOL CALLBACK FindThreadWindowProc(HWND hwnd, LPARAM lParam)
{
    HWND* pResult = (HWND*)lParam;
    *pResult = hwnd;
    return FALSE;  // Stop after first window
}

// Execute a function on the main STA thread.
// Tries SetWindowsHookEx first, falls back to window subclass.
HRESULT ExecuteOnMainSTA(MainSTAFunc func)
{
    Log(L"=== ExecuteOnMainSTA ===");

    g_pMainSTAFunc = func;

    // Find main thread
    g_mainThreadId = FindMainThreadId();
    Log(L"  Main thread TID: %d (our TID: %d)", g_mainThreadId, GetCurrentThreadId());

    if (g_mainThreadId == 0 || g_mainThreadId == GetCurrentThreadId())
    {
        Log(L"  FAIL: Could not find main thread (or we ARE the main thread)");
        return E_FAIL;
    }

    // --- Try 1: SetWindowsHookEx(WH_GETMESSAGE) ---
    Log(L"  Trying SetWindowsHookEx(WH_GETMESSAGE)...");
    InterlockedExchange(&g_browseRequested, 1);
    InterlockedExchange(&g_browseResult, (LONG)E_PENDING);

    g_hHook = SetWindowsHookExW(WH_GETMESSAGE, MainSTAHookProc,
                                 GetModuleHandle(NULL), g_mainThreadId);
    if (g_hHook)
    {
        Log(L"  Hook installed: 0x%p", g_hHook);

        // Post WM_NULL with magic wParam to trigger the hook
        BOOL posted = PostThreadMessageW(g_mainThreadId, WM_NULL, HOOK_MAGIC_WPARAM, 0);
        Log(L"  PostThreadMessage: %s", posted ? L"OK" : L"FAILED");

        if (posted)
        {
            // Wait for result (up to 30s)
            DWORD t0 = GetTickCount();
            while (!g_shouldStop && InterlockedCompareExchange(&g_browseResult, 0, 0) == (LONG)E_PENDING)
            {
                if (GetTickCount() - t0 > 30000)
                {
                    Log(L"  TIMEOUT waiting for hook result after 30s");
                    break;
                }
                Sleep(100);
            }
        }
        else
        {
            Log(L"  PostThreadMessage failed (err=%d), will try fallback", GetLastError());
            InterlockedExchange(&g_browseRequested, 1);
        }

        UnhookWindowsHookEx(g_hHook);
        g_hHook = NULL;

        HRESULT result = (HRESULT)InterlockedCompareExchange(&g_browseResult, 0, 0);
        if (result != (LONG)E_PENDING)
        {
            Log(L"  Hook approach result: 0x%08x", result);
            return result;
        }
    }
    else
    {
        Log(L"  SetWindowsHookEx failed: %d", GetLastError());
    }

    // --- Try 2: Window subclass fallback ---
    Log(L"  Trying window subclass fallback...");
    HWND hTargetWnd = NULL;
    EnumThreadWindows(g_mainThreadId, FindThreadWindowProc, (LPARAM)&hTargetWnd);

    if (!hTargetWnd)
    {
        Log(L"  FAIL: No windows found on main thread TID=%d", g_mainThreadId);
        return E_FAIL;
    }

    wchar_t cls[256] = {};
    GetClassNameW(hTargetWnd, cls, 256);
    Log(L"  Found window: HWND=0x%p class=\"%s\"", hTargetWnd, cls);

    // Subclass the window
    InterlockedExchange(&g_browseRequested, 1);
    InterlockedExchange(&g_browseResult, (LONG)E_PENDING);

    g_origWndProc = (WNDPROC)SetWindowLongPtrW(hTargetWnd, GWLP_WNDPROC, (LONG_PTR)HookWndProc);
    if (g_origWndProc)
    {
        Log(L"  Subclassed window, sending SUBCLASS_MSG...");
        SendMessageW(hTargetWnd, SUBCLASS_MSG, 0, 0);

        // Restore original wndproc
        SetWindowLongPtrW(hTargetWnd, GWLP_WNDPROC, (LONG_PTR)g_origWndProc);
        g_origWndProc = nullptr;

        HRESULT result = (HRESULT)InterlockedCompareExchange(&g_browseResult, 0, 0);
        Log(L"  Subclass approach result: 0x%08x", result);
        return result;
    }
    else
    {
        Log(L"  SetWindowLongPtrW failed: %d", GetLastError());
        return E_FAIL;
    }
}
