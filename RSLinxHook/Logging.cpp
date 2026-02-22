#include "Logging.h"

// ============================================================
// Logging globals
// ============================================================

FILE* g_logFile = nullptr;
CRITICAL_SECTION g_logCS;
HANDLE g_hPipe = INVALID_HANDLE_VALUE;
bool g_pipeConnected = false;

// ============================================================
// Logging functions
// ============================================================

void PipeSend(const char* data, int len)
{
    if (!g_pipeConnected) return;
    DWORD written;
    if (!WriteFile(g_hPipe, data, len, &written, NULL))
    {
        g_pipeConnected = false;
        CloseHandle(g_hPipe);
        g_hPipe = INVALID_HANDLE_VALUE;
    }
}

void Log(const wchar_t* fmt, ...)
{
    EnterCriticalSection(&g_logCS);

    // Write to log file
    va_list args;
    va_start(args, fmt);
    if (g_logFile) {
        vfwprintf(g_logFile, fmt, args);
        fwprintf(g_logFile, L"\n");
        fflush(g_logFile);
    }
    va_end(args);

    // Stream to TUI pipe as L|<text>\n
    if (g_pipeConnected) {
        wchar_t buf[2048];
        va_start(args, fmt);
        int len = _vsnwprintf(buf, 2046, fmt, args);
        va_end(args);
        if (len > 0) {
            char utf8[4096];
            utf8[0] = 'L'; utf8[1] = '|';
            int n = WideCharToMultiByte(CP_UTF8, 0, buf, len, utf8 + 2, 4090, NULL, NULL);
            if (n > 0) {
                utf8[n + 2] = '\n';
                PipeSend(utf8, n + 3);
            }
        }
    }

    LeaveCriticalSection(&g_logCS);
}

void PipeSendTopology(const wchar_t* xmlPath)
{
    if (!g_pipeConnected) return;
    FILE* f = _wfopen(xmlPath, L"r");
    if (!f) return;

    EnterCriticalSection(&g_logCS);
    PipeSend("X|BEGIN\n", 8);
    char line[4096];
    while (fgets(line, sizeof(line), f))
        PipeSend(line, (int)strlen(line));
    PipeSend("X|END\n", 6);
    LeaveCriticalSection(&g_logCS);

    fclose(f);
}

void PipeSendStatus(int total, int identified, int events)
{
    if (!g_pipeConnected) return;
    char buf[128];
    int n = snprintf(buf, sizeof(buf), "S|%d|%d|%d\n", total, identified, events);
    EnterCriticalSection(&g_logCS);
    PipeSend(buf, n);
    LeaveCriticalSection(&g_logCS);
}

bool PipeCheckStop()
{
    if (!g_pipeConnected) return false;
    DWORD bytesAvail = 0;
    if (PeekNamedPipe(g_hPipe, NULL, 0, NULL, &bytesAvail, NULL) && bytesAvail > 0) {
        char stopBuf[64];
        DWORD bytesRead = 0;
        DWORD toRead = bytesAvail < 63 ? bytesAvail : 63;
        if (ReadFile(g_hPipe, stopBuf, toRead, &bytesRead, NULL) && bytesRead > 0) {
            stopBuf[bytesRead] = '\0';
            if (strstr(stopBuf, "STOP")) {
                return true;
            }
        }
    }
    return false;
}

std::wstring LogPath(const std::wstring& logDir, const wchar_t* filename)
{
    std::wstring path = logDir;
    if (!path.empty() && path.back() != L'\\')
        path += L'\\';
    path += filename;
    return path;
}
