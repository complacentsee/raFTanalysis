#include "Logging.h"

// ============================================================
// Logging globals
// ============================================================

FILE* g_logFile = nullptr;
CRITICAL_SECTION g_logCS;
HANDLE g_hPipe = INVALID_HANDLE_VALUE;
HANDLE g_hStopEvent = NULL;
bool g_pipeConnected = false;

// ============================================================
// Logging functions
// ============================================================

void PipeSend(const char* data, int len)
{
    if (!g_pipeConnected) return;
    OVERLAPPED ov = {};
    ov.hEvent = CreateEventW(NULL, TRUE, FALSE, NULL);
    if (!ov.hEvent) { g_pipeConnected = false; return; }
    DWORD written;
    BOOL ok = WriteFile(g_hPipe, data, len, &written, &ov);
    if (!ok)
    {
        if (GetLastError() == ERROR_IO_PENDING)
            GetOverlappedResult(g_hPipe, &ov, &written, TRUE);
        else
            g_pipeConnected = false;
    }
    CloseHandle(ov.hEvent);
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
        OVERLAPPED ov = {};
        ov.hEvent = CreateEventW(NULL, TRUE, FALSE, NULL);
        if (!ov.hEvent) return false;
        BOOL ok = ReadFile(g_hPipe, stopBuf, toRead, &bytesRead, &ov);
        if (!ok && GetLastError() == ERROR_IO_PENDING)
            GetOverlappedResult(g_hPipe, &ov, &bytesRead, TRUE);
        CloseHandle(ov.hEvent);
        if (bytesRead > 0) {
            stopBuf[bytesRead] = '\0';
            if (strstr(stopBuf, "STOP")) {
                return true;
            }
        }
    }
    return false;
}

// ============================================================
// Pipe server functions (hook becomes the named pipe server)
// ============================================================

bool PipeCreateServer()
{
    g_hStopEvent = CreateEventW(NULL, TRUE, FALSE, NULL);
    g_hPipe = CreateNamedPipeW(
        L"\\\\.\\pipe\\RSLinxHook",
        PIPE_ACCESS_DUPLEX | FILE_FLAG_OVERLAPPED,
        PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT,
        1, 4096, 4096, 0, NULL);
    return g_hPipe != INVALID_HANDLE_VALUE;
}

bool PipeAcceptClient()
{
    if (g_hPipe == INVALID_HANDLE_VALUE) return false;

    OVERLAPPED ov = {};
    ov.hEvent = CreateEventW(NULL, TRUE, FALSE, NULL);
    if (!ov.hEvent) return false;

    BOOL ok = ConnectNamedPipe(g_hPipe, &ov);
    if (!ok)
    {
        DWORD err = GetLastError();
        if (err == ERROR_PIPE_CONNECTED)
        {
            CloseHandle(ov.hEvent);
            g_pipeConnected = true;
            return true;
        }
        if (err == ERROR_IO_PENDING)
        {
            // Wait for either a client connection or the stop event
            HANDLE handles[2] = { ov.hEvent, g_hStopEvent };
            DWORD waitResult = WaitForMultipleObjects(g_hStopEvent ? 2 : 1, handles, FALSE, INFINITE);
            if (waitResult == WAIT_OBJECT_0)
            {
                // Client connected
                DWORD dummy;
                GetOverlappedResult(g_hPipe, &ov, &dummy, FALSE);
                CloseHandle(ov.hEvent);
                g_pipeConnected = true;
                return true;
            }
            else
            {
                // Stop event signaled or error — cancel the pending I/O
                CancelIoEx(g_hPipe, &ov);
                CloseHandle(ov.hEvent);
                return false;
            }
        }
        // Other error
        CloseHandle(ov.hEvent);
        return false;
    }

    // ConnectNamedPipe returned TRUE (rare but possible)
    CloseHandle(ov.hEvent);
    g_pipeConnected = true;
    return true;
}

void PipeDisconnectClient()
{
    DisconnectNamedPipe(g_hPipe);
    g_pipeConnected = false;
}

void PipeDestroyServer()
{
    if (g_hPipe != INVALID_HANDLE_VALUE)
    {
        CloseHandle(g_hPipe);
        g_hPipe = INVALID_HANDLE_VALUE;
    }
    if (g_hStopEvent != NULL)
    {
        CloseHandle(g_hStopEvent);
        g_hStopEvent = NULL;
    }
}

// Read one newline-terminated line. Returns false on disconnect or g_shouldStop.
bool PipeReadLine(char* buf, int maxLen)
{
    if (!g_pipeConnected || g_hPipe == INVALID_HANDLE_VALUE) return false;
    int i = 0;
    while (i < maxLen - 1 && !g_shouldStop)
    {
        OVERLAPPED ov = {};
        ov.hEvent = CreateEventW(NULL, TRUE, FALSE, NULL);
        if (!ov.hEvent) { buf[i] = '\0'; return false; }

        DWORD bytesRead = 0;
        BOOL ok = ReadFile(g_hPipe, buf + i, 1, &bytesRead, &ov);
        if (!ok && GetLastError() == ERROR_IO_PENDING)
        {
            HANDLE handles[2] = { ov.hEvent, g_hStopEvent };
            DWORD wait = WaitForMultipleObjects(g_hStopEvent ? 2 : 1, handles, FALSE, INFINITE);
            if (wait == WAIT_OBJECT_0)
            {
                GetOverlappedResult(g_hPipe, &ov, &bytesRead, FALSE);
            }
            else
            {
                CancelIoEx(g_hPipe, &ov);
                CloseHandle(ov.hEvent);
                buf[i] = '\0';
                return false;
            }
        }
        CloseHandle(ov.hEvent);

        if (bytesRead == 0)
        {
            g_pipeConnected = false;
            buf[i] = '\0';
            return false;
        }
        if (buf[i] == '\n') { buf[i] = '\0'; return true; }
        if (buf[i] != '\r') i++;  // skip \r if present
    }
    buf[i] = '\0';
    return false;
}

void PipeSendLine(const char* line)
{
    PipeSend(line, (int)strlen(line));
    PipeSend("\n", 1);
}

std::wstring LogPath(const std::wstring& logDir, const wchar_t* filename)
{
    std::wstring path = logDir;
    if (!path.empty() && path.back() != L'\\')
        path += L'\\';
    path += filename;
    return path;
}
