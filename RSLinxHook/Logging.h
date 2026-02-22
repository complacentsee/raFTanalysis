#pragma once
#include "RSLinxHook_fwd.h"

// ============================================================
// Logging and pipe communication
// Globals defined in Logging.cpp
// ============================================================

extern FILE* g_logFile;
extern CRITICAL_SECTION g_logCS;
extern HANDLE g_hPipe;
extern bool g_pipeConnected;

void PipeSend(const char* data, int len);
void Log(const wchar_t* fmt, ...);
void PipeSendTopology(const wchar_t* xmlPath);
void PipeSendStatus(int total, int identified, int events);
bool PipeCheckStop();
std::wstring LogPath(const std::wstring& logDir, const wchar_t* filename);
