#pragma once
#include "RSLinxHook_fwd.h"
#include "Config.h"

// ============================================================
// Main-STA thread hook mechanism
// Globals defined in STAHook.cpp
// ============================================================

extern MainSTAFunc g_pMainSTAFunc;
extern HHOOK g_hHook;
extern DWORD g_mainThreadId;
extern volatile LONG g_browseRequested;
extern volatile LONG g_browseResult;
extern HookConfig* g_pSharedConfig;
extern WNDPROC g_origWndProc;

DWORD FindMainThreadId();
HRESULT ExecuteOnMainSTA(MainSTAFunc func);

// Callbacks (accessible for hook registration)
LRESULT CALLBACK MainSTAHookProc(int code, WPARAM wp, LPARAM lp);
LRESULT CALLBACK HookWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
BOOL CALLBACK FindThreadWindowProc(HWND hwnd, LPARAM lParam);
