// MinimalFTArchiveDir.cpp
// Build x86: cl /EHsc /W4 MinimalFTArchiveDir.cpp ole32.lib oleaut32.lib

#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <cstdio>
#include <cwchar>

static void PrintHresult(const char* label, HRESULT hr) {
    std::printf("%s: 0x%08lX (%ld)\n", label, hr, (long)hr);
}

static void PrintFuncLocation(const char* label, void* fnPtr) {
    if (!fnPtr) {
        std::printf("%s: <null>\n", label);
        return;
    }

    HMODULE hMod = nullptr;
    if (!GetModuleHandleExA(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
        GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
        reinterpret_cast<LPCSTR>(fnPtr),
        &hMod) ||
        !hMod) {
        std::printf("%s: %p (module not found)\n", label, fnPtr);
        return;
    }

    char path[MAX_PATH]{};
    DWORD n = GetModuleFileNameA(hMod, path, MAX_PATH);

    auto base = reinterpret_cast<uintptr_t>(hMod);
    auto addr = reinterpret_cast<uintptr_t>(fnPtr);
    auto rva = addr - base;

    std::printf("%s: %p  moduleBase=%p  RVA=0x%08lX  module=%s\n",
        label,
        fnPtr,
        hMod,
        static_cast<unsigned long>(rva),
        (n ? path : "<unknown>"));
}

static bool EnsureDirExists(const wchar_t* path) {
    if (!path || !*path) return false;

    if (CreateDirectoryW(path, nullptr)) return true;

    DWORD e = GetLastError();
    if (e == ERROR_ALREADY_EXISTS) return true;

    std::fwprintf(stderr, L"CreateDirectoryW failed (%lu) for: %ls\n", e, path);
    return false;
}

int wmain(int argc, wchar_t** argv) {
    if (argc < 3) {
        std::fwprintf(stderr, L"Usage: %s <path-to-mer> <extraction-dir>\n", argv[0]);
        return 2;
    }

    const wchar_t* merPathW = argv[1];
    const wchar_t* outDirW = argv[2];

    GUID clsidFTArchiveDir{};
    GUID iidIFTArchiveDir{};

    HRESULT hr = CLSIDFromString(L"{BE87C5E3-E3CB-4BAB-8427-578ECCE263F7}", &clsidFTArchiveDir);
    if (FAILED(hr)) {
        PrintHresult("CLSIDFromString(FTArchiveDir CLSID) failed", hr);
        return 3;
    }

    hr = CLSIDFromString(L"{49AC140C-83C5-445F-A4E2-786C490B8FFC}", &iidIFTArchiveDir);
    if (FAILED(hr)) {
        PrintHresult("CLSIDFromString(IID_IFTArchiveDir) failed", hr);
        return 4;
    }

    hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        PrintHresult("CoInitializeEx failed", hr);
        return 5;
    }

    void* pIf = nullptr;
    hr = CoCreateInstance(clsidFTArchiveDir, nullptr, CLSCTX_INPROC_SERVER, iidIFTArchiveDir, &pIf);
    std::printf("IFTArchiveDir object ptr: %p\n", pIf);
    PrintHresult("CoCreateInstance(FTArchiveDir)", hr);

    if (FAILED(hr) || !pIf) {
        CoUninitialize();
        return 6;
    }

    void** vtbl = *(void***)pIf;
    std::printf("vtbl ptr: %p\n", vtbl);

    // Standard IUnknown slots
    PrintFuncLocation("vtbl[0] QueryInterface", vtbl[0]);
    PrintFuncLocation("vtbl[1] AddRef", vtbl[1]);
    PrintFuncLocation("vtbl[2] Release", vtbl[2]);


    const int kSlot_ArchiveOp = 5;   // QueryInterface + 0x14
    const int kSlot_SetFlags = 13;  // QueryInterface + 0x34

    PrintFuncLocation("vtbl[5]  ArchiveOp (QI+0x14)", vtbl[kSlot_ArchiveOp]);
    PrintFuncLocation("vtbl[13] SetFlags  (QI+0x34)", vtbl[kSlot_SetFlags]);

    //   SetFlags(this, 0xC, 3, 0, 0)
    using FnSetFlags = HRESULT(__stdcall*)(void* thisPtr, ULONG a, ULONG b, ULONG c, ULONG d);

    //   ArchiveOp(this, bstrExtractionPath, bstrMerFileName)
    using FnArchiveOp = HRESULT(__stdcall*)(void* thisPtr, BSTR extractionPath, BSTR merNameOrPath);

    auto SetFlags = reinterpret_cast<FnSetFlags>(vtbl[kSlot_SetFlags]);
    auto ArchiveOp = reinterpret_cast<FnArchiveOp>(vtbl[kSlot_ArchiveOp]);

    // Call SetFlags with the same values you showed
    const ULONG uVar18 = 0xC;
    const ULONG uVar19 = 3;
    const ULONG uVar20 = 0;
    const ULONG uVar21 = 0;

    hr = SetFlags(pIf, uVar18, uVar19, uVar20, uVar21);
    PrintHresult("SetFlags(0xC,3,0,0)", hr);
    if (FAILED(hr)) {
        using FnRelease = ULONG(__stdcall*)(void*);
        reinterpret_cast<FnRelease>(vtbl[2])(pIf);
        CoUninitialize();
        return 7;
    }

    // Ensure output dir exists (your decompile bails if CreateDirectoryW fails)
    if (!EnsureDirExists(outDirW)) {
        using FnRelease = ULONG(__stdcall*)(void*);
        reinterpret_cast<FnRelease>(vtbl[2])(pIf);
        CoUninitialize();
        return 8;
    }

    // Allocate BSTRs like the app does
    BSTR bstrOut = SysAllocString(outDirW);
    if (!bstrOut) {
        std::fprintf(stderr, "SysAllocString(outDir) failed\n");
        using FnRelease = ULONG(__stdcall*)(void*);
        reinterpret_cast<FnRelease>(vtbl[2])(pIf);
        CoUninitialize();
        return 9;
    }

    BSTR bstrMer = SysAllocString(merPathW);
    if (!bstrMer) {
        std::fprintf(stderr, "SysAllocString(merPath) failed\n");
        SysFreeString(bstrOut);
        using FnRelease = ULONG(__stdcall*)(void*);
        reinterpret_cast<FnRelease>(vtbl[2])(pIf);
        CoUninitialize();
        return 10;
    }

    hr = ArchiveOp(pIf, bstrOut, bstrMer);
    PrintHresult("ArchiveOp(outDir, merPathOrName)", hr);

    SysFreeString(bstrMer);
    SysFreeString(bstrOut);

    using FnRelease = ULONG(__stdcall*)(void*);
    reinterpret_cast<FnRelease>(vtbl[2])(pIf);

    CoUninitialize();
    return FAILED(hr) ? 1 : 0;
}
