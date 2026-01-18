// MinimalRsvCRC.cpp
// Build x86: cl /EHsc /W4 MinimalRsvCRC.cpp ole32.lib oleaut32.lib

#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <cstdio>
#include <stdio.h>

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
        &hMod) || !hMod) {
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

int wmain(int argc, wchar_t** argv) {
    if (argc < 2) {
        fwprintf(stderr, L"Usage: %s <path-to-mer>\n", argv[0]);
        return 2;
    }

    const wchar_t* merPathW = argv[1];

    GUID clsidRsvCRC{};
    // {D19BE1A7-1D25-40F8-A71B-3E08AC7C219D}
    HRESULT hr = CLSIDFromString(L"{D19BE1A7-1D25-40F8-A71B-3E08AC7C219D}", &clsidRsvCRC);
    if (FAILED(hr)) {
        PrintHresult("CLSIDFromString(CLSID_RsvCRC) failed", hr);
        return 3;
    }

    GUID iidIRsvCRC{};
    hr = CLSIDFromString(L"{3B206953-0FE0-4A38-8C18-DCF29B9FA7AE}", &iidIRsvCRC);
    if (FAILED(hr)) {
        PrintHresult("CLSIDFromString(IID_IRsvCRC) failed", hr);
        return 4;
    }

    hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        PrintHresult("CoInitializeEx failed", hr);
        return 5;
    }

    void* pUnk = nullptr;
    hr = CoCreateInstance(clsidRsvCRC, nullptr, CLSCTX_INPROC_SERVER, iidIRsvCRC, &pUnk);
    std::printf("IRsvCRC object ptr: %p\n", pUnk);

    void** vtbl = *(void***)pUnk;
    std::printf("vtbl ptr: %p\n", vtbl);

    // Print where key vtable entries live
    PrintFuncLocation("vtbl[0] QueryInterface", vtbl[0]);
    PrintFuncLocation("vtbl[1] AddRef", vtbl[1]);
    PrintFuncLocation("vtbl[2] Release", vtbl[2]);
    PrintFuncLocation("vtbl[3] SignMerFile", vtbl[3]);
    PrintFuncLocation("vtbl[4] CheckMer", vtbl[4]);
    PrintFuncLocation("vtbl[5] Unknown", vtbl[5]);
    PrintFuncLocation("vtbl[6] Unknown", vtbl[6]);
    PrintFuncLocation("vtbl[7] Unknown", vtbl[7]);

    using FnCheckMer = HRESULT(__stdcall*)(void* thisPtr, BSTR merPath);
    auto CheckMer = reinterpret_cast<FnCheckMer>(vtbl[4]);

    BSTR bstrMer = SysAllocString(merPathW);
    if (!bstrMer) {
        std::fprintf(stderr, "SysAllocString failed\n");
        using FnRelease = ULONG(__stdcall*)(void*);
        reinterpret_cast<FnRelease>(vtbl[2])(pUnk);
        CoUninitialize();
        return 7;
    }

    HRESULT raw = CheckMer(pUnk, bstrMer);
    std::printf("raw return = 0x%08lX (%ld)\n", raw, (long)raw);

    hr = raw;
    if (hr == 1) {
        std::printf("raw==1 -> treating as failure (matches app)\n");
        hr = E_FAIL;
    }

    PrintHresult("IRsvCRC vtbl[4](merPath)", hr);

    SysFreeString(bstrMer);

    using FnRelease = ULONG(__stdcall*)(void*);
    reinterpret_cast<FnRelease>(vtbl[2])(pUnk);

    CoUninitialize();
    return FAILED(hr) ? 1 : 0;
}
