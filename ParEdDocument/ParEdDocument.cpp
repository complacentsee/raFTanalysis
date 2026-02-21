// ParEdTypeInfo.cpp
// Build x86: cl /EHsc /W4 /nologo ParEdTypeInfo.cpp ole32.lib oleaut32.lib shlwapi.lib

#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <propsys.h>     // IInitializeWithStream
#include <shlwapi.h>     // SHCreateStreamOnFileEx
#include <cstdio>
#include <cstdint>
#include <vector>
#include <string>
#include <ocidl.h>       // IProvideClassInfo
#include <conio.h>

#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "OleAut32.lib")

static void PrintHresult(const char* label, HRESULT hr) {
    std::printf("%s: 0x%08lX (%ld)\n", label, hr, (long)hr);
}

static const char* InvKindToStr(INVOKEKIND k) {
    switch (k) {
    case INVOKE_FUNC: return "FUNC";
    case INVOKE_PROPERTYGET: return "PROPGET";
    case INVOKE_PROPERTYPUT: return "PROPPUT";
    case INVOKE_PROPERTYPUTREF: return "PROPPUTREF";
    default: return "INVKIND?";
    }
}

static const char* VarTypeToStr(VARTYPE vt) {
    switch (vt) {
    case VT_EMPTY: return "VT_EMPTY";
    case VT_NULL: return "VT_NULL";
    case VT_I2: return "VT_I2";
    case VT_I4: return "VT_I4";
    case VT_R4: return "VT_R4";
    case VT_R8: return "VT_R8";
    case VT_CY: return "VT_CY";
    case VT_DATE: return "VT_DATE";
    case VT_BSTR: return "VT_BSTR";
    case VT_DISPATCH: return "VT_DISPATCH";
    case VT_ERROR: return "VT_ERROR";
    case VT_BOOL: return "VT_BOOL";
    case VT_VARIANT: return "VT_VARIANT";
    case VT_UNKNOWN: return "VT_UNKNOWN";
    case VT_UI1: return "VT_UI1";
    case VT_UI2: return "VT_UI2";
    case VT_UI4: return "VT_UI4";
    case VT_I8: return "VT_I8";
    case VT_UI8: return "VT_UI8";
    case VT_INT: return "VT_INT";
    case VT_UINT: return "VT_UINT";
    case VT_VOID: return "VT_VOID";
    case VT_HRESULT: return "VT_HRESULT";
    case VT_PTR: return "VT_PTR";
    case VT_SAFEARRAY: return "VT_SAFEARRAY";
    case VT_CARRAY: return "VT_CARRAY";
    case VT_USERDEFINED: return "VT_USERDEFINED";
    default: return "VT_???";
    }
}

static void PrintTypeDesc(ITypeInfo* pti, const TYPEDESC& td);

static void PrintUserDefined(ITypeInfo* pti, HREFTYPE href) {
    ITypeInfo* ref = nullptr;
    if (FAILED(pti->GetRefTypeInfo(href, &ref)) || !ref) {
        std::printf("USERDEF(<GetRefTypeInfo failed>)");
        return;
    }

    BSTR name = nullptr;
    if (SUCCEEDED(ref->GetDocumentation(MEMBERID_NIL, &name, nullptr, nullptr, nullptr)) && name) {
        // BSTR is wide; print as UTF-8-ish best-effort
        std::wstring ws(name, SysStringLen(name));
        std::string s(ws.begin(), ws.end());
        std::printf("USERDEF(%s)", s.c_str());
        SysFreeString(name);
    }
    else {
        std::printf("USERDEF(<unnamed>)");
    }
    ref->Release();
}

static void PrintTypeDesc(ITypeInfo* pti, const TYPEDESC& td) {
    switch (td.vt) {
    case VT_PTR:
        std::printf("PTR(");
        PrintTypeDesc(pti, *td.lptdesc);
        std::printf(")");
        break;
    case VT_SAFEARRAY:
        std::printf("SAFEARRAY(");
        PrintTypeDesc(pti, *td.lptdesc);
        std::printf(")");
        break;
    case VT_USERDEFINED:
        PrintUserDefined(pti, td.hreftype);
        break;
    default:
        std::printf("%s", VarTypeToStr(td.vt));
        break;
    }
}

static void PrintElemDesc(ITypeInfo* pti, const ELEMDESC& ed) {
    PrintTypeDesc(pti, ed.tdesc);

    // Basic param flags (optional)
    if (ed.paramdesc.wParamFlags) {
        std::printf(" [");
        bool first = true;
        auto flag = [&](USHORT f, const char* n) {
            if (ed.paramdesc.wParamFlags & f) {
                if (!first) std::printf("|");
                std::printf("%s", n);
                first = false;
            }
        };
        flag(PARAMFLAG_FIN, "in");
        flag(PARAMFLAG_FOUT, "out");
        flag(PARAMFLAG_FRETVAL, "retval");
        flag(PARAMFLAG_FOPT, "opt");
        std::printf("]");
    }
}

static void DumpTypeInfo(ITypeInfo* pti) {
    TYPEATTR* ta = nullptr;
    HRESULT hr = pti->GetTypeAttr(&ta);
    if (FAILED(hr) || !ta) {
        PrintHresult("GetTypeAttr failed", hr);
        return;
    }

    std::printf("TypeAttr: cFuncs=%u cVars=%u typekind=%u\n",
        (unsigned)ta->cFuncs, (unsigned)ta->cVars, (unsigned)ta->typekind);

    for (UINT i = 0; i < (UINT)ta->cFuncs; i++) {
        FUNCDESC* fd = nullptr;
        hr = pti->GetFuncDesc(i, &fd);
        if (FAILED(hr) || !fd) {
            std::printf("  GetFuncDesc(%u) failed: 0x%08lX\n", i, (long)hr);
            continue;
        }

        BSTR bName = nullptr;
        BSTR bDoc = nullptr;
        hr = pti->GetDocumentation(fd->memid, &bName, &bDoc, nullptr, nullptr);

        std::wstring wname = (bName ? std::wstring(bName, SysStringLen(bName)) : L"<noname>");
        std::string name(wname.begin(), wname.end());

        std::printf("\n[%u] %s  memid=0x%08lX  inv=%s  cParams=%u  ret=",
            i,
            name.c_str(),
            (long)fd->memid,
            InvKindToStr(fd->invkind),
            (unsigned)fd->cParams);

        PrintElemDesc(pti, fd->elemdescFunc);
        std::printf("\n");

        for (UINT p = 0; p < (UINT)fd->cParams; p++) {
            std::printf("    arg[%u]: ", p);
            PrintElemDesc(pti, fd->lprgelemdescParam[p]);
            std::printf("\n");
        }

        if (bName) SysFreeString(bName);
        if (bDoc) SysFreeString(bDoc);

        pti->ReleaseFuncDesc(fd);
    }

    pti->ReleaseTypeAttr(ta);
}

// Try invoke by name; supports 0-arg or 1-arg(BSTR) calls.
// If it finds multiple overload-ish entries, it tries the first matching arity.
static HRESULT TryInvokeByName(
    IDispatch* disp,
    ITypeInfo* pti,
    const wchar_t* methodName,          // <-- const now
    const wchar_t* maybePathBstr
) {
    if (!disp || !pti || !methodName) return E_INVALIDARG;

    // Resolve name -> MEMBERID using the typeinfo to avoid locale/name mangling issues.
    MEMBERID memid = MEMBERID_NIL;

    // GetIDsOfNames requires LPOLESTR* (mutable). We only cast at the boundary.
    LPOLESTR namesTI[1] = { const_cast<LPOLESTR>(methodName) };
    HRESULT hr = pti->GetIDsOfNames(namesTI, 1, &memid);
    if (FAILED(hr)) {
        // Fallback: IDispatch lookup
        LPOLESTR namesDisp[1] = { const_cast<LPOLESTR>(methodName) };
        DISPID id = DISPID_UNKNOWN;
        hr = disp->GetIDsOfNames(IID_NULL, namesDisp, 1, LOCALE_USER_DEFAULT, &id);
        if (FAILED(hr)) return hr;
        memid = (MEMBERID)id;
    }

    // Find a FUNCDESC with matching memid (and prefer INVOKE_FUNC / PROPGET)
    TYPEATTR* ta = nullptr;
    hr = pti->GetTypeAttr(&ta);
    if (FAILED(hr) || !ta) return hr;

    FUNCDESC* chosen0 = nullptr;  // 0-arg
    FUNCDESC* chosen1b = nullptr; // 1-arg maybe bstr/variant

    for (UINT i = 0; i < (UINT)ta->cFuncs; i++) {
        FUNCDESC* fd = nullptr;
        if (FAILED(pti->GetFuncDesc(i, &fd)) || !fd) continue;

        if (fd->memid == memid && (fd->invkind == INVOKE_FUNC || fd->invkind == INVOKE_PROPERTYGET)) {
            if (fd->cParams == 0 && !chosen0) {
                chosen0 = fd;
                continue; // keep fd
            }
            else if (fd->cParams == 1 && !chosen1b) {
                chosen1b = fd;
                continue; // keep fd
            }
        }

        pti->ReleaseFuncDesc(fd);
        if (chosen0 && chosen1b) break;
    }

    pti->ReleaseTypeAttr(ta);

    auto doInvoke = [&](DISPID dispid, WORD flags, VARIANT* args, UINT cArgs) -> HRESULT {
        DISPPARAMS dp{};
        dp.cArgs = cArgs;
        dp.rgvarg = args;

        VARIANT ret{};
        VariantInit(&ret);
        EXCEPINFO ex{};
        UINT argErr = 0;

        HRESULT r = disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, flags, &dp, &ret, &ex, &argErr);

        std::printf("  Invoke(%ls) -> 0x%08lX", methodName, (long)r);
        if (FAILED(r)) {
            std::printf(" (argErr=%u)\n", argErr);
        }
        else {
            std::printf("  ret.vt=%u\n", (unsigned)ret.vt);
        }

        VariantClear(&ret);
        return r;
    };

    HRESULT best = E_FAIL;

    if (chosen0) {
        DISPID id = (DISPID)chosen0->memid;
        WORD flags = (chosen0->invkind == INVOKE_PROPERTYGET) ? DISPATCH_PROPERTYGET : DISPATCH_METHOD;
        best = doInvoke(id, flags, nullptr, 0);
        pti->ReleaseFuncDesc(chosen0);

        if (SUCCEEDED(best)) {
            if (chosen1b) pti->ReleaseFuncDesc(chosen1b);
            return best;
        }
    }

    if (chosen1b && maybePathBstr) {
        // args are passed right-to-left in rgvarg (only 1 arg here)
        VARIANT arg{};
        VariantInit(&arg);
        arg.vt = VT_BSTR;
        arg.bstrVal = SysAllocString(maybePathBstr);

        DISPID id = (DISPID)chosen1b->memid;
        WORD flags = (chosen1b->invkind == INVOKE_PROPERTYGET) ? DISPATCH_PROPERTYGET : DISPATCH_METHOD;
        best = doInvoke(id, flags, &arg, 1);

        VariantClear(&arg);
        pti->ReleaseFuncDesc(chosen1b);
        return best;
    }

    if (chosen1b) pti->ReleaseFuncDesc(chosen1b);
    return best;
}

static HRESULT GetTypeInfoViaProvideClassInfo(IDispatch* disp, ITypeInfo** outTI) {
    *outTI = nullptr;

    IProvideClassInfo* pci = nullptr;
    HRESULT hr = disp->QueryInterface(IID_IProvideClassInfo, (void**)&pci);
    PrintHresult("QI(IProvideClassInfo)", hr);
    if (FAILED(hr) || !pci) return hr;

    ITypeInfo* ti = nullptr;
    hr = pci->GetClassInfo(&ti);
    PrintHresult("IProvideClassInfo::GetClassInfo", hr);
    pci->Release();

    if (FAILED(hr) || !ti) return hr;

    *outTI = ti; // caller releases
    return S_OK;
}

static HRESULT GetDefaultInterfaceTypeInfoFromCoclass(ITypeInfo* coclassTI, ITypeInfo** outIfaceTI) {
    *outIfaceTI = nullptr;
    if (!coclassTI) return E_INVALIDARG;

    TYPEATTR* ta = nullptr;
    HRESULT hr = coclassTI->GetTypeAttr(&ta);
    if (FAILED(hr) || !ta) return hr;

    if (ta->typekind != TKIND_COCLASS) {
        coclassTI->ReleaseTypeAttr(ta);
        return TYPE_E_WRONGTYPEKIND;
    }

    // Iterate implemented interfaces
    for (UINT i = 0; i < (UINT)ta->cImplTypes; i++) {
        INT flags = 0;
        hr = coclassTI->GetImplTypeFlags(i, &flags);
        if (FAILED(hr)) continue;

        // We want the default interface (not the default source/event interface)
        const bool isDefault = (flags & IMPLTYPEFLAG_FDEFAULT) != 0;
        const bool isSource = (flags & IMPLTYPEFLAG_FSOURCE) != 0;

        if (!isDefault || isSource) {
            continue;
        }

        HREFTYPE href = 0;
        hr = coclassTI->GetRefTypeOfImplType(i, &href);
        if (FAILED(hr)) continue;

        ITypeInfo* iface = nullptr;
        hr = coclassTI->GetRefTypeInfo(href, &iface);
        if (SUCCEEDED(hr) && iface) {
            *outIfaceTI = iface; // caller releases
            coclassTI->ReleaseTypeAttr(ta);
            return S_OK;
        }
    }

    coclassTI->ReleaseTypeAttr(ta);
    return E_NOINTERFACE;
}

int wmain(int argc, wchar_t** argv) {
    if (argc < 2) {
        fwprintf(stderr, L"Usage: %s <path-to-par>\n", argv[0]);
        return 2;
    }
    const wchar_t* parPath = argv[1];

    std::printf("Press any key to continue...\n");
    _getch();

    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        PrintHresult("CoInitializeEx failed", hr);
        return 3;
    }

    GUID clsid{};
    hr = CLSIDFromString(L"{345F59D1-50EF-11D3-8B7F-00104B701D36}", &clsid);
    if (FAILED(hr)) {
        PrintHresult("CLSIDFromString(ParEd) failed", hr);
        CoUninitialize();
        return 4;
    }

    IDispatch* pDisp = nullptr;
    hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pDisp);
    PrintHresult("CoCreateInstance(ParEd, IID_IDispatch)", hr);
    if (FAILED(hr) || !pDisp) {
        CoUninitialize();
        return 5;
    }

    // --- Initialize from stream first ---
    {
        IInitializeWithStream* pInit = nullptr;
        hr = pDisp->QueryInterface(__uuidof(IInitializeWithStream), (void**)&pInit);
        PrintHresult("QI(IInitializeWithStream)", hr);

        if (SUCCEEDED(hr) && pInit) {
            IStream* stm = nullptr;
            hr = SHCreateStreamOnFileEx(
                parPath,
                STGM_READ | STGM_SHARE_DENY_NONE,
                FILE_ATTRIBUTE_NORMAL,
                FALSE,
                nullptr,
                &stm
            );
            PrintHresult("SHCreateStreamOnFileEx(parPath)", hr);

            if (SUCCEEDED(hr) && stm) {
                hr = pInit->Initialize(stm, STGM_READ);
                PrintHresult("IInitializeWithStream::Initialize(stream, STGM_READ)", hr);
                stm->Release();
            }

            pInit->Release();
            pInit = nullptr;
        }
    }

    // --- Acquire type info ---
    ITypeInfo* tiFromDisp = nullptr;

    UINT n = 0;
    hr = pDisp->GetTypeInfoCount(&n);
    PrintHresult("IDispatch::GetTypeInfoCount", hr);
    std::printf("TypeInfoCount = %u\n", n);

    if (SUCCEEDED(hr) && n > 0) {
        hr = pDisp->GetTypeInfo(0, LOCALE_USER_DEFAULT, &tiFromDisp);
        PrintHresult("IDispatch::GetTypeInfo(0)", hr);
    }

    if (!tiFromDisp) {
        std::printf("IDispatch::GetTypeInfo failed; trying IProvideClassInfo...\n");
        hr = GetTypeInfoViaProvideClassInfo(pDisp, &tiFromDisp);
        PrintHresult("GetTypeInfoViaProvideClassInfo", hr);
    }

    if (!tiFromDisp) {
        std::printf("No type info available via either path.\n");
        pDisp->Release();
        CoUninitialize();
        return 6;
    }

    // If we got a coclass, pivot to its default (non-source) interface.
    ITypeInfo* tiToDumpAndInvoke = tiFromDisp;

    TYPEATTR* ta = nullptr;
    hr = tiFromDisp->GetTypeAttr(&ta);
    if (SUCCEEDED(hr) && ta) {
        std::printf("TypeAttr: cFuncs=%u cVars=%u typekind=%u cImplTypes=%u\n",
            (unsigned)ta->cFuncs,
            (unsigned)ta->cVars,
            (unsigned)ta->typekind,
            (unsigned)ta->cImplTypes);

        bool isCoclass = (ta->typekind == TKIND_COCLASS);
        tiFromDisp->ReleaseTypeAttr(ta);
        ta = nullptr;

        if (isCoclass) {
            ITypeInfo* defIface = nullptr;
            hr = GetDefaultInterfaceTypeInfoFromCoclass(tiFromDisp, &defIface);
            PrintHresult("GetDefaultInterfaceTypeInfoFromCoclass", hr);

            if (SUCCEEDED(hr) && defIface) {
                tiToDumpAndInvoke = defIface; // use this instead of coclass TI
            }
            else {
                std::printf("Could not locate default interface; dumping coclass info only.\n");
            }
        }
    }

    std::printf("\n=== Dumping methods from selected typeinfo ===\n");
    DumpTypeInfo(tiToDumpAndInvoke);

    // Try calling some common method names (0-arg or 1-BSTR)
    const wchar_t* suspects[] = {
        L"AOAInitialize",
        L"AOAPrint",
        L"AOASave",
        L"AOAGetWindowStyle",
        L"Compile",
        L"Decompile"
    };

    std::printf("\n=== Attempting suspect method invokes (0-arg or 1-BSTR) ===\n");
    for (auto s : suspects) {
        HRESULT r = TryInvokeByName(pDisp, tiToDumpAndInvoke, s, parPath);
        if (SUCCEEDED(r)) {
            std::printf("  %ls: SUCCESS\n", s);
        }
    }

    // Release typeinfos
    if (tiToDumpAndInvoke && tiToDumpAndInvoke != tiFromDisp) {
        tiToDumpAndInvoke->Release();
        tiToDumpAndInvoke = nullptr;
    }
    if (tiFromDisp) {
        tiFromDisp->Release();
        tiFromDisp = nullptr;
    }

    std::printf("\n=== Trying AOASave ===\n");
    {
        // Pick an output filename (you can change extension as needed)
        std::wstring outPath = std::wstring(parPath) + L"C:\\temp\\";

        HRESULT r = TryInvokeByName(pDisp, tiToDumpAndInvoke, L"AOASave", outPath.c_str());
        PrintHresult("TryInvokeByName(AOASave, outPath)", r);
    }

    pDisp->Release();
    CoUninitialize();
    return 0;
}
