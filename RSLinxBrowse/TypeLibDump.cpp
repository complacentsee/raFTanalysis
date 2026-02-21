/**
 * TypeLibDump.cpp
 *
 * Standalone tool to extract and dump TypeLib information from RSLinx DLLs.
 * This reveals the exact interface definitions, method signatures, and DISPIDs
 * for RSLinx topology browsing.
 *
 * Build: cl /EHsc TypeLibDump.cpp ole32.lib oleaut32.lib
 * Run: TypeLibDump.exe
 */

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <windows.h>
#include <oleauto.h>
#include <objbase.h>
#include <iostream>
#include <iomanip>
#include <string>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

std::wstring VarTypeToString(VARTYPE vt)
{
    switch (vt & ~VT_BYREF & ~VT_ARRAY)
    {
        case VT_EMPTY:    return L"void";
        case VT_I2:       return L"short";
        case VT_I4:       return L"long";
        case VT_R4:       return L"float";
        case VT_R8:       return L"double";
        case VT_BSTR:     return L"BSTR";
        case VT_BOOL:     return L"BOOL";
        case VT_VARIANT:  return L"VARIANT";
        case VT_DISPATCH: return L"IDispatch*";
        case VT_UNKNOWN:  return L"IUnknown*";
        case VT_HRESULT:  return L"HRESULT";
        case VT_VOID:     return L"void";
        case VT_UI4:      return L"ULONG";
        case VT_UI2:      return L"USHORT";
        case VT_INT:      return L"int";
        case VT_UINT:     return L"UINT";
        case VT_LPSTR:    return L"LPSTR";
        case VT_LPWSTR:   return L"LPWSTR";
        case VT_PTR:      return L"PTR";
        case VT_USERDEFINED: return L"USERDEFINED";
        default:
        {
            wchar_t buf[32];
            swprintf_s(buf, L"VT_%d", vt);
            return buf;
        }
    }
}

std::wstring InvKindToString(INVOKEKIND invKind)
{
    switch (invKind)
    {
        case INVOKE_FUNC:         return L"METHOD";
        case INVOKE_PROPERTYGET:  return L"PROPGET";
        case INVOKE_PROPERTYPUT:  return L"PROPPUT";
        case INVOKE_PROPERTYPUTREF: return L"PROPPUTREF";
        default:                  return L"UNKNOWN";
    }
}

std::wstring TypeKindToString(TYPEKIND tk)
{
    switch (tk)
    {
        case TKIND_ENUM:     return L"ENUM";
        case TKIND_RECORD:   return L"STRUCT";
        case TKIND_MODULE:   return L"MODULE";
        case TKIND_INTERFACE: return L"INTERFACE";
        case TKIND_DISPATCH: return L"DISPATCH";
        case TKIND_COCLASS:  return L"COCLASS";
        case TKIND_ALIAS:    return L"ALIAS";
        case TKIND_UNION:    return L"UNION";
        default:             return L"OTHER";
    }
}

void DumpTypeInfo(ITypeInfo* pTypeInfo, int indent = 0)
{
    TYPEATTR* pTypeAttr = nullptr;
    HRESULT hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr)) return;

    BSTR bstrName = nullptr;
    pTypeInfo->GetDocumentation(-1, &bstrName, nullptr, nullptr, nullptr);

    std::wstring prefix(indent * 2, L' ');

    // GUID
    wchar_t guidStr[64];
    StringFromGUID2(pTypeAttr->guid, guidStr, 64);

    std::wcout << prefix << TypeKindToString(pTypeAttr->typekind) << L" "
               << (bstrName ? bstrName : L"<unnamed>")
               << L"  " << guidStr << std::endl;

    if (bstrName) SysFreeString(bstrName);

    // For interfaces/dispatches, dump methods and properties
    if (pTypeAttr->typekind == TKIND_DISPATCH || pTypeAttr->typekind == TKIND_INTERFACE)
    {
        std::wcout << prefix << L"  Functions: " << pTypeAttr->cFuncs
                   << L"  Vars: " << pTypeAttr->cVars << std::endl;

        for (UINT i = 0; i < pTypeAttr->cFuncs; i++)
        {
            FUNCDESC* pFuncDesc = nullptr;
            hr = pTypeInfo->GetFuncDesc(i, &pFuncDesc);
            if (FAILED(hr)) continue;

            BSTR funcName = nullptr;
            UINT nameCount = 0;
            BSTR paramNames[32] = {};
            pTypeInfo->GetNames(pFuncDesc->memid, paramNames, 32, &nameCount);
            funcName = (nameCount > 0) ? paramNames[0] : nullptr;

            // Return type
            std::wstring retType = VarTypeToString(pFuncDesc->elemdescFunc.tdesc.vt);

            std::wcout << prefix << L"  [" << std::setw(3) << i << L"] "
                       << InvKindToString(pFuncDesc->invkind)
                       << L" DISPID=" << pFuncDesc->memid
                       << L" vtable=" << pFuncDesc->oVft / 4  // vtable slot (4 bytes per ptr on x86)
                       << L"  " << retType << L" "
                       << (funcName ? funcName : L"<unnamed>") << L"(";

            // Parameters
            for (SHORT p = 0; p < pFuncDesc->cParams; p++)
            {
                if (p > 0) std::wcout << L", ";

                BSTR paramName = (p + 1 < (SHORT)nameCount) ? paramNames[p + 1] : nullptr;
                std::wstring paramType = VarTypeToString(pFuncDesc->lprgelemdescParam[p].tdesc.vt);

                // Check for pointer types
                if (pFuncDesc->lprgelemdescParam[p].tdesc.vt == VT_PTR)
                {
                    if (pFuncDesc->lprgelemdescParam[p].tdesc.lptdesc)
                    {
                        paramType = VarTypeToString(pFuncDesc->lprgelemdescParam[p].tdesc.lptdesc->vt) + L"*";
                    }
                }

                // Check for [in], [out], [retval] flags
                USHORT flags = pFuncDesc->lprgelemdescParam[p].paramdesc.wParamFlags;
                std::wstring flagStr;
                if (flags & PARAMFLAG_FIN) flagStr += L"in";
                if (flags & PARAMFLAG_FOUT) { if (!flagStr.empty()) flagStr += L","; flagStr += L"out"; }
                if (flags & PARAMFLAG_FRETVAL) { if (!flagStr.empty()) flagStr += L","; flagStr += L"retval"; }

                if (!flagStr.empty()) std::wcout << L"[" << flagStr << L"] ";
                std::wcout << paramType;
                if (paramName) std::wcout << L" " << paramName;
            }

            std::wcout << L")" << std::endl;

            // Free names
            for (UINT n = 0; n < nameCount; n++)
                SysFreeString(paramNames[n]);

            pTypeInfo->ReleaseFuncDesc(pFuncDesc);
        }
    }

    // For coclasses, list implemented interfaces
    if (pTypeAttr->typekind == TKIND_COCLASS)
    {
        for (UINT i = 0; i < pTypeAttr->cImplTypes; i++)
        {
            HREFTYPE hRefType;
            hr = pTypeInfo->GetRefTypeOfImplType(i, &hRefType);
            if (FAILED(hr)) continue;

            ITypeInfo* pImplType = nullptr;
            hr = pTypeInfo->GetRefTypeInfo(hRefType, &pImplType);
            if (FAILED(hr)) continue;

            BSTR implName = nullptr;
            pImplType->GetDocumentation(-1, &implName, nullptr, nullptr, nullptr);

            int implFlags = 0;
            pTypeInfo->GetImplTypeFlags(i, &implFlags);

            std::wcout << prefix << L"  Implements: " << (implName ? implName : L"?");
            if (implFlags & IMPLTYPEFLAG_FDEFAULT) std::wcout << L" [default]";
            if (implFlags & IMPLTYPEFLAG_FSOURCE) std::wcout << L" [source]";
            std::wcout << std::endl;

            if (implName) SysFreeString(implName);
            pImplType->Release();
        }
    }

    pTypeInfo->ReleaseTypeAttr(pTypeAttr);
}

void DumpTypeLib(const wchar_t* path, const wchar_t* label)
{
    std::wcout << std::endl;
    std::wcout << L"================================================================" << std::endl;
    std::wcout << L"TypeLib: " << label << std::endl;
    std::wcout << L"Path: " << path << std::endl;
    std::wcout << L"================================================================" << std::endl;

    ITypeLib* pTypeLib = nullptr;
    HRESULT hr = LoadTypeLib(path, &pTypeLib);
    if (FAILED(hr))
    {
        std::wcerr << L"[ERROR] LoadTypeLib failed: hr=0x" << std::hex << hr << std::dec << std::endl;
        return;
    }

    UINT count = pTypeLib->GetTypeInfoCount();
    std::wcout << L"Type count: " << count << std::endl;

    BSTR libName = nullptr;
    BSTR libDoc = nullptr;
    pTypeLib->GetDocumentation(-1, &libName, &libDoc, nullptr, nullptr);
    if (libName) { std::wcout << L"Name: " << libName << std::endl; SysFreeString(libName); }
    if (libDoc) { std::wcout << L"Doc: " << libDoc << std::endl; SysFreeString(libDoc); }

    std::wcout << std::endl;

    for (UINT i = 0; i < count; i++)
    {
        ITypeInfo* pTypeInfo = nullptr;
        hr = pTypeLib->GetTypeInfo(i, &pTypeInfo);
        if (FAILED(hr)) continue;

        DumpTypeInfo(pTypeInfo, 0);
        std::wcout << std::endl;

        pTypeInfo->Release();
    }

    pTypeLib->Release();
}

int wmain()
{
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    // Dump RSTOP.DLL TypeLib (Harmony Topology Services)
    DumpTypeLib(
        L"C:\\Program Files (x86)\\Rockwell Software\\RSCommon\\RSTOP.DLL",
        L"Harmony Topology Services (RSTOP.DLL)");

    // Dump RSWHO.OCX TypeLib
    DumpTypeLib(
        L"C:\\Program Files (x86)\\Rockwell Software\\RSCommon\\RSWHO.OCX",
        L"RSWho Control (RSWHO.OCX)");

    // Also dump LINXCOMM.DLL if it has a TypeLib
    DumpTypeLib(
        L"C:\\Program Files (x86)\\Rockwell Software\\RSLinx\\LINXCOMM.DLL",
        L"LinxComm (LINXCOMM.DLL)");

    CoUninitialize();
    return 0;
}
