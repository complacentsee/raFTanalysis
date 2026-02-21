/**
 * RSLinxInterfaces.h
 *
 * COM interface definitions for RSLinx Classic topology browsing.
 * These interfaces were reverse-engineered from RSLinx Classic.
 *
 * IMPORTANT: RSLinx Classic is 32-bit only. This code must be compiled
 * for Win32 (x86) platform.
 */

#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <windows.h>
#include <objbase.h>
#include <unknwn.h>

// ============================================================
// CLSIDs
// ============================================================

// {043D7910-38FB-11D2-903D-00C04FA363C1}
static const GUID CLSID_HarmonyServices =
{ 0x043D7910, 0x38FB, 0x11D2, { 0x90, 0x3D, 0x00, 0xC0, 0x4F, 0xA3, 0x63, 0xC1 } };

// {38593054-38E4-11D0-BE25-00C04FC2AA48}
static const GUID CLSID_RSTopologyGlobals =
{ 0x38593054, 0x38E4, 0x11D0, { 0xBE, 0x25, 0x00, 0xC0, 0x4F, 0xC2, 0xAA, 0x48 } };

// {C92DFEA6-1D29-11D0-AD3F-00C04FD915B9}
static const GUID CLSID_RSProjectGlobal =
{ 0xC92DFEA6, 0x1D29, 0x11D0, { 0xAD, 0x3F, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };

// {2EC6B980-C629-11D0-BDCF-080009DC75C8} - Browse a bus/network
static const GUID CLSID_OnlineBusExt =
{ 0x2EC6B980, 0xC629, 0x11D0, { 0xBD, 0xCF, 0x08, 0x00, 0x09, 0xDC, 0x75, 0xC8 } };

// ============================================================
// IIDs
// ============================================================

// {19EECB80-3868-11D2-903D-00C04FA363C1}
static const GUID IID_IHarmonyConnector =
{ 0x19EECB80, 0x3868, 0x11D2, { 0x90, 0x3D, 0x00, 0xC0, 0x4F, 0xA3, 0x63, 0xC1 } };

// {640DAC76-38E3-11D0-BE25-00C04FC2AA48}
static const GUID IID_IRSTopologyGlobals =
{ 0x640DAC76, 0x38E3, 0x11D0, { 0xBE, 0x25, 0x00, 0xC0, 0x4F, 0xC2, 0xAA, 0x48 } };

// {B286B85E-211C-11D0-AD42-00C04FD915B9}
static const GUID IID_IRSProjectGlobal =
{ 0xB286B85E, 0x211C, 0x11D0, { 0xAD, 0x42, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };

// {D61BFDA0-EEAB-11CE-B4B5-C46F03C10000}
static const GUID IID_IRSProject =
{ 0xD61BFDA0, 0xEEAB, 0x11CE, { 0xB4, 0xB5, 0xC4, 0x6F, 0x03, 0xC1, 0x00, 0x00 } };

// {FC357A88-0A98-11D1-AD78-00C04FD915B9} - OLD/WRONG IID (actually IOnlineDeviceNotify-related)
static const GUID IID_IOnlineEnumerator_OLD =
{ 0xFC357A88, 0x0A98, 0x11D1, { 0xAD, 0x78, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };

// {91748520-A51B-11D0-BDC9-080009DC75C8} - CORRECT IOnlineEnumerator (from API Monitor trace)
static const GUID IID_IOnlineEnumerator =
{ 0x91748520, 0xA51B, 0x11D0, { 0xBD, 0xC9, 0x08, 0x00, 0x09, 0xDC, 0x75, 0xC8 } };

// {EEA64C84-18D1-11d1-AD81-00C04FD915B9} - IOnlineBusFinder (from API Monitor trace)
static const GUID IID_IOnlineBusFinder =
{ 0xEEA64C84, 0x18D1, 0x11D1, { 0xAD, 0x81, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };

// {B04746EA-5C5B-11d2-904A-00C04FA363C1} - IOnlineDeviceNotify
static const GUID IID_IOnlineDeviceNotify =
{ 0xB04746EA, 0x5C5B, 0x11D2, { 0x90, 0x4A, 0x00, 0xC0, 0x4F, 0xA3, 0x63, 0xC1 } };

// {FC357A89-0A98-11d1-AD78-00C04FD915B9} - IOnlineDeviceNotify2
static const GUID IID_IOnlineDeviceNotify2 =
{ 0xFC357A89, 0x0A98, 0x11D1, { 0xAD, 0x78, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };

// {6DBDFEB8-F703-11D0-AD73-00C04FD915B9} - RSPath / ITopologyPathComposer
static const GUID CLSID_RSPath =
{ 0x6DBDFEB8, 0xF703, 0x11D0, { 0xAD, 0x73, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };

// {AFF2BF80-8D86-11D0-B77F-F87205C10000} - Browse events (SOURCE)
static const GUID IID_ITopologyBusEvents =
{ 0xAFF2BF80, 0x8D86, 0x11D0, { 0xB7, 0x7F, 0xF8, 0x72, 0x05, 0xC1, 0x00, 0x00 } };

// {25C81D16-F7BA-11D0-AD73-00C04FD915B9} - ITopologyBus (dual interface, inherits IDispatch)
static const GUID IID_ITopologyBus =
{ 0x25C81D16, 0xF7BA, 0x11D0, { 0xAD, 0x73, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };

// {EF3A0832-940C-11D1-ADAC-00C04FD915B9} - ITopologyPathComposer (dual interface, inherits IDispatch)
static const GUID IID_ITopologyPathComposer =
{ 0xEF3A0832, 0x940C, 0x11D1, { 0xAD, 0xAC, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };

// {DCEAD8E0-2E7A-11CF-B4B5-C46F03C10000} - IRSTopologyObject (dual interface, has GetPath)
static const GUID IID_IRSTopologyObject =
{ 0xDCEAD8E0, 0x2E7A, 0x11CF, { 0xB4, 0xB5, 0xC4, 0x6F, 0x03, 0xC1, 0x00, 0x00 } };

// {FC357A88-0A98-11D1-AD78-00C04FD915B9} - IOnlineEnumerator from TypeLib (dual, inherits IDispatch)
static const GUID IID_IOnlineEnumeratorTypeLib =
{ 0xFC357A88, 0x0A98, 0x11D1, { 0xAD, 0x78, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };

// {46DAD8E4-4048-11D0-BE26-00C04FC2AA48}
static const GUID IID_IRSTopologyNetwork =
{ 0x46DAD8E4, 0x4048, 0x11D0, { 0xBE, 0x26, 0x00, 0xC0, 0x4F, 0xC2, 0xAA, 0x48 } };

// {DCEAD8E1-2E7A-11CF-B4B5-C46F03C10000}
static const GUID IID_IRSTopologyDevice =
{ 0xDCEAD8E1, 0x2E7A, 0x11CF, { 0xB4, 0xB5, 0xC4, 0x6F, 0x03, 0xC1, 0x00, 0x00 } };

// {98E549B2-B27E-11D0-AD5E-00C04FD915B9}
static const GUID IID_IRSTopologyPort =
{ 0x98E549B2, 0xB27E, 0x11D0, { 0xAD, 0x5E, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9 } };

// ============================================================
// Interface Declarations
// ============================================================

#undef INTERFACE
#define INTERFACE IHarmonyConnector
DECLARE_INTERFACE_(IHarmonyConnector, IUnknown)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // IHarmonyConnector
    STDMETHOD(SetServerOptions)(THIS_ DWORD options, LPCWSTR servername) PURE;
    STDMETHOD(GetSpecialObject)(THIS_ const GUID* clsID, const GUID* iid, IUnknown** ppObject) PURE;
};

#undef INTERFACE
#define INTERFACE IRSTopologyGlobals
DECLARE_INTERFACE_(IRSTopologyGlobals, IUnknown)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // IRSTopologyGlobals (partial - only methods we use)
    STDMETHOD(GetClassFromPCCCReply)(THIS_ void* pReply, GUID* pClassID) PURE;
    STDMETHOD(GetClassFromASAKey)(THIS_ void* pASAKey, GUID* pClassID) PURE;
    STDMETHOD(FindSinglePath)(THIS_ void* pIDeviceStart, void* pIDeviceEnd, void** ppIPath) PURE;
    STDMETHOD(FindPartialPath)(THIS_ void* pIDeviceStart, void* pIBusEnd, void** ppIPath) PURE;
    STDMETHOD(AddToPath)(THIS_ void* pIDeviceEnd, LPCWSTR pwszAddress, void* pIPathIn, void** ppIPathOut) PURE;
    STDMETHOD(ReleasePATHPART)(THIS_ void* pPathPart) PURE;
    STDMETHOD(GetThisWorkstationObject)(THIS_ void* container, IUnknown** workstation) PURE;
};

#undef INTERFACE
#define INTERFACE IRSProjectGlobal
DECLARE_INTERFACE_(IRSProjectGlobal, IUnknown)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // IRSProjectGlobal
    STDMETHOD(OpenProject)(THIS_ LPCWSTR pwszProjectName, DWORD flags, HWND hAppWnd, void* pApp, const GUID* iid, void** ppProject) PURE;
};

#undef INTERFACE
#define INTERFACE IRSProject
DECLARE_INTERFACE_(IRSProject, IUnknown)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // IRSProject (partial - only methods we use)
    STDMETHOD(BindToObject)(THIS_ LPCWSTR path, IUnknown** ppObject) PURE;
};

#undef INTERFACE
#define INTERFACE IOnlineEnumerator
DECLARE_INTERFACE_(IOnlineEnumerator, IUnknown)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // IOnlineEnumerator
    STDMETHOD(Start)(THIS_ LPCWSTR path) PURE;
    STDMETHOD(Stop)(THIS) PURE;
};

#undef INTERFACE
#define INTERFACE IRSTopologyNetwork
DECLARE_INTERFACE_(IRSTopologyNetwork, IUnknown)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // IRSTopologyNetwork
    STDMETHOD(Browse)(THIS_ DWORD flags) PURE;
    STDMETHOD(GetDevices)(THIS_ IUnknown** ppDevices) PURE;
    // ... other methods omitted for brevity
    STDMETHOD(ConnectNewDevice)(THIS_ DWORD flags, const GUID* deviceClassGuid, void* reserved1,
                                 LPCWSTR name, void* reserved2, LPCWSTR address) PURE;
};

#undef INTERFACE
#define INTERFACE IRSTopologyDevice
DECLARE_INTERFACE_(IRSTopologyDevice, IUnknown)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // IRSTopologyDevice
    STDMETHOD(GetPorts)(THIS_ IUnknown** ppPorts) PURE;
    STDMETHOD(GetBusses)(THIS_ IUnknown** ppBusses) PURE;
    // ... other methods omitted for brevity
};

#undef INTERFACE
#define INTERFACE IRSTopologyPort
DECLARE_INTERFACE_(IRSTopologyPort, IUnknown)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // IRSTopologyPort
    STDMETHOD(GetBus)(THIS_ IUnknown** ppBus) PURE;
    STDMETHOD(IsBackplane)(THIS_ BOOL* pbIsBackplane) PURE;
    // ... other methods omitted for brevity
};

#undef INTERFACE
#define INTERFACE ITopologyBusEvents
DECLARE_INTERFACE_(ITopologyBusEvents, IUnknown)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // ITopologyBusEvents (EVENT SINK)
    STDMETHOD(OnPortConnect)(THIS_ IUnknown* pBus, IUnknown* pPort, VARIANT addr) PURE;
    STDMETHOD(OnPortDisconnect)(THIS_ IUnknown* pBus, IUnknown* pPort, VARIANT addr) PURE;
    STDMETHOD(OnPortChangeAddress)(THIS_ IUnknown* pBus, IUnknown* pPort, VARIANT oldAddr, VARIANT newAddr) PURE;
    STDMETHOD(OnPortChangeState)(THIS_ IUnknown* pPort, long State) PURE;
    STDMETHOD(OnBrowseStarted)(THIS_ IUnknown* pBus) PURE;
    STDMETHOD(OnBrowseCycled)(THIS_ IUnknown* pBus) PURE;
    STDMETHOD(OnBrowseEnded)(THIS_ IUnknown* pBus) PURE;
    STDMETHOD(OnBrowseAddressFound)(THIS_ IUnknown* pBus, VARIANT addr) PURE;
    STDMETHOD(OnBrowseAddressNotFound)(THIS_ IUnknown* pBus, VARIANT addr) PURE;
};
