#pragma once
#include "RSLinxHook_fwd.h"

// ============================================================
// COM GUIDs (extern declarations — defined in ComInterfaces.cpp)
// ============================================================

extern const GUID CLSID_HarmonyServices;
extern const GUID CLSID_RSTopologyGlobals;
extern const GUID CLSID_RSProjectGlobal;
extern const GUID CLSID_OnlineBusExt;
extern const GUID CLSID_RSPath;
extern const GUID CLSID_EthernetPort;
extern const GUID CLSID_EthernetBus;

extern const GUID IID_IHarmonyConnector;
extern const GUID IID_IRSTopologyGlobals;
extern const GUID IID_IRSProjectGlobal;
extern const GUID IID_IRSProject;
extern const GUID IID_ITopologyBus;
extern const GUID IID_ITopologyBusEvents;
extern const GUID IID_ITopologyPathComposer;
extern const GUID IID_ITopologyDevice_Dual;
extern const GUID IID_IOnlineEnumeratorTypeLib;
extern const GUID IID_IOnlineEnumerator;
extern const GUID IID_IRSTopologyNetwork;
extern const GUID IID_IRSTopologyObject;
extern const GUID IID_IRSTopologyDevice;
extern const GUID IID_ITopologyCollection;
extern const GUID IID_ITopologyChassis;
extern const GUID IID_ITopologyObject;
extern const GUID IID_IRSTopologyPort;
extern const GUID IID_IRSObject;
extern const GUID IID_IRSTopologyOnlineNotify;

// ============================================================
// Interface declarations (minimal)
// ============================================================

#undef INTERFACE
#define INTERFACE IHarmonyConnector
DECLARE_INTERFACE_(IHarmonyConnector, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(SetServerOptions)(THIS_ DWORD options, LPCWSTR servername) PURE;
    STDMETHOD(GetSpecialObject)(THIS_ const GUID* clsID, const GUID* iid, IUnknown** ppObject) PURE;
};

#undef INTERFACE
#define INTERFACE IRSTopologyGlobals
DECLARE_INTERFACE_(IRSTopologyGlobals, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(GetClassFromPCCCReply)(THIS_ void* pReply, GUID* pClassID) PURE;
    STDMETHOD(GetClassFromASAKey)(THIS_ void* pASAKey, GUID* pClassID) PURE;
    STDMETHOD(FindSinglePath)(THIS_ void* a, void* b, void** c) PURE;
    STDMETHOD(FindPartialPath)(THIS_ void* a, void* b, void** c) PURE;
    STDMETHOD(AddToPath)(THIS_ void* a, LPCWSTR b, void* c, void** d) PURE;
    STDMETHOD(ReleasePATHPART)(THIS_ void* a) PURE;
    STDMETHOD(GetThisWorkstationObject)(THIS_ void* container, IUnknown** workstation) PURE;
};

#undef INTERFACE
#define INTERFACE IRSProjectGlobal
DECLARE_INTERFACE_(IRSProjectGlobal, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(OpenProject)(THIS_ LPCWSTR name, DWORD flags, HWND hWnd, void* pApp, const GUID* iid, void** pp) PURE;
};

#undef INTERFACE
#define INTERFACE IRSProject
DECLARE_INTERFACE_(IRSProject, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(BindToObject)(THIS_ LPCWSTR path, IUnknown** ppObject) PURE;
};

// ============================================================
// ITopologyBusEvents interface
// ============================================================

#undef INTERFACE
#define INTERFACE ITopologyBusEvents
DECLARE_INTERFACE_(ITopologyBusEvents, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(OnPortConnect)(THIS_ IUnknown*, IUnknown*, VARIANT) PURE;
    STDMETHOD(OnPortDisconnect)(THIS_ IUnknown*, IUnknown*, VARIANT) PURE;
    STDMETHOD(OnPortChangeAddress)(THIS_ IUnknown*, IUnknown*, VARIANT, VARIANT) PURE;
    STDMETHOD(OnPortChangeState)(THIS_ IUnknown*, long) PURE;
    STDMETHOD(OnBrowseStarted)(THIS_ IUnknown*) PURE;
    STDMETHOD(OnBrowseCycled)(THIS_ IUnknown*) PURE;
    STDMETHOD(OnBrowseEnded)(THIS_ IUnknown*) PURE;
    STDMETHOD(OnBrowseAddressFound)(THIS_ IUnknown*, VARIANT) PURE;
    STDMETHOD(OnBrowseAddressNotFound)(THIS_ IUnknown*, VARIANT) PURE;
};

// ============================================================
// IRSTopologyOnlineNotify interface (browse event callback)
// ============================================================
// Vtable layout (from analysis):
//   [0] QI  [1] AddRef  [2] Release
//   [3] BrowseStarted(IUnknown* pBus)
//   [4] BrowseCycled(IUnknown* pBus)
//   [5] BrowseEnded(IUnknown* pBus)
//   [6] Found(IUnknown* pBus, VARIANT addr)
//   [7] NothingAtAddress(IUnknown* pBus, VARIANT addr)

#undef INTERFACE
#define INTERFACE IRSTopologyOnlineNotify
DECLARE_INTERFACE_(IRSTopologyOnlineNotify, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(BrowseStarted)(THIS_ IUnknown* pBus) PURE;
    STDMETHOD(BrowseCycled)(THIS_ IUnknown* pBus) PURE;
    STDMETHOD(BrowseEnded)(THIS_ IUnknown* pBus) PURE;
    STDMETHOD(Found)(THIS_ IUnknown* pBus, VARIANT addr) PURE;
    STDMETHOD(NothingAtAddress)(THIS_ IUnknown* pBus, VARIANT addr) PURE;
};
