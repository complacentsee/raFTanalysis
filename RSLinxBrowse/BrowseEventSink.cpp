/**
 * BrowseEventSink.cpp
 *
 * Implementation of COM event sink for RSLinx browse operations.
 */

#include <winsock2.h>  // Must come before windows.h
#include "BrowseEventSink.h"
#include <iostream>
#include <comutil.h>

#pragma comment(lib, "comsuppw.lib")  // For _bstr_t

BrowseEventSink::BrowseEventSink()
    : m_refCount(1)
    , m_browseStarted(false)
    , m_browseEnded(false)
{
}

BrowseEventSink::~BrowseEventSink()
{
}

// ========================================
// IUnknown Implementation
// ========================================

STDMETHODIMP BrowseEventSink::QueryInterface(REFIID riid, void** ppvObject)
{
    if (ppvObject == nullptr)
        return E_INVALIDARG;

    *ppvObject = nullptr;

    if (riid == IID_IUnknown || riid == IID_ITopologyBusEvents)
    {
        *ppvObject = static_cast<ITopologyBusEvents*>(this);
        AddRef();
        return S_OK;
    }

    // Accept Ethernet driver connection point IIDs
    // These CPs deliver browse events on the standalone IOnlineEnumerator
    // CP[0] {FA5D9CF0-A259-11D1-BE10-080009DC75C8}
    // CP[1] {F0B077A1-7483-4BCE-A33C-5E27B6A5FEA1}
    // CP[2] {E548068D-20F5-48EC-8519-12BA9AC0002D}
    static const GUID IID_EthernetCP0 =
        {0xFA5D9CF0, 0xA259, 0x11D1, {0xBE, 0x10, 0x08, 0x00, 0x09, 0xDC, 0x75, 0xC8}};
    static const GUID IID_EthernetCP1 =
        {0xF0B077A1, 0x7483, 0x4BCE, {0xA3, 0x3C, 0x5E, 0x27, 0xB6, 0xA5, 0xFE, 0xA1}};
    static const GUID IID_EthernetCP2 =
        {0xE548068D, 0x20F5, 0x48EC, {0x85, 0x19, 0x12, 0xBA, 0x9A, 0xC0, 0x00, 0x2D}};

    // IRSTopologyOnlineNotify {DC9A007C-BA7F-11D0-AD5F-00C04FD915B9}
    // This is the CRITICAL event interface for browse notifications from the bus-based enumerator.
    // Methods: BrowseStarted, BrowseCycled, BrowseEnded, Found, NothingAtAddress
    // We accept this and map it to our ITopologyBusEvents vtable (compatible layout)
    static const GUID IID_IRSTopologyOnlineNotify =
        {0xDC9A007C, 0xBA7F, 0x11D0, {0xAD, 0x5F, 0x00, 0xC0, 0x4F, 0xD9, 0x15, 0xB9}};

    // Unknown CP from bus-based enumerator
    static const GUID IID_UnknownCP5 =
        {0x595E68A0, 0x61FB, 0x11D1, {0x89, 0x22, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00}};

    if (riid == IID_EthernetCP0 || riid == IID_EthernetCP1 || riid == IID_EthernetCP2 ||
        riid == IID_IRSTopologyOnlineNotify || riid == IID_UnknownCP5)
    {
        wchar_t guidStr[64];
        StringFromGUID2(riid, guidStr, 64);
        std::wcout << L"[SINK] Accepting CP QI: " << guidStr << std::endl;

        // Return our ITopologyBusEvents vtable - these event interfaces likely have
        // a compatible layout: IUnknown + event methods dispatched by DISPID
        *ppvObject = static_cast<ITopologyBusEvents*>(this);
        AddRef();
        return S_OK;
    }

    // Also accept IDispatch - some event sources use IDispatch::Invoke to fire events
    if (riid == IID_IDispatch)
    {
        std::wcout << L"[SINK] Accepting IDispatch QI" << std::endl;
        *ppvObject = static_cast<ITopologyBusEvents*>(this);
        AddRef();
        return S_OK;
    }

    // Log unknown QI requests for debugging
    wchar_t guidStr[64];
    StringFromGUID2(riid, guidStr, 64);
    std::wcout << L"[SINK] QI for unknown: " << guidStr << std::endl;

    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) BrowseEventSink::AddRef()
{
    return InterlockedIncrement(&m_refCount);
}

STDMETHODIMP_(ULONG) BrowseEventSink::Release()
{
    ULONG newCount = InterlockedDecrement(&m_refCount);
    if (newCount == 0)
    {
        delete this;
    }
    return newCount;
}

// ========================================
// ITopologyBusEvents Implementation
// ========================================

STDMETHODIMP BrowseEventSink::OnPortConnect(IUnknown* pBus, IUnknown* pPort, VARIANT addr)
{
    std::wcout << L"[EVENT] OnPortConnect: " << VariantToString(addr) << std::endl;
    return S_OK;
}

STDMETHODIMP BrowseEventSink::OnPortDisconnect(IUnknown* pBus, IUnknown* pPort, VARIANT addr)
{
    std::wcout << L"[EVENT] OnPortDisconnect: " << VariantToString(addr) << std::endl;
    return S_OK;
}

STDMETHODIMP BrowseEventSink::OnPortChangeAddress(IUnknown* pBus, IUnknown* pPort, VARIANT oldAddr, VARIANT newAddr)
{
    std::wcout << L"[EVENT] OnPortChangeAddress" << std::endl;
    return S_OK;
}

STDMETHODIMP BrowseEventSink::OnPortChangeState(IUnknown* pPort, long State)
{
    std::wcout << L"[EVENT] OnPortChangeState: " << State << std::endl;
    return S_OK;
}

STDMETHODIMP BrowseEventSink::OnBrowseStarted(IUnknown* pBus)
{
    std::wcout << L"[EVENT] Browse Started" << std::endl;
    m_browseStarted = true;
    m_browseEnded = false;

    if (m_browseStateCallback)
    {
        m_browseStateCallback(true);
    }

    return S_OK;
}

STDMETHODIMP BrowseEventSink::OnBrowseCycled(IUnknown* pBus)
{
    std::wcout << L"[EVENT] Browse Cycled (" << m_discoveredDevices.size() << L" devices so far)" << std::endl;
    return S_OK;
}

STDMETHODIMP BrowseEventSink::OnBrowseEnded(IUnknown* pBus)
{
    std::wcout << L"[EVENT] Browse Ended (total: " << m_discoveredDevices.size() << L" devices)" << std::endl;
    m_browseEnded = true;

    if (m_browseStateCallback)
    {
        m_browseStateCallback(false);
    }

    return S_OK;
}

STDMETHODIMP BrowseEventSink::OnBrowseAddressFound(IUnknown* pBus, VARIANT addr)
{
    // **THIS IS THE CRITICAL EVENT**
    // This is called for each device discovered on the network

    std::wstring address = VariantToString(addr);
    std::wcout << L"[EVENT] >> Device Found: " << address << std::endl;

    m_discoveredDevices.push_back(address);

    if (m_deviceFoundCallback)
    {
        m_deviceFoundCallback(address);
    }

    return S_OK;
}

STDMETHODIMP BrowseEventSink::OnBrowseAddressNotFound(IUnknown* pBus, VARIANT addr)
{
    std::wstring address = VariantToString(addr);
    std::wcout << L"[EVENT] No Device at: " << address << std::endl;
    return S_OK;
}

// ========================================
// Helper Methods
// ========================================

std::wstring BrowseEventSink::VariantToString(const VARIANT& var)
{
    try
    {
        if (var.vt == VT_BSTR)
        {
            return std::wstring(var.bstrVal);
        }
        else if (var.vt == VT_LPWSTR)
        {
            return std::wstring((LPCWSTR)var.byref);
        }
        else
        {
            // Try to convert to string
            VARIANT varCopy;
            VariantInit(&varCopy);
            if (SUCCEEDED(VariantChangeType(&varCopy, &var, 0, VT_BSTR)))
            {
                std::wstring result(varCopy.bstrVal);
                VariantClear(&varCopy);
                return result;
            }
        }
    }
    catch (...)
    {
        return L"<error>";
    }

    return L"<unknown>";
}
