/**
 * BrowseEventSink.h
 *
 * COM event sink for ITopologyBusEvents.
 * Receives notifications when devices are discovered during browse operations.
 *
 * This is the CRITICAL component that was missing from the Python implementation.
 * Without this event sink, discovered devices are not added to the topology.
 */

#pragma once

#include "RSLinxInterfaces.h"
#include <vector>
#include <string>
#include <functional>

/**
 * BrowseEventSink - Implements ITopologyBusEvents to receive browse notifications
 *
 * Usage:
 *   1. Create instance: BrowseEventSink* sink = new BrowseEventSink();
 *   2. Set callback: sink->SetDeviceFoundCallback(my_callback);
 *   3. Connect to bus via IConnectionPoint
 *   4. Call IOnlineEnumerator::Start() to trigger browse
 *   5. Receive OnBrowseAddressFound() events in callback
 *   6. Disconnect and release when done
 */
class BrowseEventSink : public ITopologyBusEvents
{
public:
    BrowseEventSink();
    virtual ~BrowseEventSink();

    // Callback types
    using DeviceFoundCallback = std::function<void(const std::wstring& address)>;
    using BrowseStateCallback = std::function<void(bool started)>;

    // Set callbacks
    void SetDeviceFoundCallback(DeviceFoundCallback callback) { m_deviceFoundCallback = callback; }
    void SetBrowseStateCallback(BrowseStateCallback callback) { m_browseStateCallback = callback; }

    // Get discovered devices
    const std::vector<std::wstring>& GetDiscoveredDevices() const { return m_discoveredDevices; }

    // Browse state
    bool IsBrowseStarted() const { return m_browseStarted; }
    bool IsBrowseEnded() const { return m_browseEnded; }

    // ========================================
    // IUnknown implementation
    // ========================================
    STDMETHOD(QueryInterface)(REFIID riid, void** ppvObject) override;
    STDMETHOD_(ULONG, AddRef)() override;
    STDMETHOD_(ULONG, Release)() override;

    // ========================================
    // ITopologyBusEvents implementation
    // ========================================
    STDMETHOD(OnPortConnect)(IUnknown* pBus, IUnknown* pPort, VARIANT addr) override;
    STDMETHOD(OnPortDisconnect)(IUnknown* pBus, IUnknown* pPort, VARIANT addr) override;
    STDMETHOD(OnPortChangeAddress)(IUnknown* pBus, IUnknown* pPort, VARIANT oldAddr, VARIANT newAddr) override;
    STDMETHOD(OnPortChangeState)(IUnknown* pPort, long State) override;

    // **CRITICAL EVENTS**
    STDMETHOD(OnBrowseStarted)(IUnknown* pBus) override;
    STDMETHOD(OnBrowseCycled)(IUnknown* pBus) override;
    STDMETHOD(OnBrowseEnded)(IUnknown* pBus) override;
    STDMETHOD(OnBrowseAddressFound)(IUnknown* pBus, VARIANT addr) override;  // **KEY EVENT**
    STDMETHOD(OnBrowseAddressNotFound)(IUnknown* pBus, VARIANT addr) override;

private:
    ULONG m_refCount;
    std::vector<std::wstring> m_discoveredDevices;
    bool m_browseStarted;
    bool m_browseEnded;

    DeviceFoundCallback m_deviceFoundCallback;
    BrowseStateCallback m_browseStateCallback;

    // Helper to convert VARIANT to string
    std::wstring VariantToString(const VARIANT& var);
};
