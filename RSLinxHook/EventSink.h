#pragma once
#include "RSLinxHook_fwd.h"
#include "ComInterfaces.h"

// ============================================================
// DualEventSink — COM event sink for browse callbacks
// Globals defined in EventSink.cpp
// ============================================================

extern std::vector<std::wstring> g_discoveredDevices;
extern std::vector<IUnknown*> g_capturedBuses;
extern volatile bool g_captureBuses;

// ============================================================
// Dual-interface event sink with FTM + padding.
//
// Key design points:
// 1. Dual inheritance gives correct vtables for both CPs
// 2. Accept-all QI lets CP[1] and CP[2] connect (required for browse)
// 3. FTM (Free Threaded Marshaler) enables cross-apartment event
//    delivery without proxy/stubs
// 4. Padding at +8 accommodates Start()'s direct memory writes
//    (Start writes up to offset ~68 from object start)
//
// Object layout:
//   +0:  vtable_1 (IRSTopologyOnlineNotify*)  [4 bytes]
//   +4:  vtable_2 (ITopologyBusEvents*)       [4 bytes]
//   +8:  m_pad[128] - absorbs Start()'s direct writes
//   +136: m_refCount (safe from Start)
//   +140: m_pFTM (IMarshal* for FTM)
//   +144: m_magic
// ============================================================

class DualEventSink : public IRSTopologyOnlineNotify, public ITopologyBusEvents
{
public:
    BYTE m_pad[2048];           // +8: absorbs Start()'s writes (observed up to ~+528)
    LONG m_refCount;            // +136: safe from Start
    IUnknown* m_pFTM;           // +140: Free Threaded Marshaler
    DWORD m_magic;              // +144
    std::wstring m_label;       // identity for log messages (e.g., "Test/Ethernet", "5069-L320ER/Backplane")

    // Cycle detection: track seen addresses to detect when browse repeats
    CRITICAL_SECTION m_cs;
    std::set<std::wstring> m_seenAddresses;
    std::vector<std::wstring> m_addressesInOrder;  // Preserves discovery order for address-device correlation
    volatile bool m_cycleComplete;  // true when repeat address, BrowseCycled, or BrowseEnded
    volatile bool m_browseEnded;    // true when BrowseEnded fires
    int m_addressCount;             // total Found() calls

    DualEventSink(const wchar_t* label = L"");
    ~DualEventSink();

    void DumpCounters(const wchar_t* label);
    void DumpDWords(const wchar_t* label, int count = 32);

    // --- IUnknown ---
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;

    // --- IRSTopologyOnlineNotify methods ---
    STDMETHODIMP BrowseStarted(IUnknown* pBus) override;
    STDMETHODIMP BrowseCycled(IUnknown* pBus) override;
    STDMETHODIMP BrowseEnded(IUnknown* pBus) override;
    STDMETHODIMP Found(IUnknown* pBus, VARIANT addr) override;
    STDMETHODIMP NothingAtAddress(IUnknown* pBus, VARIANT addr) override;

    // --- ITopologyBusEvents methods ---
    STDMETHODIMP OnPortConnect(IUnknown*, IUnknown*, VARIANT) override;
    STDMETHODIMP OnPortDisconnect(IUnknown*, IUnknown*, VARIANT) override;
    STDMETHODIMP OnPortChangeAddress(IUnknown*, IUnknown*, VARIANT, VARIANT) override;
    STDMETHODIMP OnPortChangeState(IUnknown*, long) override;
    STDMETHODIMP OnBrowseStarted(IUnknown* pBus) override;
    STDMETHODIMP OnBrowseCycled(IUnknown*) override;
    STDMETHODIMP OnBrowseEnded(IUnknown*) override;
    STDMETHODIMP OnBrowseAddressFound(IUnknown*, VARIANT addr) override;
    STDMETHODIMP OnBrowseAddressNotFound(IUnknown*, VARIANT) override;
};
