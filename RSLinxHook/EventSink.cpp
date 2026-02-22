#include "EventSink.h"
#include "Logging.h"
#include "SEHHelpers.h"

// ============================================================
// EventSink globals
// ============================================================

std::vector<std::wstring> g_discoveredDevices;
std::vector<IUnknown*> g_capturedBuses;
volatile bool g_captureBuses = false;

// ============================================================
// DualEventSink implementation
// ============================================================

DualEventSink::DualEventSink(const wchar_t* label)
    : m_refCount(1), m_pFTM(nullptr),
      m_magic(0xDEADBEEF), m_label(label ? label : L""),
      m_cycleComplete(false), m_browseEnded(false), m_addressCount(0)
{
    memset(m_pad, 0, sizeof(m_pad));
    // +8 must be non-zero for Start to succeed
    DWORD* pDW = (DWORD*)m_pad;
    pDW[0] = 1;  // fake initial value at +8

    InitializeCriticalSection(&m_cs);

    // Create Free Threaded Marshaler for cross-apartment calls
    HRESULT hr = CoCreateFreeThreadedMarshaler(
        static_cast<IRSTopologyOnlineNotify*>(this), &m_pFTM);
    if (FAILED(hr))
        m_pFTM = nullptr;
}

DualEventSink::~DualEventSink()
{
    DeleteCriticalSection(&m_cs);
    if (m_pFTM) m_pFTM->Release();
}

void DualEventSink::DumpCounters(const wchar_t* label)
{
    DWORD* pDW = (DWORD*)m_pad;
    Log(L"[COUNTERS@%s] magic=0x%08x ref=%d pad+0=0x%08x pad+4=0x%08x pad+8=0x%08x pad+16=0x%08x",
        label, m_magic, m_refCount, pDW[0], pDW[1], pDW[2], pDW[4]);
}

void DualEventSink::DumpDWords(const wchar_t* label, int count)
{
    Log(L"=== DWORD DUMP @ %s (from obj+8) ===", label);
    DWORD* pBase = (DWORD*)m_pad;
    for (int i = 0; i < count; i++)
    {
        DWORD val = pBase[i];
        if (val != 0)
            Log(L"  +%03d: 0x%08X (%u / %d)", i * 4 + 8, val, val, (int)val);
    }
}

// --- IUnknown ---

STDMETHODIMP DualEventSink::QueryInterface(REFIID riid, void** ppv)
{
    if (!ppv) return E_INVALIDARG;
    wchar_t gs[64];
    StringFromGUID2(riid, gs, 64);

    if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IRSTopologyOnlineNotify))
    {
        Log(L"[SINK] QI for %s - accepting (notify vtable)", gs);
        *ppv = static_cast<IRSTopologyOnlineNotify*>(this);
        AddRef();
        return S_OK;
    }
    if (IsEqualIID(riid, IID_ITopologyBusEvents))
    {
        Log(L"[SINK] QI for %s - accepting (bus vtable)", gs);
        *ppv = static_cast<ITopologyBusEvents*>(this);
        AddRef();
        return S_OK;
    }
    // FTM: enables cross-apartment marshaling without proxy/stubs
    if (IsEqualIID(riid, IID_IMarshal) && m_pFTM)
    {
        Log(L"[SINK] QI for IMarshal - delegating to FTM");
        return m_pFTM->QueryInterface(riid, ppv);
    }

    // Accept ALL unknown IIDs - return the notify vtable.
    // CP[1] {F0B077A1-...} and CP[2] {E548068D-...} MUST connect.
    Log(L"[SINK] QI for %s - accepting (unknown, returning notify vtable)", gs);
    *ppv = static_cast<IRSTopologyOnlineNotify*>(this);
    AddRef();
    return S_OK;
}

STDMETHODIMP_(ULONG) DualEventSink::AddRef() { return InterlockedIncrement(&m_refCount); }
STDMETHODIMP_(ULONG) DualEventSink::Release()
{
    ULONG c = InterlockedDecrement(&m_refCount);
    return c;  // Don't self-delete
}

// --- IRSTopologyOnlineNotify methods ---

STDMETHODIMP DualEventSink::BrowseStarted(IUnknown* pBus)
{
    Log(L"[ENUM:%s] Browse started (pBus=0x%p capture=%d)",
        m_label.c_str(), pBus, (int)g_captureBuses);
    // Capture backplane buses during Phase 4a
    if (g_captureBuses && pBus && (uintptr_t)pBus > 0x10000)
    {
        IDispatch* pDisp = nullptr;
        if (SUCCEEDED(pBus->QueryInterface(IID_IDispatch, (void**)&pDisp)) && pDisp)
        {
            pDisp->Release();
            bool dup = false;
            for (size_t i = 0; i < g_capturedBuses.size(); i++)
            {
                if (g_capturedBuses[i] == pBus) { dup = true; break; }
            }
            if (!dup)
            {
                pBus->AddRef();
                g_capturedBuses.push_back(pBus);
                Log(L"[ENUM:%s] Captured object #%d (0x%p)",
                    m_label.c_str(), (int)g_capturedBuses.size(), pBus);
            }
        }
    }
    return S_OK;
}

STDMETHODIMP DualEventSink::BrowseCycled(IUnknown* pBus)
{
    m_cycleComplete = true;
    Log(L"[ENUM:%s] BrowseCycled (explicit)", m_label.c_str());
    return S_OK;
}

STDMETHODIMP DualEventSink::BrowseEnded(IUnknown* pBus)
{
    m_cycleComplete = true;
    m_browseEnded = true;
    Log(L"[ENUM:%s] BrowseEnded (%d addresses seen)", m_label.c_str(), m_addressCount);
    return S_OK;
}

STDMETHODIMP DualEventSink::Found(IUnknown* pBus, VARIANT addr)
{
    wchar_t addrBuf[256] = L"<unknown>";
    SafeVariantToString(&addr, addrBuf, 256);
    if (addr.vt == VT_I2 || addr.vt == VT_I4)
        Log(L"[ENUM:%s] Slot %s found", m_label.c_str(), addrBuf);
    else
        Log(L"[ENUM:%s] Address %s found", m_label.c_str(), addrBuf);

    EnterCriticalSection(&m_cs);
    m_addressCount++;
    if (m_seenAddresses.count(addrBuf) > 0)
    {
        if (!m_cycleComplete)
        {
            m_cycleComplete = true;
            Log(L"[ENUM:%s] Cycle complete -- repeat address %s (after %d addresses)",
                m_label.c_str(), addrBuf, m_addressCount);
        }
    }
    else
        m_seenAddresses.insert(addrBuf);
    LeaveCriticalSection(&m_cs);

    g_discoveredDevices.push_back(addrBuf);
    return S_OK;
}

STDMETHODIMP DualEventSink::NothingAtAddress(IUnknown* pBus, VARIANT addr)
{
    wchar_t addrBuf[256] = L"<unknown>";
    SafeVariantToString(&addr, addrBuf, 256);
    Log(L"[ENUM:%s] Nothing at %s", m_label.c_str(), addrBuf);
    return S_OK;
}

// --- ITopologyBusEvents methods ---

STDMETHODIMP DualEventSink::OnPortConnect(IUnknown*, IUnknown*, VARIANT) { return S_OK; }
STDMETHODIMP DualEventSink::OnPortDisconnect(IUnknown*, IUnknown*, VARIANT) { return S_OK; }
STDMETHODIMP DualEventSink::OnPortChangeAddress(IUnknown*, IUnknown*, VARIANT, VARIANT) { return S_OK; }
STDMETHODIMP DualEventSink::OnPortChangeState(IUnknown*, long) { return S_OK; }

STDMETHODIMP DualEventSink::OnBrowseStarted(IUnknown* pBus)
{
    Log(L"[BUS:%s] Browse started", m_label.c_str());
    return S_OK;
}

STDMETHODIMP DualEventSink::OnBrowseCycled(IUnknown*)
{
    m_cycleComplete = true;
    Log(L"[BUS:%s] OnBrowseCycled (explicit)", m_label.c_str());
    return S_OK;
}

STDMETHODIMP DualEventSink::OnBrowseEnded(IUnknown*)
{
    m_cycleComplete = true;
    m_browseEnded = true;
    Log(L"[BUS:%s] BrowseEnded (%d addresses seen)", m_label.c_str(), m_addressCount);
    return S_OK;
}

STDMETHODIMP DualEventSink::OnBrowseAddressFound(IUnknown*, VARIANT addr)
{
    wchar_t addrBuf[256] = L"<unknown>";
    SafeVariantToString(&addr, addrBuf, 256);
    if (addr.vt == VT_I2 || addr.vt == VT_I4)
        Log(L"[BUS:%s] Slot %s found", m_label.c_str(), addrBuf);
    else
        Log(L"[BUS:%s] Address %s found", m_label.c_str(), addrBuf);

    EnterCriticalSection(&m_cs);
    m_addressCount++;
    if (m_seenAddresses.count(addrBuf) > 0)
    {
        if (!m_cycleComplete)
        {
            m_cycleComplete = true;
            Log(L"[BUS:%s] Cycle complete -- repeat address %s (after %d addresses)",
                m_label.c_str(), addrBuf, m_addressCount);
        }
    }
    else
        m_seenAddresses.insert(addrBuf);
    LeaveCriticalSection(&m_cs);

    g_discoveredDevices.push_back(addrBuf);
    return S_OK;
}

STDMETHODIMP DualEventSink::OnBrowseAddressNotFound(IUnknown*, VARIANT) { return S_OK; }
