/**
 * TopologyBrowser.h
 *
 * High-level interface for browsing RSLinx topology.
 * Handles COM initialization, event sink connection, and browse orchestration.
 */

#pragma once

#include "RSLinxInterfaces.h"
#include "BrowseEventSink.h"
#include <string>
#include <vector>
#include <memory>

/**
 * TopologyBrowser - Main class for RSLinx topology browsing
 *
 * Usage:
 *   TopologyBrowser browser;
 *   if (browser.Initialize())
 *   {
 *       std::vector<std::wstring> devices;
 *       if (browser.BrowseDriver(L"Test", devices))
 *       {
 *           // devices now contains discovered IP addresses
 *       }
 *       browser.Uninitialize();
 *   }
 */
class TopologyBrowser
{
public:
    TopologyBrowser();
    ~TopologyBrowser();

    // Initialization
    bool Initialize();
    void Uninitialize();

    // Browse operations
    bool BrowseDriver(const std::wstring& driverName, std::vector<std::wstring>& discoveredDevices, DWORD timeoutMs = 30000);

    // Direct topology access
    bool SaveTopologyXML(const std::wstring& filename);

    // Manual device management
    bool AddDeviceManually(const std::wstring& driverName, const std::wstring& ipAddress, const std::wstring& deviceName = L"");

    // Error information
    std::wstring GetLastError() const { return m_lastError; }

    // Parse devices from a topology XML file for a given bus name
    std::vector<std::wstring> ParseDevicesFromXML(const std::wstring& xmlPath, const std::wstring& busName);

private:
    bool m_initialized;
    std::wstring m_lastError;

    // COM objects
    IHarmonyConnector* m_pHarmonyConnector;
    IRSTopologyGlobals* m_pTopologyGlobals;
    IRSProjectGlobal* m_pProjectGlobal;
    IUnknown* m_pProject;
    IUnknown* m_pWorkstation;

    // Helper methods
    bool InitializeHarmonyServices();
    bool InitializeTopologyGlobals();
    bool GetWorkstation();
    IUnknown* GetBusObjectByName(const std::wstring& driverName);

    bool ConnectEventSink(IUnknown* pBus, BrowseEventSink* pSink, DWORD* pCookie);
    bool DisconnectEventSink(IUnknown* pBus, DWORD cookie);

    void SetLastError(const std::wstring& error);
    void ReleaseAll();
};
