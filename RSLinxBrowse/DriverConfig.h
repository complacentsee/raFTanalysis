/**
 * DriverConfig.h
 *
 * RSLinx driver configuration and EtherNet/IP CIP ListIdentity scanner.
 * Reads device IPs from RSLinx Node Table registry and optionally
 * sends directed UDP ListIdentity packets on port 44818.
 *
 * Protocol: EtherNet/IP Encapsulation (ODVA standard)
 * Command: ListIdentity (0x0063)
 */

#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#define _WINSOCK_DEPRECATED_NO_WARNINGS

#include <winsock2.h>
#include <ws2tcpip.h>
#include <string>
#include <vector>

#pragma comment(lib, "ws2_32.lib")

struct CIPDevice
{
    std::wstring ipAddress;
    std::wstring productName;
    WORD vendorId;
    WORD deviceType;
    WORD productCode;
    BYTE revisionMajor;
    BYTE revisionMinor;
    WORD status;
    DWORD serialNumber;
    bool online;  // true if device responded
};

class DriverConfig
{
public:
    DriverConfig();
    ~DriverConfig();

    /**
     * Scan specific IPs for EtherNet/IP devices using directed ListIdentity.
     *
     * @param ipAddresses  List of IP addresses to probe
     * @param timeoutMs    How long to wait for all responses (default 5 seconds)
     * @return Vector of discovered devices (only those that responded)
     */
    std::vector<CIPDevice> ScanIPs(const std::vector<std::wstring>& ipAddresses, DWORD timeoutMs = 5000);

    /**
     * Read device IPs from RSLinx Node Table registry for a given driver.
     *
     * @param driverName  The RSLinx driver name (e.g. "Test")
     * @return Vector of IP address strings from the Node Table
     */
    static std::vector<std::wstring> ReadNodeTable(const std::wstring& driverName);

    /**
     * Convenience: scan all devices configured in RSLinx Node Table.
     */
    std::vector<CIPDevice> ScanNodeTable(const std::wstring& driverName, DWORD timeoutMs = 5000);

    /**
     * Get just the IP addresses from a scan result.
     */
    static std::vector<std::wstring> GetIPAddresses(const std::vector<CIPDevice>& devices);

    /**
     * Print device details to console.
     */
    static void PrintDevices(const std::vector<CIPDevice>& devices);

private:
    bool m_wsaInitialized;

    bool InitWinsock();
    void CleanupWinsock();

    std::vector<CIPDevice> ParseResponses(SOCKET sock, DWORD timeoutMs, size_t expectedCount);
    bool ParseListIdentityResponse(const BYTE* data, int dataLen, const sockaddr_in& fromAddr, CIPDevice& device);
};
