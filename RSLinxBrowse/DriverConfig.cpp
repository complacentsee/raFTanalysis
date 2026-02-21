/**
 * DriverConfig.cpp
 *
 * RSLinx driver configuration and EtherNet/IP CIP ListIdentity scanner.
 *
 * Reads target IPs from RSLinx Node Table registry and optionally
 * sends directed ListIdentity (0x0063) UDP packets to specific device IPs
 * on port 44818.
 *
 * Packet format follows ODVA EtherNet/IP specification.
 */

#include "DriverConfig.h"
#include <iostream>
#include <iomanip>
#include <chrono>

// EtherNet/IP constants
#define ENCAP_PORT          44818   // 0xAF12
#define CMD_LIST_IDENTITY   0x0063
#define ITEM_ID_IDENTITY    0x000C

// EtherNet/IP Encapsulation Header (24 bytes)
#pragma pack(push, 1)
struct EncapsulationHeader
{
    WORD    command;        // Command code
    WORD    length;         // Data length (after header)
    DWORD   sessionHandle;  // Session handle (0 for unconnected)
    DWORD   status;         // Status (0 for request)
    BYTE    senderContext[8]; // Sender context (echoed in response)
    DWORD   options;        // Options (0)
};
#pragma pack(pop)

DriverConfig::DriverConfig()
    : m_wsaInitialized(false)
{
}

DriverConfig::~DriverConfig()
{
    CleanupWinsock();
}

bool DriverConfig::InitWinsock()
{
    if (m_wsaInitialized)
        return true;

    WSADATA wsaData;
    int result = WSAStartup(MAKEWORD(2, 2), &wsaData);
    if (result != 0)
    {
        std::wcerr << L"[ERROR] WSAStartup failed: " << result << std::endl;
        return false;
    }

    m_wsaInitialized = true;
    return true;
}

void DriverConfig::CleanupWinsock()
{
    if (m_wsaInitialized)
    {
        WSACleanup();
        m_wsaInitialized = false;
    }
}

std::vector<std::wstring> DriverConfig::ReadNodeTable(const std::wstring& driverName)
{
    std::vector<std::wstring> ips;

    // Find the driver in registry
    // Path: HKLM\SOFTWARE\WOW6432Node\Rockwell Software\RSLinx\Drivers\AB_ETH
    // We need to find which AB_ETH-N has Name == driverName
    std::wstring basePath = L"SOFTWARE\\WOW6432Node\\Rockwell Software\\RSLinx\\Drivers\\AB_ETH";

    HKEY hBaseKey;
    LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, basePath.c_str(), 0, KEY_READ, &hBaseKey);
    if (result != ERROR_SUCCESS)
    {
        std::wcerr << L"[ERROR] Cannot open RSLinx Drivers registry: " << result << std::endl;
        return ips;
    }

    // Enumerate subkeys (AB_ETH-1, AB_ETH-2, etc.)
    wchar_t subKeyName[256];
    DWORD subKeyNameLen;
    std::wstring foundDriverPath;

    for (DWORD i = 0; ; i++)
    {
        subKeyNameLen = 256;
        result = RegEnumKeyExW(hBaseKey, i, subKeyName, &subKeyNameLen, NULL, NULL, NULL, NULL);
        if (result != ERROR_SUCCESS)
            break;

        // Open this subkey and check its Name value
        HKEY hDriverKey;
        result = RegOpenKeyExW(hBaseKey, subKeyName, 0, KEY_READ, &hDriverKey);
        if (result != ERROR_SUCCESS)
            continue;

        wchar_t nameValue[256] = {};
        DWORD nameValueLen = sizeof(nameValue);
        DWORD nameType;
        result = RegQueryValueExW(hDriverKey, L"Name", NULL, &nameType, (BYTE*)nameValue, &nameValueLen);

        if (result == ERROR_SUCCESS && nameType == REG_SZ && driverName == nameValue)
        {
            foundDriverPath = basePath + L"\\" + subKeyName;
            RegCloseKey(hDriverKey);
            break;
        }

        RegCloseKey(hDriverKey);
    }

    RegCloseKey(hBaseKey);

    if (foundDriverPath.empty())
    {
        std::wcerr << L"[ERROR] Driver '" << driverName << L"' not found in registry" << std::endl;
        return ips;
    }

    std::wcout << L"[OK] Found driver '" << driverName << L"' at: " << foundDriverPath << std::endl;

    // Open Node Table subkey
    std::wstring nodeTablePath = foundDriverPath + L"\\Node Table";
    HKEY hNodeTable;
    result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, nodeTablePath.c_str(), 0, KEY_READ, &hNodeTable);
    if (result != ERROR_SUCCESS)
    {
        std::wcerr << L"[ERROR] Cannot open Node Table: " << result << std::endl;
        return ips;
    }

    // Enumerate values (0, 1, 2, ...)
    wchar_t valueName[64];
    DWORD valueNameLen;
    wchar_t valueData[256];
    DWORD valueDataLen;
    DWORD valueType;

    for (DWORD i = 0; ; i++)
    {
        valueNameLen = 64;
        valueDataLen = sizeof(valueData);
        result = RegEnumValueW(hNodeTable, i, valueName, &valueNameLen, NULL,
                               &valueType, (BYTE*)valueData, &valueDataLen);
        if (result != ERROR_SUCCESS)
            break;

        if (valueType == REG_SZ)
        {
            ips.push_back(valueData);
        }
    }

    RegCloseKey(hNodeTable);

    std::wcout << L"[OK] Read " << ips.size() << L" IP(s) from Node Table" << std::endl;
    for (size_t i = 0; i < ips.size(); i++)
    {
        std::wcout << L"  " << i << L": " << ips[i] << std::endl;
    }

    return ips;
}

std::vector<CIPDevice> DriverConfig::ScanNodeTable(const std::wstring& driverName, DWORD timeoutMs)
{
    std::vector<std::wstring> nodeIPs = ReadNodeTable(driverName);
    if (nodeIPs.empty())
    {
        std::wcerr << L"[ERROR] No IPs in Node Table for driver '" << driverName << L"'" << std::endl;
        return {};
    }

    return ScanIPs(nodeIPs, timeoutMs);
}

std::vector<CIPDevice> DriverConfig::ScanIPs(const std::vector<std::wstring>& ipAddresses, DWORD timeoutMs)
{
    std::vector<CIPDevice> devices;

    if (ipAddresses.empty())
        return devices;

    if (!InitWinsock())
        return devices;

    // Create UDP socket
    SOCKET sock = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP);
    if (sock == INVALID_SOCKET)
    {
        std::wcerr << L"[ERROR] socket() failed: " << WSAGetLastError() << std::endl;
        return devices;
    }

    // Bind to any local address
    sockaddr_in localAddr = {};
    localAddr.sin_family = AF_INET;
    localAddr.sin_addr.s_addr = INADDR_ANY;
    localAddr.sin_port = 0;

    if (bind(sock, (sockaddr*)&localAddr, sizeof(localAddr)) == SOCKET_ERROR)
    {
        std::wcerr << L"[ERROR] bind() failed: " << WSAGetLastError() << std::endl;
        closesocket(sock);
        return devices;
    }

    // Build ListIdentity request
    EncapsulationHeader request = {};
    request.command = CMD_LIST_IDENTITY;
    request.length = 0;
    request.sessionHandle = 0;
    request.status = 0;
    memset(request.senderContext, 0, 8);
    request.senderContext[0] = 'R';
    request.senderContext[1] = 'S';
    request.options = 0;

    // Send directed ListIdentity to each IP
    std::wcout << L"[INFO] Sending CIP ListIdentity to " << ipAddresses.size()
               << L" device(s) on port " << ENCAP_PORT << L"..." << std::endl;

    int sentCount = 0;
    for (const auto& ip : ipAddresses)
    {
        sockaddr_in destAddr = {};
        destAddr.sin_family = AF_INET;
        destAddr.sin_port = htons(ENCAP_PORT);

        // Convert wide string IP to narrow
        char narrowIP[64] = {};
        WideCharToMultiByte(CP_ACP, 0, ip.c_str(), -1, narrowIP, sizeof(narrowIP), NULL, NULL);

        // Use inet_pton for the conversion
        if (inet_pton(AF_INET, narrowIP, &destAddr.sin_addr) != 1)
        {
            std::wcerr << L"[WARN] Invalid IP: " << ip << std::endl;
            continue;
        }

        int sent = sendto(sock, (char*)&request, sizeof(request), 0,
                          (sockaddr*)&destAddr, sizeof(destAddr));
        if (sent == SOCKET_ERROR)
        {
            std::wcerr << L"[WARN] sendto " << ip << L" failed: " << WSAGetLastError() << std::endl;
            continue;
        }

        std::wcout << L"  -> " << ip << std::endl;
        sentCount++;
    }

    std::wcout << L"[OK] Sent to " << sentCount << L" device(s), waiting for responses ("
               << timeoutMs / 1000 << L"s timeout)..." << std::endl;

    // Collect responses
    devices = ParseResponses(sock, timeoutMs, ipAddresses.size());

    closesocket(sock);
    return devices;
}

std::vector<CIPDevice> DriverConfig::ParseResponses(SOCKET sock, DWORD timeoutMs, size_t expectedCount)
{
    std::vector<CIPDevice> devices;
    BYTE buffer[4096];

    auto startTime = std::chrono::steady_clock::now();

    while (true)
    {
        auto now = std::chrono::steady_clock::now();
        DWORD elapsed = static_cast<DWORD>(
            std::chrono::duration_cast<std::chrono::milliseconds>(now - startTime).count());
        if (elapsed >= timeoutMs)
            break;

        // Early exit if we got all expected responses
        if (devices.size() >= expectedCount)
        {
            std::wcout << L"[OK] All " << expectedCount << L" devices responded" << std::endl;
            break;
        }

        // Set receive timeout
        DWORD remaining = timeoutMs - elapsed;
        DWORD recvTimeout = (remaining < 500) ? remaining : 500;
        setsockopt(sock, SOL_SOCKET, SO_RCVTIMEO, (char*)&recvTimeout, sizeof(recvTimeout));

        sockaddr_in fromAddr = {};
        int fromLen = sizeof(fromAddr);

        int received = recvfrom(sock, (char*)buffer, sizeof(buffer), 0,
                                (sockaddr*)&fromAddr, &fromLen);

        if (received == SOCKET_ERROR)
        {
            int err = WSAGetLastError();
            if (err == WSAETIMEDOUT || err == WSAEWOULDBLOCK)
                continue;
            break;
        }

        if (received < (int)sizeof(EncapsulationHeader))
            continue;

        // Verify it's a ListIdentity response
        EncapsulationHeader* header = (EncapsulationHeader*)buffer;
        if (header->command != CMD_LIST_IDENTITY)
            continue;

        // Parse the identity data
        CIPDevice device;
        device.online = true;
        if (ParseListIdentityResponse(buffer + sizeof(EncapsulationHeader),
                                       received - sizeof(EncapsulationHeader),
                                       fromAddr, device))
        {
            // Check for duplicate
            bool isDuplicate = false;
            for (const auto& existing : devices)
            {
                if (existing.ipAddress == device.ipAddress)
                {
                    isDuplicate = true;
                    break;
                }
            }

            if (!isDuplicate)
            {
                devices.push_back(device);
                std::wcout << L"  [ONLINE] " << device.ipAddress;
                if (!device.productName.empty())
                    std::wcout << L" - " << device.productName;
                std::wcout << std::endl;
            }
        }
    }

    return devices;
}

bool DriverConfig::ParseListIdentityResponse(const BYTE* data, int dataLen,
                                             const sockaddr_in& fromAddr, CIPDevice& device)
{
    // Response data format (after encapsulation header):
    // Item Count (2 bytes)
    // For each item:
    //   Item Type ID (2 bytes): 0x000C
    //   Item Length (2 bytes)
    //   Protocol Version (2 bytes)
    //   Socket Address (16 bytes)
    //   Vendor ID (2 bytes, little-endian)
    //   Device Type (2 bytes, little-endian)
    //   Product Code (2 bytes, little-endian)
    //   Revision Major (1 byte)
    //   Revision Minor (1 byte)
    //   Status (2 bytes, little-endian)
    //   Serial Number (4 bytes, little-endian)
    //   Product Name Length (1 byte)
    //   Product Name (variable, ASCII)
    //   State (1 byte)

    if (dataLen < 2)
        return false;

    WORD itemCount = *(WORD*)data;
    data += 2;
    dataLen -= 2;

    if (itemCount < 1)
        return false;

    if (dataLen < 4)
        return false;

    WORD itemTypeId = *(WORD*)data;
    WORD itemLength = *(WORD*)(data + 2);
    data += 4;
    dataLen -= 4;

    if (itemTypeId != ITEM_ID_IDENTITY)
        return false;

    if (dataLen < itemLength || itemLength < 33)
        return false;

    // Protocol version (skip)
    data += 2;

    // Socket Address (16 bytes) - skip but extract IP
    data += 16;

    // Use the actual source IP from recvfrom
    char ipStr[32];
    inet_ntop(AF_INET, &fromAddr.sin_addr, ipStr, sizeof(ipStr));

    wchar_t wideIp[32];
    MultiByteToWideChar(CP_ACP, 0, ipStr, -1, wideIp, 32);
    device.ipAddress = wideIp;

    // Vendor ID
    device.vendorId = *(WORD*)data;
    data += 2;

    // Device Type
    device.deviceType = *(WORD*)data;
    data += 2;

    // Product Code
    device.productCode = *(WORD*)data;
    data += 2;

    // Revision
    device.revisionMajor = data[0];
    device.revisionMinor = data[1];
    data += 2;

    // Status
    device.status = *(WORD*)data;
    data += 2;

    // Serial Number
    device.serialNumber = *(DWORD*)data;
    data += 4;

    // Product Name
    BYTE nameLen = data[0];
    data += 1;

    if (nameLen > 0 && nameLen <= 128)
    {
        char nameBuffer[129] = {};
        memcpy(nameBuffer, data, nameLen);
        nameBuffer[nameLen] = '\0';

        wchar_t wideName[129] = {};
        MultiByteToWideChar(CP_ACP, 0, nameBuffer, -1, wideName, 129);
        device.productName = wideName;
    }

    return true;
}

std::vector<std::wstring> DriverConfig::GetIPAddresses(const std::vector<CIPDevice>& devices)
{
    std::vector<std::wstring> ips;
    for (const auto& dev : devices)
    {
        ips.push_back(dev.ipAddress);
    }
    return ips;
}

void DriverConfig::PrintDevices(const std::vector<CIPDevice>& devices)
{
    std::wcout << std::endl;
    std::wcout << L"========================================" << std::endl;
    std::wcout << L"CIP ListIdentity Scan Results" << std::endl;
    std::wcout << L"========================================" << std::endl;
    std::wcout << L"Responded: " << devices.size() << L" device(s)" << std::endl;
    std::wcout << std::endl;

    for (size_t i = 0; i < devices.size(); i++)
    {
        const auto& dev = devices[i];
        std::wcout << L"  " << std::setw(2) << (i + 1) << L". " << dev.ipAddress;

        if (!dev.productName.empty())
            std::wcout << L" - " << dev.productName;

        std::wcout << std::endl;

        std::wcout << L"      Vendor: " << dev.vendorId
                   << L"  Type: " << dev.deviceType
                   << L"  Product: " << dev.productCode
                   << L"  Rev: " << (int)dev.revisionMajor << L"." << (int)dev.revisionMinor
                   << std::endl;
        std::wcout << L"      Serial: 0x" << std::hex << std::setw(8) << std::setfill(L'0')
                   << dev.serialNumber << std::dec << std::setfill(L' ')
                   << L"  Status: 0x" << std::hex << dev.status << std::dec
                   << std::endl;
    }

    std::wcout << L"========================================" << std::endl;
}
