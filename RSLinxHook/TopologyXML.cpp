#include "TopologyXML.h"
#include "Logging.h"

// Extract attribute value from position in buf: attr="value" → value into out[outMax]
static bool ExtractAttr(const char* pos, const char* attrName, char* out, int outMax)
{
    // Build search: " attrName=\""
    char needle[128];
    snprintf(needle, sizeof(needle), " %s=\"", attrName);
    const char* found = strstr(pos, needle);
    if (!found || found > pos + 400) return false;
    found += strlen(needle);
    const char* end = strchr(found, '"');
    if (!end) return false;
    int len = (int)(end - found);
    if (len >= outMax) len = outMax - 1;
    memcpy(out, found, len);
    out[len] = '\0';
    return true;
}

// ============================================================
// TopologyXML globals
// ============================================================

std::map<std::wstring, DeviceInfo> g_deviceDetails;
std::map<std::wstring, QueryResult> g_queryCache;

// ============================================================
// TopologyXML implementations
// ============================================================

bool SaveTopologyXML(IRSTopologyGlobals* pGlobals, const wchar_t* filename)
{
    IDispatch* pDisp = nullptr;
    HRESULT hr = pGlobals->QueryInterface(IID_IDispatch, (void**)&pDisp);
    if (FAILED(hr)) return false;

    VARIANT args[3];
    VariantInit(&args[0]);
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(filename);
    VariantInit(&args[1]);
    args[1].vt = VT_I4;
    args[1].lVal = 100;
    VariantInit(&args[2]);
    args[2].vt = VT_BSTR;
    args[2].bstrVal = SysAllocString(L"");

    DISPPARAMS params = {};
    params.rgvarg = args;
    params.cArgs = 3;

    VARIANT result;
    VariantInit(&result);

    hr = pDisp->Invoke(1610743808, IID_NULL, LOCALE_USER_DEFAULT,
                        DISPATCH_METHOD, &params, &result, nullptr, nullptr);

    VariantClear(&args[0]);
    VariantClear(&args[1]);
    VariantClear(&args[2]);
    VariantClear(&result);
    pDisp->Release();

    return SUCCEEDED(hr);
}

TopologyCounts CountDevicesInXML(const wchar_t* filename)
{
    TopologyCounts counts = { 0, 0 };
    FILE* f = _wfopen(filename, L"r");
    if (!f) return counts;
    char buf[65536];
    size_t n = fread(buf, 1, sizeof(buf) - 1, f);
    buf[n] = 0;
    fclose(f);
    char* p = buf;
    while ((p = strstr(p, "<device ")) != nullptr)
    {
        counts.totalDevices++;
        char* cn = strstr(p, "classname=\"");
        if (cn && cn < p + 300)
        {
            cn += 11;
            if (strncmp(cn, "Unrecognized Device", 19) != 0 &&
                strncmp(cn, "Workstation", 11) != 0)
                counts.identifiedDevices++;
        }
        p++;
    }
    return counts;
}

// Check if a specific IP address has been identified (non-Unrecognized) in topology XML
bool IsTargetIdentifiedInXML(const wchar_t* filename, const std::vector<std::wstring>& targetIPs)
{
    FILE* f = _wfopen(filename, L"r");
    if (!f) return false;
    char buf[65536];
    size_t n = fread(buf, 1, sizeof(buf) - 1, f);
    buf[n] = 0;
    fclose(f);

    for (auto& wip : targetIPs)
    {
        // Convert wide IP to narrow
        char ipNarrow[64];
        WideCharToMultiByte(CP_ACP, 0, wip.c_str(), -1, ipNarrow, 64, NULL, NULL);

        // Build search pattern: value="IP"
        char pattern[128];
        snprintf(pattern, sizeof(pattern), "value=\"%s\"", ipNarrow);

        // Find the address element containing this IP
        char* addrPos = strstr(buf, pattern);
        if (!addrPos) continue;

        // Look for a <device> with a non-Unrecognized classname after this address
        char* p = addrPos;
        char* searchEnd = addrPos + 2000;
        if (searchEnd > buf + n) searchEnd = buf + n;

        while (p < searchEnd && (p = strstr(p, "<device ")) != nullptr && p < searchEnd)
        {
            char* cn = strstr(p, "classname=\"");
            if (cn && cn < p + 500)
            {
                cn += 11;
                if (strncmp(cn, "Unrecognized Device", 19) != 0 &&
                    strncmp(cn, "Workstation", 11) != 0 &&
                    strncmp(cn, "\"", 1) != 0)  // empty classname
                    return true;
            }
            p++;
        }
    }
    return false;
}

// Query topology XML for a device at an IP/port/slot path
QueryResult QueryXMLForPath(const wchar_t* xmlFile,
                             const std::wstring& ip,
                             const std::wstring& portName,
                             int slot)
{
    QueryResult result;
    result.ip = ip;
    result.portName = portName;
    result.slot = slot;

    FILE* f = _wfopen(xmlFile, L"r");
    if (!f) return result;

    static char buf[262144];
    size_t n = fread(buf, 1, sizeof(buf) - 1, f);
    buf[n] = '\0';
    fclose(f);

    // Convert ip to narrow
    char ipA[64];
    WideCharToMultiByte(CP_ACP, 0, ip.c_str(), -1, ipA, sizeof(ipA), NULL, NULL);

    // Find <address type="String" value="IP">
    char addrPattern[128];
    snprintf(addrPattern, sizeof(addrPattern), "value=\"%s\"", ipA);
    char* pos1 = strstr(buf, addrPattern);
    if (!pos1) return result;

    if (portName.empty())
    {
        // IP-only query: find first <device> after the address element
        char* devPos = strstr(pos1, "<device ");
        if (!devPos || devPos > pos1 + 500) return result;
        // Skip <device reference="..."> entries
        if (strncmp(devPos + 8, "reference", 9) == 0) {
            devPos = strstr(devPos + 1, "<device ");
            if (!devPos || devPos > pos1 + 800) return result;
        }

        char cn[256] = {}, nm[256] = {};
        ExtractAttr(devPos, "classname", cn, sizeof(cn));
        ExtractAttr(devPos, "name", nm, sizeof(nm));

        result.found = true;
        result.classname = std::wstring(cn, cn + strlen(cn));
        result.deviceName = std::wstring(nm, nm + strlen(nm));
        return result;
    }

    // portName + slot query: find name="portName" then value="slot" then <device>
    char portA[128];
    WideCharToMultiByte(CP_ACP, 0, portName.c_str(), -1, portA, sizeof(portA), NULL, NULL);

    char portPattern[256];
    snprintf(portPattern, sizeof(portPattern), "name=\"%s\"", portA);
    char* pos2 = strstr(pos1, portPattern);
    if (!pos2) return result;

    char slotPattern[64];
    snprintf(slotPattern, sizeof(slotPattern), "value=\"%d\"", slot);
    char* pos3 = strstr(pos2, slotPattern);
    if (!pos3) return result;

    // Find <device> within 300 bytes of slot address element
    char* pos4 = strstr(pos3, "<device ");
    if (!pos4 || pos4 > pos3 + 300) return result;
    if (strncmp(pos4 + 8, "reference", 9) == 0)
    {
        pos4 = strstr(pos4 + 1, "<device ");
        if (!pos4 || pos4 > pos3 + 500) return result;
    }

    char cn[256] = {}, nm[256] = {};
    ExtractAttr(pos4, "classname", cn, sizeof(cn));
    ExtractAttr(pos4, "name", nm, sizeof(nm));

    result.found = true;
    result.classname = std::wstring(cn, cn + strlen(cn));
    result.deviceName = std::wstring(nm, nm + strlen(nm));
    return result;
}

// Populate g_queryCache from topology XML — called once after each browse phase.
// Walks every <address type="String" value="IP"> block and extracts:
//   "ip"              -> top-level device (classname, name)
//   "ip\Port\slot"    -> per-slot device on any named backplane-style bus
void PopulateQueryCache(const wchar_t* xmlFile)
{
    FILE* f = _wfopen(xmlFile, L"r");
    if (!f) return;
    static char buf[262144];
    size_t n = fread(buf, 1, sizeof(buf) - 1, f);
    buf[n] = '\0';
    fclose(f);

    char* p = buf;
    while ((p = strstr(p, "<address type=\"String\" value=\"")) != nullptr)
    {
        p += 30; // skip to start of IP value
        char* ipEnd = strchr(p, '"');
        if (!ipEnd) break;
        std::string ipA(p, ipEnd);
        std::wstring ipW(ipA.begin(), ipA.end());
        p = ipEnd;

        // Find top-level <device> for this IP (within 500 bytes)
        char* searchEnd = p + 500;
        if (searchEnd > buf + n) searchEnd = buf + n;
        char* devPos = strstr(p, "<device ");
        if (!devPos || devPos >= searchEnd) continue;
        if (strncmp(devPos + 8, "reference", 9) == 0) {
            devPos = strstr(devPos + 1, "<device ");
            if (!devPos || devPos >= searchEnd) continue;
        }

        char cn[256] = {}, nm[256] = {};
        ExtractAttr(devPos, "classname", cn, sizeof(cn));
        ExtractAttr(devPos, "name", nm, sizeof(nm));

        // Store IP-only entry
        QueryResult ipResult;
        ipResult.found = true;
        ipResult.ip = ipW;
        ipResult.classname = std::wstring(cn, cn + strlen(cn));
        ipResult.deviceName = std::wstring(nm, nm + strlen(nm));
        ipResult.slot = -1;
        g_queryCache[ipW] = ipResult;

        // Walk into ports → buses → slots
        // Find all <port name="..."> within the next 8KB of this device block
        char* portSearch = devPos;
        char* portEnd = devPos + 8192;
        if (portEnd > buf + n) portEnd = buf + n;

        while (portSearch < portEnd)
        {
            char* portPos = strstr(portSearch, "<port ");
            if (!portPos || portPos >= portEnd) break;

            char portName[128] = {};
            ExtractAttr(portPos, "name", portName, sizeof(portName));
            std::wstring portNameW(portName, portName + strlen(portName));

            // Find <bus> within the next 200 bytes of this port
            char* busPos = strstr(portPos, "<bus ");
            if (!busPos || busPos > portPos + 200) { portSearch = portPos + 1; continue; }

            // Enumerate <address type="Short" value="N"> within the bus block (next 4KB)
            char* slotSearch = busPos;
            char* slotEnd = busPos + 4096;
            if (slotEnd > buf + n) slotEnd = buf + n;

            while (slotSearch < slotEnd)
            {
                char* slotPos = strstr(slotSearch, "<address type=\"Short\" value=\"");
                if (!slotPos || slotPos >= slotEnd) break;

                char slotStr[16] = {};
                // value is right after 'value="'
                char* sv = slotPos + 29; // skip '<address type="Short" value="' (29 chars)
                char* svEnd = strchr(sv, '"');
                if (!svEnd) { slotSearch = slotPos + 1; continue; }
                int slotLen = (int)(svEnd - sv < 15 ? svEnd - sv : 15);
                memcpy(slotStr, sv, slotLen);
                int slotN = atoi(slotStr);

                // Find <device> within 200 bytes of this slot address
                char* slotDevPos = strstr(slotPos, "<device ");
                if (!slotDevPos || slotDevPos > slotPos + 200) { slotSearch = slotPos + 1; continue; }
                if (strncmp(slotDevPos + 8, "reference", 9) == 0) {
                    slotDevPos = strstr(slotDevPos + 1, "<device ");
                    if (!slotDevPos || slotDevPos > slotPos + 300) { slotSearch = slotPos + 1; continue; }
                }

                char scn[256] = {}, snm[256] = {};
                ExtractAttr(slotDevPos, "classname", scn, sizeof(scn));
                ExtractAttr(slotDevPos, "name", snm, sizeof(snm));

                // Build cache key: "ip\portName\slot"
                wchar_t slotKeyBuf[512];
                swprintf(slotKeyBuf, 512, L"%s\\%s\\%d", ipW.c_str(), portNameW.c_str(), slotN);

                QueryResult slotResult;
                slotResult.found = true;
                slotResult.ip = ipW;
                slotResult.portName = portNameW;
                slotResult.slot = slotN;
                slotResult.classname = std::wstring(scn, scn + strlen(scn));
                slotResult.deviceName = std::wstring(snm, snm + strlen(snm));
                g_queryCache[slotKeyBuf] = slotResult;

                slotSearch = slotDevPos + 1;
            }

            portSearch = portPos + 1;
        }
    }
}

// Update g_deviceDetails with IP addresses extracted from topology XML
void UpdateDeviceIPsFromXML(const wchar_t* filename)
{
    FILE* f = _wfopen(filename, L"r");
    if (!f) return;
    char buf[65536];
    size_t n = fread(buf, 1, sizeof(buf) - 1, f);
    buf[n] = 0;
    fclose(f);

    // Find each <address type="String" value="IP">
    char* p = buf;
    while ((p = strstr(p, "<address type=\"String\" value=\"")) != nullptr)
    {
        p += 30;  // skip past '<address type="String" value="' to start of IP
        char* ipEnd = strchr(p, '"');
        if (!ipEnd) break;
        std::string ipA(p, ipEnd);
        p = ipEnd;

        // Look for <device name="..." within this address block (next 2KB)
        char* searchEnd = p + 2000;
        if (searchEnd > buf + n) searchEnd = buf + n;
        char* dev = strstr(p, "<device ");
        if (dev && dev < searchEnd)
        {
            // Skip <device reference="..."> entries (no name attribute)
            if (strncmp(dev + 8, "reference", 9) == 0) continue;

            char* nameAttr = strstr(dev, "name=\"");
            if (nameAttr && nameAttr < dev + 200)
            {
                nameAttr += 6;
                char* nameEnd = strchr(nameAttr, '"');
                if (nameEnd)
                {
                    std::string nameA(nameAttr, nameEnd);
                    std::wstring nameW(nameA.begin(), nameA.end());
                    std::wstring ipW(ipA.begin(), ipA.end());

                    auto it = g_deviceDetails.find(nameW);
                    if (it != g_deviceDetails.end())
                        it->second.ip = ipW;
                    else
                    {
                        DeviceInfo info;
                        info.productName = nameW;
                        info.ip = ipW;
                        g_deviceDetails[nameW] = info;
                    }
                }
            }
        }
    }
}
