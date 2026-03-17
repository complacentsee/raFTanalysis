#include "TopologyXML.h"
#include "Logging.h"
#include "DispatchHelpers.h"
#include "STAHook.h"

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

// Resolve a <device reference="GUID"/> by finding objectid="GUID" elsewhere in buf.
// Fills cn and nm from the actual device element. Returns true if found.
static bool ResolveReference(const char* buf, const char* refDevPos,
                              char* cn, int cnMax, char* nm, int nmMax)
{
    // Extract GUID from reference="..."
    const char* refAttr = strstr(refDevPos, "reference=\"");
    if (!refAttr || refAttr > refDevPos + 200) return false;
    refAttr += 11; // skip 'reference="'
    const char* refEnd = strchr(refAttr, '"');
    if (!refEnd) return false;
    int guidLen = (int)(refEnd - refAttr);
    if (guidLen < 1 || guidLen >= 120) return false;

    // Build: objectid="GUID"
    char oidPattern[140];
    memcpy(oidPattern, "objectid=\"", 10);
    memcpy(oidPattern + 10, refAttr, guidLen);
    oidPattern[10 + guidLen] = '"';
    oidPattern[11 + guidLen] = '\0';

    const char* found = strstr(buf, oidPattern);
    if (!found) return false;

    // Walk back to the '<' that starts the <device> element
    const char* devStart = found;
    while (devStart > buf && *devStart != '<') devStart--;

    ExtractAttr(devStart, "classname", cn, cnMax);
    ExtractAttr(devStart, "name", nm, nmMax);
    return (*cn != '\0' || *nm != '\0');
}

// ============================================================
// TopologyXML globals
// ============================================================

std::map<std::wstring, DeviceInfo> g_deviceDetails;
std::map<std::wstring, QueryResult> g_queryCache;
std::map<std::wstring, std::vector<std::wstring>> g_driverDeviceNames;

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
    fseek(f, 0, SEEK_END);
    long fileSize = ftell(f);
    fseek(f, 0, SEEK_SET);
    size_t allocSize = (fileSize > 0) ? (size_t)fileSize + 1 : 262144;
    char* buf = (char*)malloc(allocSize);
    if (!buf) { fclose(f); return counts; }
    size_t n = fread(buf, 1, allocSize - 1, f);
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
    free(buf);
    return counts;
}

// Count how many target IPs have been identified (non-Unrecognized) in topology XML
int CountTargetsIdentifiedInXML(const wchar_t* filename, const std::vector<std::wstring>& targetIPs)
{
    FILE* f = _wfopen(filename, L"r");
    if (!f) return 0;
    fseek(f, 0, SEEK_END);
    long fileSize = ftell(f);
    fseek(f, 0, SEEK_SET);
    size_t allocSize = (fileSize > 0) ? (size_t)fileSize + 1 : 262144;
    char* buf = (char*)malloc(allocSize);
    if (!buf) { fclose(f); return 0; }
    size_t n = fread(buf, 1, allocSize - 1, f);
    buf[n] = 0;
    fclose(f);

    int count = 0;
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
                {
                    count++;
                    break;  // This target IP is identified, move to next
                }
            }
            p++;
        }
    }
    free(buf);
    return count;
}

// Check if ANY target IP has been identified — backwards compat wrapper
bool IsTargetIdentifiedInXML(const wchar_t* filename, const std::vector<std::wstring>& targetIPs)
{
    return CountTargetsIdentifiedInXML(filename, targetIPs) > 0;
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

    char cn[256] = {}, nm[256] = {};
    if (strncmp(pos4 + 8, "reference", 9) == 0)
    {
        if (!ResolveReference(buf, pos4, cn, sizeof(cn), nm, sizeof(nm)))
            return result;
    }
    else
    {
        ExtractAttr(pos4, "classname", cn, sizeof(cn));
        ExtractAttr(pos4, "name", nm, sizeof(nm));
    }

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
        // Find end of this device's block: next top-level <address type="String" (next IP)
        // or end of buffer, whichever comes first
        char* portSearch = devPos;
        char* nextIp = strstr(devPos + 1, "<address type=\"String\" value=\"");
        char* portEnd = nextIp ? nextIp : buf + n;

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

            // Enumerate <address type="Short" value="N"> within the bus block
            // Use </bus> as boundary instead of fixed window
            char* slotSearch = busPos;
            char* busClose = strstr(busPos, "</bus>");
            char* slotEnd = busClose ? busClose : portEnd;

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
                char scn[256] = {}, snm[256] = {};
                if (strncmp(slotDevPos + 8, "reference", 9) == 0) {
                    if (!ResolveReference(buf, slotDevPos, scn, sizeof(scn), snm, sizeof(snm)))
                    { slotSearch = slotPos + 1; continue; }
                } else {
                    ExtractAttr(slotDevPos, "classname", scn, sizeof(scn));
                    ExtractAttr(slotDevPos, "name", snm, sizeof(snm));
                }

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

// ============================================================
// WalkTopologyTree — emit N| topology block from cache + COM globals
// ============================================================

// Convert wide string to UTF-8, replacing pipe/newline chars with spaces.
static std::string WalkWideToUtf8(const std::wstring& w)
{
    if (w.empty()) return "";
    int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(),
                                nullptr, 0, nullptr, nullptr);
    std::string s(n, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(),
                        &s[0], n, nullptr, nullptr);
    for (char& c : s)
        if (c == '|' || c == '\n' || c == '\r') c = ' ';
    return s;
}

// Format: N|TYPE|field1|field2\n
static std::string WalkNodeLine(const char* msgType,
                                 const std::wstring& f1,
                                 const std::wstring& f2)
{
    return std::string(msgType) + "|" +
           WalkWideToUtf8(f1) + "|" +
           WalkWideToUtf8(f2) + "\n";
}

// Format: N|ADDR|addrType|addrVal|devName|classname\n
static std::string WalkAddrLine(const char* addrType,
                                 const std::wstring& addrVal,
                                 const std::wstring& devName,
                                 const std::wstring& classname)
{
    return std::string("N|ADDR|") + addrType + "|" +
           WalkWideToUtf8(addrVal) + "|" +
           WalkWideToUtf8(devName) + "|" +
           WalkWideToUtf8(classname) + "\n";
}

// Emit N|BEGIN...N|END block using g_driverDeviceNames (set by DoBusBrowse),
// g_deviceDetails (set by UpdateDeviceIPsFromXML), and g_queryCache (set by PopulateQueryCache).
// No COM calls made here — all data comes from in-memory caches.
// pGlobals is accepted for API consistency and null-guard only.
void WalkTopologyTree(IRSTopologyGlobals* pGlobals)
{
    if (!g_pipeConnected || !pGlobals || !g_pSharedConfig) return;

    std::vector<std::string> out;
    out.reserve(64);
    out.push_back("N|BEGIN\n");
    out.push_back(WalkNodeLine("N|ROOT", L"WORKSTATION", L"Workstation"));

    // Dedup by network address (IP, or devName when IP is unknown) so that
    // two driver configs pointing at the same device don't emit it twice.
    std::set<std::wstring> emittedAddrs;

    for (const auto& drv : g_pSharedConfig->drivers)
    {
        auto dit = g_driverDeviceNames.find(drv.name);
        if (dit == g_driverDeviceNames.end() || dit->second.empty()) continue;

        // Pre-scan: collect devices not yet emitted for this driver.
        // Skip the bus entirely if there is nothing new to show.
        std::vector<std::wstring> newDevices;
        for (const auto& devName : dit->second)
        {
            if (devName.empty()) continue;
            std::wstring ip;
            {
                auto it = g_deviceDetails.find(devName);
                if (it != g_deviceDetails.end()) ip = it->second.ip;
            }
            std::wstring addrVal = ip.empty() ? devName : ip;
            if (emittedAddrs.count(addrVal) == 0)
                newDevices.push_back(devName);
        }
        if (newDevices.empty()) continue;

        out.push_back(WalkNodeLine("N|BUS", drv.name, L""));

        for (const auto& devName : newDevices)
        {
            // IP from g_deviceDetails
            std::wstring ip;
            {
                auto it = g_deviceDetails.find(devName);
                if (it != g_deviceDetails.end())
                    ip = it->second.ip;
            }

            // Classname from g_queryCache (keyed by IP)
            std::wstring classname;
            if (!ip.empty())
            {
                auto it = g_queryCache.find(ip);
                if (it != g_queryCache.end())
                    classname = it->second.classname;
            }

            std::wstring addrVal = ip.empty() ? devName : ip;
            emittedAddrs.insert(addrVal);
            out.push_back(WalkAddrLine("String", addrVal, devName, classname));

            if (ip.empty()) continue;

            // Find backplane slot entries: g_queryCache keys "ip\portName\slot"
            std::map<std::wstring, std::map<int, QueryResult>> ports;
            for (const auto& kv : g_queryCache)
            {
                const std::wstring& key = kv.first;
                if (key.size() <= ip.size() + 1) continue;
                if (key.compare(0, ip.size(), ip) != 0) continue;
                if (key[ip.size()] != L'\\') continue;
                size_t portEnd = key.find(L'\\', ip.size() + 1);
                if (portEnd == std::wstring::npos) continue;
                std::wstring portName = key.substr(ip.size() + 1, portEnd - ip.size() - 1);
                if (!portName.empty() && kv.second.slot >= 0)
                    ports[portName][kv.second.slot] = kv.second;
            }

            if (!ports.empty())
            {
                out.push_back("N|PUSH\n");
                for (const auto& portKv : ports)
                {
                    out.push_back(WalkNodeLine("N|BUS", portKv.first, L""));
                    for (const auto& slotKv : portKv.second)
                    {
                        const QueryResult& qr = slotKv.second;
                        out.push_back(WalkAddrLine("Short",
                            std::to_wstring(slotKv.first),
                            qr.deviceName,
                            qr.classname));
                    }
                }
                out.push_back("N|POP\n");
            }
        }
    }

    out.push_back("N|END\n");

    EnterCriticalSection(&g_logCS);
    for (const auto& line : out)
        PipeSend(line.c_str(), (int)line.size());
    LeaveCriticalSection(&g_logCS);

    Log(L"[WALK] N| block sent: %d lines", (int)out.size());
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
