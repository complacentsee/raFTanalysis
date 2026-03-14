/**
 * Standalone test for QueryXMLForPath
 * Compiles without COM dependencies — copies the implementation inline.
 */
#include <windows.h>
#include <stdio.h>
#include <string>
#include <map>

// ---- Inline copy of ExtractAttr + QueryXMLForPath ----

struct QueryResult {
    bool found = false;
    std::wstring classname;
    std::wstring deviceName;
    std::wstring ip;
    std::wstring portName;
    int slot = -1;
};

static bool ExtractAttr(const char* pos, const char* attrName, char* out, int outMax)
{
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

static QueryResult QueryXMLForPath(const wchar_t* xmlFile,
                                    const std::wstring& ip,
                                    const std::wstring& portName,
                                    int slot)
{
    QueryResult result;
    result.ip = ip;
    result.portName = portName;
    result.slot = slot;

    FILE* f = _wfopen(xmlFile, L"r");
    if (!f) { wprintf(L"  [ERR] Cannot open XML\n"); return result; }
    static char buf[262144];
    size_t n = fread(buf, 1, sizeof(buf) - 1, f);
    buf[n] = '\0';
    fclose(f);

    char ipA[64]; WideCharToMultiByte(CP_ACP, 0, ip.c_str(), -1, ipA, sizeof(ipA), NULL, NULL);
    char addrPattern[128]; snprintf(addrPattern, sizeof(addrPattern), "value=\"%s\"", ipA);
    char* pos1 = strstr(buf, addrPattern);
    if (!pos1) return result;

    if (portName.empty())
    {
        char* devPos = strstr(pos1, "<device ");
        if (!devPos || devPos > pos1 + 500) return result;
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

    char portA[128]; WideCharToMultiByte(CP_ACP, 0, portName.c_str(), -1, portA, sizeof(portA), NULL, NULL);
    char portPattern[256]; snprintf(portPattern, sizeof(portPattern), "name=\"%s\"", portA);
    char* pos2 = strstr(pos1, portPattern);
    if (!pos2) return result;

    char slotPattern[64]; snprintf(slotPattern, sizeof(slotPattern), "value=\"%d\"", slot);
    char* pos3 = strstr(pos2, slotPattern);
    if (!pos3) return result;

    char* pos4 = strstr(pos3, "<device ");
    if (!pos4 || pos4 > pos3 + 300) return result;
    if (strncmp(pos4 + 8, "reference", 9) == 0) {
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

// ---- Inline PopulateQueryCache for test ----

static std::map<std::wstring, QueryResult> g_queryCache;

static std::wstring MakeCacheKey(const std::wstring& ip, const std::wstring& portName, int slot)
{
    if (portName.empty()) return ip;
    wchar_t buf[512];
    swprintf(buf, 512, L"%s\\%s\\%d", ip.c_str(), portName.c_str(), slot);
    return buf;
}

static void PopulateQueryCache(const wchar_t* xmlFile)
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
        p += 30;
        char* ipEnd = strchr(p, '"');
        if (!ipEnd) break;
        std::string ipA(p, ipEnd);
        std::wstring ipW(ipA.begin(), ipA.end());
        p = ipEnd;

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

        QueryResult ipResult;
        ipResult.found = true; ipResult.ip = ipW; ipResult.slot = -1;
        ipResult.classname = std::wstring(cn, cn + strlen(cn));
        ipResult.deviceName = std::wstring(nm, nm + strlen(nm));
        g_queryCache[ipW] = ipResult;

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

            char* busPos = strstr(portPos, "<bus ");
            if (!busPos || busPos > portPos + 200) { portSearch = portPos + 1; continue; }

            char* slotSearch = busPos;
            char* slotEnd = busPos + 4096;
            if (slotEnd > buf + n) slotEnd = buf + n;

            while (slotSearch < slotEnd)
            {
                char* slotPos = strstr(slotSearch, "<address type=\"Short\" value=\"");
                if (!slotPos || slotPos >= slotEnd) break;
                char* sv = slotPos + 29;
                char* svEnd = strchr(sv, '"');
                if (!svEnd) { slotSearch = slotPos + 1; continue; }
                char slotStr[16] = {};
                int slotLen = (int)(svEnd - sv < 15 ? svEnd - sv : 15);
                memcpy(slotStr, sv, slotLen);
                int slotN = atoi(slotStr);

                char* slotDevPos = strstr(slotPos, "<device ");
                if (!slotDevPos || slotDevPos > slotPos + 200) { slotSearch = slotPos + 1; continue; }
                if (strncmp(slotDevPos + 8, "reference", 9) == 0) {
                    slotDevPos = strstr(slotDevPos + 1, "<device ");
                    if (!slotDevPos || slotDevPos > slotPos + 300) { slotSearch = slotPos + 1; continue; }
                }
                char scn[256] = {}, snm[256] = {};
                ExtractAttr(slotDevPos, "classname", scn, sizeof(scn));
                ExtractAttr(slotDevPos, "name", snm, sizeof(snm));

                std::wstring key = MakeCacheKey(ipW, portNameW, slotN);
                QueryResult sr;
                sr.found = true; sr.ip = ipW; sr.portName = portNameW; sr.slot = slotN;
                sr.classname = std::wstring(scn, scn + strlen(scn));
                sr.deviceName = std::wstring(snm, snm + strlen(snm));
                g_queryCache[key] = sr;

                slotSearch = slotDevPos + 1;
            }
            portSearch = portPos + 1;
        }
    }
}

// ---- Test harness ----

static int pass = 0, fail = 0;

static void Check(const char* desc, bool condition)
{
    if (condition) { printf("  [PASS] %s\n", desc); pass++; }
    else           { printf("  [FAIL] %s\n", desc); fail++; }
}

static void RunTests(const wchar_t* xmlFile)
{
    printf("\nXML: %ls\n\n", xmlFile);

    // === QueryXMLForPath tests ===
    printf("-- QueryXMLForPath --\n");

    // T1: IP-only query — known device
    {
        QueryResult r = QueryXMLForPath(xmlFile, L"192.168.1.55", L"", -1);
        Check("T1 found=true",               r.found);
        Check("T1 classname=1756-L85E/B",    r.classname == L"1756-L85E/B");
        Check("T1 deviceName has LOGIX5585E", r.deviceName.find(L"LOGIX5585E") != std::wstring::npos);
    }
    {
        QueryResult r = QueryXMLForPath(xmlFile, L"192.168.1.99", L"", -1);
        Check("T2 found=true (Unrecognized)", r.found);
        Check("T2 classname=Unrecognized Device", r.classname == L"Unrecognized Device");
    }
    {
        QueryResult r = QueryXMLForPath(xmlFile, L"192.168.1.55", L"Backplane", 0);
        Check("T3 Backplane/0 found",         r.found);
        Check("T3 classname=1756-L85E/B",     r.classname == L"1756-L85E/B");
        Check("T3 deviceName=1756-L85E",      r.deviceName == L"1756-L85E");
    }
    {
        QueryResult r = QueryXMLForPath(xmlFile, L"192.168.1.55", L"Backplane", 1);
        Check("T4 Backplane/1 classname=1756-EN2T/D", r.classname == L"1756-EN2T/D");
    }
    {
        QueryResult r = QueryXMLForPath(xmlFile, L"192.168.1.55", L"Backplane", 2);
        Check("T5 Backplane/2 classname=1756-IA16I/B", r.classname == L"1756-IA16I/B");
    }
    {
        QueryResult r = QueryXMLForPath(xmlFile, L"10.0.0.1", L"", -1);
        Check("T6 non-existent IP → not found", !r.found);
    }
    {
        QueryResult r = QueryXMLForPath(xmlFile, L"192.168.1.55", L"Backplane", 99);
        Check("T7 non-existent slot → not found", !r.found);
    }

    // === PopulateQueryCache tests ===
    printf("\n-- PopulateQueryCache (cache-first lookups) --\n");
    g_queryCache.clear();
    PopulateQueryCache(xmlFile);
    printf("  Cache entries: %d\n", (int)g_queryCache.size());

    {
        auto it = g_queryCache.find(L"192.168.1.55");
        Check("C1 IP key exists",            it != g_queryCache.end());
        if (it != g_queryCache.end()) {
            Check("C1 classname=1756-L85E/B", it->second.classname == L"1756-L85E/B");
        }
    }
    {
        auto it = g_queryCache.find(L"192.168.1.99");
        Check("C2 Unrecognized IP in cache", it != g_queryCache.end());
        if (it != g_queryCache.end())
            Check("C2 classname=Unrecognized Device", it->second.classname == L"Unrecognized Device");
    }
    {
        auto it = g_queryCache.find(L"192.168.1.55\\Backplane\\0");
        Check("C3 Backplane\\0 in cache",    it != g_queryCache.end());
        if (it != g_queryCache.end())
            Check("C3 classname=1756-L85E/B", it->second.classname == L"1756-L85E/B");
    }
    {
        auto it = g_queryCache.find(L"192.168.1.55\\Backplane\\1");
        Check("C4 Backplane\\1 classname=1756-EN2T/D",
              it != g_queryCache.end() && it->second.classname == L"1756-EN2T/D");
    }
    {
        auto it = g_queryCache.find(L"192.168.1.55\\Backplane\\2");
        Check("C5 Backplane\\2 classname=1756-IA16I/B",
              it != g_queryCache.end() && it->second.classname == L"1756-IA16I/B");
    }
    {
        auto it = g_queryCache.find(L"192.168.1.55\\Backplane\\3");
        Check("C6 Backplane\\3 classname=Unrecognized Device",
              it != g_queryCache.end() && it->second.classname == L"Unrecognized Device");
    }
    {
        Check("C7 non-existent key absent",
              g_queryCache.find(L"10.0.0.1") == g_queryCache.end());
    }
}

int main()
{
    RunTests(L"C:\\temp\\test_topo.xml");

    printf("\n--- Results: %d passed, %d failed ---\n", pass, fail);
    return fail > 0 ? 1 : 0;
}
