#include "TopologyXML.h"
#include "Logging.h"

// ============================================================
// TopologyXML globals
// ============================================================

std::map<std::wstring, DeviceInfo> g_deviceDetails;

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
                        info.name = nameW;
                        info.slot = -1;
                        info.ip = ipW;
                        g_deviceDetails[nameW] = info;
                    }
                }
            }
        }
    }
}

// ============================================================
// Tree-based topology functions (no XML dependency)
// ============================================================

static void CountDevicesRecursive(const TopoNode& node, TopologyCounts& counts)
{
    if (node.type == TopoNode::Device && !node.name.empty())
    {
        counts.totalDevices++;
        if (!node.classname.empty() &&
            node.classname != L"Unrecognized Device" &&
            node.classname != L"Workstation")
            counts.identifiedDevices++;
    }
    for (const auto& child : node.children)
        CountDevicesRecursive(child, counts);
}

TopologyCounts CountDevicesInTree(const TopoNode& root)
{
    TopologyCounts counts = { 0, 0 };
    CountDevicesRecursive(root, counts);
    return counts;
}

static bool FindTargetIPInTree(const TopoNode& node, const std::vector<std::wstring>& targetIPs)
{
    if (node.type == TopoNode::Device && node.addressType == L"String")
    {
        for (const auto& ip : targetIPs)
        {
            if (node.address == ip &&
                !node.classname.empty() &&
                node.classname != L"Unrecognized Device")
                return true;
        }
    }
    for (const auto& child : node.children)
    {
        if (FindTargetIPInTree(child, targetIPs))
            return true;
    }
    return false;
}

bool IsTargetIdentifiedInTree(const TopoNode& root, const std::vector<std::wstring>& targetIPs)
{
    return FindTargetIPInTree(root, targetIPs);
}
