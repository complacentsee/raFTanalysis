#include "Config.h"
#include "Logging.h"

// ============================================================
// Configuration parsing
// ============================================================

std::wstring Utf8ToWide(const char* utf8)
{
    if (!utf8 || !*utf8) return L"";
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
    if (wlen <= 0) return L"";
    std::wstring result(wlen - 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, &result[0], wlen);
    return result;
}

bool ReadConfigFromPipe(HookConfig& config)
{
    if (!g_pipeConnected) return false;
    char buf[4096];
    std::string accumulated;
    while (true) {
        DWORD bytesRead = 0;
        if (!ReadFile(g_hPipe, buf, sizeof(buf) - 1, &bytesRead, NULL) || bytesRead == 0)
            return false;
        buf[bytesRead] = '\0';
        accumulated += buf;
        // Process complete lines
        size_t pos;
        while ((pos = accumulated.find('\n')) != std::string::npos) {
            std::string line = accumulated.substr(0, pos);
            accumulated.erase(0, pos + 1);
            if (!line.empty() && line.back() == '\r') line.pop_back();
            if (line == "C|END") return !config.drivers.empty();
            if (line.length() >= 2 && line[0] == 'C' && line[1] == '|') {
                std::wstring wval = Utf8ToWide(line.c_str() + 2);
                if (wval == L"MODE=inject") config.mode = HookMode::Inject;
                else if (wval == L"MODE=monitor") config.mode = HookMode::Monitor;
                else if (wval.length() >= 7 && wval.substr(0, 7) == L"LOGDIR=") config.logDir = wval.substr(7);
                else if (wval == L"DEBUGXML=1") config.debugXml = true;
                else if (wval == L"PROBE=1") config.probeDispids = true;
                else if (wval.length() >= 7 && wval.substr(0, 7) == L"DRIVER=") config.drivers.push_back({wval.substr(7), {}, false});
                else if (wval == L"NEWDRIVER=1" && !config.drivers.empty()) config.drivers.back().newDriver = true;
                else if (wval.length() >= 3 && wval.substr(0, 3) == L"IP=" && !config.drivers.empty()) config.drivers.back().ipAddresses.push_back(wval.substr(3));
            }
        }
    }
}
