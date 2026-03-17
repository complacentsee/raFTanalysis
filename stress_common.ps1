# stress_common.ps1 — shared helpers for stress test scripts
#
# Dot-source this file from stress_test.ps1, stress_rebrowse.ps1, and stress_monitor.ps1:
#   . (Join-Path $PSScriptRoot "stress_common.ps1")

function Find-NodeTablePath {
    <#
    .SYNOPSIS
    Find the registry Node Table path for an RSLinx AB_ETH driver by name.
    Enumerates HKLM:\...\Drivers\AB_ETH\AB_ETH-* subkeys and returns the one
    whose Name value matches $DriverName (case-insensitive).
    #>
    param([string]$DriverName)

    $basePath = "HKLM:\SOFTWARE\WOW6432Node\Rockwell Software\RSLinx\Drivers\AB_ETH"
    if (-not (Test-Path $basePath)) { return $null }

    foreach ($sub in Get-ChildItem $basePath -ErrorAction SilentlyContinue) {
        $name = (Get-ItemProperty $sub.PSPath -Name "Name" -ErrorAction SilentlyContinue).Name
        if ($name -and $name -eq $DriverName) {
            return Join-Path $sub.PSPath "Node Table"
        }
    }
    return $null
}

function Read-NodeTableIPs {
    <#
    .SYNOPSIS
    Read all IP addresses from an RSLinx driver Node Table registry key.
    Returns a string array of IPs.
    #>
    param([string]$NodeTablePath)

    $ips = @()
    if (-not (Test-Path $NodeTablePath)) { return $ips }

    $props = Get-ItemProperty $NodeTablePath -ErrorAction SilentlyContinue
    $props.PSObject.Properties |
        Where-Object { $_.Name -notmatch '^PS' } |
        ForEach-Object { $ips += $_.Value }
    return $ips
}

function Parse-TopologyXML {
    <#
    .SYNOPSIS
    Parse a topology XML file and build an ordered hashtable of query paths
    to expected results, matching the format used by stress test $QUERIES.

    Returns [ordered]@{ "IP" = "FOUND|classname"; "IP\Port\Slot" = "FOUND|classname"; ... }
    Also adds one NOTFOUND entry per IP: "IP\Backplane\99" = "NOTFOUND"

    Mirrors the C++ PopulateQueryCache logic in TopologyXML.cpp.
    #>
    param([string]$XmlPath)

    $queries = [ordered]@{}
    if (-not (Test-Path $XmlPath)) { return $queries }

    $raw = [System.IO.File]::ReadAllText($XmlPath)

    # Helper: extract attribute value from a position in the string
    function Get-Attr($text, $startPos, $attrName, $maxDist = 400) {
        $needle = " $attrName=`""
        $idx = $text.IndexOf($needle, $startPos, [System.StringComparison]::Ordinal)
        if ($idx -lt 0 -or ($idx - $startPos) -gt $maxDist) { return $null }
        $valStart = $idx + $needle.Length
        $valEnd = $text.IndexOf('"', $valStart)
        if ($valEnd -lt 0) { return $null }
        return $text.Substring($valStart, $valEnd - $valStart)
    }

    # Helper: resolve <device reference="GUID"> by finding objectid="GUID" elsewhere
    function Resolve-Reference($text, $devPos) {
        $refAttr = Get-Attr $text $devPos "reference" 200
        if (-not $refAttr) { return $null }
        $oidPattern = "objectid=`"$refAttr`""
        $oidIdx = $text.IndexOf($oidPattern, [System.StringComparison]::Ordinal)
        if ($oidIdx -lt 0) { return $null }
        # Walk back to '<'
        $ds = $oidIdx
        while ($ds -gt 0 -and $text[$ds] -ne '<') { $ds-- }
        $cn = Get-Attr $text $ds "classname"
        return $cn
    }

    # Walk all <address type="String" value="IP"> blocks
    $addrPattern = '<address type="String" value="'
    $pos = 0
    while ($true) {
        $pos = $raw.IndexOf($addrPattern, $pos, [System.StringComparison]::Ordinal)
        if ($pos -lt 0) { break }
        $pos += $addrPattern.Length
        $ipEnd = $raw.IndexOf('"', $pos)
        if ($ipEnd -lt 0) { break }
        $ip = $raw.Substring($pos, $ipEnd - $pos)
        $pos = $ipEnd

        # Skip internal 0.0.0.0 addresses
        if ($ip -eq '0.0.0.0') { continue }

        # Find top-level <device> within 500 chars
        $searchEnd = [Math]::Min($pos + 500, $raw.Length)
        $devIdx = $raw.IndexOf("<device ", $pos, [System.StringComparison]::Ordinal)
        if ($devIdx -lt 0 -or $devIdx -ge $searchEnd) { continue }

        # Skip <device reference="..."> — resolve or find next real device
        $cn = $null
        if ($raw.Substring($devIdx + 8, [Math]::Min(9, $raw.Length - $devIdx - 8)) -eq "reference") {
            $cn = Resolve-Reference $raw $devIdx
            if (-not $cn) {
                # Try next <device> tag
                $devIdx2 = $raw.IndexOf("<device ", $devIdx + 1, [System.StringComparison]::Ordinal)
                if ($devIdx2 -ge 0 -and $devIdx2 -lt $searchEnd) {
                    $cn = Get-Attr $raw $devIdx2 "classname"
                }
            }
        } else {
            $cn = Get-Attr $raw $devIdx "classname"
        }

        # Skip unrecognized devices — they are transient and unreliable for baseline
        if (-not $cn -or $cn -eq "Unrecognized Device") { continue }
        $queries[$ip] = "FOUND|$cn"

        # Walk ports -> buses -> slots, bounded by next IP block
        # Find the next top-level IP address (skip 0.0.0.0 entries which are internal)
        $portSearch = $devIdx
        $nextIpPos = $ipEnd + 1
        while ($true) {
            $nextIpPos = $raw.IndexOf($addrPattern, $nextIpPos, [System.StringComparison]::Ordinal)
            if ($nextIpPos -lt 0) { $nextIpPos = $raw.Length; break }
            $nextIpValStart = $nextIpPos + $addrPattern.Length
            if ($raw.Substring($nextIpValStart, [Math]::Min(7, $raw.Length - $nextIpValStart)) -ne '0.0.0.0') { break }
            $nextIpPos = $nextIpValStart
        }
        $portEnd = $nextIpPos

        while ($portSearch -lt $portEnd) {
            $portPos = $raw.IndexOf("<port ", $portSearch, [System.StringComparison]::Ordinal)
            if ($portPos -lt 0 -or $portPos -ge $portEnd) { break }

            $portName = Get-Attr $raw $portPos "name"
            if (-not $portName) { $portSearch = $portPos + 1; continue }

            # Find <bus> within 200 chars of port
            $busPos = $raw.IndexOf("<bus ", $portPos, [System.StringComparison]::Ordinal)
            if ($busPos -lt 0 -or ($busPos - $portPos) -gt 200) { $portSearch = $portPos + 1; continue }

            # Enumerate <address type="Short" value="N"> within bus, bounded by next IP
            $slotSearch = $busPos
            $slotEnd = $nextIpPos

            while ($slotSearch -lt $slotEnd) {
                $slotPattern = '<address type="Short" value="'
                $slotPos = $raw.IndexOf($slotPattern, $slotSearch, [System.StringComparison]::Ordinal)
                if ($slotPos -lt 0 -or $slotPos -ge $slotEnd) { break }

                $svStart = $slotPos + $slotPattern.Length
                $svEnd = $raw.IndexOf('"', $svStart)
                if ($svEnd -lt 0) { $slotSearch = $slotPos + 1; continue }
                $slotNum = $raw.Substring($svStart, $svEnd - $svStart)

                # Find <device> within 200 chars of slot
                $slotDevPos = $raw.IndexOf("<device ", $slotPos, [System.StringComparison]::Ordinal)
                if ($slotDevPos -lt 0 -or ($slotDevPos - $slotPos) -gt 200) { $slotSearch = $slotPos + 1; continue }

                $slotCn = $null
                if ($raw.Substring($slotDevPos + 8, [Math]::Min(9, $raw.Length - $slotDevPos - 8)) -eq "reference") {
                    $slotCn = Resolve-Reference $raw $slotDevPos
                } else {
                    $slotCn = Get-Attr $raw $slotDevPos "classname"
                }

                if ($slotCn -and $slotCn -ne "Unrecognized Device") {
                    $key = "$ip\$portName\$slotNum"
                    $queries[$key] = "FOUND|$slotCn"
                }

                $slotSearch = $slotDevPos + 1
            }

            $portSearch = $portPos + 1
        }

    }

    return $queries
}

function Do-BaselineSnapshot {
    <#
    .SYNOPSIS
    Run a full browse with --debug-xml to establish baseline query expectations.
    The XML files are preserved for debugging. RSLinx must be running.
    Returns $null on failure (caller should exit).
    #>
    param(
        [string]$BrowseExe,
        [string]$DriverName,
        [string[]]$TargetIPs,
        [string]$LogDir
    )

    Log '============================================================'
    Log 'BASELINE: Establish topology query expectations'
    Log '============================================================'

    # Check for existing baseline XML: project-level golden master first, then LogDir
    $xmlPath = $null
    $gmPath = Join-Path $PSScriptRoot 'GM_topo.xml'
    if (Test-Path $gmPath) {
        $xmlPath = $gmPath
    } else {
        $candidates = @('hook_topo_after.xml', 'hook_topo_monitor_final.xml')
        foreach ($candidate in $candidates) {
            $p = Join-Path $LogDir $candidate
            if (Test-Path $p) { $xmlPath = $p; break }
        }
    }

    if ($xmlPath) {
        Log ('  Reusing existing baseline XML: ' + $xmlPath)
    } else {
        # No existing XML -- do a full browse to create one
        $svc = Get-Service -Name 'RSLinx' -ErrorAction SilentlyContinue
        if (-not $svc -or $svc.Status -ne 'Running') {
            Log 'FATAL: No baseline XML found and RSLinx service is not running.'
            return $null
        }
        $rslinxProc = Get-Process -Name 'RSLinx' -ErrorAction SilentlyContinue
        Log '  No existing baseline XML found -- running full browse...'
        Log ('  RSLinx running, PID: ' + $rslinxProc.Id)

        $t0 = Get-Date
        $ipList = @()
        foreach ($oneIp in $TargetIPs) { $ipList += '--ip'; $ipList += $oneIp }
        $browseOut = & $BrowseExe --debug-xml --driver $DriverName @ipList --logdir $LogDir 2>&1
        $browseS = [int]((Get-Date) - $t0).TotalSeconds
        $browseOk = $browseOut | Where-Object { $_ -match 'Browse complete' }
        $devLine = ($browseOut | Where-Object { $_ -match 'DEVICES_IDENTIFIED' } | Select-Object -First 1) -replace '^\s+',''

        if ($browseOk -and $LASTEXITCODE -eq 0) {
            Log ('  Baseline browse OK in ' + $browseS + 's -- ' + $devLine)
        } else {
            $tail = ($browseOut | Select-Object -Last 4) -join ' :: '
            Log ('FATAL: Baseline browse FAILED after ' + $browseS + 's (exit=' + $LASTEXITCODE + '): ' + $tail)
            $hookLogPath = Join-Path $LogDir 'hook_log.txt'
            if (Test-Path $hookLogPath) {
                $hookTail = Get-Content $hookLogPath -Tail 15
                Log '  HOOK LOG (last 15 lines):'
                foreach ($hl in $hookTail) {
                    Log ('    ' + $hl)
                }
            }
            return $null
        }

        # Find the XML that was just written
        $candidates2 = @('hook_topo_after.xml', 'hook_topo_monitor_final.xml')
        foreach ($candidate in $candidates2) {
            $p = Join-Path $LogDir $candidate
            if (Test-Path $p) { $xmlPath = $p; break }
        }
    }

    if (-not $xmlPath) {
        Log ('FATAL: No topology XML found in ' + $LogDir + ' after snapshot')
        Log '  Files present:'
        $topoFiles = Get-ChildItem (Join-Path $LogDir 'hook_topo*') -ErrorAction SilentlyContinue
        foreach ($tf in $topoFiles) {
            Log ('    ' + $tf.Name + ' (' + $tf.Length + ' bytes)')
        }
        return $null
    }

    Log ('  Parsing topology XML: ' + $xmlPath)
    $queries = Parse-TopologyXML $xmlPath
    if ($queries.Count -eq 0) {
        Log ('FATAL: No queries generated from topology XML (' + $xmlPath + ')')
        $fsize = (Get-Item $xmlPath).Length
        Log ('  XML file size: ' + $fsize + ' bytes')
        return $null
    }
    $qcount = $queries.Count
    Log ('  Generated ' + $qcount + ' query expectations from topology')
    Log ''

    return $queries
}
