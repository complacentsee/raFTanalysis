param(
    [switch]$Debug
)

# RSLinxHook Stress Test - 50 cycle kill/clear/browse/query
#
# Requires RSLinx to be running with the driver already browsed (manually or
# from a prior run). Uses --monitor --debug-xml to snapshot the existing
# topology as a baseline, then validates subsequent stress cycles against it.
#
# All browse cycles use --debug-xml so topology XML files are preserved
# in $LOGDIR for debugging.

# ---- TESTBENCH CONFIGURATION ----
$DRIVER_NAME = "ATL"
# ---- END CONFIGURATION ----

. (Join-Path $PSScriptRoot "stress_common.ps1")

$CYCLES      = 50
$BROWSE_EXE  = Join-Path $PSScriptRoot "RSLinxBrowse\Release\RSLinxBrowse.exe"
$HARMONY_HRC = "C:\Program Files (x86)\Rockwell Software\RSCommon\Harmony.hrc"
$HARMONY_RSH = "C:\Program Files (x86)\Rockwell Software\RSCommon\Harmony.rsh"
$LOGDIR      = "C:\temp"
$LOGFILE     = "C:\temp\stress_results.txt"
$SVC_NAME    = "RSLinx"   # RSLinx Classic service name

$totalPass = 0
$totalFail = 0

function Log($msg) {
    $ts   = Get-Date -Format "HH:mm:ss"
    $line = "[$ts] $msg"
    Write-Host $line
    Add-Content $LOGFILE $line
}

function DeleteIfExists($path) {
    if (Test-Path $path) {
        Remove-Item $path -Force -ErrorAction SilentlyContinue
        if (Test-Path $path) { Log "      WARN: could not delete $path" }
        else                  { Log "      Deleted: $path" }
    } else {
        Log "      (not present): $path"
    }
}

function ClearNodeTable($regPath) {
    if (Test-Path $regPath) {
        $vals = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
        $vals.PSObject.Properties |
            Where-Object { $_.Name -notmatch '^PS' } |
            ForEach-Object {
                Remove-ItemProperty -Path $regPath -Name $_.Name -Force -ErrorAction SilentlyContinue
            }
        Log "      Node Table cleared: $regPath"
    } else {
        Log "      (Node Table not present): $regPath"
    }
}

function ShowProgress($cycle, $cycles, $pass, $fail) {
    $done     = $pass + $fail
    $cyclePct = [int](100.0 * $cycle / $cycles)
    $passRate = if ($done -gt 0) { "$([math]::Round(100.0 * $pass / $done, 1))%" } else { "n/a" }
    $filled   = [int]($cyclePct / 5)
    $bar      = "[" + ("#" * $filled).PadRight(20, '-') + "]"
    $color    = if ($fail -eq 0) { "Green" } else { "Yellow" }
    $ts       = Get-Date -Format "HH:mm:ss"
    $line     = "[$ts]   $bar $cyclePct%  Cycle $cycle/$cycles  Pass=$pass  Fail=$fail  SuccessRate=$passRate"
    Write-Host $line -ForegroundColor $color
    Add-Content $LOGFILE $line
}

# Clear log
if (Test-Path $LOGFILE) { Remove-Item $LOGFILE -Force }

# --- Resolve Node Table path and IPs dynamically ---
$NODE_TABLE = Find-NodeTablePath $DRIVER_NAME
if (-not $NODE_TABLE) {
    Log "FATAL: Driver '$DRIVER_NAME' not found in registry."
    exit 1
}

$TARGET_IPS = Read-NodeTableIPs $NODE_TABLE
if ($TARGET_IPS.Count -eq 0) {
    Log "FATAL: No IPs found in Node Table for driver '$DRIVER_NAME'."
    exit 1
}

# --- Baseline snapshot: read existing topology (RSLinx must be running + browsed) ---
$QUERIES = Do-BaselineSnapshot `
    -BrowseExe  $BROWSE_EXE `
    -DriverName $DRIVER_NAME `
    -TargetIPs  $TARGET_IPS `
    -LogDir     $LOGDIR

if (-not $QUERIES -or $QUERIES.Count -eq 0) {
    Log "FATAL: Baseline snapshot failed -- cannot establish expectations"
    exit 1
}

# Build query set: all GM entries for reachable IPs (IP-level + backplane slots)
$ipSet = @{}
foreach ($tip in $TARGET_IPS) { $ipSet[$tip] = $true }

$qPaths   = @()
$qExpects = @()
foreach ($k in $QUERIES.Keys) {
    $v = $QUERIES[$k]
    # Extract the IP (first segment before backslash)
    $qIp = if ($k -match '\\') { $k.Split('\')[0] } else { $k }
    if (-not $ipSet.ContainsKey($qIp)) { continue }
    $qPaths   += $k
    $qExpects += $v
}

$qCount = $qPaths.Count
$gmCount = $QUERIES.Count
Log "=== RSLinxHook Stress Test: $CYCLES cycles, $qCount queries each (from $gmCount GM entries) ==="
Log "Target: $($TARGET_IPS.Count) IPs  Driver: $DRIVER_NAME  Service: $SVC_NAME"
Log ""

for ($cycle = 1; $cycle -le $CYCLES; $cycle++) {
    Log "------------------------------------------------------------"
    Log "CYCLE $cycle / $CYCLES   (running total: $totalPass pass, $totalFail fail)"
    Log "------------------------------------------------------------"

    # --- Step 1: Stop RSLinx service ---
    Log "  [1] Stopping RSLinx service..."
    # Kill RSOBSERV first to release the injected hook DLL, then stop the service
    Stop-Process -Name "RSOBSERV" -Force -ErrorAction SilentlyContinue
    Stop-Service -Name $SVC_NAME -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    # Wait up to 15s for the service to stop
    $stopDeadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $stopDeadline) {
        $svc = Get-Service -Name $SVC_NAME
        if ($svc.Status -eq "Stopped") { break }
        Start-Sleep -Seconds 1
    }
    $svc = Get-Service -Name $SVC_NAME
    if ($svc.Status -ne "Stopped") {
        Log "      WARN: service still $($svc.Status), force-killing processes..."
        Stop-Process -Name "RSLinx" -Force -ErrorAction SilentlyContinue
        Stop-Process -Name "RSOBSERV" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
    }
    Log "      Service stopped."

    # --- Step 2: Delete harmony files and clear node table ---
    Log "  [2] Removing harmony files and clearing node table..."
    DeleteIfExists $HARMONY_HRC
    DeleteIfExists $HARMONY_RSH
    ClearNodeTable $NODE_TABLE

    # --- Step 3: Start RSLinx service (synchronous) ---
    Log "  [3] Starting RSLinx service..."
    Start-Service -Name $SVC_NAME -ErrorAction SilentlyContinue
    $svc = Get-Service -Name $SVC_NAME
    if ($svc.Status -ne "Running") {
        Log "  SKIP: Service failed to start ($($svc.Status)) -- $($qPaths.Count) failures"
        $totalFail += $qPaths.Count
        ShowProgress $cycle $CYCLES $totalPass $totalFail
        continue
    }
    # Wait for RSLinx COM objects to be ready (8s prevents hot-load race on fresh start)
    Start-Sleep -Seconds 8
    $rslinxProc = Get-Process -Name "RSLinx" -ErrorAction SilentlyContinue
    Log "      Service running, PID: $($rslinxProc.Id)"

    # --- Step 4: Inject + browse (--debug-xml preserves XML for debugging) ---
    Log "  [4] Running browse (fresh injection)..."
    $t0        = Get-Date
    $ipArgs    = $TARGET_IPS | ForEach-Object { "--ip"; $_ }
    if ($Debug) {
        $browseCmd = ('"' + $BROWSE_EXE + '" --debug-xml --driver ' + $DRIVER_NAME + ' ' + ($ipArgs -join ' ') + ' --logdir ' + $LOGDIR)
        Log "      CMD: $browseCmd"
    }
    $browseOut = & $BROWSE_EXE --debug-xml --driver $DRIVER_NAME @ipArgs --logdir $LOGDIR 2>&1
    $browseS   = [int]((Get-Date) - $t0).TotalSeconds
    $browseOk  = $browseOut | Where-Object { $_ -match 'Browse complete' }
    $devLine   = ($browseOut | Where-Object { $_ -match 'DEVICES_IDENTIFIED' } | Select-Object -First 1) -replace '^\s+',''

    $browseExitCode = $LASTEXITCODE
    if ($browseOk -and $browseExitCode -eq 0) {
        Log "      Browse OK in ${browseS}s -- $devLine"
    } else {
        $tail = ($browseOut | Select-Object -Last 4) -join " | "
        Log "      Browse FAILED after ${browseS}s (exit=$browseExitCode): $tail"
        # Capture hook log for diagnosis
        if (Test-Path "C:\temp\hook_log.txt") {
            $hookTail = Get-Content "C:\temp\hook_log.txt" -Tail 15
            Log "      HOOK LOG (last 15 lines):"
            $hookTail | ForEach-Object { Log "        | $_" }
        }
        $totalFail += $qPaths.Count
        ShowProgress $cycle $CYCLES $totalPass $totalFail
        continue
    }

    # --- Step 5: Query hook for each GM entry (FOUND/NOTFOUND status check) ---
    Log "  [5] Querying ($($qPaths.Count) checks)..."
    $cycleOk   = $true
    $cyclePass = 0
    $cycleFail = 0

    for ($qi = 0; $qi -lt $qPaths.Count; $qi++) {
        $qPath   = $qPaths[$qi]
        $qExpect = $qExpects[$qi]
        $wantFound = $qExpect -match '^FOUND'

        $qOutFile = "C:\temp\query_diag.txt"
        if ($Debug) {
            $qCmd = ('"' + $BROWSE_EXE + '" --query "' + $qPath + '" --logdir ' + $LOGDIR)
            Log "      CMD: $qCmd"
        }
        Start-Process -FilePath $BROWSE_EXE `
            -ArgumentList "--query", $qPath, "--logdir", $LOGDIR `
            -RedirectStandardOutput $qOutFile `
            -RedirectStandardError  "C:\temp\query_diag_err.txt" `
            -NoNewWindow -Wait
        $qOutRaw = if (Test-Path $qOutFile) { Get-Content $qOutFile -Raw } else { "" }
        $result  = (($qOutRaw -split "`n") | Where-Object { $_ -match '^\[FOUND\]|^\[NOTFOUND\]' } | Select-Object -First 1) -replace '\r',''

        $gotFound = $result -match '^\[FOUND\]'

        if ($wantFound -eq $gotFound) {
            Log "      PASS  $qPath"
            $cyclePass++
            $totalPass++
        } else {
            $got = if ($result) { $result } else { "(no output)" }
            $want = if ($wantFound) { "FOUND" } else { "NOTFOUND" }
            Log "      FAIL  $qPath  want=$want  got=$got"
            $cycleFail++
            $totalFail++
            $cycleOk = $false
        }
    }

    $status = if ($cycleOk) { "PASSED" } else { "FAILED" }
    Log "  => Cycle $cycle $status  (pass=$cyclePass fail=$cycleFail)"
    ShowProgress $cycle $CYCLES $totalPass $totalFail
    Log ""
}

# --- Final summary ---
$total  = $totalPass + $totalFail
$pctStr = if ($total -gt 0) { "$([math]::Round(100.0 * $totalFail / $total, 1))" + "%" } else { "n/a" }
Log "============================================================"
Log "STRESS TEST COMPLETE: $CYCLES cycles, $total total checks"
Log "  PASS: $totalPass"
Log "  FAIL: $totalFail  ($pctStr failure rate)"
if ($totalFail -eq 0) {
    Log "  RESULT: ALL PASS"
} else {
    Log "  RESULT: FAILURES DETECTED -- review log above"
}
Log "  Log: $LOGFILE"
Log "============================================================"
