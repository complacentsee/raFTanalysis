# RSLinxHook Stress Test - 50 cycle kill/clear/browse/query
#
# CONFIGURATION: edit the variables below for your testbench before running.
#
# $TARGET_IP    - IP of the primary Ethernet device to browse to
# $DRIVER_NAME  - RSLinx driver name (case-insensitive match)
# $QUERIES      - Ordered hashtable of query path -> expected result substring.
#                 Keys use $TARGET_IP so only the block below needs updating.
#
# Expected result format: "FOUND|<classname>" or "NOTFOUND"
# Backplane paths: "$TARGET_IP\Backplane\<slot>"

# ---- TESTBENCH CONFIGURATION ----
$TARGET_IP   = "192.0.2.1"        # replace with your device IP
$DRIVER_NAME = "MyDriver"          # replace with your RSLinx driver name

$QUERIES = [ordered]@{
    $TARGET_IP                          = "FOUND|<primary-device-classname>"
    "$TARGET_IP\Backplane\0"            = "FOUND|<slot0-classname>"
    "$TARGET_IP\Backplane\1"            = "FOUND|<slot1-classname>"
    "$TARGET_IP\Backplane\2"            = "FOUND|<slot2-classname>"
    "$TARGET_IP\Backplane\3"            = "FOUND|<slot3-classname>"
    "$TARGET_IP\Backplane\99"           = "NOTFOUND"
}
# ---- END CONFIGURATION ----

$CYCLES      = 50
$BROWSE_EXE  = Join-Path $PSScriptRoot "RSLinxBrowse\Release\RSLinxBrowse.exe"
$HARMONY_HRC = "C:\Program Files (x86)\Rockwell Software\RSCommon\Harmony.hrc"
$HARMONY_RSH = "C:\Program Files (x86)\Rockwell Software\RSCommon\Harmony.rsh"
$NODE_TABLE  = "HKLM:\SOFTWARE\WOW6432Node\Rockwell Software\RSLinx\Drivers\AB_ETH\AB_ETH-1\Node Table"
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

# Snapshot keys/values into plain arrays once so the loop never touches $QUERIES
# NOTE: do NOT use $expected or $EXPECTED as a loop variable — PS names are
# case-insensitive and would overwrite this hashtable.
$qPaths   = @($QUERIES.Keys)
$qExpects = @($QUERIES.Values)

# Clear log
if (Test-Path $LOGFILE) { Remove-Item $LOGFILE -Force }
Log "=== RSLinxHook Stress Test: $CYCLES cycles, $($qPaths.Count) checks each ==="
Log "Target: $TARGET_IP  Driver: $DRIVER_NAME  Service: $SVC_NAME"
Log ""

for ($cycle = 1; $cycle -le $CYCLES; $cycle++) {
    Log "------------------------------------------------------------"
    Log "CYCLE $cycle / $CYCLES   (running total: $totalPass pass, $totalFail fail)"
    Log "------------------------------------------------------------"

    # --- Step 1: Stop RSLinx service (synchronous) ---
    Log "  [1] Stopping RSLinx service..."
    Stop-Service -Name $SVC_NAME -Force -ErrorAction SilentlyContinue
    $svc = Get-Service -Name $SVC_NAME
    if ($svc.Status -ne "Stopped") {
        Log "      WARN: service still $($svc.Status) after Stop-Service, waiting 5s..."
        Start-Sleep -Seconds 5
    }
    # Also kill RSOBSERV if hanging around
    Stop-Process -Name "RSOBSERV" -Force -ErrorAction SilentlyContinue
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
        continue
    }
    # Wait for RSLinx COM objects to be ready (8s prevents hot-load race on fresh start)
    Start-Sleep -Seconds 8
    $rslinxProc = Get-Process -Name "RSLinx" -ErrorAction SilentlyContinue
    Log "      Service running, PID: $($rslinxProc.Id)"

    # --- Step 4: Inject + browse ---
    Log "  [4] Running browse (fresh injection)..."
    $t0        = Get-Date
    $browseOut = & $BROWSE_EXE --driver $DRIVER_NAME --ip $TARGET_IP --logdir $LOGDIR 2>&1
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
        continue
    }

    # --- Step 5: Run queries ---
    Log "  [5] Querying ($($qPaths.Count) checks)..."
    $cycleOk = $true

    for ($qi = 0; $qi -lt $qPaths.Count; $qi++) {
        $qPath   = $qPaths[$qi]
        $qExpect = $qExpects[$qi]

        $qOutFile = "C:\temp\query_diag.txt"
        $proc = Start-Process -FilePath $BROWSE_EXE `
            -ArgumentList "--query", $qPath, "--logdir", $LOGDIR `
            -RedirectStandardOutput $qOutFile `
            -RedirectStandardError  "C:\temp\query_diag_err.txt" `
            -NoNewWindow -Wait -PassThru
        $qOutRaw = if (Test-Path $qOutFile) { Get-Content $qOutFile -Raw } else { "" }
        $qErrRaw = if (Test-Path "C:\temp\query_diag_err.txt") { Get-Content "C:\temp\query_diag_err.txt" -Raw } else { "" }
        $result  = (($qOutRaw -split "`n") | Where-Object { $_ -match '^\[FOUND\]|^\[NOTFOUND\]' } | Select-Object -First 1) -replace '\r',''

        if ($result -and ($result -match [regex]::Escape($qExpect))) {
            Log "      PASS  $qPath  =>  $result"
            $totalPass++
        } else {
            $got = if ($result) { $result } else { "(no output)" }
            Log "      FAIL  $qPath  want: $qExpect  got: $got  exit=$($proc.ExitCode)"
            Log "        STDOUT: $($qOutRaw -replace '\r?\n',' | ')"
            Log "        STDERR: $($qErrRaw -replace '\r?\n',' | ')"
            $totalFail++
            $cycleOk = $false
        }
    }

    $status = if ($cycleOk) { "PASSED" } else { "FAILED" }
    Log "  => Cycle $cycle $status"
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
