# RSLinxHook Monitor Mode Stress Test - 20 cycles, --monitor flag
#
# Requires RSLinx to be running with the driver already browsed. Uses
# --monitor --debug-xml to snapshot existing topology as a baseline, then
# validates monitor mode across N cycles.
#
# PURPOSE: Validates the --monitor code path across N runs. Monitor mode skips
# DoEngineHotLoad and node-table/harmony-file changes, browsing the existing
# driver in-place. Verify via hook_log.txt that no "hot-load" lines appear
# during the monitor-mode cycles.
#
# All browse cycles use --debug-xml so topology XML files are preserved
# in $LOGDIR for debugging.

# ---- TESTBENCH CONFIGURATION ----
$DRIVER_NAME = "ATL"
# ---- END CONFIGURATION ----

. (Join-Path $PSScriptRoot "stress_common.ps1")

$CYCLES      = 20
$BROWSE_EXE  = Join-Path $PSScriptRoot "RSLinxBrowse\Release\RSLinxBrowse.exe"
$LOGDIR      = "C:\temp"
$LOGFILE     = "C:\temp\stress_monitor_results.txt"
$SVC_NAME    = "RSLinx"   # RSLinx Classic service name

$totalPass = 0
$totalFail = 0

function Log($msg) {
    $ts   = Get-Date -Format "HH:mm:ss"
    $line = "[$ts] $msg"
    Write-Host $line
    Add-Content $LOGFILE $line
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

# Snapshot keys/values into plain arrays once so the loop never touches $QUERIES
# NOTE: do NOT use $expected or $EXPECTED as a loop variable — PS names are
# case-insensitive and would overwrite this hashtable.
$qPaths   = @($QUERIES.Keys)
$qExpects = @($QUERIES.Values)

Log "=== RSLinxHook Monitor Mode Stress Test: $CYCLES cycles, $($qPaths.Count) checks each ==="
Log "Target: $($TARGET_IPS.Count) IPs  Driver: $DRIVER_NAME  Service: $SVC_NAME"
Log ""

# ---- MAIN LOOP: monitor mode against the same running RSLinx process ----
for ($cycle = 1; $cycle -le $CYCLES; $cycle++) {
    Log "------------------------------------------------------------"
    Log "CYCLE $cycle / $CYCLES   (running total: $totalPass pass, $totalFail fail)"
    Log "------------------------------------------------------------"

    # --- Step 1: Verify RSLinx is still running ---
    Log "  [1] Verifying RSLinx is still running..."
    $svc = Get-Service -Name $SVC_NAME -ErrorAction SilentlyContinue
    if (-not $svc -or $svc.Status -ne "Running") {
        Log "  ABORT: RSLinx service is no longer running ($($svc.Status)) -- monitor-mode COM path may have caused a crash"
        $totalFail += $qPaths.Count
        ShowProgress $cycle $CYCLES $totalPass $totalFail
        break
    }
    $rslinxProc = Get-Process -Name "RSLinx" -ErrorAction SilentlyContinue
    Log "      RSLinx running, PID: $($rslinxProc.Id)"

    # --- Step 2: Monitor-mode browse with --debug-xml ---
    Log "  [2] Running --monitor browse (no node-table or harmony-file changes)..."
    $t0        = Get-Date
    $ipArgs    = $TARGET_IPS | ForEach-Object { "--ip"; $_ }
    $browseOut = & $BROWSE_EXE --monitor --debug-xml --driver $DRIVER_NAME @ipArgs --logdir $LOGDIR 2>&1
    $browseS   = [int]((Get-Date) - $t0).TotalSeconds
    $browseOk  = $browseOut | Where-Object { $_ -match 'Browse complete' }
    $devLine   = ($browseOut | Where-Object { $_ -match 'DEVICES_IDENTIFIED' } | Select-Object -First 1) -replace '^\s+',''

    $browseExitCode = $LASTEXITCODE
    if ($browseOk -and $browseExitCode -eq 0) {
        Log "      Browse OK in ${browseS}s -- $devLine"
    } else {
        $tail = ($browseOut | Select-Object -Last 4) -join " | "
        Log "      Browse FAILED after ${browseS}s (exit=$browseExitCode): $tail"
        if (Test-Path "C:\temp\hook_log.txt") {
            $hookTail = Get-Content "C:\temp\hook_log.txt" -Tail 15
            Log "      HOOK LOG (last 15 lines):"
            $hookTail | ForEach-Object { Log "        | $_" }
        }
        $totalFail += $qPaths.Count
        ShowProgress $cycle $CYCLES $totalPass $totalFail
        continue
    }

    # --- Step 3: Run queries ---
    Log "  [3] Querying ($($qPaths.Count) checks)..."
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
    ShowProgress $cycle $CYCLES $totalPass $totalFail
    Log ""
}

# --- Final summary ---
$total  = $totalPass + $totalFail
$pctStr = if ($total -gt 0) { "$([math]::Round(100.0 * $totalFail / $total, 1))" + "%" } else { "n/a" }
Log "============================================================"
Log "MONITOR MODE STRESS TEST COMPLETE: $CYCLES cycles, $total total checks"
Log "  PASS: $totalPass"
Log "  FAIL: $totalFail  ($pctStr failure rate)"
if ($totalFail -eq 0) {
    Log "  RESULT: ALL PASS"
} else {
    Log "  RESULT: FAILURES DETECTED -- review log above"
}
Log "  NOTE: verify C:\temp\hook_log.txt contains no 'hot-load' lines (monitor must skip DoEngineHotLoad)"
Log "  Log: $LOGFILE"
Log "============================================================"
