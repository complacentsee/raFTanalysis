<#
.SYNOPSIS
    Builds all raFTMEanalysis projects and stages output.

.DESCRIPTION
    1. Builds RSLinxBrowse + RSLinxHook (C++ Release|Win32) via MSBuild
    2. Builds RSLinxViewer (.NET 8 x86 self-contained) via dotnet publish
    3. Copies RSLinxHook.dll into the RSLinxViewer publish directory

.PARAMETER Clean
    Remove build output directories before building.

.PARAMETER SkipCpp
    Skip C++ MSBuild projects (only build RSLinxViewer).

.PARAMETER SkipViewer
    Skip RSLinxViewer (only build C++ projects).

.EXAMPLE
    .\build.ps1
    .\build.ps1 -Clean
    .\build.ps1 -SkipViewer
#>
param(
    [switch]$Clean,
    [switch]$SkipCpp,
    [switch]$SkipViewer
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

$releaseDir    = Join-Path $root 'Release'
$viewerPublish = Join-Path (Join-Path $root 'RSLinxViewer') 'publish'

# --- Clean ---
if ($Clean) {
    Write-Host '--- Cleaning build outputs ---' -ForegroundColor Yellow
    foreach ($dir in @($releaseDir, $viewerPublish)) {
        if (Test-Path $dir) {
            Remove-Item $dir -Recurse -Force
            Write-Host "  Removed $dir"
        }
    }
    $viewerBin = Join-Path $root 'RSLinxViewer' 'bin'
    $viewerObj = Join-Path $root 'RSLinxViewer' 'obj'
    foreach ($dir in @($viewerBin, $viewerObj)) {
        if (Test-Path $dir) {
            Remove-Item $dir -Recurse -Force
            Write-Host "  Removed $dir"
        }
    }
    Write-Host ''
}

# --- Locate MSBuild ---
if (-not $SkipCpp) {
    $msbuild = $null

    # Try vswhere first (VS2019+)
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $vsPath = & $vswhere -latest -requires Microsoft.Component.MSBuild -property installationPath 2>$null
        if ($vsPath) {
            $candidate = Join-Path $vsPath 'MSBuild\Current\Bin\MSBuild.exe'
            if (Test-Path $candidate) { $msbuild = $candidate }
        }
    }

    # Fallback: check PATH
    if (-not $msbuild) {
        $msbuild = Get-Command MSBuild.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
    }

    if (-not $msbuild) {
        Write-Error 'MSBuild.exe not found. Install Visual Studio with C++ workload or run from a Developer Command Prompt.'
    }

    Write-Host "MSBuild: $msbuild" -ForegroundColor Cyan
}

# --- Build C++ projects ---
if (-not $SkipCpp) {
    # Set SolutionDir so output goes to solution-level Release\
    $slnDir = "$root\"

    Write-Host ''
    Write-Host '=== Building RSLinxHook (Release|Win32) ===' -ForegroundColor Green
    $hookProj = Join-Path $root 'RSLinxHook\RSLinxHook.vcxproj'
    & $msbuild $hookProj /p:Configuration=Release /p:Platform=Win32 "/p:SolutionDir=$slnDir" /verbosity:minimal
    if ($LASTEXITCODE -ne 0) { Write-Error 'RSLinxHook build failed.' }

    Write-Host ''
    Write-Host '=== Building RSLinxBrowse (Release|Win32) ===' -ForegroundColor Green
    $browseProj = Join-Path $root 'RSLinxBrowse\RSLinxBrowse.vcxproj'
    & $msbuild $browseProj /p:Configuration=Release /p:Platform=Win32 "/p:SolutionDir=$slnDir" /verbosity:minimal
    if ($LASTEXITCODE -ne 0) { Write-Error 'RSLinxBrowse build failed.' }

    Write-Host ''
    Write-Host "C++ binaries in: $releaseDir" -ForegroundColor Cyan
    Write-Host '  RSLinxBrowse.exe + RSLinxHook.dll both land in Release\'
}

# --- Build RSLinxViewer ---
if (-not $SkipViewer) {
    Write-Host ''
    Write-Host '=== Publishing RSLinxViewer (win-x86 self-contained) ===' -ForegroundColor Green
    $viewerProj = Join-Path $root 'RSLinxViewer\RSLinxViewer.csproj'
    & dotnet publish $viewerProj -r win-x86 --self-contained -c Release -o $viewerPublish
    if ($LASTEXITCODE -ne 0) { Write-Error 'RSLinxViewer publish failed.' }

    # Copy RSLinxHook.dll into viewer publish directory
    $hookDll = Join-Path $releaseDir 'RSLinxHook.dll'
    if (Test-Path $hookDll) {
        Copy-Item $hookDll -Destination $viewerPublish -Force
        Write-Host ''
        Write-Host "Copied RSLinxHook.dll -> $viewerPublish" -ForegroundColor Cyan
    }
    elseif (-not $SkipCpp) {
        Write-Warning "RSLinxHook.dll not found at $hookDll - DLL copy skipped."
    }
    else {
        Write-Host ''
        Write-Host 'Skipped RSLinxHook.dll copy (C++ build was skipped).' -ForegroundColor Yellow
    }
}

# --- Summary ---
Write-Host ''
Write-Host '=== Build Complete ===' -ForegroundColor Green
if (-not $SkipCpp) {
    Write-Host "  C++ output:    $releaseDir"
    Write-Host '    RSLinxBrowse.exe'
    Write-Host '    RSLinxHook.dll'
}
if (-not $SkipViewer) {
    Write-Host "  Viewer output: $viewerPublish"
    Write-Host '    RSLinxViewer.exe (self-contained)'
    Write-Host '    RSLinxHook.dll (copied from C++ build)'
}
