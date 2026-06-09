# Reproducible portable build for TSMIS Report Consolidator.
#
# Produces a self-contained onefolder under dist\<AppName>\ that bundles Python
# and every dependency -- no installer and no Python required on the target
# machine. No browser automation of any kind (this app only reads files already
# exported by the TSMIS Reports Exporter). Zip that folder to distribute.
#
# Usage (from the repo root):
#   powershell -ExecutionPolicy Bypass -File build\build.ps1
#
# See build\app.spec for the packaging recipe.

param([switch]$SelfTest)   # -SelfTest builds a headless self-test instead of the windowed app

$ErrorActionPreference = "Stop"

$BuildDir = $PSScriptRoot
$RepoRoot = Split-Path -Parent $BuildDir
$VenvDir  = Join-Path $BuildDir ".venv"
$VenvPy   = Join-Path $VenvDir  "Scripts\python.exe"
$WorkDir  = Join-Path $BuildDir "pyi-work"
$DistDir  = Join-Path $RepoRoot "dist"

function Assert-LastExit($what) {
    if ($LASTEXITCODE -ne 0) { throw "$what failed (exit $LASTEXITCODE)" }
}

# --- 1. Isolated build venv ----------------------------------------------
if (-not (Test-Path $VenvPy)) {
    Write-Host "==> Creating build venv"
    python -m venv $VenvDir; Assert-LastExit "venv creation"
}
Write-Host "==> Installing pinned build dependencies"
& $VenvPy -m pip install --upgrade pip --quiet; Assert-LastExit "pip upgrade"
& $VenvPy -m pip install -r (Join-Path $RepoRoot "requirements-build.txt"); Assert-LastExit "pip install"

# --- 2. Package as a portable onefolder -----------------------------------
if ($SelfTest) {
    # Headless self-test (console): writes a minimal PDF, runs pdfplumber
    # text/word extraction (the Ramp Summary path), an openpyxl round-trip, a
    # real consolidator run over synthetic inputs, and constructs the GUI window
    # withdrawn. Verifies the pruned frozen bundle still runs every real code
    # path -- without a visible window or a blocking mainloop.
    $env:TSMIS_ENTRY    = Join-Path $BuildDir "full_smoke.py"
    $env:TSMIS_APP_NAME = "TSMIS Consolidator SelfTest"
    $env:TSMIS_CONSOLE  = "1"
} else {
    # The real windowed deliverable.
    $env:TSMIS_ENTRY    = Join-Path $RepoRoot "scripts\gui_main.py"
    $env:TSMIS_APP_NAME = "TSMIS Consolidator"
    $env:TSMIS_CONSOLE  = "0"
}

Write-Host "==> Running PyInstaller"
& $VenvPy -m PyInstaller (Join-Path $BuildDir "app.spec") `
    --distpath $DistDir --workpath $WorkDir --noconfirm
Assert-LastExit "PyInstaller"

# --- 3. Trim to runtime-only files + DLP guard ----------------------------
# Bundled third-party docs are the proven corporate-DLP risk (an upstream
# markdown example once carried a test credit-card number that SharePoint
# blocked). Strip non-runtime files and FAIL the build if any doc or
# credit-card-like content remains, so a release can never reintroduce it.
$AppDir = Join-Path $DistDir $env:TSMIS_APP_NAME
Write-Host "==> Pruning bundle to runtime-only files and scanning for DLP-blocked content"
& (Join-Path $BuildDir "prune_bundle.ps1") -Target $AppDir

# --- 3b. Run the frozen self-test (the real release gate) -----------------
# Building the self-test exe only proves it links; RUN it so -SelfTest actually
# verifies the PRUNED frozen bundle exercises every real code path
# (pdfplumber, openpyxl, the consolidators, the GUI). Nonzero exit fails the build.
if ($SelfTest) {
    $SelfTestExe = Join-Path $AppDir ("{0}.exe" -f $env:TSMIS_APP_NAME)
    Write-Host "==> Running frozen self-test: $SelfTestExe"
    & $SelfTestExe
    Assert-LastExit "frozen self-test"
    Write-Host "==> Frozen self-test PASSED (pruned bundle runs every code path)."
}

# --- 4. Report ------------------------------------------------------------
if (-not $SelfTest) {
    Copy-Item (Join-Path $BuildDir "dist_readme.txt") (Join-Path $AppDir "Start Here.txt") -Force
}
$SizeMB = (Get-ChildItem $AppDir -Recurse -File | Measure-Object Length -Sum).Sum / 1MB
Write-Host ("`n==> Built {0}  ({1:N0} MB onefolder)" -f $AppDir, $SizeMB)
Write-Host "    Zip this folder to distribute (right-click -> Send to -> Compressed folder)."
