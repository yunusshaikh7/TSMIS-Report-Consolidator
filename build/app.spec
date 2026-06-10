# PyInstaller spec for TSMIS Report Consolidator (portable onefolder).
#
# Driven by build\build.ps1, which sets these environment variables:
#   TSMIS_ENTRY     path to the entry-point .py to package
#   TSMIS_APP_NAME  output folder / exe name (e.g. "TSMIS Consolidator")
#   TSMIS_CONSOLE   "1" to show a console window, "0" for a windowed GUI app
#
# Recipe:
#   * NO browser automation: unlike the TSMIS Reports Exporter there is no
#     Playwright, no Node driver, and no browser -- only the PDF/Excel libraries.
#   * pdfminer ships CMap data files that must be collected or pdfplumber text
#     extraction breaks when frozen -> collect_data_files('pdfminer').
#   * The .exe carries a version-info resource (from version.py), an icon, and a
#     manifest (asInvoker). Those are trust signals that reduce Windows Defender /
#     corporate-IT (DLP/SmartScreen) false-positives on an unsigned build. Code
#     signing is still the only complete fix.
import os
import sys
from PyInstaller.utils.hooks import collect_all, collect_data_files

ENTRY    = os.environ.get("TSMIS_ENTRY", os.path.join(SPECPATH, "full_smoke.py"))
APP_NAME = os.environ.get("TSMIS_APP_NAME", "TSMIS Consolidator")
CONSOLE  = os.environ.get("TSMIS_CONSOLE", "1") == "1"

# The app uses flat modules in scripts/ (imported by bare name) plus version.py
# at the repo root. Put both on pathex so PyInstaller resolves them, and list
# them as hidden imports because several are imported lazily (inside functions).
REPO_ROOT = os.path.dirname(SPECPATH)               # build/ -> repo root
SCRIPTS   = os.path.join(REPO_ROOT, "scripts")

# --- Windows .exe metadata: version resource + icon + manifest ---------------
# The single source of truth for the version is version.py.
sys.path.insert(0, REPO_ROOT)
from version import __version__ as APP_VERSION       # noqa: E402

_parts  = (APP_VERSION.split(".") + ["0", "0", "0", "0"])[:4]
_vtuple = tuple(int(p) if p.isdigit() else 0 for p in _parts)

from PyInstaller.utils.win32.versioninfo import (   # noqa: E402
    VSVersionInfo, FixedFileInfo, StringFileInfo, StringTable, StringStruct,
    VarFileInfo, VarStruct,
)
VERSION_INFO = VSVersionInfo(
    ffi=FixedFileInfo(filevers=_vtuple, prodvers=_vtuple, mask=0x3F, flags=0x0,
                      OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
    kids=[
        StringFileInfo([StringTable("040904B0", [
            StringStruct("CompanyName", "TSMIS Report Consolidator"),
            StringStruct("FileDescription", "TSMIS Report Consolidator"),
            StringStruct("FileVersion", APP_VERSION),
            StringStruct("InternalName", APP_NAME),
            StringStruct("LegalCopyright", "Internal tool. Provided as-is, no warranty."),
            StringStruct("OriginalFilename", APP_NAME + ".exe"),
            StringStruct("ProductName", "TSMIS Report Consolidator"),
            StringStruct("ProductVersion", APP_VERSION),
        ])]),
        VarFileInfo([VarStruct("Translation", [0x0409, 1200])]),  # US English, Unicode
    ],
)

ICON     = os.path.join(SPECPATH, "app.ico")
MANIFEST = os.path.join(SPECPATH, "app.manifest")
APP_MODULES = [
    "version", "paths", "events", "logging_setup", "cli", "reports",
    "consolidate_xlsx_base", "consolidate_ramp_summary", "consolidate_ramp_detail",
    "consolidate_highway_sequence", "consolidate_highway_log",
    "consolidate_tsn_highway_log", "compare_highway_log",
    "gui_main", "gui_app", "gui_worker", "gui_theme",
]

datas, binaries, hiddenimports = [], [], list(APP_MODULES)

# Bundle the icon so the GUI can set the window/taskbar icon at runtime
# (resolved via sys._MEIPASS -> _internal/app.ico). Binary, so the DLP text scan
# in prune_bundle.ps1 skips it.
if os.path.exists(ICON):
    datas += [(ICON, ".")]

# PDF + Excel consolidators. pdfminer's CMap data is the known frozen-build trap.
datas += collect_data_files("pdfminer")
for _pkg in ("pdfplumber", "openpyxl"):
    _d, _b, _h = collect_all(_pkg)
    datas += _d; binaries += _b; hiddenimports += _h

# Drop optional image libraries the app never needs at runtime: Pillow (PIL) and
# pypdfium2 (pdfplumber.to_image). openpyxl imports Pillow EAGERLY when present,
# but the code paths the app actually uses -- text/table extraction and writing
# plain workbooks, never image insert or rasterizing a PDF -- don't need it, and
# openpyxl tolerates a missing Pillow. The proof is the FROZEN self-test
# (build.ps1 -SelfTest runs full_smoke.py) passing with PIL excluded. pypdfium2
# is only touched by pdfplumber.to_image, which the app never calls.
EXCLUDES = ["PIL", "pypdfium2", "pypdfium2_raw"]
_excl = set(EXCLUDES)
hiddenimports = [h for h in hiddenimports if h.split(".")[0] not in _excl]

a = Analysis(
    [ENTRY],
    pathex=[SCRIPTS, REPO_ROOT],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                                       # UPX-packed exes are a classic AV false-positive trigger
    console=CONSOLE,
    disable_windowed_traceback=False,
    icon=(ICON if os.path.exists(ICON) else None),
    version=VERSION_INFO,
    manifest=(MANIFEST if os.path.exists(MANIFEST) else None),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name=APP_NAME,
)
