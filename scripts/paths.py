"""Frozen-aware filesystem paths for TSMIS Report Consolidator.

One place that decides WHERE the app reads and writes, so the rest of the
code never has to care whether it is running as a dev script or as the
packaged portable .exe.

Policy ("portable by default, never break"):
  * Packaged build (sys.frozen): write next to the .exe -- the intuitive
    "my reports are right here in the folder" model. If that folder is not
    writable (e.g. unzipped into Program Files or a read-only network share),
    fall back automatically to %LOCALAPPDATA%\\TSMIS Consolidator so the app
    still runs. Callers should surface DATA_ROOT in the UI so the rare
    fallback is never a mystery.
  * Dev / .bat workflow (not frozen): the repo root (./output), matching the
    layout the TSMIS Reports Exporter uses, so dropping exported files into
    output/<report>/ works identically in both apps.
"""
import os
import sys
from pathlib import Path

APP_NAME = "TSMIS Consolidator"


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def _writable(directory: Path) -> bool:
    """True if we can create a file in `directory` (creating it if needed)."""
    try:
        directory.mkdir(parents=True, exist_ok=True)
        probe = directory / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return True
    except OSError:
        return False


def _localappdata_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / APP_NAME


def _resolve_data_root() -> Path:
    """Base directory for everything the app reads and writes by default."""
    if is_frozen():
        exe_dir = Path(sys.executable).resolve().parent   # the onefolder app dir
        if _writable(exe_dir):
            return exe_dir
        fallback = _localappdata_dir()                     # read-only location
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback
    # Dev: repo root (this file lives in scripts/).
    return Path(__file__).resolve().parent.parent


# Resolved once at import time.
DATA_ROOT = _resolve_data_root()

# Default location of the per-route exports to combine: each report type reads
# its own subfolder under here, and the combined workbooks land in
# OUTPUT_ROOT / "consolidated". The GUI lets the user point a run at any other
# input folder (e.g. the TSMIS Reports Exporter's output) without changing this.
OUTPUT_ROOT = DATA_ROOT / "output"

# App-private data (logs).
_PRIVATE = (DATA_ROOT / "data") if is_frozen() else DATA_ROOT

LOG_DIR = _PRIVATE / "logs"
