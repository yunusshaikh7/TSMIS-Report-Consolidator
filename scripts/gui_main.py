"""GUI entry point for the TSMIS Report Consolidator.

Run in dev:   python scripts\\gui_main.py
Packaged:     this is the PyInstaller entry (build\\build.ps1 sets TSMIS_ENTRY
              here and TSMIS_CONSOLE=0 for a windowed app).
"""
import sys
from pathlib import Path


def _bootstrap():
    # Dev only: make the flat scripts/ modules and the repo-root version.py
    # importable regardless of the working directory. Frozen builds bundle these.
    if not getattr(sys, "frozen", False):
        here = Path(__file__).resolve().parent          # scripts/
        sys.path.insert(0, str(here))                   # events, gui_app, ...
        sys.path.insert(0, str(here.parent))            # version.py (repo root)


_bootstrap()

from gui_app import App  # noqa: E402  (must follow _bootstrap)


def main():
    from logging_setup import setup_logging
    setup_logging()
    App().mainloop()


if __name__ == "__main__":
    main()
