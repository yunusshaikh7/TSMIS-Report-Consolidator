"""Single source of truth for the report registry.

Every report type appears here exactly once, so adding one is a one-place change
on the Python side: both the GUI (gui_app.py) and the console menu read this
list. (The `.bat` menu is static text and is still edited by hand — see
CLAUDE.md "Extending".)

Kept import-light and console-free: it only pulls in the consolidate_* modules,
whose third-party imports are guarded, so importing it never fails on a missing
dependency and never does any I/O.
"""
import consolidate_ramp_summary as _c_ramp_summary
import consolidate_ramp_detail as _c_ramp_detail
import consolidate_highway_sequence as _c_highway_seq
import consolidate_highway_log as _c_highway_log

# (menu label, module). Each module provides consolidate(), INPUT_DIR, OUT_PATH,
# INPUT_GLOB, and REPORT_NAME. Order here is the display order in the GUI and
# the numbering in the console menu.
CONSOLIDATE_REPORTS = [
    ("TSAR: Ramp Summary", _c_ramp_summary),
    ("TSAR: Ramp Detail", _c_ramp_detail),
    ("Highway Sequence Listing", _c_highway_seq),
    ("Highway Log", _c_highway_log),
]
