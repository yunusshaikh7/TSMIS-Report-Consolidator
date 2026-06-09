"""Consolidate Highway Sequence Listing XLSX files into a single workbook.

Reads every XLSX in   input/highway_sequence/   (or a caller-chosen folder)
Writes one workbook in output/highway_sequence_consolidated.xlsx
with a leading "Route" column added so rows from different routes are
distinguishable in the combined file.

Thin wrapper over consolidate_xlsx_base, which is shared with Ramp Detail and
Highway Log -- all three are "one sheet, header row + data rows" exports that
differ only by input folder, sheet name, and output name. (The Ramp Summary
consolidator stays standalone because it parses PDFs, not XLSX.)

Importable: consolidate(events, confirm_overwrite) returns a ConsolidateResult
and never prints/prompts/exits, so the GUI can drive it. The console UX lives in
cli.run_consolidate_cli, used by the __main__ entry (and therefore by
"2. consolidate (combine reports).bat").
"""
from consolidate_xlsx_base import consolidate_xlsx
from paths import INPUT_ROOT, OUTPUT_ROOT

INPUT_DIR = INPUT_ROOT / "highway_sequence"
OUT_PATH = OUTPUT_ROOT / "highway_sequence_consolidated.xlsx"

# Sheet name produced by the TSMIS export — must match exactly.
SHEET_NAME = "Highway Locations"

# Friendly report name for user-facing messages (shown in both the GUI and the
# console, so keep it UI-neutral -- no ".bat" / "menu option" wording).
REPORT_NAME = "Highway Sequence"

# File pattern the GUI uses to preview how many inputs a folder holds.
INPUT_GLOB = "*.xlsx"


def consolidate(events=None, confirm_overwrite=None, input_dir=None, out_path=None):
    """Combine every per-route Highway Sequence XLSX into one workbook."""
    return consolidate_xlsx(
        input_dir=input_dir or INPUT_DIR, out_path=out_path or OUT_PATH,
        sheet_name=SHEET_NAME, report_name=REPORT_NAME,
        title="Highway Sequence Consolidation",
        events=events, confirm_overwrite=confirm_overwrite,
    )


if __name__ == "__main__":
    from cli import run_consolidate_cli
    run_consolidate_cli(consolidate)
