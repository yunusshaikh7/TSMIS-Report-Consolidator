"""Shared core for the XLSX-input consolidators.

Three reports export a single sheet with a header row (row 1) followed by data
rows (row 2+): TSAR: Ramp Detail, Highway Sequence Listing, and Highway Log.
Combining them is identical except for four values -- the input subfolder, the
output file name, the exact sheet name, and the friendly report name -- so that
common logic lives here once. A fix (or a tweak to the combined layout) now
benefits all three instead of drifting across near-duplicate copies.

The TSAR: Ramp Summary consolidator deliberately does NOT use this: it parses
PDFs with report-specific column/wrap logic, not XLSX, so it stays standalone.

Console-free, like the rest of the core: progress is reported via Events.on_log,
overwrite is confirmed through the confirm_overwrite(path)->bool callback, the run
honors events.is_cancelled(), and a ConsolidateResult is returned. Never
prints/prompts/exits. Third-party imports are guarded so importing this module
(and therefore each thin wrapper) never fails when a dependency is missing -- the
caller gets a clean error result instead of an ImportError.
"""
import re
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    _DEPS_OK = True
except ImportError:
    _DEPS_OK = False

from events import ConsolidateResult, Events

# Pull the route token out of "<prefix>_route_<ROUTE>.xlsx".
ROUTE_FROM_NAME = re.compile(r"_route_(\w+)\.xlsx$", re.IGNORECASE)


def _extract_route(path):
    m = ROUTE_FROM_NAME.search(path.name)
    return m.group(1).upper() if m else path.stem


def _read_header(path, sheet_name):
    """Return row 1 of `sheet_name`, or None if the file isn't shaped like this
    report's export (sheet missing or empty)."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return None
        ws = wb[sheet_name]
        row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        return list(row) if row else None
    finally:
        wb.close()


def _stream_data_rows(path, sheet_name):
    """Yield each data row (row 2 onwards) of `sheet_name` as a tuple."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, values_only=True):
            yield row
    finally:
        wb.close()


def consolidate_xlsx(*, input_dir, out_path, sheet_name, report_name, title,
                     events=None, confirm_overwrite=None):
    """Combine every per-route XLSX in `input_dir` (reading worksheet
    `sheet_name`) into one workbook at `out_path`, prepending a "Route" column so
    rows from different routes stay distinguishable.

    input_dir / out_path are pathlib.Path (strings accepted). report_name is the
    friendly name used in user-facing messages (UI-neutral). title is the banner
    shown in the log.
    """
    input_dir, out_path = Path(input_dir), Path(out_path)
    events = events or Events()
    if not _DEPS_OK:
        return ConsolidateResult(
            status="error",
            message="Required components are missing (openpyxl).",
        )
    confirm = confirm_overwrite or (lambda _p: True)

    if not input_dir.exists():
        return ConsolidateResult(
            status="error",
            message=(f"The {report_name} input folder doesn't exist:\n{input_dir}\n\n"
                     f"Export the {report_name} report first, then consolidate."),
        )

    files = sorted(input_dir.glob("*.xlsx"))
    if not files:
        return ConsolidateResult(
            status="error",
            message=(f"No {report_name} files were found in:\n{input_dir}\n\n"
                     f"Export the {report_name} report first, then consolidate."),
        )

    # Confirm overwrite before reading any inputs.
    if out_path.exists() and not confirm(out_path):
        return ConsolidateResult(status="cancelled", message="Cancelled. Existing file kept.")

    events.on_log("=" * 60)
    events.on_log(f"{title} - {len(files)} file(s)")
    events.on_log("=" * 60)
    events.on_log("")

    # Lock in the first readable file's header as the canonical layout. Files that
    # disagree are skipped so misaligned columns can't silently corrupt the
    # combined workbook.
    canonical_header = None
    canonical_source = None
    for p in files:
        try:
            h = _read_header(p, sheet_name)
        except Exception as e:
            events.on_log(f"  {p.name}: header read FAILED ({type(e).__name__}); skipping")
            continue
        if h is None:
            events.on_log(f"  {p.name}: sheet '{sheet_name}' missing; skipping")
            continue
        canonical_header = h
        canonical_source = p.name
        break

    if canonical_header is None:
        return ConsolidateResult(
            status="error",
            message=f"No readable {report_name} files were found in:\n{input_dir}.")

    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Style objects are built here (not at module scope) so importing this module
    # never touches openpyxl -- the GUI can import it even if a dep is missing and
    # get a clean error result instead of an ImportError.
    header_fill = PatternFill("solid", start_color="305496")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # write_only mode streams rows to disk so the consolidated file can be
    # hundreds of thousands of rows without exhausting memory.
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(sheet_name)
    ws.freeze_panes = "B2"

    ws.column_dimensions["A"].width = 9
    for i, _ in enumerate(canonical_header, start=2):
        ws.column_dimensions[get_column_letter(i)].width = 16

    header_values = ["Route"] + [v if v is not None else "" for v in canonical_header]
    header_cells = []
    for v in header_values:
        cell = WriteOnlyCell(ws, value=str(v))
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        header_cells.append(cell)
    ws.append(header_cells)

    used_files = 0
    appended_rows = 0
    skipped = []
    failed = []

    for i, p in enumerate(files, 1):
        if events.is_cancelled():
            ws.close()  # finalize the write_only lazy writer (closes its temp file)
            wb.close()  # then release the archive
            return ConsolidateResult(status="cancelled", message="Cancelled by user.")
        prefix = f"[{i:>3}/{len(files)}] {p.name}"
        try:
            h = _read_header(p, sheet_name)
        except Exception as e:
            events.on_log(f"{prefix} FAILED reading header ({type(e).__name__}): {e}")
            failed.append(p.name)
            continue
        if h is None:
            events.on_log(f"{prefix} skipped: sheet '{sheet_name}' missing")
            skipped.append(p.name)
            continue
        if h != canonical_header:
            events.on_log(f"{prefix} skipped: header differs from {canonical_source}")
            skipped.append(p.name)
            continue

        route = _extract_route(p)
        try:
            count = 0
            for row in _stream_data_rows(p, sheet_name):
                ws.append([route] + list(row))
                count += 1
            events.on_log(f"{prefix} +{count} rows  (route {route})")
            appended_rows += count
            used_files += 1
        except Exception as e:
            events.on_log(f"{prefix} FAILED reading rows ({type(e).__name__}): {e}")
            failed.append(p.name)

    events.on_log("")
    events.on_log("Writing consolidated workbook...")
    try:
        wb.save(out_path)
    except PermissionError:
        return ConsolidateResult(
            status="error",
            message=(f"Could not save {out_path.name}.\n\n"
                     "The file is probably open in Excel. Close it and try again.\n"
                     "(Your exported files were not changed.)"),
        )

    return ConsolidateResult(
        status="ok",
        output_path=str(out_path),
        summary_lines=[
            f"Files combined: {used_files}",
            f"Rows added:     {appended_rows}",
            f"Files skipped:  {len(skipped)} {skipped if skipped else ''}",
            f"Files failed:   {len(failed)} {failed if failed else ''}",
            f"Output file:    {out_path}",
        ],
    )
