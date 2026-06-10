# CLAUDE.md — TSMIS Report Consolidator

A portable Windows desktop tool that combines per-route TSMIS (Caltrans
Transportation System Management Information System) report files — exported by
the companion **TSMIS Reports Exporter** — into one Excel workbook per report
type. **No login, no browser, no Playwright, no network**: only `pdfplumber`
(PDF parsing) and `openpyxl` (Excel read/write) over files already on disk.

This is the consolidation half of `yunusshaikh7/TSMIS-Reports-Exporter`, split
out as a standalone app. The consolidator cores are copied from there and kept
behaviorally identical (plus an input-folder override); fixes proven here should
be considered for porting back, and vice versa.

## Supported Reports

| # | Report | Input | Default input folder | Combined output (`output/`) |
|---|---|---|---|---|
| 1 | TSAR: Ramp Summary | PDF | `input/ramp_summary/` | `tsar_ramp_summary_consolidated.xlsx` |
| 2 | TSAR: Ramp Detail | XLSX | `input/ramp_detail/` | `tsar_ramp_detail_consolidated.xlsx` |
| 3 | Highway Sequence Listing | XLSX | `input/highway_sequence/` | `highway_sequence_consolidated.xlsx` |
| 4 | Highway Log | XLSX | `input/highway_log/` | `highway_log_consolidated.xlsx` |
| 5 | TSN Highway Log | PDF (per district) | `input/tsn_highway_log/` | `tsn_highway_log_consolidated.xlsx` (+ per-route conversions in `output/tsn_highway_log/`) |
| 6 | Compare: TSN vs TSMIS Highway Log | the consolidated workbooks of 4 + 5 | `output/` (override via picker) | `highway_log_comparison.xlsx` |

## Two Run Modes, One Core

The consolidator cores are **console-free** and back both:
- **`.bat` console flow** (development + fallback): `2. consolidate (combine
  reports).bat` → a menu → `python scripts/consolidate_<name>.py`.
- **Packaged GUI** (`scripts/gui_*.py`, Tkinter) — the shipped desktop app.

Only `cli.py` and `gui_*.py` touch `print`/`input`/widgets. Core code reports
via the `Events` sink (`scripts/events.py`), confirms overwrites through the
`confirm_overwrite(path)->bool` callback, honors `events.is_cancelled()`, and
returns a `ConsolidateResult` — never `print`/`input`/`sys.exit`. User-facing
strings from the core must be **UI-neutral** (no ".bat" names, no "this window").

**Threading:** all file work (PDF parsing, workbook writing) runs on a worker
thread (`gui_worker.py`); only the main thread touches Tk. Workers talk to the
UI through a `queue.Queue` drained by `root.after()`.

## The App

A **portable single-folder Windows desktop app** (bundled Python + deps +
Tkinter GUI; no installer, no Python needed on target): staff unzip one folder
and double-click the `.exe`.

**Design decisions (don't relitigate without reason):**
- **Packaging:** PyInstaller **onefolder**, shipped as a portable zip.
- **No browser / Playwright** — that's the whole point of this split. If a task
  seems to need one, it belongs in the Exporter repo.
- **Data location:** the packaged app reads `input/`, writes `output/` and logs **next to the
  `.exe`**, falling back to `%LOCALAPPDATA%\TSMIS Consolidator` if read-only.
  See `scripts/paths.py`.
- **Input folder is per-run plumbing:** every `consolidate()` accepts
  `input_dir=`/`out_path=` overrides; the GUI exposes a folder picker so users
  can point at the Exporter's output without copying files.

**Pinned versions:** `pdfplumber==0.11.9` (→ `pdfminer.six`), `openpyxl==3.1.5`,
`pyinstaller==6.20.0`, `pyinstaller-hooks-contrib==2026.5`. Python 3.11.

## Repository Layout

```
1. setup (one time).bat           # pip install -r requirements.txt
2. consolidate (combine reports).bat  # console menu (dev / fallback)
run app (GUI preview).bat         # dev launcher for the GUI
requirements.txt / -build.txt     # pinned runtime / build deps
version.py                        # app name/version (single source of truth)
scripts/
  paths.py            # frozen-aware paths: DATA_ROOT, INPUT_ROOT, OUTPUT_ROOT, LOG_DIR
  logging_setup.py    # rotating file log under LOG_DIR (every entry point calls it)
  events.py           # Events sink + ConsolidateResult
  cli.py              # console adapter: run_consolidate_cli (overwrite prompt, exit codes)
  reports.py          # SINGLE registry: CONSOLIDATE_REPORTS = [(label, module)]
  consolidate_xlsx_base.py    # shared XLSX consolidator core
  consolidate_ramp_summary.py # standalone (parses PDFs; audited workbook + Combined sheet)
  consolidate_{ramp_detail,highway_sequence,highway_log}.py  # thin wrappers over the base
  consolidate_tsn_highway_log.py  # standalone: TSN district PDFs -> TSMIS-format per-route XLSX -> combined
  compare_highway_log.py          # TSN-vs-TSMIS comparison report (vendor fidelity; see its docstring)
  gui_main.py / gui_app.py / gui_worker.py / gui_theme.py    # GUI entry / window / workers / styles
build/
  build.ps1           # one-command onefolder build (-SelfTest = headless verify gate)
  prune_bundle.ps1    # strip to runtime-only files + DLP guard (run by build.ps1)
  app.spec            # PyInstaller spec (pdf/excel only; excludes image libs; version-info + icon + manifest)
  app.ico / app.manifest / full_smoke.py / dist_readme.txt / .venv/ (git-ignored)
dist/                 # build output: dist/TSMIS Consolidator/ (git-ignored)
input/                # the user's exported files go here (.gitkeep stubs; contents git-ignored)
  ramp_summary/ ramp_detail/ highway_sequence/ highway_log/ tsn_highway_log/
output/               # everything the app writes (consolidated workbooks); contents git-ignored
```

Don't commit `input/`/`output/` contents (only the `.gitkeep` stubs), build
artifacts (`build/.venv`, `dist/`), or `.claude/` permission state.

## Architecture Notes

- **One shared XLSX core, per-report differences in the thin wrappers.** Ramp
  Detail / Highway Sequence / Highway Log differ only by `INPUT_DIR`,
  `OUT_PATH`, `SHEET_NAME`, `REPORT_NAME` — the logic lives once in
  `consolidate_xlsx_base.consolidate_xlsx`. Ramp Summary stays standalone (it
  parses a specific PDF layout, not XLSX).
- **Single report registry** (`reports.py`): `CONSOLIDATE_REPORTS = [(label,
  module)]`; each module exposes `consolidate()`, `INPUT_DIR`, `OUT_PATH`,
  `INPUT_GLOB`, `REPORT_NAME`. The GUI radios and console menu both follow it.
  (The `.bat` menu is hand-edited text.)
- **Header lock-in (XLSX core):** the first readable file's header row is the
  canonical layout; files that disagree are *skipped and reported*, never
  merged, so misaligned columns can't silently corrupt the combined workbook.
- **Route extraction** comes from the filename (`…_route_<ROUTE>.xlsx`), falling
  back to the file stem; Ramp Summary reads the route from the PDF title.
- **TSN Highway Log is a converter + consolidator.** The TSN district log
  (OTM52010) is a fixed-layout PDF in proportional Helvetica, parsed by
  x-position windows (`COLUMN_WINDOWS`) **character by character** — never by
  words, because adjacent columns print closer than word-segmentation
  tolerances (a City code starts ~2pt after the county odometer, fusing into
  `042.010LKPT`). Lines are clustered with a 3pt y tolerance (data rows wrap
  1pt), `* *` totals and the per-page header band are skipped, the centered
  `<district> <county> <route>` header switches context, and description lines
  attach to the data row above them. `_normalize_row` matches TSMIS number
  formats (MI `000.075`; T-W unpadded `36`). Output uses the
  **exact** TSMIS Highway Log sheet name + 31-column header (`TSMIS_HEADER`),
  TSN-only ADT columns dropped, so the combined workbook is column-compatible
  with the TSMIS `highway_log_consolidated.xlsx` for comparison. Previously
  converted files are cleared each run so the result mirrors the input PDFs.
- **The Highway Log comparison** (`compare_highway_log.py`) answers "did the
  vendor (TSN) correctly represent TSMIS" without false positives. Core ideas:
  county sections are detected by odometer resets and matched by postmile
  overlap (the files order counties differently); rows are paired exactly on
  (Location, Cnty Odom), then unique Location (absorbs odometer drift), then
  per-section for duplicates; the vendor is judged AT TSMIS BREAKPOINTS so
  TSN's finer segmentation (bridges, DVMS stations) can't create findings; a
  blank TSMIS cell makes no claim (TSN values there are informational);
  verified conventions are suppressed (TSN omits Sig Chg == Date of Rec, '+'
  placeholders, Med Wid zero padding, '(DVMS)' description notes); TSMIS L/R-
  suffix rows are separate alignment series TSN folds into LB/RB columns (not
  row-comparable); TSMIS rows dated after the newest date in the TSN file are
  classed "post-snapshot", not vendor errors. Every rule is listed on the
  report's Summary sheet.
- **write_only streaming** in the XLSX core keeps memory flat for
  hundreds-of-thousands-row outputs; openpyxl style objects are built inside
  functions (never at module scope) so importing a core never touches openpyxl —
  `_DEPS_OK` guards give a clean error result instead of an ImportError.

## Key Behaviors

- **Overwrite confirmation** happens *before* any input is read (GUI: dialog
  pre-resolved; console: Y/N prompt; EOF = No).
- **Cancel:** the GUI Cancel button sets an event; cores check
  `is_cancelled()` between files and return a clean `cancelled` result.
- **Per-file failures don't stop the run:** unreadable/mismatched files are
  recorded as skipped/failed and listed in the summary.
- **Excel-has-the-file-open** (`PermissionError` on save) is caught and turned
  into a clear "close it and try again" error message.

## Build & Packaging (portable onefolder)

From the repo root: `powershell -ExecutionPolicy Bypass -File build\build.ps1`
→ windowed `dist\TSMIS Consolidator\`. Add `-SelfTest` for a headless console
build that **builds AND runs** `full_smoke.py` over the pruned frozen bundle
(pdfplumber over a hand-written PDF, openpyxl round-trip, a real consolidator
run over synthetic inputs, GUI construction) — the release gate.

`app.spec` highlights:
- `collect_data_files('pdfminer')` + `collect_all('pdfplumber'/'openpyxl')` —
  the pdfminer CMap data is the classic frozen trap. `cryptography` is a hard
  pdfminer import and **must stay**.
- `excludes=['PIL','pypdfium2','pypdfium2_raw']` — image libs the runtime paths
  (text extraction + plain workbooks) don't need; proven safe by the frozen
  `-SelfTest` passing.
- **Trust metadata** (reduces IT/Defender/DLP false-positives on the unsigned
  exe): version-info resource from `version.py`, `app.ico`, `app.manifest`
  (`asInvoker`), `upx=False`. Code-signing is the only complete fix (not done).

**Bundle hygiene / DLP (`prune_bundle.ps1`):** strips the bundle to runtime-only
files and **fails the build** if DLP-blocked content remains (same guard as the
Exporter, minus the Playwright-specific pruning). Deletes all prose docs
bundle-wide (licenses kept), sanitizes `dist-info` METADATA to headers,
`tests/`/`*.pyi`, stray image-lib dirs. Guards: non-license docs, credit cards
(IIN + length + Luhn), PEM private keys, AWS keys, US SSNs. Re-runnable on an
extracted release: `prune_bundle.ps1 -Target "…\TSMIS Consolidator"`
(`-GuardOnly` to audit).

## Extending

**New report type:**
1. For an XLSX report, create `scripts/consolidate_<name>.py` wrapping
   `consolidate_xlsx_base.consolidate_xlsx` (set `INPUT_DIR`, `OUT_PATH`,
   `SHEET_NAME`, `REPORT_NAME`, `INPUT_GLOB`) like `consolidate_highway_log.py`;
   for a different input format, write standalone (like
   `consolidate_ramp_summary.py`) implementing
   `consolidate(events, confirm_overwrite, input_dir=None, out_path=None) ->
   ConsolidateResult` (console-free, `_DEPS_OK`-guarded imports, styles built
   inside functions).
2. Add the `__main__` → `run_consolidate_cli` block.
3. Add one `(label, module)` entry to `CONSOLIDATE_REPORTS` in `reports.py`
   (feeds the GUI) and a branch to `2. consolidate…bat`. (A converter that
   produces TSMIS-format files and then consolidates them fits the same
   contract — see `consolidate_tsn_highway_log.py`.)
4. List the module in `APP_MODULES` in `build/app.spec`.
5. Add `input/<name>/.gitkeep`, whitelist it in `.gitignore`.
6. Document in the table at the top.

## Conventions

- Keep core code console-free; messages UI-neutral (see *Two Run Modes*).
- Runtime deps pinned. End-user setup uses global `pip` (no venv); the build
  uses `build\.venv`.
- **No test suite** — true verification is running a consolidator over real
  exported files. Logic-level checks can run any consolidator over synthetic
  inputs (see `build/full_smoke.py` step 3) without Windows.
- Commit messages: short, imperative (`add route column width override`).

## Common Issues

| Symptom | Cause | Fix |
|---|---|---|
| "Required components are missing" | Setup not run | Run `1. setup…bat` (dev) — packaged builds bundle them |
| "No … files were found" | Wrong input folder | Point the folder picker at the per-route exports |
| Files skipped: "header differs" | Mixed layouts in one folder | Expected guard — remove the odd files or fix the export |
| Could not save (file open in Excel) | Output workbook is open | Close Excel, run again |
| Frozen build can't parse PDFs | pdfminer CMap data missing | Keep `collect_data_files('pdfminer')` in `app.spec` |
| Build: "GUARD FAILED" | A dep shipped DLP-blocked content | Extend the prune list in `prune_bundle.ps1` |
