# TSMIS Report Consolidator

> Combine per-route Caltrans TSMIS report files into one Excel workbook — no login, no browser.

[![Version](https://img.shields.io/badge/version-0.1.0-blue)](version.py)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-0078D6?logo=windows)](#)
[![Python](https://img.shields.io/badge/python-3.11-3776AB?logo=python&logoColor=white)](#)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

A portable Windows desktop tool that combines the per-route report files
exported from the Caltrans **Transportation System Management Information
System (TSMIS)** — typically by the companion
[TSMIS Reports Exporter](https://github.com/yunusshaikh7/TSMIS-Reports-Exporter)
— into **one consolidated Excel workbook per report type**.

It is the consolidation half of the Exporter, shipped on its own: there is **no
sign-in, no browser automation, and no network access** — just PDF parsing
(`pdfplumber`) and Excel reading/writing (`openpyxl`) over files already on the
machine. Distributed as a single zip: unzip, double-click, done.

## Supported reports

| Report | Input | Default input folder | Combined output (`output/`) |
|---|---|---|---|
| TSAR: Ramp Summary | PDFs | `input/ramp_summary/` | `tsar_ramp_summary_consolidated.xlsx` (one audited row per route + a live "Combined" summary sheet) |
| TSAR: Ramp Detail | XLSX | `input/ramp_detail/` | `tsar_ramp_detail_consolidated.xlsx` (rows stacked, leading `Route` column) |
| Highway Sequence Listing | XLSX | `input/highway_sequence/` | `highway_sequence_consolidated.xlsx` (rows stacked, leading `Route` column) |
| Highway Log | XLSX | `input/highway_log/` | `highway_log_consolidated.xlsx` (rows stacked, leading `Route` column) |
| TSN Highway Log | district PDFs | `input/tsn_highway_log/` | `tsn_highway_log_consolidated.xlsx` (PDFs converted to the TSMIS Highway Log format, then combined) |
| Compare: TSN vs TSMIS Highway Log | the two consolidated workbooks above | `output/` | `highway_log_comparison.xlsx` (where the vendor data does not represent TSMIS) |

**TSN Highway Log:** the TSN "California State Highway Log" district PDFs
(report OTM52010, e.g. `D01_Highway_Log_TSN.pdf`) are first converted to
per-route workbooks in `output/tsn_highway_log/` using the **exact** sheet name
and 31-column layout of the TSMIS Highway Log export, then combined. The result
is column-for-column compatible with `highway_log_consolidated.xlsx`, so the
TSN and TSMIS data can be compared directly.

**The comparison** judges the vendor at TSMIS breakpoints (so TSN's finer
segmentation never produces false differences), matches county sections by
postmile overlap, pairs rows exactly on Location + odometer, and suppresses
every representational convention verified against the data (blank TSMIS cells,
TSN placeholder codes, sig-date printing rules, epoch dates, description
encodings). TSMIS rows dated after the TSN snapshot are reported separately.

## Getting started (end users)

1. Unzip the release anywhere (right-click the zip → Properties → **Unblock**
   first, if shown).
2. Double-click **`TSMIS Consolidator.exe`**.
3. Pick the report type, point **"Folder with the exported files"** at the
   folder holding the per-route files (e.g. the Exporter's
   `output\ramp_summary`), and click **Start consolidation**.

The combined workbook lands in `output\` next to the app. Files
whose layout doesn't match the report are skipped (and listed in the log) so a
stray file can't corrupt the combined workbook.

## Developer setup

```bat
1. setup (one time).bat              :: pip install -r requirements.txt
2. consolidate (combine reports).bat :: console menu (reads input\<report>\)
run app (GUI preview).bat            :: the Tkinter GUI from source
```

Python 3.11. Runtime deps: `pdfplumber==0.11.9`, `openpyxl==3.1.5`.

## Building the app

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1            # windowed app
powershell -ExecutionPolicy Bypass -File build\build.ps1 -SelfTest  # headless release gate
```

Produces a portable PyInstaller **onefolder** under `dist\TSMIS Consolidator\`.
The build prunes third-party docs and **fails** if any DLP-flagged content
(credit-card-like numbers, private keys, …) remains in the bundle. `-SelfTest`
builds and **runs** a frozen smoke test covering pdfplumber, openpyxl, a real
consolidator run, and GUI construction.

## Project structure

See [CLAUDE.md](CLAUDE.md) for the full layout, architecture notes, and how to
add a new report type.

## License & disclaimer

MIT. Internal tool, provided as-is, no warranty. Not affiliated with Caltrans.
