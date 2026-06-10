Portable Windows desktop app that combines per-route Caltrans TSMIS report
files into one Excel workbook per report type — no login, no browser, no
network. Unzip, double-click `TSMIS Consolidator.exe`, done (see
`Start Here.txt` inside the zip).

## Report types

| Report | Input | Combined output |
|---|---|---|
| TSAR: Ramp Summary | per-route PDFs | `tsar_ramp_summary_consolidated.xlsx` (audited, with a live Combined sheet) |
| TSAR: Ramp Detail | per-route XLSX | `tsar_ramp_detail_consolidated.xlsx` |
| Highway Sequence Listing | per-route XLSX | `highway_sequence_consolidated.xlsx` |
| Highway Log | per-route XLSX | `highway_log_consolidated.xlsx` |
| TSN Highway Log | district PDFs (OTM52010) | `tsn_highway_log_consolidated.xlsx` |
| **Compare: TSN vs TSMIS Highway Log** (new) | the two consolidated workbooks | `highway_log_comparison.xlsx` |

## Highlights

- **TSN vs TSMIS comparison:** reports where the TSN data does not represent
  the TSMIS Highway Log — engineered against false positives (rows pair on
  Location + odometer, the vendor is judged at TSMIS breakpoints so TSN's
  finer segmentation never counts as a difference, blank TSMIS cells make no
  claim, and every verified report convention is suppressed and listed on the
  report's Summary sheet). Changes TSMIS made after the TSN snapshot are
  reported separately.
- **TSN Highway Log conversion:** district PDFs are parsed and rewritten as
  per-route workbooks in the exact TSMIS Highway Log layout (same sheet name,
  same 31 columns), then combined — so the result lines up column-for-column
  with the consolidated TSMIS Highway Log for comparison.
- **Separated folders:** source files go in `input\<report>\` (or browse to
  any folder, e.g. the TSMIS Reports Exporter's output); everything the app
  writes lands in `output\`.
- Mismatched-layout files are skipped and reported, never silently merged.
- First run: if Windows warns about an unknown publisher, choose
  "More info" → "Run anyway" (in-house unsigned tool). If downloaded as a
  zip, right-click → Properties → Unblock before extracting.
