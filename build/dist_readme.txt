TSMIS Report Consolidator
=========================

WHAT IT DOES
  Combines the per-route TSMIS report files (exported by the TSMIS Reports
  Exporter) into ONE Excel workbook per report type. Reports supported:
    - TSAR: Ramp Summary        (reads PDFs)
    - TSAR: Ramp Detail         (reads Excel)
    - Highway Sequence Listing  (reads Excel)
    - Highway Log               (reads Excel)
    - TSN Highway Log           (reads the TSN district PDFs and converts
                                 them to the TSMIS Excel layout first)
    - Compare TSN vs TSMIS      (reports where the TSN data does not match
                                 the TSMIS Highway Log; run 4 and 5 first)
  No login, no browser, no internet -- it only reads files already on this PC.

HOW TO RUN
  Double-click  "TSMIS Consolidator.exe"  in this folder.
  Keep this whole folder together -- the app needs the "_internal" folder next
  to the .exe. Don't move the .exe out on its own. You don't need Python.

COMBINE FILES
  1. Pick the report type.
  2. Point "Folder with the exported files" at wherever the per-route files
     are -- e.g. the Exporter's  output\ramp_summary  folder -- or copy them
     into this app's matching  input\<report>  folder and leave the default.
  3. Click  "Start consolidation".
  The combined workbook is saved in  output  next to this app
  ("Open folder" takes you there). If one already exists you'll be asked
  before it is overwritten.

GOOD TO KNOW
  * Files whose layout doesn't match the report (wrong sheet, different
    columns) are skipped and listed in the log -- they can't silently corrupt
    the combined workbook.
  * The first time you run it, Windows may say the publisher is unknown. That's
    expected for an in-house, unsigned tool: choose "More info" -> "Run anyway".
  * If you received this as a .zip, right-click the zip -> Properties -> tick
    "Unblock" -> OK, BEFORE extracting it. This also helps with IT/Defender.
  * Logs are under  "data\logs"  (or click "Logs" in the app). Include them if
    you report a problem.
