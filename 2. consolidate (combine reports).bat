@echo off
cd /d "%~dp0"

:menu
cls
echo ================================================================
echo            TSMIS Reports - Bulk Consolidator
echo ================================================================
echo.
echo  Combine per-route exports into one Excel file in output\.
echo  Put the exported files in the matching input\^<report^>\ folder
echo  first (or run the GUI to pick any folder).
echo.
echo  Which report do you want to consolidate?
echo.
echo     1.  TSAR: Ramp Summary       (PDFs -^> XLSX)
echo     2.  TSAR: Ramp Detail        (XLSX -^> XLSX)
echo     3.  Highway Sequence Listing (XLSX -^> XLSX)
echo     4.  Highway Log              (XLSX -^> XLSX)
echo     5.  TSN Highway Log          (district PDFs -^> XLSX)
echo     6.  Compare TSN vs TSMIS Highway Log (needs 4 + 5 done)
echo.
echo     Q.  Quit
echo.
echo ================================================================
echo.
set "choice="
set /p choice="Enter your choice [1, 2, 3, 4, 5, 6, Q]: "

if /i "%choice%"=="1" goto ramp_summary
if /i "%choice%"=="2" goto ramp_detail
if /i "%choice%"=="3" goto highway_sequence
if /i "%choice%"=="4" goto highway_log
if /i "%choice%"=="5" goto tsn_highway_log
if /i "%choice%"=="6" goto compare_hl
if /i "%choice%"=="Q" exit /b 0
if /i "%choice%"=="quit" exit /b 0
echo.
echo Invalid choice "%choice%". Please pick 1, 2, 3, 4, 5, 6, or Q.
echo.
pause
goto menu

:ramp_summary
python scripts\consolidate_ramp_summary.py
pause
exit /b 0

:ramp_detail
python scripts\consolidate_ramp_detail.py
pause
exit /b 0

:highway_sequence
python scripts\consolidate_highway_sequence.py
pause
exit /b 0

:highway_log
python scripts\consolidate_highway_log.py
pause
exit /b 0

:tsn_highway_log
python scripts\consolidate_tsn_highway_log.py
pause
exit /b 0

:compare_hl
python scripts\compare_highway_log.py
pause
exit /b 0
