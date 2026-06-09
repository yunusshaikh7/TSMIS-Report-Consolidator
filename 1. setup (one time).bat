@echo off
cd /d "%~dp0"
echo ================================================================
echo            TSMIS Report Consolidator - one-time setup
echo ================================================================
echo.
echo Installing the PDF / Excel libraries this tool needs...
echo (No browser and no login -- it only combines files already
echo  exported by the TSMIS Reports Exporter.)
echo.
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: pip install failed. Is Python installed and on PATH?
    pause
    exit /b 1
)
echo.
echo Setup complete. Next, run "2. consolidate (combine reports).bat".
pause
