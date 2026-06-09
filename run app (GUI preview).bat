@echo off
cd /d "%~dp0"
rem Dev launcher for the GUI (the packaged app runs gui_main frozen).
python scripts\gui_main.py
if errorlevel 1 pause
