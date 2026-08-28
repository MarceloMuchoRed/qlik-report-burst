@echo off
REM Launches the Qlik report burst. Called by the AHK Ctrl+G hotkey, but you can
REM also just double-click this file to run it. The window stays open at the end
REM (pause) so you can read the run summary / any errors.
cd /d "%~dp0"
python qlik_report_burst.py
echo.
pause
