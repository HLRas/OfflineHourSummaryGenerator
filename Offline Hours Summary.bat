@echo off
rem ---------------------------------------------------------------
rem  Offline Hours Summary
rem  Double-click to open the window, or drag a calendar .CSV
rem  straight onto this file to generate the report immediately.
rem ---------------------------------------------------------------
cd /d "%~dp0"

where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" pythonw "offline_hours.py" %*
    goto :eof
)

where python >nul 2>&1
if %errorlevel%==0 (
    start "" python "offline_hours.py" %*
    goto :eof
)

echo.
echo Python was not found on this PC.
echo Install Python 3 from https://www.python.org/downloads/
echo and tick "Add python.exe to PATH" during setup.
echo.
pause
