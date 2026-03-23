@echo off
echo.
echo  Refreshing Market Feedback Dashboard...
echo.
cd /d "%~dp0"
python scripts\extract_data.py
echo.
pause
