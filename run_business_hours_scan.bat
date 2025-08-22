@echo off
REM Van Paper Business Hours Scan
echo Van Paper Business Hours Email Scan
echo ===================================
echo Time: %date% %time%
echo.

cd /d "c:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"

echo Running business hours email scan...
python business_hours_automation.py

echo.
echo Scan complete!
