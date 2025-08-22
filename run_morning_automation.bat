@echo off
REM Van Paper Morning Automation - 7:32 AM CST
echo Van Paper Morning Report Automation
echo ===================================
echo Time: %date% %time%
echo.

cd /d "c:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"

echo Running morning automation...
python scheduled_automation.py

echo.
echo Morning automation complete!
pause
