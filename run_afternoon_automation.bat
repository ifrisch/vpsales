@echo off
REM Van Paper Afternoon Automation - 2:02 PM CST
echo Van Paper Afternoon Report Automation
echo =====================================
echo Time: %date% %time%
echo.

cd /d "c:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"

echo Running afternoon automation...
python scheduled_automation.py

echo.
echo Afternoon automation complete!
pause
