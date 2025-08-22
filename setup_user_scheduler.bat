@echo off
echo Setting up Van Paper Business Hours Automation (User Level)
echo ==========================================================

cd /d "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"

echo Creating user-level scheduled task for 1:30 PM...
schtasks /create /tn "VanPaper1330" /tr "python \"C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\business_hours_automation.py\"" /sc daily /st 13:30 /f

echo.
echo Creating user-level scheduled task for 3:30 PM...
schtasks /create /tn "VanPaper1530" /tr "python \"C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\business_hours_automation.py\"" /sc daily /st 15:30 /f

echo.
echo Verifying tasks...
schtasks /query /tn "VanPaper1330" /fo LIST
echo.
schtasks /query /tn "VanPaper1530" /fo LIST

echo.
echo ==========================================================
echo User-level automation tasks created!
echo - 1:30 PM daily automation
echo - 3:30 PM daily automation
echo ==========================================================

pause
