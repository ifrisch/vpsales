@echo off
echo Creating BULLETPROOF Van Paper Automation Schedule
echo ==================================================

REM Create multiple scheduled tasks with proper paths and logging

echo Creating 9:30 AM automation...
schtasks /create /tn "VanPaper_0930" /tr "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\automation_wrapper.bat" /sc daily /st 09:30 /f

echo Creating 11:30 AM automation...
schtasks /create /tn "VanPaper_1130" /tr "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\automation_wrapper.bat" /sc daily /st 11:30 /f

echo Creating 1:30 PM automation...
schtasks /create /tn "VanPaper_1330" /tr "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\automation_wrapper.bat" /sc daily /st 13:30 /f

echo Creating 3:30 PM automation...
schtasks /create /tn "VanPaper_1530" /tr "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\automation_wrapper.bat" /sc daily /st 15:30 /f

echo.
echo Testing the automation wrapper...
call "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\automation_wrapper.bat"

echo.
echo ==================================================
echo AUTOMATION SCHEDULE CREATED!
echo.
echo Tasks will run at:
echo - 9:30 AM daily
echo - 11:30 AM daily  
echo - 1:30 PM daily
echo - 3:30 PM daily
echo.
echo Log file: automation.log
echo.
echo Verifying tasks...
schtasks /query | findstr VanPaper

echo ==================================================

pause
