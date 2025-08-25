@echo off
echo Creating WORKING Van Paper Automation Tasks
echo ==========================================

echo Creating 9:30 AM task...
schtasks /create /tn "VanPaper_0930_Working" /tr "cmd /c \"cd /d C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard && C:\Users\Isaac\AppData\Local\Programs\Python\Python313\python.exe silent_automation.py\"" /sc daily /st 09:30 /f

echo Creating 11:30 AM task...
schtasks /create /tn "VanPaper_1130_Working" /tr "cmd /c \"cd /d C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard && C:\Users\Isaac\AppData\Local\Programs\Python\Python313\python.exe silent_automation.py\"" /sc daily /st 11:30 /f

echo Creating 1:30 PM task...
schtasks /create /tn "VanPaper_1330_Working" /tr "cmd /c \"cd /d C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard && C:\Users\Isaac\AppData\Local\Programs\Python\Python313\python.exe silent_automation.py\"" /sc daily /st 13:30 /f

echo Creating 3:30 PM task...
schtasks /create /tn "VanPaper_1530_Working" /tr "cmd /c \"cd /d C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard && C:\Users\Isaac\AppData\Local\Programs\Python\Python313\python.exe silent_automation.py\"" /sc daily /st 15:30 /f

echo.
echo ==========================================
echo WORKING automation tasks created!
echo.
echo Testing one task manually...
cmd /c "cd /d C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard && C:\Users\Isaac\AppData\Local\Programs\Python\Python313\python.exe silent_automation.py"

echo.
echo Verifying tasks...
schtasks /query | findstr Working

echo ==========================================

pause
