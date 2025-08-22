@echo off
REM Van Paper Automation Wrapper with Full Paths
REM This handles all the environment setup that scheduled tasks miss

echo [%time%] Van Paper Automation Starting... >> "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\automation.log"

REM Set working directory
cd /d "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"

REM Set full Python path (adjust if needed)
set PYTHON_PATH=C:\Users\Isaac\AppData\Local\Programs\Python\Python313\python.exe

REM Run the automation with full path
"%PYTHON_PATH%" "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\business_hours_automation.py" >> "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\automation.log" 2>&1

echo [%time%] Van Paper Automation Complete >> "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\automation.log"
