@echo off
REM Van Paper Automation Wrapper - Silent Version
REM Runs completely silently without any popup windows

REM Log start time
echo [%time%] Van Paper Automation Starting... >> "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\automation.log"

REM Set working directory
cd /d "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"

REM Set full Python path
set PYTHON_PATH=C:\Users\Isaac\AppData\Local\Programs\Python\Python313\python.exe

REM Run the automation silently - no window will appear
"%PYTHON_PATH%" "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\one_click_update.py" > nul 2>&1

REM Log completion
echo [%time%] Van Paper Automation Complete >> "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\automation.log"

REM Exit immediately without pause
exit /b
