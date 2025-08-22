@echo off
title Van Paper Leaderboard Automation
color 0A

echo ========================================
echo  Van Paper Sales Leaderboard Automation
echo ========================================
echo.
echo This will:
echo  - Connect to your Outlook
echo  - Look for emails from noreply@vanpaper.com
echo  - Download the latest leaderboard report
echo  - Update your live app automatically
echo.
echo ========================================
echo.

cd /d "c:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"

echo Checking dependencies...
python -c "import win32com.client; print('✅ Outlook integration ready')" 2>nul || (
    echo Installing required packages...
    pip install pywin32
)

echo.
echo 🔍 Starting automation...
echo.
python outlook_automation.py

echo.
echo ========================================
if %ERRORLEVEL% == 0 (
    echo ✅ SUCCESS: Automation completed!
    echo 🚀 Your live app will update automatically
) else (
    echo ❌ ERROR: Something went wrong
    echo 📋 Check outlook_automation.log for details
)
echo ========================================
echo.
echo Press any key to exit...
pause >nul
