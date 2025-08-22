@echo off
echo ===============================================
echo Van Paper Manual Update Runner
echo ===============================================
echo.
echo This will check for new Van Paper emails and update the live app
echo.

cd /d "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"

echo Running Van Paper automation...
python business_hours_automation.py

echo.
echo ===============================================
echo Update complete!
echo Check your live app: https://vpsales.streamlit.app/
echo ===============================================

pause
