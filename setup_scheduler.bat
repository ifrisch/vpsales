@echo off
echo Setting up Van Paper Business Hours Automation Schedule
echo =====================================================

REM Create the scheduled task for business hours automation
REM Runs every 2 hours from 7:30 AM to 3:30 PM, Monday-Friday

cd /d "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"

echo Creating scheduled task: VanPaperBusinessHours
schtasks /create /tn "VanPaperBusinessHours" /tr "python \"C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\business_hours_automation.py\"" /sc daily /st 07:30 /mo 1 /ru SYSTEM /f

echo.
echo Creating additional triggers for 9:30 AM...
schtasks /create /tn "VanPaperBusinessHours930" /tr "python \"C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\business_hours_automation.py\"" /sc daily /st 09:30 /mo 1 /ru SYSTEM /f

echo.
echo Creating additional triggers for 11:30 AM...
schtasks /create /tn "VanPaperBusinessHours1130" /tr "python \"C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\business_hours_automation.py\"" /sc daily /st 11:30 /mo 1 /ru SYSTEM /f

echo.
echo Creating additional triggers for 1:30 PM...
schtasks /create /tn "VanPaperBusinessHours1330" /tr "python \"C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\business_hours_automation.py\"" /sc daily /st 13:30 /mo 1 /ru SYSTEM /f

echo.
echo Creating additional triggers for 3:30 PM...
schtasks /create /tn "VanPaperBusinessHours1530" /tr "python \"C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\business_hours_automation.py\"" /sc daily /st 15:30 /mo 1 /ru SYSTEM /f

echo.
echo =====================================================
echo Business Hours Automation Schedule Created!
echo.
echo The following tasks will run Monday-Friday:
echo - 7:30 AM  - Morning scan
echo - 9:30 AM  - Mid-morning scan  
echo - 11:30 AM - Late morning scan
echo - 1:30 PM  - Afternoon scan
echo - 3:30 PM  - End of day scan
echo.
echo To verify: schtasks /query /tn "VanPaperBusinessHours1130"
echo =====================================================

pause
