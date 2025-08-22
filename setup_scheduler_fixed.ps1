# Simple Van Paper Task Setup
# Fixed version with proper encoding and error handling

Write-Host "Van Paper Email Automation - Task Scheduler Setup" -ForegroundColor Green
Write-Host "======================================================" -ForegroundColor Green

$currentDir = "c:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    pause
    exit 1
}

Write-Host "SUCCESS: Running as Administrator" -ForegroundColor Green

# Create Morning Task (7:32 AM CST)
Write-Host "`nCreating Morning Task (7:32 AM CST)..." -ForegroundColor Yellow

$morningAction = New-ScheduledTaskAction -Execute "$currentDir\run_morning_automation.bat"
$morningTrigger = New-ScheduledTaskTrigger -Daily -At "07:32"
$morningSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

try {
    # Use current user context instead of requiring password
    Register-ScheduledTask -TaskName "Van Paper Morning Automation" -Action $morningAction -Trigger $morningTrigger -Settings $morningSettings -Force
    Write-Host "SUCCESS: Morning task created!" -ForegroundColor Green
} catch {
    Write-Host "ERROR creating morning task: $($_.Exception.Message)" -ForegroundColor Red
}

# Create Afternoon Task (2:02 PM CST)
Write-Host "`nCreating Afternoon Task (2:02 PM CST)..." -ForegroundColor Yellow

$afternoonAction = New-ScheduledTaskAction -Execute "$currentDir\run_afternoon_automation.bat"
$afternoonTrigger = New-ScheduledTaskTrigger -Daily -At "14:02"
$afternoonSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

try {
    # Use current user context instead of requiring password
    Register-ScheduledTask -TaskName "Van Paper Afternoon Automation" -Action $afternoonAction -Trigger $afternoonTrigger -Settings $afternoonSettings -Force
    Write-Host "SUCCESS: Afternoon task created!" -ForegroundColor Green
} catch {
    Write-Host "ERROR creating afternoon task: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nAUTOMATION SETUP COMPLETE!" -ForegroundColor Green
Write-Host "======================================================" -ForegroundColor Green

Write-Host "`nSCHEDULED TASKS:" -ForegroundColor Cyan
Write-Host "Morning:   7:32 AM CST (processes 7:30 AM Van Paper report)" -ForegroundColor White
Write-Host "Afternoon: 2:02 PM CST (processes 2:00 PM Van Paper report)" -ForegroundColor White

Write-Host "`nTO MANAGE TASKS:" -ForegroundColor Cyan
Write-Host "- Open Task Scheduler (taskschd.msc)" -ForegroundColor White
Write-Host "- Look for 'Van Paper Morning Automation' and 'Van Paper Afternoon Automation'" -ForegroundColor White
Write-Host "- Right-click to Run, Disable, or Modify tasks" -ForegroundColor White

Write-Host "`nTO TEST AUTOMATION NOW:" -ForegroundColor Cyan
Write-Host "- Run: python scheduled_automation.py" -ForegroundColor White
Write-Host "- Or double-click: run_morning_automation.bat" -ForegroundColor White

Write-Host "`nLIVE APP:" -ForegroundColor Cyan
Write-Host "- https://vpsales.streamlit.app/" -ForegroundColor White
Write-Host "- Updates automatically 1-2 minutes after processing" -ForegroundColor White

Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
