# Van Paper Automation - Task Scheduler Setup
# Run this script as Administrator to set up automated tasks

Write-Host "🤖 Van Paper Email Automation - Task Scheduler Setup" -ForegroundColor Green
Write-Host "=" * 60 -ForegroundColor Green

$currentDir = "c:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "❌ This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    pause
    exit 1
}

Write-Host "✅ Running as Administrator" -ForegroundColor Green

# Create Morning Task (7:32 AM CST)
Write-Host "`n📅 Creating Morning Task (7:32 AM CST)..." -ForegroundColor Yellow

$morningAction = New-ScheduledTaskAction -Execute "$currentDir\run_morning_automation.bat"
$morningTrigger = New-ScheduledTaskTrigger -Daily -At "07:32"
$morningSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
$morningPrincipal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType InteractiveOrPassword

try {
    Register-ScheduledTask -TaskName "Van Paper Morning Automation" -Action $morningAction -Trigger $morningTrigger -Settings $morningSettings -Principal $morningPrincipal -Force
    Write-Host "✅ Morning task created successfully!" -ForegroundColor Green
} catch {
    Write-Host "❌ Error creating morning task: $($_.Exception.Message)" -ForegroundColor Red
}

# Create Afternoon Task (2:02 PM CST)
Write-Host "`n📅 Creating Afternoon Task (2:02 PM CST)..." -ForegroundColor Yellow

$afternoonAction = New-ScheduledTaskAction -Execute "$currentDir\run_afternoon_automation.bat"
$afternoonTrigger = New-ScheduledTaskTrigger -Daily -At "14:02"
$afternoonSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
$afternoonPrincipal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType InteractiveOrPassword

try {
    Register-ScheduledTask -TaskName "Van Paper Afternoon Automation" -Action $afternoonAction -Trigger $afternoonTrigger -Settings $afternoonSettings -Principal $afternoonPrincipal -Force
    Write-Host "✅ Afternoon task created successfully!" -ForegroundColor Green
} catch {
    Write-Host "❌ Error creating afternoon task: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n🎯 AUTOMATION SETUP COMPLETE!" -ForegroundColor Green
Write-Host "=" * 60 -ForegroundColor Green

Write-Host "`n📋 SCHEDULED TASKS:" -ForegroundColor Cyan
Write-Host "🌅 Morning:   7:32 AM CST (processes 7:30 AM Van Paper report)" -ForegroundColor White
Write-Host "🌆 Afternoon: 2:02 PM CST (processes 2:00 PM Van Paper report)" -ForegroundColor White

Write-Host "`n🔧 TO MANAGE TASKS:" -ForegroundColor Cyan
Write-Host "• Open Task Scheduler (taskschd.msc)" -ForegroundColor White
Write-Host "• Look for 'Van Paper Morning Automation' and 'Van Paper Afternoon Automation'" -ForegroundColor White
Write-Host "• Right-click to Run, Disable, or Modify tasks" -ForegroundColor White

Write-Host "`n🧪 TO TEST AUTOMATION NOW:" -ForegroundColor Cyan
Write-Host "• Run: python scheduled_automation.py" -ForegroundColor White
Write-Host "• Or double-click: run_morning_automation.bat" -ForegroundColor White

Write-Host "`n🌐 LIVE APP:" -ForegroundColor Cyan
Write-Host "• https://vpsales.streamlit.app/" -ForegroundColor White
Write-Host "• Updates automatically 1-2 minutes after processing" -ForegroundColor White

Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
