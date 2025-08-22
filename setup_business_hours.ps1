# Van Paper Business Hours Automation Setup
# Scans every 2 hours from 7:30 AM - 3:00 PM, Monday-Friday

Write-Host "Van Paper Business Hours Email Automation Setup" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green

$currentDir = "c:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    pause
    exit 1
}

Write-Host "SUCCESS: Running as Administrator" -ForegroundColor Green

# Remove old tasks if they exist
Write-Host "`nRemoving old Van Paper tasks..." -ForegroundColor Yellow
try {
    Unregister-ScheduledTask -TaskName "Van Paper Morning Automation" -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName "Van Paper Afternoon Automation" -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Old tasks removed" -ForegroundColor Green
} catch {
    Write-Host "No old tasks to remove" -ForegroundColor Gray
}

# Business hours schedule: 7:30 AM, 9:30 AM, 11:30 AM, 1:30 PM, 3:30 PM
$businessHours = @("07:30", "09:30", "11:30", "13:30", "15:30")
$taskNames = @(
    "Van Paper 7:30 AM Scan",
    "Van Paper 9:30 AM Scan", 
    "Van Paper 11:30 AM Scan",
    "Van Paper 1:30 PM Scan",
    "Van Paper 3:30 PM Scan"
)

Write-Host "`nCreating business hours automation tasks..." -ForegroundColor Yellow

for ($i = 0; $i -lt $businessHours.Length; $i++) {
    $time = $businessHours[$i]
    $taskName = $taskNames[$i]
    
    Write-Host "Creating task: $taskName at $time" -ForegroundColor Cyan
    
    try {
        # Create action to run the automation script
        $action = New-ScheduledTaskAction -Execute "$currentDir\run_business_hours_scan.bat"
        
        # Create trigger for weekdays only (Monday=2, Friday=6)
        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At $time
        
        # Create settings
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable
        
        # Register the task
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
        
        Write-Host "  SUCCESS: $taskName created" -ForegroundColor Green
        
    } catch {
        Write-Host "  ERROR: Failed to create $taskName - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`nBUSINESS HOURS AUTOMATION COMPLETE!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green

Write-Host "`nSCHEDULE OVERVIEW:" -ForegroundColor Cyan
Write-Host "Monday - Friday:" -ForegroundColor White
Write-Host "  7:30 AM - Scan for Van Paper reports" -ForegroundColor White
Write-Host "  9:30 AM - Scan for Van Paper reports" -ForegroundColor White
Write-Host " 11:30 AM - Scan for Van Paper reports" -ForegroundColor White
Write-Host "  1:30 PM - Scan for Van Paper reports" -ForegroundColor White
Write-Host "  3:30 PM - Scan for Van Paper reports" -ForegroundColor White

Write-Host "`nWEEKEND BEHAVIOR:" -ForegroundColor Cyan
Write-Host "- No scans on Saturday/Sunday" -ForegroundColor White
Write-Host "- Automation resumes Monday morning" -ForegroundColor White

Write-Host "`nHOW IT WORKS:" -ForegroundColor Cyan
Write-Host "- Scans for Van Paper emails from last 3 hours" -ForegroundColor White
Write-Host "- If report found: processes and updates live app" -ForegroundColor White
Write-Host "- If no report: logs 'no new reports' and exits quietly" -ForegroundColor White
Write-Host "- Updates: https://vpsales.streamlit.app/" -ForegroundColor White

Write-Host "`nTO MANAGE:" -ForegroundColor Cyan
Write-Host "- Open Task Scheduler (taskschd.msc)" -ForegroundColor White
Write-Host "- Look for 'Van Paper [Time] Scan' tasks" -ForegroundColor White
Write-Host "- Right-click to Run, Disable, or Modify" -ForegroundColor White

Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
