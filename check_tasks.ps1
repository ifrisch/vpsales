# Check if Van Paper tasks were created successfully

Write-Host "Checking Van Paper Scheduled Tasks..." -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green

# Check for Morning Task
try {
    $morningTask = Get-ScheduledTask -TaskName "Van Paper Morning Automation" -ErrorAction Stop
    Write-Host "SUCCESS: Morning task found!" -ForegroundColor Green
    Write-Host "  Name: $($morningTask.TaskName)"
    Write-Host "  State: $($morningTask.State)"
} catch {
    Write-Host "WARNING: Morning task not found" -ForegroundColor Yellow
}

# Check for Afternoon Task  
try {
    $afternoonTask = Get-ScheduledTask -TaskName "Van Paper Afternoon Automation" -ErrorAction Stop
    Write-Host "SUCCESS: Afternoon task found!" -ForegroundColor Green
    Write-Host "  Name: $($afternoonTask.TaskName)"
    Write-Host "  State: $($afternoonTask.State)"
} catch {
    Write-Host "WARNING: Afternoon task not found" -ForegroundColor Yellow
}

Write-Host "`nOpening Task Scheduler for manual verification..." -ForegroundColor Cyan
Start-Process "taskschd.msc"

Write-Host "`nDone! Check Task Scheduler window for your Van Paper tasks." -ForegroundColor Green
