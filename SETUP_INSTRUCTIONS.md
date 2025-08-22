# Van Paper Email Automation - Setup Instructions

## 🎯 AUTOMATION OVERVIEW
Your automation is now ready! It will:
- Run automatically at **7:32 AM CST** (2 minutes after your 7:30 AM report)
- Run automatically at **2:02 PM CST** (2 minutes after your 2:00 PM report)
- Find Van Paper emails with Excel attachments
- Update your live Streamlit app: https://vpsales.streamlit.app/

## 🚀 SETUP STEPS

### Step 1: Set up Windows Task Scheduler
**Option A - Automatic Setup (Recommended):**
1. Right-click PowerShell and select "Run as Administrator"
2. Navigate to your project folder:
   ```powershell
   cd "c:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"
   ```
3. Run the setup script:
   ```powershell
   .\setup_scheduler.ps1
   ```

**Option B - Manual Setup:**
1. Open Task Scheduler (Windows key + R, type `taskschd.msc`)
2. Click "Create Basic Task"
3. Create two tasks:
   - **Morning Task**: Run at 7:32 AM daily, execute `run_morning_automation.bat`
   - **Afternoon Task**: Run at 2:02 PM daily, execute `run_afternoon_automation.bat`

### Step 2: Verify Setup
1. Check Task Scheduler shows your two Van Paper tasks
2. Test by running: `python test_automation.py`
3. Monitor your live app: https://vpsales.streamlit.app/

## 📋 FILES CREATED
- `scheduled_automation.py` - Main automation script
- `run_morning_automation.bat` - Morning batch file (7:32 AM)
- `run_afternoon_automation.bat` - Afternoon batch file (2:02 PM)
- `setup_scheduler.ps1` - PowerShell setup script
- `test_automation.py` - Test script
- `automation_config.txt` - Configuration (already updated with Van Paper settings)

## 🔧 TROUBLESHOOTING
- **No emails found**: Check if Van Paper reports arrived on time
- **Permission errors**: Run setup as Administrator
- **App not updating**: Check git credentials and internet connection
- **Wrong time zone**: Automation uses local computer time (CST)

## 🎉 SUCCESS INDICATORS
When working correctly, you'll see:
- Scheduled tasks in Task Scheduler
- Live app updates within 2-3 minutes of Van Paper email arrival
- Backup files created automatically
- Git commits with timestamps

Your automation is ready to go! 🚀
