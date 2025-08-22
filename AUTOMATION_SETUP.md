# 📧 Outlook Email Automation Setup Guide

## 🚀 What This Does
Automatically downloads Excel files from your Outlook emails and updates your leaderboard app!

## 📁 Files Created
- `outlook_automation.py` - Main automation script
- `run_outlook_automation.bat` - Easy-to-run batch file
- `automation_config.txt` - Configuration settings
- `automation_requirements.txt` - Required packages
- `backups/` folder - Automatic backups of old files

## ⚙️ Setup Instructions

### 1. Configure Email Settings
Edit `automation_config.txt` and fill in these details:

```
SENDER_EMAIL = your-report-sender@company.com
SUBJECT_CONTAINS = leaderboard
ATTACHMENT_NAME_CONTAINS = leaderboard
```

### 2. Test the Automation
- Double-click `run_outlook_automation.bat`
- OR run: `python outlook_automation.py`

### 3. Set Up Automatic Schedule (Optional)

#### Option A: Windows Task Scheduler
1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., daily at 9 AM)
4. Set action to run: `run_outlook_automation.bat`

#### Option B: Manual Run
Just double-click `run_outlook_automation.bat` whenever you want to check for new emails.

## 🔍 How It Works

1. **Connects to Outlook** - Uses your existing Outlook installation
2. **Searches Emails** - Looks for recent emails matching your criteria
3. **Downloads Excel** - Gets the newest Excel attachment
4. **Backs Up Old File** - Saves your previous file in `backups/` folder
5. **Updates App** - Replaces `leaderboard.xlsx` with new data
6. **Pushes to Git** - Automatically updates your live app at vpsales.streamlit.app

## 📊 What You'll See

The script will log everything:
```
2025-08-22 09:00:01 - INFO - Starting Outlook automation...
2025-08-22 09:00:02 - INFO - Successfully connected to Outlook
2025-08-22 09:00:03 - INFO - Found 1 messages matching criteria
2025-08-22 09:00:04 - INFO - Found Excel attachment: leaderboard.xlsx from reports@company.com
2025-08-22 09:00:05 - INFO - Successfully downloaded and replaced leaderboard.xlsx
2025-08-22 09:00:06 - INFO - Successfully updated git repository
2025-08-22 09:00:07 - INFO - Live app will update automatically
2025-08-22 09:00:08 - INFO - Automation completed successfully!
```

## 🛠️ Customization Options

### Search Criteria
- `SENDER_EMAIL` - Only emails from this address
- `SUBJECT_CONTAINS` - Only emails with this text in subject
- `ATTACHMENT_NAME_CONTAINS` - Only Excel files with this text in name
- `DAYS_BACK` - How many days to search (default: 7)

### Automation Options
- `AUTO_UPDATE_GIT` - Automatically push to live app (True/False)
- `CREATE_BACKUPS` - Save old files before replacing (True/False)

## 🔧 Troubleshooting

### "Failed to connect to Outlook"
- Make sure Outlook is installed and configured
- Try opening Outlook manually first

### "No messages found"
- Check your configuration settings in `automation_config.txt`
- Try removing some filters (set them to empty)

### "Git operation failed"
- Make sure you're in the correct folder
- Check if git is installed and configured

## 📞 Need Help?
Check the `outlook_automation.log` file for detailed error messages!

## 🎯 Quick Start
1. Edit `automation_config.txt` with your email details
2. Double-click `run_outlook_automation.bat`
3. Check if it worked!

That's it! Your leaderboard will now auto-update! 🚀
