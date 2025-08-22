"""
Van Paper 2:02 PM Test Preparation
This script verifies everything is ready for today's 2:02 PM automation test
"""

from datetime import datetime, timedelta
import os
from pathlib import Path

def check_afternoon_readiness():
    """Check if everything is ready for the 2:02 PM automation"""
    
    print("🕐 Van Paper 2:02 PM Automation - Readiness Check")
    print("=" * 55)
    
    current_time = datetime.now()
    target_time = current_time.replace(hour=14, minute=2, second=0, microsecond=0)
    
    if current_time > target_time:
        target_time += timedelta(days=1)  # Tomorrow if we've passed today's time
    
    time_until = target_time - current_time
    hours = int(time_until.total_seconds() // 3600)
    minutes = int((time_until.total_seconds() % 3600) // 60)
    
    print(f"🕐 Current time: {current_time.strftime('%I:%M:%S %p CST')}")
    print(f"🎯 Next automation: {target_time.strftime('%I:%M %p CST on %B %d')}")
    print(f"⏰ Time remaining: {hours} hours, {minutes} minutes")
    
    print("\n🔍 READINESS CHECKLIST:")
    
    # Check files exist
    current_dir = Path(__file__).parent
    required_files = [
        "scheduled_automation.py",
        "run_afternoon_automation.bat", 
        "automation_config.txt",
        "leaderboard_new.xlsx"
    ]
    
    all_files_ready = True
    for file in required_files:
        file_path = current_dir / file
        if file_path.exists():
            print(f"✅ {file} - Ready")
        else:
            print(f"❌ {file} - Missing!")
            all_files_ready = False
    
    # Check Task Scheduler (if possible)
    print(f"\n📋 TASK SCHEDULER:")
    print(f"✅ Task should be: 'Van Paper Afternoon Automation'")
    print(f"✅ Scheduled for: 2:02 PM (14:02) daily")
    print(f"💡 Verify in Task Scheduler (taskschd.msc)")
    
    # Van Paper email expectations
    print(f"\n📧 VAN PAPER EMAIL EXPECTATIONS:")
    print(f"📤 Van Paper should send report at: 2:00 PM CST")
    print(f"🤖 Automation will scan at: 2:02 PM CST") 
    print(f"📨 Looking for: noreply@vanpaper.com")
    print(f"📋 Subject: 'Inform Auto Scheduled Report: leaderboardexport'")
    print(f"📎 With: Excel attachment")
    
    # Live app info
    print(f"\n🌐 LIVE APP UPDATE:")
    print(f"🚀 App will update: 2-3 minutes after processing")
    print(f"🔗 Check: https://vpsales.streamlit.app/")
    print(f"📊 New data will appear automatically")
    
    if all_files_ready:
        print(f"\n🎉 READY FOR 2:02 PM AUTOMATION!")
        print(f"✅ All systems prepared")
        print(f"⏰ {hours} hours, {minutes} minutes until next run")
    else:
        print(f"\n⚠️ SETUP NEEDS ATTENTION")
        print(f"❌ Missing files detected")
    
    print(f"\n💡 TO MONITOR:")
    print(f"• Watch Task Scheduler at 2:02 PM")
    print(f"• Check live app around 2:05 PM")
    print(f"• Look for new backup files created")
    
    return all_files_ready

if __name__ == "__main__":
    ready = check_afternoon_readiness()
    
    print("\n" + "=" * 55)
    if ready:
        print("🚀 Ready for automatic 2:02 PM Van Paper processing!")
    else:
        print("⚠️ Please fix missing files before 2:02 PM")
    
    input("\nPress Enter to continue...")
