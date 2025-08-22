#!/usr/bin/env python3
"""
Van Paper Business Hours Email Automation
Scans every 2 hours during business hours (7:30 AM - 3:30 PM, Mon-Fri)
Processes Van Paper reports when found, exits quietly when none found
"""

import win32com.client
import os
import shutil
import subprocess
import pandas as pd
import configparser
from datetime import datetime, timedelta
from pathlib import Path
import time

def load_config():
    """Load configuration from automation_config.txt"""
    config = configparser.ConfigParser()
    config_file = Path(__file__).parent / "automation_config.txt"
    
    if config_file.exists():
        config.read(config_file)
        return config
    else:
        print("❌ Configuration file not found!")
        return None

def find_recent_van_paper_email():
    """Find Van Paper emails from the last 3 hours"""
    
    print("🔍 Business Hours Van Paper Email Scan...")
    print(f"🕐 Current time: {datetime.now().strftime('%I:%M %p CST')}")
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("✅ Connected to Outlook")
        
        # Look for emails from the last 3 hours (to catch reports between scans)
        cutoff_time = datetime.now() - timedelta(hours=3)
        
        print(f"📅 Looking for Van Paper emails since {cutoff_time.strftime('%I:%M %p')}")
        
        # Get messages and sort by received time (newest first)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        # Look specifically for Van Paper emails
        van_paper_emails = []
        
        for message in messages:
            try:
                # Skip if no received time
                if not hasattr(message, 'ReceivedTime') or not message.ReceivedTime:
                    continue
                
                # Skip if too old
                if message.ReceivedTime < cutoff_time:
                    break  # Since sorted newest first, we can break here
                
                # Check for Van Paper sender
                sender = getattr(message, 'SenderEmailAddress', '')
                if 'noreply@vanpaper.com' not in str(sender).lower():
                    continue
                
                # Check for leaderboard export subject
                subject = getattr(message, 'Subject', '')
                if 'leaderboardexport' not in str(subject).lower():
                    continue
                
                # Check for Excel attachments
                if message.Attachments.Count == 0:
                    continue
                
                # Find Excel attachment
                excel_attachment = None
                for attachment in message.Attachments:
                    filename = attachment.FileName
                    if filename.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                        excel_attachment = attachment
                        break
                
                if excel_attachment:
                    van_paper_emails.append({
                        'message': message,
                        'attachment': excel_attachment,
                        'received_time': message.ReceivedTime
                    })
                    
            except Exception as e:
                continue
        
        if van_paper_emails:
            # Return the most recent one
            latest_email = van_paper_emails[0]  # Already sorted newest first
            print(f"🎯 FOUND Van Paper email!")
            print(f"   📅 Received: {latest_email['received_time'].strftime('%I:%M %p')}")
            print(f"   📧 Subject: {latest_email['message'].Subject}")
            print(f"   📎 Excel file: {latest_email['attachment'].FileName}")
            return latest_email
        else:
            print("ℹ️ No new Van Paper emails found")
            print("💡 This is normal between report times")
            return None
        
    except Exception as e:
        print(f"❌ Error connecting to Outlook: {e}")
        return None

def process_van_paper_email(email_data):
    """Process the Van Paper email and update the leaderboard"""
    
    print("\n📊 Processing Van Paper email...")
    
    try:
        current_dir = Path(__file__).parent
        
        # Create timestamp for files
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Save the Excel attachment
        temp_excel = current_dir / f"vanpaper_temp_{timestamp}.xlsx"
        
        print(f"💾 Saving Excel attachment...")
        email_data['attachment'].SaveAsFile(str(temp_excel))
        
        # Verify the Excel file
        try:
            df = pd.read_excel(temp_excel)
            print(f"✅ Excel verified: {len(df)} rows, {len(df.columns)} columns")
            print(f"📋 Columns: {list(df.columns)}")
        except Exception as e:
            print(f"❌ Error reading Excel file: {e}")
            return False
        
        # Create backup of current leaderboard
        main_leaderboard = current_dir / "leaderboard_new.xlsx"
        if main_leaderboard.exists():
            backup_name = f"leaderboard_backup_{timestamp}.xlsx"
            backup_path = current_dir / backup_name
            shutil.copy2(main_leaderboard, backup_path)
            print(f"💾 Created backup: {backup_name}")
        
        # Replace the main leaderboard file
        try:
            # Remove old file if it exists
            if main_leaderboard.exists():
                main_leaderboard.unlink()
            
            # Copy new file
            shutil.copy2(temp_excel, main_leaderboard)
            print(f"✅ Updated leaderboard_new.xlsx")
            
            # Clean up temp file
            temp_excel.unlink()
            
        except Exception as e:
            print(f"⚠️ File replacement issue: {e}")
            # Just rename temp file if we can't replace
            final_name = current_dir / f"leaderboard_vanpaper_{timestamp}.xlsx"
            shutil.move(temp_excel, final_name)
            print(f"💾 Saved as: {final_name.name}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error processing email: {e}")
        return False

def update_live_app(email_received_time):
    """Update git and push to live Streamlit app"""
    
    print("\n🚀 Updating live Streamlit app...")
    
    try:
        current_dir = Path(__file__).parent
        
        # Git operations
        print("📝 Adding files to git...")
        subprocess.run(['git', 'add', 'leaderboard_new.xlsx'], 
                      cwd=current_dir, capture_output=True, check=True)
        
        # Commit with timestamp
        commit_message = f"Auto-update from Van Paper report {email_received_time.strftime('%Y-%m-%d %I:%M %p CST')}"
        print(f"📝 Committing: {commit_message}")
        subprocess.run(['git', 'commit', '-m', commit_message], 
                      cwd=current_dir, capture_output=True, check=True)
        
        # Push to live app
        print("🌐 Pushing to live app...")
        result = subprocess.run(['git', 'push'], 
                              cwd=current_dir, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ Successfully updated live app!")
            print("🌐 Live app: https://vpsales.streamlit.app/")
            print("⏱️ App will refresh in 1-2 minutes")
            return True
        else:
            print(f"⚠️ Git push failed: {result.stderr}")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"⚠️ Git operation failed: {e}")
        return False
    except Exception as e:
        print(f"❌ Update error: {e}")
        return False

def main():
    """Main business hours automation function"""
    
    print("🕐 Van Paper Business Hours Email Automation")
    print("=" * 50)
    print(f"🕐 Scan time: {datetime.now().strftime('%Y-%m-%d %I:%M:%S %p CST')}")
    
    # Check if it's a business day
    current_day = datetime.now().weekday()  # 0=Monday, 6=Sunday
    if current_day >= 5:  # Saturday or Sunday
        print("📅 Weekend detected - skipping scan")
        print("💤 Business hours automation runs Monday-Friday only")
        return True
    
    # Check if it's business hours (7:30 AM to 3:30 PM)
    current_hour = datetime.now().hour
    current_minute = datetime.now().minute
    current_time_minutes = current_hour * 60 + current_minute
    
    # Business hours: 7:30 AM (450 min) to 3:30 PM (930 min)
    if current_time_minutes < 450 or current_time_minutes > 930:
        print("🕐 Outside business hours - skipping scan")
        print("💼 Business hours: 7:30 AM - 3:30 PM, Monday-Friday")
        return True
    
    # Load configuration
    config = load_config()
    if not config:
        print("❌ Failed to load configuration")
        return False
    
    print("✅ Configuration loaded")
    
    # Find Van Paper email
    email_data = find_recent_van_paper_email()
    if not email_data:
        print("\n✅ No new Van Paper reports - scan complete")
        print("💡 Will check again in 2 hours")
        return True  # This is normal, not an error
    
    # Process the email
    if not process_van_paper_email(email_data):
        print("❌ Failed to process Van Paper email")
        return False
    
    # Update the live app
    if not update_live_app(email_data['received_time']):
        print("⚠️ Live app update had issues")
        return False
    
    print("\n🎉 SUCCESS! Van Paper report processed!")
    print(f"📧 Processed email from: {email_data['received_time'].strftime('%I:%M %p')}")
    print("🌐 Live app updated with fresh data!")
    
    return True

if __name__ == "__main__":
    success = main()
    
    print("\n" + "=" * 50)
    if success:
        print("✅ Business hours scan completed!")
    else:
        print("❌ Business hours scan had issues")
    
    print(f"🕐 Next scan: In ~2 hours (business hours only)")
    print(f"🔗 Live app: https://vpsales.streamlit.app/")
