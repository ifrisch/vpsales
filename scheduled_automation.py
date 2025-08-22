#!/usr/bin/env python3
"""
Van Paper Scheduled Email Automation
Runs at 7:32 AM CST and 2:02 PM CST (14:02) to process Van Paper reports
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

def find_van_paper_email():
    """Find the most recent Van Paper email with Excel attachment"""
    
    print("🔍 Looking for Van Paper email...")
    print(f"🕐 Current time: {datetime.now().strftime('%I:%M %p CST')}")
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("✅ Connected to Outlook")
        
        # Look for emails from the last 2 hours (to catch the 7:30 AM or 2:00 PM reports)
        cutoff_time = datetime.now() - timedelta(hours=2)
        
        print(f"📅 Looking for emails since {cutoff_time.strftime('%I:%M %p')}")
        
        # Get messages and sort by received time (newest first)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        # Look specifically for Van Paper emails
        for message in messages:
            try:
                # Skip if no received time
                if not hasattr(message, 'ReceivedTime') or not message.ReceivedTime:
                    continue
                
                # Skip if too old
                if message.ReceivedTime < cutoff_time:
                    continue
                
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
                    print(f"🎯 FOUND Van Paper email!")
                    print(f"   📅 Received: {message.ReceivedTime.strftime('%I:%M %p')}")
                    print(f"   📧 Subject: {subject}")
                    print(f"   📎 Excel file: {excel_attachment.FileName}")
                    
                    return {
                        'message': message,
                        'attachment': excel_attachment,
                        'received_time': message.ReceivedTime
                    }
                    
            except Exception as e:
                continue
        
        print("❌ No recent Van Paper email found")
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
    """Main automation function"""
    
    print("🤖 Van Paper Email Automation")
    print("=" * 50)
    print(f"🕐 Started at: {datetime.now().strftime('%Y-%m-%d %I:%M:%S %p CST')}")
    
    # Load configuration
    config = load_config()
    if not config:
        print("❌ Failed to load configuration")
        return False
    
    print("✅ Configuration loaded")
    
    # Find the Van Paper email
    email_data = find_van_paper_email()
    if not email_data:
        print("\n❌ No Van Paper email found in the last 2 hours")
        print("💡 This is normal if:")
        print("   - Running outside of 7:30 AM or 2:00 PM CST schedule")
        print("   - Van Paper report hasn't been sent yet")
        print("   - Email is taking longer than usual to arrive")
        return False
    
    # Process the email
    if not process_van_paper_email(email_data):
        print("❌ Failed to process Van Paper email")
        return False
    
    # Update the live app
    if not update_live_app(email_data['received_time']):
        print("⚠️ Live app update had issues")
        return False
    
    print("\n🎉 SUCCESS! Van Paper automation completed!")
    print(f"📧 Processed email from: {email_data['received_time'].strftime('%I:%M %p')}")
    print("🌐 Live app updated with fresh data!")
    
    return True

if __name__ == "__main__":
    success = main()
    
    print("\n" + "=" * 50)
    if success:
        print("✅ Automation completed successfully!")
    else:
        print("❌ Automation completed with issues")
    
    print(f"🕐 Finished at: {datetime.now().strftime('%Y-%m-%d %I:%M:%S %p CST')}")
