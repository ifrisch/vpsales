#!/usr/bin/env python3
"""
Silent Van Paper Automation
ZERO user interaction - designed for scheduled tasks
"""

import win32com.client
import os
import shutil
import subprocess
import pandas as pd
import configparser
from datetime import datetime, timedelta
from pathlib import Path
import sys

def load_config():
    """Load configuration silently"""
    config = configparser.ConfigParser()
    config_file = Path(__file__).parent / "automation_config.txt"
    
    if config_file.exists():
        config.read(config_file)
        return config
    else:
        return None

def find_recent_van_paper_email():
    """Find Van Paper emails from the last 3 hours - silent operation"""
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        # Look for emails from the last 3 hours
        cutoff_time = datetime.now() - timedelta(hours=3)
        
        # Get messages and sort by received time (newest first)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        van_paper_emails = []
        
        for message in messages:
            try:
                if not hasattr(message, 'ReceivedTime') or not message.ReceivedTime:
                    continue
                
                # Fix timezone comparison
                try:
                    msg_time = message.ReceivedTime
                    if hasattr(msg_time, 'replace'):
                        msg_time_naive = msg_time.replace(tzinfo=None)
                    else:
                        msg_time_naive = msg_time
                    
                    if msg_time_naive < cutoff_time:
                        break
                except Exception:
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
                    van_paper_emails.append({
                        'message': message,
                        'attachment': excel_attachment,
                        'received_time': message.ReceivedTime
                    })
                    
            except Exception:
                continue
        
        if van_paper_emails:
            return van_paper_emails[0]  # Return most recent
        else:
            return None
        
    except Exception:
        return None

def process_van_paper_email(email_data):
    """Process the Van Paper email silently"""
    
    try:
        current_dir = Path(__file__).parent
        
        # Create timestamp for files
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Save the Excel attachment
        temp_excel = current_dir / f"vanpaper_temp_{timestamp}.xlsx"
        email_data['attachment'].SaveAsFile(str(temp_excel))
        
        # Create backup of current file
        main_leaderboard = current_dir / "leaderboard_new.xlsx"
        if main_leaderboard.exists():
            backup_name = f"leaderboard_backup_{timestamp}.xlsx"
            backup_path = current_dir / backup_name
            shutil.copy2(main_leaderboard, backup_path)
        
        # Replace the main leaderboard file
        shutil.copy2(temp_excel, main_leaderboard)
        
        # Save a timestamped copy
        timestamped_copy = current_dir / f"leaderboard_from_vanpaper_{timestamp}.xlsx"
        shutil.copy2(temp_excel, timestamped_copy)
        
        # Clean up temp file
        temp_excel.unlink()
        
        # Git update - silently
        git_commands = [
            ["git", "add", "."],
            ["git", "commit", "-m", f"Auto-update from Van Paper {email_data['received_time'].strftime('%I:%M %p')} on {email_data['received_time'].strftime('%Y-%m-%d')}"],
            ["git", "push"]
        ]
        
        for cmd in git_commands:
            try:
                subprocess.run(cmd, capture_output=True, text=True, cwd=current_dir, timeout=30)
            except Exception:
                pass
        
        return True
        
    except Exception:
        return False

def main():
    """Main silent automation function"""
    
    # Silent operation - minimal output
    start_time = datetime.now()
    
    # Check if it's a business day
    current_day = start_time.weekday()  # 0=Monday, 6=Sunday
    if current_day >= 5:  # Saturday or Sunday
        return True
    
    # Load configuration
    config = load_config()
    if not config:
        return False
    
    # Find recent Van Paper emails
    email_data = find_recent_van_paper_email()
    
    if email_data:
        # Process the email
        success = process_van_paper_email(email_data)
        # Create sync timestamp regardless of success
        with open("last_sync.txt", "w") as f:
            f.write(start_time.strftime('%Y-%m-%d %H:%M:%S'))
        return success
    else:
        # No emails found - still create sync timestamp to show we checked
        with open("last_sync.txt", "w") as f:
            f.write(start_time.strftime('%Y-%m-%d %H:%M:%S'))
        return True

if __name__ == "__main__":
    # ZERO user interaction - just run and exit
    success = main()
    sys.exit(0 if success else 1)
