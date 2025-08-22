"""
Continuous Van Paper Monitor
Runs every 30 minutes during business hours and processes any new emails
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

def is_business_hours():
    """Check if it's currently business hours (7 AM - 4 PM, Mon-Fri)"""
    now = datetime.now()
    
    # Check if it's a weekday (Monday = 0, Sunday = 6)
    if now.weekday() > 4:  # Saturday or Sunday
        return False
    
    # Check if it's during business hours (7 AM - 4 PM)
    if now.hour < 7 or now.hour >= 16:
        return False
    
    return True

def find_new_van_paper_emails():
    """Find Van Paper emails from the last 45 minutes"""
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        # Look for emails from the last 45 minutes
        cutoff_time = datetime.now() - timedelta(minutes=45)
        
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
        
        return van_paper_emails[0] if van_paper_emails else None
        
    except Exception:
        return None

def process_van_paper_email(email_data):
    """Process the Van Paper email and update the leaderboard"""
    
    try:
        current_dir = Path(__file__).parent
        
        # Create timestamp for files
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Save the Excel attachment
        temp_excel = current_dir / f"vanpaper_temp_{timestamp}.xlsx"
        email_data['attachment'].SaveAsFile(str(temp_excel))
        
        # Replace the main leaderboard file
        main_leaderboard = current_dir / "leaderboard_new.xlsx"
        if main_leaderboard.exists():
            backup_name = f"leaderboard_backup_{timestamp}.xlsx"
            backup_path = current_dir / backup_name
            shutil.copy2(main_leaderboard, backup_path)
        
        shutil.copy2(temp_excel, main_leaderboard)
        
        # Save a timestamped copy
        timestamped_copy = current_dir / f"leaderboard_from_vanpaper_{timestamp}.xlsx"
        shutil.copy2(temp_excel, timestamped_copy)
        
        # Clean up
        temp_excel.unlink()
        
        # Git update
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
    """Main monitoring function"""
    
    log_time = datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')
    
    # Only run during business hours
    if not is_business_hours():
        return
    
    # Check for new Van Paper emails
    email_data = find_new_van_paper_emails()
    
    if email_data:
        # Process the email
        success = process_van_paper_email(email_data)
        
        if success:
            # Log success
            with open("automation.log", "a") as f:
                f.write(f"[{log_time}] SUCCESS: Processed Van Paper email from {email_data['received_time'].strftime('%I:%M %p')}\n")
        else:
            # Log failure
            with open("automation.log", "a") as f:
                f.write(f"[{log_time}] ERROR: Failed to process Van Paper email\n")
    else:
        # Log no emails found (but quietly)
        with open("automation.log", "a") as f:
            f.write(f"[{log_time}] INFO: No new Van Paper emails found\n")

if __name__ == "__main__":
    main()
