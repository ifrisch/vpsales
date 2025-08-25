"""
Process 7:29 AM Van Paper Email
"""

import win32com.client
import pandas as pd
import os
from datetime import datetime, timedelta
import subprocess
import shutil

def process_729_email():
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("Looking for 7:29 AM Van Paper email...")
        
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        target_email = None
        today = datetime.now().date()
        
        for message in messages:
            try:
                if not hasattr(message, 'ReceivedTime'):
                    continue
                
                if message.ReceivedTime.date() == today:
                    if (message.ReceivedTime.hour == 7 and 
                        28 <= message.ReceivedTime.minute <= 30):
                        
                        sender = str(getattr(message, 'SenderEmailAddress', ''))
                        subject = str(getattr(message, 'Subject', ''))
                        
                        if ('noreply@vanpaper.com' in sender.lower() and
                            'leaderboardexport' in subject.lower()):
                            target_email = message
                            print(f"FOUND: {message.ReceivedTime.strftime('%I:%M:%S %p')} - {subject}")
                            break
                            
            except Exception:
                continue
        
        if not target_email:
            print("Could not find 7:29 AM email")
            return False
        
        if target_email.Attachments.Count == 0:
            print("No attachments found")
            return False
        
        # Process the attachment
        attachment = target_email.Attachments[0]
        filename = attachment.FileName
        print(f"Processing: {filename}")
        
        # Save attachment
        temp_path = os.path.join(os.getcwd(), f"temp_{filename}")
        attachment.SaveAsFile(temp_path)
        
        # Backup current file
        current_file = "leaderboard_new.xlsx"
        backup_name = f"leaderboard_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        if os.path.exists(current_file):
            shutil.copy2(current_file, backup_name)
            print(f"Created backup: {backup_name}")
        
        # Update main file
        shutil.copy2(temp_path, current_file)
        print(f"Updated {current_file}")
        
        # Save timestamped copy
        timestamped_name = f"leaderboard_from_vanpaper_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        shutil.copy2(temp_path, timestamped_name)
        
        # Clean up
        os.remove(temp_path)
        
        # Verify data
        try:
            df = pd.read_excel(current_file)
            print(f"Data verified: {len(df)} rows loaded")
        except Exception as e:
            print(f"Data verification failed: {e}")
        
        # Git update
        git_commands = [
            ["git", "add", "."],
            ["git", "commit", "-m", f"MANUAL: Process 7:29 AM Van Paper email from {target_email.ReceivedTime.strftime('%Y-%m-%d')}"],
            ["git", "push"]
        ]
        
        for cmd in git_commands:
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, cwd=os.getcwd())
                print(f"{' '.join(cmd)}: {result.returncode}")
            except Exception as e:
                print(f"Git error: {e}")
        
        print("Successfully processed 7:29 AM email!")
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        return False

if __name__ == "__main__":
    success = process_729_email()
    if success:
        print("\nSUCCESS: 7:29 AM email processed!")
        print("Live app should update at: https://vpsales.streamlit.app/")
    else:
        print("\nFAILED to process email")
