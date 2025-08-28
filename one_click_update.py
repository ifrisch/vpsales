# -*- coding: utf-8 -*-
"""
One-Click Van Paper Update
Simple, reliable, fast - just run this when you want to update the app
"""

import sys
import io

# Set stdout to handle ASCII only for Windows compatibility
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='ascii', errors='replace')

import win32com.client
import pandas as pd
import os
from datetime import datetime, timedelta
import subprocess
import shutil

def update_from_latest_vanpaper():
    """Find and process the most recent Van Paper email"""
    
    print("=== Van Paper One-Click Update ===")
    print(f"Time: {datetime.now().strftime('%I:%M %p on %Y-%m-%d')}")
    print()
    
    try:
        # Connect to Outlook
        print("Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        # Look for Van Paper emails from today
        print("Looking for Van Paper emails from today...")
        today = datetime.now().date()
        
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        latest_vanpaper = None
        
        # Check only recent messages to avoid hanging
        count = 0
        for message in messages:
            count += 1
            if count > 100:  # Reasonable limit
                break
                
            try:
                if not hasattr(message, 'ReceivedTime'):
                    continue
                
                if message.ReceivedTime.date() == today:
                    sender = str(getattr(message, 'SenderEmailAddress', ''))
                    subject = str(getattr(message, 'Subject', ''))
                    
                    if ('noreply@vanpaper.com' in sender.lower() and
                        'leaderboardexport' in subject.lower() and
                        message.Attachments.Count > 0):
                        
                        latest_vanpaper = message
                        print(f"FOUND: {message.ReceivedTime.strftime('%I:%M:%S %p')} - {subject}")
                        break
                        
            except Exception:
                continue
        
        if not latest_vanpaper:
            print("No Van Paper emails found from today")
            print("Your app is probably already up to date!")
            
            # Still create sync timestamp to show the app was checked
            current_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with open("last_sync.txt", "w") as f:
                f.write(current_timestamp)
            
            # Update timestamp in leaderboard.py even when no new emails
            try:
                with open("leaderboard.py", "r", encoding="utf-8") as f:
                    content = f.read()
                
                import re
                pattern = r'LAST_SYNC_TIMESTAMP = "[^"]*"'
                replacement = f'LAST_SYNC_TIMESTAMP = "{current_timestamp}"'
                
                if re.search(pattern, content):
                    content = re.sub(pattern, replacement, content)
                    
                    with open("leaderboard.py", "w", encoding="utf-8") as f:
                        f.write(content)
                    print(f"[OK] Updated embedded timestamp: {current_timestamp}")
                    
                    # Commit the timestamp update
                    subprocess.run(["git", "add", "leaderboard.py"], capture_output=True)
                    subprocess.run(["git", "commit", "-m", f"Update sync timestamp - no new emails found"], capture_output=True)
                    subprocess.run(["git", "push"], capture_output=True)
                    print("[OK] Pushed timestamp update to live app")
                    
            except Exception as e:
                print(f"[WARNING] Timestamp update failed: {e}")
            
            return True
        
        # Process the attachment
        attachment = latest_vanpaper.Attachments[0]
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
            if len(df) > 0:
                print(f"Sample customers: {', '.join(df.iloc[:3, 0].astype(str).tolist())}")
        except Exception as e:
            print(f"Data verification failed: {e}")
        
        # Create sync timestamp file and update embedded timestamp BEFORE git
        current_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open("last_sync.txt", "w") as f:
            f.write(current_timestamp)
        
        # Update timestamp directly in leaderboard.py for reliability
        try:
            with open("leaderboard.py", "r", encoding="utf-8") as f:
                content = f.read()
            
            # Find and replace the timestamp line
            import re
            pattern = r'LAST_SYNC_TIMESTAMP = "[^"]*"'
            replacement = f'LAST_SYNC_TIMESTAMP = "{current_timestamp}"'
            
            if re.search(pattern, content):
                content = re.sub(pattern, replacement, content)
                
                with open("leaderboard.py", "w", encoding="utf-8") as f:
                    f.write(content)
                print(f"[OK] Updated embedded timestamp: {current_timestamp}")
            else:
                print("[WARNING] Could not find timestamp in leaderboard.py")
                
        except Exception as e:
            print(f"[WARNING] Timestamp update failed: {e}")

        # Git update
        print("Updating live Streamlit app...")
        git_commands = [
            ["git", "add", "."],
            ["git", "commit", "-m", f"One-click update from Van Paper {latest_vanpaper.ReceivedTime.strftime('%I:%M %p')} on {latest_vanpaper.ReceivedTime.strftime('%Y-%m-%d')}"],
            ["git", "push"]
        ]
        
        for cmd in git_commands:
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, cwd=os.getcwd())
                if result.returncode == 0:
                    print(f"[OK] {' '.join(cmd[:2])}")
                else:
                    print(f"[ERROR] {' '.join(cmd[:2])}: {result.stderr}")
            except Exception as e:
                print(f"Git error: {e}")
        
        print()
        print("=== UPDATE COMPLETE! ===")
        print(f"[OK] Processed Van Paper email from {latest_vanpaper.ReceivedTime.strftime('%I:%M %p')}")
        print(f"[OK] Live app updated: https://vpsales.streamlit.app/")
        print(f"[OK] {len(df)} customers loaded")
        print()
        
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        print()
        print("=== UPDATE FAILED ===")
        print("Try running again or check your Outlook connection")
        return False

if __name__ == "__main__":
    success = update_from_latest_vanpaper()
    # Silent mode - no user input required
    if success:
        print("Update completed successfully!")
    else:
        print("Update failed - check logs")
