"""
Process the LATEST Van Paper Email (1:04 PM)
Get the most recent leaderboard data
"""

import win32com.client
import pandas as pd
import os
from datetime import datetime, timedelta
import subprocess
import shutil

def process_latest_vanpaper():
    """Find and process the most recent Van Paper email"""
    
    print("🎯 Processing LATEST Van Paper Email")
    print("=" * 40)
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("✅ Connected to Outlook")
        
        # Look for the most recent Van Paper email from today
        today = datetime.now().date()
        
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        latest_vanpaper = None
        
        for message in messages:
            try:
                if not hasattr(message, 'ReceivedTime') or not message.ReceivedTime:
                    continue
                
                # Check if from today
                if message.ReceivedTime.date() == today:
                    
                    sender = getattr(message, 'SenderEmailAddress', '')
                    subject = getattr(message, 'Subject', '')
                    
                    if ('noreply@vanpaper.com' in str(sender) and
                        'leaderboardexport' in str(subject) and
                        message.Attachments.Count > 0):
                        
                        latest_vanpaper = message
                        print(f"✅ Found LATEST Van Paper email!")
                        print(f"   Time: {message.ReceivedTime.strftime('%I:%M:%S %p')}")
                        print(f"   Subject: {subject}")
                        print(f"   Attachments: {message.Attachments.Count}")
                        break
                        
            except Exception as e:
                continue
        
        if not latest_vanpaper:
            print("❌ Could not find any Van Paper emails")
            return False
        
        # Process the attachment
        attachment = latest_vanpaper.Attachments[0]
        filename = attachment.FileName
        
        print(f"📎 Processing attachment: {filename}")
        
        # Save attachment temporarily
        temp_path = os.path.join(os.getcwd(), f"temp_{filename}")
        attachment.SaveAsFile(temp_path)
        
        print(f"💾 Saved to: {temp_path}")
        
        # Create backup of current file
        current_file = "leaderboard_new.xlsx"
        backup_name = f"leaderboard_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        if os.path.exists(current_file):
            shutil.copy2(current_file, backup_name)
            print(f"📋 Backed up current file to: {backup_name}")
        
        # Copy the new file
        shutil.copy2(temp_path, current_file)
        print(f"✅ Updated {current_file}")
        
        # Also save with timestamp for reference
        timestamped_name = f"leaderboard_from_vanpaper_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        shutil.copy2(temp_path, timestamped_name)
        print(f"📁 Saved copy as: {timestamped_name}")
        
        # Clean up temp file
        os.remove(temp_path)
        
        # Verify the data
        try:
            df = pd.read_excel(current_file)
            print(f"📊 Data verification: {len(df)} rows loaded")
            if len(df) > 0:
                print(f"   First few customers: {', '.join(df.iloc[:3, 0].astype(str).tolist())}")
        except Exception as e:
            print(f"⚠️ Data verification failed: {e}")
        
        # Git update
        print(f"\n🔄 Updating Git Repository...")
        
        git_commands = [
            ["git", "add", "."],
            ["git", "commit", "-m", f"MANUAL: Latest Van Paper update from {latest_vanpaper.ReceivedTime.strftime('%I:%M %p')} on {latest_vanpaper.ReceivedTime.strftime('%Y-%m-%d')}"],
            ["git", "push"]
        ]
        
        for cmd in git_commands:
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, cwd=os.getcwd())
                print(f"✅ {' '.join(cmd)}: {result.returncode}")
                if result.returncode != 0:
                    print(f"   Output: {result.stdout}")
                    print(f"   Error: {result.stderr}")
            except Exception as e:
                print(f"❌ Git command failed: {e}")
        
        print(f"\n✅ Successfully processed LATEST Van Paper email!")
        print(f"🔗 Live app will update at: https://vpsales.streamlit.app/")
        
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = process_latest_vanpaper()
    
    print(f"\n" + "=" * 50)
    if success:
        print("✅ LATEST EMAIL SUCCESSFULLY PROCESSED!")
        print("🚀 Streamlit app will refresh with newest data")
    else:
        print("❌ Failed to process latest email")
    
    input("\nPress Enter to continue...")
