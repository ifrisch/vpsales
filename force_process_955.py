"""
Force Process 9:55 AM Van Paper Email
This will specifically find and process the 9:55 AM email
"""

import win32com.client
import pandas as pd
import os
from datetime import datetime, timedelta
import subprocess
import shutil

def force_process_955_email():
    """Find and process the 9:55 AM Van Paper email specifically"""
    
    print("🎯 Force Processing 9:55 AM Van Paper Email")
    print("=" * 45)
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("✅ Connected to Outlook")
        
        # Look for the specific 9:55 AM email
        target_time = datetime.now().replace(hour=9, minute=55, second=0, microsecond=0)
        print(f"🕐 Looking for email around {target_time.strftime('%I:%M %p')}")
        
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        target_email = None
        
        for message in messages:
            try:
                if not hasattr(message, 'ReceivedTime') or not message.ReceivedTime:
                    continue
                
                # Check if this is from today and around 9:55 AM
                if (message.ReceivedTime.date() == datetime.now().date() and
                    message.ReceivedTime.hour == 9 and
                    54 <= message.ReceivedTime.minute <= 56):
                    
                    sender = getattr(message, 'SenderEmailAddress', '')
                    subject = getattr(message, 'Subject', '')
                    
                    if ('noreply@vanpaper.com' in str(sender) and
                        'leaderboardexport' in str(subject)):
                        target_email = message
                        print(f"✅ Found target email!")
                        print(f"   Time: {message.ReceivedTime.strftime('%I:%M:%S %p')}")
                        print(f"   Subject: {subject}")
                        print(f"   Attachments: {message.Attachments.Count}")
                        break
                        
            except Exception as e:
                continue
        
        if not target_email:
            print("❌ Could not find the 9:55 AM email")
            return False
        
        # Process the attachment
        if target_email.Attachments.Count == 0:
            print("❌ Email has no attachments")
            return False
        
        attachment = target_email.Attachments[0]
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
            ["git", "commit", "-m", f"Auto-update leaderboard from Van Paper email at {target_email.ReceivedTime.strftime('%I:%M %p')} on {target_email.ReceivedTime.strftime('%Y-%m-%d')}"],
            ["git", "push"]
        ]
        
        for cmd in git_commands:
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, cwd=os.getcwd())
                print(f"✅ {' '.join(cmd)}: {result.returncode}")
                if result.returncode != 0 and result.stderr:
                    print(f"   Warning: {result.stderr.strip()}")
            except Exception as e:
                print(f"❌ Git command failed: {e}")
        
        print(f"\n✅ Successfully processed 9:55 AM Van Paper email!")
        print(f"🔗 Live app will update at: https://vpsales.streamlit.app/")
        
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = force_process_955_email()
    
    print(f"\n" + "=" * 50)
    if success:
        print("✅ 9:55 AM EMAIL SUCCESSFULLY PROCESSED!")
        print("🚀 Streamlit app will refresh with latest data")
    else:
        print("❌ Failed to process 9:55 AM email")
    
    input("\nPress Enter to continue...")
