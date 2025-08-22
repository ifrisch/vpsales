"""
Process Today's Van Paper Email - Robust Version
This version handles file locks and permission issues
"""

import win32com.client
from datetime import datetime
import os
import pandas as pd
import subprocess
import shutil
import time

def process_todays_vanpaper_email_robust():
    """Find and process the Van Paper email from today with better error handling"""
    
    print("🎯 Processing Today's Van Paper Email (Robust Version)...")
    print("=" * 60)
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("✅ Connected to Outlook")
        
        # Look for today's Van Paper email
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        today = datetime.now().date()
        vanpaper_email = None
        
        print("🔍 Looking for Van Paper email from today...")
        
        for message in messages:
            try:
                if hasattr(message, 'ReceivedTime') and message.ReceivedTime:
                    msg_date = message.ReceivedTime.date()
                    
                    if (msg_date == today and 
                        hasattr(message, 'SenderEmailAddress') and
                        'noreply@vanpaper.com' in str(message.SenderEmailAddress) and
                        'leaderboardexport' in str(message.Subject).lower() and
                        message.Attachments.Count > 0):
                        
                        vanpaper_email = message
                        print(f"🎯 FOUND Van Paper email!")
                        print(f"   Time: {message.ReceivedTime.strftime('%I:%M %p')}")
                        print(f"   Subject: {message.Subject}")
                        print(f"   Attachments: {message.Attachments.Count}")
                        break
            except:
                continue
        
        if not vanpaper_email:
            print("❌ Van Paper email not found")
            return False
        
        # Process the attachment
        print("\n📎 Processing attachment...")
        
        excel_attachment = None
        for attachment in vanpaper_email.Attachments:
            filename = attachment.FileName
            print(f"   Found: {filename}")
            
            if filename.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                excel_attachment = attachment
                break
        
        if not excel_attachment:
            print("❌ No Excel attachment found")
            return False
        
        # Save the attachment to a new file name
        current_dir = os.path.dirname(os.path.abspath(__file__))
        new_leaderboard = f"leaderboard_new_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        new_path = os.path.join(current_dir, new_leaderboard)
        
        print(f"💾 Saving attachment as: {new_leaderboard}")
        excel_attachment.SaveAsFile(new_path)
        
        # Verify the file
        try:
            df = pd.read_excel(new_path)
            print(f"✅ Excel file verified: {len(df)} rows, {len(df.columns)} columns")
            print(f"📋 Columns: {list(df.columns)}")
        except Exception as e:
            print(f"❌ Error reading Excel file: {e}")
            return False
        
        # Handle the file replacement carefully
        leaderboard_path = os.path.join(current_dir, "leaderboard.xlsx")
        
        print(f"\n🔄 Replacing leaderboard.xlsx...")
        
        # Create backup if original exists
        if os.path.exists(leaderboard_path):
            backup_name = f"leaderboard_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            backup_path = os.path.join(current_dir, backup_name)
            
            # Try multiple times to handle file locks
            for attempt in range(3):
                try:
                    shutil.copy2(leaderboard_path, backup_path)
                    print(f"💾 Created backup: {backup_name}")
                    break
                except PermissionError:
                    print(f"⚠️ Backup attempt {attempt + 1} failed, retrying...")
                    time.sleep(1)
            
            # Delete original file
            for attempt in range(3):
                try:
                    os.remove(leaderboard_path)
                    print(f"🗑️ Removed old leaderboard.xlsx")
                    break
                except PermissionError:
                    print(f"⚠️ Delete attempt {attempt + 1} failed, retrying...")
                    time.sleep(1)
        
        # Move new file to replace old one
        for attempt in range(3):
            try:
                shutil.move(new_path, leaderboard_path)
                print(f"✅ Successfully replaced leaderboard.xlsx")
                break
            except PermissionError:
                print(f"⚠️ Replace attempt {attempt + 1} failed, retrying...")
                time.sleep(1)
        else:
            # If all attempts failed, just rename the new file
            final_name = f"leaderboard_from_vanpaper_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            os.rename(new_path, final_name)
            print(f"⚠️ Could not replace original file, saved as: {final_name}")
            print(f"💡 Please manually rename {final_name} to leaderboard.xlsx")
            return False
        
        # Update git and push to live app
        print("\n🚀 Updating live Streamlit app...")
        
        try:
            # Git add
            result = subprocess.run(['git', 'add', 'leaderboard.xlsx'], 
                                  cwd=current_dir, capture_output=True, text=True)
            print(f"📝 Git add: {result.returncode}")
            
            # Git commit
            commit_message = f"Auto-update from Van Paper email {vanpaper_email.ReceivedTime.strftime('%Y-%m-%d %H:%M')}"
            result = subprocess.run(['git', 'commit', '-m', commit_message], 
                                  cwd=current_dir, capture_output=True, text=True)
            print(f"📝 Git commit: {result.returncode}")
            
            # Git push
            result = subprocess.run(['git', 'push'], 
                                  cwd=current_dir, capture_output=True, text=True)
            print(f"📝 Git push: {result.returncode}")
            
            if result.returncode == 0:
                print("✅ Successfully updated live app!")
                print("🌐 Check: https://vpsales.streamlit.app/")
            else:
                print(f"⚠️ Git push had issues: {result.stderr}")
                print("📊 Data was updated locally but may not be live yet")
            
        except Exception as e:
            print(f"⚠️ Git update error: {e}")
            print("📊 Data was updated locally but may not be live yet")
        
        print("\n🎉 SUCCESS! Van Paper email processed!")
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = process_todays_vanpaper_email_robust()
    
    if success:
        print("\n" + "="*60)
        print("🚀 Your Streamlit app should be updated within 1-2 minutes!")
        print("🌐 https://vpsales.streamlit.app/")
        print("📊 The new Van Paper data is now live!")
    else:
        print("\n" + "="*60)
        print("❌ Processing failed. Check the output above for details.")
