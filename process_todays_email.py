"""
Process Today's Van Paper Email
This will find and process the 8:55 AM Van Paper email we just found
"""

import win32com.client
from datetime import datetime, timedelta
import os
import pandas as pd
import subprocess
import shutil

def process_todays_vanpaper_email():
    """Find and process the Van Paper email from today"""
    
    print("🎯 Processing Today's Van Paper Email...")
    print("=" * 50)
    
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
        
        # Save the attachment
        current_dir = os.path.dirname(os.path.abspath(__file__))
        temp_filename = f"vanpaper_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        temp_path = os.path.join(current_dir, temp_filename)
        
        print(f"💾 Saving attachment as: {temp_filename}")
        excel_attachment.SaveAsFile(temp_path)
        
        # Verify the file
        try:
            df = pd.read_excel(temp_path)
            print(f"✅ Excel file verified: {len(df)} rows, {len(df.columns)} columns")
            print(f"📋 Columns: {list(df.columns)}")
        except Exception as e:
            print(f"❌ Error reading Excel file: {e}")
            return False
        
        # Create backup of current leaderboard.xlsx
        leaderboard_path = os.path.join(current_dir, "leaderboard.xlsx")
        if os.path.exists(leaderboard_path):
            backup_name = f"leaderboard_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            backup_path = os.path.join(current_dir, backup_name)
            shutil.copy2(leaderboard_path, backup_path)
            print(f"💾 Created backup: {backup_name}")
        
        # Replace leaderboard.xlsx with new data
        shutil.copy2(temp_path, leaderboard_path)
        print(f"✅ Updated leaderboard.xlsx")
        
        # Update git and push to live app
        print("\n🚀 Updating live Streamlit app...")
        
        try:
            # Git add
            subprocess.run(['git', 'add', 'leaderboard.xlsx'], 
                         cwd=current_dir, capture_output=True, check=True)
            
            # Git commit
            commit_message = f"Auto-update from Van Paper email {vanpaper_email.ReceivedTime.strftime('%Y-%m-%d %H:%M')}"
            subprocess.run(['git', 'commit', '-m', commit_message], 
                         cwd=current_dir, capture_output=True, check=True)
            
            # Git push
            subprocess.run(['git', 'push'], 
                         cwd=current_dir, capture_output=True, check=True)
            
            print("✅ Successfully updated live app!")
            print("🌐 Check: https://vpsales.streamlit.app/")
            
        except subprocess.CalledProcessError as e:
            print(f"⚠️ Git update had issues: {e}")
            print("📊 Data was updated locally but may not be live yet")
        
        # Clean up temp file
        os.remove(temp_path)
        print(f"🗑️ Cleaned up temporary file")
        
        print("\n🎉 SUCCESS! Van Paper email processed and live app updated!")
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = process_todays_vanpaper_email()
    
    if success:
        print("\n" + "="*50)
        print("🚀 Your Streamlit app should be updated within 1 minute!")
        print("🌐 https://vpsales.streamlit.app/")
    else:
        print("\n" + "="*50)
        print("❌ Processing failed. Check the output above for details.")
