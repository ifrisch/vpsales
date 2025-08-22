"""
Check Outlook Profiles and Find the Right Email
This will help us connect to the correct Outlook profile/account
"""

import win32com.client
from datetime import datetime, timedelta
import os

def check_outlook_profiles():
    """Check all available Outlook profiles and accounts"""
    
    print("🔍 Checking Outlook Profiles and Accounts...")
    print("=" * 60)
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print("✅ Connected to Outlook")
        
        # Check all accounts
        print("\n📧 AVAILABLE EMAIL ACCOUNTS:")
        print("-" * 40)
        
        accounts = namespace.Accounts
        for i, account in enumerate(accounts, 1):
            print(f"{i}. {account.DisplayName}")
            print(f"   Type: {account.AccountType}")
            print(f"   Email: {getattr(account, 'SmtpAddress', 'N/A')}")
            print()
        
        # Check different folder types for each account
        print("\n📁 CHECKING FOLDERS IN DEFAULT ACCOUNT:")
        print("-" * 40)
        
        # Default folders
        folder_types = {
            6: "Inbox",
            5: "Sent Items", 
            3: "Deleted Items",
            23: "Junk Email",
            4: "Outbox",
            16: "Drafts"
        }
        
        for folder_id, folder_name in folder_types.items():
            try:
                folder = namespace.GetDefaultFolder(folder_id)
                print(f"📂 {folder_name}: {folder.Items.Count} items")
            except:
                print(f"📂 {folder_name}: Not accessible")
        
        # Now let's specifically look for emails from today
        print(f"\n🎯 SEARCHING FOR TODAY'S EMAILS (8/22/2025):")
        print("-" * 40)
        
        inbox = namespace.GetDefaultFolder(6)  # Inbox
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)  # Newest first
        
        today = datetime.now().date()
        today_emails = []
        
        count = 0
        for message in messages:
            count += 1
            if count > 200:  # Check more emails
                break
                
            try:
                if hasattr(message, 'ReceivedTime') and message.ReceivedTime:
                    msg_date = message.ReceivedTime.date()
                    
                    if msg_date == today:
                        today_emails.append({
                            'time': message.ReceivedTime.strftime('%I:%M %p'),
                            'subject': message.Subject,
                            'sender': getattr(message, 'SenderEmailAddress', 'Unknown'),
                            'attachments': message.Attachments.Count
                        })
            except:
                continue
        
        print(f"📊 Found {len(today_emails)} emails from today")
        
        if today_emails:
            print("\n📋 TODAY'S EMAILS:")
            for email in today_emails:
                attachment_text = f"📎 {email['attachments']}" if email['attachments'] > 0 else "📭"
                print(f"{email['time']} - {attachment_text} - {email['subject'][:50]}...")
                print(f"         From: {email['sender']}")
                print()
        
        # Look specifically for the 8:56 AM email
        print("\n🎯 LOOKING FOR 8:56 AM EMAIL:")
        print("-" * 40)
        
        found_856_email = False
        for email in today_emails:
            if "8:56" in email['time']:
                found_856_email = True
                print(f"🎯 FOUND IT!")
                print(f"   Time: {email['time']}")
                print(f"   Subject: {email['subject']}")
                print(f"   From: {email['sender']}")
                print(f"   Attachments: {email['attachments']}")
                
                # Let's get the actual message and check attachments
                for message in messages:
                    try:
                        if (hasattr(message, 'ReceivedTime') and 
                            message.ReceivedTime.strftime('%I:%M %p') == email['time'] and
                            message.Subject == email['subject']):
                            
                            print(f"\n📎 ATTACHMENT DETAILS:")
                            for j, attachment in enumerate(message.Attachments, 1):
                                filename = attachment.FileName
                                print(f"   {j}. {filename}")
                                if filename.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                                    print(f"      ✅ This is an Excel file!")
                            break
                    except:
                        continue
                break
        
        if not found_856_email:
            print("❌ 8:56 AM email not found in default inbox")
            print("\n💡 POSSIBLE SOLUTIONS:")
            print("1. The email might be in a different account")
            print("2. The email might be in a different folder (Junk, etc.)")
            print("3. The time zone might be different")
            print("4. We need to check all accounts, not just the default")
        
        return found_856_email
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

if __name__ == "__main__":
    check_outlook_profiles()
    print("\n" + "="*60)
    print("💡 If the 8:56 AM email isn't found, we need to check")
    print("   other Outlook accounts or folders!")
