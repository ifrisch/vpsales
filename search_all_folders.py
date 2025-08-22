"""
Search ALL Outlook Folders for 9:55 AM Email
This will check Inbox, Junk, Sent Items, and other folders
"""

import win32com.client
from datetime import datetime, timedelta

def search_all_outlook_folders():
    """Search ALL Outlook folders for the 9:55 AM email"""
    
    print("🔍 Searching ALL Outlook Folders for 9:55 AM Email")
    print("=" * 55)
    print(f"🕐 Current time: {datetime.now().strftime('%Y-%m-%d %I:%M %p CST')}")
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print("✅ Connected to Outlook")
        
        # Define timeframe around 9:55 AM
        today = datetime.now().date()
        start_time = datetime.combine(today, datetime.min.time().replace(hour=9, minute=50))
        end_time = datetime.combine(today, datetime.min.time().replace(hour=10, minute=5))
        
        print(f"📅 Looking for emails between {start_time.strftime('%I:%M %p')} and {end_time.strftime('%I:%M %p')}")
        
        # Check different folder types
        folder_types = {
            6: "Inbox",
            5: "Sent Items", 
            3: "Deleted Items",
            23: "Junk Email",
            4: "Outbox",
            16: "Drafts"
        }
        
        all_found_emails = []
        
        for folder_id, folder_name in folder_types.items():
            print(f"\n📂 Searching {folder_name}...")
            
            try:
                folder = namespace.GetDefaultFolder(folder_id)
                messages = folder.Items
                messages.Sort("[ReceivedTime]", True)
                
                folder_emails = []
                
                count = 0
                for message in messages:
                    count += 1
                    if count > 100:  # Limit per folder
                        break
                        
                    try:
                        # Skip if no received time
                        if not hasattr(message, 'ReceivedTime') or not message.ReceivedTime:
                            continue
                        
                        # Check if in our timeframe (9:50-10:05 AM)
                        if message.ReceivedTime >= start_time and message.ReceivedTime <= end_time:
                            
                            sender = getattr(message, 'SenderEmailAddress', 'Unknown')
                            subject = getattr(message, 'Subject', 'No Subject')
                            
                            email_info = {
                                'folder': folder_name,
                                'time': message.ReceivedTime,
                                'sender': sender,
                                'subject': subject,
                                'attachments': message.Attachments.Count,
                                'attachment_names': [att.FileName for att in message.Attachments] if message.Attachments.Count > 0 else []
                            }
                            
                            folder_emails.append(email_info)
                            all_found_emails.append(email_info)
                
                    except Exception as e:
                        # Skip problematic messages
                        continue
                
                if folder_emails:
                    print(f"   📧 Found {len(folder_emails)} emails in {folder_name}")
                    for email in folder_emails:
                        time_str = email['time'].strftime('%I:%M:%S %p')
                        print(f"   {time_str} - {email['subject'][:50]}...")
                        print(f"   From: {email['sender']}")
                        if email['attachments'] > 0:
                            print(f"   📎 {email['attachments']} attachments: {', '.join(email['attachment_names'])}")
                else:
                    print(f"   ❌ No emails found in {folder_name}")
                    
            except Exception as e:
                print(f"   ⚠️ Could not access {folder_name}: {e}")
        
        print(f"\n📊 TOTAL EMAILS FOUND IN 9:50-10:05 AM TIMEFRAME: {len(all_found_emails)}")
        
        if all_found_emails:
            print("\n🎯 ALL EMAILS FOUND:")
            print("-" * 70)
            for i, email in enumerate(all_found_emails, 1):
                time_str = email['time'].strftime('%I:%M:%S %p')
                print(f"{i}. [{email['folder']}] {time_str}")
                print(f"   Subject: {email['subject']}")
                print(f"   From: {email['sender']}")
                print(f"   Attachments: {email['attachments']}")
                if email['attachment_names']:
                    print(f"   Files: {', '.join(email['attachment_names'])}")
                
                # Check if this looks like Van Paper
                is_vanpaper = False
                if ('vanpaper' in str(email['sender']).lower() or 
                    'vanpaper' in str(email['subject']).lower() or
                    'noreply@vanpaper.com' in str(email['sender']).lower() or
                    'leaderboard' in str(email['subject']).lower()):
                    print(f"   🎯 THIS LOOKS LIKE A VAN PAPER EMAIL!")
                    is_vanpaper = True
                
                print("-" * 50)
        else:
            print("\n❌ NO EMAILS FOUND IN ANY FOLDER for 9:50-10:05 AM timeframe")
            print("\n💡 Possible explanations:")
            print("- Email might be in a different Outlook account")
            print("- Email timestamp might be in different time zone")
            print("- Email might be in a subfolder not checked")
            print("- Outlook might not be fully synced")
        
        # Also check what accounts are available
        print(f"\n📧 AVAILABLE OUTLOOK ACCOUNTS:")
        print("-" * 40)
        accounts = namespace.Accounts
        for i, account in enumerate(accounts, 1):
            print(f"{i}. {account.DisplayName}")
            try:
                smtp = getattr(account, 'SmtpAddress', 'N/A')
                print(f"   Email: {smtp}")
            except:
                print(f"   Email: Unable to retrieve")
        
        return all_found_emails
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return []

if __name__ == "__main__":
    emails = search_all_outlook_folders()
    
    print("\n" + "=" * 70)
    print("🔍 This search checked ALL major Outlook folders")
    print("📧 If your 9:55 AM email still isn't found, it might be:")
    print("   - In a different Outlook account")
    print("   - In a custom folder")
    print("   - Not fully synced to local Outlook cache")
    
    input("\nPress Enter to continue...")
