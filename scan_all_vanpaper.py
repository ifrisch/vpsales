"""
Find ALL Van Paper emails from today - comprehensive scan
"""

import win32com.client
from datetime import datetime, timedelta

def find_all_vanpaper_today():
    """Find ALL Van Paper emails from today"""
    
    print("🔍 Finding ALL Van Paper Emails From Today")
    print("=" * 45)
    print(f"🕐 Current time: {datetime.now().strftime('%Y-%m-%d %I:%M %p CST')}")
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("✅ Connected to Outlook")
        
        # Look for emails from today
        today = datetime.now().date()
        
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        vanpaper_emails = []
        
        print(f"🔎 Scanning ALL emails from today...")
        
        count = 0
        for message in messages:
            count += 1
            if count > 200:  # Reasonable limit
                break
                
            try:
                # Skip if no received time
                if not hasattr(message, 'ReceivedTime') or not message.ReceivedTime:
                    continue
                
                # Check if from today
                if message.ReceivedTime.date() == today:
                    
                    sender = getattr(message, 'SenderEmailAddress', 'Unknown')
                    subject = getattr(message, 'Subject', 'No Subject')
                    
                    # Check if this is Van Paper related
                    sender_lower = str(sender).lower()
                    subject_lower = str(subject).lower()
                    
                    if ('vanpaper' in sender_lower or 
                        'noreply@vanpaper.com' in sender_lower or
                        'leaderboard' in subject_lower or
                        'leaderboardexport' in subject_lower):
                        
                        email_info = {
                            'time': message.ReceivedTime,
                            'sender': sender,
                            'subject': subject,
                            'attachments': message.Attachments.Count,
                            'attachment_names': [att.FileName for att in message.Attachments] if message.Attachments.Count > 0 else []
                        }
                        
                        vanpaper_emails.append(email_info)
                
            except Exception as e:
                continue
        
        print(f"🎯 Found {len(vanpaper_emails)} Van Paper emails from today")
        
        if vanpaper_emails:
            print(f"\n🎯 ALL VAN PAPER EMAILS FROM TODAY:")
            print("-" * 50)
            for i, email in enumerate(vanpaper_emails, 1):
                time_str = email['time'].strftime('%I:%M:%S %p')
                print(f"{i}. {time_str}")
                print(f"   Subject: {email['subject']}")
                print(f"   From: {email['sender']}")
                print(f"   Attachments: {email['attachments']}")
                if email['attachment_names']:
                    print(f"   Files: {', '.join(email['attachment_names'])}")
                print("-" * 30)
                
            # Find emails after 11:16 AM (the last one we processed)
            recent_emails = [e for e in vanpaper_emails if e['time'].hour > 11 or (e['time'].hour == 11 and e['time'].minute > 16)]
            
            if recent_emails:
                print(f"\n🚨 EMAILS AFTER 11:16 AM (SHOULD BE PROCESSED):")
                print("-" * 45)
                for i, email in enumerate(recent_emails, 1):
                    time_str = email['time'].strftime('%I:%M:%S %p')
                    print(f"{i}. {time_str} - NOT PROCESSED!")
                    print(f"   Subject: {email['subject']}")
                    print(f"   Files: {', '.join(email['attachment_names'])}")
                    print()
            else:
                print(f"\n✅ No new emails after 11:16 AM")
        else:
            print("\n❌ NO VAN PAPER EMAILS FOUND FROM TODAY")
        
        return vanpaper_emails
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return []

if __name__ == "__main__":
    emails = find_all_vanpaper_today()
    
    print("\n" + "=" * 60)
    if emails:
        recent_count = len([e for e in emails if e['time'].hour > 11 or (e['time'].hour == 11 and e['time'].minute > 16)])
        if recent_count > 0:
            print(f"🚨 FOUND {recent_count} UNPROCESSED VAN PAPER EMAILS!")
            print("This explains why automation should have run but didn't work!")
        else:
            print("✅ All Van Paper emails have been processed")
    else:
        print("❌ No Van Paper emails found")
    
    input("\nPress Enter to continue...")
