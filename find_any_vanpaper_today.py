"""
Find ANY Van Paper Email From Today - No Time Restrictions
This will find ANY email containing Van Paper keywords from today
"""

import win32com.client
from datetime import datetime, timedelta

def find_any_vanpaper_today():
    """Find ANY Van Paper related email from today, regardless of time"""
    
    print("🔍 Finding ANY Van Paper Email From Today")
    print("=" * 45)
    print(f"🕐 Current time: {datetime.now().strftime('%Y-%m-%d %I:%M %p CST')}")
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("✅ Connected to Outlook")
        
        # Look for emails from today (no time restrictions)
        today = datetime.now().date()
        today_start = datetime.combine(today, datetime.min.time())
        
        print(f"📅 Looking for ANY emails from {today_start.strftime('%Y-%m-%d')} onwards")
        
        # Get messages and sort by received time (newest first)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        all_todays_emails = []
        vanpaper_related = []
        
        print("\n🔎 Scanning ALL emails from today...")
        
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
                    
                    email_info = {
                        'time': message.ReceivedTime,
                        'sender': sender,
                        'subject': subject,
                        'attachments': message.Attachments.Count,
                        'attachment_names': [att.FileName for att in message.Attachments] if message.Attachments.Count > 0 else []
                    }
                    
                    all_todays_emails.append(email_info)
                    
                    # Check if this is Van Paper related
                    sender_lower = str(sender).lower()
                    subject_lower = str(subject).lower()
                    
                    if ('vanpaper' in sender_lower or 
                        'vanpaper' in subject_lower or
                        'noreply@vanpaper.com' in sender_lower or
                        'leaderboard' in subject_lower or
                        'leaderboardexport' in subject_lower):
                        vanpaper_related.append(email_info)
                
            except Exception as e:
                continue
        
        print(f"📊 Found {len(all_todays_emails)} total emails from today")
        print(f"🎯 Found {len(vanpaper_related)} Van Paper related emails")
        
        if all_todays_emails:
            print(f"\n📧 ALL EMAILS FROM TODAY ({len(all_todays_emails)} total):")
            print("-" * 60)
            for i, email in enumerate(all_todays_emails, 1):
                time_str = email['time'].strftime('%I:%M:%S %p')
                print(f"{i}. {time_str} - {email['subject'][:40]}...")
                print(f"   From: {email['sender']}")
                if email['attachments'] > 0:
                    print(f"   📎 {email['attachments']} attachments")
                print()
                
                # Stop after showing first 10 to avoid too much output
                if i >= 10:
                    print(f"   ... and {len(all_todays_emails) - 10} more emails")
                    break
        
        if vanpaper_related:
            print(f"\n🎯 VAN PAPER RELATED EMAILS:")
            print("-" * 40)
            for i, email in enumerate(vanpaper_related, 1):
                time_str = email['time'].strftime('%I:%M:%S %p')
                print(f"{i}. {time_str}")
                print(f"   Subject: {email['subject']}")
                print(f"   From: {email['sender']}")
                print(f"   Attachments: {email['attachments']}")
                if email['attachment_names']:
                    print(f"   Files: {', '.join(email['attachment_names'])}")
                print("-" * 30)
        else:
            print("\n❌ NO VAN PAPER RELATED EMAILS FOUND FROM TODAY")
            print("\n💡 This means:")
            print("- No emails contain 'vanpaper' in sender or subject")
            print("- No emails from 'noreply@vanpaper.com'")
            print("- No emails contain 'leaderboard' in subject")
            print("- The 9:55 AM email might have different keywords")
        
        # Show a few sample emails to see what we ARE finding
        if all_todays_emails and not vanpaper_related:
            print(f"\n📧 SAMPLE OF WHAT WE ARE FINDING TODAY:")
            print("-" * 45)
            for i, email in enumerate(all_todays_emails[:5], 1):
                time_str = email['time'].strftime('%I:%M:%S %p')
                print(f"{i}. {time_str}")
                print(f"   Subject: {email['subject']}")
                print(f"   From: {email['sender']}")
                print()
        
        return vanpaper_related
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return []

if __name__ == "__main__":
    emails = find_any_vanpaper_today()
    
    print("\n" + "=" * 60)
    if emails:
        print("✅ Found Van Paper emails!")
    else:
        print("❌ No Van Paper emails found with our search criteria")
        print("💡 The 9:55 AM email you're seeing might:")
        print("   - Use different sender/subject text than expected")
        print("   - Be in a different Outlook data file")
        print("   - Not be fully synced to the Python Outlook connection")
    
    input("\nPress Enter to continue...")
