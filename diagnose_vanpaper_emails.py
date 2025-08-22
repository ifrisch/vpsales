"""
Van Paper Email Diagnostic - Find All Recent Van Paper Emails
This will show us exactly what Van Paper emails are in your inbox
"""

import win32com.client
from datetime import datetime, timedelta

def find_all_vanpaper_emails():
    """Find all Van Paper emails from today"""
    
    print("🔍 Van Paper Email Diagnostic")
    print("=" * 40)
    print(f"🕐 Current time: {datetime.now().strftime('%I:%M %p CST')}")
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("✅ Connected to Outlook")
        
        # Look for emails from today
        today_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        
        print(f"📅 Looking for emails since {today_start.strftime('%I:%M %p')}")
        
        # Get messages and sort by received time (newest first)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        vanpaper_emails = []
        
        print("\n🔎 Scanning for Van Paper emails...")
        
        count = 0
        for message in messages:
            count += 1
            if count > 200:  # Don't scan too many
                break
                
            try:
                # Skip if no received time
                if not hasattr(message, 'ReceivedTime') or not message.ReceivedTime:
                    continue
                
                # Skip if before today
                if message.ReceivedTime < today_start:
                    continue
                
                # Check sender
                sender = getattr(message, 'SenderEmailAddress', '')
                subject = getattr(message, 'Subject', '')
                
                # Look for any Van Paper related emails
                if ('vanpaper' in str(sender).lower() or 
                    'vanpaper' in str(subject).lower() or
                    'noreply@vanpaper.com' in str(sender).lower()):
                    
                    vanpaper_emails.append({
                        'time': message.ReceivedTime,
                        'sender': sender,
                        'subject': subject,
                        'attachments': message.Attachments.Count,
                        'attachment_names': [att.FileName for att in message.Attachments] if message.Attachments.Count > 0 else []
                    })
                    
            except Exception as e:
                continue
        
        print(f"📊 Found {len(vanpaper_emails)} Van Paper emails from today")
        
        if vanpaper_emails:
            print("\n📧 VAN PAPER EMAILS FROM TODAY:")
            print("-" * 50)
            for i, email in enumerate(vanpaper_emails, 1):
                print(f"{i}. {email['time'].strftime('%I:%M %p')} - {email['subject']}")
                print(f"   From: {email['sender']}")
                print(f"   Attachments: {email['attachments']}")
                if email['attachment_names']:
                    print(f"   Files: {', '.join(email['attachment_names'])}")
                print()
                
                # Check if this would match our automation criteria
                if ('noreply@vanpaper.com' in str(email['sender']).lower() and
                    'leaderboardexport' in str(email['subject']).lower() and
                    email['attachments'] > 0):
                    print(f"   ✅ MATCHES automation criteria!")
                else:
                    print(f"   ❌ Does not match automation criteria")
                    if 'noreply@vanpaper.com' not in str(email['sender']).lower():
                        print(f"      - Sender doesn't match 'noreply@vanpaper.com'")
                    if 'leaderboardexport' not in str(email['subject']).lower():
                        print(f"      - Subject doesn't contain 'leaderboardexport'")
                    if email['attachments'] == 0:
                        print(f"      - No attachments")
                print("-" * 30)
        else:
            print("\n❌ No Van Paper emails found from today")
            print("\n💡 Possible reasons:")
            print("- Van Paper emails might be from yesterday")
            print("- Sender might be different than expected")
            print("- Emails might be in Junk/Spam folder")
        
        return vanpaper_emails
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return []

if __name__ == "__main__":
    emails = find_all_vanpaper_emails()
    
    print("\n" + "=" * 50)
    print("💡 This diagnostic helps us understand:")
    print("- What Van Paper emails are actually in your inbox")
    print("- Whether they match the automation criteria")
    print("- Why the automation might not be finding them")
    
    input("\nPress Enter to continue...")
