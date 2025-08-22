"""
Extended Van Paper Email Search
Look at the last 24 hours to find all Van Paper emails
"""

import win32com.client
from datetime import datetime, timedelta

def find_all_recent_vanpaper_emails():
    """Find all Van Paper emails from the last 24 hours"""
    
    print("🔍 Extended Van Paper Email Search (Last 24 Hours)")
    print("=" * 55)
    print(f"🕐 Current time: {datetime.now().strftime('%Y-%m-%d %I:%M %p CST')}")
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("✅ Connected to Outlook")
        
        # Look for emails from the last 24 hours
        cutoff_time = datetime.now() - timedelta(hours=24)
        
        print(f"📅 Looking for emails since {cutoff_time.strftime('%Y-%m-%d %I:%M %p')}")
        
        # Get messages and sort by received time (newest first)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        all_vanpaper_emails = []
        matching_emails = []
        
        print("\n🔎 Scanning recent emails...")
        
        count = 0
        for message in messages:
            count += 1
            if count > 300:  # Scan more emails
                break
                
            try:
                # Skip if no received time
                if not hasattr(message, 'ReceivedTime') or not message.ReceivedTime:
                    continue
                
                # Skip if too old
                if message.ReceivedTime < cutoff_time:
                    continue
                
                # Check sender and subject
                sender = getattr(message, 'SenderEmailAddress', '')
                subject = getattr(message, 'Subject', '')
                
                # Look for any Van Paper related emails (broader search)
                if ('vanpaper' in str(sender).lower() or 
                    'vanpaper' in str(subject).lower() or
                    'noreply@vanpaper.com' in str(sender).lower() or
                    'leaderboard' in str(subject).lower()):
                    
                    email_info = {
                        'time': message.ReceivedTime,
                        'sender': sender,
                        'subject': subject,
                        'attachments': message.Attachments.Count,
                        'attachment_names': [att.FileName for att in message.Attachments] if message.Attachments.Count > 0 else [],
                        'matches_criteria': False
                    }
                    
                    # Check if this would match our automation criteria
                    if ('noreply@vanpaper.com' in str(sender).lower() and
                        'leaderboardexport' in str(subject).lower() and
                        message.Attachments.Count > 0):
                        email_info['matches_criteria'] = True
                        matching_emails.append(email_info)
                    
                    all_vanpaper_emails.append(email_info)
                    
            except Exception as e:
                continue
        
        print(f"📊 Found {len(all_vanpaper_emails)} Van Paper-related emails")
        print(f"🎯 Found {len(matching_emails)} emails that match automation criteria")
        
        if all_vanpaper_emails:
            print("\n📧 ALL VAN PAPER-RELATED EMAILS (Last 24 Hours):")
            print("-" * 60)
            for i, email in enumerate(all_vanpaper_emails, 1):
                status = "✅ MATCHES" if email['matches_criteria'] else "❌ No match"
                print(f"{i}. {email['time'].strftime('%Y-%m-%d %I:%M %p')} - {status}")
                print(f"   Subject: {email['subject']}")
                print(f"   From: {email['sender']}")
                print(f"   Attachments: {email['attachments']}")
                if email['attachment_names']:
                    print(f"   Files: {', '.join(email['attachment_names'])}")
                print()
                
        if matching_emails:
            print(f"\n🎯 EMAILS THAT MATCH AUTOMATION CRITERIA:")
            print("-" * 50)
            for i, email in enumerate(matching_emails, 1):
                print(f"{i}. {email['time'].strftime('%Y-%m-%d %I:%M %p')}")
                print(f"   Subject: {email['subject']}")
                print(f"   From: {email['sender']}")
                print(f"   Excel files: {[f for f in email['attachment_names'] if f.lower().endswith(('.xlsx', '.xls', '.xlsm'))]}")
                print()
                
                # Check if this is within the 3-hour window
                time_diff = datetime.now() - email['time']
                hours_ago = time_diff.total_seconds() / 3600
                print(f"   ⏰ {hours_ago:.1f} hours ago")
                if hours_ago <= 3:
                    print(f"   ✅ Within 3-hour automation window")
                else:
                    print(f"   ❌ Outside 3-hour automation window")
                print("-" * 30)
        else:
            print("\n❌ No emails found that match automation criteria")
            print("\nAutomation looks for:")
            print("- Sender: noreply@vanpaper.com")
            print("- Subject containing: leaderboardexport")  
            print("- Has Excel attachments")
            print("- Within last 3 hours")
        
        return matching_emails
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return []

if __name__ == "__main__":
    emails = find_all_recent_vanpaper_emails()
    
    print("\n" + "=" * 60)
    if emails:
        print("🎉 Found matching Van Paper emails!")
        print("💡 If automation isn't picking them up, they might be outside the 3-hour window")
    else:
        print("❌ No matching Van Paper emails found in last 24 hours")
        print("💡 Check if the sender/subject format has changed")
    
    input("\nPress Enter to continue...")
