"""
Comprehensive Email Search - Find 9:55 AM Email
This will search for ALL emails around 9:55 AM and show exact details
"""

import win32com.client
from datetime import datetime, timedelta

def find_all_emails_around_955():
    """Find ALL emails received around 9:55 AM today"""
    
    print("🔍 Comprehensive Email Search Around 9:55 AM")
    print("=" * 50)
    print(f"🕐 Current time: {datetime.now().strftime('%Y-%m-%d %I:%M %p CST')}")
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("✅ Connected to Outlook")
        
        # Look for emails from 9:00 AM to 10:30 AM today
        today = datetime.now().date()
        start_time = datetime.combine(today, datetime.min.time().replace(hour=9, minute=0))
        end_time = datetime.combine(today, datetime.min.time().replace(hour=10, minute=30))
        
        print(f"📅 Looking for emails between {start_time.strftime('%I:%M %p')} and {end_time.strftime('%I:%M %p')}")
        
        # Get messages and sort by received time (newest first)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        emails_in_timeframe = []
        
        print("\n🔎 Scanning ALL emails in timeframe...")
        
        count = 0
        for message in messages:
            count += 1
            if count > 500:  # Scan many emails
                break
                
            try:
                # Skip if no received time
                if not hasattr(message, 'ReceivedTime') or not message.ReceivedTime:
                    continue
                
                # Check if in our timeframe
                if message.ReceivedTime >= start_time and message.ReceivedTime <= end_time:
                    
                    sender = getattr(message, 'SenderEmailAddress', 'Unknown')
                    subject = getattr(message, 'Subject', 'No Subject')
                    
                    emails_in_timeframe.append({
                        'time': message.ReceivedTime,
                        'sender': sender,
                        'subject': subject,
                        'attachments': message.Attachments.Count,
                        'attachment_names': [att.FileName for att in message.Attachments] if message.Attachments.Count > 0 else []
                    })
                    
            except Exception as e:
                continue
        
        print(f"📊 Found {len(emails_in_timeframe)} emails between 9:00-10:30 AM")
        
        if emails_in_timeframe:
            print("\n📧 ALL EMAILS IN 9:00-10:30 AM TIMEFRAME:")
            print("-" * 70)
            for i, email in enumerate(emails_in_timeframe, 1):
                time_str = email['time'].strftime('%I:%M:%S %p')
                print(f"{i}. {time_str}")
                print(f"   From: {email['sender']}")
                print(f"   Subject: {email['subject']}")
                print(f"   Attachments: {email['attachments']}")
                if email['attachment_names']:
                    print(f"   Files: {', '.join(email['attachment_names'])}")
                
                # Check if this could be the Van Paper email
                is_vanpaper = False
                reasons = []
                
                if 'vanpaper' in str(email['sender']).lower():
                    is_vanpaper = True
                    reasons.append("sender contains 'vanpaper'")
                
                if 'noreply@vanpaper.com' in str(email['sender']).lower():
                    is_vanpaper = True
                    reasons.append("sender is noreply@vanpaper.com")
                
                if 'leaderboard' in str(email['subject']).lower():
                    is_vanpaper = True
                    reasons.append("subject contains 'leaderboard'")
                
                if 'vanpaper' in str(email['subject']).lower():
                    is_vanpaper = True
                    reasons.append("subject contains 'vanpaper'")
                
                if email['attachments'] > 0:
                    excel_files = [f for f in email['attachment_names'] if f.lower().endswith(('.xlsx', '.xls', '.xlsm'))]
                    if excel_files:
                        reasons.append(f"has Excel attachments: {excel_files}")
                
                if is_vanpaper or reasons:
                    print(f"   🎯 POTENTIAL VAN PAPER EMAIL!")
                    print(f"   🔍 Reasons: {', '.join(reasons) if reasons else 'matches Van Paper patterns'}")
                
                print("-" * 50)
                
        else:
            print("\n❌ No emails found in the 9:00-10:30 AM timeframe")
            print("💡 The 9:55 AM email might be:")
            print("   - In a different time zone format")
            print("   - In Junk/Spam folder")
            print("   - In a different Outlook account/folder")
        
        # Also look for emails with 'leaderboard' in subject from today
        print(f"\n🔍 SEARCHING FOR 'LEADERBOARD' EMAILS FROM TODAY:")
        print("-" * 50)
        
        today_start = datetime.combine(today, datetime.min.time())
        leaderboard_emails = []
        
        for message in messages:
            try:
                if (hasattr(message, 'ReceivedTime') and 
                    message.ReceivedTime and
                    message.ReceivedTime >= today_start):
                    
                    subject = getattr(message, 'Subject', '')
                    if 'leaderboard' in str(subject).lower():
                        leaderboard_emails.append({
                            'time': message.ReceivedTime,
                            'sender': getattr(message, 'SenderEmailAddress', 'Unknown'),
                            'subject': subject,
                            'attachments': message.Attachments.Count
                        })
            except:
                continue
        
        if leaderboard_emails:
            for email in leaderboard_emails:
                time_str = email['time'].strftime('%I:%M:%S %p')
                print(f"📧 {time_str} - {email['subject']}")
                print(f"   From: {email['sender']}")
                print(f"   Attachments: {email['attachments']}")
                print()
        else:
            print("❌ No emails with 'leaderboard' in subject found today")
        
        return emails_in_timeframe
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return []

if __name__ == "__main__":
    emails = find_all_emails_around_955()
    
    print("\n" + "=" * 70)
    print("💡 This comprehensive search shows:")
    print("- Every email received between 9:00-10:30 AM")
    print("- Exact sender addresses and subjects")
    print("- All attachments")
    print("- Why the automation might not be finding your 9:55 AM email")
    
    input("\nPress Enter to continue...")
