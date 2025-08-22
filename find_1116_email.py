"""
Find 11:16 AM Van Paper Email
Look specifically for the email you're seeing at 11:16 AM
"""

import win32com.client
from datetime import datetime, timedelta

def find_1116_email():
    """Find the specific 11:16 AM Van Paper email"""
    
    print("🔍 Looking for 11:16 AM Van Paper Email")
    print("=" * 40)
    print(f"🕐 Current time: {datetime.now().strftime('%Y-%m-%d %I:%M %p CST')}")
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("✅ Connected to Outlook")
        
        # Look for emails around 11:16 AM today
        today = datetime.now().date()
        
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        emails_found = []
        
        print(f"🔎 Scanning for emails from 11:00-11:30 AM today...")
        
        for message in messages:
            try:
                if not hasattr(message, 'ReceivedTime') or not message.ReceivedTime:
                    continue
                
                # Check if from today and in the 11:00-11:30 AM window
                if (message.ReceivedTime.date() == today and
                    message.ReceivedTime.hour == 11 and
                    0 <= message.ReceivedTime.minute <= 30):
                    
                    sender = getattr(message, 'SenderEmailAddress', 'Unknown')
                    subject = getattr(message, 'Subject', 'No Subject')
                    
                    email_info = {
                        'time': message.ReceivedTime,
                        'sender': sender,
                        'subject': subject,
                        'attachments': message.Attachments.Count,
                        'attachment_names': [att.FileName for att in message.Attachments] if message.Attachments.Count > 0 else []
                    }
                    
                    emails_found.append(email_info)
                    
                    # Check if this is Van Paper related
                    sender_lower = str(sender).lower()
                    subject_lower = str(subject).lower()
                    
                    is_vanpaper = ('vanpaper' in sender_lower or 
                                  'noreply@vanpaper.com' in sender_lower or
                                  'leaderboard' in subject_lower)
                    
                    time_str = message.ReceivedTime.strftime('%I:%M:%S %p')
                    print(f"📧 {time_str} - {'🎯 VAN PAPER! ' if is_vanpaper else ''}")
                    print(f"   Subject: {subject}")
                    print(f"   From: {sender}")
                    print(f"   Attachments: {message.Attachments.Count}")
                    if message.Attachments.Count > 0:
                        print(f"   Files: {', '.join([att.FileName for att in message.Attachments])}")
                    print()
                
            except Exception as e:
                continue
        
        if not emails_found:
            print("❌ No emails found in the 11:00-11:30 AM timeframe")
            print("💡 Let me check a wider time range...")
            
            # Check 10:30-11:45 AM
            print(f"\n🔍 Expanding search to 10:30 AM - 11:45 AM...")
            
            for message in messages:
                try:
                    if not hasattr(message, 'ReceivedTime') or not message.ReceivedTime:
                        continue
                    
                    msg_time = message.ReceivedTime
                    if (msg_time.date() == today and
                        ((msg_time.hour == 10 and msg_time.minute >= 30) or
                         (msg_time.hour == 11 and msg_time.minute <= 45))):
                        
                        sender = getattr(message, 'SenderEmailAddress', 'Unknown')
                        subject = getattr(message, 'Subject', 'No Subject')
                        
                        time_str = msg_time.strftime('%I:%M:%S %p')
                        print(f"📧 {time_str}")
                        print(f"   Subject: {subject[:50]}...")
                        print(f"   From: {sender}")
                        print()
                        
                except Exception as e:
                    continue
        
        return emails_found
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return []

if __name__ == "__main__":
    emails = find_1116_email()
    
    print("\n" + "=" * 50)
    if emails:
        van_paper_emails = [e for e in emails if 'vanpaper' in str(e['sender']).lower() or 'leaderboard' in str(e['subject']).lower()]
        if van_paper_emails:
            print(f"✅ Found {len(van_paper_emails)} Van Paper emails!")
        else:
            print(f"📧 Found {len(emails)} emails but none from Van Paper")
    else:
        print("❌ No emails found in search window")
    
    input("\nPress Enter to continue...")
