"""
Debug Business Hours Automation
Show exactly what the automation is finding and why it might miss emails
"""

import win32com.client
from datetime import datetime, timedelta

def debug_business_hours_scan():
    """Debug version of the business hours scan"""
    
    print("🔍 DEBUG: Business Hours Van Paper Email Scan")
    print("=" * 50)
    print(f"🕐 Current time: {datetime.now().strftime('%I:%M %p CST')}")
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("✅ Connected to Outlook")
        
        # Look for emails from the last 3 hours
        cutoff_time = datetime.now() - timedelta(hours=3)
        
        print(f"📅 Looking for emails since {cutoff_time.strftime('%I:%M %p')}")
        print(f"📋 Cutoff timestamp: {cutoff_time}")
        
        # Get messages and sort by received time (newest first)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        van_paper_emails = []
        all_recent_emails = []
        
        print(f"\n🔎 Scanning messages...")
        
        count = 0
        for message in messages:
            count += 1
            if count > 100:  # Reasonable limit
                break
                
            try:
                # Skip if no received time
                if not hasattr(message, 'ReceivedTime') or not message.ReceivedTime:
                    continue
                
                # Skip if too old
                if message.ReceivedTime < cutoff_time:
                    print(f"⏰ Stopping scan - reached cutoff time at message {count}")
                    break  # Since sorted newest first, we can break here
                
                # Track all recent emails
                sender = getattr(message, 'SenderEmailAddress', 'Unknown')
                subject = getattr(message, 'Subject', 'No Subject')
                
                email_info = {
                    'time': message.ReceivedTime,
                    'sender': sender,
                    'subject': subject,
                    'attachments': message.Attachments.Count
                }
                
                all_recent_emails.append(email_info)
                
                # Check for Van Paper sender
                if 'noreply@vanpaper.com' not in str(sender).lower():
                    continue
                
                print(f"🎯 FOUND Van Paper email!")
                print(f"   📅 Time: {message.ReceivedTime.strftime('%I:%M:%S %p')}")
                print(f"   📧 Subject: {subject}")
                print(f"   👤 Sender: {sender}")
                print(f"   📎 Attachments: {message.Attachments.Count}")
                
                # Check for leaderboard export subject
                if 'leaderboardexport' not in str(subject).lower():
                    print(f"   ❌ SKIPPED: Subject doesn't contain 'leaderboardexport'")
                    continue
                
                # Check for Excel attachments
                if message.Attachments.Count == 0:
                    print(f"   ❌ SKIPPED: No attachments")
                    continue
                
                # Find Excel attachment
                excel_attachment = None
                for attachment in message.Attachments:
                    filename = attachment.FileName
                    print(f"   📄 Attachment: {filename}")
                    if filename.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                        excel_attachment = attachment
                        print(f"   ✅ VALID Excel attachment found!")
                        break
                
                if excel_attachment:
                    van_paper_emails.append({
                        'message': message,
                        'attachment': excel_attachment,
                        'received_time': message.ReceivedTime
                    })
                    print(f"   ✅ ADDED to processing list!")
                else:
                    print(f"   ❌ SKIPPED: No Excel attachments")
                    
            except Exception as e:
                print(f"   ⚠️ Error processing message: {e}")
                continue
        
        print(f"\n📊 SCAN SUMMARY:")
        print(f"   📧 Total recent emails scanned: {len(all_recent_emails)}")
        print(f"   🎯 Van Paper emails found: {len(van_paper_emails)}")
        
        if all_recent_emails:
            print(f"\n📧 ALL RECENT EMAILS (last 3 hours):")
            print("-" * 40)
            for i, email in enumerate(all_recent_emails[:10], 1):  # Show first 10
                time_str = email['time'].strftime('%I:%M:%S %p')
                is_vanpaper = 'noreply@vanpaper.com' in str(email['sender']).lower()
                print(f"{i}. {time_str} {'🎯' if is_vanpaper else '📧'}")
                print(f"   {email['subject'][:50]}...")
                print(f"   From: {email['sender']}")
                if email['attachments'] > 0:
                    print(f"   📎 {email['attachments']} attachments")
                print()
                
            if len(all_recent_emails) > 10:
                print(f"   ... and {len(all_recent_emails) - 10} more emails")
        
        if van_paper_emails:
            print(f"\n🎯 VAN PAPER EMAILS READY FOR PROCESSING:")
            print("-" * 45)
            for i, email in enumerate(van_paper_emails, 1):
                time_str = email['received_time'].strftime('%I:%M:%S %p')
                print(f"{i}. {time_str}")
                print(f"   Subject: {email['message'].Subject}")
                print(f"   Excel: {email['attachment'].FileName}")
                print()
                
            # Return the most recent one
            latest_email = van_paper_emails[0]  # Already sorted newest first
            print(f"✅ WOULD PROCESS: {latest_email['received_time'].strftime('%I:%M %p')} email")
            return latest_email
        else:
            print(f"\n❌ NO VAN PAPER EMAILS FOUND")
            print(f"💡 Automation would exit quietly (normal behavior)")
            return None
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    print("🔧 DEBUGGING BUSINESS HOURS AUTOMATION")
    print("This will show exactly what the automation sees")
    print()
    
    result = debug_business_hours_scan()
    
    print("\n" + "=" * 60)
    if result:
        print("✅ DEBUG: Van Paper email would be processed")
    else:
        print("❌ DEBUG: No Van Paper emails would be processed")
        print("This explains why your 11:16 AM email was missed!")
    
    input("\nPress Enter to continue...")
