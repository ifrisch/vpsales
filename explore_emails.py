"""
Explore and Find Emails Tool
This script helps you find emails with Excel attachments in your Outlook
"""

import win32com.client
from datetime import datetime, timedelta
import os

def explore_all_emails():
    """Search through all emails to find ANY with Excel attachments"""
    
    print("🔍 Exploring your Outlook for Excel attachments...")
    print("-" * 50)
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
        
        print(f"✅ Connected to Outlook inbox")
        print(f"📧 Total emails in inbox: {inbox.Items.Count}")
        
        # Get emails from last 30 days
        cutoff_date = datetime.now() - timedelta(days=30)
        
        emails_with_attachments = []
        excel_emails = []
        
        print("\n🔎 Scanning recent emails...")
        
        # Sort by received time, newest first
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        count = 0
        for message in messages:
            count += 1
            
            # Only check last 100 emails for speed
            if count > 100:
                break
                
            try:
                # Check if message has received time and is recent enough
                if hasattr(message, 'ReceivedTime') and message.ReceivedTime:
                    if message.ReceivedTime < cutoff_date:
                        continue
                
                # Check for attachments
                if message.Attachments.Count > 0:
                    emails_with_attachments.append({
                        'subject': message.Subject,
                        'sender': message.SenderEmailAddress if hasattr(message, 'SenderEmailAddress') else 'Unknown',
                        'received': message.ReceivedTime if hasattr(message, 'ReceivedTime') else 'Unknown',
                        'attachments': []
                    })
                    
                    # Check each attachment
                    for attachment in message.Attachments:
                        filename = attachment.FileName
                        emails_with_attachments[-1]['attachments'].append(filename)
                        
                        # Check if it's an Excel file
                        if filename.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                            excel_emails.append(emails_with_attachments[-1])
                            break
                            
            except Exception as e:
                # Skip problematic emails
                continue
        
        # Report findings
        print(f"\n📊 RESULTS:")
        print(f"🔍 Checked {count} recent emails")
        print(f"📎 Found {len(emails_with_attachments)} emails with attachments")
        print(f"📊 Found {len(excel_emails)} emails with Excel attachments")
        
        if emails_with_attachments:
            print("\n📎 ALL EMAILS WITH ATTACHMENTS:")
            print("-" * 50)
            for i, email in enumerate(emails_with_attachments, 1):
                print(f"{i}. Subject: {email['subject']}")
                print(f"   From: {email['sender']}")
                print(f"   Date: {email['received']}")
                print(f"   Attachments: {', '.join(email['attachments'])}")
                print()
        
        if excel_emails:
            print("\n🎯 EMAILS WITH EXCEL ATTACHMENTS:")
            print("-" * 50)
            for i, email in enumerate(excel_emails, 1):
                print(f"{i}. Subject: {email['subject']}")
                print(f"   From: {email['sender']}")
                print(f"   Date: {email['received']}")
                print(f"   Excel files: {[f for f in email['attachments'] if f.lower().endswith(('.xlsx', '.xls', '.xlsm'))]}")
                print()
        else:
            print("\n❌ No Excel attachments found in recent emails")
            print("\nPossible reasons:")
            print("1. Van Paper hasn't set up automated reports yet")
            print("2. Reports are sent to a different email address")
            print("3. Reports are in a different Outlook folder (Spam/Junk)")
            print("4. Reports use a different file format")
            print("5. Reports haven't been sent in the last 30 days")
        
        return excel_emails
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return []

if __name__ == "__main__":
    explore_all_emails()
    print("\n" + "="*50)
    print("💡 TIP: If you find the right emails, update automation_config.txt")
    print("    with the correct sender and subject information!")
