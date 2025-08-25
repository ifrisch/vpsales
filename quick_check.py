"""
Quick Van Paper Check - Simple and Fast
"""

import win32com.client
from datetime import datetime

def quick_vanpaper_check():
    try:
        print("Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("Checking for Van Paper emails today...")
        today = datetime.now().date()
        
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        vanpaper_count = 0
        found_emails = []
        
        # Only check first 50 messages to avoid hanging
        count = 0
        for message in messages:
            count += 1
            if count > 50:
                break
                
            try:
                if not hasattr(message, 'ReceivedTime'):
                    continue
                    
                # Only check today's emails
                if message.ReceivedTime.date() != today:
                    continue
                    
                sender = str(getattr(message, 'SenderEmailAddress', ''))
                subject = str(getattr(message, 'Subject', ''))
                
                if 'noreply@vanpaper.com' in sender.lower() and 'leaderboard' in subject.lower():
                    vanpaper_count += 1
                    time_str = message.ReceivedTime.strftime('%I:%M:%S %p')
                    found_emails.append(f"{time_str} - {subject}")
                    print(f"FOUND: {time_str} - {subject}")
                    
            except Exception:
                continue
        
        print(f"\nTotal Van Paper emails today: {vanpaper_count}")
        return found_emails
        
    except Exception as e:
        print(f"Error: {e}")
        return []

if __name__ == "__main__":
    emails = quick_vanpaper_check()
    print(f"\nDone! Found {len(emails)} Van Paper emails.")
