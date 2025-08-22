#!/usr/bin/env python3
"""
Simple test script to debug Outlook connection and list recent emails.
"""

import win32com.client
import datetime

def test_outlook():
    print("Testing Outlook connection...")
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print("✅ Successfully connected to Outlook")
        
        # Check different folders
        folders_to_check = [
            (6, "Inbox"),
            (3, "Deleted Items"), 
            (4, "Outbox"),
            (5, "Sent Items"),
            (23, "Junk Email")
        ]
        
        for folder_id, folder_name in folders_to_check:
            try:
                print(f"\n📁 Checking {folder_name}...")
                folder = namespace.GetDefaultFolder(folder_id)
                messages = folder.Items
                
                if len(messages) > 0:
                    print(f"   Found {len(messages)} total messages")
                    messages.Sort("[ReceivedTime]", True)
                    
                    # Show last 5 emails
                    count = 0
                    cutoff_date = datetime.datetime.now() - datetime.timedelta(days=30)  # Look back 30 days
                    
                    for message in messages:
                        try:
                            if message.ReceivedTime >= cutoff_date and count < 5:
                                print(f"   {count + 1}. Subject: {message.Subject}")
                                print(f"      From: {message.SenderEmailAddress}")
                                print(f"      Date: {message.ReceivedTime}")
                                
                                # Check for Van Paper
                                if "vanpaper" in message.SenderEmailAddress.lower() or "vanpaper" in message.Subject.lower():
                                    print(f"      🎯 THIS IS A VAN PAPER EMAIL!")
                                
                                print()
                                count += 1
                            elif count >= 5:
                                break
                        except Exception as e:
                            continue
                else:
                    print(f"   No messages in {folder_name}")
                    
            except Exception as e:
                print(f"   Error accessing {folder_name}: {e}")
        
        # Try to find ANY Van Paper emails in any folder
        print("\n🔍 Searching ALL folders for Van Paper emails...")
        van_paper_found = False
        
        for folder_id, folder_name in folders_to_check:
            try:
                folder = namespace.GetDefaultFolder(folder_id)
                messages = folder.Items
                
                for message in messages:
                    try:
                        sender = getattr(message, 'SenderEmailAddress', '')
                        subject = getattr(message, 'Subject', '')
                        
                        if "vanpaper" in sender.lower() or "leaderboard" in subject.lower():
                            print(f"📨 Found Van Paper email in {folder_name}:")
                            print(f"   Subject: {subject}")
                            print(f"   From: {sender}")
                            print(f"   Date: {message.ReceivedTime}")
                            print(f"   Attachments: {message.Attachments.Count}")
                            van_paper_found = True
                            print()
                    except:
                        continue
                        
            except:
                continue
        
        if not van_paper_found:
            print("❌ No Van Paper emails found in any folder")
            print("\n💡 Suggestions:")
            print("1. Check if the emails are being sent to a different email account")
            print("2. Check if they're being filtered to a specific folder")
            print("3. Verify the sender email address is exactly 'noreply@vanpaper.com'")
            print("4. Try manually sending yourself a test email with an Excel attachment")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False
    
    return True

if __name__ == "__main__":
    test_outlook()
    input("\nPress Enter to exit...")
