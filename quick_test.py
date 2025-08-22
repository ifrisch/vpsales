#!/usr/bin/env python3
"""
Quick test - just find any recent Excel attachments in Outlook
"""

import win32com.client
import datetime

def quick_test():
    print("🔍 Quick test: Finding ANY Excel attachments in recent emails...")
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print("✅ Connected to Outlook")
        
        # Get inbox
        inbox = namespace.GetDefaultFolder(6)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        cutoff_date = datetime.datetime.now() - datetime.timedelta(days=7)
        print(f"Looking for emails since {cutoff_date.strftime('%Y-%m-%d')}")
        
        excel_found = 0
        emails_checked = 0
        
        for message in messages:
            try:
                emails_checked += 1
                
                # Stop after checking 50 emails to keep it fast
                if emails_checked > 50:
                    break
                    
                # Skip old emails
                if message.ReceivedTime < cutoff_date:
                    break
                
                # Check for Excel attachments
                if message.Attachments.Count > 0:
                    for attachment in message.Attachments:
                        filename = attachment.FileName.lower()
                        if filename.endswith(('.xlsx', '.xls')):
                            excel_found += 1
                            print(f"📊 Excel file found!")
                            print(f"   Email: {message.Subject}")
                            print(f"   From: {message.SenderEmailAddress}")
                            print(f"   File: {attachment.FileName}")
                            print(f"   Date: {message.ReceivedTime}")
                            print()
                            
                            # Check if this looks like a Van Paper email
                            if "vanpaper" in message.SenderEmailAddress.lower() or "leaderboard" in filename:
                                print("   🎯 THIS LOOKS LIKE YOUR VAN PAPER FILE!")
                                print()
                
            except Exception as e:
                continue
        
        print(f"✅ Checked {emails_checked} recent emails")
        print(f"📊 Found {excel_found} Excel attachments")
        
        if excel_found == 0:
            print("\n💡 No Excel files found. Possible reasons:")
            print("1. No recent emails with Excel attachments")
            print("2. Van Paper emails might be going to a different folder")
            print("3. The automation reports might not be set up yet")
            
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    quick_test()
    input("\nPress Enter to exit...")
