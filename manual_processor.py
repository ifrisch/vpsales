"""
Manual Excel Processor
Use this when you manually download an Excel file and want to process it
"""

import pandas as pd
import subprocess
import shutil
import os
from datetime import datetime

def process_manual_excel():
    """Process a manually downloaded Excel file"""
    
    print("📂 Manual Excel File Processor")
    print("-" * 40)
    
    # Look for Excel files in the directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_files = []
    
    for file in os.listdir(current_dir):
        if file.lower().endswith(('.xlsx', '.xls', '.xlsm')):
            if file != 'leaderboard.xlsx':  # Don't process the main file
                excel_files.append(file)
    
    if not excel_files:
        print("❌ No Excel files found in the directory")
        print("💡 Place your downloaded Excel file in this folder and run again")
        return False
    
    print(f"📊 Found Excel files: {excel_files}")
    
    # Use the first (or most recent) Excel file
    if len(excel_files) == 1:
        source_file = excel_files[0]
    else:
        print("\nMultiple Excel files found:")
        for i, file in enumerate(excel_files, 1):
            print(f"{i}. {file}")
        
        try:
            choice = int(input("Enter the number of the file to process: ")) - 1
            source_file = excel_files[choice]
        except (ValueError, IndexError):
            print("❌ Invalid choice")
            return False
    
    print(f"📂 Processing: {source_file}")
    
    try:
        # Create backup of existing leaderboard.xlsx
        backup_name = f"leaderboard_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        if os.path.exists("leaderboard.xlsx"):
            shutil.copy2("leaderboard.xlsx", backup_name)
            print(f"💾 Created backup: {backup_name}")
        
        # Copy new file to leaderboard.xlsx
        shutil.copy2(source_file, "leaderboard.xlsx")
        print(f"✅ Updated leaderboard.xlsx with data from {source_file}")
        
        # Test if we can read the file
        df = pd.read_excel("leaderboard.xlsx")
        print(f"📊 Verified: Excel file has {len(df)} rows and {len(df.columns)} columns")
        print(f"📋 Columns: {list(df.columns)}")
        
        # Offer to update git and live app
        update_git = input("\n🚀 Update the live Streamlit app? (y/n): ").lower().strip()
        
        if update_git == 'y':
            print("\n📤 Updating live app...")
            
            # Git operations
            subprocess.run(['git', 'add', 'leaderboard.xlsx'], cwd=current_dir, capture_output=True)
            
            commit_message = f"Update leaderboard data from {source_file} - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            result = subprocess.run(['git', 'commit', '-m', commit_message], 
                                  cwd=current_dir, capture_output=True, text=True)
            
            if result.returncode == 0:
                print("✅ Changes committed to git")
                
                push_result = subprocess.run(['git', 'push'], cwd=current_dir, capture_output=True, text=True)
                if push_result.returncode == 0:
                    print("🚀 Changes pushed! Live app will update in ~1 minute")
                    print("🌐 Check: https://vpsales.streamlit.app/")
                else:
                    print(f"❌ Git push failed: {push_result.stderr}")
            else:
                print(f"ℹ️ Git commit result: {result.stderr}")
        
        # Clean up - ask if they want to delete the processed file
        cleanup = input(f"\n🗑️ Delete the processed file '{source_file}'? (y/n): ").lower().strip()
        if cleanup == 'y':
            os.remove(source_file)
            print(f"🗑️ Deleted {source_file}")
        
        print("\n✅ Manual processing complete!")
        return True
        
    except Exception as e:
        print(f"❌ Error processing file: {e}")
        return False

if __name__ == "__main__":
    process_manual_excel()
    print("\n" + "="*50)
    print("💡 TIP: This tool is perfect for testing until Van Paper")
    print("    sets up their automated email reports!")
