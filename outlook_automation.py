#!/usr/bin/env python3
"""
Outlook Email Automation for Sales Leaderboard
Automatically downloads Excel attachments from Outlook and updates the leaderboard app.
Configured for Van Paper Company automated reports.
"""

import win32com.client
import os
import shutil
import subprocess
import datetime
import logging
import configparser
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('outlook_automation.log'),
        logging.StreamHandler()
    ]
)

def load_config():
    """Load configuration from automation_config.txt file."""
    config = configparser.ConfigParser()
    config_file = Path(__file__).parent / "automation_config.txt"
    
    # Default values
    defaults = {
        'SENDER_EMAIL': 'noreply@vanpaper.com',
        'SUBJECT_CONTAINS': 'Inform Auto Scheduled Report: leaderboardexport',
        'ATTACHMENT_NAME_CONTAINS': '',
        'DAYS_BACK': '3',
        'AUTO_UPDATE_GIT': 'True',
        'CREATE_BACKUPS': 'True'
    }
    
    if config_file.exists():
        try:
            config.read(config_file)
            return {
                'sender_email': config.get('EMAIL_SETTINGS', 'SENDER_EMAIL', fallback=defaults['SENDER_EMAIL']).strip(),
                'subject_contains': config.get('EMAIL_SETTINGS', 'SUBJECT_CONTAINS', fallback=defaults['SUBJECT_CONTAINS']).strip(),
                'attachment_name_contains': config.get('EMAIL_SETTINGS', 'ATTACHMENT_NAME_CONTAINS', fallback=defaults['ATTACHMENT_NAME_CONTAINS']).strip(),
                'days_back': int(config.get('EMAIL_SETTINGS', 'DAYS_BACK', fallback=defaults['DAYS_BACK'])),
                'auto_update_git': config.getboolean('AUTOMATION_SETTINGS', 'AUTO_UPDATE_GIT', fallback=True),
                'create_backups': config.getboolean('AUTOMATION_SETTINGS', 'CREATE_BACKUPS', fallback=True)
            }
        except Exception as e:
            logging.warning(f"Error reading config file, using defaults: {e}")
    
    # Return defaults if config file doesn't exist or has errors
    return {
        'sender_email': defaults['SENDER_EMAIL'],
        'subject_contains': defaults['SUBJECT_CONTAINS'],
        'attachment_name_contains': defaults['ATTACHMENT_NAME_CONTAINS'],
        'days_back': int(defaults['DAYS_BACK']),
        'auto_update_git': True,
        'create_backups': True
    }

class OutlookAutomation:
    def __init__(self, project_folder=None):
        """Initialize the Outlook automation."""
        self.project_folder = project_folder or Path(__file__).parent
        self.excel_filename = "leaderboard.xlsx"
        self.backup_folder = self.project_folder / "backups"
        self.backup_folder.mkdir(exist_ok=True)
        
    def connect_to_outlook(self):
        """Connect to Outlook application."""
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            logging.info("Successfully connected to Outlook")
            return outlook, namespace
        except Exception as e:
            logging.error(f"Failed to connect to Outlook: {e}")
            return None, None
    
    def search_for_emails(self, namespace, sender_email=None, subject_contains=None, days_back=1):
        """Search for emails with Excel attachments."""
        try:
            inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
            
            # Build search criteria
            search_criteria = []
            
            # Date filter - only look at recent emails
            date_filter = (datetime.datetime.now() - datetime.timedelta(days=days_back)).strftime("%m/%d/%Y")
            search_criteria.append(f"[ReceivedTime] >= '{date_filter}'")
            
            # Sender filter
            if sender_email:
                search_criteria.append(f"[SenderEmailAddress] = '{sender_email}'")
            
            # Subject filter
            if subject_contains:
                search_criteria.append(f"[Subject] LIKE '%{subject_contains}%'")
            
            # Combine criteria
            filter_string = " AND ".join(search_criteria) if search_criteria else None
            
            if filter_string:
                messages = inbox.Items.Restrict(filter_string)
                logging.info(f"Found {len(messages)} messages matching criteria")
            else:
                messages = inbox.Items
                logging.info("Searching all recent messages for Excel attachments")
            
            # Sort by received time (newest first)
            messages.Sort("[ReceivedTime]", True)
            
            return messages
            
        except Exception as e:
            logging.error(f"Error searching emails: {e}")
            return []
    
    def find_excel_attachments(self, messages, attachment_name_contains=None):
        """Find messages with Excel attachments."""
        excel_attachments = []
        
        for message in messages:
            try:
                if message.Attachments.Count > 0:
                    for attachment in message.Attachments:
                        filename = attachment.FileName.lower()
                        
                        # Check if it's an Excel file
                        if filename.endswith(('.xlsx', '.xls')):
                            # Optional: filter by filename pattern
                            if attachment_name_contains:
                                if attachment_name_contains.lower() not in filename:
                                    continue
                            
                            excel_attachments.append({
                                'message': message,
                                'attachment': attachment,
                                'filename': attachment.FileName,
                                'received_time': message.ReceivedTime,
                                'sender': message.SenderEmailAddress,
                                'subject': message.Subject
                            })
                            
                            logging.info(f"Found Excel attachment: {attachment.FileName} from {message.SenderEmailAddress}")
                            
            except Exception as e:
                logging.warning(f"Error processing message: {e}")
                continue
        
        # Sort by received time (newest first)
        excel_attachments.sort(key=lambda x: x['received_time'], reverse=True)
        return excel_attachments
    
    def backup_current_file(self):
        """Backup the current Excel file."""
        current_file = self.project_folder / self.excel_filename
        if current_file.exists():
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"leaderboard_backup_{timestamp}.xlsx"
            backup_path = self.backup_folder / backup_name
            
            try:
                shutil.copy2(current_file, backup_path)
                logging.info(f"Backed up current file to: {backup_name}")
                return True
            except Exception as e:
                logging.error(f"Failed to backup file: {e}")
                return False
        return True
    
    def download_attachment(self, attachment_info):
        """Download the Excel attachment."""
        try:
            # Backup current file first
            if not self.backup_current_file():
                logging.warning("Failed to backup current file, proceeding anyway...")
            
            # Download new file
            attachment = attachment_info['attachment']
            temp_path = self.project_folder / f"temp_{attachment_info['filename']}"
            final_path = self.project_folder / self.excel_filename
            
            # Save attachment
            attachment.SaveAsFile(str(temp_path))
            
            # Move to final location
            if temp_path.exists():
                if final_path.exists():
                    final_path.unlink()  # Remove old file
                shutil.move(str(temp_path), str(final_path))
                
                logging.info(f"Successfully downloaded and replaced {self.excel_filename}")
                logging.info(f"Source: {attachment_info['sender']} - {attachment_info['subject']}")
                logging.info(f"Received: {attachment_info['received_time']}")
                return True
            else:
                logging.error("Failed to save attachment")
                return False
                
        except Exception as e:
            logging.error(f"Error downloading attachment: {e}")
            return False
    
    def update_git_repo(self, commit_message=None):
        """Commit and push changes to git repository."""
        try:
            os.chdir(self.project_folder)
            
            # Add changes
            subprocess.run(["git", "add", "."], check=True, capture_output=True)
            
            # Create commit message
            if not commit_message:
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                commit_message = f"Auto-update leaderboard data - {timestamp}"
            
            # Commit changes
            result = subprocess.run(
                ["git", "commit", "-m", commit_message], 
                capture_output=True, text=True
            )
            
            if result.returncode == 0:
                # Push to remote
                subprocess.run(["git", "push"], check=True, capture_output=True)
                logging.info("Successfully updated git repository")
                logging.info("Live app will update automatically")
                return True
            else:
                if "nothing to commit" in result.stdout:
                    logging.info("No changes to commit")
                    return True
                else:
                    logging.error(f"Git commit failed: {result.stderr}")
                    return False
                    
        except subprocess.CalledProcessError as e:
            logging.error(f"Git operation failed: {e}")
            return False
        except Exception as e:
            logging.error(f"Error updating git repo: {e}")
            return False
    
    def run_automation(self, sender_email=None, subject_contains=None, 
                      attachment_name_contains=None, auto_update_git=True, days_back=1):
        """Run the complete automation process."""
        logging.info("Starting Outlook automation...")
        
        # Connect to Outlook
        outlook, namespace = self.connect_to_outlook()
        if not outlook:
            return False
        
        # Search for emails
        messages = self.search_for_emails(
            namespace, 
            sender_email=sender_email,
            subject_contains=subject_contains,
            days_back=days_back
        )
        
        if not messages:
            logging.info("No messages found matching criteria")
            return False
        
        # Find Excel attachments
        excel_attachments = self.find_excel_attachments(
            messages, 
            attachment_name_contains=attachment_name_contains
        )
        
        if not excel_attachments:
            logging.info("No Excel attachments found in matching emails")
            return False
        
        # Download the most recent attachment
        latest_attachment = excel_attachments[0]
        logging.info(f"Processing most recent attachment: {latest_attachment['filename']}")
        
        if self.download_attachment(latest_attachment):
            logging.info("✅ Excel file downloaded and replaced successfully")
            
            # Update git repository
            if auto_update_git:
                commit_msg = f"Auto-update leaderboard from Van Paper report - {latest_attachment['received_time'].strftime('%Y-%m-%d %H:%M')}"
                if self.update_git_repo(commit_msg):
                    logging.info("🚀 Git repository updated - live app will refresh automatically!")
                    return True
                else:
                    logging.warning("⚠️ File updated locally but git push failed")
                    return False
            else:
                logging.info("📁 File updated locally (git update skipped)")
                return True
        else:
            logging.error("❌ Failed to download attachment")
            return False


def main():
    """Main function to run the automation."""
    
    # Load configuration from file
    config = load_config()
    
    logging.info("Starting Van Paper Company leaderboard automation...")
    logging.info(f"Looking for emails from: {config['sender_email']}")
    logging.info(f"Subject contains: {config['subject_contains']}")
    logging.info(f"Searching last {config['days_back']} days")
    
    # Create automation instance
    automation = OutlookAutomation()
    
    # Run automation with loaded config
    success = automation.run_automation(
        sender_email=config['sender_email'] if config['sender_email'] else None,
        subject_contains=config['subject_contains'] if config['subject_contains'] else None,
        attachment_name_contains=config['attachment_name_contains'] if config['attachment_name_contains'] else None,
        auto_update_git=config['auto_update_git'],
        days_back=config['days_back']
    )
    
    if success:
        print("✅ Van Paper leaderboard automation completed successfully!")
        print("🚀 Your live app at vpsales.streamlit.app will update automatically!")
    else:
        print("❌ Automation failed. Check the outlook_automation.log file for details.")
    
    return success


if __name__ == "__main__":
    main()
