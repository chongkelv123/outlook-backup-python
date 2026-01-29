"""
Outlook Connector Module
Handles connection to local Outlook application and email retrieval
"""

import win32com.client
import pywintypes
import time
from datetime import datetime
from typing import List, Optional


class OutlookConnector:
    """Manages connection to Outlook application and folder access"""

    def __init__(self):
        self.outlook = None
        self.namespace = None

    def connect(self, retry_count=3, wait_seconds=2) -> bool:
        """
        Connect to the local Outlook application with retry logic

        NOTE: This tool requires CLASSIC Outlook (Desktop version with COM support).
        The new Outlook (web-based) is NOT supported.

        Args:
            retry_count: Number of connection attempts
            wait_seconds: Seconds to wait between retries

        Returns True if successful, raises ConnectionError otherwise
        """
        last_error = None

        for attempt in range(retry_count):
            try:
                # Check if Outlook process is running
                if not self._is_outlook_process_running():
                    raise ConnectionError(
                        "Outlook is not running.\n\n"
                        "IMPORTANT: This tool requires CLASSIC Outlook.\n"
                        "The new Outlook (web-based version) is NOT supported.\n\n"
                        "Please start Classic Microsoft Outlook and try again.\n"
                        "If you're using new Outlook, switch to Classic Outlook using\n"
                        "the toggle switch in the top-right corner of Outlook."
                    )

                # Try different connection methods
                try:
                    # Method 1: GetActiveObject (connect to existing instance)
                    self.outlook = win32com.client.GetActiveObject("Outlook.Application")
                except:
                    # Method 2: Dispatch (create new or connect to existing)
                    self.outlook = win32com.client.Dispatch("Outlook.Application")

                # Get MAPI namespace
                self.namespace = self.outlook.GetNamespace("MAPI")

                # Try to access folders to verify connection
                try:
                    # This will fail if Outlook is not properly initialized
                    test_folder = self.namespace.GetDefaultFolder(6)  # Inbox
                    return True
                except:
                    raise ConnectionError(
                        "Connected to Outlook but cannot access folders.\n\n"
                        "This usually means:\n"
                        "1. You're using the new Outlook (not supported)\n"
                        "2. Outlook profile is not properly configured\n\n"
                        "SOLUTION: Switch to Classic Outlook using the toggle\n"
                        "switch in the top-right corner of Outlook."
                    )

            except pywintypes.com_error as e:
                error_code = e.args[0] if e.args else 0

                # Parse specific error codes
                if error_code == -2146959355:  # 0x80080005 - Server execution failed
                    last_error = ConnectionError(
                        "Outlook connection failed (Server execution failed).\n\n"
                        "IMPORTANT: This tool requires CLASSIC Outlook.\n"
                        "The new Outlook (web-based) does NOT support COM automation.\n\n"
                        "Solutions:\n"
                        "1. Switch to Classic Outlook using the toggle in top-right corner\n"
                        "2. Make sure Classic Outlook is running and fully loaded\n"
                        "3. Try running this application as Administrator\n"
                        "4. Close and restart Classic Outlook, then try again\n"
                        "5. Check if Outlook has a profile configured"
                    )
                elif error_code == -2147221005:  # 0x800401F3 - Invalid class string
                    last_error = ConnectionError(
                        "Outlook is not properly installed or registered.\n\n"
                        "Please reinstall Microsoft Office/Outlook."
                    )
                else:
                    last_error = ConnectionError(
                        f"COM Error connecting to Outlook (Code: {error_code}):\n{str(e)}\n\n"
                        "Try:\n"
                        "1. Start Outlook manually first\n"
                        "2. Run this application as Administrator\n"
                        "3. Restart your computer"
                    )

                # Wait before retry
                if attempt < retry_count - 1:
                    time.sleep(wait_seconds)

            except Exception as e:
                last_error = ConnectionError(
                    f"Unexpected error connecting to Outlook:\n{str(e)}\n\n"
                    "Make sure:\n"
                    "1. Outlook is installed and running\n"
                    "2. You have pywin32 installed (pip install pywin32)\n"
                    "3. Outlook has at least one email account configured"
                )

                if attempt < retry_count - 1:
                    time.sleep(wait_seconds)

        # All retries failed
        if last_error:
            raise last_error
        else:
            raise ConnectionError("Failed to connect to Outlook after multiple attempts")

    def _is_outlook_process_running(self) -> bool:
        """Check if Outlook process is running"""
        try:
            import subprocess
            # Use tasklist to check for OUTLOOK.EXE
            result = subprocess.run(
                ['tasklist', '/FI', 'IMAGENAME eq OUTLOOK.EXE'],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            return 'OUTLOOK.EXE' in result.stdout
        except:
            # If we can't check, assume it might be running
            return True

    def get_default_folders(self) -> dict:
        """
        Get commonly used Outlook folders
        Returns dict with folder names and objects
        """
        try:
            folders = {
                "Inbox": self.namespace.GetDefaultFolder(6),  # olFolderInbox
                "Sent Items": self.namespace.GetDefaultFolder(5),  # olFolderSentMail
                "Drafts": self.namespace.GetDefaultFolder(16),  # olFolderDrafts
                "Deleted Items": self.namespace.GetDefaultFolder(3),  # olFolderDeletedItems
                "Junk Email": self.namespace.GetDefaultFolder(23),  # olFolderJunk
            }
            return folders
        except Exception as e:
            raise Exception(f"Failed to retrieve default folders: {str(e)}")

    def get_folder_by_path(self, folder_path: str):
        """
        Get a folder by its path (e.g., "Inbox\\Subfolder")
        """
        try:
            folder = self.namespace.Folders.Item(1)
            for folder_name in folder_path.split("\\"):
                folder = folder.Folders[folder_name]
            return folder
        except Exception as e:
            raise Exception(f"Failed to access folder '{folder_path}': {str(e)}")

    def get_all_folder_names(self, folder=None, prefix="") -> List[str]:
        """
        Recursively get all folder names in the mailbox
        Returns list of folder paths
        """
        folder_names = []

        try:
            if folder is None:
                # Start with the root folders
                for account in self.namespace.Folders:
                    folder_names.extend(self._get_subfolders(account, account.Name))
            else:
                folder_names.extend(self._get_subfolders(folder, prefix))

            return folder_names
        except Exception as e:
            print(f"Warning: Could not retrieve all folders: {str(e)}")
            return folder_names

    def _get_subfolders(self, folder, prefix: str) -> List[str]:
        """Helper method to recursively get subfolder names"""
        folder_names = [prefix]

        try:
            for subfolder in folder.Folders:
                subfolder_path = f"{prefix}\\{subfolder.Name}"
                folder_names.extend(self._get_subfolders(subfolder, subfolder_path))
        except Exception:
            pass  # Some folders may not be accessible

        return folder_names

    def get_emails_from_folder(self, folder, filters: dict = None) -> List:
        """
        Retrieve emails from a folder with optional filters

        Args:
            folder: Outlook folder object
            filters: Dictionary with filter criteria (applied by filter_engine)

        Returns:
            List of email items
        """
        try:
            items = folder.Items
            # Sort by received time, newest first
            items.Sort("[ReceivedTime]", True)

            # Convert to list to allow iteration
            emails = []
            for item in items:
                # Only process MailItem objects
                if item.Class == 43:  # olMail
                    emails.append(item)

            return emails
        except Exception as e:
            raise Exception(f"Failed to retrieve emails from folder: {str(e)}")

    def get_email_count(self, folder) -> int:
        """Get the total number of emails in a folder"""
        try:
            count = 0
            for item in folder.Items:
                if item.Class == 43:  # olMail
                    count += 1
            return count
        except Exception as e:
            return 0

    def is_outlook_running(self) -> bool:
        """Check if Outlook application is running"""
        return self._is_outlook_process_running()

    def is_new_outlook_running(self) -> bool:
        """
        Check if new Outlook (web-based version) is running
        Returns True if new Outlook process is detected
        """
        try:
            import subprocess
            # Check for new Outlook process (olk.exe or HxOutlook.exe)
            result = subprocess.run(
                ['tasklist', '/FI', 'IMAGENAME eq olk.exe'],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            if 'olk.exe' in result.stdout:
                return True

            # Also check for HxOutlook.exe (another new Outlook process name)
            result = subprocess.run(
                ['tasklist', '/FI', 'IMAGENAME eq HxOutlook.exe'],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            return 'HxOutlook.exe' in result.stdout
        except:
            return False

    def get_outlook_version_info(self) -> dict:
        """
        Get information about which Outlook version is running
        Returns dict with version info and compatibility status
        """
        classic_running = self._is_outlook_process_running()
        new_running = self.is_new_outlook_running()

        return {
            'classic_outlook_running': classic_running,
            'new_outlook_running': new_running,
            'is_compatible': classic_running,
            'warning_message': self._get_compatibility_warning(classic_running, new_running)
        }

    def _get_compatibility_warning(self, classic_running: bool, new_running: bool) -> str:
        """Generate appropriate warning message based on detected Outlook versions"""
        if new_running and not classic_running:
            return (
                "⚠️ NEW OUTLOOK DETECTED\n\n"
                "You are running the new Outlook (web-based version).\n"
                "This tool requires Classic Outlook with COM automation support.\n\n"
                "HOW TO SWITCH:\n"
                "1. Open new Outlook\n"
                "2. Look for the toggle switch in the top-right corner\n"
                "3. Click 'Try the new Outlook' to turn it OFF\n"
                "4. Outlook will restart in Classic mode\n"
                "5. Run this backup tool again\n\n"
                "Classic Outlook is fully supported by Microsoft and\n"
                "provides full COM automation capabilities."
            )
        elif not classic_running and not new_running:
            return (
                "⚠️ OUTLOOK NOT DETECTED\n\n"
                "Please start Classic Microsoft Outlook.\n\n"
                "IMPORTANT: This tool requires Classic Outlook.\n"
                "The new Outlook (web-based) is NOT supported."
            )
        else:
            return ""
