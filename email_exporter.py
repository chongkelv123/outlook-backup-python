"""
Email Exporter Module
Handles exporting emails to .msg files (Outlook native format) with flexible folder organization
"""

import os
import re
from datetime import datetime
from typing import Tuple, Optional


class EmailExporter:
    """Exports Outlook emails to .msg format with flexible folder organization"""

    def __init__(self, backup_location: str, organize_by_date: bool = False,
                 include_attachments: bool = True, sender_filter_enabled: bool = False,
                 subject_filter_enabled: bool = False):
        """
        Initialize the exporter

        Args:
            backup_location: Root directory for backups
            organize_by_date: Create YYYY/MM folder structure
            include_attachments: Save email attachments (Note: .msg format embeds attachments automatically)
            sender_filter_enabled: Organize emails by sender email address
            subject_filter_enabled: Organize emails by subject
        """
        self.backup_location = backup_location
        self.organize_by_date = organize_by_date
        self.include_attachments = include_attachments
        self.sender_filter_enabled = sender_filter_enabled
        self.subject_filter_enabled = subject_filter_enabled
        self.exported_count = 0
        self.total_size = 0
        self.errors = []

    def export_email(self, email, progress_callback=None) -> Tuple[bool, str]:
        """
        Export a single email to .msg file

        Args:
            email: Outlook mail item
            progress_callback: Optional callback function for progress updates

        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            # Get email date
            try:
                email_date = email.ReceivedTime
            except AttributeError:
                email_date = email.CreationTime

            # Convert to datetime
            if hasattr(email_date, 'year'):
                dt = email_date
            else:
                dt = datetime(email_date.year, email_date.month, email_date.day,
                              email_date.hour, email_date.minute, email_date.second)

            # Determine output directory based on priority logic
            output_dir = self._get_output_directory(email, dt)

            # Create directory if it doesn't exist
            os.makedirs(output_dir, exist_ok=True)

            # Generate filename
            filename = self._generate_filename(email, dt)
            filepath = os.path.join(output_dir, filename)

            # Handle filename collisions
            filepath = self._handle_collision(filepath)

            # Save email as .msg file (Outlook native format)
            # olMSG = 3 (preserves ALL Outlook-specific data: categories, flags, importance, etc.)
            email.SaveAs(filepath, 3)

            # Get file size
            file_size = os.path.getsize(filepath)
            self.total_size += file_size

            # Note: .msg format already embeds attachments, so no separate extraction needed
            # If include_attachments is False and we want to strip attachments, we would need
            # additional logic, but typically .msg with embedded attachments is preferred

            self.exported_count += 1

            if progress_callback:
                progress_callback(f"Exported: {os.path.basename(filepath)}")

            return True, filepath

        except Exception as e:
            error_msg = f"Failed to export email: {str(e)}"
            self.errors.append(error_msg)
            if progress_callback:
                progress_callback(f"Error: {error_msg}")
            return False, error_msg

    def _get_output_directory(self, email, email_date: datetime) -> str:
        """
        Determine output directory based on organizational settings

        Priority Logic:
        1. If sender_filter_enabled → organize by sender email first
        2. Else if subject_filter_enabled → organize by subject first
        3. Then apply date organization if organize_by_date is enabled

        Examples:
        - Sender only: BackupFolder/sender@email.com/
        - Sender + Date: BackupFolder/sender@email.com/2025/01/
        - Subject only: BackupFolder/Subject_Name/
        - Subject + Date: BackupFolder/Subject_Name/2025/01/
        - Date only: BackupFolder/2025/01/
        """
        base_dir = self.backup_location

        # Priority 1: Sender organization
        if self.sender_filter_enabled:
            sender_email = self._extract_sender_email(email)
            sender_folder = self._sanitize_folder_name(sender_email)
            base_dir = os.path.join(base_dir, sender_folder)

        # Priority 2: Subject organization (only if sender not enabled)
        elif self.subject_filter_enabled:
            subject = self._extract_subject(email)
            subject_folder = self._sanitize_subject_folder(subject)
            base_dir = os.path.join(base_dir, subject_folder)

        # Apply date organization on top of sender/subject organization
        if self.organize_by_date:
            year_dir = str(email_date.year)
            month_dir = f"{email_date.month:02d}"
            base_dir = os.path.join(base_dir, year_dir, month_dir)

        return base_dir

    def _extract_sender_email(self, email) -> str:
        """
        Extract sender's SMTP email address
        Handles both Exchange and SMTP email types
        """
        try:
            sender_email = ""

            # Check if it's an Exchange email type
            if hasattr(email, 'SenderEmailType') and email.SenderEmailType == "EX":
                # For Exchange emails, get SMTP address from Exchange User
                try:
                    if hasattr(email, 'Sender') and email.Sender:
                        exchange_user = email.Sender.GetExchangeUser()
                        if exchange_user and hasattr(exchange_user, 'PrimarySmtpAddress'):
                            sender_email = exchange_user.PrimarySmtpAddress
                except:
                    pass

            # If still empty, try standard properties
            if not sender_email:
                if hasattr(email, 'SenderEmailAddress'):
                    sender_email = email.SenderEmailAddress
                elif hasattr(email, 'Sender') and email.Sender:
                    if hasattr(email.Sender, 'Address'):
                        sender_email = email.Sender.Address

            # If still empty, try SenderName
            if not sender_email and hasattr(email, 'SenderName'):
                sender_email = email.SenderName

            # Default if all else fails
            if not sender_email:
                sender_email = "unknown_sender"

            return sender_email

        except Exception as e:
            return "unknown_sender"

    def _extract_subject(self, email) -> str:
        """Extract email subject"""
        try:
            subject = email.Subject if hasattr(email, 'Subject') and email.Subject else ""
            return subject.strip() if subject else ""
        except:
            return ""

    def _sanitize_folder_name(self, folder_name: str) -> str:
        """
        Sanitize folder name by removing invalid filesystem characters
        Used for email addresses: removes < > and other invalid chars
        """
        # Replace invalid characters with underscore
        invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
        sanitized = re.sub(invalid_chars, '_', folder_name)

        # Remove leading/trailing spaces and dots
        sanitized = sanitized.strip('. ')

        # Ensure not empty
        if not sanitized:
            sanitized = "unknown"

        return sanitized

    def _sanitize_subject_folder(self, subject: str) -> str:
        """
        Sanitize subject for use as folder name
        - Removes special chars: / \\ : * ? " < > |
        - Truncates to max 50 characters
        - Handles empty subjects
        """
        if not subject:
            return "No_Subject"

        # Replace invalid characters with underscore
        invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
        sanitized = re.sub(invalid_chars, '_', subject)

        # Replace multiple underscores with single underscore
        sanitized = re.sub(r'_+', '_', sanitized)

        # Remove leading/trailing spaces, dots, and underscores
        sanitized = sanitized.strip('. _')

        # Truncate to 50 characters
        if len(sanitized) > 50:
            sanitized = sanitized[:50].rstrip('_')

        # Ensure not empty after sanitization
        if not sanitized:
            sanitized = "No_Subject"

        return sanitized

    def _generate_filename(self, email, email_date: datetime) -> str:
        """
        Generate filename for email
        Format: email_YYYYMMDD_HHMMSS_[first20chars_of_subject].msg
        """
        # Get date/time string
        date_str = email_date.strftime("%Y%m%d_%H%M%S")

        # Get subject and sanitize
        try:
            subject = email.Subject if email.Subject else "no_subject"
        except:
            subject = "no_subject"

        # Take first 20 characters and sanitize
        subject = subject[:20]
        subject = self._sanitize_filename(subject)

        # Construct filename with .msg extension
        filename = f"email_{date_str}_{subject}.msg"

        return filename

    def _sanitize_filename(self, filename: str) -> str:
        """Remove invalid characters from filename"""
        # Replace invalid characters with underscore
        invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
        sanitized = re.sub(invalid_chars, '_', filename)

        # Remove leading/trailing spaces and dots
        sanitized = sanitized.strip('. ')

        # Ensure not empty
        if not sanitized:
            sanitized = "unnamed"

        return sanitized

    def _handle_collision(self, filepath: str) -> str:
        """Handle filename collisions by appending numbers"""
        if not os.path.exists(filepath):
            return filepath

        base, ext = os.path.splitext(filepath)
        counter = 1

        while os.path.exists(f"{base}_{counter}{ext}"):
            counter += 1

        return f"{base}_{counter}{ext}"

    def get_summary(self) -> dict:
        """Get export summary statistics"""
        return {
            'exported_count': self.exported_count,
            'total_size': self.total_size,
            'total_size_mb': round(self.total_size / (1024 * 1024), 2),
            'errors': self.errors,
            'error_count': len(self.errors),
            'backup_location': self.backup_location
        }

    def reset_stats(self):
        """Reset statistics for new export session"""
        self.exported_count = 0
        self.total_size = 0
        self.errors = []
