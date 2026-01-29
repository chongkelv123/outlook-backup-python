"""
Email Exporter Module
Handles exporting emails to .eml files and saving attachments
"""

import os
import re
from datetime import datetime
from typing import Tuple, Optional
import tempfile


class EmailExporter:
    """Exports Outlook emails to .eml format with attachments"""

    def __init__(self, backup_location: str, organize_by_date: bool = False,
                 include_attachments: bool = False):
        """
        Initialize the exporter

        Args:
            backup_location: Root directory for backups
            organize_by_date: Create YYYY/MM folder structure
            include_attachments: Save email attachments
        """
        self.backup_location = backup_location
        self.organize_by_date = organize_by_date
        self.include_attachments = include_attachments
        self.exported_count = 0
        self.total_size = 0
        self.errors = []

    def export_email(self, email, progress_callback=None) -> Tuple[bool, str]:
        """
        Export a single email to .eml file

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

            # Determine output directory
            output_dir = self._get_output_directory(dt)

            # Create directory if it doesn't exist
            os.makedirs(output_dir, exist_ok=True)

            # Generate filename
            filename = self._generate_filename(email, dt)
            filepath = os.path.join(output_dir, filename)

            # Handle filename collisions
            filepath = self._handle_collision(filepath)

            # Save email as .eml file
            # Use SaveAs method with olMSG format, then manually convert
            # Outlook doesn't directly support .eml export, so we use a workaround
            temp_msg = os.path.join(tempfile.gettempdir(), "temp_email.msg")

            try:
                # Save as MSG first
                email.SaveAs(temp_msg, 3)  # 3 = olMSG format

                # For .eml, we need to save in a different way
                # Best approach: use MIME format if available
                try:
                    # Try to save directly as .eml (MIME format)
                    # This might not work in all Outlook versions
                    email.SaveAs(filepath, 5)  # 5 = olMHTML
                except:
                    # Fallback: save as .msg and rename to .eml
                    # Or use the MSG file directly
                    msg_filepath = filepath.replace('.eml', '.msg')
                    email.SaveAs(msg_filepath, 3)  # 3 = olMSG
                    filepath = msg_filepath

                # Get file size
                file_size = os.path.getsize(filepath)
                self.total_size += file_size

                # Clean up temp file
                if os.path.exists(temp_msg):
                    os.remove(temp_msg)

            except Exception as e:
                # If SaveAs fails, try alternative method
                print(f"SaveAs failed, trying alternative: {str(e)}")
                # Save as MSG format as fallback
                msg_filepath = filepath.replace('.eml', '.msg')
                email.SaveAs(msg_filepath, 3)  # olMSG
                filepath = msg_filepath
                file_size = os.path.getsize(filepath)
                self.total_size += file_size

            # Handle attachments
            if self.include_attachments and email.Attachments.Count > 0:
                self._save_attachments(email, filepath)

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

    def _get_output_directory(self, email_date: datetime) -> str:
        """Determine output directory based on settings"""
        if self.organize_by_date:
            year_dir = str(email_date.year)
            month_dir = f"{email_date.month:02d}"
            return os.path.join(self.backup_location, year_dir, month_dir)
        else:
            return self.backup_location

    def _generate_filename(self, email, email_date: datetime) -> str:
        """
        Generate filename for email
        Format: email_YYYYMMDD_HHMMSS_[first20chars_of_subject].eml
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

        # Construct filename
        filename = f"email_{date_str}_{subject}.eml"

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

    def _save_attachments(self, email, email_filepath: str) -> None:
        """Save email attachments to subfolder"""
        try:
            # Create attachments folder
            base_name = os.path.splitext(email_filepath)[0]
            attachments_dir = f"{base_name}_attachments"
            os.makedirs(attachments_dir, exist_ok=True)

            # Save each attachment
            for attachment in email.Attachments:
                try:
                    filename = self._sanitize_filename(attachment.FileName)
                    attachment_path = os.path.join(attachments_dir, filename)

                    # Handle collisions
                    attachment_path = self._handle_collision(attachment_path)

                    # Save attachment
                    attachment.SaveAsFile(attachment_path)

                    # Update total size
                    self.total_size += os.path.getsize(attachment_path)

                except Exception as e:
                    self.errors.append(f"Failed to save attachment {attachment.FileName}: {str(e)}")

        except Exception as e:
            self.errors.append(f"Failed to process attachments: {str(e)}")

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
