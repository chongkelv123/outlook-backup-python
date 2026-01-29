"""
Filter Engine Module
Handles email filtering logic based on various criteria
"""

from datetime import datetime
from typing import List, Optional


class FilterEngine:
    """Applies filters to email lists"""

    @staticmethod
    def apply_filters(emails: List, filters: dict) -> List:
        """
        Apply multiple filters to email list

        Args:
            emails: List of Outlook mail items
            filters: Dictionary containing filter criteria:
                - date_from: datetime object (start date)
                - date_to: datetime object (end date)
                - sender: str (email address to filter)
                - subject: str (keyword to search in subject)

        Returns:
            Filtered list of emails
        """
        filtered_emails = emails

        # Apply date range filter
        if filters.get('date_from') or filters.get('date_to'):
            filtered_emails = FilterEngine._filter_by_date_range(
                filtered_emails,
                filters.get('date_from'),
                filters.get('date_to')
            )

        # Apply sender filter
        if filters.get('sender'):
            filtered_emails = FilterEngine._filter_by_sender(
                filtered_emails,
                filters.get('sender')
            )

        # Apply subject filter
        if filters.get('subject'):
            filtered_emails = FilterEngine._filter_by_subject(
                filtered_emails,
                filters.get('subject')
            )

        return filtered_emails

    @staticmethod
    def _filter_by_date_range(emails: List, date_from: Optional[datetime],
                               date_to: Optional[datetime]) -> List:
        """Filter emails by date range"""
        filtered = []

        for email in emails:
            try:
                # Get received time (handle both ReceivedTime and CreationTime)
                try:
                    email_date = email.ReceivedTime
                except AttributeError:
                    email_date = email.CreationTime

                # Convert to datetime if needed (COM date format)
                if hasattr(email_date, 'year'):
                    email_datetime = email_date
                else:
                    # Handle pywintypes.datetime
                    email_datetime = datetime(
                        email_date.year,
                        email_date.month,
                        email_date.day,
                        email_date.hour,
                        email_date.minute,
                        email_date.second
                    )

                # Remove timezone info for comparison
                email_datetime = email_datetime.replace(tzinfo=None)

                # Apply date filters
                if date_from and email_datetime < date_from:
                    continue
                if date_to and email_datetime > date_to:
                    continue

                filtered.append(email)
            except Exception as e:
                print(f"Warning: Could not process email date: {str(e)}")
                # Include emails with date processing errors
                filtered.append(email)

        return filtered

    @staticmethod
    def _filter_by_sender(emails: List, sender: str) -> List:
        """Filter emails by sender address"""
        filtered = []
        sender_lower = sender.lower().strip()

        for email in emails:
            try:
                # Get sender email address
                email_sender = ""

                # Check if it's an Exchange email type
                if hasattr(email, 'SenderEmailType') and email.SenderEmailType == "EX":
                    # For Exchange emails, get SMTP address from Exchange User
                    try:
                        if hasattr(email, 'Sender') and email.Sender:
                            exchange_user = email.Sender.GetExchangeUser()
                            if exchange_user and hasattr(exchange_user, 'PrimarySmtpAddress'):
                                email_sender = exchange_user.PrimarySmtpAddress
                    except:
                        pass

                # If still empty, try standard properties
                if not email_sender:
                    if hasattr(email, 'SenderEmailAddress'):
                        email_sender = email.SenderEmailAddress
                    elif hasattr(email, 'Sender') and email.Sender:
                        if hasattr(email.Sender, 'Address'):
                            email_sender = email.Sender.Address

                # Also check SenderName
                sender_name = ""
                if hasattr(email, 'SenderName'):
                    sender_name = email.SenderName

                # Match if sender string is found in either email or name
                if (sender_lower in email_sender.lower() or
                        sender_lower in sender_name.lower()):
                    filtered.append(email)
            except Exception as e:
                print(f"Warning: Could not process email sender: {str(e)}")
                continue

        return filtered

    @staticmethod
    def _filter_by_subject(emails: List, subject_keyword: str) -> List:
        """Filter emails by subject keyword (case-insensitive contains)"""
        filtered = []
        keyword_lower = subject_keyword.lower().strip()

        for email in emails:
            try:
                email_subject = email.Subject if hasattr(email, 'Subject') else ""
                if email_subject and keyword_lower in email_subject.lower():
                    filtered.append(email)
            except Exception as e:
                print(f"Warning: Could not process email subject: {str(e)}")
                continue

        return filtered

    @staticmethod
    def get_filter_summary(filters: dict) -> str:
        """
        Generate a human-readable summary of active filters

        Args:
            filters: Dictionary containing filter criteria

        Returns:
            String description of active filters
        """
        summary_parts = []

        if filters.get('date_from') or filters.get('date_to'):
            date_from_str = filters['date_from'].strftime('%Y-%m-%d') if filters.get('date_from') else 'Any'
            date_to_str = filters['date_to'].strftime('%Y-%m-%d') if filters.get('date_to') else 'Any'
            summary_parts.append(f"Date: {date_from_str} to {date_to_str}")

        if filters.get('sender'):
            summary_parts.append(f"Sender: {filters['sender']}")

        if filters.get('subject'):
            summary_parts.append(f"Subject contains: '{filters['subject']}'")

        if not summary_parts:
            return "No filters applied"

        return " | ".join(summary_parts)
