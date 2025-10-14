"""Concrete filter implementations for email filtering.

This module contains specific filter implementations for common filtering
scenarios like date ranges, senders, and keyword matching.
"""

from datetime import datetime
from typing import List

from outlook_exporter.core.models import EmailMetadata
from outlook_exporter.filters.base import EmailFilter

class DateRangeFilter(EmailFilter):
    """Filter emails by received date range.
    
    Attributes:
        start_date: Minimum received date (inclusive), None for no lower bound
        end_date: Maximum received date (inclusive), None for no upper bound
    """
    
    def __init__(self, start_date: datetime | None = None, end_date: datetime | None = None):
        """Initialize the date range filter.
        
        Args:
            start_date: Filter emails received on or after this date
            end_date: Filter emails received on or before this date
        """
        self.start_date = start_date
        self.end_date = end_date
    
    def matches(self, email: EmailMetadata) -> bool:
        """Check if email's received date is within the specified range.
        
        Args:
            email: The email metadata to check
            
        Returns:
            True if email date is within range, False otherwise
        """
        if self.start_date and email.received_time < self.start_date:
            return False
        if self.end_date and email.received_time > self.end_date:
            return False
        return True

class SenderFilter(EmailFilter):
    """Filter emails by sender email address.
    
    Performs case-insensitive exact matching against a list of allowed senders.
    
    Attributes:
        senders: List of allowed sender email addresses
    """
    
    def __init__(self, senders: List[str]):
        """Initialize the sender filter.
        
        Args:
            senders: List of sender email addresses to match (case-insensitive)
        """
        self.senders = [s.lower() for s in senders]
    
    def matches(self, email: EmailMetadata) -> bool:
        """Check if email sender matches any of the allowed senders.
        
        Args:
            email: The email metadata to check
            
        Returns:
            True if sender matches, False otherwise
        """
        return email.sender_email.lower() in self.senders

class SubjectKeywordFilter(EmailFilter):
    """Filter emails by keywords in the subject line.
    
    All keywords must be present in the subject (AND logic).
    Matching is case-insensitive.
    
    Attributes:
        keywords: List of keywords that must all appear in subject
    """
    
    def __init__(self, keywords: List[str]):
        """Initialize the subject keyword filter.
        
        Args:
            keywords: Keywords that must all appear in subject (case-insensitive)
        """
        self.keywords = [kw.lower() for kw in keywords]
    
    def matches(self, email: EmailMetadata) -> bool:
        """Check if all keywords appear in the subject.
        
        Args:
            email: The email metadata to check
            
        Returns:
            True if all keywords are found in subject, False otherwise
        """
        subject_lower = email.subject.lower()
        return all(kw in subject_lower for kw in self.keywords)

class BodyKeywordFilter(EmailFilter):
    """Filter emails by keywords in the body text.
    
    All keywords must be present in the body (AND logic).
    Matching is case-insensitive.
    
    Attributes:
        keywords: List of keywords that must all appear in body
    """
    
    def __init__(self, keywords: List[str]):
        """Initialize the body keyword filter.
        
        Args:
            keywords: Keywords that must all appear in body (case-insensitive)
        """
        self.keywords = [kw.lower() for kw in keywords]
    
    def matches(self, email: EmailMetadata) -> bool:
        """Check if all keywords appear in the body.
        
        Args:
            email: The email metadata to check
            
        Returns:
            True if all keywords are found in body, False otherwise
        """
        body_lower = email.body.lower()
        return all(kw in body_lower for kw in self.keywords)

class AttachmentPresenceFilter(EmailFilter):
    """Filter emails based on attachment presence.
    
    This is a placeholder that will be evaluated at runtime based on
    actual attachment count from the Outlook message.
    
    Attributes:
        requires_attachments: True to require attachments, False to exclude them
    """
    
    def __init__(self, requires_attachments: bool):
        """Initialize the attachment presence filter.
        
        Args:
            requires_attachments: True to match only emails with attachments,
                                False to match only emails without attachments
        """
        self.requires_attachments = requires_attachments
    
    def matches(self, email: EmailMetadata) -> bool:
        """This filter requires attachment count from COM object.
        
        This method should not be called directly. The filter criteria
        should be evaluated in the Outlook adapter layer.
        
        Args:
            email: The email metadata (not used)
            
        Returns:
            Always True (actual filtering done at COM layer)
        """
        # This filter needs to be applied at the COM object level
        # where we have access to Attachments.Count
        return True
