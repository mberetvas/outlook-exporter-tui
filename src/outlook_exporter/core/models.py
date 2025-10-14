"""Domain models for the Outlook exporter.

This module defines the core data structures used throughout the application,
following domain-driven design principles.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional, List


@dataclass
class EmailMetadata:
    """Represents email metadata extracted from an Outlook message.
    
    Attributes:
        subject: Email subject line
        sender_email: Sender's email address
        sender_name: Sender's display name
        received_time: When the email was received
        sent_time: When the email was sent
        to_recipients: Semicolon-separated list of TO recipients
        cc_recipients: Semicolon-separated list of CC recipients
        body: Plain text body of the email
        html_body: HTML body of the email
    """
    subject: str
    sender_email: str
    sender_name: str
    received_time: datetime
    sent_time: Optional[datetime]
    to_recipients: str
    cc_recipients: str
    body: str
    html_body: str


@dataclass
class Attachment:
    """Represents an email attachment.
    
    Attributes:
        filename: Name of the attachment file
        size: Size in bytes
        content_hash: Hash of file content (for duplicate detection)
        is_inline: Whether this is an inline attachment (e.g., signature image)
        position: Position in the email (0 for inline attachments)
    """
    filename: str
    size: int
    content_hash: Optional[str] = None
    is_inline: bool = False
    position: int = 0


@dataclass
class ExportConfig:
    """Configuration for the export operation.
    
    Attributes:
        output_dir: Base directory for exported files
        include_inline: Whether to include inline attachments
        duplicates_subfolder: Name of subfolder for duplicate files
        hash_algorithm: Algorithm to use for duplicate detection
        dry_run: If True, simulate export without writing files
        subject_sanitize_length: Max length for subject in folder names
        batch_size: Number of messages to process in each batch
        limit: Maximum number of messages to process (None for unlimited)
    """
    output_dir: Path
    include_inline: bool = False
    duplicates_subfolder: str = "duplicates"
    hash_algorithm: str = "sha256"
    dry_run: bool = False
    subject_sanitize_length: int = 80
    batch_size: int = 200
    limit: Optional[int] = None


@dataclass
class FilterCriteria:
    """Criteria for filtering emails.
    
    Attributes:
        start_date: Filter emails received on or after this date
        end_date: Filter emails received on or before this date
        senders: List of sender email addresses to match
        subject_keywords: Keywords that must all appear in subject
        body_keywords: Keywords that must all appear in body
        with_attachments: If True, only match emails with attachments
        without_attachments: If True, only match emails without attachments
        folder_path: Outlook folder path (e.g., 'Inbox/Projects')
    """
    start_date: Optional[datetime] = None
    end_date: Optional[datetime] = None
    senders: List[str] = field(default_factory=list)
    subject_keywords: List[str] = field(default_factory=list)
    body_keywords: List[str] = field(default_factory=list)
    with_attachments: bool = False
    without_attachments: bool = False
    folder_path: Optional[str] = None


@dataclass
class ExportResult:
    """Result of an export operation.
    
    Attributes:
        messages_processed: Number of messages processed
        messages_matched: Number of messages matching filters
        attachments_saved: Number of attachments saved
        duplicates_found: Number of duplicate attachments found
        errors: List of error messages encountered
    """
    messages_processed: int = 0
    messages_matched: int = 0
    attachments_saved: int = 0
    duplicates_found: int = 0
    errors: List[str] = field(default_factory=list)
    
    def add_error(self, error: str) -> None:
        """Add an error message to the result."""
        self.errors.append(error)
