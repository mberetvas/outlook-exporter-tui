"""Adapters for Outlook COM objects.

This module provides adapter classes that safely extract data from Outlook
COM objects and convert them to domain models, isolating the rest of the
application from COM-specific code.
"""

from datetime import datetime
from typing import Any, Optional
import logging

from outlook_exporter.core.models import EmailMetadata, Attachment
from outlook_exporter.utils.exceptions import OutlookConnectionError

logger = logging.getLogger(__name__)

def safe_get_com_property(obj: Any, property_name: str, default: Any = None) -> Any:
    """Safely retrieve a property from a COM object.
    
    COM property access can fail for various reasons (e.g., "Operation aborted").
    This function catches exceptions and returns a default value.
    
    Args:
        obj: The COM object
        property_name: Name of the property to retrieve
        default: Value to return if property access fails
        
    Returns:
        The property value if successful, default value otherwise
    """
    try:
        return getattr(obj, property_name)
    except Exception as e:
        logger.debug(
            f"Failed to access {property_name} on {type(obj).__name__}: {e}"
        )
        return default

def parse_com_datetime(com_date: Any) -> datetime:
    """Parse a datetime from a COM object.
    
    COM datetime objects may have different formats. This function attempts
    to convert them to Python datetime objects.
    
    Args:
        com_date: The COM datetime object
        
    Returns:
        A Python datetime object
    """
    if isinstance(com_date, datetime):
        return com_date
    
    # Try to convert COM date string format
    try:
        if hasattr(com_date, 'Format'):
            # Some COM dates have a Format method
            date_str = str(com_date)
            return datetime.strptime(date_str, '%m/%d/%y %H:%M:%S')
    except (ValueError, AttributeError):
        pass
    
    # Fallback to current time
    logger.warning(f"Could not parse COM datetime: {com_date}, using current time")
    return datetime.now()

class OutlookMessageAdapter:
    """Adapter for Outlook MailItem COM objects.
    
    Provides a safe interface for extracting email metadata from Outlook
    COM objects, handling errors and converting to domain models.
    
    Attributes:
        com_message: The underlying COM MailItem object
    """
    
    def __init__(self, com_message: Any):
        """Initialize the message adapter.
        
        Args:
            com_message: An Outlook MailItem COM object
        """
        self.com_message = com_message
    
    def to_metadata(self) -> EmailMetadata:
        """Convert the COM message to an EmailMetadata domain model.
        
        Returns:
            EmailMetadata object with extracted information
        """
        subject = safe_get_com_property(self.com_message, 'Subject', '') or ''
        sender_email = safe_get_com_property(
            self.com_message, 'SenderEmailAddress', ''
        ) or 'unknown'
        sender_name = safe_get_com_property(
            self.com_message, 'SenderName', ''
        ) or sender_email
        
        # Get datetime fields
        received_raw = safe_get_com_property(
            self.com_message, 'ReceivedTime', datetime.now()
        )
        received_time = parse_com_datetime(received_raw)
        
        sent_raw = safe_get_com_property(self.com_message, 'SentOn', None)
        sent_time = parse_com_datetime(sent_raw) if sent_raw else None
        
        to_recipients = safe_get_com_property(self.com_message, 'To', '') or ''
        cc_recipients = safe_get_com_property(self.com_message, 'CC', '') or ''
        body = safe_get_com_property(self.com_message, 'Body', '') or ''
        html_body = safe_get_com_property(self.com_message, 'HTMLBody', '') or ''
        
        return EmailMetadata(
            subject=subject,
            sender_email=sender_email,
            sender_name=sender_name,
            received_time=received_time,
            sent_time=sent_time,
            to_recipients=to_recipients,
            cc_recipients=cc_recipients,
            body=body,
            html_body=html_body
        )
    
    def get_attachment_count(self) -> int:
        """Get the number of attachments in the message.
        
        Returns:
            Number of attachments (0 if none or if access fails)
        """
        attachments = safe_get_com_property(self.com_message, 'Attachments', None)
        if attachments is None:
            return 0
        
        # COM collections have .Count property
        if hasattr(attachments, 'Count'):
            try:
                return int(attachments.Count)
            except Exception:
                return 0
        
        # Fallback: try len()
        try:
            return len(attachments)
        except Exception:
            return 0
    
    def get_attachments(self) -> list[Any]:
        """Get all attachments from the message.
        
        Returns:
            List of COM Attachment objects
        """
        attachments = safe_get_com_property(self.com_message, 'Attachments', None)
        if attachments is None:
            return []
        
        result = []
        count = self.get_attachment_count()
        
        # COM collections are 1-based
        for idx in range(1, count + 1):
            try:
                if hasattr(attachments, 'Item'):
                    # Use Item method for safe access
                    result.append(attachments.Item(idx))
                else:
                    # Fallback to 0-based indexing
                    result.append(attachments[idx - 1])
            except (IndexError, Exception) as e:
                logger.debug(f"Failed to access attachment {idx}/{count}: {e}")
                continue
        
        return result
    
    def save_as_msg(self, file_path: str) -> None:
        """Save the entire message as a .msg file.
        
        Args:
            file_path: Path where to save the .msg file
            
        Raises:
            OutlookConnectionError: If the save operation fails
        """
        try:
            self.com_message.SaveAs(file_path)
        except Exception as e:
            raise OutlookConnectionError(
                f"Failed to save message as .msg: {e}"
            ) from e

class OutlookAttachmentAdapter:
    """Adapter for Outlook Attachment COM objects.
    
    Provides a safe interface for extracting attachment information and
    saving attachments to disk.
    
    Attributes:
        com_attachment: The underlying COM Attachment object
    """
    
    def __init__(self, com_attachment: Any):
        """Initialize the attachment adapter.
        
        Args:
            com_attachment: An Outlook Attachment COM object
        """
        self.com_attachment = com_attachment
    
    def to_attachment_info(self) -> Attachment:
        """Convert the COM attachment to an Attachment domain model.
        
        Returns:
            Attachment object with extracted information
        """
        filename = safe_get_com_property(
            self.com_attachment, 'FileName', 'attachment.bin'
        ) or 'attachment.bin'
        
        size = safe_get_com_property(self.com_attachment, 'Size', 0) or 0
        position = safe_get_com_property(self.com_attachment, 'Position', 0) or 0
        
        # Heuristic for inline attachments: Position = 0 often indicates inline
        # Type 5 is olByValue which may be inline images
        attachment_type = safe_get_com_property(self.com_attachment, 'Type', None)
        is_inline = position == 0 and attachment_type == 5
        
        return Attachment(
            filename=filename,
            size=size,
            is_inline=is_inline,
            position=position
        )
    
    def save_to_file(self, file_path: str) -> None:
        """Save the attachment to a file.
        
        Args:
            file_path: Path where to save the attachment
            
        Raises:
            OutlookConnectionError: If the save operation fails
        """
        try:
            self.com_attachment.SaveAsFile(file_path)
        except Exception as e:
            raise OutlookConnectionError(
                f"Failed to save attachment: {e}"
            ) from e
