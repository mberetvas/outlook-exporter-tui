"""Outlook client for managing connection and folder access.

This module provides a high-level client for interacting with Outlook,
managing the COM connection and providing access to folders and messages.
"""

import logging
from typing import Any, Iterator, Optional

try:
    import win32com.client  # type: ignore
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

from outlook_exporter.utils.exceptions import (
    OutlookConnectionError,
    FolderNotFoundError
)

logger = logging.getLogger(__name__)

# Outlook constants
OUTLOOK_INBOX_ID = 6

class OutlookClient:
    """High-level client for Outlook COM interactions.
    
    Manages the Outlook COM connection, folder navigation, and message iteration.
    Automatically handles COM initialization and cleanup.
    
    Attributes:
        namespace: The Outlook MAPI namespace object
        _com_initialized: Whether COM has been initialized
    """
    
    def __init__(self):
        """Initialize the Outlook client.
        
        Raises:
            OutlookConnectionError: If pywin32 is not available or connection fails
        """
        if win32com is None or pythoncom is None:
            raise OutlookConnectionError(
                "pywin32 is required to connect to Outlook. "
                "Install it with: pip install pywin32"
            )
        
        self._com_initialized = False
        self.namespace: Optional[Any] = None
    
    def __enter__(self):
        """Context manager entry - initializes COM and connects to Outlook."""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - cleans up COM resources."""
        self.disconnect()
    
    def connect(self) -> None:
        """Connect to Outlook and initialize COM.
        
        Raises:
            OutlookConnectionError: If connection fails
        """
        try:
            pythoncom.CoInitialize()
            self._com_initialized = True
            
            outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = outlook.GetNamespace("MAPI")
            
            logger.info("Successfully connected to Outlook")
        
        except Exception as e:
            if self._com_initialized:
                pythoncom.CoUninitialize()
                self._com_initialized = False
            
            raise OutlookConnectionError(
                f"Failed to connect to Outlook: {e}"
            ) from e
    
    def disconnect(self) -> None:
        """Disconnect from Outlook and cleanup COM resources."""
        if self._com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                logger.warning(f"Error during COM cleanup: {e}")
            finally:
                self._com_initialized = False
                self.namespace = None
    
    def get_folder(self, folder_path: Optional[str] = None) -> Any:
        """Get an Outlook folder by path.
        
        If no path is specified, returns the default Inbox folder.
        Path components are separated by '/' or '\\'.
        
        Args:
            folder_path: Folder path (e.g., 'Inbox/SubFolder') or None for Inbox
            
        Returns:
            An Outlook Folder COM object
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            FolderNotFoundError: If the folder path is invalid
            
        Examples:
            >>> client.get_folder()  # Returns Inbox
            >>> client.get_folder('Inbox/Projects')  # Returns Projects subfolder
        """
        if self.namespace is None:
            raise OutlookConnectionError("Not connected to Outlook")
        
        try:
            # Get default Inbox
            folder = self.namespace.GetDefaultFolder(OUTLOOK_INBOX_ID)
            
            if not folder_path:
                return folder
            
            # Navigate to subfolder
            parts = folder_path.replace('\\', '/').split('/')
            
            for part in parts:
                # Skip empty parts and 'Inbox' (we start from Inbox)
                if not part or part.lower() == 'inbox':
                    continue
                
                try:
                    folder = folder.Folders[part]
                except Exception:
                    raise FolderNotFoundError(
                        f"Folder '{part}' not found in path '{folder_path}'"
                    )
            
            logger.info(f"Opened folder: {folder_path or 'Inbox'}")
            return folder
        
        except FolderNotFoundError:
            raise
        except Exception as e:
            raise FolderNotFoundError(
                f"Failed to access folder '{folder_path}': {e}"
            ) from e
    
    def iterate_messages(
        self,
        folder: Any,
        batch_size: int = 200,
        limit: Optional[int] = None
    ) -> Iterator[Any]:
        """Iterate through messages in a folder with batching.
        
        Uses safe iteration to handle COM collection issues. Processes messages
        in batches to avoid loading the entire collection into memory.
        
        Args:
            folder: An Outlook Folder COM object
            batch_size: Number of messages to process per batch
            limit: Maximum number of messages to return (None for unlimited)
            
        Yields:
            Outlook MailItem COM objects
            
        Examples:
            >>> folder = client.get_folder()
            >>> for message in client.iterate_messages(folder, limit=100):
            ...     print(message.Subject)
        """
        try:
            items = folder.Items
            total = items.Count
            processed = 0
            
            logger.info(f"Found {total} messages in folder")
            
            while processed < total:
                if limit and processed >= limit:
                    logger.info(f"Reached message limit of {limit}")
                    break
                
                end_index = min(processed + batch_size, total)
                if limit:
                    end_index = min(end_index, limit)
                
                # COM collections are 1-based
                for i in range(processed + 1, end_index + 1):
                    try:
                        # Use Item method for safe 1-based access
                        if hasattr(items, 'Item'):
                            yield items.Item(i)
                        else:
                            # Fallback to direct indexing
                            yield items[i]
                    except (IndexError, Exception) as e:
                        logger.debug(f"Failed to access message {i}/{total}: {e}")
                        continue
                
                processed = end_index
            
            logger.info(f"Processed {processed} messages")
        
        except Exception as e:
            logger.error(f"Error iterating messages: {e}")
            raise OutlookConnectionError(
                f"Failed to iterate messages: {e}"
            ) from e
