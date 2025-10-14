"""Custom exceptions for the Outlook exporter.

This module defines application-specific exceptions to provide clear
error handling and better error messages.
"""

class OutlookExporterError(Exception):
    """Base exception for all Outlook exporter errors."""
    pass

class OutlookConnectionError(OutlookExporterError):
    """Raised when unable to connect to Outlook or access folders."""
    pass

class FolderNotFoundError(OutlookExporterError):
    """Raised when a specified Outlook folder cannot be found."""
    pass

class FileSystemError(OutlookExporterError):
    """Raised when file system operations fail."""
    pass

class DuplicateDetectionError(OutlookExporterError):
    """Raised when duplicate detection fails."""
    pass

class FilterError(OutlookExporterError):
    """Raised when filter configuration or execution fails."""
    pass

class ExportError(OutlookExporterError):
    """Raised when export operations fail."""
    pass
