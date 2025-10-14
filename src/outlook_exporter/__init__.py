"""Outlook Exporter - Export email attachments from Microsoft Outlook.

A Python tool for efficiently exporting email attachments from Outlook with
advanced filtering, duplicate detection, and organized folder structures.
"""

__version__ = "0.2.0"
__author__ = "mberetvas"
__license__ = "MIT"

from outlook_exporter.core.models import (
    EmailMetadata,
    Attachment,
    ExportConfig,
    FilterCriteria,
    ExportResult,
)
from outlook_exporter.utils.exceptions import (
    OutlookExporterError,
    OutlookConnectionError,
    FolderNotFoundError,
    FileSystemError,
    DuplicateDetectionError,
    FilterError,
    ExportError,
)

__all__ = [
    # Models
    "EmailMetadata",
    "Attachment",
    "ExportConfig",
    "FilterCriteria",
    "ExportResult",
    # Exceptions
    "OutlookExporterError",
    "OutlookConnectionError",
    "FolderNotFoundError",
    "FileSystemError",
    "DuplicateDetectionError",
    "FilterError",
    "ExportError",
]
