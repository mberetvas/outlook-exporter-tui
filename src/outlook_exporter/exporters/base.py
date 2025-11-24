"""Base exporter interface and common functionality.

This module defines the abstract base class for all exporters,
following the Strategy pattern for flexible export implementations.
"""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional

from outlook_exporter.core.models import EmailMetadata, ExportConfig, ExportResult


class BaseExporter(ABC):
    """Abstract base class for all exporters.

    Exporters handle the actual export of emails and attachments to disk,
    implementing different strategies (attachments only, full .msg, markdown, etc.).
    """

    def __init__(self, config: ExportConfig):
        """Initialize the exporter with configuration.

        Args:
            config: Export configuration settings
        """
        self.config = config

    @abstractmethod
    def export(
        self,
        email: EmailMetadata,
        outlook_message,
        result: ExportResult,
    ) -> Optional[Path]:
        """Export an email according to the exporter's strategy.

        Args:
            email: Email metadata extracted from Outlook
            outlook_message: Raw Outlook COM object for saving
            result: Export result to update with statistics

        Returns:
            Path to the exported file/folder, or None if export failed
        """
        pass

    def _create_base_folder(self, email: EmailMetadata) -> Path:
        """Create the base folder structure for an email.

        Uses the pattern: {output_dir}/{sender}/{subject}/{date}/

        Args:
            email: Email metadata containing sender, subject, and date info

        Returns:
            Path to the created folder
        """
        from outlook_exporter.storage.path_utils import create_email_folder_path

        folder_path = create_email_folder_path(
            base_dir=self.config.output_dir,
            sender_email=email.sender_email,
            subject=email.subject,
            received_time=email.received_time,
            subject_max_length=self.config.subject_sanitize_length,
        )

        if not self.config.dry_run:
            folder_path.mkdir(parents=True, exist_ok=True)

        return folder_path


class ExporterFactory:
    """Factory for creating appropriate exporter instances.

    This factory selects the right exporter based on configuration,
    following the Factory pattern for clean object creation.
    """

    @staticmethod
    def create_exporter(config: ExportConfig, export_type: str) -> BaseExporter:
        """Create an exporter based on the export type.

        Args:
            config: Export configuration
            export_type: Type of export ('attachments', 'msg', 'markdown')

        Returns:
            Appropriate exporter instance

        Raises:
            ValueError: If export_type is not recognized
        """
        from outlook_exporter.exporters.attachment import AttachmentExporter
        from outlook_exporter.exporters.markdown import MarkdownExporter
        from outlook_exporter.exporters.message import MessageExporter

        exporters = {
            "attachments": AttachmentExporter,
            "msg": MessageExporter,
            "markdown": MarkdownExporter,
        }

        exporter_class = exporters.get(export_type)
        if not exporter_class:
            raise ValueError(
                f"Unknown export type: {export_type}. "
                f"Valid types: {', '.join(exporters.keys())}"
            )

        return exporter_class(config)
