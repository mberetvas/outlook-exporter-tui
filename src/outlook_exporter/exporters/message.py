"""Message exporter implementation.

This module exports complete emails as .msg files.
"""

import logging
from pathlib import Path
from typing import Optional

from outlook_exporter.core.models import EmailMetadata, ExportResult
from outlook_exporter.exporters.base import BaseExporter
from outlook_exporter.storage.path_utils import (
    ensure_unique_path,
    sanitize_for_filesystem,
)


logger = logging.getLogger(__name__)


class MessageExporter(BaseExporter):
    """Exports complete emails as .msg files.

    This exporter saves the entire email message in Outlook's .msg format,
    preserving all metadata, attachments, and formatting.
    """

    def export(
        self,
        email: EmailMetadata,
        outlook_message,
        result: ExportResult,
    ) -> Optional[Path]:
        """Export an email as a .msg file.

        Args:
            email: Email metadata
            outlook_message: Outlook COM message object
            result: Export result to update

        Returns:
            Path to the saved .msg file, or None if failed
        """
        folder = self._create_base_folder(email)

        # Create filename from subject
        filename = sanitize_for_filesystem(email.subject, max_length=255) + ".msg"
        dest_path = folder / filename

        if self.config.dry_run:
            logger.info(f"[DRY-RUN] Would save email: {dest_path}")
            result.attachments_saved += 1
            return dest_path

        # Ensure unique path
        final_path = ensure_unique_path(dest_path)

        # Save the message
        try:
            outlook_message.SaveAs(str(final_path))
            logger.info(f"Saved email: {final_path}")
            result.files_exported += 1
            return final_path
        except Exception as e:
            logger.warning(f"Failed to save email '{email.subject}': {e}")
            result.add_error(f"Failed to save email '{email.subject}': {e}")
            return None
