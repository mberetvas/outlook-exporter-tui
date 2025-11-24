"""Attachment exporter implementation.

This module exports email attachments to disk with duplicate detection
and organized folder structures.
"""

import logging
from pathlib import Path
from typing import Optional, Dict, List

from outlook_exporter.core.duplicates import DuplicateTracker
from outlook_exporter.core.models import EmailMetadata, ExportResult
from outlook_exporter.exporters.base import BaseExporter
from outlook_exporter.storage.path_utils import (
    ensure_unique_path,
    sanitize_for_filesystem,
)


logger = logging.getLogger(__name__)


class AttachmentExporter(BaseExporter):
    """Exports email attachments to organized folder structure.

    This exporter saves individual attachments from emails,
    organizing them by sender/subject/date and detecting duplicates.
    """

    def __init__(self, *args, **kwargs):
        """Initialize the attachment exporter."""
        super().__init__(*args, **kwargs)
        self.duplicate_tracker = DuplicateTracker(
            algorithm=self.config.hash_algorithm
        )

    def export(
        self,
        email: EmailMetadata,
        outlook_message,
        result: ExportResult,
    ) -> Optional[Path]:
        """Export all attachments from an email.

        Args:
            email: Email metadata
            outlook_message: Outlook COM message object
            result: Export result to update

        Returns:
            Path to the folder containing exported attachments
        """
        folder = self._create_base_folder(email)
        attachments_saved = 0

        # Get attachments from COM object
        try:
            attachments = outlook_message.Attachments
            count = attachments.Count if hasattr(attachments, "Count") else len(attachments)
        except Exception as e:
            logger.warning(f"Failed to access attachments: {e}")
            result.add_error(f"Failed to access attachments for '{email.subject}': {e}")
            return None

        # Process each attachment
        for i in range(1, count + 1):
            try:
                if hasattr(attachments, "Item"):
                    attachment = attachments.Item(i)
                else:
                    attachment = attachments[i - 1]

                if self._save_attachment(attachment, folder, email):
                    attachments_saved += 1

            except Exception as e:
                logger.debug(f"Error accessing attachment {i}: {e}")
                continue

        result.attachments_saved += attachments_saved
        return folder if attachments_saved > 0 else None

    def _save_attachment(
        self,
        attachment,
        folder: Path,
        email: EmailMetadata,
    ) -> bool:
        """Save a single attachment to disk.

        Args:
            attachment: Outlook attachment COM object
            folder: Folder to save the attachment in
            email: Email metadata for logging

        Returns:
            True if attachment was saved successfully
        """
        # Skip inline attachments unless configured to include them
        if not self.config.include_inline:
            try:
                position = getattr(attachment, "Position", 0)
                att_type = getattr(attachment, "Type", None)
                if position == 0 and att_type == 5:  # Likely inline
                    logger.debug("Skipping inline attachment")
                    return False
            except Exception:
                pass

        # Get filename
        try:
            filename = getattr(attachment, "FileName", "attachment.bin") or "attachment.bin"
            filename = sanitize_for_filesystem(filename, max_length=255)
        except Exception as e:
            logger.warning(f"Failed to get attachment filename: {e}")
            return False

        # Determine save path
        dest_path = folder / filename

        if self.config.dry_run:
            logger.info(f"[DRY-RUN] Would save attachment: {dest_path}")
            return True

        # Save to temporary file first
        temp_path = dest_path.with_suffix(dest_path.suffix + ".tmp")
        try:
            attachment.SaveAsFile(str(temp_path))
        except Exception as e:
            logger.warning(f"Failed to save attachment {filename}: {e}")
            return False

        # Check for duplicates
        file_hash = self.duplicate_tracker.compute_hash(temp_path)
        is_duplicate = self.duplicate_tracker.is_duplicate(file_hash)

        if is_duplicate:
            # Move to duplicates subfolder
            dup_folder = folder / self.config.duplicates_subfolder
            dup_folder.mkdir(parents=True, exist_ok=True)
            final_path = ensure_unique_path(dup_folder / filename)
            temp_path.rename(final_path)
            logger.info(f"Duplicate detected: {final_path}")
            self.duplicate_tracker.add_duplicate(file_hash, final_path)
        else:
            # Save as new file
            final_path = ensure_unique_path(dest_path)
            temp_path.rename(final_path)
            logger.debug(f"Saved attachment: {final_path}")
            self.duplicate_tracker.add_file(file_hash, final_path)

        return True
