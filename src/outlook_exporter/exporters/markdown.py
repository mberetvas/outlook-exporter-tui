"""Markdown exporter implementation.

This module exports emails as markdown files with YAML frontmatter,
ideal for archival and LLM processing.
"""

import logging
from pathlib import Path
from typing import Optional, List

from outlook_exporter.core.models import EmailMetadata, ExportResult
from outlook_exporter.exporters.attachment import AttachmentExporter
from outlook_exporter.storage.path_utils import (
    ensure_unique_path,
    format_file_size,
    get_relative_path,
    sanitize_for_filesystem,
)

try:
    from markdownify import markdownify as md
except ImportError:  # pragma: no cover
    md = None

logger = logging.getLogger(__name__)


class MarkdownExporter(AttachmentExporter):
    """Exports emails as markdown files with attachments.

    This exporter creates markdown files with YAML frontmatter containing
    email metadata, converts HTML body to markdown, and includes references
    to saved attachments. Inherits from AttachmentExporter to reuse
    attachment saving logic.
    """

    def export(
        self,
        email: EmailMetadata,
        outlook_message,
        result: ExportResult,
    ) -> Optional[Path]:
        """Export an email as markdown with attachments.

        Args:
            email: Email metadata
            outlook_message: Outlook COM message object
            result: Export result to update

        Returns:
            Path to the saved markdown file, or None if failed
        """
        if md is None:
            logger.warning("markdownify not available. Install with 'pip install markdownify'")
            result.add_error("markdownify library not installed")
            return None

        folder = self._create_base_folder(email)

        # First, save all attachments to reference them
        saved_attachments = self._export_attachments(outlook_message, folder, result)

        # Create markdown content
        markdown_content = self._create_markdown(email, saved_attachments, folder)

        # Save markdown file
        filename = sanitize_for_filesystem(email.subject, max_length=255) + ".md"
        dest_path = folder / filename

        if self.config.dry_run:
            logger.info(f"[DRY-RUN] Would save markdown: {dest_path}")
            result.attachments_saved += 1
            return dest_path

        final_path = ensure_unique_path(dest_path)

        try:
            with final_path.open("w", encoding="utf-8") as f:
                f.write(markdown_content)
            logger.info(f"Saved markdown: {final_path}")
            result.files_exported += 1
            return final_path
        except Exception as e:
            logger.warning(f"Failed to save markdown '{email.subject}': {e}")
            result.add_error(f"Failed to save markdown '{email.subject}': {e}")
            return None

    def _export_attachments(
        self,
        outlook_message,
        folder: Path,
        result: ExportResult,
    ) -> List[Path]:
        """Export all attachments and return their paths.

        Args:
            outlook_message: Outlook COM message object
            folder: Folder to save attachments in
            result: Export result to update

        Returns:
            List of paths to saved attachments
        """
        saved = []
        try:
            attachments = outlook_message.Attachments
            count = attachments.Count if hasattr(attachments, "Count") else len(attachments)

            for i in range(1, count + 1):
                try:
                    if hasattr(attachments, "Item"):
                        attachment = attachments.Item(i)
                    else:
                        attachment = attachments[i - 1]

                    # Create a fake email metadata for attachment saving
                    # (reusing parent method which needs it)
                    if self._save_attachment(attachment, folder, None):
                        # Get the filename to add to saved list
                        filename = getattr(attachment, "FileName", "attachment.bin")
                        saved.append(folder / sanitize_for_filesystem(filename, 255))
                except Exception as e:
                    logger.debug(f"Error accessing attachment {i}: {e}")
                    continue

        except Exception as e:
            logger.warning(f"Failed to access attachments: {e}")

        return saved

    def _create_markdown(
        self,
        email: EmailMetadata,
        attachments: List[Path],
        markdown_path: Path,
    ) -> str:
        """Create markdown content with YAML frontmatter.

        Args:
            email: Email metadata
            attachments: List of saved attachment paths
            markdown_path: Path where markdown will be saved (for relative paths)

        Returns:
            Complete markdown content as string
        """
        lines = []

        # YAML frontmatter
        lines.append("---")
        lines.append(f"from: {email.sender_email}")
        if email.to_recipients:
            lines.append(f"to: {email.to_recipients}")
        if email.cc_recipients:
            lines.append(f"cc: {email.cc_recipients}")
        lines.append(f"subject: {email.subject}")
        lines.append(f"date: {email.sent_time.isoformat() if email.sent_time else email.received_time.isoformat()}")
        lines.append(f"received: {email.received_time.isoformat()}")
        lines.append("---")
        lines.append("")

        # Email header section
        lines.append(f"# {email.subject}")
        lines.append("")
        lines.append(f"**From:** {email.sender_name} ({email.sender_email})  ")
        if email.to_recipients:
            lines.append(f"**To:** {email.to_recipients}  ")
        if email.cc_recipients:
            lines.append(f"**CC:** {email.cc_recipients}  ")

        sent_time = email.sent_time or email.received_time
        lines.append(f"**Date:** {sent_time.strftime('%B %d, %Y %I:%M %p')}")
        lines.append("")
        lines.append("---")
        lines.append("")

        # Email body (convert HTML to markdown)
        if email.html_body:
            try:
                body_md = md(email.html_body, heading_style="ATX", bullets="-")
                lines.append(body_md)
            except Exception as e:
                logger.warning(f"Failed to convert HTML to markdown: {e}")
                lines.append(email.body or "")
        else:
            lines.append(email.body or "")

        lines.append("")

        # Attachments section
        if attachments:
            lines.append("---")
            lines.append("")
            lines.append("## Attachments")
            lines.append("")
            for att_path in attachments:
                try:
                    size = att_path.stat().st_size
                    size_str = format_file_size(size)
                    rel_path = get_relative_path(att_path, markdown_path.parent)
                    lines.append(f"- [{att_path.name}]({rel_path}) ({size_str})")
                except Exception as e:
                    logger.debug(f"Failed to get attachment info for {att_path}: {e}")
                    lines.append(f"- {att_path.name}")

        return "\n".join(lines)
