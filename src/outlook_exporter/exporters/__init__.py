"""Export strategies for different output formats.

This package contains exporters that implement different strategies
for exporting emails and attachments.
"""

from outlook_exporter.exporters.attachment import AttachmentExporter
from outlook_exporter.exporters.base import BaseExporter, ExporterFactory
from outlook_exporter.exporters.markdown import MarkdownExporter
from outlook_exporter.exporters.message import MessageExporter

__all__ = [
    "BaseExporter",
    "ExporterFactory",
    "AttachmentExporter",
    "MessageExporter",
    "MarkdownExporter",
]
