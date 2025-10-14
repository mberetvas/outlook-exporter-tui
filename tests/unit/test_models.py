"""Unit tests for domain models."""

from datetime import datetime
from pathlib import Path

import pytest

from outlook_exporter.core.models import (
    EmailMetadata,
    Attachment,
    ExportConfig,
    FilterCriteria,
    ExportResult,
)

def test_email_metadata_creation():
    """Test creating an EmailMetadata instance."""
    now = datetime.now()
    email = EmailMetadata(
        subject="Test Subject",
        sender_email="sender@example.com",
        sender_name="Test Sender",
        received_time=now,
        sent_time=now,
        to_recipients="recipient@example.com",
        cc_recipients="cc@example.com",
        body="Test body",
        html_body="<p>Test body</p>"
    )
    
    assert email.subject == "Test Subject"
    assert email.sender_email == "sender@example.com"
    assert email.received_time == now

def test_attachment_creation():
    """Test creating an Attachment instance."""
    attachment = Attachment(
        filename="test.pdf",
        size=1024,
        content_hash="abc123",
        is_inline=False
    )
    
    assert attachment.filename == "test.pdf"
    assert attachment.size == 1024
    assert attachment.content_hash == "abc123"
    assert not attachment.is_inline

def test_attachment_defaults():
    """Test Attachment default values."""
    attachment = Attachment(filename="test.txt", size=100)
    
    assert attachment.content_hash is None
    assert not attachment.is_inline
    assert attachment.position == 0

def test_export_config_creation():
    """Test creating an ExportConfig instance."""
    config = ExportConfig(
        output_dir=Path("/tmp/exports"),
        include_inline=True,
        duplicates_subfolder="dupes",
        hash_algorithm="md5"
    )
    
    assert config.output_dir == Path("/tmp/exports")
    assert config.include_inline
    assert config.duplicates_subfolder == "dupes"
    assert config.hash_algorithm == "md5"

def test_export_config_defaults():
    """Test ExportConfig default values."""
    config = ExportConfig(output_dir=Path("/tmp"))
    
    assert not config.include_inline
    assert config.duplicates_subfolder == "duplicates"
    assert config.hash_algorithm == "sha256"
    assert not config.dry_run
    assert config.subject_sanitize_length == 80
    assert config.batch_size == 200
    assert config.limit is None

def test_filter_criteria_defaults():
    """Test FilterCriteria default values."""
    criteria = FilterCriteria()
    
    assert criteria.start_date is None
    assert criteria.end_date is None
    assert criteria.senders == []
    assert criteria.subject_keywords == []
    assert criteria.body_keywords == []
    assert not criteria.with_attachments
    assert not criteria.without_attachments
    assert criteria.folder_path is None

def test_export_result_initialization():
    """Test ExportResult initialization."""
    result = ExportResult()
    
    assert result.messages_processed == 0
    assert result.messages_matched == 0
    assert result.attachments_saved == 0
    assert result.duplicates_found == 0
    assert result.errors == []

def test_export_result_add_error():
    """Test adding errors to ExportResult."""
    result = ExportResult()
    result.add_error("Error 1")
    result.add_error("Error 2")
    
    assert len(result.errors) == 2
    assert "Error 1" in result.errors
    assert "Error 2" in result.errors
