"""Unit tests for path utilities."""

from pathlib import Path
import pytest

from outlook_exporter.storage.path_utils import (
    sanitize_for_filesystem,
    create_email_folder_path,
    ensure_unique_path,
    format_file_size,
)

class TestSanitizeForFilesystem:
    """Tests for sanitize_for_filesystem function."""
    
    def test_basic_string(self):
        """Test sanitizing a basic string."""
        result = sanitize_for_filesystem("Normal File Name")
        assert result == "Normal File Name"
    
    def test_invalid_characters(self):
        """Test that invalid Windows characters are replaced."""
        result = sanitize_for_filesystem('File:With*Invalid?Chars<>|"')
        assert result == "File_With_Invalid_Chars____"
    
    def test_forward_slash(self):
        """Test that forward slashes are replaced."""
        result = sanitize_for_filesystem("path/to/file")
        assert result == "path_to_file"
    
    def test_backslash(self):
        """Test that backslashes are replaced."""
        result = sanitize_for_filesystem("path\\to\\file")
        assert result == "path_to_file"
    
    def test_trailing_spaces(self):
        """Test that trailing spaces are removed."""
        result = sanitize_for_filesystem("file  ")
        assert result == "file"
    
    def test_trailing_dots(self):
        """Test that trailing dots are removed."""
        result = sanitize_for_filesystem("file...")
        assert result == "file"
    
    def test_leading_whitespace(self):
        """Test that leading whitespace is removed."""
        result = sanitize_for_filesystem("  file")
        assert result == "file"
    
    def test_max_length(self):
        """Test that strings are truncated to max length."""
        long_string = "a" * 300
        result = sanitize_for_filesystem(long_string, max_length=100)
        assert len(result) == 100
    
    def test_empty_string(self):
        """Test that empty strings become underscore."""
        result = sanitize_for_filesystem("")
        assert result == "_"
    
    def test_whitespace_only(self):
        """Test that whitespace-only strings become underscore."""
        result = sanitize_for_filesystem("   ")
        assert result == "_"

class TestCreateEmailFolderPath:
    """Tests for create_email_folder_path function."""
    
    def test_basic_path_creation(self):
        """Test creating a basic email folder path."""
        base = Path("/tmp/exports")
        result = create_email_folder_path(
            base_dir=base,
            sender_email="sender@example.com",
            subject="Test Email",
            date_str="2024-01-15"
        )
        
        expected = base / "sender@example.com" / "Test Email" / "2024-01-15"
        assert result == expected
    
    def test_sanitizes_sender(self):
        """Test that sender email is sanitized."""
        base = Path("/tmp/exports")
        result = create_email_folder_path(
            base_dir=base,
            sender_email="sender:invalid@example.com",
            subject="Test",
            date_str="2024-01-15"
        )
        
        # Colon should be replaced
        assert "sender_invalid@example.com" in str(result)
    
    def test_sanitizes_subject(self):
        """Test that subject is sanitized."""
        base = Path("/tmp/exports")
        result = create_email_folder_path(
            base_dir=base,
            sender_email="sender@example.com",
            subject="Test: Invalid? Subject",
            date_str="2024-01-15"
        )
        
        # Invalid chars should be replaced
        assert "Test_ Invalid_ Subject" in str(result)
    
    def test_subject_max_length(self):
        """Test that subject is truncated to max length."""
        base = Path("/tmp/exports")
        long_subject = "a" * 200
        result = create_email_folder_path(
            base_dir=base,
            sender_email="sender@example.com",
            subject=long_subject,
            date_str="2024-01-15",
            subject_max_length=50
        )
        
        # Subject component should be 50 chars or less
        parts = result.parts
        subject_part = parts[-2]  # Second to last part
        assert len(subject_part) <= 50

class TestEnsureUniquePath:
    """Tests for ensure_unique_path function."""
    
    def test_non_existing_path(self, tmp_path):
        """Test that non-existing paths are returned unchanged."""
        file_path = tmp_path / "test.txt"
        result = ensure_unique_path(file_path)
        assert result == file_path
    
    def test_existing_path_gets_suffix(self, tmp_path):
        """Test that existing paths get a numeric suffix."""
        file_path = tmp_path / "test.txt"
        file_path.write_text("content")
        
        result = ensure_unique_path(file_path)
        assert result == tmp_path / "test_1.txt"
    
    def test_multiple_existing_paths(self, tmp_path):
        """Test incrementing suffix for multiple collisions."""
        base = tmp_path / "test.txt"
        base.write_text("content")
        (tmp_path / "test_1.txt").write_text("content")
        (tmp_path / "test_2.txt").write_text("content")
        
        result = ensure_unique_path(base)
        assert result == tmp_path / "test_3.txt"

class TestFormatFileSize:
    """Tests for format_file_size function."""
    
    def test_bytes(self):
        """Test formatting sizes in bytes."""
        assert format_file_size(100) == "100.0 B"
        assert format_file_size(512) == "512.0 B"
    
    def test_kilobytes(self):
        """Test formatting sizes in kilobytes."""
        assert format_file_size(1024) == "1.0 KB"
        assert format_file_size(2048) == "2.0 KB"
        assert format_file_size(1536) == "1.5 KB"
    
    def test_megabytes(self):
        """Test formatting sizes in megabytes."""
        assert format_file_size(1024 * 1024) == "1.0 MB"
        assert format_file_size(1024 * 1024 * 2.5) == "2.5 MB"
    
    def test_gigabytes(self):
        """Test formatting sizes in gigabytes."""
        assert format_file_size(1024 * 1024 * 1024) == "1.0 GB"
        assert format_file_size(1024 * 1024 * 1024 * 3.7) == "3.7 GB"
    
    def test_terabytes(self):
        """Test formatting sizes in terabytes."""
        assert format_file_size(1024 * 1024 * 1024 * 1024) == "1.0 TB"
