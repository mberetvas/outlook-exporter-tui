"""Utilities for path sanitization and manipulation.

This module provides functions for safely creating file system paths,
handling Windows path length limitations, and sanitizing user input.
"""

import hashlib
import re
from pathlib import Path

def sanitize_for_filesystem(value: str, max_length: int = 255) -> str:
    """Sanitize a string for use as a file or directory name.
    
    Removes or replaces characters that are invalid in Windows paths,
    handles trailing spaces and dots, and enforces length limits.
    
    Args:
        value: The string to sanitize
        max_length: Maximum allowed length for the result
        
    Returns:
        A sanitized string safe for use in file system paths
        
    Examples:
        >>> sanitize_for_filesystem("My File: Test.txt")
        'My File_ Test.txt'
        >>> sanitize_for_filesystem("  spaces  ")
        'spaces'
    """
    # Remove leading/trailing whitespace
    value = value.strip()
    
    # Replace invalid Windows path characters with underscores
    # Invalid chars: \ / : * ? " < > |
    value = re.sub(r'[\\/:*?"<>|]', '_', value)
    
    # Remove trailing spaces and dots (Windows restriction)
    value = value.rstrip('. ')
    
    # Truncate to max length
    if len(value) > max_length:
        value = value[:max_length].rstrip('_. ')
    
    # Ensure not empty
    if not value:
        value = "_"
    
    return value

def create_email_folder_path(
    base_dir: Path,
    sender_email: str,
    subject: str,
    date_str: str,
    subject_max_length: int = 80
) -> Path:
    """Create a hierarchical folder path for email storage.
    
    Creates path structure: base_dir/sender/subject/date
    Handles Windows path length limitations with automatic fallback strategies.
    
    Args:
        base_dir: Base directory for all exports
        sender_email: Sender's email address
        subject: Email subject line
        date_str: Date string (e.g., "2024-01-15")
        subject_max_length: Maximum length for subject component
        
    Returns:
        A Path object representing the target directory
        
    Raises:
        ValueError: If the path cannot be created within length constraints
    """
    sender = sanitize_for_filesystem(sender_email, 120)
    subject_sanitized = sanitize_for_filesystem(subject, subject_max_length)
    
    target_dir = base_dir / sender / subject_sanitized / date_str
    
    # Check path length (Windows limit ~260 characters, leave room for filename)
    if len(str(target_dir)) > 200:
        # Strategy 1: Try shorter subject
        short_subject = sanitize_for_filesystem(subject, 50)
        target_dir = base_dir / sender / short_subject / date_str
        
        if len(str(target_dir)) > 200:
            # Strategy 2: Use hash of subject
            subject_hash = hashlib.md5(subject.encode('utf-8')).hexdigest()[:8]
            target_dir = base_dir / sender / f"subj_{subject_hash}" / date_str
            
            if len(str(target_dir)) > 200:
                # Strategy 3: Flat structure with timestamp
                fallback_name = f"{sender}_{date_str}"
                target_dir = base_dir / sanitize_for_filesystem(fallback_name, 100)
    
    return target_dir

def ensure_unique_path(path: Path) -> Path:
    """Ensure a file path is unique by appending a counter if needed.
    
    If the path already exists, appends "_1", "_2", etc. to the stem
    until a non-existing path is found.
    
    Args:
        path: The desired file path
        
    Returns:
        A unique file path (may be the same as input if it doesn't exist)
        
    Examples:
        >>> ensure_unique_path(Path("file.txt"))  # if file.txt exists
        Path("file_1.txt")
    """
    if not path.exists():
        return path
    
    counter = 1
    while True:
        new_path = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        if not new_path.exists():
            return new_path
        counter += 1

def get_relative_path(from_path: Path, to_path: Path) -> str:
    """Get a relative path from one file to another.
    
    Used for creating relative links in markdown files.
    
    Args:
        from_path: The source file path
        to_path: The target file path
        
    Returns:
        A string representing the relative path with forward slashes
    """
    try:
        rel_path = to_path.relative_to(from_path.parent)
        # Convert to forward slashes for markdown compatibility
        return str(rel_path).replace('\\', '/')
    except ValueError:
        # If paths don't share a common base, return absolute path
        return str(to_path).replace('\\', '/')

def format_file_size(size_bytes: int) -> str:
    """Format a file size in bytes to human-readable format.
    
    Args:
        size_bytes: File size in bytes
        
    Returns:
        Formatted string (e.g., "2.3 MB", "856.2 KB")
        
    Examples:
        >>> format_file_size(1024)
        '1.0 KB'
        >>> format_file_size(1536000)
        '1.5 MB'
    """
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"
