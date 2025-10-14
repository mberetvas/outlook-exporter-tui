"""Duplicate detection using content hashing.

This module provides functionality for detecting duplicate files based on
their content hash, allowing intelligent handling of duplicate attachments.
"""

import hashlib
import logging
from pathlib import Path
from typing import Dict, List

from outlook_exporter.utils.exceptions import DuplicateDetectionError

logger = logging.getLogger(__name__)

class DuplicateTracker:
    """Tracks file hashes to detect duplicate attachments.
    
    Uses content-based hashing to identify duplicate files, regardless of
    their filename or location. Maintains a registry of all seen files.
    
    Attributes:
        hash_algorithm: The hash algorithm to use (e.g., 'sha256', 'md5')
        seen_hashes: Dictionary mapping hash values to lists of file paths
    """
    
    def __init__(self, hash_algorithm: str = "sha256"):
        """Initialize the duplicate tracker.
        
        Args:
            hash_algorithm: Hash algorithm to use for duplicate detection
            
        Raises:
            DuplicateDetectionError: If the hash algorithm is not available
        """
        if hash_algorithm not in hashlib.algorithms_available:
            raise DuplicateDetectionError(
                f"Hash algorithm '{hash_algorithm}' is not available. "
                f"Available algorithms: {', '.join(sorted(hashlib.algorithms_available))}"
            )
        
        self.hash_algorithm = hash_algorithm
        self.seen_hashes: Dict[str, List[Path]] = {}
    
    def compute_file_hash(self, file_path: Path) -> str:
        """Compute the hash of a file's contents.
        
        Reads the file in chunks to handle large files efficiently.
        
        Args:
            file_path: Path to the file to hash
            
        Returns:
            Hexadecimal string representation of the file's hash
            
        Raises:
            DuplicateDetectionError: If the file cannot be read or hashed
        """
        try:
            hasher = hashlib.new(self.hash_algorithm)
            
            with file_path.open('rb') as f:
                # Read in 8KB chunks to handle large files
                for chunk in iter(lambda: f.read(8192), b''):
                    hasher.update(chunk)
            
            return hasher.hexdigest()
        
        except (IOError, OSError) as e:
            raise DuplicateDetectionError(
                f"Failed to compute hash for {file_path}: {e}"
            ) from e
    
    def is_duplicate(self, file_hash: str) -> bool:
        """Check if a file hash has been seen before.
        
        Args:
            file_hash: The hash to check
            
        Returns:
            True if this hash has been seen before, False otherwise
        """
        return file_hash in self.seen_hashes
    
    def register_file(self, file_path: Path, file_hash: str) -> None:
        """Register a file in the duplicate tracker.
        
        Args:
            file_path: Path where the file was saved
            file_hash: Hash of the file's contents
        """
        if file_hash not in self.seen_hashes:
            self.seen_hashes[file_hash] = []
        
        self.seen_hashes[file_hash].append(file_path)
        logger.debug(f"Registered file {file_path.name} with hash {file_hash[:8]}...")
    
    def get_original_locations(self, file_hash: str) -> List[Path]:
        """Get all locations where this hash has been saved.
        
        Args:
            file_hash: The hash to look up
            
        Returns:
            List of paths where files with this hash have been saved
        """
        return self.seen_hashes.get(file_hash, [])
    
    def get_statistics(self) -> Dict[str, int]:
        """Get statistics about duplicate detection.
        
        Returns:
            Dictionary with statistics:
                - unique_files: Number of unique file hashes
                - total_files: Total number of files tracked
                - duplicate_files: Number of duplicate files detected
        """
        unique_files = len(self.seen_hashes)
        total_files = sum(len(paths) for paths in self.seen_hashes.values())
        duplicate_files = total_files - unique_files
        
        return {
            "unique_files": unique_files,
            "total_files": total_files,
            "duplicate_files": duplicate_files
        }
    
    def clear(self) -> None:
        """Clear all tracked hashes and start fresh."""
        self.seen_hashes.clear()
        logger.debug("Cleared duplicate tracker")
