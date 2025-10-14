"""Base classes for email filtering.

This module defines the abstract base class for all email filters,
following the Strategy pattern for flexible and composable filtering.
"""

from abc import ABC, abstractmethod
from typing import List

from outlook_exporter.core.models import EmailMetadata


class EmailFilter(ABC):
    """Abstract base class for email filters.
    
    All concrete filters must implement the matches() method to define
    their filtering logic.
    """
    
    @abstractmethod
    def matches(self, email: EmailMetadata) -> bool:
        """Check if an email matches this filter's criteria.
        
        Args:
            email: The email metadata to check
            
        Returns:
            True if the email matches the filter criteria, False otherwise
        """
        pass


class CompositeFilter(EmailFilter):
    """Combines multiple filters using AND logic.
    
    An email must match all component filters to pass this composite filter.
    This allows building complex filter logic from simple components.
    
    Example:
        >>> date_filter = DateRangeFilter(start_date, end_date)
        >>> sender_filter = SenderFilter(['user@example.com'])
        >>> composite = CompositeFilter([date_filter, sender_filter])
        >>> composite.matches(email)  # True only if both filters match
    """
    
    def __init__(self, filters: List[EmailFilter]):
        """Initialize the composite filter.
        
        Args:
            filters: List of filters to combine with AND logic
        """
        self.filters = filters
    
    def matches(self, email: EmailMetadata) -> bool:
        """Check if email matches all component filters.
        
        Args:
            email: The email metadata to check
            
        Returns:
            True if email matches all filters, False otherwise
        """
        return all(filter.matches(email) for filter in self.filters)


class PassThroughFilter(EmailFilter):
    """A filter that always returns True.
    
    Useful as a default or null filter when no filtering is needed.
    """
    
    def matches(self, email: EmailMetadata) -> bool:
        """Always returns True.
        
        Args:
            email: The email metadata (ignored)
            
        Returns:
            Always True
        """
        return True
