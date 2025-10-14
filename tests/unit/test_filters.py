"""Unit tests for email filters."""

from datetime import datetime, timedelta

import pytest

from outlook_exporter.core.models import EmailMetadata
from outlook_exporter.filters.base import CompositeFilter, PassThroughFilter
from outlook_exporter.filters.email_filters import (
    DateRangeFilter,
    SenderFilter,
    SubjectKeywordFilter,
    BodyKeywordFilter,
)

def create_test_email(
    subject: str = "Test Subject",
    body: str = "Test body",
    sender: str = "sender@example.com",
    days_offset: int = 0
) -> EmailMetadata:
    """Create a test EmailMetadata instance."""
    received = datetime.now() + timedelta(days=days_offset)
    return EmailMetadata(
        subject=subject,
        sender_email=sender,
        sender_name="Test Sender",
        received_time=received,
        sent_time=received,
        to_recipients="",
        cc_recipients="",
        body=body,
        html_body=""
    )

class TestPassThroughFilter:
    """Tests for PassThroughFilter."""
    
    def test_always_matches(self):
        """Test that PassThroughFilter always returns True."""
        filter = PassThroughFilter()
        email = create_test_email()
        assert filter.matches(email)

class TestDateRangeFilter:
    """Tests for DateRangeFilter."""
    
    def test_no_date_constraints(self):
        """Test filter with no date constraints matches everything."""
        filter = DateRangeFilter()
        email = create_test_email(days_offset=-10)
        assert filter.matches(email)
    
    def test_start_date_only(self):
        """Test filter with only start date."""
        start = datetime.now()
        filter = DateRangeFilter(start_date=start)
        
        # Email from 2 days ago should not match
        old_email = create_test_email(days_offset=-2)
        assert not filter.matches(old_email)
        
        # Email from today should match
        recent_email = create_test_email(days_offset=0)
        assert filter.matches(recent_email)
    
    def test_end_date_only(self):
        """Test filter with only end date."""
        end = datetime.now()
        filter = DateRangeFilter(end_date=end)
        
        # Email from 2 days ago should match
        old_email = create_test_email(days_offset=-2)
        assert filter.matches(old_email)
        
        # Email from 2 days in future should not match
        future_email = create_test_email(days_offset=2)
        assert not filter.matches(future_email)
    
    def test_date_range(self):
        """Test filter with both start and end dates."""
        start = datetime.now() - timedelta(days=7)
        end = datetime.now()
        filter = DateRangeFilter(start_date=start, end_date=end)
        
        # Email within range
        valid_email = create_test_email(days_offset=-3)
        assert filter.matches(valid_email)
        
        # Email before range
        too_old = create_test_email(days_offset=-10)
        assert not filter.matches(too_old)
        
        # Email after range
        too_new = create_test_email(days_offset=2)
        assert not filter.matches(too_new)

class TestSenderFilter:
    """Tests for SenderFilter."""
    
    def test_single_sender_match(self):
        """Test matching a single sender."""
        filter = SenderFilter(["sender@example.com"])
        email = create_test_email(sender="sender@example.com")
        assert filter.matches(email)
    
    def test_single_sender_no_match(self):
        """Test non-matching sender."""
        filter = SenderFilter(["sender@example.com"])
        email = create_test_email(sender="other@example.com")
        assert not filter.matches(email)
    
    def test_case_insensitive(self):
        """Test that sender matching is case-insensitive."""
        filter = SenderFilter(["SENDER@EXAMPLE.COM"])
        email = create_test_email(sender="sender@example.com")
        assert filter.matches(email)
    
    def test_multiple_senders(self):
        """Test matching multiple allowed senders."""
        filter = SenderFilter(["sender1@example.com", "sender2@example.com"])
        
        email1 = create_test_email(sender="sender1@example.com")
        assert filter.matches(email1)
        
        email2 = create_test_email(sender="sender2@example.com")
        assert filter.matches(email2)
        
        email3 = create_test_email(sender="sender3@example.com")
        assert not filter.matches(email3)

class TestSubjectKeywordFilter:
    """Tests for SubjectKeywordFilter."""
    
    def test_single_keyword_match(self):
        """Test matching a single keyword in subject."""
        filter = SubjectKeywordFilter(["invoice"])
        email = create_test_email(subject="Quarterly Invoice Report")
        assert filter.matches(email)
    
    def test_single_keyword_no_match(self):
        """Test non-matching keyword."""
        filter = SubjectKeywordFilter(["invoice"])
        email = create_test_email(subject="Monthly Report")
        assert not filter.matches(email)
    
    def test_case_insensitive(self):
        """Test that keyword matching is case-insensitive."""
        filter = SubjectKeywordFilter(["INVOICE"])
        email = create_test_email(subject="quarterly invoice")
        assert filter.matches(email)
    
    def test_multiple_keywords_all_required(self):
        """Test that all keywords must be present (AND logic)."""
        filter = SubjectKeywordFilter(["urgent", "invoice"])
        
        # Both keywords present
        email1 = create_test_email(subject="Urgent: Invoice Required")
        assert filter.matches(email1)
        
        # Only one keyword present
        email2 = create_test_email(subject="Urgent Report")
        assert not filter.matches(email2)

class TestBodyKeywordFilter:
    """Tests for BodyKeywordFilter."""
    
    def test_single_keyword_match(self):
        """Test matching a single keyword in body."""
        filter = BodyKeywordFilter(["urgent"])
        email = create_test_email(body="This is an urgent message")
        assert filter.matches(email)
    
    def test_single_keyword_no_match(self):
        """Test non-matching keyword."""
        filter = BodyKeywordFilter(["urgent"])
        email = create_test_email(body="This is a normal message")
        assert not filter.matches(email)
    
    def test_case_insensitive(self):
        """Test that keyword matching is case-insensitive."""
        filter = BodyKeywordFilter(["URGENT"])
        email = create_test_email(body="this is urgent")
        assert filter.matches(email)
    
    def test_multiple_keywords_all_required(self):
        """Test that all keywords must be present (AND logic)."""
        filter = BodyKeywordFilter(["urgent", "review"])
        
        # Both keywords present
        email1 = create_test_email(body="Please urgent action and review this")
        assert filter.matches(email1)
        
        # Only one keyword present
        email2 = create_test_email(body="Please urgent action")
        assert not filter.matches(email2)

class TestCompositeFilter:
    """Tests for CompositeFilter."""
    
    def test_empty_filter_list(self):
        """Test composite filter with no sub-filters."""
        filter = CompositeFilter([])
        email = create_test_email()
        assert filter.matches(email)
    
    def test_single_filter(self):
        """Test composite filter with one sub-filter."""
        sender_filter = SenderFilter(["sender@example.com"])
        composite = CompositeFilter([sender_filter])
        
        email1 = create_test_email(sender="sender@example.com")
        assert composite.matches(email1)
        
        email2 = create_test_email(sender="other@example.com")
        assert not composite.matches(email2)
    
    def test_multiple_filters_all_match(self):
        """Test composite filter where all sub-filters match."""
        date_filter = DateRangeFilter(start_date=datetime.now() - timedelta(days=7))
        sender_filter = SenderFilter(["sender@example.com"])
        subject_filter = SubjectKeywordFilter(["invoice"])
        
        composite = CompositeFilter([date_filter, sender_filter, subject_filter])
        
        email = create_test_email(
            sender="sender@example.com",
            subject="Monthly Invoice",
            days_offset=-3
        )
        assert composite.matches(email)
    
    def test_multiple_filters_one_fails(self):
        """Test composite filter where one sub-filter fails."""
        sender_filter = SenderFilter(["sender@example.com"])
        subject_filter = SubjectKeywordFilter(["invoice"])
        
        composite = CompositeFilter([sender_filter, subject_filter])
        
        # Wrong subject
        email = create_test_email(
            sender="sender@example.com",
            subject="Monthly Report"
        )
        assert not composite.matches(email)
