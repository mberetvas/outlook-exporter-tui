"""Minimal tests for message_matches logic using a fake message object.

These tests do not interact with Outlook; they verify pure filtering logic
based on attributes required by message_matches.
"""

from datetime import datetime, timedelta

import argparse

import main  # type: ignore


class FakeAttachments:
    def __init__(self, count: int):
        self._count = count

    @property
    def Count(self):  # Outlook COM style property
        return self._count


def make_message(subject: str = "Subject", body: str = "Body", sender: str = "sender@example.com", days_offset: int = 0, attachments: int = 0):
    msg = argparse.Namespace()
    msg.Subject = subject
    msg.Body = body
    msg.SenderEmailAddress = sender
    msg.ReceivedTime = datetime.now() + timedelta(days=days_offset)
    msg.Attachments = FakeAttachments(attachments)
    return msg


def build_args(**overrides):
    base = {
        'start_date': None,
        'end_date': None,
        'sender': None,
        'subject_keyword': None,
        'body_keyword': None,
        'with_attachments': False,
        'without_attachments': False,
    }
    base.update(overrides)
    ns = argparse.Namespace()
    for k, v in base.items():
        setattr(ns, k, v)
    return ns


def test_subject_keyword_match():
    args = build_args(subject_keyword=["Invoice"])
    msg = make_message(subject="Quarterly Invoice")
    assert main.message_matches(msg, args) is True


def test_subject_keyword_miss():
    args = build_args(subject_keyword=["Invoice"])
    msg = make_message(subject="Report")
    assert main.message_matches(msg, args) is False


def test_body_keyword_all_required():
    args = build_args(body_keyword=["urgent", "review"])
    msg = make_message(body="Please urgent action and review this")
    assert main.message_matches(msg, args) is True


def test_body_keyword_missing_one():
    args = build_args(body_keyword=["urgent", "review"])
    msg = make_message(body="Please urgent action")
    assert main.message_matches(msg, args) is False


def test_sender_filter():
    args = build_args(sender=["sender@example.com"])  # exact match lowercased
    msg = make_message(sender="SENDER@example.com")
    assert main.message_matches(msg, args) is True


def test_date_range_excludes_old():
    today = datetime.now().strftime("%Y-%m-%d")
    args = build_args(start_date=today)
    msg = make_message(days_offset=-2)
    assert main.message_matches(msg, args) is False


def test_with_attachments_only():
    args = build_args(with_attachments=True)
    msg = make_message(attachments=2)
    assert main.message_matches(msg, args) is True


def test_without_attachments_only():
    args = build_args(without_attachments=True)
    msg = make_message(attachments=0)
    assert main.message_matches(msg, args) is True
