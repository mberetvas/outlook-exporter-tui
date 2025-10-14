"""Outlook Attachments Export Tool

Features:
 - Filter emails by date range, sender, subject keywords, body content keywords, presence/absence of attachments.
 - Save attachments to structured subfolders: sender/subject/date.
 - Hash-based duplicate detection; duplicates saved under a dedicated duplicates/ subfolder.
 - CLI arguments for configuration, including log file path and verbosity (-v / -q).
 - Dry-run mode lists intended actions without writing files.
 - Optional opening of target folder after completion.
 - Progress bar and robust logging with error continuation.
 - Ability to choose default Inbox or a specific folder path.
 - Pagination / batching to avoid loading huge collections at once.

Requires: Outlook desktop (MAPI) and pywin32.
"""

from __future__ import annotations

import argparse
import hashlib
import logging
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Optional, Sequence

try:
    import win32com.client  # type: ignore
    import pythoncom
except ImportError:  # pragma: no cover
    print("pywin32 is required to run this script. Install via 'pip install pywin32'.", file=sys.stderr)
    raise

try:
    from tqdm import tqdm
except ImportError:  # pragma: no cover
    tqdm = None  # fallback: no progress bar

try:
    from markdownify import markdownify as md
except ImportError:  # pragma: no cover
    md = None  # fallback: no markdown conversion


OUTLOOK_INBOX_ID = 6  # Default Inbox folder constant


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export Outlook attachments with filtering and duplicate handling")
    parser.add_argument("--output", "-o", type=Path, required=True, help="Base output directory for attachments")
    parser.add_argument("--start-date", type=str, help="Filter: start date (YYYY-MM-DD)")
    parser.add_argument("--end-date", type=str, help="Filter: end date (YYYY-MM-DD)")
    parser.add_argument("--sender", action="append", help="Filter: sender email (can repeat)")
    parser.add_argument("--subject-keyword", action="append", help="Filter: subject contains keyword (can repeat)")
    parser.add_argument("--body-keyword", action="append", help="Filter: body contains keyword (can repeat)")
    parser.add_argument("--with-attachments", action="store_true", help="Filter: only messages that have attachments")
    parser.add_argument("--without-attachments", action="store_true", help="Filter: only messages without attachments")
    parser.add_argument("--folder", type=str, help="Custom Outlook folder path (e.g. 'Inbox/SubFolder')")
    parser.add_argument("--limit", type=int, help="Max number of messages to process (pagination)")
    parser.add_argument("--batch-size", type=int, default=200, help="Batch size for iterating messages")
    parser.add_argument("--dry-run", action="store_true", help="Dry run: list actions without saving")
    parser.add_argument("--open-folder", action="store_true", help="Open output folder after completion")
    parser.add_argument("--log-file", type=Path, help="Path to log file")
    parser.add_argument("--duplicates-subfolder", default="duplicates", help="Name of subfolder for duplicate attachments")
    parser.add_argument("--hash-algorithm", default="sha256", choices=hashlib.algorithms_available, help="Hash algorithm for duplicate detection")
    parser.add_argument("--include-inline", action="store_true", help="Include inline attachments (images in signatures). Default: skip them")
    parser.add_argument("--export-mail", action="store_true", help="Export the entire email as a .msg file")
    parser.add_argument("--export-markdown", action="store_true", help="Export the entire email as a .md (markdown) file")
    parser.add_argument("--subject-sanitize-length", type=int, default=80, help="Max length of subject used in folder names")
    parser.add_argument("-v", "--verbose", action="count", default=0, help="Increase verbosity (can repeat)")
    parser.add_argument("-q", "--quiet", action="store_true", help="Quiet mode: only warnings/errors")
    return parser.parse_args(argv)


def setup_logging(args: argparse.Namespace) -> None:
    level = logging.INFO
    if args.quiet:
        level = logging.WARNING
    elif args.verbose >= 2:
        level = logging.DEBUG
    elif args.verbose == 1:
        level = logging.INFO
    else:
        level = logging.INFO

    handlers: List[logging.Handler] = [logging.StreamHandler(sys.stdout)]
    if args.log_file:
        log_file_parent = args.log_file.parent
        log_file_parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(args.log_file, encoding="utf-8"))

    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)-8s %(message)s",
        handlers=handlers,
    )


def get_outlook_folder(namespace, folder_path: Optional[str]):  # pragma: no cover (Outlook interaction)
    if not folder_path:
        return namespace.GetDefaultFolder(OUTLOOK_INBOX_ID)
    parts = folder_path.split('/') if '/' in folder_path else folder_path.split('\\')
    folder = namespace.GetDefaultFolder(OUTLOOK_INBOX_ID)
    for part in parts:
        if not part or part.lower() == 'inbox':
            continue
        try:
            folder = folder.Folders[part]
        except Exception:
            raise ValueError(f"Folder part not found: {part}")
    return folder


def parse_date(date_str: Optional[str]) -> Optional[datetime]:
    if not date_str:
        return None
    return datetime.strptime(date_str, "%Y-%m-%d")


def message_matches(message, args: argparse.Namespace) -> bool:  # pragma: no cover (Outlook interaction)
    try:
        # Date filter
        start = parse_date(args.start_date)
        end = parse_date(args.end_date)
        if start or end:
            msg_date = message.ReceivedTime  # COM datetime
            # Convert to python datetime if needed
            if start and msg_date < start:
                return False
            if end and msg_date > end:
                return False

        # Sender filter
        if args.sender:
            sender_email = getattr(message, 'SenderEmailAddress', '') or ''
            if not any(sender_email.lower() == s.lower() for s in args.sender):
                return False

        # Subject filter
        if args.subject_keyword:
            subject = (getattr(message, 'Subject', '') or '')
            if not all(kw.lower() in subject.lower() for kw in args.subject_keyword):
                return False

        # Body filter
        if args.body_keyword:
            body = (getattr(message, 'Body', '') or '')
            if not all(kw.lower() in body.lower() for kw in args.body_keyword):
                return False

        # Attachment presence filter
        atts = getattr(message, 'Attachments', [])
        # COM collection has .Count; fallback if python list
        if hasattr(atts, 'Count'):
            has_attachments = int(atts.Count) > 0  # type: ignore[attr-defined]
        else:
            has_attachments = len(atts) > 0
        if args.with_attachments and not has_attachments:
            return False
        if args.without_attachments and has_attachments:
            return False

        return True
    except Exception as e:
        logging.debug("Filter evaluation error: %s", e)
        return False


def sanitize_for_fs(value: str, max_length: int) -> str:
    value = value.strip()
    # Remove/replace invalid Windows path characters
    value = re.sub(r"[\\/:*?\"<>|]", "_", value)
    # Remove trailing spaces and dots (Windows restriction)
    value = value.rstrip('. ')
    # Truncate to max length
    if len(value) > max_length:
        value = value[:max_length].rstrip('_. ')
    # Ensure not empty
    if not value:
        value = "_"
    return value


def compute_hash(path: Path, algo: str) -> str:
    h = hashlib.new(algo)
    with path.open('rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    return h.hexdigest()


def safe_get(obj, attr: str, default=None):  # pragma: no cover
    """Safely get COM attribute, returning default if access fails (e.g. Operation aborted)."""
    try:
        return getattr(obj, attr)
    except Exception as e:
        logging.debug("safe_get failed for %s.%s: %s", type(obj).__name__, attr, e)
        return default


def save_attachment(base_dir: Path, message, attachment, args: argparse.Namespace, seen_hashes: dict[str, List[Path]]) -> Optional[Path]:  # pragma: no cover
    # Skip inline attachments unless requested
    if not args.include_inline:
        try:
            if getattr(attachment, 'Position', 0) == 0 and getattr(attachment, 'Type', None) == 5:  # 5: olByValue maybe inline
                # Heuristic skip: inline images often have Position = 0
                return None
        except Exception:
            pass

    subject = sanitize_for_fs(safe_get(message, 'Subject', '') or '', args.subject_sanitize_length)
    sender = sanitize_for_fs(safe_get(message, 'SenderEmailAddress', '') or 'unknown', 120)
    received = safe_get(message, 'ReceivedTime', datetime.now())
    if hasattr(received, 'Format'):  # COM date
        # Some COM objects have Format method; rely on python cast when stringifying
        try:
            received = datetime.strptime(str(received), '%m/%d/%y %H:%M:%S')
        except Exception:
            received = datetime.now()

    date_folder = received.strftime('%Y-%m-%d')
    target_dir = base_dir / sender / subject / date_folder
    
    # Check path length and truncate if needed (Windows limit ~260 chars)
    if len(str(target_dir)) > 200:  # Leave room for filename
        # Try shorter subject first
        short_subject = sanitize_for_fs(safe_get(message, 'Subject', '') or '', 50)
        target_dir = base_dir / sender / short_subject / date_folder
        if len(str(target_dir)) > 200:
            # Fallback: use hash of original subject
            subject_hash = hashlib.md5((safe_get(message, 'Subject', '') or '').encode('utf-8')).hexdigest()[:8]
            target_dir = base_dir / sender / f"subj_{subject_hash}" / date_folder
    
    try:
        target_dir.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        logging.warning("Failed to create directory %s: %s. Using fallback.", target_dir, e)
        # Ultimate fallback: flat structure with timestamp
        fallback_name = f"{sender}_{date_folder}_{received.strftime('%H%M%S')}"
        target_dir = base_dir / sanitize_for_fs(fallback_name, 100)
        target_dir.mkdir(parents=True, exist_ok=True)

    filename = sanitize_for_fs(safe_get(attachment, 'FileName', 'attachment.bin') or 'attachment.bin', 255)
    dest_path = target_dir / filename

    if args.dry_run:
        logging.info("[DRY-RUN] Would save attachment: %s", dest_path)
        return dest_path

    # Write to a temp file first to compute hash reliably
    temp_path = dest_path.with_suffix(dest_path.suffix + '.tmp')
    try:
        attachment.SaveAsFile(str(temp_path))
    except Exception as e:
        logging.warning("Failed to save attachment temp file %s: %s", temp_path, e)
        return None

    file_hash = compute_hash(temp_path, args.hash_algorithm)
    duplicate = file_hash in seen_hashes

    if duplicate:
        duplicates_dir = target_dir / args.duplicates_subfolder
        duplicates_dir.mkdir(parents=True, exist_ok=True)
        dup_dest = duplicates_dir / filename
        # Rename if already exists
        counter = 1
        while dup_dest.exists():
            dup_dest = duplicates_dir / f"{dup_dest.stem}_{counter}{dup_dest.suffix}"
            counter += 1
        temp_path.rename(dup_dest)
        logging.info("Duplicate detected (hash=%s). Saved to %s", file_hash, dup_dest)
        seen_hashes[file_hash].append(dup_dest)
        return dup_dest
    else:
        # Rename if existing
        final_path = dest_path
        counter = 1
        while final_path.exists():
            final_path = dest_path.with_name(f"{dest_path.stem}_{counter}{dest_path.suffix}")
            counter += 1
        temp_path.rename(final_path)
        logging.debug("Saved attachment %s (hash=%s)", final_path, file_hash)
        seen_hashes[file_hash] = [final_path]
        return final_path


def iterate_messages(folder, limit: Optional[int], batch_size: int) -> Iterable:  # pragma: no cover
    items = folder.Items
    total = items.Count
    processed = 0
    
    # Use COM safe iteration with Item() method and error handling
    while processed < total:
        if limit and processed >= limit:
            break
        end_index = min(processed + batch_size, total)
        for i in range(processed + 1, end_index + 1):
            try:
                # Use Item method for 1-based COM access
                if hasattr(items, 'Item'):
                    yield items.Item(i)  # type: ignore[attr-defined]
                else:
                    # Fallback to direct indexing
                    yield items[i]
            except (IndexError, Exception) as e:
                logging.debug("Failed to access message %d/%d: %s", i, total, e)
                # Skip this message and continue
                continue
        processed = end_index


def save_mail(base_dir: Path, message, args: argparse.Namespace) -> Optional[Path]:
    subject = sanitize_for_fs(
        safe_get(message, "Subject", "") or "", args.subject_sanitize_length
    )
    sender = sanitize_for_fs(
        safe_get(message, "SenderEmailAddress", "") or "unknown", 120
    )
    received = safe_get(message, "ReceivedTime", datetime.now())
    if hasattr(received, "Format"):  # COM date
        try:
            received = datetime.strptime(str(received), "%m/%d/%y %H:%M:%S")
        except Exception:
            received = datetime.now()

    date_folder = received.strftime("%Y-%m-%d")
    target_dir = base_dir / sender / subject / date_folder
    target_dir.mkdir(parents=True, exist_ok=True)

    filename = sanitize_for_fs(subject, 255) + ".msg"
    dest_path = target_dir / filename

    if args.dry_run:
        logging.info("[DRY-RUN] Would save email: %s", dest_path)
        return dest_path

    try:
        message.SaveAs(str(dest_path))
        logging.info("Saved email: %s", dest_path)
        return dest_path
    except Exception as e:
        logging.warning("Failed to save email %s: %s", dest_path, e)
        return None

def format_size(size_bytes: int) -> str:
    """Format file size in human-readable format."""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"

def save_markdown(base_dir: Path, message, args: argparse.Namespace, saved_attachments: List[Path]) -> Optional[Path]:  # pragma: no cover
    """Save email as markdown file with YAML frontmatter."""
    if md is None:
        logging.warning("markdownify not available. Install with 'pip install markdownify'")
        return None
    
    # Extract email metadata
    subject_raw = safe_get(message, "Subject", "") or "No Subject"
    subject = sanitize_for_fs(subject_raw, args.subject_sanitize_length)
    sender_email = safe_get(message, "SenderEmailAddress", "") or "unknown"
    sender_name = safe_get(message, "SenderName", "") or sender_email
    sender = sanitize_for_fs(sender_email, 120)
    
    # Recipients
    to_recipients = safe_get(message, "To", "") or ""
    cc_recipients = safe_get(message, "CC", "") or ""
    
    # Dates
    received = safe_get(message, "ReceivedTime", datetime.now())
    sent = safe_get(message, "SentOn", None)
    
    if hasattr(received, "Format"):  # COM date
        try:
            received = datetime.strptime(str(received), "%m/%d/%y %H:%M:%S")
        except Exception:
            received = datetime.now()
    
    if sent and hasattr(sent, "Format"):
        try:
            sent = datetime.strptime(str(sent), "%m/%d/%y %H:%M:%S")
        except Exception:
            sent = received
    else:
        sent = received
    
    # Create target directory
    date_folder = received.strftime("%Y-%m-%d")
    target_dir = base_dir / sender / subject / date_folder
    target_dir.mkdir(parents=True, exist_ok=True)
    
    filename = sanitize_for_fs(subject, 255) + ".md"
    dest_path = target_dir / filename
    
    if args.dry_run:
        logging.info("[DRY-RUN] Would save markdown: %s", dest_path)
        return dest_path
    
    # Get email body
    html_body = safe_get(message, "HTMLBody", "")
    plain_body = safe_get(message, "Body", "")
    rtf_body = None
    
    if not html_body and not plain_body:
        try:
            # Fallback to RTFBody if others are empty
            rtf_body_bytes = safe_get(message, "RTFBody")
            if rtf_body_bytes:
                # This is a simplified conversion, might not handle all RTF features
                # For a more robust solution, a dedicated RTF-to-text library would be better
                rtf_body = rtf_body_bytes.decode('ascii', errors='ignore')
                # Basic parsing to remove RTF control words
                rtf_body = re.sub(r'{\\[^{}]+}|\\s\d+|\\pard|\\par|\n|\r', '', rtf_body)
                rtf_body = re.sub(r'\\.[a-z0-9]+', '', rtf_body).strip()
        except Exception as e:
            logging.warning("Could not process RTFBody: %s", e)

    # Convert body to markdown
    if html_body:
        try:
            body_markdown = md(html_body, heading_style="ATX", bullets="-")
        except Exception as e:
            logging.warning("Failed to convert HTML to markdown: %s. Using plain text.", e)
            body_markdown = plain_body or rtf_body or ""
    else:
        body_markdown = plain_body or rtf_body or ""

    if not body_markdown:
        logging.warning(
            "Email body is empty for subject: '%s'. "
            "This could be due to Outlook security settings or an add-in blocking access. "
            "Consider checking your Outlook COM Add-ins.",
            subject_raw
        )
    
    # Build markdown content
    markdown_lines = []
    
    # YAML frontmatter
    markdown_lines.append("---")
    markdown_lines.append(f"from: {sender_email}")
    if to_recipients:
        markdown_lines.append(f"to: {to_recipients}")
    if cc_recipients:
        markdown_lines.append(f"cc: {cc_recipients}")
    markdown_lines.append(f"subject: {subject_raw}")
    markdown_lines.append(f"date: {sent.isoformat()}")
    markdown_lines.append(f"received: {received.isoformat()}")
    markdown_lines.append("---")
    markdown_lines.append("")
    
    # Email header section
    markdown_lines.append(f"# {subject_raw}")
    markdown_lines.append("")
    markdown_lines.append(f"**From:** {sender_name} ({sender_email})  ")
    if to_recipients:
        markdown_lines.append(f"**To:** {to_recipients}  ")
    if cc_recipients:
        markdown_lines.append(f"**CC:** {cc_recipients}  ")
    markdown_lines.append(f"**Date:** {sent.strftime('%B %d, %Y %I:%M %p')}")
    markdown_lines.append("")
    markdown_lines.append("---")
    markdown_lines.append("")
    
    # Email body
    markdown_lines.append(body_markdown)
    markdown_lines.append("")
    
    # Attachments section
    if saved_attachments:
        markdown_lines.append("---")
        markdown_lines.append("")
        markdown_lines.append("## Attachments")
        markdown_lines.append("")
        for attachment_path in saved_attachments:
            try:
                size = attachment_path.stat().st_size
                size_str = format_size(size)
                # Get relative path from markdown file to attachment
                rel_path = os.path.relpath(attachment_path, dest_path.parent)
                markdown_lines.append(f"- [{attachment_path.name}]({rel_path}) ({size_str})")
            except Exception as e:
                logging.debug("Failed to get attachment info for %s: %s", attachment_path, e)
                markdown_lines.append(f"- {attachment_path.name}")
    
    # Write markdown file
    try:
        with dest_path.open('w', encoding='utf-8') as f:
            f.write('\n'.join(markdown_lines))
        logging.info("Saved markdown: %s", dest_path)
        return dest_path
    except Exception as e:
        logging.warning("Failed to save markdown %s: %s", dest_path, e)
        return None


def process_messages(args: argparse.Namespace) -> int:  # pragma: no cover
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = get_outlook_folder(outlook, args.folder)
    base_dir: Path = args.output
    base_dir.mkdir(parents=True, exist_ok=True)

    seen_hashes: dict[str, List[Path]] = {}
    saved_count = 0
    matched_messages = 0

    messages_iter = iterate_messages(folder, args.limit, args.batch_size)
    wrapper = tqdm(messages_iter, desc="Messages", unit="msg") if (tqdm and not args.quiet) else messages_iter
    for message in wrapper:
        if message_matches(message, args):
            matched_messages += 1
            if args.export_mail:
                path = save_mail(base_dir, message, args)
                if path:
                    saved_count += 1
                continue
            if args.export_markdown:
                # Save attachments first, then create markdown with references
                saved_attachments: List[Path] = []
                attachments = getattr(message, "Attachments", [])
                if hasattr(attachments, "Count"):
                    attach_count = int(attachments.Count)  # type: ignore[attr-defined]
                else:
                    attach_count = len(attachments)
                
                # Save all attachments
                for idx in range(1, attach_count + 1):
                    try:
                        if hasattr(attachments, "Item"):
                            attachment = attachments.Item(idx)  # type: ignore[attr-defined]
                        else:
                            attachment = attachments[idx - 1]
                    except (IndexError, Exception) as e:
                        logging.debug("Error accessing attachment %d: %s", idx, e)
                        continue
                    
                    path = save_attachment(base_dir, message, attachment, args, seen_hashes)
                    if path:
                        saved_attachments.append(path)
                        saved_count += 1
                
                # Save markdown with attachment references
                path = save_markdown(base_dir, message, args, saved_attachments)
                if path:
                    saved_count += 1
                continue
            
            # Default: just save attachments
            attachments = getattr(message, "Attachments", [])
            if hasattr(attachments, "Count"):
                attach_count = int(attachments.Count)  # type: ignore[attr-defined]
            else:
                attach_count = len(attachments)
            if attach_count == 0 and args.with_attachments:
                continue
            if attach_count > 0 and args.without_attachments:
                continue
            # Iterate attachments safely (COM collections are 1-based via Item())
            for idx in range(1, attach_count + 1):
                try:
                    # Use Item method if available to avoid IndexError from python enumerator
                    if hasattr(attachments, "Item"):
                        attachment = attachments.Item(idx)  # type: ignore[attr-defined]
                    else:
                        # Fallback: python sequence (0-based)
                        attachment = attachments[idx - 1]
                except IndexError:
                    logging.debug(
                        "IndexError accessing attachment %d/%d", idx, attach_count
                    )
                    continue
                except Exception as e:
                    logging.debug("Error accessing attachment %d: %s", idx, e)
                    continue
                path = save_attachment(
                    base_dir, message, attachment, args, seen_hashes
                )
                if path:
                    saved_count += 1
        else:
            continue
    logging.info("Messages matched filters: %d", matched_messages)
    logging.info("Attachments saved: %d", saved_count)
    return saved_count


def run_export(args: argparse.Namespace) -> int:
    """Runs the export process with the given arguments."""
    setup_logging(args)
    logging.info("Starting Outlook attachments export")
    if args.with_attachments and args.without_attachments:
        logging.error("Cannot specify both --with-attachments and --without-attachments")
        return 2
    if args.export_mail and args.export_markdown:
        logging.error("Cannot specify both --export-mail and --export-markdown")
        return 2
    pythoncom.CoInitialize()
    try:
        saved = process_messages(args)
        if args.open_folder:
            try:
                os.startfile(args.output)  # Windows only
            except Exception as e:
                logging.warning("Failed to open folder: %s", e)
        logging.info("Completed. Saved %d attachments", saved)
        return 0
    except Exception as e:
        logging.exception("Unhandled error: %s", e)
        return 1
    finally:
        pythoncom.CoUninitialize()


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    return run_export(args)


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
