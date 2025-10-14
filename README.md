# Outlook Attachments Export

Advanced script to export email attachments from Outlook with rich filtering, duplicate detection, structured folders, and a dry-run mode.

## Features

* Filter by date range, sender email, subject keywords, body keywords.
* Require presence / absence of attachments.
* Structured folder hierarchy: `sender/subject/date/`.
* Hash-based duplicate detection; duplicates stored under `duplicates/` subfolder within each message folder.
* Skip inline images by default (enable with `--include-inline`).
* Dry-run mode (`--dry-run`) to preview actions.
* Progress bar (requires `tqdm` installed).
* Logging with configurable verbosity (`-v`, `-vv`, `-q`) and optional log file.
* Pagination / batching for large inboxes.
* Open output folder after completion (`--open-folder`).

## Installation

Requires Python 3.12+ and Outlook desktop.

Install dependencies:

```powershell
pip install .
```

Or if using `uv`:

```powershell
uv sync
```

## Usage Examples

Basic export of all attachments from Inbox:

```powershell
python main.py --output C:\Attachments
```

Filter by date range and sender, only messages with attachments:

```powershell
python main.py --output C:\Attachments --start-date 2025-10-01 --end-date 2025-10-14 --sender someone@example.com --with-attachments
```

Subject and body keyword filtering (all keywords must appear):

```powershell
python main.py --output C:\Attachments --subject-keyword invoice --body-keyword urgent
```

Specify a subfolder of Inbox:

```powershell
python main.py --output C:\Attachments --folder Inbox/SubFolder
```

Dry run (no files written):

```powershell
python main.py --output C:\Attachments --dry-run --sender boss@example.com --subject-keyword Q4
```

Increase verbosity:

```powershell
python main.py --output C:\Attachments -v
python main.py --output C:\Attachments -vv  # debug level
```

Quiet mode (only warnings/errors):

```powershell
python main.py --output C:\Attachments -q
```

Limit processed messages (pagination) and change batch size:

```powershell
python main.py --output C:\Attachments --limit 500 --batch-size 100
```

Save logs to file:

```powershell
python main.py --output C:\Attachments --log-file C:\Attachments\export.log
```

Open output folder automatically:

```powershell
python main.py --output C:\Attachments --open-folder
```

Include inline images:

```powershell
python main.py --output C:\Attachments --include-inline
```

## Duplicate Handling

Duplicates are identified using a cryptographic hash (`--hash-algorithm`, default `sha256`). The first occurrence is saved in the normal folder path; subsequent duplicates are placed under `duplicates/` within that message's folder. Existing filename clashes are resolved by appending an incrementing suffix.

## Dry Run Mode

When `--dry-run` is specified, the script logs intended save paths but does not create or modify files.

## Exit Codes

* `0` – Success
* `1` – Unhandled error
* `2` – Invalid argument combination

## Notes

* Outlook COM interactions cannot easily be unit tested; logic functions are isolated for potential testing.
* The script skips inline images heuristically unless `--include-inline` is passed.
* Folder name sanitization replaces characters invalid on Windows with `_` and truncates subjects to a configurable length.

## Future Improvements

* Add unit tests for filtering logic.
* Add option to export a JSON/CSV report of saved attachments.
* Support selecting other stores/mailboxes.

