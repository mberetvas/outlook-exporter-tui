# Outlook Attachments Export Tool

*Bulk export email attachments from Outlook with filtering and duplicate detection*

[![License](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)](LICENSE)
![Python version](https://img.shields.io/badge/Python->=3.12-3c873a?style=flat-square)
![Platform](https://img.shields.io/badge/Platform-Windows-blue?style=flat-square)

⭐ If you like this project, star it on GitHub!

[Overview](#overview) • [Features](#features) • [Getting Started](#getting-started) • [Usage](#usage) • [Examples](#examples)

---

A Windows-based CLI and TUI tool for efficiently exporting email attachments from Microsoft Outlook desktop. Features advanced filtering options, intelligent duplicate detection, and organized folder structures to help you manage large volumes of email attachments.

## Overview

This tool provides both a command-line interface (CLI) and a terminal user interface (TUI) for exporting attachments from your Outlook mailbox. It uses Outlook's MAPI interface through COM automation to access your emails and extract attachments based on sophisticated filtering criteria.

**Key capabilities:**
- Filter emails by date range, sender, subject, or body content
- Export attachments to structured folders organized by sender, subject, and date
- Hash-based duplicate detection to avoid saving the same file multiple times
- Export entire emails as `.msg` files
- Interactive TUI for easier configuration
- Dry-run mode to preview actions before execution

> [!NOTE]
> This tool requires Microsoft Outlook desktop (Windows) and uses the Win32 COM interface. It will not work with Outlook Web Access or on non-Windows platforms.

## Features

- **Advanced Filtering** - Filter messages by date range, sender email, subject keywords, body content, and attachment presence
- **Organized Export Structure** - Automatically creates hierarchical folders: `{sender}/{subject}/{date}/` for easy navigation
- **Duplicate Detection** - SHA256 hash-based detection saves duplicates to a separate subfolder
- **Interactive TUI** - User-friendly terminal interface built with Textual for easy configuration
- **Powerful CLI** - Full-featured command-line interface for automation and scripting
- **Dry-Run Mode** - Preview what will be exported without actually saving files
- **Batch Processing** - Handle large mailboxes efficiently with configurable batch sizes
- **Flexible Options** - Include/exclude inline attachments, export full emails, custom folder paths
- **Robust Error Handling** - Continues processing on errors with detailed logging
- **Progress Tracking** - Real-time progress bars for long-running operations

## Getting Started

### Prerequisites

- **Windows OS** - Required for Outlook COM automation
- **Python 3.12+** - Modern Python with type hints support
- **Microsoft Outlook Desktop** - Must be installed and configured with at least one email account
- **uv** (recommended) - Fast Python package installer and environment manager

### Installation

#### Using uv (recommended)

```powershell
# Install uv if you don't have it
# Follow instructions at https://github.com/astral-sh/uv

# Clone the repository
git clone https://github.com/mberetvas/outlook-exporter-tui.git
cd outlook-exporter-tui

# Install dependencies
uv sync
```

#### Using pip

```powershell
# Clone the repository
git clone https://github.com/mberetvas/outlook-exporter-tui.git
cd outlook-exporter-tui

# Create a virtual environment (optional but recommended)
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# Install dependencies
pip install -e .
```

## Usage

### Terminal User Interface (TUI)

The easiest way to use the tool is through the interactive TUI:

```powershell
# Using uv
uv run outlook-tui

# Or if installed with pip
outlook-tui
```

The TUI provides a form-based interface where you can:
- Set the output directory
- Configure date ranges and filter criteria
- Choose attachment options
- View real-time logs during export
- Run in dry-run mode to preview results

### Command-Line Interface (CLI)

For automation and scripting, use the CLI directly:

```powershell
# Using uv
uv run python main.py [OPTIONS]

# Or if installed with pip
python main.py [OPTIONS]
```

#### Basic CLI Options

```
Required:
  -o, --output PATH          Base output directory for attachments

Filtering:
  --start-date YYYY-MM-DD    Filter: start date
  --end-date YYYY-MM-DD      Filter: end date
  --sender EMAIL             Filter: sender email (can repeat)
  --subject-keyword TEXT     Filter: subject contains keyword (can repeat)
  --body-keyword TEXT        Filter: body contains keyword (can repeat)
  --with-attachments         Filter: only messages with attachments
  --without-attachments      Filter: only messages without attachments

Export Options:
  --folder PATH              Custom Outlook folder path (e.g. 'Inbox/SubFolder')
  --include-inline           Include inline attachments (images in signatures)
  --export-mail              Export entire email as .msg file
  --duplicates-subfolder     Name of subfolder for duplicates (default: duplicates)
  --hash-algorithm ALGO      Hash algorithm for duplicate detection (default: sha256)

Performance:
  --limit N                  Max number of messages to process
  --batch-size N             Batch size for iterating messages (default: 200)

Execution:
  --dry-run                  Preview actions without saving files
  --open-folder              Open output folder after completion

Logging:
  --log-file PATH            Path to log file
  -v, --verbose              Increase verbosity (can repeat: -v, -vv)
  -q, --quiet                Quiet mode: only warnings/errors
```

## Examples

### Basic Export

Export all attachments from inbox to a folder:

```powershell
uv run python main.py --output ./exports
```

### Filter by Date Range

Export attachments from emails received in January 2024:

```powershell
uv run python main.py --output ./exports --start-date 2024-01-01 --end-date 2024-01-31
```

### Filter by Sender

Export attachments from specific senders:

```powershell
uv run python main.py --output ./exports --sender john@example.com --sender jane@example.com
```

### Filter by Keywords

Export attachments from emails containing specific keywords:

```powershell
uv run python main.py --output ./exports --subject-keyword "invoice" --subject-keyword "receipt"
```

### Export from Specific Folder

Export from a subfolder in your mailbox:

```powershell
uv run python main.py --output ./exports --folder "Inbox/Projects/ProjectX"
```

### Dry Run

Preview what would be exported without actually saving files:

```powershell
uv run python main.py --output ./exports --dry-run --start-date 2024-01-01
```

### Export Full Emails

Export entire emails as .msg files instead of just attachments:

```powershell
uv run python main.py --output ./exports --export-mail --sender boss@company.com
```

### Complex Filtering

Combine multiple filters for precise control:

```powershell
uv run python main.py --output ./exports `
  --start-date 2024-01-01 `
  --end-date 2024-03-31 `
  --sender vendor@supplier.com `
  --subject-keyword "purchase order" `
  --with-attachments `
  --include-inline `
  --log-file ./export.log `
  -v
```

## Output Structure

The tool creates a hierarchical folder structure for exported attachments:

```
output/
├── sender1@example.com/
│   ├── Email Subject 1/
│   │   ├── 2024-01-15/
│   │   │   ├── attachment1.pdf
│   │   │   ├── attachment2.xlsx
│   │   │   └── duplicates/
│   │   │       └── attachment1.pdf  # Duplicate file
│   │   └── 2024-01-16/
│   │       └── document.docx
│   └── Another Subject/
│       └── 2024-01-20/
│           └── file.zip
└── sender2@example.com/
    └── Subject/
        └── 2024-02-01/
            └── attachment.pdf
```

**Path sanitization:**
- Invalid Windows characters (`\/:*?"<>|`) are replaced with underscores
- Paths are truncated to avoid Windows path length limits (~260 characters)
- Fallback naming strategies ensure files are always saved

## Development

### Running Tests

```powershell
# Run all tests
uv run pytest

# Run with coverage
uv run pytest --cov=. --cov-report=html

# Run specific test file
uv run pytest tests/test_filter_logic.py
```

> [!NOTE]
> Due to the nature of COM automation, only pure Python functions (like filtering logic) are unit tested. Outlook interaction functions are marked `# pragma: no cover` and require manual testing.

### Project Structure

```
outlook-exporter-tui/
├── main.py                  # CLI entry point and core export logic
├── tui/
│   ├── __init__.py
│   ├── app.py              # Textual TUI application
│   └── app.css             # TUI styling
├── tests/
│   └── test_filter_logic.py
├── pyproject.toml          # Project configuration
└── README.md
```

## Troubleshooting

### "pywin32 is required" Error

Make sure pywin32 is installed:
```powershell
uv sync
# or
pip install pywin32
```

### Outlook Not Responding

If Outlook hangs during export:
- Close Outlook completely and restart
- Reduce the `--batch-size` parameter
- Use `--limit` to process fewer messages at once

### Path Too Long Errors

Windows has a 260-character path limit. The tool includes automatic path truncation, but if you still encounter issues:
- Use a shorter output directory path (e.g., `C:\exports` instead of a deep path)
- The tool will automatically shorten folder names and use hash-based fallbacks

### No Attachments Exported

Check your filters:
- Use `--dry-run` to see what would be exported
- Enable verbose logging with `-vv` to see detailed filtering information
- Verify your date format is `YYYY-MM-DD`
- Check that sender emails match exactly (case-insensitive)

### Inline Attachments

By default, inline attachments (like signature images) are skipped. Use `--include-inline` to export them.

## Technical Notes

- **COM Collections**: Outlook COM collections are 1-based, not 0-based like Python lists
- **Safe Property Access**: All COM property accesses use the `safe_get()` wrapper to handle "Operation aborted" errors
- **Hash Algorithm**: Uses SHA256 by default for duplicate detection (configurable via `--hash-algorithm`)
- **Windows-Specific**: Uses `os.startfile()` for opening folders (Windows only)

## Frequently Asked Questions

**Q: Can this work with Office 365/Outlook.com?**  
A: Only if you have Outlook desktop installed and configured. It requires the local MAPI interface.

**Q: Will this work on Mac or Linux?**  
A: No, this tool requires Windows and the Win32 COM interface to Outlook.

**Q: Can I schedule automatic exports?**  
A: Yes! Use Windows Task Scheduler to run the CLI command at specified intervals.

**Q: How are duplicates determined?**  
A: Files are considered duplicates if they have the same SHA256 hash (content-based, not filename-based).

**Q: Can I export from multiple folders at once?**  
A: Not in a single command, but you can run the tool multiple times with different `--folder` parameters.

## Resources

- [pywin32 Documentation](https://github.com/mhammond/pywin32)
- [Textual Documentation](https://textual.textualize.io/)
- [Outlook Object Model Reference](https://learn.microsoft.com/office/vba/api/overview/outlook/object-model)

## Acknowledgments

Built with:
- [pywin32](https://github.com/mhammond/pywin32) - Python for Windows Extensions
- [Textual](https://textual.textualize.io/) - Modern TUI framework
- [tqdm](https://github.com/tqdm/tqdm) - Progress bars

---

Made with ❤️ for anyone drowning in email attachments

