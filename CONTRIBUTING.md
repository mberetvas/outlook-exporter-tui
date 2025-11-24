# Contributing to Outlook Exporter

Thank you for your interest in contributing to the Outlook Exporter project! This guide will help you set up your development environment and understand our development workflow.

## Table of Contents

- [Development Setup](#development-setup)
- [Code Quality Standards](#code-quality-standards)
- [Testing](#testing)
- [Architecture Overview](#architecture-overview)
- [Making Changes](#making-changes)
- [Pull Request Process](#pull-request-process)
- [Troubleshooting](#troubleshooting)

## Development Setup

### Prerequisites

- **Python 3.12+** - Required for modern type hints
- **Windows OS** - Required for Outlook COM automation (for production use)
- **Git** - For version control
- **uv** (recommended) or **pip** - For dependency management

### Initial Setup

```bash
# Clone the repository
git clone https://github.com/mberetvas/outlook-exporter-tui.git
cd outlook-exporter-tui

# Option 1: Using uv (recommended)
uv sync --all-groups

# Option 2: Using pip
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -e ".[dev]"
```

### Install Pre-commit Hooks

We use pre-commit hooks to ensure code quality:

```bash
# Install pre-commit hooks
pre-commit install

# Run hooks manually on all files
pre-commit run --all-files
```

## Code Quality Standards

We use [Ruff](https://github.com/astral-sh/ruff) for linting and formatting, which replaces multiple tools (flake8, black, isort, pyupgrade) with a single fast tool.

### Running Code Quality Checks

```bash
# Check code style (reports issues)
ruff check .

# Fix auto-fixable issues
ruff check --fix .

# Check formatting
ruff format --check .

# Apply formatting
ruff format .
```

### Configuration

All Ruff configuration is in `pyproject.toml`. Key standards:

- **Line length**: 120 characters
- **Python version**: 3.12+
- **Quotes**: Double quotes for strings
- **Imports**: Sorted with isort-compatible style
- **Type hints**: Required for all public functions

### Pre-commit Hooks

Pre-commit hooks run automatically on `git commit`:

- Ruff linting and formatting
- Trailing whitespace removal
- End-of-file fixer
- YAML/JSON/TOML syntax checks
- Debug statement detection
- Large file prevention

## Testing

### Running Tests

```bash
# Run all tests
pytest

# Run with coverage report
pytest --cov=src/outlook_exporter --cov-report=html

# Run specific test file
pytest tests/unit/test_filters.py

# Run with verbose output
pytest -v

# Run tests matching a pattern
pytest -k "test_date_filter"
```

### Test Structure

```
tests/
├── unit/              # Unit tests (no external dependencies)
│   ├── test_models.py
│   ├── test_filters.py
│   └── test_path_utils.py
├── integration/       # Integration tests (Outlook required)
└── fixtures/          # Test fixtures and data
```

### Writing Tests

- **Unit tests**: Test pure Python logic without Outlook
- **Coverage target**: 90%+ for new code
- **Test naming**: Use descriptive names like `test_date_filter_excludes_old_emails`
- **Fixtures**: Use pytest fixtures for reusable test data

Example test:

```python
from datetime import datetime
from outlook_exporter.filters.email_filters import DateRangeFilter
from outlook_exporter.core.models import EmailMetadata

def test_date_range_filter_excludes_old():
    """Test that date filter excludes emails before start date."""
    filter = DateRangeFilter(
        start_date=datetime(2024, 1, 1),
        end_date=None
    )
    
    old_email = EmailMetadata(
        subject="Old Email",
        sender_email="test@example.com",
        sender_name="Test",
        received_time=datetime(2023, 12, 31),
        # ... other fields
    )
    
    assert not filter.matches(old_email)
```

## Architecture Overview

The project follows a layered architecture with clear separation of concerns:

### Current Structure (v0.2.0)

```
outlook-exporter-tui/
├── main.py                    # LEGACY: Monolithic script (still in use)
├── tui/                       # LEGACY: TUI using main.py
│   └── app.py
└── src/outlook_exporter/      # NEW: Refactored architecture
    ├── core/                  # Domain models and business logic
    │   ├── models.py          # Data classes (EmailMetadata, ExportConfig, etc.)
    │   └── duplicates.py      # Duplicate detection
    ├── filters/               # Email filtering system
    │   ├── base.py            # Abstract filter classes
    │   └── email_filters.py   # Concrete filter implementations
    ├── outlook/               # Outlook COM interaction
    │   ├── client.py          # High-level Outlook client
    │   └── adapters.py        # COM object adapters
    ├── storage/               # File system operations
    │   └── path_utils.py      # Path sanitization and utilities
    ├── exporters/             # Export strategies (NEW in this PR!)
    │   ├── base.py            # Base exporter and factory
    │   ├── attachment.py      # Attachment exporter
    │   ├── message.py         # .msg file exporter
    │   └── markdown.py        # Markdown exporter
    └── utils/                 # Utilities
        └── exceptions.py      # Custom exceptions
```

### Architecture Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                         CLI / TUI                           │
│                    (User Interface Layer)                   │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│                    Exporter Factory                         │
│              (Creates appropriate exporter)                 │
└────────────────────┬────────────────────────────────────────┘
                     │
         ┌───────────┼───────────┬──────────────┐
         ▼           ▼           ▼              ▼
    ┌────────┐  ┌────────┐  ┌────────┐    ┌──────────┐
    │Attachment││Message │  │Markdown│    │Future    │
    │Exporter  ││Exporter│  │Exporter│    │Exporters │
    └────┬─────┘└───┬────┘  └───┬────┘    └──────────┘
         │          │           │
         └──────────┼───────────┘
                    ▼
┌─────────────────────────────────────────────────────────────┐
│                   Business Logic Layer                      │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌──────────┐   │
│  │ Filters  │  │Duplicate │  │  Path    │  │  Models  │   │
│  │          │  │ Tracker  │  │  Utils   │  │          │   │
│  └──────────┘  └──────────┘  └──────────┘  └──────────┘   │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│                  Outlook COM Layer                          │
│  ┌──────────────┐        ┌──────────────┐                  │
│  │Outlook Client│  ◄───► │   Adapters   │                  │
│  └──────────────┘        └──────────────┘                  │
└─────────────────────────────────────────────────────────────┘
```

### Design Patterns Used

1. **Strategy Pattern**: Filters and Exporters (interchangeable algorithms)
2. **Factory Pattern**: ExporterFactory (creates appropriate exporters)
3. **Adapter Pattern**: Outlook adapters (wraps COM objects)
4. **Composite Pattern**: CompositeFilter (combines filters)
5. **Context Manager**: OutlookClient (resource management)

### Migration Status

The refactoring is partially complete:

- ✅ **Complete**: Core models, filters, outlook client, storage utils, exporters
- ⚠️ **In Progress**: Migrating main.py to use new exporters
- ❌ **Pending**: Updating TUI to import from `outlook_exporter` package

## Making Changes

### Workflow

1. **Create a branch** from `main`:
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make your changes** following the code quality standards

3. **Write tests** for new functionality

4. **Run quality checks**:
   ```bash
   # Linting and formatting
   ruff check --fix .
   ruff format .
   
   # Tests
   pytest
   ```

5. **Commit your changes**:
   ```bash
   git add .
   git commit -m "feat: add new feature"
   ```
   
   Use [Conventional Commits](https://www.conventionalcommits.org/):
   - `feat:` - New features
   - `fix:` - Bug fixes
   - `docs:` - Documentation changes
   - `style:` - Code style changes
   - `refactor:` - Code refactoring
   - `test:` - Test changes
   - `chore:` - Build/tooling changes

6. **Push and create PR**:
   ```bash
   git push origin feature/your-feature-name
   ```

### Code Style Guidelines

#### Type Hints

Always use type hints for function signatures:

```python
def process_email(
    email: EmailMetadata,
    config: ExportConfig
) -> Optional[Path]:
    """Process an email and return the output path."""
    pass
```

#### Docstrings

Use Google-style docstrings:

```python
def export_attachment(attachment, folder: Path) -> bool:
    """Export a single attachment to the specified folder.
    
    Args:
        attachment: Outlook attachment COM object
        folder: Destination folder path
        
    Returns:
        True if export succeeded, False otherwise
        
    Raises:
        FileSystemError: If folder cannot be created
    """
    pass
```

#### Error Handling

Use custom exceptions from `utils/exceptions.py`:

```python
from outlook_exporter.utils.exceptions import FileSystemError

def create_folder(path: Path) -> None:
    """Create a folder with error handling."""
    try:
        path.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        raise FileSystemError(f"Failed to create folder {path}: {e}") from e
```

## Pull Request Process

1. **Update documentation** if you changed APIs or added features
2. **Ensure all tests pass** and coverage is maintained
3. **Update CHANGELOG.md** with your changes
4. **Request review** from maintainers
5. **Address feedback** and update PR as needed

### PR Checklist

- [ ] Code follows style guidelines (Ruff passes)
- [ ] Tests added/updated and passing
- [ ] Documentation updated
- [ ] Changelog updated
- [ ] No breaking changes (or clearly documented)
- [ ] Commits follow conventional commits format

## Troubleshooting

### Common Development Issues

#### "pywin32 not found" on Linux/Mac

The project requires Windows for Outlook COM, but you can still develop and test non-COM code:

```bash
# Install without pywin32
pip install -e . --no-deps
pip install tqdm textual rich markdownify markitdown pytest
```

#### Pre-commit hooks failing

```bash
# Update hooks to latest version
pre-commit autoupdate

# Clear cache if needed
pre-commit clean

# Run manually
pre-commit run --all-files
```

#### Tests failing on import

Make sure the package is installed in editable mode:

```bash
pip install -e .
```

#### Ruff configuration issues

Configuration is in `pyproject.toml` under `[tool.ruff]`. To check what rules are enabled:

```bash
ruff check --show-settings
```

### Getting Help

- **Issues**: Open an issue on GitHub
- **Discussions**: Use GitHub Discussions for questions
- **Code Review**: Request review from maintainers

## Development Tools

### Recommended VSCode Extensions

- **Python** (ms-python.python)
- **Ruff** (charliermarsh.ruff)
- **Even Better TOML** (tamasfe.even-better-toml)

### Recommended VSCode Settings

```json
{
  "editor.formatOnSave": true,
  "editor.codeActionsOnSave": {
    "source.fixAll": true
  },
  "[python]": {
    "editor.defaultFormatter": "charliermarsh.ruff",
    "editor.formatOnSave": true
  },
  "python.testing.pytestEnabled": true,
  "python.testing.unittestEnabled": false
}
```

## Release Process

Releases are managed by maintainers:

1. Update version in `pyproject.toml` and `__init__.py`
2. Update CHANGELOG.md
3. Create release tag: `git tag v0.3.0`
4. Push tag: `git push origin v0.3.0`
5. GitHub Actions builds and publishes package

---

## Questions?

If you have questions not covered here, please:
- Check existing [GitHub Issues](https://github.com/mberetvas/outlook-exporter-tui/issues)
- Open a new issue with the `question` label
- Start a [GitHub Discussion](https://github.com/mberetvas/outlook-exporter-tui/discussions)

Thank you for contributing! 🎉
