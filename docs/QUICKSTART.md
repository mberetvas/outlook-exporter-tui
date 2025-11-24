# Quick Start Guide - Activating the Improvements

This guide helps you activate the four critical improvements implemented in this PR.

## Overview

Four improvements were added across quality pillars:
1. **Code Quality**: Ruff linting + pre-commit hooks
2. **Architecture**: Complete exporter layer
3. **Testing**: GitHub Actions CI/CD
4. **Documentation**: Comprehensive guides

## Immediate Next Steps

### 1. Install Development Dependencies

```bash
# Using uv (recommended)
uv sync --all-groups

# Or using pip
pip install -e ".[dev]"
```

This installs:
- `pytest` and `pytest-cov` for testing
- `ruff` for linting and formatting
- `pre-commit` for automated hooks

### 2. Set Up Pre-commit Hooks

```bash
# Install the hooks
pre-commit install

# Test the hooks (optional)
pre-commit run --all-files
```

The hooks will now run automatically on every `git commit`, checking:
- Code style (Ruff)
- Trailing whitespace
- File endings
- YAML/JSON/TOML syntax
- Debug statements
- Large files

### 3. Run Code Quality Checks

```bash
# Check code style
ruff check .

# Auto-fix issues
ruff check --fix .

# Format code
ruff format .
```

### 4. Verify CI/CD Pipeline

The GitHub Actions workflow (`.github/workflows/ci.yml`) will run automatically on:
- Every push to `main` or `develop`
- Every pull request to `main` or `develop`
- Manual trigger via GitHub UI

No setup needed - it's ready to go!

### 5. Use the New Exporter Architecture

The exporter layer is now complete. To use it in new code:

```python
from outlook_exporter.core.models import ExportConfig
from outlook_exporter.exporters import ExporterFactory

# Create configuration
config = ExportConfig(
    output_dir=Path("./exports"),
    include_inline=False,
    hash_algorithm="sha256"
)

# Create the appropriate exporter
exporter = ExporterFactory.create_exporter(
    config=config,
    export_type="attachments"  # or "msg" or "markdown"
)

# Use it
result = ExportResult()
exporter.export(email_metadata, outlook_message, result)
```

### 6. Read the Documentation

- **CONTRIBUTING.md** - Complete developer guide
- **CHANGELOG.md** - Version history
- **docs/AUDIT_SUMMARY.md** - Implementation details

## What Happens Automatically

### On Every Commit (Local)
- Pre-commit hooks run quality checks
- Code is automatically formatted
- Invalid commits are rejected

### On Every Push/PR (GitHub)
- CI/CD pipeline runs 4 jobs:
  1. Code quality (Ruff)
  2. Tests (pytest with coverage)
  3. Security scanning (Safety + Bandit)
  4. Package build validation

### Results
- PRs get automatic quality feedback
- Coverage reports uploaded to Codecov
- Security issues flagged in artifacts

## Common Commands

```bash
# Run all quality checks locally
ruff check . && ruff format --check .

# Run tests with coverage
pytest --cov=src/outlook_exporter

# Update pre-commit hooks
pre-commit autoupdate

# Run pre-commit on all files
pre-commit run --all-files

# Build the package
python -m build
```

## Troubleshooting

### Pre-commit hooks not running
```bash
# Re-install
pre-commit uninstall
pre-commit install

# Clear cache
pre-commit clean
```

### Ruff errors
```bash
# See what's wrong
ruff check .

# Auto-fix what's possible
ruff check --fix .

# Check specific file
ruff check main.py
```

### Tests failing on Linux (pywin32)
This is expected - the project requires Windows for Outlook. The CI is configured to handle this gracefully.

## Next Development Steps

1. **Migrate main.py** to use new exporter architecture
2. **Update TUI** to import from `outlook_exporter` package
3. **Add integration tests** for Outlook interaction
4. **Write tests** for new exporters

See **CONTRIBUTING.md** for detailed development workflow.

## Questions?

- Read **CONTRIBUTING.md** for comprehensive guide
- Check **docs/AUDIT_SUMMARY.md** for implementation details
- Open a GitHub issue for questions

---

**Note**: All improvements are backward compatible. Existing code continues to work while you gradually adopt the new architecture.
