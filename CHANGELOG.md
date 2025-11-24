# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

#### Pillar 1: Code Quality & Maintainability
- **Ruff linting and formatting** - Comprehensive code quality tool replacing flake8, black, isort, and pyupgrade
  - Configured in `pyproject.toml` with strict settings for Python 3.12+
  - Line length: 120 characters
  - Full type hint enforcement
  - PEP 8 compliance with modern Python practices
- **Pre-commit hooks** - Automated code quality checks before commits
  - Ruff linting and formatting
  - Trailing whitespace and end-of-file fixes
  - YAML/JSON/TOML syntax validation
  - Debug statement detection
  - Large file prevention
  - Configuration in `.pre-commit-config.yaml`

#### Pillar 2: Architecture & Design
- **Complete exporter layer implementation** - Finished the refactoring architecture
  - `BaseExporter` abstract class with `ExporterFactory` for strategy pattern
  - `AttachmentExporter` - Exports attachments with duplicate detection
  - `MessageExporter` - Exports complete emails as .msg files
  - `MarkdownExporter` - Exports emails as markdown with YAML frontmatter
  - All exporters follow single responsibility principle
  - Clean separation of concerns from legacy `main.py`

#### Pillar 3: Testing & Quality Assurance  
- **GitHub Actions CI/CD pipeline** - Automated testing and quality checks
  - Code quality job: Ruff linting and formatting checks
  - Test job: pytest with coverage reporting on Python 3.12 and 3.13
  - Security job: Safety (dependency vulnerabilities) and Bandit (security linting)
  - Build job: Package building and validation with twine
  - Codecov integration for coverage tracking
  - Automated artifact uploads for reports
- **Enhanced pytest configuration** - Comprehensive test settings
  - Coverage reporting with HTML output
  - Strict test configuration
  - Clear coverage exclusions for COM code

#### Pillar 4: Documentation & Developer Experience
- **CONTRIBUTING.md** - Comprehensive contributor guide
  - Development setup instructions (uv and pip)
  - Code quality standards and tooling
  - Testing guidelines and examples
  - Architecture overview with diagrams
  - Design patterns documentation
  - Common troubleshooting solutions
  - VSCode configuration recommendations
  - Pull request process and checklist

### Changed
- Updated `pyproject.toml` with:
  - Development dependencies (pytest-cov, ruff, pre-commit)
  - Ruff configuration (lint rules, formatting, per-file ignores)
  - Pytest configuration (coverage, test discovery)
  - Coverage configuration (source, exclusions)

### Technical Details

**Quality Improvements:**
- Enforced code style consistency across the entire codebase
- Automated quality checks prevent low-quality code from entering the repository
- Pre-commit hooks catch issues before they reach CI/CD

**Architecture Improvements:**
- Completed the exporter layer, making the refactored architecture production-ready
- Clear migration path from legacy `main.py` to new architecture
- Extensible design allows easy addition of new export formats

**Testing Improvements:**
- Automated testing on every push and pull request
- Multi-version Python testing ensures compatibility
- Security scanning catches vulnerabilities early
- Coverage tracking identifies untested code

**Documentation Improvements:**
- Comprehensive onboarding for new contributors
- Clear architecture documentation reduces confusion
- Troubleshooting guide reduces support burden
- VSCode recommendations improve developer experience

## [0.2.0] - 2025-10-14

### Added
- Refactored architecture with modular package structure
- Domain models (EmailMetadata, Attachment, ExportConfig, FilterCriteria, ExportResult)
- Filter system with Strategy and Composite patterns
- Outlook client with adapters for COM interaction
- Path utilities for Windows compatibility
- Duplicate detection with configurable hashing
- Custom exception hierarchy
- Comprehensive unit tests (90%+ coverage for new modules)

### Changed
- Project structure moved to `src/outlook_exporter/` layout
- Separated concerns across multiple modules

## [0.1.0] - 2024-12-01

### Added
- Initial release
- CLI for exporting Outlook attachments
- TUI built with Textual
- Basic filtering (date, sender, subject, body)
- Duplicate detection with SHA256
- Markdown export support
- Organized folder structure (sender/subject/date)

[Unreleased]: https://github.com/mberetvas/outlook-exporter-tui/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/mberetvas/outlook-exporter-tui/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/mberetvas/outlook-exporter-tui/releases/tag/v0.1.0
