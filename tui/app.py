import argparse
import logging
import sys
from pathlib import Path

# Add the project root to the Python path to allow importing 'main'
project_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(project_root))

from textual.app import App, ComposeResult
from textual.containers import Horizontal, VerticalScroll, GridLayout, Vertical
from textual.widgets import (
    Button,
    Checkbox,
    Header,
    Footer,
    Input,
    Label,
    RadioButton,
    RadioSet,
    RichLog,
)
from textual import work

from main import run_export


class TuiLogger(logging.Handler):
    """A logging handler that sends logs to a RichLog widget."""

    def __init__(self, rich_log: RichLog):
        super().__init__()
        self.rich_log = rich_log

    def emit(self, record: logging.LogRecord) -> None:
        """Emit a record."""
        self.rich_log.write(self.format(record))


class OutlookExportTUI(App):
    """A Textual TUI for the Outlook Attachments Export tool."""

    TITLE = "Outlook Attachments Export"
    SUB_TITLE = "Export attachments with filtering and duplicate handling"

    CSS_PATH = "app.css"

    def compose(self) -> ComposeResult:
        """Create child widgets for the app."""
        yield Header()
        with Horizontal():
            with VerticalScroll(id="filters"):
                with Vertical(id="filters-grid"):
                    yield Label("Output Directory (*)")
                    yield Input(placeholder="/path/to/output", id="output-dir")
                    yield Label("Start Date (YYYY-MM-DD)")
                    yield Input(placeholder="2023-01-01", id="start-date")
                    yield Label("End Date (YYYY-MM-DD)")
                    yield Input(placeholder="2023-12-31", id="end-date")
                    yield Label("Sender Emails (comma-separated)")
                    yield Input(placeholder="sender@example.com", id="sender")
                    yield Label("Subject Keywords (comma-separated)")
                    yield Input(
                        placeholder="keyword1, keyword2", id="subject-keyword"
                    )
                    yield Label("Body Keywords (comma-separated)")
                    yield Input(placeholder="keyword1, keyword2", id="body-keyword")
                    label = Label("Include Inline Attachments")
                    label.tooltip = "Include images and other files embedded in the email body (e.g., signatures)."
                    yield label
                    yield Checkbox(id="include-inline")
                    label = Label("Export Full Email (.msg)")
                    label.tooltip = "Export the entire email as a .msg file instead of just the attachments."
                    yield label
                    yield Checkbox(id="export-mail")
                with Vertical(id="attachment-options-container"):
                    yield Label("Attachment Options:", classes="grid-span-2")
                    with RadioSet(id="attachment-options", classes="grid-span-2"):
                        yield RadioButton("All emails", id="all-emails")
                        yield RadioButton(
                            "Emails with attachments", id="with-attachments"
                        )
                        yield RadioButton(
                            "Emails without attachments", id="without-attachments"
                        )
            yield RichLog(id="log", wrap=True)
        with Horizontal(id="buttons"):
            yield Button("Run", id="run", variant="primary")
            yield Button("Dry Run", id="dry-run")
            yield Button("Quit", id="quit")
        yield Footer()

    def on_mount(self) -> None:
        """Called when the app is mounted."""
        self.query_one("#all-emails", RadioButton).value = True
        log_widget = self.query_one(RichLog)
        handler = TuiLogger(log_widget)
        logging.basicConfig(
            level=logging.INFO,
            handlers=[handler],
            format="%(asctime)s %(levelname)-8s %(message)s",
        )

    def on_button_pressed(self, event: Button.Pressed) -> None:
        """Event handler called when a button is pressed."""
        if event.button.id == "run" or event.button.id == "dry-run":
            self.run_export_worker(dry_run=event.button.id == "dry-run")
        elif event.button.id == "quit":
            self.exit()

    @work(exclusive=True, thread=True)
    def run_export_worker(self, dry_run: bool) -> None:
        """Runs the export process in a worker."""
        args = self.get_args(dry_run)
        if args:
            run_export(args)

    def get_args(self, dry_run: bool) -> argparse.Namespace | None:
        """Gathers the arguments from the TUI."""
        try:
            output_dir = self.query_one("#output-dir", Input).value
            if not output_dir:
                logging.error("Output directory is required.")
                return None

            pressed_button = self.query_one(RadioSet).pressed_button
            attachment_option_id = (
                pressed_button.id if pressed_button else "all-emails"
            )

            with_attachments = attachment_option_id == "with-attachments"
            without_attachments = attachment_option_id == "without-attachments"

            args = argparse.Namespace(
                output=Path(output_dir),
                start_date=self.query_one("#start-date", Input).value or None,
                end_date=self.query_one("#end-date", Input).value or None,
                sender=self.query_one("#sender", Input).value.split(",")
                if self.query_one("#sender", Input).value
                else None,
                subject_keyword=self.query_one("#subject-keyword", Input).value.split(",")
                if self.query_one("#subject-keyword", Input).value
                else None,
                body_keyword=self.query_one("#body-keyword", Input).value.split(",")
                if self.query_one("#body-keyword", Input).value
                else None,
                with_attachments=with_attachments,
                without_attachments=without_attachments,
                folder=None,
                limit=None,
                batch_size=200,
                dry_run=dry_run,
                open_folder=False,
                log_file=None,
                duplicates_subfolder="duplicates",
                hash_algorithm="sha256",
                include_inline=self.query_one("#include-inline", Checkbox).value,
                export_mail=self.query_one("#export-mail", Checkbox).value,
                subject_sanitize_length=80,
                verbose=0,
                quiet=False,
            )
            return args
        except Exception as e:
            logging.error(f"Error getting arguments: {e}")
            return None


def main() -> None:
    """The main function for the TUI."""
    app = OutlookExportTUI()
    app.run()


if __name__ == "__main__":
    main()
