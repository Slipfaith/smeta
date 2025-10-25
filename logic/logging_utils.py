"""Application-wide logging configuration helpers."""

from __future__ import annotations

import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from datetime import datetime
from pathlib import Path
from tempfile import gettempdir
from typing import Optional

if os.name == "nt":  # pragma: no cover - Windows-specific behaviour
    from ctypes import windll


_CONSOLE_HANDLER_MARK = "__smeta_console_handler__"


LOG_DIR_NAME = "logs"
LAST_RUN_LOG_NAME = "last_run.md"
ROTATING_LOG_NAME = "app.md"
ROTATING_MAX_BYTES = 5 * 1024 * 1024  # 5 MB
ROTATING_BACKUP_COUNT = 5

_configured = False
_last_run_log_path: Optional[Path] = None


def _candidate_log_directories() -> list[Path]:
    cwd_logs = Path.cwd() / LOG_DIR_NAME
    home_logs = Path.home() / ".smeta" / LOG_DIR_NAME
    temp_logs = Path(gettempdir()) / "smeta_logs"
    return [cwd_logs, home_logs, temp_logs]


def _ensure_log_directory() -> Path:
    for directory in _candidate_log_directories():
        try:
            directory.mkdir(parents=True, exist_ok=True)
        except OSError:
            continue
        else:
            return directory
    # As a last resort, fall back to the current working directory even if
    # ``mkdir`` kept failing (unlikely).
    return Path.cwd()


class _MarkdownFormatter(logging.Formatter):
    """Render log records as Markdown blocks."""

    def format(self, record: logging.LogRecord) -> str:  # pragma: no cover - formatting
        timestamp = datetime.fromtimestamp(record.created).strftime("%Y-%m-%d %H:%M:%S")
        message = super().format(record).rstrip()
        if message:
            message = f"{message}\n"
        return f"---\n### {timestamp} · {record.levelname}\n\n{message}"


def _initialise_markdown_file(path: Path, title: str, fresh: bool) -> None:
    """Ensure *path* starts with a Markdown heading."""

    header = f"# {title}\n\n"
    if fresh:
        path.write_text(header, encoding="utf-8")
    elif not path.exists() or path.stat().st_size == 0:
        path.write_text(header, encoding="utf-8")


def _attach_to_parent_console() -> bool:  # pragma: no cover - requires Windows
    """Attach the process to the parent console if possible."""

    if os.name != "nt":
        return False

    ATTACH_PARENT_PROCESS = -1
    try:
        attached = bool(windll.kernel32.AttachConsole(ATTACH_PARENT_PROCESS))
    except OSError:
        return False

    if not attached:
        return False

    # ``AttachConsole`` succeeds but standard streams may still be missing.
    # Re-open them against the console device so ``print`` / logging works.
    if getattr(sys, "stdout", None) is None:
        sys.stdout = open("CONOUT$", "w", encoding="utf-8", buffering=1)
    if getattr(sys, "stderr", None) is None:
        sys.stderr = open("CONOUT$", "w", encoding="utf-8", buffering=1)
    if getattr(sys, "stdin", None) is None:
        sys.stdin = open("CONIN$", "r", encoding="utf-8", buffering=1)

    return True


def enable_console_logging() -> bool:
    """Stream log records to the console when available.

    Returns
    -------
    bool
        ``True`` when logging to a console stream is now active.
    """

    root_logger = logging.getLogger()

    # Avoid configuring duplicate console handlers on repeated calls.
    for handler in root_logger.handlers:
        if getattr(handler, _CONSOLE_HANDLER_MARK, False):
            return True

    stream = sys.stderr if sys.stderr is not None else sys.stdout

    if stream is None and not _attach_to_parent_console():
        return False

    stream = sys.stderr if sys.stderr is not None else sys.stdout
    if stream is None:
        return False

    console_handler = logging.StreamHandler(stream)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(
        logging.Formatter("[%(levelname)s] %(message)s")
    )
    setattr(console_handler, _CONSOLE_HANDLER_MARK, True)
    root_logger.addHandler(console_handler)
    return True


def setup_logging() -> Path:
    """Configure logging handlers for the application.

    Returns
    -------
    Path
        Path to the ``last_run.log`` file.
    """

    global _configured, _last_run_log_path

    if _configured and _last_run_log_path is not None:
        return _last_run_log_path

    log_dir = _ensure_log_directory()
    last_run_log = log_dir / LAST_RUN_LOG_NAME
    rotating_log = log_dir / ROTATING_LOG_NAME

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    # Remove any existing handlers to avoid duplicate logs when reconfiguring.
    for handler in list(root_logger.handlers):
        root_logger.removeHandler(handler)

    formatter = _MarkdownFormatter("%(message)s")

    _initialise_markdown_file(last_run_log, "Журнал последнего запуска", fresh=True)
    last_run_handler = logging.FileHandler(
        last_run_log, mode="a", encoding="utf-8"
    )
    last_run_handler.setLevel(logging.DEBUG)
    last_run_handler.setFormatter(formatter)

    _initialise_markdown_file(
        rotating_log,
        "История активности приложения",
        fresh=not rotating_log.exists(),
    )
    rotating_handler = RotatingFileHandler(
        rotating_log,
        maxBytes=ROTATING_MAX_BYTES,
        backupCount=ROTATING_BACKUP_COUNT,
        encoding="utf-8",
    )
    rotating_handler.setLevel(logging.DEBUG)
    rotating_handler.setFormatter(formatter)

    root_logger.addHandler(last_run_handler)
    root_logger.addHandler(rotating_handler)

    _configured = True
    _last_run_log_path = last_run_log

    root_logger.debug("Logging configured. Logs directory: %s", log_dir)

    return last_run_log


def get_last_run_log_path() -> Path:
    """Return the path to ``last_run.log`` ensuring logging is configured."""

    if not _configured or _last_run_log_path is None:
        return setup_logging()
    return _last_run_log_path

