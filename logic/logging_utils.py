"""Application-wide logging configuration helpers."""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
from pathlib import Path
from tempfile import gettempdir
from typing import Optional


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

