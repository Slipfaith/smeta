"""Application-wide logging configuration helpers."""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from tempfile import gettempdir
from typing import Optional


LOG_DIR_NAME = "logs"
LAST_RUN_LOG_NAME = "last_run.log"
ROTATING_LOG_NAME = "app.log"
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

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
    )

    last_run_handler = logging.FileHandler(
        last_run_log, mode="w", encoding="utf-8"
    )
    last_run_handler.setLevel(logging.DEBUG)
    last_run_handler.setFormatter(formatter)

    rotating_handler = RotatingFileHandler(
        rotating_log,
        maxBytes=ROTATING_MAX_BYTES,
        backupCount=ROTATING_BACKUP_COUNT,
        encoding="utf-8",
    )
    rotating_handler.setLevel(logging.DEBUG)
    rotating_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(
        logging.Formatter("%(levelname)s | %(name)s | %(message)s")
    )

    root_logger.addHandler(last_run_handler)
    root_logger.addHandler(rotating_handler)
    root_logger.addHandler(console_handler)

    _configured = True
    _last_run_log_path = last_run_log

    root_logger.debug("Logging configured. Logs directory: %s", log_dir)

    return last_run_log


def get_last_run_log_path() -> Path:
    """Return the path to ``last_run.log`` ensuring logging is configured."""

    if not _configured or _last_run_log_path is None:
        return setup_logging()
    return _last_run_log_path

