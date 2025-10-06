"""Application-wide logging configuration."""
from __future__ import annotations

import logging
import os
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

_LOGGER_NAME = "ProjectCalculator"
_MAX_BYTES = 5 * 1024 * 1024  # 5 MB
_BACKUP_COUNT = 3
_LOG_PATH: Optional[Path] = None


def _determine_log_path() -> Path:
    """Return the path to the application log file, creating directories if needed."""
    global _LOG_PATH
    appdata = os.getenv("APPDATA")
    if appdata:
        base_dir = Path(appdata) / "ProjectCalculator"
    else:
        # Fallback for non-Windows environments.
        base_dir = Path.home() / ".config" / "ProjectCalculator"

    log_dir = base_dir / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    _LOG_PATH = log_dir / "app.log"
    return _LOG_PATH


def _configure_root_logger() -> logging.Logger:
    """Configure and return the application's root logger."""
    logger = logging.getLogger(_LOGGER_NAME)
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)
    log_path = _determine_log_path()
    handler = RotatingFileHandler(
        log_path,
        maxBytes=_MAX_BYTES,
        backupCount=_BACKUP_COUNT,
        encoding="utf-8",
    )
    formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(name)s - %(message)s"
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.propagate = False
    return logger


def get_logger(name: Optional[str] = None) -> logging.Logger:
    """Return a logger instance scoped to the given name."""
    root_logger = _configure_root_logger()
    if name:
        return root_logger.getChild(name)
    return root_logger


def get_log_file_path() -> Path:
    """Return the path to the current log file."""
    if _LOG_PATH is None:
        return _determine_log_path()
    return _LOG_PATH


# Initialize the root logger on module import.
logger = get_logger()
