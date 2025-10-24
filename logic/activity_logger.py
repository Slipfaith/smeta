"""High-level helpers for structured user activity logging."""

from __future__ import annotations

import json
import logging
from typing import Any, Mapping, Optional

from .logging_utils import setup_logging

_ACTIVITY_LOGGER_NAME = "activity"


def _ensure_logger() -> logging.Logger:
    """Return the activity logger ensuring handlers are configured."""

    setup_logging()
    logger = logging.getLogger(_ACTIVITY_LOGGER_NAME)
    if logger.level == logging.NOTSET:
        logger.setLevel(logging.DEBUG)
    return logger


def _stringify(value: Any) -> str:
    if isinstance(value, str):
        return value
    try:
        return json.dumps(value, ensure_ascii=False)
    except TypeError:
        return repr(value)


def _format_details(details: Mapping[str, Any]) -> str:
    lines: list[str] = []
    for key, value in details.items():
        title = key.replace("_", " ").strip() or "значение"
        lines.append(f"- **{title.title()}:** {_stringify(value)}")
    return "\n".join(lines)


def log_user_action(
    action: str,
    *,
    details: Optional[Mapping[str, Any]] = None,
    snapshot: Optional[Mapping[str, Any]] = None,
    level: int = logging.INFO,
) -> None:
    """Log a high-level user action with optional structured details."""

    logger = _ensure_logger()

    body_lines = [f"#### {action}"]
    if details:
        formatted = _format_details(details)
        if formatted:
            body_lines.append("")
            body_lines.append(formatted)
    if snapshot:
        try:
            snapshot_json = json.dumps(snapshot, ensure_ascii=False, indent=2)
        except TypeError:
            snapshot_json = json.dumps(str(snapshot), ensure_ascii=False)
        body_lines.extend(
            [
                "",
                "<details><summary>Снимок данных</summary>",
                "",
                "```json",
                snapshot_json,
                "```",
                "</details>",
            ]
        )

    logger.log(level, "\n".join(body_lines))


def log_window_action(
    action: str,
    window: Any,
    *,
    details: Optional[Mapping[str, Any]] = None,
    level: int = logging.INFO,
    include_snapshot: bool = True,
) -> None:
    """Log an action and optionally attach the current project snapshot."""

    snapshot: Optional[Mapping[str, Any]] = None
    snapshot_error: Optional[str] = None

    if include_snapshot:
        try:
            from .project_data import ProjectData  # Local import to avoid cycles

            snapshot = ProjectData.from_window(window).to_mapping()
        except Exception as exc:  # pragma: no cover - defensive
            snapshot_error = str(exc)

    merged_details: dict[str, Any] = dict(details or {})
    if snapshot_error:
        merged_details.setdefault("Ошибка снимка", snapshot_error)

    log_user_action(action, details=merged_details, snapshot=snapshot, level=level)
