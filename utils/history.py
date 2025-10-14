"""Simple persistence layer for the rate selection history."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List

_HISTORY_FILE = Path.home() / ".smeta_rates_history.json"
_MAX_ENTRIES = 50


def _read_history() -> List[Dict[str, object]]:
    if not _HISTORY_FILE.exists():
        return []
    try:
        with _HISTORY_FILE.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, list):
            return [entry for entry in data if isinstance(entry, dict)]
    except (json.JSONDecodeError, OSError):
        pass
    return []


def _write_history(entries: List[Dict[str, object]]) -> None:
    try:
        with _HISTORY_FILE.open("w", encoding="utf-8") as fh:
            json.dump(entries[:_MAX_ENTRIES], fh, ensure_ascii=False, indent=2)
    except OSError:
        # Persistence errors should not break the main workflow.
        pass


def load_history() -> List[Dict[str, object]]:
    """Return previously saved rate selections."""
    return _read_history()


def add_entry(source: str, targets: List[str], is_second_file: bool) -> None:
    """Append a new history entry and persist it to disk."""
    entries = _read_history()
    entry = {
        "source": source,
        "targets": targets,
        "file": 2 if is_second_file else 1,
    }

    # Remove duplicate entries while preserving order.
    entries = [e for e in entries if not _compare_entries(e, entry)]
    entries.insert(0, entry)
    _write_history(entries)


def _compare_entries(left: Dict[str, object], right: Dict[str, object]) -> bool:
    return (
        left.get("source") == right.get("source")
        and list(left.get("targets", [])) == list(right.get("targets", []))
        and left.get("file", 1) == right.get("file", 1)
    )
